#!/usr/bin/env python3
"""
Knowledge Chest - promote working papers to published.

When a working paper appears in a journal, this updates its existing catalog
entry IN PLACE (no duplicate): it matches the entry by title + author last
names, replaces "Working Paper" with the journal name, recomputes the
standardized name, relabels the PDF in the repo if it's present, and rebuilds
the dashboard + spreadsheet.

Claude fills in PROMOTIONS below for each paper that got published, then you
upload this file and run the "promote" workflow (or `python kc_promote.py`).
Matching is bibliographic only - the paper text is not re-read - so the
summary, logical flow, research design, and datasets are kept as they were.
"""
import re, subprocess, sys, json
from pathlib import Path
import kc

REPO = Path(".").resolve()
kc.DB_PATH    = REPO / "data" / "knowledge_chest.db"
kc.OUTPUT_DIR = REPO / "docs"
kc.MOVE_PROCESSED = False

# -------------------------------------------------------------------------
# Claude fills this in per publication. `authors` is optional but makes the
# match unambiguous. `year` is optional (use it only if the published year
# differs from what's stored).
PROMOTIONS = [
    # {"title": "Homemade Foreign Trading",
    #  "authors": ["He", "Wang", "Zhu"],
    #  "journal": "Journal of Finance",
    #  "year": 2027},
]
# -------------------------------------------------------------------------

def norm(s):
    return re.sub(r"[^a-z0-9 ]", "", (s or "").lower()).split()

def title_match(promo_title, db_title):
    a, b = norm(promo_title), norm(db_title)
    if not a or not b:
        return False
    sa, sb = set(a), set(b)
    # one contains the other, or strong token overlap
    contains = " ".join(a) in " ".join(b) or " ".join(b) in " ".join(a)
    overlap = len(sa & sb) / max(len(sa), len(sb))
    return contains or overlap >= 0.8

def authors_ok(promo_authors, db_authors):
    if not promo_authors:
        return True
    pa = {x.lower() for x in promo_authors}
    da = {x.lower() for x in db_authors}
    return len(pa & da) >= 1

def git(*a):
    subprocess.run(["git", *a], cwd=REPO, check=True)

def main():
    if not PROMOTIONS:
        print("No promotions listed - nothing to do."); return 0
    conn = kc.get_conn()
    rows = [dict(r) for r in conn.execute("SELECT * FROM papers")]
    changed = 0
    for promo in PROMOTIONS:
        cands = [r for r in rows if title_match(promo["title"], r["title"])
                 and authors_ok(promo.get("authors"), json.loads(r["authors_json"]))]
        if len(cands) == 0:
            print(f"NO MATCH for {promo['title']!r} - skipped (check spelling/authors)."); continue
        if len(cands) > 1:
            print(f"AMBIGUOUS: {promo['title']!r} matched {len(cands)} entries - skipped:")
            for c in cands: print("   -", c["std_name"]); 
            continue
        p = cands[0]
        old_std = p["std_name"]
        rec = {"is_working_paper": False, "journal": promo["journal"],
               "year": promo.get("year") or p["year"],
               "authors": json.loads(p["authors_json"]), "title": p["title"]}
        new_std = kc.build_std_name(rec)
        if new_std == old_std:
            print(f"Already up to date: {old_std[:70]}"); continue
        conn.execute("UPDATE papers SET journal=?, is_working_paper=0, year=?, std_name=? WHERE id=?",
                     (promo["journal"], rec["year"], new_std, p["id"]))
        # relabel the PDF in the repo if it happens to be there
        old_pdf = REPO / (old_std[:kc.MAX_FILENAME].rstrip(" .") + ".pdf")
        new_pdf = REPO / (new_std[:kc.MAX_FILENAME].rstrip(" .") + ".pdf")
        if old_pdf.exists() and old_pdf != new_pdf:
            git("mv", "--", old_pdf.name, new_pdf.name)
            print(f"  relabeled file -> {new_pdf.name[:66]}")
        changed += 1
        print(f"PROMOTED: {p['title'][:50]}")
        print(f"    was: {old_std[:74]}")
        print(f"    now: {new_std[:74]}")
    conn.commit(); conn.close()
    if changed:
        kc.build(); git("add", "-A")
    print(f"\n{changed} entr{'y' if changed==1 else 'ies'} promoted.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
