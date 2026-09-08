#!/usr/bin/env python3
"""
Knowledge Chest - remove working-paper duplicates that are now published.

When a working paper later appears in a journal AND its published version is
already catalogued as a separate entry, the working-paper row is a duplicate.
This script finds each such working paper (matched by title + author last names)
and DELETES the working-paper row, keeping the published entry. It also removes
the row's dataset links and any dataset left with no papers, then rebuilds the
dashboard + spreadsheet.

Difference from kc_promote.py: promote UPDATES a lone working-paper entry into
its published form (no published duplicate exists yet). This script is for the
case where BOTH the working paper and the published paper are in the database,
so the right fix is to drop the working paper.

Claude fills in REMOVALS below (title + authors of the working paper to drop).
Safe to re-run: a removal is applied only when the working paper still exists
AND a matching published entry is present; otherwise it is skipped.

Run from repo root:  python kc_dedupe.py
Then:                git commit -am "remove working-paper duplicates" && git push
Requires: kc.py in the repo, and  pip install openpyxl
"""
import re, sys, json
from pathlib import Path
import kc

REPO = Path(".").resolve()
kc.DB_PATH    = REPO / "data" / "knowledge_chest.db"
kc.OUTPUT_DIR = REPO / "docs"

# -------------------------------------------------------------------------
# Each entry identifies the WORKING PAPER to remove. `authors` (last names)
# makes the match unambiguous. The removal happens only if a PUBLISHED entry
# with a matching title + author also exists.
REMOVALS = [
    {"title": "Automation and Rent Dissipation: Implications for Wages, Inequality, Productivity, and Growth",
     "authors": ["Acemoglu", "Restrepo"]},
    {"title": "The Cost of Regulatory Compliance in the United States",
     "authors": ["Trebbi", "Zhang", "Simkovic"]},
]
# -------------------------------------------------------------------------

def norm(s):
    return re.sub(r"[^a-z0-9 ]", " ", (s or "").lower()).split()

def title_match(a_title, b_title):
    a, b = norm(a_title), norm(b_title)
    if not a or not b:
        return False
    contains = " ".join(a) in " ".join(b) or " ".join(b) in " ".join(a)
    overlap = len(set(a) & set(b)) / max(len(set(a)), len(set(b)))
    return contains or overlap >= 0.8

def authors_ok(want, have):
    if not want:
        return True
    return len({x.lower() for x in want} & {x.lower() for x in have}) >= 1

def main():
    if not REMOVALS:
        print("No removals listed - nothing to do."); return 0
    conn = kc.get_conn()
    rows = [dict(r) for r in conn.execute("SELECT * FROM papers")]
    removed = 0
    touched_ds = set()   # datasets that were linked to a removed working paper
    for want in REMOVALS:
        auth = want.get("authors")
        wps  = [r for r in rows if r["is_working_paper"] == 1
                and title_match(want["title"], r["title"])
                and authors_ok(auth, json.loads(r["authors_json"] or "[]"))]
        pubs = [r for r in rows if r["is_working_paper"] == 0
                and title_match(want["title"], r["title"])
                and authors_ok(auth, json.loads(r["authors_json"] or "[]"))]
        if not wps:
            print(f"No working paper found for {want['title']!r} - skipped (already removed?)."); continue
        if not pubs:
            print(f"NO PUBLISHED match for {want['title']!r} - skipped (nothing to dedupe against)."); continue
        if len(wps) > 1:
            print(f"AMBIGUOUS: {want['title']!r} matched {len(wps)} working papers - skipped:")
            for w in wps: print("   -", w["std_name"])
            continue
        w = wps[0]
        # Remember this paper's datasets, drop its links, then delete the row.
        touched_ds.update(r["dataset_id"] for r in conn.execute(
            "SELECT dataset_id FROM paper_datasets WHERE paper_id=?", (w["id"],)))
        conn.execute("DELETE FROM paper_datasets WHERE paper_id=?", (w["id"],))
        conn.execute("DELETE FROM papers WHERE id=?", (w["id"],))
        removed += 1
        print(f"REMOVED working paper: {w['std_name'][:74]}")
        print(f"    kept published : {pubs[0]['std_name'][:74]}")
    # Clean up ONLY datasets that this removal orphaned. Standalone catalogue
    # datasets that were never linked to a paper are left untouched.
    cleaned = 0
    for ds_id in touched_ds:
        still = conn.execute(
            "SELECT 1 FROM paper_datasets WHERE dataset_id=? LIMIT 1", (ds_id,)).fetchone()
        if still:
            continue
        o = conn.execute("SELECT provider, product FROM datasets WHERE id=?", (ds_id,)).fetchone()
        conn.execute("DELETE FROM datasets WHERE id=?", (ds_id,))
        cleaned += 1
        print(f"  removed dataset orphaned by this removal: {o['provider']}"
              + (f" - {o['product']}" if o and o["product"] else ""))
    conn.commit(); conn.close()
    if removed or cleaned:
        kc.build()
    print(f"\n{removed} working-paper duplicate(s) removed, {cleaned} orphaned dataset(s) cleaned.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
