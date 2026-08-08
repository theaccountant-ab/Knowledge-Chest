#!/usr/bin/env python3
"""
Knowledge Chest - apply one catalogued batch to THIS repo.

Claude catalogued these papers in chat; the results are baked in below. Running
this script renames each PDF to its standardized name (via `git mv`), updates
data/knowledge_chest.db, and rebuilds docs/index.html + docs/knowledge_chest.xlsx.

No API key, no model, no internet - it just replays work already done.
Run it from the repo root:   python kc_apply.py
Then:                        git commit -am "catalog batch" && git push

Requires: kc.py in the repo, and  pip install openpyxl pymupdf
Safe to re-run: papers already in the database are skipped (dedup by file hash).
"""
import json, subprocess, sys
from pathlib import Path
import kc

REPO = Path(".").resolve()
kc.DB_PATH = REPO / "data" / "knowledge_chest.db"
kc.OUTPUT_DIR = REPO / "docs"
kc.MOVE_PROCESSED = False
kc.DB_PATH.parent.mkdir(parents=True, exist_ok=True)

RECORDS = json.loads(r'''[
 {
  "journal": "American Economic Review",
  "is_working_paper": false,
  "year": 2025,
  "authors": [
   "Lerche"
  ],
  "title": "Direct and Indirect Effects of Investment Tax Incentives",
  "summary": "The paper estimates both the direct effects and the local spillover effects of investment tax credits on firms, using a German policy that changed the tax-credit rate differently by firm size. Lerche finds that lowering a firm's investment cost by 7.6 percent raises its capital stock by 17.7 percent and employment by 12.0 percent. Positive local spillovers create roughly one additional manufacturing job for each directly created job; they are strongest between firms linked through input-output relationships and operate within about five kilometers. Firms reliant on local consumer demand also expand employment, while within-industry spillovers are slightly negative. The paper shows that general-equilibrium spillovers substantially amplify the direct effects of investment tax incentives.",
  "logical_flow": "The paper begins from the long-standing use of investment tax credits to lower the cost of capital and recent firm-level evidence that they raise investment, while stressing that general-equilibrium spillovers determine how firm-level effects translate into aggregate outcomes—agglomeration and local-demand forces may amplify them, while fixed labor supply and product-market competition may offset them through reallocation. To estimate both direct and indirect effects, Lerche exploits an investment tax credit introduced after German reunification whose rate fell more for small firms, creating a differential change in the user cost of capital by firm size. Using administrative matched employer-employee data and firm addresses, he first estimates direct (partial-equilibrium) effects with a difference-in-differences design comparing firms just below and above the 250-employee cutoff (excluding those right at the cutoff). He finds large direct increases in capital and employment. He then combines this with a second difference-in-differences design based on the local (county) exposure to the policy—excluding a firm's own employment—to identify local spillovers. He finds positive local spillovers of about one extra manufacturing job per directly created job, strongest among input-output-linked industries and within five kilometers, plus positive local-demand effects and small negative within-industry spillovers. The paper concludes that investment tax incentives have sizable direct effects that are meaningfully amplified by local general-equilibrium spillovers.",
  "research_design": "A quasi-experimental study exploiting a post-reunification German investment tax credit whose rate changed differentially by firm size (around a 250-employee cutoff), using administrative matched employer-employee data with firm addresses. Direct effects are estimated with a difference-in-differences design comparing firms below and above the size cutoff; local spillovers are identified by combining this with a second difference-in-differences design based on county-level policy exposure (excluding a firm's own employment). Identification comes from the size-differential rate change and local exposure.",
  "categories": [
   "Public Economics",
   "Labor Economics",
   "Regional & Urban Economics"
  ],
  "datasets": [
   {
    "provider": "Institute for Employment Research (IAB, Germany)",
    "product": "administrative matched employer-employee data",
    "description": "German administrative surveys and matched employer-employee data with firm addresses, used to measure firm capital and employment responses and local spillovers around the tax-credit change.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "Germany; matched employer-employee; investment; tax credit"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Post-reunification German investment tax credit (size-differential rate change)",
   "type": "policy natural experiment (difference-in-differences)",
   "what": "An investment tax credit whose rate changed more for small firms (around a 250-employee cutoff), providing differential variation in the user cost of capital used to identify direct and local spillover effects."
  },
  "missing_notes": null,
  "orig_filename": "American Economic Review – 2025 – Lerche – Direct and Indirect Effects of Investment Tax Incentives.pdf",
  "std_name": "American Economic Review - 2025 - Lerche - Direct and Indirect Effects of Investment Tax Incentives"
 },
 {
  "journal": "American Law and Economics Review",
  "is_working_paper": false,
  "year": 2023,
  "authors": [
   "Ohlrogge"
  ],
  "title": "Down the Tubes: Financial Distress, Bankruptcy, and Industrial Water Pollution",
  "summary": "The paper examines how firms' efforts to avoid harming third parties change as they approach, enter, and exit bankruptcy, focusing on industrial water pollution. Studying about 350 US firms regulated under the Clean Water Act that declared bankruptcy, Ohlrogge finds that as firms approach bankruptcy their rate of releasing pollutants beyond permitted limits rises by 50 percent. Once they file, compliance improves dramatically, returning to a baseline from well before filing. He argues the pattern is best explained by a mix of moral-hazard problems and financing frictions affecting thinly capitalized firms. The findings support policy interventions such as changes to bankruptcy claim priority and suggest an underappreciated public benefit of encouraging firms to file early.",
  "logical_flow": "The paper starts from the recognized idea that financial distress distorts firm behavior—underinvestment from debt overhang and shareholder-creditor conflicts—and asks how distress and bankruptcy affect a firm's efforts to avoid harming third parties, here through water pollution. It assembles a panel of roughly 350 US firms regulated under the Clean Water Act (via NPDES permits) that all eventually declare bankruptcy, and studies their permitted-limit exceedances over time. Using a difference-in-differences-style comparison to similar facilities owned by non-bankrupt firms, Ohlrogge finds that as firms approach bankruptcy their rate of exceeding pollution limits rises by about 50 percent, and that after filing, compliance improves sharply back to a pre-distress baseline. He explores mechanisms and argues the pattern reflects a nexus of moral hazard (thinly capitalized firms have weak incentives to invest in compliance when liabilities can be discharged) and financing frictions that ease once bankruptcy resolves distress; he notes the bankruptcy code's higher administrative priority for post-filing environmental penalties as relevant. He draws policy implications, including changing claim priority to mitigate moral hazard and encouraging early filing. The paper concludes that bankruptcy can reduce public harms from distressed firms, implying benefits to policies that promote timely resolution of distress.",
  "research_design": "An empirical panel study of about 350 US Clean Water Act (NPDES)-regulated firms that declare bankruptcy, using discharge-monitoring data on permitted-limit exceedances over time. A difference-in-differences-style design compares bankrupt firms' compliance around approach, filing, and exit to similar facilities owned by non-bankrupt firms. Identification comes from the timing of financial distress and bankruptcy filing relative to comparison facilities.",
  "categories": [
   "Law & Economics",
   "Corporate Finance",
   "Environmental Economics"
  ],
  "datasets": [
   {
    "provider": "US Environmental Protection Agency",
    "product": "NPDES discharge monitoring reports (Clean Water Act)",
    "description": "Facility-level water-pollution discharge monitoring reports and permitted-limit exceedances under the Clean Water Act's NPDES program, for roughly 350 bankrupt firms and comparison facilities.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "water pollution; NPDES; compliance; Clean Water Act"
   },
   {
    "provider": "US bankruptcy filings",
    "product": "corporate bankruptcy records",
    "description": "Records of corporate bankruptcy filings used to date firms' approach to, entry into, and exit from bankruptcy.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "bankruptcy; financial distress; firms"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Corporate bankruptcy filing",
   "type": "event study / difference-in-differences",
   "what": "The approach to, filing of, and exit from bankruptcy by CWA-regulated firms, whose timing is used to identify changes in pollution compliance relative to non-bankrupt comparison facilities."
  },
  "missing_notes": null,
  "orig_filename": "American Law and Economics Review - 2023 - Ohlrogge - Down the Tubes Financial Distress, Bankruptcy, and Industrial Water Pollution​.pdf",
  "std_name": "American Law and Economics Review - 2023 - Ohlrogge - Down the Tubes Financial Distress, Bankruptcy, and Industrial Water Pollution"
 },
 {
  "journal": "Annual Review of Economics",
  "is_working_paper": false,
  "year": 2022,
  "authors": [
   "Baker",
   "Kueng"
  ],
  "title": "Household Financial Transaction Data",
  "summary": "This review surveys the fast-growing use of detailed household financial transaction micro-data in economics and finance. Such data—drawn from banks, FinTech aggregator apps, payment intermediaries, and card companies—provide comprehensive, high-frequency panels of financial flows and balances for many thousands to millions of individuals. The authors review how these data have advanced research on consumption, household balance sheets, and responses to income shocks and policy, and they weigh their benefits and limitations against traditional survey, scanner, and tax-based measures. They also discuss the data's future potential for firm-focused research, real-time policy analysis, and macroeconomic statistics. The paper serves as a guide to the sources, uses, and tradeoffs of household transaction data.",
  "logical_flow": "The review begins by noting that measuring household consumption and finances well is central to many questions, yet traditional sources—government surveys, scanner data, and tax-based imputation—suffer from high cost, small samples, short panels, slow updating, partial coverage, or measurement error. It then introduces a newer class of data: comprehensive panels of financial transactions from credit, debit, and checking accounts, sourced from banks, financial aggregators (FinTech apps), and card companies, often at high frequency and spanning many account types. The authors organize how researchers have used these data to study consumption behavior, household balance sheets, and responses to income fluctuations and policy changes, improving understanding of heterogeneity in income, balance sheets, beliefs, and preferences. They systematically compare the benefits (coverage, frequency, sample size, timeliness) and limitations (representativeness, incomplete account linkage, categorization, access) of transaction data relative to traditional measures. Finally, they look ahead to applications in firm-focused research, real-time policy analysis, and macro statistics. The review concludes by positioning household financial transaction data as a flexible, powerful, but imperfect complement to existing measurement tools.",
  "research_design": "A survey/review article that synthesizes the literature using household financial transaction micro-data (from banks, FinTech aggregators, payment intermediaries, and card companies), rather than conducting original empirical estimation. It categorizes applications across household finance and macroeconomics and assesses the benefits and limitations of transaction data relative to survey, scanner, and tax-based measures. There is no research-design shock or dataset of the authors' own.",
  "categories": [
   "Household Finance",
   "Data & Measurement",
   "Macroeconomics"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": null,
  "missing_notes": "Review/survey article; it surveys the class of household financial transaction data (from banks, FinTech apps, and card companies) rather than analyzing a specific dataset of its own.",
  "orig_filename": "Annual Review of Economics - 2022 - Baker & Kueng - Household Financial Transaction Data.pdf",
  "std_name": "Annual Review of Economics - 2022 - Baker and Kueng - Household Financial Transaction Data"
 },
 {
  "journal": "Econometrica",
  "is_working_paper": false,
  "year": 2023,
  "authors": [
   "Eeckhout",
   "Veldkamp"
  ],
  "title": "Data and Markups: A Macro-Finance Perspective",
  "summary": "The paper asks how to measure the extent to which data-intensive firms use market power, and argues that markups—the usual proxy—can be misleading. Building a model in which firms price risk in their capital-allocation and production decisions, the authors show data has competing effects on markups, so markups are an unreliable measure of data-derived market power. Instead, they show that comparing markups measured at different levels of aggregation (product, firm, industry) reveals the role of data and distinguishes data from other intangibles. This reconciles seemingly contradictory empirical markup findings across aggregation levels. The paper offers a new way to measure data and its effects on competition.",
  "logical_flow": "The paper enters the debate over rising market power and whether the dominance of large, data-intensive firms reflects economies of scale in information that reduce competition. It asks how to measure such data-derived market power and questions the standard reliance on markups. The authors build a model drawing on macro theory (data as information that reduces uncertainty), corporate finance (firms price risk), and industrial organization (firms exploit market power), in which economies of scale in data lead a data-rich firm to lower its marginal cost and capture more market share. Working through the model, they show data exerts competing effects on markups, and that the net effect depends on the level of aggregation at which markups are measured, so a single markup number is an unreliable gauge of data-driven power. Crucially, they show the difference between markups measured at the firm versus product level acts as a sufficient statistic for the amount of relevant data. They demonstrate that this framework reconciles existing empirical findings that product-, firm-, and industry-level markups behave differently over the cycle and trend. The paper concludes that examining how markups differ across aggregation levels is the right way to detect and measure the effects of data on markets.",
  "research_design": "A theoretical macro-finance model in which firms price risk in capital-allocation and production decisions and economies of scale in data lower marginal cost and raise market share. The analysis derives how data affects markups differently across levels of aggregation (product, firm, industry) and shows the firm-versus-product markup gap is a sufficient statistic for data, reconciling existing empirical markup evidence. It is analytical, using existing empirical facts for validation rather than estimating a new dataset.",
  "categories": [
   "Macroeconomics",
   "Industrial Organization",
   "Expectations & Information"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": null,
  "missing_notes": "Theoretical paper; no dataset. It reconciles existing empirical markup evidence (product/firm/industry level) rather than analyzing an original dataset. The supplied PDF is the February 2023 manuscript version of the article published in Econometrica (2023).",
  "orig_filename": "Econometrica - 2023 - Eeckhout, Veldkamp - Data and Markups A Macro-Finance Perspective​.pdf",
  "std_name": "Econometrica - 2023 - Eeckhout and Veldkamp - Data and Markups A Macro-Finance Perspective"
 }
]''')


def git(*a): subprocess.run(["git", *a], cwd=REPO, check=True)

def main():
    conn = kc.get_conn()
    done = {r["file_hash"] for r in conn.execute("SELECT file_hash FROM papers")}
    added = skipped = missing = 0
    for rec in RECORDS:
        rec = dict(rec)
        orig = rec.pop("orig_filename")
        std  = rec["std_name"][:kc.MAX_FILENAME].rstrip(" .") + ".pdf"
        src, dest = REPO / orig, REPO / std
        path = dest if (dest.exists() and not src.exists()) else src   # tolerate already-renamed
        if not path.exists():
            print("MISSING (skip):", orig); missing += 1; continue
        h = kc.sha256_of(path)
        if h in done:
            print("already done:", std[:70]); skipped += 1; continue
        if path == src and dest.resolve() != src.resolve():
            git("mv", "--", orig, std)
        kc.save_paper(conn, rec, h, str(dest), orig)
        done.add(h); added += 1
        print("OK ->", std[:74])
    conn.close()
    kc.build()
    git("add", "-A")
    print(f"\nAdded {added}, skipped {skipped}, missing {missing}. "
          f"Now: git commit -am 'catalog batch' && git push")

if __name__ == "__main__":
    sys.exit(main())
