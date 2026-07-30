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
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Diwan",
   "Eliason",
   "League",
   "Leder-Luis",
   "McDevitt",
   "Roberts"
  ],
  "title": "Competition and Fraud in Health Care",
  "summary": "This paper examines how competition affects fraud when governments procure goods and services from private firms, using Medicare's durable medical equipment (DME) market. The authors argue competition has an ambiguous effect on fraud: it dissipates the rents that attract fraudulent firms, but also squeezes margins so that legitimate firms may exit. Exploiting Medicare's switch from administratively regulated prices to competitive bidding for DME, they show that fraudulent suppliers' cost advantage let them gain market share while legitimate suppliers left the market. The result implies that price competition can perversely concentrate a market in fraudulent providers. The paper connects procurement design to the composition of firms and the prevalence of fraud.",
  "logical_flow": "The paper frames government procurement as a setting where competition lowers prices but may change which firms survive, since fraudulent firms have an artificial cost advantage from not delivering legitimate goods or services. It uses Medicare's DME competitive-bidding reform as a shock that sharply lowered reimbursement, predicting that legitimate suppliers with real costs are squeezed out while fraudulent suppliers persist. Linking supplier-level Medicare claims to fraud-enforcement records, it traces how market shares and exit differ between fraudulent and legitimate firms after the reform. It interprets the resulting shift as evidence that competition alone does not discipline fraud and can even entrench it.",
  "research_design": "Event-study/difference-in-differences around Medicare's switch to competitive bidding for durable medical equipment, using supplier-level claims linked to fraud-enforcement outcomes to compare fraudulent and legitimate firms.",
  "categories": [
   "Health Economics",
   "Industrial Organization & Procurement",
   "Fraud & Enforcement"
  ],
  "datasets": [
   {
    "provider": "Centers for Medicare & Medicaid Services (CMS)",
    "product": "Medicare DMEPOS claims",
    "description": "Supplier- (NPI-) level Medicare durable medical equipment claims used to measure market shares, entry, and exit around the competitive-bidding reform.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "Medicare; DME; claims; suppliers"
   },
   {
    "provider": "DOJ / HHS-OIG",
    "product": "Fraud enforcement records",
    "description": "Federal fraud enforcement actions - indictments, civil settlements, and program exclusions - used to identify fraudulent suppliers.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "fraud enforcement; exclusions; indictments"
   },
   {
    "provider": "Yunan Ji",
    "product": "Medicare DME auction bids",
    "description": "Bids submitted in Medicare's DME competitive-bidding auctions, shared by the author, used in the bidding analysis.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "auction bids; competitive bidding"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null,
  "journal": null,
  "std_name": "Working Paper - 2026 - Diwan et al. - Competition and Fraud in Health Care",
  "orig_filename": "w34802.pdf"
 },
 {
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Dyck",
   "Fang",
   "Hebert",
   "Xu"
  ],
  "title": "Venture Fraud",
  "summary": "This paper assembles the first comprehensive sample of venture fraud - 614 U.S. venture-capital-backed startups founded since 2000 that faced fraud allegations - and studies its governance roots. Within a high-detection subsample of newly public firms, VC-backed firms are 54% more likely to face fraud charges than comparable non-VC-backed firms. Fraud is more likely where founders hold stronger control rights, more convex cash-flow rights, more investors, and more non-traditional investors, with founder-controlled boards 88% more likely to commit fraud than VC- or shared-control boards. Governance features predict fraud far better than founder characteristics, and hot funding conditions at the first round - which weaken governance - forecast later fraud. The paper reframes venture fraud as substantially a product of governance structures rather than individual bad actors.",
  "logical_flow": "The paper begins by building a large hand-collected sample of venture-fraud cases, addressing the measurement gap that has limited prior work. It first establishes that VC backing is associated with more fraud in a setting where detection is high and roughly uniform, ruling out pure detection differences. It then turns to mechanisms, arguing that founder-friendly structures and cap-table complexity weaken the monitoring that should deter fraud. Using a panel prediction model, it shows governance variables - founder control, convex payoffs, investor composition - dominate founder traits, and that hot-market financing conditions that erode governance at inception predict future fraud.",
  "research_design": "Hand-collected case-sample construction plus a matched comparison (VC vs. non-VC among newly public firms) and a hazard-style panel prediction model relating fraud to governance and financing-market conditions.",
  "categories": [
   "Venture Capital",
   "Corporate Governance",
   "Financial Fraud"
  ],
  "datasets": [
   {
    "provider": "PitchBook",
    "product": null,
    "description": "Venture-capital-backed startups, their founders, financing rounds, and cap-table/board structures, used to build the sample and governance variables.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "venture capital; startups; cap tables"
   },
   {
    "provider": "Crunchbase",
    "product": null,
    "description": "Startup and news-feed data used alongside PitchBook to identify fraud allegations.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "startups; news; allegations"
   },
   {
    "provider": "Dyck, Fang, Hebert & Xu (this paper)",
    "product": "Venture-fraud case sample",
    "description": "Hand-collected sample of 614 U.S. VC-backed startups facing fraud charges (SEC, DOJ, and civil litigation) since 2000, with fraud commitment periods.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "venture fraud; hand-collected; litigation"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null,
  "journal": null,
  "std_name": "Working Paper - 2026 - Dyck et al. - Venture Fraud",
  "orig_filename": "w34868.pdf"
 },
 {
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Amiti",
   "Kashyap",
   "Kovner",
   "Weinstein"
  ],
  "title": "Why Do Firms Pay Different Interest Rates on Their Bank Loans?",
  "summary": "This paper documents large dispersion in interest rates on similar commercial and industrial loans using confidential supervisory data on the largest U.S. banks, and argues the spread is not explained by risk. The authors rationalize the dispersion with a search-cost model in which borrowers pay to solicit competing quotes. Estimated search costs are highest for smaller and riskier borrowers and lowest for public firms, matching predictable differences in screening and monitoring costs. The search costs are economically large: over a third of firms behave as if they never comparison-shop, half appear to obtain only two quotes, and the rest search widely. The paper interprets loan-rate dispersion as evidence of substantial search frictions in credit markets rather than pure risk pricing.",
  "logical_flow": "The paper starts from the puzzle that observably similar borrowers pay very different loan rates even after accounting for risk, suggesting frictions beyond credit quality. It proposes a search-cost model in which firms trade off the cost of gathering additional quotes against the rate savings, generating dispersion that depends on borrower type. Using confidential loan-level supervisory data, it strips out risk and attributes residual dispersion to search costs, which it estimates structurally. The estimated pattern - higher search costs for small, risky, private borrowers - lines up with the economics of screening and monitoring, and implies many firms effectively fail to shop around.",
  "research_design": "Structural estimation of a borrower search-cost model on confidential loan-level supervisory (FR Y-14) data, after controlling for risk, to recover the distribution of search costs across borrower types.",
  "categories": [
   "Banking",
   "Credit Markets",
   "Search Frictions"
  ],
  "datasets": [
   {
    "provider": "Federal Reserve",
    "product": "FR Y-14 supervisory data",
    "description": "Confidential loan-level data on commercial and industrial loans at the largest U.S. banks, collected for Dodd-Frank stress testing, used to measure rate dispersion and borrower/loan characteristics.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "supervisory; C&I loans; bank stress test"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null,
  "journal": null,
  "std_name": "Working Paper - 2026 - Amiti et al. - Why Do Firms Pay Different Interest Rates on Their Bank Loans",
  "orig_filename": "w34870.pdf"
 },
 {
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Eisfeldt",
   "Hartman-Glaser",
   "Kim",
   "Lee"
  ],
  "title": "Intangible Intensity",
  "summary": "This paper builds a text-based measure of firms' intangible investment intensity from 10-K filings and introduces a general 'semantic theme scoring' (STS) methodology. The approach further decomposes disclosure text into knowledge, customer, and organization capital. High-intangible-intensity firms are smaller, younger, and invest heavily in R&D and human capital, and the three subcomponents map to distinct firm types - knowledge-intensive (R&D-driven, high valuation, skilled labor), customer-intensive (mature, profitable, commercial), and organization-intensive (large, asset-heavy incumbents). The authors show managerial expenditure descriptions carry informative signals about intangible capital that complement financial statements. The paper offers both a new measure and a reusable NLP methodology for capturing corporate intangibles.",
  "logical_flow": "The paper motivates that intangible capital is poorly captured by accounting statements, so it turns to the language firms use to describe their spending. It develops semantic theme scoring to convert 10-K disclosure text into an intangible-intensity measure and to separate knowledge, customer, and organization capital. It validates the measure by showing high-intangible firms have the expected profile and that the three components correspond to economically distinct firm types. It argues these text-derived signals complement financial data, opening avenues for measuring firms' position in the intangible lifecycle.",
  "research_design": "Text-as-data measurement: a semantic-theme-scoring (NLP) methodology applied to 10-K filings to construct and validate an intangible-intensity measure and its subcomponents, with cross-sectional characterization.",
  "categories": [
   "Intangible Capital",
   "Text-as-Data / NLP",
   "Corporate Finance"
  ],
  "datasets": [
   {
    "provider": "SEC (WRDS SEC Analytics Suite)",
    "product": "10-K filings",
    "description": "Full-text corporate 10-K filings, accessed via the WRDS SEC Analytics Suite, used to build the text-based intangible-intensity measure.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "10-K; disclosure text; NLP"
   },
   {
    "provider": "S&P Capital IQ",
    "product": null,
    "description": "Supplementary financial data used to improve customer-capital expense measures.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "financials; customer capital"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null,
  "journal": null,
  "std_name": "Working Paper - 2026 - Eisfeldt et al. - Intangible Intensity",
  "orig_filename": "w34882.pdf"
 },
 {
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Baslandze",
   "Edwards",
   "Graham",
   "McClure",
   "Meyer",
   "Sparks",
   "Waddell",
   "Weitz"
  ],
  "title": "Artificial Intelligence, Productivity, and the Workforce: Evidence from Corporate Executives",
  "summary": "This paper uses a novel survey of nearly 750 corporate executives to study how artificial intelligence is affecting productivity and the workforce. Adoption is highly heterogeneous: more than half of firms have already invested, though many smaller firms are just starting. Labor-productivity gains are positive, vary by sector, are expected to strengthen in 2026, and are concentrated in high-skill services and finance; they reflect higher revenue-based total factor productivity through innovation and demand channels rather than capital deepening. The authors document a 'productivity paradox' in which perceived gains exceed measured gains, likely due to delayed revenue realization. In labor markets they find little near-term aggregate employment decline, though large firms anticipate AI-driven reductions while smaller firms expect modest gains, alongside compositional labor reallocation.",
  "logical_flow": "Because firm-level effects of AI are hard to observe in standard data, the paper gathers direct evidence by surveying corporate executives about adoption, productivity, and workforce plans. It first documents wide heterogeneity in adoption and then characterizes the size, sectoral pattern, and channels of productivity gains, attributing them to revenue-based TFP rather than capital deepening. It reconciles optimistic perceptions with smaller measured effects through a productivity paradox driven by lagged revenue. Finally it turns to labor, showing limited aggregate employment effects so far but divergent expectations and reallocation across firm size and worker types.",
  "research_design": "Survey-based descriptive and cross-sectional analysis of ~750 corporate executives, with comparisons of adoption, productivity, and workforce outcomes across firm size and sector.",
  "categories": [
   "Economics of AI",
   "Productivity",
   "Labor Economics"
  ],
  "datasets": [
   {
    "provider": "Federal Reserve Bank of Atlanta & Duke University",
    "product": "Survey of corporate executives",
    "description": "Novel survey of nearly 750 corporate executives on AI adoption, productivity effects, and workforce plans, benchmarked to the U.S. Census industry/size distribution.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "executive survey; AI adoption; productivity"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null,
  "journal": null,
  "std_name": "Working Paper - 2026 - Baslandze et al. - Artificial Intelligence, Productivity, and the Workforce Evidence from Corporate Executives",
  "orig_filename": "w34984.pdf"
 },
 {
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Matvos",
   "Piskorski",
   "Seru"
  ],
  "title": "Private Credit, Balance Sheets and Financial Stability",
  "summary": "This paper uses new, comprehensive fund- and asset-level data covering most of the private-credit industry to assess its balance sheets and financial-stability implications. Private-credit funds are highly capitalized, with equity typically 65-80% of assets - more than six times the capitalization of U.S. banks - and use only moderate debt, largely bank credit lines for liquidity management. Fund lives of 10-12 years exceed the shorter maturities of underlying loans, implying little maturity mismatch, unlike banks that fund long-term assets with short-term callable deposits. Portfolios are diversified across industries, geographies, and strategies, and performance shows positive average net returns with losses borne mainly by equity. The authors conclude that, as currently structured, private-credit funds are conservatively built and unlikely to pose bank-like systemic risks.",
  "logical_flow": "The paper addresses the financial-stability debate around private credit by bringing comprehensive fund- and asset-level data to bear on how these funds are actually capitalized and funded. It documents high equity ratios and limited, liquidity-oriented debt, contrasting this with banks' thin capital and deposit funding. It then shows the maturity structure implies little mismatch and that portfolios are diversified, limiting correlated-shock exposure. Finally it examines performance and loss-bearing, arguing the equity-heavy structure absorbs losses, so current private-credit configurations are unlikely to generate bank-style systemic risk.",
  "research_design": "Descriptive analysis of new proprietary fund- and asset-level private-credit data (~1,300 funds), benchmarked against U.S. commercial banks using regulatory data.",
  "categories": [
   "Private Credit",
   "Financial Stability",
   "Banking"
  ],
  "datasets": [
   {
    "provider": "Proprietary private-credit data provider",
    "product": null,
    "description": "New proprietary fund- and asset-level data covering roughly 1,300 private-credit funds and their loans (capitalization, funding, maturities, performance), covering most of the industry.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "private credit; funds; balance sheets"
   },
   {
    "provider": "Preqin",
    "product": null,
    "description": "Industry-level private-credit assets used to gauge coverage and situate the sample within the broader market.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "private credit; industry assets"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null,
  "journal": null,
  "std_name": "Working Paper - 2026 - Matvos et al. - Private Credit, Balance Sheets and Financial Stability",
  "orig_filename": "w34991.pdf"
 },
 {
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Pagel",
   "Sridhar",
   "Williams"
  ],
  "title": "Bank Fees and Household Financial Well-Being",
  "summary": "This paper studies 2017-2022 policy changes at large U.S. banks that eliminated non-sufficient-funds (NSF) fees and relaxed overdraft policies, using individual transaction-level data. Eliminating NSF fees immediately reduced NSF charges across the income distribution, as expected. Relaxing overdraft policies, however, reduced overdraft fees only for wealthier (higher-income, higher-liquidity) households, and only they saw subsequent declines in late fees, interest, maintenance fees, and use of alternatives like payday loans. The authors conclude the changes were not substantial enough to meaningfully relieve financial stress for the most vulnerable households. Methodologically, the multi-treatment, varying-intensity setting motivates a new stacked event-study estimator related to de Chaisemartin et al. (2024) to handle staggered-DiD biases.",
  "logical_flow": "The paper asks whether widely publicized reductions in bank fees actually improved household financial well-being, especially for vulnerable households. Using transaction-level data, it separates two distinct policy changes - NSF-fee elimination and overdraft relaxation - and traces their effects across the income and liquidity distribution. It finds broad relief from NSF elimination but overdraft benefits accruing only to wealthier households, who then experience knock-on reductions in other fees and costly credit. Because the setting has multiple staggered treatments of differing intensity, it develops and applies a stacked event-study estimator to obtain unbiased effects.",
  "research_design": "Stacked event-study / staggered difference-in-differences (a new estimator related to de Chaisemartin et al. 2024) on individual transaction-level banking data around NSF/overdraft policy changes.",
  "categories": [
   "Household Finance",
   "Banking",
   "Consumer Protection"
  ],
  "datasets": [
   {
    "provider": "Consumer transaction-data provider",
    "product": null,
    "description": "Individual account- and transaction-level banking data used to measure NSF, overdraft, late, and maintenance fees and use of alternative financial services across households.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "transaction-level; bank fees; households"
   },
   {
    "provider": "FDIC",
    "product": "Summary of Deposits",
    "description": "Branch/deposit data used to determine which metropolitan areas were exposed to each bank's policy change.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "deposits; branch exposure; MSA"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null,
  "journal": null,
  "std_name": "Working Paper - 2026 - Pagel et al. - Bank Fees and Household Financial Well-Being",
  "orig_filename": "w34993.pdf"
 },
 {
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Chen",
   "Sacerdote"
  ],
  "title": "Capital in the Capitol: Congressional Trades Resemble Uninformed Retail Trading",
  "summary": "This paper asks whether members of Congress exploit informational advantages in their personal stock trading, using a new hand-collected dataset of all U.S. legislators' and their families' trades from 2012 to 2023. After the STOCK Act, legislators' portfolios on average underperform or merely match market benchmarks. To explain this mediocre performance, the authors show legislators' positions track financial professionals' recommendations and that their trade timing reflects prevailing market sentiment - estimated from retail investors' social-media posts - rather than anticipating price moves. Congressional trading thus looks like public-signal-following and resembles uninformed retail behavior. The paper pushes back on the view that legislators systematically trade on private information.",
  "logical_flow": "The paper revisits the contested question of congressional trading advantage, noting prior studies used limited samples and said little about mechanisms. It builds a comprehensive hand-collected record of legislator and family trades and evaluates performance against multiple benchmarks, finding no outperformance after the STOCK Act. It then probes why, testing whether trades follow analyst recommendations and align with market sentiment measured from retail social-media activity. Finding that timing mirrors public sentiment rather than leading prices, it concludes legislators' observable trading resembles uninformed retail trading.",
  "research_design": "Event-study and calendar-time portfolio performance analysis on a hand-collected congressional-trading dataset, plus mechanism tests linking trades to analyst recommendations and social-media sentiment.",
  "categories": [
   "Political Economy",
   "Informed Trading",
   "Behavioral Finance"
  ],
  "datasets": [
   {
    "provider": "Chen & Sacerdote (this paper)",
    "product": "Congressional trading records",
    "description": "Hand-collected transactions of all U.S. members of Congress and their immediate families, 2012-2023, from periodic financial-disclosure (STOCK Act) filings.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "congressional trades; disclosures; hand-collected"
   },
   {
    "provider": "Context Analytics",
    "product": "S-Score (Twitter/X sentiment)",
    "description": "Security-level social-media sentiment scores derived from the Twitter/X firehose via proprietary NLP, used to proxy retail market sentiment.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "social media sentiment; retail; NLP"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null,
  "journal": null,
  "std_name": "Working Paper - 2026 - Chen and Sacerdote - Capital in the Capitol Congressional Trades Resemble Uninformed Retail Trading",
  "orig_filename": "w35041.pdf"
 },
 {
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Drechsler",
   "Savov",
   "Schnabl"
  ],
  "title": "Credit Crunches and the Great Stagflation",
  "summary": "This paper argues that severe banking-system credit crunches helped cause the Great Stagflation of the 1970s, working through Regulation Q, which capped deposit rates. Under Reg Q, Fed tightening triggered large deposit outflows that forced banks to contract lending, and these credit crunches line up closely with stagflation over time. Adding Reg Q to a standard model where firms finance working capital with bank loans, the authors show binding caps make working capital costlier, leading firms to raise prices and cut output. This yields an augmented Phillips curve in which monetary tightening reduces aggregate supply as well as demand, with effects increasing in credit-crunch severity, external-finance dependence, and working-capital intensity. Cross-industry tests confirm that more exposed manufacturing industries raised prices and cut output relative to others.",
  "logical_flow": "The paper links a monetary-institutional feature - Regulation Q deposit-rate caps - to the macro puzzle of simultaneous high inflation and stagnation. It explains how Fed tightening under Reg Q caused deposit outflows and lending contractions, then documents that these credit crunches track stagflation in the time series. It embeds Reg Q in a working-capital model, deriving an augmented Phillips curve where tighter credit raises firms' costs and shifts aggregate supply inward. It then tests the model's cross-sectional predictions across manufacturing industries by exposure, external-finance dependence, and working-capital intensity.",
  "research_design": "A quantitative macro model (working-capital channel with Regulation Q) combined with time-series evidence and cross-industry exposure-based empirical tests of its supply-side predictions.",
  "categories": [
   "Monetary Economics",
   "Banking",
   "Macro-Finance"
  ],
  "datasets": [
   {
    "provider": "U.S. Call Reports (via FOIA)",
    "product": null,
    "description": "Bank-level regulatory balance-sheet data obtained back to 1959 through a Freedom of Information Act request, used to measure deposits and lending contractions.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "call reports; bank lending; 1959-"
   },
   {
    "provider": "NBER-CES",
    "product": "Manufacturing Industry Database",
    "description": "Industry-level prices and quantities (derived from the U.S. Census) used to test cross-industry supply effects.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "manufacturing; prices; quantities"
   },
   {
    "provider": "FDIC",
    "product": "Historical Bank Data",
    "description": "Aggregate historical deposits and deposit interest expense used to characterize Reg Q-era bank funding.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "historical deposits; Reg Q"
   },
   {
    "provider": "Federal Reserve",
    "product": "Senior Loan Officer Opinion Survey (SLOOS)",
    "description": "Bank lending-standards survey (back to 1964) used as a measure of credit tightening.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "lending standards; SLOOS"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null,
  "journal": null,
  "std_name": "Working Paper - 2026 - Drechsler et al. - Credit Crunches and the Great Stagflation",
  "orig_filename": "w35057.pdf"
 },
 {
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Egan",
   "Matvos",
   "Seru",
   "Wang",
   "Yao"
  ],
  "title": "Who Pays for Payments?",
  "summary": "This paper uses novel data on the composition and cost of payments across U.S. merchants to quantify redistribution in the payment system. Card interchange fees fund consumer rewards, and when merchants raise prices for everyone, users of low-cost methods like cash and debit cross-subsidize high-reward credit-card users at the same merchant. The authors show incidence actually depends on the joint distribution of payment choices across merchants, not the standard assumption that everyone shops at the same stores facing the same fees. Two forces shape redistribution: consumer sorting, where users of different payment methods shop at different merchants, limits cash and debit users' exposure to high interchange fees; and interchange fees vary across merchants, being lower where payment types overlap (e.g., large grocery stores) due to sector discounts and negotiations. Embedding these forces in a sufficient-statistics framework, the paper revises who ultimately bears payment-system costs.",
  "logical_flow": "The paper reconsiders the common claim that low-reward payers subsidize high-reward card users, noting it assumes uniform shopping and fees. Using merchant-level payment data, it documents two overlooked forces - consumer sorting across merchants and cross-merchant variation in interchange fees. It shows sorting insulates cash and debit users from high-fee exposure, while fee variation further shapes overlap points. It then embeds these forces in a sufficient-statistics incidence framework to recompute the direction and magnitude of cross-subsidies in the payment system.",
  "research_design": "A sufficient-statistics incidence framework estimated on proprietary merchant-acquirer payment data, combined with consumer payment-diary data, to quantify cross-subsidies under consumer sorting and merchant-level fee variation.",
  "categories": [
   "Payments",
   "Household Finance",
   "Industrial Organization"
  ],
  "datasets": [
   {
    "provider": "Fiserv",
    "product": null,
    "description": "Two proprietary datasets from Fiserv, a large U.S. merchant acquirer, on the composition and cost (interchange/settlement) of payments across merchants.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "merchant acquiring; interchange; settlement"
   },
   {
    "provider": "Federal Reserve Bank of Atlanta",
    "product": "Diary of Consumer Payment Choice (DCPC)",
    "description": "Consumer payment-diary data on payment-method use and shopping, used to measure sorting across merchant sectors.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "consumer payments; diary; sorting"
   },
   {
    "provider": "Nilson Report",
    "product": null,
    "description": "Industry statistics on interchange and network fees used to benchmark payment costs.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "interchange; network fees; benchmarks"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null,
  "journal": null,
  "std_name": "Working Paper - 2026 - Egan et al. - Who Pays for Payments",
  "orig_filename": "w35067.pdf"
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
