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
  "orig_filename": "w35095.pdf",
  "std_name": "Working Paper - 2026 - He et al. - Homemade Foreign Trading",
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "He",
   "Wang",
   "Zhu"
  ],
  "title": "Homemade Foreign Trading",
  "summary": "This paper documents how Chinese mainland insiders disguise domestic trades as foreign investment by round-tripping capital through the Stock Connect program, a practice the authors term 'homemade foreign trading.' Using cross-border holding data from all custodians in Stock Connect, they show that northbound flows predict returns and correlate with insider trading, but that this predictability decays after the 2018 Northbound Investor Identification reform introduced see-through surveillance. The decay is concentrated among less prestigious foreign custodians and cross-operating mainland custodians, precisely where insiders can most easily hide. The reform also reduces price informativeness in stocks most exposed to homemade foreign investors. The paper highlights regulatory cooperation as a key ingredient for the integrity of cross-border capital market integration.",
  "logical_flow": "The paper starts from the observation that Stock Connect lets mainland capital exit and re-enter disguised as foreign 'northbound' flow, giving insiders a channel to trade on private information while appearing to be offshore investors. It argues that if such round-tripping is real, northbound flows should predict returns and track insider activity, and that a reform imposing investor-level identification should break this link by making insiders identifiable. Exploiting the 2018 Northbound Investor Identification reform as a shock to surveillance, the authors trace how return predictability and its correlation with insider trading decay afterward. They then sharpen identification by showing the effect is strongest where hiding is easiest (opaque custodians) and by documenting a decline in price informativeness for the most exposed stocks.",
  "research_design": "Event-study / difference-in-differences around the 2018 Northbound Investor Identification reform, using custodian-level cross-border holding data and return-predictability regressions, with cross-sectional heterogeneity by custodian type and stock exposure for identification.",
  "categories": [
   "International Finance",
   "Insider Trading",
   "Market Microstructure"
  ],
  "datasets": [
   {
    "provider": "China Securities Depository and Clearing (CSDC) / Stock Connect",
    "product": "",
    "description": "Cross-border holding and order data from all custodians in China's Stock Connect (northbound), including the 2018 Northbound Investor Identification reform, used to trace round-tripping insider flows.",
    "access_type": "Restricted",
    "delivery": "",
    "topic_tags": "China; Stock Connect; custodians; cross-border holdings"
   },
   {
    "provider": "CSMAR",
    "product": "China Stock Market & Accounting Research",
    "description": "Chinese firms' financial statements and corporate announcements used for firm characteristics and insider-trading events.",
    "access_type": "Proprietary",
    "delivery": "",
    "topic_tags": "China; accounting; announcements"
   },
   {
    "provider": "WIND",
    "product": "",
    "description": "Adjusted opening prices and free-floating shares for Chinese listed stocks.",
    "access_type": "Proprietary",
    "delivery": "",
    "topic_tags": "China; prices; shares"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null
 },
 {
  "orig_filename": "w35206.pdf",
  "std_name": "Working Paper - 2026 - Dou et al. - The Cost of Intermediary Market Power for Distressed Borrowers",
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Dou",
   "Wang",
   "Wang"
  ],
  "title": "The Cost of Intermediary Market Power for Distressed Borrowers",
  "summary": "This paper measures how much lender market power raises borrowing costs for firms in financial distress, who must raise urgent financing in thin, concentrated credit markets. The authors show that distressed borrowers pay very high loan spreads even after stripping out compensation for credit risk, liquidity risk, and ordinary loan-making costs. To decompose the residual, they build and estimate a dynamic game-theoretic model of distressed lending featuring latent demand heterogeneity, endogenous lender participation, creditor blocking power, and tacit collusion sustained by repeated syndication. Lender market power explains 533 basis points of risk-adjusted spreads in debtor-in-possession (DIP) loans and 300 basis points in highly speculative loans, with roughly 140 basis points from tacit collusion in each. Because this markup reduces survival-critical liquidity by 16–20%, market power emerges as a major, previously underappreciated source of financial-distress costs.",
  "logical_flow": "The argument begins by noting that distressed firms borrow in markets dominated by a few specialized lenders and by existing creditors who hold blocking power, so observed spreads may reflect market power rather than pure risk. The authors first document that DIP and highly speculative loan spreads remain large after removing risk and cost components, motivating a structural explanation. They then specify a dynamic game in which lenders choose whether to participate, exploit blocking power, and sustain tacit collusion through repeated syndication, allowing the markup to be identified and decomposed. Estimating the model on facility-level loan data, they quantify the market-power markup and trace its consequences for borrower liquidity and asset-value destruction.",
  "research_design": "Structural estimation of a dynamic game-theoretic model of distressed lending (with endogenous participation, creditor blocking power, and tacit collusion), disciplined by facility-level DIP and highly speculative loan data, used to decompose risk-adjusted spreads into market-power components.",
  "categories": [
   "Financial Distress & Bankruptcy",
   "Banking & Lending",
   "Market Power / Industrial Organization"
  ],
  "datasets": [
   {
    "provider": "Refinitiv LPC",
    "product": "DealScan",
    "description": "Loan-level covenant terms (e.g., debt-to-EBITDA covenants) used to identify firms subject to earnings-based constraints. Facility-level syndicated loan terms used to build the sample of debtor-in-possession and highly speculative loans.",
    "access_type": "Proprietary",
    "delivery": "",
    "topic_tags": "syndicated loans; DIP; facilities; loan covenants; debt-to-EBITDA"
   },
   {
    "provider": "LPC (Refinitiv)",
    "product": "LSTA/LPC Mark-to-Market Pricing Data",
    "description": "Secondary-market loan trading/pricing data used to measure loan liquidity and pricing.",
    "access_type": "Proprietary",
    "delivery": "",
    "topic_tags": "secondary loan prices; liquidity"
   },
   {
    "provider": "PACER",
    "product": "Public Access to Court Electronic Records",
    "description": "U.S. bankruptcy court records used to identify DIP financing and Chapter 11 cases.",
    "access_type": "Public",
    "delivery": "",
    "topic_tags": "bankruptcy; DIP; court records"
   },
   {
    "provider": "Moody's",
    "product": "Default and Recovery Database",
    "description": "Nominal recovery rates aggregated by industry, used in the risk-adjustment of spreads.",
    "access_type": "Proprietary",
    "delivery": "",
    "topic_tags": "defaults; recoveries"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null
 },
 {
  "orig_filename": "w35213.pdf",
  "std_name": "Working Paper - 2026 - Ebsim et al. - Sophisticated Borrowing Constraints and Macroeconomic Dynamics",
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Ebsim",
   "Lian",
   "Ma",
   "Ottonello",
   "Perez"
  ],
  "title": "Sophisticated Borrowing Constraints and Macroeconomic Dynamics",
  "summary": "This paper argues that the way debt constraints are modeled fundamentally changes their macroeconomic implications. Traditional models impose hard borrowing limits that force firms to cut borrowing and investment whenever adverse shocks bind, producing financial acceleration. The authors instead model 'sophisticated borrowing constraints' resembling real-world financial covenants, where breaching an earnings-based debt threshold transfers control rights to creditors who then act to maximize their own value rather than mechanically shrinking credit. Calibrated to micro evidence on investment and earnings around covenant violations, the model matches firm-level dynamics while, at the macro level, generating no financial acceleration because violations do not trigger indiscriminate downscaling. The result challenges a workhorse mechanism linking credit conditions to aggregate fluctuations.",
  "logical_flow": "The paper contrasts textbook hard constraints, under which binding limits mechanically force deleveraging and amplify shocks, with the institutional reality of covenants, which reallocate control rather than dictate fixed borrowing ratios. It models covenant violation as a state-contingent transfer of control to creditors who choose firm policies to maximize creditor value, so tightening need not cause uniform credit cuts. The authors first show the model reproduces the empirical behavior of investment and earnings around violations at the micro level. They then embed the mechanism in a general-equilibrium setting and demonstrate that, unlike hard constraints, sophisticated constraints do not generate financial acceleration.",
  "research_design": "A quantitative macro-finance model with state-contingent, covenant-like borrowing constraints and creditor control rights, calibrated to firm-level dynamics around covenant violations, then evaluated for its aggregate (financial-acceleration) implications.",
  "categories": [
   "Macro-Finance",
   "Corporate Debt & Covenants",
   "Firm Investment"
  ],
  "datasets": [
   {
    "provider": "Refinitiv LPC",
    "product": "DealScan",
    "description": "Loan-level covenant terms (e.g., debt-to-EBITDA covenants) used to identify firms subject to earnings-based constraints. Facility-level syndicated loan terms used to build the sample of debtor-in-possession and highly speculative loans.",
    "access_type": "Proprietary",
    "delivery": "",
    "topic_tags": "syndicated loans; DIP; facilities; loan covenants; debt-to-EBITDA"
   },
   {
    "provider": "Greg Nini (Nini, Smith & Sufi)",
    "product": "Financial covenant violation data",
    "description": "Hand-shared extended data on financial covenant violations disclosed in firms' filings, 1996 onward, used to construct micro moments around violations.",
    "access_type": "Restricted",
    "delivery": "",
    "topic_tags": "covenant violations; control rights"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null
 },
 {
  "orig_filename": "w35227.pdf",
  "std_name": "Working Paper - 2026 - Clayton and Coppola - The Optimal Use of AI in Financial Regulation",
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Clayton",
   "Coppola"
  ],
  "title": "The Optimal Use of AI in Financial Regulation",
  "summary": "This paper asks whether modern AI methods applied to granular portfolio-holdings data can improve macroprudential regulation, and how policymakers should use such tools. The authors build a graph-based deep learning model over security-level holdings of financial intermediaries that embeds economic priors and learns latent representations of both assets and investors from the network of positions. Applied to the near-universe of non-bank financial intermediaries (about $40 trillion), the model forecasts intermediary trading far better than existing approaches, including during crises, and has more than ten times the explanatory power for cross-sectional stress-period returns. Its learned embeddings encode economically interpretable information about fire-sale vulnerability, and the architecture is inductive enough to produce estimates even for withheld asset classes or investors. The authors then embed the empirical model in a macroprudential optimal-policy framework, addressing the Lucas critique, to show why these objects matter for welfare.",
  "logical_flow": "The paper frames systemic risk measurement as a prediction problem over the network of who holds what, arguing that the structure of intermediary holdings encodes fire-sale vulnerability that standard metrics miss. It develops a graph neural network with built-in economic priors that learns asset and investor representations from holdings, then validates the model out-of-sample on trading behavior and stress-period returns against existing benchmarks. Having shown the learned objects are predictive and interpretable, it asks how a regulator should optimally use them. It closes by embedding the empirical model in an equilibrium optimal-policy framework that confronts the Lucas critique, formalizing the welfare rationale for AI-based supervision.",
  "research_design": "Design and out-of-sample evaluation of a graph-based (network) deep-learning model on security-level intermediary holdings, benchmarked against existing systemic-risk metrics, embedded within a macroprudential optimal-policy framework robust to the Lucas critique.",
  "categories": [
   "Machine Learning in Finance",
   "Financial Regulation",
   "Systemic Risk"
  ],
  "datasets": [
   {
    "provider": "FactSet & Morningstar",
    "product": "Security-level intermediary holdings",
    "description": "Holdings data covering the universe of non-bank financial intermediaries (mutual funds, ETFs, separate accounts, etc.), ~$40 trillion, used to train the holdings-network model.",
    "access_type": "Proprietary",
    "delivery": "",
    "topic_tags": "holdings network; non-bank intermediaries; fire sales"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null
 },
 {
  "orig_filename": "w35228.pdf",
  "std_name": "Working Paper - 2026 - Ahern - Industrial Concentration, Property Values, and Municipal Bond Spreads",
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Ahern"
  ],
  "title": "Industrial Concentration, Property Values, and Municipal Bond Spreads",
  "summary": "This paper shows that the industrial composition of a city's economy shapes its cost of municipal borrowing. In a panel of 1,177 U.S. cities from 2005 to 2022, greater sectoral concentration raises default risk and municipal bond spreads, especially where dominant industries are associated with low property values. Instrumental variables built from national sector-employment trends and regional house-price variation support a causal reading. A calibrated city-default model implies the estimated spread effect understates the gross risk from concentration, because concentration also brings agglomeration benefits that lower spreads, particularly in high-property-value cities. The paper connects local economic structure to public finance and credit risk.",
  "logical_flow": "The paper reasons that a city dependent on few industries faces more concentrated economic risk, which should raise default probability and hence bond spreads, but that the effect may depend on whether those industries support high property values that anchor the tax base. It documents the concentration–spread relationship in a large city panel and shows it is stronger where dominant sectors imply low property values. To move from correlation to causation, it instruments concentration using national sector-employment trends interacted with regional house-price variation. Finally, a calibrated default model reconciles the estimates with offsetting agglomeration benefits, implying the gross risk effect is larger than the net spread effect suggests.",
  "research_design": "Panel regressions with an instrumental-variables strategy (national sector-employment trends and regional house-price variation) linking industrial concentration to municipal bond spreads, complemented by a calibrated structural model of city default.",
  "categories": [
   "Municipal Finance",
   "Credit Risk",
   "Urban & Regional Economics"
  ],
  "datasets": [
   {
    "provider": "MSRB",
    "product": "Municipal Securities Rulemaking Board trade data",
    "description": "Secondary-market municipal bond trades used to construct city-level bond spreads.",
    "access_type": "Public",
    "delivery": "",
    "topic_tags": "municipal bonds; spreads; secondary market"
   },
   {
    "provider": "U.S. Census Bureau",
    "product": "LODES (LEHD Origin-Destination Employment Statistics)",
    "description": "Sector-level job counts used to measure a city's industrial concentration.",
    "access_type": "Public",
    "delivery": "",
    "topic_tags": "employment; industry concentration; LODES"
   },
   {
    "provider": "FHFA",
    "product": "House Price Index",
    "description": "Regional house-price data (pre-sample 1975-2004) used to build instruments and proxy property values.",
    "access_type": "Public",
    "delivery": "",
    "topic_tags": "house prices; property values"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null
 },
 {
  "orig_filename": "w35286.pdf",
  "std_name": "Working Paper - 2026 - Thesmar and Verner - Beliefs and Stock Market Fluctuations New Evidence from the Past Seven Decades",
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Thesmar",
   "Verner"
  ],
  "title": "Beliefs and Stock Market Fluctuations: New Evidence from the Past Seven Decades",
  "summary": "This paper builds a new seven-decade series of subjective expected equity returns (1956-2024) from an independent equity-analysis firm and uses it to reassess how beliefs move markets. Unlike the expectations of individual investors and professional forecasters, this measure is strongly positively correlated with the earnings-price ratio, responds negatively to past returns, and positively predicts future returns. Individual and professional-forecaster expectations are only weakly or negatively correlated with the new series, and disagreement between sophisticated and individual investors is associated with higher trading volume. The evidence fits a model of heterogeneous beliefs in which naive investors extrapolate past returns while sophisticated investors are close to rational. The paper offers a rare long-run window on the belief dynamics behind valuations.",
  "logical_flow": "The paper motivates the need for a long, consistent measure of sophisticated investors' return expectations, which it constructs by hand from decades of an independent equity-analysis firm's forecasts. It characterizes the new series' properties—co-movement with the earnings-price ratio, mean-reverting response to past returns, and predictive power for future returns—and contrasts them with survey measures of individuals and forecasters. Finding that sophisticated and naive expectations diverge, and that their disagreement tracks trading volume, it argues a single-belief model cannot fit the facts. It then shows a heterogeneous-beliefs model, with extrapolative naive investors and near-rational sophisticated investors, rationalizes the joint patterns.",
  "research_design": "Construction and analysis of a novel long-run (1956-2024) subjective-expected-return series hand-collected from an independent equity-analysis firm, with predictive and correlation regressions against survey expectations and a heterogeneous-beliefs asset-pricing model.",
  "categories": [
   "Behavioral Finance",
   "Asset Pricing",
   "Expectations & Beliefs"
  ],
  "datasets": [
   {
    "provider": "Value Line Investment Survey",
    "product": "",
    "description": "Analyst forecasts hand-collected from microfilm records (1956-2024) used to construct a long-run series of sophisticated investors' subjective expected returns.",
    "access_type": "Restricted",
    "delivery": "",
    "topic_tags": "subjective expectations; analyst forecasts; microfilm"
   },
   {
    "provider": "Robert Shiller",
    "product": "Investor Confidence Survey",
    "description": "Investor-confidence survey data used as a comparison measure of investor beliefs.",
    "access_type": "Public",
    "delivery": "",
    "topic_tags": "investor confidence; beliefs"
   },
   {
    "provider": "UBS/Gallup and Livingston Survey",
    "product": "",
    "description": "Survey measures of individual-investor and professional-forecaster return/earnings expectations used for comparison.",
    "access_type": "Public",
    "delivery": "",
    "topic_tags": "investor surveys; forecaster expectations"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null
 },
 {
  "orig_filename": "w35335.pdf",
  "std_name": "Working Paper - 2026 - Bardóczy et al. - Monopsony Power and the Transmission of Monetary Policy",
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Bardóczy",
   "Bornstein",
   "Salgado"
  ],
  "title": "Monopsony Power and the Transmission of Monetary Policy",
  "summary": "This paper studies how employer labor-market power changes the way monetary policy transmits to wages and employment. Using administrative U.S. Census data, the authors show that high-monopsony firms—those accounting for over 10 percent of their local wage bill—adjust their wage bill and employment less in response to monetary policy. They rationalize this with a New Keynesian model featuring heterogeneous firms and oligopsonistic labor competition, where wage stickiness combined with labor-market power drives the muted response. The model isolates two channels—partial passthrough and misallocation—through which oligopsony shapes aggregate effects. Calibrated to U.S. labor markets, it implies that the decline in labor-market power since the 1980s raised the output response to monetary policy by about 10 percent and accounts for roughly 15 percent of the flattening of the Phillips curve.",
  "logical_flow": "The paper posits that firms with labor-market power set wages strategically, so their response to monetary policy may differ systematically from competitive firms. Using administrative data, it establishes that high-monopsony firms adjust wage bill and employment less after monetary shocks. To interpret this, it builds a heterogeneous-firm New Keynesian model with oligopsony in which wage stickiness and market power jointly generate the muted response. The model formalizes partial passthrough and misallocation channels, and its calibration links the secular decline in monopsony to changes in the potency of monetary policy and the slope of the Phillips curve.",
  "research_design": "Empirical estimation of firm-level monetary-policy responses by monopsony intensity using administrative Census data, combined with a calibrated heterogeneous-firm New Keynesian model with oligopsonistic labor markets.",
  "categories": [
   "Monetary Policy",
   "Labor Market Power",
   "Macroeconomics"
  ],
  "datasets": [
   {
    "provider": "U.S. Census Bureau",
    "product": "LEHD merged with Longitudinal Business Database (LBD)",
    "description": "Confidential administrative data on establishment-level wage bill and employment used to measure local monopsony power and firm responses to monetary policy.",
    "access_type": "Restricted",
    "delivery": "",
    "topic_tags": "administrative; wage bill; employment; monopsony"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null
 },
 {
  "orig_filename": "w35388.pdf",
  "std_name": "Working Paper - 2026 - Malenko et al. - Fragmentation of Shareholder Power",
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Malenko",
   "Malenko",
   "Tsoy"
  ],
  "title": "Fragmentation of Shareholder Power",
  "summary": "This paper develops a theoretical framework to evaluate how the asset-management industry's shift toward tailored portfolios, fund proliferation, and decentralized stewardship affects corporate governance. Growing heterogeneity in investor preferences better aligns products with demand but can fragment ownership and weaken managerial oversight. The authors show fund proliferation does not necessarily weaken governance: stronger manager incentives and concentrated portfolios of specialized funds can offset fragmentation, especially when investor preferences are intense. However, intense preferences can also push asset managers to compete by granting investors control—decentralizing stewardship and adopting pass-through voting—without internalizing the resulting governance costs. The paper clarifies when investor-driven customization helps or harms oversight.",
  "logical_flow": "The paper begins from real industry trends—customization, more funds, and decentralized voting—and asks whether they necessarily erode the monitoring that concentrated ownership provides. It builds a model in which asset managers choose portfolio concentration, effort, and how much control to pass through to investors, given heterogeneous investor preferences. Analyzing the model, it shows proliferation can coexist with strong governance when incentives and specialization are strong, particularly under intense preferences. It then identifies a countervailing force: competition on the 'control' margin can lead managers to decentralize stewardship inefficiently, because they do not bear the full governance cost.",
  "research_design": "Theoretical model (mechanism/contracting framework) of asset-manager portfolio choice, incentives, and stewardship decentralization under heterogeneous investor preferences; no external dataset.",
  "categories": [
   "Corporate Governance",
   "Asset Management",
   "Financial Theory"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "missing_notes": "Theoretical paper; no nonstandard datasets identified."
 },
 {
  "orig_filename": "w35479.pdf",
  "std_name": "Working Paper - 2026 - Bornstein and Castillo-Martinez - Firm Exit and Financial Frictions",
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Bornstein",
   "Castillo-Martinez"
  ],
  "title": "Firm Exit and Financial Frictions",
  "summary": "This paper examines whether financial frictions cause too many firms to fail, motivating government interventions that prevent closures during crises. The authors build a firm-dynamics model with incomplete financial markets and show that frictions generate excessive firm exit, with a key sufficient statistic being the marginal propensity to exit with debt. Using confidential U.S. Census data, they estimate the debt–exit relationship and use it to discipline the model. The calibrated model implies that eliminating financial frictions would cut firm exit from 9.3% to 5.0% and deliver welfare gains of 3.6% in consumption-equivalent terms, with costs that spike during financial crises but not during ordinary productivity recessions. Comparing government-guaranteed loans and grants, the paper quantifies the trade-off between fiscal cost and effectiveness in preventing excessive exit.",
  "logical_flow": "The paper starts from the policy fear that viable-but-constrained firms fail in crises, and asks when firm exit is actually inefficient. It develops a firm-dynamics model with incomplete markets in which financial frictions distort the exit margin, isolating the marginal propensity to exit with debt as the statistic governing dynamic inefficiency. It then estimates the debt–exit relationship in confidential Census data to discipline that statistic and calibrate the model. Using the calibrated model, it quantifies the welfare cost of frictions across crisis and non-crisis states and evaluates guaranteed loans versus grants as interventions.",
  "research_design": "A quantitative firm-dynamics model with incomplete financial markets, disciplined by a reduced-form debt–exit relationship estimated on confidential U.S. Census microdata, used for welfare analysis and policy (loan-guarantee vs. grant) comparison.",
  "categories": [
   "Firm Dynamics",
   "Financial Frictions",
   "Macroeconomic Policy"
  ],
  "datasets": [
   {
    "provider": "U.S. Census Bureau",
    "product": "Longitudinal Business Database (LBD)",
    "description": "Confidential data tracking entry and exit for the universe of U.S. employers, used to estimate the debt–exit relationship.",
    "access_type": "Restricted",
    "delivery": "",
    "topic_tags": "firm entry/exit; universe of employers"
   },
   {
    "provider": "Bureau van Dijk",
    "product": "Orbis (Italy)",
    "description": "Italian firm-level data used to replicate the U.S. analysis in a setting with deeper financial frictions.",
    "access_type": "Proprietary",
    "delivery": "",
    "topic_tags": "firm-level; Italy; cross-country"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null
 },
 {
  "orig_filename": "w35504.pdf",
  "std_name": "Working Paper - 2026 - Correia et al. - Bank Runs with and without Bank Failure",
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Correia",
   "Luck",
   "Verner"
  ],
  "title": "Bank Runs with and without Bank Failure",
  "summary": "This paper builds a comprehensive historical database of U.S. bank runs by applying large language models to historical newspapers, yielding information on 3,984 runs on individual banks from 1863 to 1934. The data show runs are far more likely at weak banks but also occur at strong banks, especially following negative news about the real economy or the broader banking system. Crucially, runs typically cause failure only when a bank's fundamentals are poor; strong banks survive through signaling, interbank cooperation, and temporary suspension. Locally, runs on weak banks produce much larger declines in deposits, lending, and manufacturing activity than runs on strong banks. The findings put poor fundamentals at the center of both when runs occur and when they are economically damaging, tempering pure self-fulfilling-panic accounts.",
  "logical_flow": "The paper's first contribution is measurement: it uses LLMs to read historical newspapers and assemble a large, bank-level database of runs spanning 1863-1934. With this data, it asks whether runs strike indiscriminately or track fundamentals, showing runs concentrate in weak banks yet can hit strong banks after adverse news. It then distinguishes runs from failures, documenting that strong banks survive runs via signaling, cooperation, and suspension, so that failure requires poor fundamentals. Finally it links runs to local real outcomes, showing damage is concentrated where fundamentals are weak, which it interprets as evidence against models where small shocks trigger discontinuous jumps to bad equilibria.",
  "research_design": "Construction of a novel bank-level historical dataset via large language models applied to newspaper archives, combined with empirical analysis relating runs to bank fundamentals, survival mechanisms, and local real-economy outcomes.",
  "categories": [
   "Banking History",
   "Financial Crises",
   "Text-as-Data / LLMs"
  ],
  "datasets": [
   {
    "provider": "Correia, Luck & Verner (this paper)",
    "product": "Historical bank-run database",
    "description": "Newly constructed database of 3,984 bank runs on individual U.S. banks, 1863-1934, extracted from historical newspapers using large language models.",
    "access_type": "Public",
    "delivery": "",
    "topic_tags": "bank runs; historical; LLM-constructed"
   },
   {
    "provider": "Library of Congress",
    "product": "Chronicling America",
    "description": "Large-scale archive of digitized historical U.S. newspapers, the main source for detecting and dating bank runs.",
    "access_type": "Public",
    "delivery": "",
    "topic_tags": "historical newspapers; text source"
   },
   {
    "provider": "OCC",
    "product": "National bank call reports (Annual Report to Congress)",
    "description": "Balance-sheet data for national banks used to measure bank fundamentals and outcomes.",
    "access_type": "Public",
    "delivery": "",
    "topic_tags": "call reports; bank balance sheets; fundamentals"
   }
  ],
  "no_nonstandard_datasets": false,
  "missing_notes": null
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
