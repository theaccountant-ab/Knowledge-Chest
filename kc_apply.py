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
  "journal": "Quarterly Journal of Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Acemoglu",
   "Restrepo"
  ],
  "title": "Automation and Rent Dissipation: Implications for Wages, Inequality, and Productivity",
  "summary": "This paper studies automation in a task-based economy where some jobs pay workers rents—wages above their outside options. The authors show that automation disproportionately targets high-rent tasks, dissipating those rents, amplifying wage losses for exposed workers, and compressing wage dispersion within exposed groups. Because dissipating rents is socially inefficient, this channel partly offsets the productivity gains automation would otherwise deliver. The framework reconciles observed wage declines and falling within-group inequality with modest aggregate productivity effects. It reframes part of automation's distributional cost as the destruction of labor rents rather than pure efficiency-enhancing substitution.",
  "logical_flow": "The paper begins from the task-based view of production, in which firms can automate individual tasks, and adds the realistic feature that some tasks pay workers rents above their outside options. It argues that firms have a heightened incentive to automate precisely the high-rent tasks, because doing so lets them capture the rent, so automation is not distributionally neutral but concentrated where labor was best paid. This implies that automation dissipates rents, which amplifies wage losses for exposed workers beyond the standard displacement effect and, because the highest-paid exposed workers lose the most, compresses within-group wage dispersion. The authors then show that this rent dissipation is socially inefficient: the resources spent automating to capture rents do not correspond to genuine productivity gains, so measured productivity rises less than a frictionless model would predict. The model thus links three observations—wage declines, falling within-group inequality in exposed groups, and disappointing aggregate productivity—to a single mechanism. It closes by drawing out the welfare implication that some automation is privately profitable but socially wasteful because it merely redistributes rents from workers to firms.",
  "research_design": "A task-based theoretical model of automation extended so that some tasks earn workers rents, solved to characterize how automation targets high-rent tasks and what that implies for wages, within-group inequality, and productivity. The predictions are illustrated with standard industry and productivity data (for example, BLS multifactor-productivity and capital-share measures) rather than a single natural experiment. The contribution is primarily conceptual: isolating rent dissipation as a distinct channel through which automation lowers wages and inequality while offsetting productivity gains.",
  "categories": [
   "Labor Economics",
   "Automation & Technology",
   "Macroeconomics"
  ],
  "datasets": [
   {
    "provider": "Bureau of Labor Statistics",
    "product": "Total Multifactor Productivity Tables",
    "description": "Industry-level multifactor productivity and capital/automation-share measures used to discipline the model's quantitative illustrations.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "productivity; capital share; industries"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Quarterly Journal of Economics - 2026 - Acemoglu and Restrepo - Automation and Rent Dissipation Implications for Wages, Inequality, and Productivity",
  "orig_filename": "qjag006.pdf"
 },
 {
  "journal": "Quarterly Journal of Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Sangani"
  ],
  "title": "Complete Pass-Through in Levels",
  "summary": "This paper revisits the long-standing finding that cost pass-through to prices is incomplete. Using microdata from gas stations, food products, and manufacturing, the author shows that incomplete pass-through measured in percentages often masks complete pass-through in levels: a one-dollar-per-unit rise in input costs raises downstream prices by about one dollar per unit. The apparent incompleteness arises because percentage measures mix together products with different markups. The paper argues that dollar-for-dollar (level) pass-through is the empirically robust regularity and derives its implications for how cost shocks propagate through supply chains. The result revises a widely used input to models of inflation and market power.",
  "logical_flow": "The paper starts from the robust empirical finding that pass-through of input costs to prices looks incomplete when measured in percentages, even over long horizons, which has shaped models of markups and inflation. The author points out that percentage pass-through conflates the size of a cost change with the level of a product's price and markup, so a common dollar cost increase can look like different percentage pass-through across products. Re-examining the data in levels rather than percentages, the paper finds that a one-dollar increase in per-unit input costs translates into roughly a one-dollar increase in per-unit prices—complete pass-through in levels. It documents this across very different settings (retail gasoline, packaged food, and manufacturing), suggesting the regularity is general rather than sector-specific. The paper then works through the theoretical implications, showing which demand and cost structures generate level pass-through and how this changes the transmission of cost shocks. The upshot is that the 'incomplete pass-through' puzzle largely dissolves once pass-through is measured in the right units.",
  "research_design": "An empirical measurement study that recomputes cost pass-through in levels (dollars per unit) rather than percentages, using price and cost microdata from three distinct settings—retail gasoline, packaged food (retail scanner data), and manufacturing (Census production data). The core comparison contrasts level and percentage pass-through estimates to show the former is complete while the latter appears incomplete because it mixes products with different markups. The analysis is descriptive and cross-sectional across products and industries, complemented by a simple model clarifying when level pass-through holds.",
  "categories": [
   "Industrial Organization & Prices",
   "Macroeconomics",
   "Applied Microeconomics"
  ],
  "datasets": [
   {
    "provider": "NielsenIQ (Nielsen Consumer LLC)",
    "product": "Retail scanner data",
    "description": "Product-level retail prices and quantities for packaged-goods categories, used to measure pass-through in the food-products setting.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "retail scanner; prices; food"
   },
   {
    "provider": "U.S. Census Bureau",
    "product": "Annual Survey of Manufactures / Census of Manufactures",
    "description": "Establishment-level production, cost, and price data used to measure pass-through in manufacturing.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "manufacturing; costs; production"
   },
   {
    "provider": "Retail gasoline price microdata",
    "product": null,
    "description": "Station-level retail gasoline prices and wholesale input costs used to measure pass-through at gas stations.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "gasoline; station prices; pass-through"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Quarterly Journal of Economics - 2026 - Sangani - Complete Pass-Through in Levels",
  "orig_filename": "qjag014.pdf"
 },
 {
  "journal": "Quarterly Journal of Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Cirera",
   "Comin",
   "Cruz"
  ],
  "title": "Technology Sophistication Across Establishments",
  "summary": "This paper introduces a new way to measure technology sophistication by separately capturing the most advanced (MAX) and the most widely used (MOST) technologies in key business functions within establishments. Using a survey of over 21,000 establishments across 15 countries, the authors find that firms generally underutilize the most sophisticated technologies they have available. These MAX–MOST gaps are large and persistent and are strongly associated with lower productivity. The paper shows that raising the intensity of use of already-available advanced technologies is an important and underappreciated margin of technology adoption. It reframes the adoption problem as one of use, not just acquisition.",
  "logical_flow": "The paper argues that standard measures of technology adoption—whether a firm has a technology—miss a crucial distinction between having an advanced technology and actually using it widely. To capture this, the authors design a measurement approach that records, for each key business function, both the most sophisticated technology an establishment can access (MAX) and the technology it uses most intensively (MOST). Applying this to a large multi-country establishment survey, they document that MAX typically exceeds MOST, meaning firms underuse the sophisticated technologies available to them, and that these gaps are widespread across functions and countries. They then show the gaps are persistent rather than transitory and are systematically related to lower productivity, implying the underuse is economically costly. This leads them to argue that the intensity of use of existing technologies is a distinct and important adoption margin, separate from acquiring new technologies. The paper thus shifts the policy question from helping firms obtain technology toward helping them use what they already have.",
  "research_design": "A measurement-and-descriptive study built on a purpose-designed survey instrument that separately elicits the most advanced (MAX) and most widely used (MOST) technologies in each business function of an establishment. Using the resulting data on 21,000+ establishments across 15 countries, the authors construct MAX–MOST gaps and relate them cross-sectionally to firm characteristics and productivity. The design's novelty is in the measurement approach itself; the analysis documents the size, persistence, and productivity correlates of the gaps rather than exploiting an external shock.",
  "categories": [
   "Technology Adoption",
   "Development Economics",
   "Firm Productivity"
  ],
  "datasets": [
   {
    "provider": "World Bank",
    "product": "Firm-level Adoption of Technology (FAT) survey",
    "description": "Novel establishment survey measuring the most advanced and most widely used technologies across business functions for 21,000+ establishments in 15 countries.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "technology adoption; establishments; cross-country survey"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Quarterly Journal of Economics - 2026 - Cirera et al. - Technology Sophistication Across Establishments",
  "orig_filename": "qjag018.pdf"
 },
 {
  "journal": "Quarterly Journal of Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Bostanci",
   "Ordoñez"
  ],
  "title": "Business, Liquidity, and Information Cycles",
  "summary": "This paper studies the dual role of stock markets as both a source of information about firm fundamentals and a source of liquidity, and shows these roles interact. When stocks are used more intensively for liquidity, their prices reveal less information about fundamentals, worsening resource allocation. The authors structurally estimate stock price informativeness for several countries and document that it declines when alternative liquidity sources contract and stocks are used more for liquidity. They embed this mechanism in a model linking business, liquidity, and information cycles. The framework connects financial-market liquidity conditions to the real economy through the informativeness of prices.",
  "logical_flow": "The paper begins by noting that stock markets serve two functions economists usually study separately: they aggregate information that guides real investment, and they provide liquid assets that agents can trade when they need cash. The authors argue these functions are in tension, because trading motivated by liquidity needs is uninformed and, when it grows, dilutes the information content of prices. They formalize a setting in which the intensity of liquidity-motivated trading endogenously reduces how much prices reveal about fundamentals, so liquidity and information become linked. To take this to data, they structurally estimate price informativeness across several countries and show it falls precisely when alternative liquidity sources dry up and stocks are used more for liquidity. This co-movement supports the model's central prediction that liquidity and information cycles are connected. The paper concludes that fluctuations in market liquidity feed into the real economy by changing the informativeness of prices and hence the quality of investment decisions.",
  "research_design": "A theoretical model in which liquidity-motivated trading endogenously lowers stock price informativeness, combined with structural estimation of price informativeness across several countries. The structural estimates are related to measures of alternative liquidity availability to test the model's prediction that informativeness falls when stocks are used more intensively for liquidity. The analysis is cross-country and structural rather than based on a single natural experiment; the unit of analysis is the country (and firm) over time.",
  "categories": [
   "Asset Pricing & Information",
   "Macro-Finance",
   "International Finance"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": null,
  "missing_notes": "Uses standard cross-country data (Worldscope and I/B/E/S via WRDS, plus World Bank GDP); no distinctive/nonstandard datasets identified.",
  "std_name": "Quarterly Journal of Economics - 2026 - Bostanci and Ordoñez - Business, Liquidity, and Information Cycles",
  "orig_filename": "qjag029.pdf"
 },
 {
  "journal": "Review of Economic Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Dempsey"
  ],
  "title": "Capital Requirements with Non-Bank Finance",
  "summary": "This paper quantitatively analyzes the macroeconomic effects of raising bank capital requirements in a model where heterogeneous firms can choose between intermediated (bank) and direct (bond-market) finance. Banks compete with each other and with the bond market, fund loans with insured deposits and costly equity subject to a minimum capital ratio, and monitor borrowers. Tighter capital requirements reduce costly bank failures but, because firms can substitute toward direct finance, have only small effects on aggregate lending and output. The availability of non-bank finance thus dampens the real costs of higher capital requirements. The paper shows that ignoring the bank–bond substitution margin overstates the output cost of bank regulation.",
  "logical_flow": "The paper starts from the policy debate over how costly higher bank capital requirements are for the real economy, noting that standard analyses often hold fixed the set of firms that borrow from banks. It builds a quantitative general-equilibrium model in which heterogeneous firms endogenously choose between bank loans and bond-market finance, and in which banks fund themselves with insured deposits and costly equity and must satisfy a capital requirement. Because firms can substitute toward direct finance when bank credit becomes more expensive, the model creates a margin absent from bank-only frameworks. Raising the capital requirement makes bank funding costlier and reduces bank lending, but much of the affected borrowing migrates to the bond market rather than disappearing, so aggregate lending and output fall only modestly. At the same time, higher capital buffers reduce the incidence of costly bank failures, generating a stability benefit. Weighing these forces, the model implies that accounting for non-bank finance substantially lowers the estimated output cost of tighter capital requirements relative to bank-only models.",
  "research_design": "A quantitative (structurally calibrated) general-equilibrium model with heterogeneous firms that endogenously choose between bank and bond finance and heterogeneous banks subject to capital requirements, deposit funding, and costly equity. The model is calibrated to standard macro-financial targets and used for counterfactual experiments that raise capital requirements and trace the effects on bank failures, the bank–bond financing mix, aggregate lending, and output. There is no external natural experiment; identification of the mechanism comes from the model's structure and the endogenous financing-choice margin.",
  "categories": [
   "Banking",
   "Macro-Finance",
   "Financial Regulation"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": null,
  "missing_notes": "Quantitative model calibrated to standard aggregates; no distinctive datasets.",
  "std_name": "Review of Economic Studies - 2026 - Dempsey - Capital Requirements with Non-Bank Finance",
  "orig_filename": "rdaf061.pdf"
 },
 {
  "journal": "Review of Economic Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Cullen",
   "Li",
   "Perez-Truglia"
  ],
  "title": "What's My Employee Worth? The Effects of Salary Benchmarking",
  "summary": "This paper studies salary benchmarking—firms using aggregated market-salary data to set pay. Using national payroll data, the authors examine firms that gain access to a tool revealing market benchmarks for each job title, and use a difference-in-differences design around that access. They find the benchmark information reduces within-firm salary dispersion by about 25%, as firms move pay toward the market rate for each role. The results show that widely available wage information meaningfully compresses pay differences rather than simply raising or lowering wages. The paper documents a concrete channel through which information affects wage setting and dispersion.",
  "logical_flow": "The paper begins with the observation that firms increasingly use aggregated market-salary data—salary benchmarking—to set pay, and asks how access to such information changes wage-setting. It frames the question around a tool that gives firms market benchmarks for each specific job title, so the relevant information is granular and directly usable in pay decisions. Exploiting the timing of when firms gain access to this tool, the authors set up a difference-in-differences comparison between firms that adopt it and otherwise similar firms that have not yet done so. They predict that benchmarking pulls each firm's pay for a given role toward the market rate, which should compress the spread of salaries within the firm rather than uniformly shift them. The data confirm a roughly 25% reduction in salary dispersion following access, consistent with firms anchoring on role-specific benchmarks. The paper interprets this as evidence that market wage information is a powerful force compressing pay differences, with implications for inequality and for debates over the transparency of compensation data.",
  "research_design": "A difference-in-differences design exploiting the staggered timing with which firms gain access to a salary-benchmarking tool that reveals market pay by job title, estimated on national payroll data. Treated firms (which adopt the tool) are compared to not-yet-treated firms, with outcomes measured as within-firm salary dispersion and pay relative to market benchmarks by O*NET job code. The unit of analysis is the firm (and job title within firm) over time; the identifying variation is the change in information availability at adoption.",
  "categories": [
   "Labor Economics",
   "Personnel Economics",
   "Wage Inequality"
  ],
  "datasets": [
   {
    "provider": "Compensation-software provider (national payroll)",
    "product": null,
    "description": "Large-scale, anonymized national payroll/compensation records with market salary benchmarks by job title (mapped to O*NET codes), used to study firms adopting a salary-benchmarking tool.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "payroll; salaries; benchmarking; O*NET"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Rollout of a salary-benchmarking tool",
   "type": "Staggered access / quasi-experiment",
   "what": "Firms gaining access to a compensation tool that reveals market salary benchmarks by job title, used in a difference-in-differences design as a quasi-exogenous shock to the wage information available to employers."
  },
  "missing_notes": null,
  "std_name": "Review of Economic Studies - 2026 - Cullen et al. - What's My Employee Worth The Effects of Salary Benchmarking",
  "orig_filename": "rdaf083.pdf"
 },
 {
  "journal": "Review of Economic Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Miranda-Agrippino",
   "Hacıoğlu-Hoke",
   "Bluwstein"
  ],
  "title": "Patents, News, and Business Cycles",
  "summary": "This paper builds a new instrument for technology news shocks—anticipated future improvements in productivity—using information contained in patent applications. The instrument relaxes the identifying assumptions traditionally used to recover news shocks and, by construction, isolates shocks that have no immediate effect on aggregate productivity. Embedding the instrument in a structural VAR, the authors trace how technology news propagates through the economy and drives business-cycle fluctuations. They find news shocks are an important source of aggregate dynamics. The paper contributes a patent-based identification strategy that sidesteps the fragilities of earlier approaches to measuring news shocks.",
  "logical_flow": "The paper addresses a core difficulty in the news-shock literature: news about future productivity is hard to identify because it does not move current productivity, and existing approaches rely on strong and contestable identifying assumptions. The authors propose that patent applications contain forward-looking information about coming technological improvements, and can therefore be used to build an instrument for technology news. They construct this instrument so that it captures anticipated future productivity gains while having, by design, no contemporaneous effect on aggregate productivity, which is exactly the signature of a news shock. Plugging the instrument into a structural VAR with rich macro aggregates, they identify the dynamic response of the economy to technology news without imposing the traditional restrictions. The estimated responses show that news shocks generate meaningful business-cycle comovement, supporting the view that expectations about future technology matter for fluctuations. The paper thus both provides a cleaner identification tool and reinforces the empirical relevance of news-driven cycles.",
  "research_design": "An instrumental-variables / structural-VAR design in which an external instrument for technology news shocks is constructed from patent-application data (via the NBER-USPTO historical patent files). The instrument is built to capture anticipated future productivity while having no contemporaneous productivity effect, and is used to identify news shocks in a VAR without the traditional identifying restrictions. The analysis is macro time-series; identification comes from the patent-based instrument's relevance and its exogeneity to current productivity.",
  "categories": [
   "Macroeconomics",
   "Business Cycles",
   "Innovation"
  ],
  "datasets": [
   {
    "provider": "NBER–USPTO",
    "product": "Historical Patent Data Files (Marco et al. 2015)",
    "description": "Comprehensive data on U.S. patent applications used to construct a forward-looking instrument for technology news shocks.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "patents; applications; news shocks; instrument"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Patent-based instrument for technology news shocks",
   "type": "Instrument / news shock",
   "what": "An instrumental variable built from information in patent applications that isolates anticipated future productivity improvements (technology news) while having no contemporaneous effect on aggregate productivity, relaxing the identifying assumptions used in prior work."
  },
  "missing_notes": null,
  "std_name": "Review of Economic Studies - 2026 - Miranda-Agrippino et al. - Patents, News, and Business Cycles",
  "orig_filename": "rdaf086.pdf"
 },
 {
  "journal": "Review of Economic Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Ivanov",
   "Pettit",
   "Whited"
  ],
  "title": "Taxes Depress Corporate Borrowing: Evidence from Private Firms",
  "summary": "This paper uses variation in U.S. state corporate income tax rates to re-examine how taxes affect corporate leverage, focusing on small private firms. Contrary to prior research, the authors find that corporate leverage rises after tax cuts for these firms. An estimated dynamic equilibrium model explains why: tax cuts make capital more productive and spur investment financed partly with debt, and they push default thresholds further away, making borrowing safer. The combination overturns the standard prediction that lower taxes reduce the debt tax shield and hence leverage. The paper shows that for financially constrained private firms, the investment and risk channels of taxes dominate the classic tax-shield channel.",
  "logical_flow": "The paper revisits the textbook relationship between taxes and leverage, where higher tax rates raise the value of the interest tax shield and should therefore increase debt, a prediction that has received mixed empirical support. The authors focus on small private firms, for which financing frictions and investment responses may reshape the tax–leverage link, and exploit changes in state corporate income tax rates as variation in the tax environment. Empirically, they find the opposite of the standard prediction: leverage rises after tax cuts for these firms. To interpret this, they build and estimate a dynamic equilibrium model in which taxes affect not only the debt tax shield but also the productivity of capital and the firm's distance to default. In the model, a tax cut raises the return to investing, and firms fund the extra investment partly with debt, while the improved cash flows move the default threshold further away and make additional borrowing safer. These investment and risk channels dominate the tax-shield channel for constrained private firms, producing the observed rise in leverage after tax cuts and reconciling the evidence with theory.",
  "research_design": "A reduced-form design exploiting staggered changes in U.S. state corporate income tax rates (and tax-base rules) to estimate their effect on the leverage of small private firms, combined with an estimated dynamic equilibrium (structural) model of firm investment and financing. The reduced-form estimates identify the sign of the tax–leverage relationship from cross-state, over-time tax variation, and the structural model decomposes the response into tax-shield, capital-productivity, and default-risk channels. Confidential bank-reported loan-level data on commercial and industrial loans provide firm borrowing outcomes.",
  "categories": [
   "Corporate Finance",
   "Taxation",
   "Capital Structure"
  ],
  "datasets": [
   {
    "provider": "State corporate income tax data",
    "product": null,
    "description": "State-by-year corporate income tax rates and tax-base rules (e.g., depreciation and apportionment), hand-collected/extended, used as variation in the tax incentive to borrow.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "state taxes; corporate income; leverage"
   },
   {
    "provider": "Federal Reserve",
    "product": "FR Y-14 supervisory data",
    "description": "Confidential bank-reported loan-level data on commercial and industrial loans, used to measure private firms' borrowing.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "supervisory; C&I loans; private firms"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "State corporate income tax rate changes",
   "type": "Staggered policy variation",
   "what": "Changes in U.S. state corporate income tax rates and tax-base rules across states and over time, used as quasi-exogenous variation in the tax incentive to use debt."
  },
  "missing_notes": null,
  "std_name": "Review of Economic Studies - 2026 - Ivanov et al. - Taxes Depress Corporate Borrowing Evidence from Private Firms",
  "orig_filename": "rdaf094.pdf"
 },
 {
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Argyle",
   "Iverson",
   "Kotter",
   "Nadauld",
   "Palmer"
  ],
  "title": "The Dynamics of Retail Deposit Balances",
  "summary": "This paper uses transaction-level data on retail deposit accounts to study how households manage their deposit balances. The authors document substantial skewness and inertia in balances, with many depositors holding large amounts—including balances above the FDIC insurance limit—and adjusting slowly to interest-rate incentives. They identify notches at product-specific thresholds (such as service or rate tiers) and show how balances bunch and respond around them. These patterns imply that retail deposits are stickier and less rate-sensitive than frictionless models assume, with implications for banks' funding costs and market power. The paper provides granular evidence on the microstructure of household deposit behavior.",
  "logical_flow": "The paper motivates that retail deposits are a huge and stable source of bank funding, yet how individual households actually manage their balances is not well documented at the transaction level. Using granular account data, the authors first characterize the distribution of balances, finding it highly skewed, with a large share of deposits held by high-balance depositors and substantial amounts above the FDIC insurance limit. They then examine dynamics, showing that balances are inert and adjust slowly to interest-rate incentives, contrary to models in which depositors quickly chase yield. To sharpen the analysis, they exploit product-specific notches—thresholds at which the interest rate or services change—and study how balances bunch and respond around them, which reveals the frictions governing deposit adjustment. These patterns imply that deposits are stickier and less rate-elastic than standard assumptions, giving banks funding stability and pricing power. The paper draws out implications for how deposit inertia shapes banks' funding costs and the transmission of interest rates.",
  "research_design": "An empirical, descriptive study using proprietary transaction-level retail deposit account data to characterize the distribution and dynamics of household deposit balances. The authors document skewness and inertia and use product-specific notches (thresholds where rates or services change) as a bunching/quasi-experimental design to identify how balances respond to incentives. The unit of analysis is the account over time; regulatory data (Call Reports, FDIC) are used to benchmark aggregate patterns.",
  "categories": [
   "Household Finance",
   "Banking",
   "Deposits"
  ],
  "datasets": [
   {
    "provider": "Proprietary retail deposit data provider",
    "product": null,
    "description": "Transaction- and account-level data on retail deposit balances and activity, used to study the distribution and dynamics of household deposits.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "transaction-level; retail deposits; accounts"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Working Paper - 2026 - Argyle et al. - The Dynamics of Retail Deposit Balances",
  "orig_filename": "w34742.pdf"
 },
 {
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Lu",
   "Wu"
  ],
  "title": "Banking on Inattention",
  "summary": "This paper argues that depositor inattention is a source of banks' deposit market power. The authors develop a model in which inattentive depositors adjust their balances sluggishly to changes in deposit rates, allowing banks to pay rates that respond incompletely to the policy rate and to earn rents on deposits. Using transaction-level evidence together with Call Reports and the FDIC Summary of Deposits, they show that more inattentive depositors adjust balances less and that this inattention maps into weaker deposit-rate pass-through. The framework links a behavioral friction—limited attention—to bank funding costs and the transmission of monetary policy. It implies that inattention, not just market concentration, underpins deposit market power.",
  "logical_flow": "The paper begins from the puzzle that bank deposit rates respond only partially to the policy rate, which is usually attributed to market concentration, and proposes an alternative, behavioral source: depositor inattention. It builds a model in which depositors do not continuously monitor deposit rates and therefore adjust their balances sluggishly, so banks can widen the spread between the policy rate and the deposit rate without losing funds quickly. In this setting, inattention directly generates deposit market power and incomplete pass-through, independent of how concentrated the market is. The authors then bring evidence to bear, using transaction-level data to show that more inattentive depositors move their balances less in response to rate changes, and linking this to aggregate patterns in Call Reports and the FDIC Summary of Deposits. The empirical mapping from inattention to weaker pass-through supports the model's central mechanism. The paper concludes that limited attention is a key micro-foundation for banks' deposit franchise and matters for how monetary policy transmits through deposits.",
  "research_design": "A theoretical model of depositor inattention generating deposit market power and incomplete deposit-rate pass-through, combined with empirical evidence. The empirical work uses transaction-level data to measure how depositor inattention relates to balance adjustment and pass-through, with U.S. Call Reports and the FDIC Summary of Deposits used for aggregate deposit sizes and branch-level patterns. Identification of the mechanism comes from relating measured inattention to the sensitivity of balances and deposit rates rather than from a single external shock.",
  "categories": [
   "Banking",
   "Monetary Economics",
   "Household Finance"
  ],
  "datasets": [
   {
    "provider": "Proprietary transaction-level data provider",
    "product": null,
    "description": "Household transaction-level data used to measure depositor inattention and the responsiveness of deposit balances to rate changes.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "transaction-level; deposits; inattention"
   },
   {
    "provider": "FDIC",
    "product": "Summary of Deposits",
    "description": "Annual branch-level deposit data (2007-2021) used to characterize deposit sizes and market structure.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "deposits; branches; market structure"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Working Paper - 2026 - Lu and Wu - Banking on Inattention",
  "orig_filename": "w34783.pdf"
 },
 {
  "journal": "Journal of Financial Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Acikalin",
   "Caskurlu",
   "Hoberg",
   "Phillips"
  ],
  "title": "Intellectual property protection lost and competition: An examination using large language models",
  "summary": "This paper examines who gains and who loses when intellectual-property protection is broadly weakened, using the 2014 Alice v. CLS Bank Supreme Court decision that invalidated many business-method patents. The authors use a large language model (ModernBERT) to score each firm's pre-Alice patents for their exposure to the decision and build a firm-level treatment. Using difference-in-differences with entropy balancing, they find an unequal impact: large firms gain while small firms lose. Small treated firms face more venture-backed entry, product-market encroachment, more litigation, and lower profits and valuations, and respond by raising R&D and nondisclosure. The paper contributes both an LLM-based method for measuring legal shocks and evidence that patents matter most for small firms.",
  "logical_flow": "The paper enters the long debate over whether patents help or hinder innovation by treating the Alice decision as a broad, technology-area-wide weakening of IP rather than the loss of a single rival's patent. Because the ex-post fate of any patent is uncertain, the authors argue a language model is needed to gauge each patent's similarity to those invalidated under Alice, yielding a continuous firm-level exposure measure. They hypothesize, following leader-versus-laggard logic, that smaller firms with fewer resources and weaker barriers to entry are hurt more, while larger firms defend their positions. These hypotheses generate predictions about differential effects on patenting, competition, performance, valuation, and litigation. The empirical design then tests each prediction, and the pattern of small-firm losses and large-firm gains supports the view that IP protection disproportionately shields smaller innovators. The paper closes by drawing out the policy implication that uniform changes in patent strength have highly unequal effects across firm size.",
  "research_design": "A continuous-treatment difference-in-differences design exploiting the Alice Supreme Court decision as a quasi-exogenous shock, with entropy balancing for covariate balance and firm/year fixed effects. The treatment is constructed from a fine-tuned ModernBERT model applied to patent text, validated against USPTO post-grant reviews. Effects on innovation, competition, performance, valuation, and litigation are compared across small and large treated firms; the unit of analysis is the firm-year.",
  "categories": [
   "Innovation & Intellectual Property",
   "Industrial Organization",
   "Machine Learning in Finance"
  ],
  "datasets": [
   {
    "provider": "KPSS patent value database",
    "product": "Kogan, Papanikolaou, Seru & Stoffman (2017)",
    "description": "Dollar valuations of individual patents, used to weight firms' Alice-exposed patents.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "patent value; innovation"
   },
   {
    "provider": "USPTO",
    "product": "Patent application office actions",
    "description": "U.S. patent applications and examiner rejections used to identify Alice-based rejections and build training labels.",
    "access_type": "Public",
    "delivery": "Bulk",
    "topic_tags": "patents; rejections"
   },
   {
    "provider": "Hoberg-Phillips",
    "product": "TNIC",
    "description": "Text-based product-market similarity network used to measure competitors and encroachment.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "product market; competition"
   },
   {
    "provider": "Stanford NPE Litigation Database / PACER",
    "product": null,
    "description": "U.S. patent-infringement lawsuits (NPE vs operating company) used to measure litigation exposure.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "patent litigation; NPE"
   },
   {
    "provider": "Unified Patents",
    "product": "Post-grant review data",
    "description": "USPTO post-grant reviews used to validate the ModernBERT Alice score.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "post-grant review; validation"
   },
   {
    "provider": "Venture Expert (Refinitiv)",
    "product": null,
    "description": "VC financing rounds and startup descriptions used to measure VC-backed entry.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "venture capital; entry"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Alice v. CLS Bank Supreme Court decision (2014)",
   "type": "Supreme Court decision",
   "what": "A 2014 ruling that invalidated many business-method and software patents, broadly and unexpectedly weakening IP protection across affected technology areas, used as a quasi-exogenous shock to firms' patent portfolios."
  },
  "missing_notes": null,
  "std_name": "Journal of Financial Economics - 2026 - Acikalin et al. - Intellectual property protection lost and competition An examination using large language models",
  "orig_filename": "1-s2.0-S0304405X26000772-main.pdf"
 },
 {
  "journal": "Journal of Financial Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Hvide",
   "Nielsen"
  ],
  "title": "Flying below the radar: Insider trading by executives below the top",
  "summary": "This paper studies whether executives just below the reportable-insider threshold trade profitably on material inside information, using Norwegian administrative register data on the population of trades. It finds abnormal returns of roughly 50-100 basis points at the one-month horizon on these executives' own-company purchases, rising over longer horizons. Because the same individuals earn no abnormal returns on non-inside or same-industry purchases, general stock-picking skill and industry knowledge are ruled out. The evidence indicates that a broad group of executives beyond top management trades on material information that escapes mandatory disclosure. It is the first transaction-level test of below-the-top executives trading on inside information.",
  "logical_flow": "The paper begins from the regulatory design that requires only primary insiders to disclose own-company trades, leaving below-the-top executives unmonitored despite their proximity to inside information. Whether they actually trade on that information is posed as an open question, since they might instead buy own-company stock out of familiarity or loyalty, which would not generate abnormal returns. Using positional codes to isolate below-the-top executives, the authors compare returns on their inside versus non-inside trades. Crucially, they benchmark inside purchases against the same individuals' other trades to net out ability and industry knowledge. Finding abnormal returns only on inside purchases, they conclude these executives trade on material information. The paper extends the analysis to indirect trades and job changes to interpret the source of the returns.",
  "research_design": "Long-run abnormal-return estimation on administrative (population) register data using a control-firm bootstrap and a calendar-time Carhart four-factor portfolio approach, with individual-specific benchmarks that difference out stock-picking ability. The unit of analysis is the individual trade; identification of information-based trading comes from comparing inside purchases to the same person's non-inside and same-industry purchases.",
  "categories": [
   "Insider Trading",
   "Financial Regulation",
   "Empirical Asset Pricing"
  ],
  "datasets": [
   {
    "provider": "Norwegian Central Securities Depository (VPS)",
    "product": null,
    "description": "All individual stock transactions on the Oslo Stock Exchange, 1997-2014.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "individual trades; Norway"
   },
   {
    "provider": "Statistics Norway",
    "product": "Employer-employee register (ISCO-88)",
    "description": "Population employer-employee data identifying each individual's position and demographics.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "employer-employee; occupation"
   },
   {
    "provider": "Oslo Stock Exchange",
    "product": null,
    "description": "Daily prices and market capitalization used to compute returns.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "prices; market cap"
   },
   {
    "provider": "Norwegian registers (LLC ownership; population)",
    "product": null,
    "description": "LLC ownership and family links used to study indirect trades.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "LLC; family links"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Financial Economics - 2026 - Hvide and Nielsen - Flying below the radar Insider trading by executives below the top",
  "orig_filename": "1-s2.0-S0304405X2600053X-main.pdf"
 },
 {
  "journal": "Journal of Financial Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Cassella",
   "Rizzo",
   "Spalt",
   "Zimmerer"
  ],
  "title": "Constrained by law: The impact of fiduciary duties on portfolios and prices in US equity markets",
  "summary": "This paper studies how a change in trust fiduciary law reshaped institutional portfolios and equity prices. The Uniform Prudent Investor Act (UPIA), adopted by U.S. states in staggered fashion from 1985 to 2006, replaced asset-by-asset prudent-man rules with a portfolio-level prudent-investor standard. The authors show trusts tilted toward 'prudent' stocks before the reform and undid these tilts after adoption, improving trust risk-return. Consistent with inelastic markets, the demand shift moved relative prices, implying demand elasticities between 0.10 and 0.36. The paper documents a decades-long, law-induced investment distortion and adds novel estimates of equity demand elasticity.",
  "logical_flow": "The paper contrasts the diversification prescriptions of modern portfolio theory with prudent-man laws that forced trustees to justify each holding in isolation, creating a binding tilt toward 'prudent' stocks. A CARA-normal model with trust managers facing a prudence-tilted benchmark yields three predictions: regulatory tilts before UPIA, an unwinding after UPIA, and price changes under inelastic markets. Because UPIA was adopted state-by-state and applies only to trusts, the design compares trusts to other institutions within state and time. The empirical sections move from holdings, to risk-return consequences, to relative prices. Finding that trusts undid their tilts and that prices moved as predicted, the authors back out demand elasticities from the regulation-induced demand shock. The result ties a legal constraint to portfolio behavior and to the inelasticity of equity markets.",
  "research_design": "Staggered difference-in-differences exploiting state-level UPIA adoption, supported by a CARA-normal inelastic-markets model. The design uses investor and state-by-quarter fixed effects and the Sun-Abraham estimator for heterogeneous treatment effects; Fama-French five-factor tests measure risk-return changes and an instrumented demand system recovers elasticities. The unit of analysis is the institution-stock-quarter; identification comes from the timing of UPIA across states.",
  "categories": [
   "Institutional Investors & Asset Management",
   "Demand-Based Asset Pricing",
   "Law and Finance"
  ],
  "datasets": [
   {
    "provider": "Thomson Reuters",
    "product": "Institutional (13F) Holdings",
    "description": "Quarterly institutional equity holdings used to observe trust and non-trust portfolios around UPIA.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "13F; holdings"
   },
   {
    "provider": "DeVault, Sias & Starks (2019)",
    "product": "Institutional classification (57 types)",
    "description": "Fine-grained institution-type classifications used to identify trusts.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "institution types; trusts"
   },
   {
    "provider": "FDIC / WestLaw",
    "product": "UPIA adoption dates",
    "description": "Hand-collected state-by-state UPIA statutes and effective dates.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "UPIA; state law"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Uniform Prudent Investor Act (UPIA) adoption",
   "type": "Staggered state-law reform",
   "what": "States' staggered 1985-2006 replacement of asset-by-asset 'prudent-man' rules with a portfolio-level prudent-investor standard, which applied only to trusts and relaxed a binding constraint on their holdings."
  },
  "missing_notes": null,
  "std_name": "Journal of Financial Economics - 2026 - Cassella et al. - Constrained by law The impact of fiduciary duties on portfolios and prices in US equity markets",
  "orig_filename": "1-s2.0-S0304405X25002351-main.pdf"
 },
 {
  "journal": "Journal of Financial Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Goyal",
   "Wahal",
   "Yavuz"
  ],
  "title": "Picking partners: Manager selection in private markets",
  "summary": "This paper studies how limited partners select private-market managers across more than 61,000 commitments (1995-2019) by constructing feasible 'opportunity sets.' LPs chase past performance but are strikingly willing to back first-time or young GPs: the probability of selecting a first-time GP is about as high as selecting a top-quartile GP, even though first-time funds earn lower expected and realized returns. After ruling out expected performance, lottery preferences, star-fund access, and co-investment, the authors show the main driver is pressure to invest as private-market allocations grow. The paper reframes manager selection around demand and supply rather than skill.",
  "logical_flow": "The paper frames manager selection as a discrete-choice problem whose central difficulty is that the feasible choice set is unobserved, which it resolves by constructing opportunity sets from institutional detail, interviews, and investment-policy statements. Estimating selection against these sets, it documents robust performance chasing alongside a high propensity to pick first-time and young GPs. It then eliminates candidate explanations, showing first-time GPs do not deliver higher expected or realized returns, greater skewness, privileged access to later star funds, or more co-investment. Having ruled these out, it develops a demand-supply account and constructs a measure of pressure to invest. It shows LPs facing high pressure move down the performance ladder toward first-time GPs. The paper concludes that excess demand for private-market exposure, met by new GP entry, explains the willingness to back unproven managers.",
  "research_design": "Discrete-choice selection modeling (conditional logit and linear-probability specifications) estimated on curated LP-year opportunity sets with two-way exclusion restrictions ('feasible-set shifters'). Robustness comes from alternative opportunity-set constructions, Heckman selection for missing performance, and expected-performance estimation. The unit of analysis is the LP-fund choice within a feasible set.",
  "categories": [
   "Private Equity & Venture Capital",
   "Institutional Investors & Asset Management",
   "Manager Selection"
  ],
  "datasets": [
   {
    "provider": "Preqin",
    "product": null,
    "description": "LP capital commitments and fund-level performance (IRR, TVPI), AuM, and flags for private-market funds 1995-2019.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "private equity; commitments; IRR"
   },
   {
    "provider": "CEM Benchmarking",
    "product": null,
    "description": "LP target and actual allocations used to measure underweighting and pressure to invest.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "target allocations"
   },
   {
    "provider": "Institute for Private Capital",
    "product": "Failed-fund data",
    "description": "Funds that failed to raise capital, used to build non-selected opportunity sets.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "failed funds"
   },
   {
    "provider": "SEC",
    "product": "Form D filings",
    "description": "Regulation D filings used to identify failed first-time fundraising.",
    "access_type": "Public",
    "delivery": "Bulk",
    "topic_tags": "Form D; first-time funds"
   },
   {
    "provider": "RelSci",
    "product": null,
    "description": "Personal/professional links between LP and GP employees.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "personal networks"
   },
   {
    "provider": "Fundmap",
    "product": null,
    "description": "Data on LPs' use of specialized search consultants.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "search consultants"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Financial Economics - 2026 - Goyal et al. - Picking partners Manager selection in private markets",
  "orig_filename": "1-s2.0-S0304405X26000218-main.pdf"
 },
 {
  "journal": "Journal of Financial Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Goldstein",
   "Liu",
   "Yang"
  ],
  "title": "Market feedback: Evidence from the horse's mouth",
  "summary": "This paper uses near-population surveys of Chinese public firms (2019 and 2022, via the CSRC, ~99% response) to provide direct evidence on whether and how the stock market affects firms, focusing on the contested learning channel. Over 90% of firms monitor the market, and the two most common reasons are learning new information relevant to investment and dependence on prices for financing. Firms that report learning have characteristics implying greater benefit and show higher investment-to-price sensitivity. The paper documents which information dimensions firms learn about and links learning to better M&A outcomes and fewer trading suspensions. It offers rare direct evidence on the real effects of markets.",
  "logical_flow": "The paper situates itself in the debate over the real effects of markets, distinguishing a financing channel from a learning channel whose existence is hard to identify because information sets are unobservable. Rather than infer learning indirectly, the authors ask managers directly through near-population surveys, establishing that monitoring is widespread and that learning and financing dominate. They then open the black box of what information is learned, a dimension prior work largely could not address. To counter the concern that survey answers are noise, they derive and test predictions about which firm characteristics should predict learning. They further test whether reported learning maps into real actions—investment-to-price sensitivity, M&A performance, and trading suspensions. Finding that responses line up with characteristics and behavior, they argue the survey evidence is credible and that learning from prices materially shapes firm decisions.",
  "research_design": "A survey-based, direct-evidence design using two rounds of near-population firm surveys, validated with cross-sectional Probit models linking responses to firm characteristics and panel investment-to-price sensitivity regressions, plus M&A and trading-suspension tests. The unit of analysis is the firm; credibility comes from tying stated learning to observable characteristics and subsequent actions.",
  "categories": [
   "Real Effects of Financial Markets",
   "Corporate Investment",
   "Survey Methods in Finance"
  ],
  "datasets": [
   {
    "provider": "CSRC & Tsinghua PBCSF",
    "product": "Survey of Chinese public firms (2019, 2022)",
    "description": "Two near-population surveys (response rates 99.9% and 98.1%) on whether/why firms monitor prices and what they learn.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "survey; learning; China"
   },
   {
    "provider": "CSMAR",
    "product": "China Stock Market & Accounting Research",
    "description": "Chinese prices, financials, M&A, and suspensions used to build variables and validate responses.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "China; prices; M&A"
   },
   {
    "provider": "Wind",
    "product": null,
    "description": "Chinese market and firm data used alongside CSMAR.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "China; firm data"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Financial Economics - 2026 - Goldstein et al. - Market feedback Evidence from the horse's mouth",
  "orig_filename": "1-s2.0-S0304405X26000267-main.pdf"
 },
 {
  "journal": "Journal of Financial Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Fisman",
   "Ghosh",
   "Sen"
  ],
  "title": "Dirty air and green investments: The impact of pollution information on portfolio allocations",
  "summary": "This paper asks whether access to local pollution information causes investors to make greener portfolio choices, exploiting the staggered rollout of air-quality monitoring stations across India. Using a triple-differences design on the trading records of 19 million retail investors, the authors show that holdings in 'brown' (high-pollution) stocks become more negatively related to local pollution once a nearby monitoring station appears. The effect is stronger on 'alert' dates when air quality is reported harmful, among tech-savvy investors likely exposed to real-time data, and among younger investors. The results show that making environmental information salient shifts capital away from polluting firms. The paper links information provision to sustainable investing behavior.",
  "logical_flow": "The paper starts from the idea that investors may want to avoid polluting firms but lack salient, local information about pollution, so providing it should change behavior. Air-quality monitoring stations, rolled out at different times across Indian locations, make local pollution observable and are used as a source of quasi-exogenous variation in information access. The authors reason that if information drives green investing, retail holdings in brown stocks should become more negatively related to local pollution after a nearby station appears, and especially on days when pollution is reported harmful. A triple-differences design compares brown-stock holdings, across investors near versus far from new stations, before versus after, and on alert versus normal days. Finding the predicted patterns—strongest for tech-savvy and younger investors—they conclude that salient pollution information causally shifts portfolios. The paper interprets this as evidence that information frictions, not just preferences, shape sustainable investment.",
  "research_design": "A triple-differences design exploiting the staggered rollout of air-quality monitoring stations across India as quasi-exogenous variation in local pollution information, estimated on the trading records of 19 million retail investors. The comparison combines proximity to a new station, timing before/after, and 'alert' versus normal pollution days, with heterogeneity by investor tech-savviness and age. The unit of analysis is the investor-stock-time; identification comes from the station rollout.",
  "categories": [
   "Sustainable Finance & ESG",
   "Household Finance",
   "Information & Markets"
  ],
  "datasets": [
   {
    "provider": "Indian retail investor trading records",
    "product": null,
    "description": "Account-level holdings and trades of ~19 million Indian retail investors, used to measure brown-stock holdings.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "retail investors; India; holdings"
   },
   {
    "provider": "Central Pollution Control Board (India)",
    "product": "Air-quality monitoring stations",
    "description": "Locations, timing, and readings of air-quality monitoring stations, providing the pollution-information shock.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "air quality; monitoring stations; India"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Rollout of air-quality monitoring stations across India",
   "type": "Staggered information rollout",
   "what": "The staggered appearance of local air-quality monitoring stations, which made pollution observable in a location at different times, used as quasi-exogenous variation in access to local pollution information."
  },
  "missing_notes": null,
  "std_name": "Journal of Financial Economics - 2026 - Fisman et al. - Dirty air and green investments The impact of pollution information on portfolio allocations",
  "orig_filename": "1-s2.0-S0304405X26000802-main.pdf"
 },
 {
  "journal": "Journal of Financial Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Benguria",
   "Garcia-Marin",
   "Schmidt-Eisenlohr"
  ],
  "title": "Trade credit and relationships",
  "summary": "This paper uses transaction-level international trade data to show that long-term firm-to-firm relationships facilitate the use of trade credit. The strength of this effect varies with firm size, payment delays, and multinational-affiliate status, and depends on the strength of contract enforcement across countries and the complexity of traded products. Because trade credit substitutes for borrowing from the financial sector, long-term relationships can reduce firms' external credit demand. A corollary is that destroying trade relationships—for example through trade conflicts—may raise firms' leverage. The paper connects relationship lending in trade to firms' financing needs.",
  "logical_flow": "The paper begins from the observation that firms often finance trade through credit extended by their trading partners rather than by banks, and asks what determines the use of such trade credit. It proposes that repeated, long-term relationships between buyers and sellers build the trust and information needed to extend credit, so relationship length should predict trade-credit use. Using detailed transaction-level trade data, it tests whether longer relationships are associated with more trade credit and how this varies with firm characteristics. It then examines institutional moderators, showing the effect is stronger where contract enforcement is weaker and products are more complex, consistent with relationships substituting for formal enforcement. Because trade credit can replace bank borrowing, the paper draws the implication that relationships reduce firms' demand for external finance. It concludes that shocks that sever trade relationships, such as trade conflicts, could push firms back toward bank debt and raise leverage.",
  "research_design": "A descriptive/reduced-form design using transaction-level (customs) international trade data to relate the length of firm-to-firm trading relationships to the use of trade credit, with heterogeneity by firm size, payment delays, and multinational status. Institutional interactions use cross-country contract-enforcement measures and product-complexity classifications. The unit of analysis is the firm-to-firm trade transaction.",
  "categories": [
   "Trade Credit & Financing",
   "International Finance",
   "Firm Financing"
  ],
  "datasets": [
   {
    "provider": "Transaction-level international trade (customs) data",
    "product": null,
    "description": "Firm-to-firm customs transactions with values, timing, payment terms, and partner identities, used to measure relationships and trade credit.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "customs; trade; firm-to-firm"
   },
   {
    "provider": "Contract-enforcement and product-complexity measures",
    "product": null,
    "description": "Cross-country contract-enforcement indicators and product-complexity/differentiation classifications used as institutional moderators.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "contract enforcement; product complexity"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Financial Economics - 2026 - Benguria et al. - Trade credit and relationships",
  "orig_filename": "1-s2.0-S0304405X26000917-main.pdf"
 },
 {
  "journal": "Journal of Financial Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Woeppel",
   "Yavuz"
  ],
  "title": "Measuring firms' capability to absorb external knowledge",
  "summary": "This paper builds a measure of firms' capability to benefit from external knowledge—'absorption intensity'—from information in their patents. To validate it, the authors exploit the American Inventors Protection Act (AIPA), which exogenously increased the availability of patent information, in a triple-difference design: firms with higher absorption intensity see greater innovation growth when more exposed to the reform. Beyond AIPA, absorption intensity predicts stronger innovation outcomes and firm growth. The measure offers a tool to study how differences in absorptive capability shape the returns to external knowledge. It contributes a patent-based proxy for a concept that was previously hard to measure.",
  "logical_flow": "The paper argues that as the economy becomes knowledge-based, a firm's ability to absorb and use external knowledge is increasingly important, yet this absorptive capability is hard to measure. The authors propose to infer it from patents, constructing an 'absorption intensity' measure that captures how well a firm draws on outside knowledge. To show the measure is meaningful rather than mechanical, they need a setting where the value of absorptive capability changes exogenously. AIPA, which increased the disclosure of patent information, provides such a shock: firms better able to absorb external knowledge should benefit more when more information becomes available. A triple-difference design comparing high- and low-absorption firms by exposure to AIPA confirms that high-absorption firms experience greater innovation growth. The paper then shows absorption intensity predicts innovation and growth more broadly, establishing it as a useful, validated proxy.",
  "research_design": "A measurement paper validated with a triple-difference design around the American Inventors Protection Act (AIPA), which exogenously increased patent-information availability. The absorption-intensity measure is built from patent data, and the triple difference compares high- versus low-absorption firms with greater versus lesser exposure to the reform, before versus after. Broader predictive regressions relate the measure to innovation and firm growth; the unit of analysis is the firm-year.",
  "categories": [
   "Innovation & Intellectual Property",
   "Corporate Finance",
   "Firm Growth"
  ],
  "datasets": [
   {
    "provider": "USPTO patent data",
    "product": null,
    "description": "Patent grants, citations, and text used to construct firms' absorption-intensity measure.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "patents; citations; absorption"
   },
   {
    "provider": "KPSS patent value database",
    "product": "Kogan, Papanikolaou, Seru & Stoffman (2017)",
    "description": "Market-based patent values used in validating innovation outcomes.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "patent value"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "American Inventors Protection Act (AIPA)",
   "type": "Legislative reform",
   "what": "A U.S. reform that required publication of most patent applications 18 months after filing, exogenously increasing the availability of patent information and raising the payoff to absorbing external knowledge."
  },
  "missing_notes": null,
  "std_name": "Journal of Financial Economics - 2026 - Woeppel and Yavuz - Measuring firms' capability to absorb external knowledge",
  "orig_filename": "1-s2.0-S0304405X26000930-main.pdf"
 },
 {
  "journal": "Journal of Financial Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Lindsey",
   "Stein"
  ],
  "title": "Angels, entrepreneurship, and employment dynamics: Evidence from investor accreditation rules",
  "summary": "This paper examines how a shock to the supply of angel finance affects entrepreneurship, using Dodd-Frank's removal of housing wealth from the definition of an accredited investor. Using U.S. Census data, the authors estimate the state-level fraction of households that lost accreditation and show that larger reductions in the investor pool reduce angel investment, firm entry, and employment at small entrants. Employment rises at small and young incumbents, suggesting competitive effects, and angel finance appears to complement other capital sources despite partial substitution. The paper provides causal evidence on the real effects of angel finance and where it matters most.",
  "logical_flow": "The paper studies angel investors, an important but hard-to-observe source of early-stage capital, and asks what happens to entrepreneurship when their supply shrinks. It exploits a specific rule change: Dodd-Frank stopped counting housing wealth toward the wealth threshold that defines an accredited (eligible angel) investor, which pushed some households below the threshold. Because house values differ across places, the fraction of households losing accreditation varies by state, generating quasi-exogenous cross-state variation in the angel-investor pool. The authors predict that a larger reduction in potential angels lowers angel investment, firm entry, and employment at small new firms. Using Census data, they confirm these effects, and additionally find employment rising at small and young incumbents, consistent with reduced entry easing competition. Documenting partial substitution toward other capital, they conclude angel finance complements rather than merely substitutes for other funding in the entrepreneurial ecosystem.",
  "research_design": "A quasi-experimental design exploiting Dodd-Frank's removal of housing wealth from accredited-investor eligibility, which reduced the angel-investor pool differentially across states depending on local house values. Using confidential U.S. Census firm and employment data, the authors relate state-level exposure (the fraction of households losing accreditation) to angel investment, firm entry, and employment at entrants and incumbents. The unit of analysis is the state/firm; identification comes from cross-state variation in exposure to the rule change.",
  "categories": [
   "Entrepreneurial Finance",
   "Firm Dynamics",
   "Financial Regulation"
  ],
  "datasets": [
   {
    "provider": "U.S. Census Bureau",
    "product": "Longitudinal Business Database / firm microdata",
    "description": "Confidential firm entry and employment data used to measure entrepreneurship outcomes.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "firm entry; employment; Census"
   },
   {
    "provider": "Angel investment and accreditation data",
    "product": null,
    "description": "Measures of angel investment activity and household accreditation (via housing-wealth thresholds) used to build the exposure shock.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "angel finance; accreditation"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Dodd-Frank removal of housing wealth from accredited-investor rules",
   "type": "Regulatory reform",
   "what": "A Dodd-Frank change that excluded home equity from the wealth test defining accredited investors, pushing some households below the threshold; because house values differ across states, it reduced the angel-investor pool by varying amounts geographically."
  },
  "missing_notes": null,
  "std_name": "Journal of Financial Economics - 2026 - Lindsey and Stein - Angels, entrepreneurship, and employment dynamics Evidence from investor accreditation rules",
  "orig_filename": "1-s2.0-S0304405X26000942-main.pdf"
 },
 {
  "journal": "Journal of Financial Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Kaviani",
   "Maleki",
   "Savor"
  ],
  "title": "Reaching for influence: Do banks use loans to establish political connections?",
  "summary": "This paper asks whether U.S. banks lend on favorable terms to firms with political ties in order to buy influence, using close congressional elections for identification. Firms connected to members of Congress receive more favorable loan terms despite no observable improvement in performance or default risk. The effect is strongest among banks facing regulatory challenges—FDIC enforcement actions, misconduct investigations, or low Community Reinvestment Act ratings—which also lend more often to connected firms, suggesting a heightened demand for influence. Politically motivated lending yields tangible benefits for banks, including reduced misconduct penalties and easier M&A approval. The paper documents a channel through which banks convert lending into political capital.",
  "logical_flow": "The paper starts from the idea that banks, which are heavily regulated, may value political influence and could try to acquire it by lending cheaply to politically connected firms. To identify politically motivated lending, it uses close elections, where whether a firm's favored candidate barely wins or loses is close to random, generating quasi-exogenous variation in a firm's political connections. It first shows that firms newly connected to winning members of Congress receive better loan terms even though their fundamentals do not improve, indicating the favorable terms reflect influence-seeking rather than credit quality. It then asks which banks do this most, finding that banks under regulatory pressure lend more to connected firms, consistent with their greater demand for political protection. Finally, it tests whether this pays off, showing connected lending is followed by reduced penalties and easier merger approvals. The paper concludes that banks strategically use loans to build political connections that yield regulatory benefits.",
  "research_design": "A regression-discontinuity-style design around close congressional elections, using the narrow victory or defeat of a firm's connected candidate as quasi-random variation in political connections. The authors compare loan terms to connected versus non-connected firms and test heterogeneity by banks' regulatory pressure (FDIC actions, misconduct investigations, CRA ratings), then relate connected lending to later bank outcomes (penalties, M&A approvals). The unit of analysis is the loan/firm-bank; identification comes from the close-election discontinuity.",
  "categories": [
   "Banking",
   "Political Economy & Finance",
   "Corporate Governance"
  ],
  "datasets": [
   {
    "provider": "Refinitiv LPC",
    "product": "DealScan",
    "description": "Syndicated loan terms used to measure loan pricing and covenants for connected and non-connected firms.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "syndicated loans; terms"
   },
   {
    "provider": "Close-election and political-connection data",
    "product": null,
    "description": "Congressional election margins and firm-politician ties used to construct quasi-random political connections.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "close elections; connections"
   },
   {
    "provider": "Bank regulatory records",
    "product": null,
    "description": "FDIC enforcement actions, misconduct investigations, and Community Reinvestment Act ratings used to measure banks' regulatory pressure and outcomes.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "FDIC; CRA; enforcement"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Close congressional elections",
   "type": "Regression discontinuity",
   "what": "Narrowly decided U.S. congressional elections, where whether a firm's connected candidate barely wins or loses is close to random, used as quasi-exogenous variation in firms' political connections."
  },
  "missing_notes": null,
  "std_name": "Journal of Financial Economics - 2026 - Kaviani et al. - Reaching for influence Do banks use loans to establish political connections",
  "orig_filename": "1-s2.0-S0304405X26001042-main.pdf"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Leonelli",
   "Muhn",
   "Rauter",
   "Sran"
  ],
  "title": "How do consumers use ESG disclosure? Evidence from a randomized field experiment with everyday product purchases",
  "summary": "This paper combines a large-scale field experiment with a customized survey of over 24,000 U.S. households to study how consumers use ESG disclosure. The authors first establish that while consumers moderately prefer ESG-responsible firms, they rarely consult corporate reporting directly and face frictions in learning about firm activities. In the experiment, households are randomly informed about real firm-disclosed activities, and consumers raise their purchase intent when exogenously shown positive ESG signals. The results show that ESG information can move consumer behavior, but only once frictions in accessing it are removed. The paper provides causal, individual-level evidence on the demand-side value of ESG disclosure.",
  "logical_flow": "The paper starts from the debate over whether ESG disclosure actually matters to the consumers it is often meant to inform. It first documents, via survey, that consumers say they prefer responsible firms but seldom read corporate reports and struggle to learn what firms actually do, implying an information friction rather than an absence of preferences. This motivates an experiment that removes the friction by exogenously delivering real firm-disclosed ESG information to randomly selected households. The authors predict that if the friction is what suppresses ESG-based choice, then providing information should raise purchase intent for firms with positive disclosures. The experiment confirms this, with purchase intent rising when households receive positive ESG signals. The paper concludes that ESG disclosure has demand-side value that is latent until consumers are actually exposed to the information.",
  "research_design": "A randomized field experiment paired with a customized survey of more than 24,000 U.S. households. The survey establishes baseline preferences and frictions, and the experiment randomly assigns households to information treatments that reveal real firm-disclosed ESG activities, identifying the causal effect on purchase intent. The unit of analysis is the household; identification comes from the random assignment of ESG information.",
  "categories": [
   "Sustainable Finance & ESG",
   "Disclosure",
   "Consumer Behavior"
  ],
  "datasets": [
   {
    "provider": "Authors' field experiment and survey",
    "product": null,
    "description": "Randomized information treatments and a customized survey covering 24,000+ U.S. households, measuring ESG preferences, information frictions, and purchase intent.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "field experiment; survey; ESG; consumers"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Randomized field experiment on ESG information",
   "type": "Randomized field experiment",
   "what": "Randomized information treatments delivered to 24,000+ U.S. households that exogenously revealed real firm-disclosed ESG activities, enabling causal estimates of how ESG information affects purchases."
  },
  "missing_notes": null,
  "std_name": "Journal of Accounting and Economics - 2026 - Leonelli et al. - How do consumers use ESG disclosure Evidence from a randomized field experiment with everyday product purchases",
  "orig_filename": "1-s2.0-S0165410125000473-main.pdf"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Miller",
   "Stockbridge",
   "Williams"
  ],
  "title": "Mandatory disclosure of investors' fossil fuel holdings",
  "summary": "This paper examines whether mandatory disclosure of fossil-fuel investments changes investors' portfolios, using a 2016 California mandate that required some U.S. insurers to publicly disclose their fossil-fuel holdings. Disclosing insurers reduced fossil-fuel investments by roughly 20% relative to non-disclosers. The average effect masks substantial heterogeneity across insurers. The results show that requiring firms to reveal environmentally sensitive holdings can shift capital away from those assets. The paper provides causal evidence on how disclosure mandates affect institutional portfolios rather than just disclosure itself.",
  "logical_flow": "The paper addresses a growing regulatory push to require investors to disclose fossil-fuel holdings, asking whether such disclosure changes behavior or is merely cosmetic. It uses a 2016 California mandate that applied to some U.S. insurers and made their fossil-fuel investments public, creating a treated group subject to disclosure and a comparison group that was not. The authors reason that if public scrutiny raises the cost of holding fossil-fuel assets, disclosing insurers should reduce those holdings relative to non-disclosers. Comparing the two groups around the mandate, they find a roughly 20% reduction in fossil-fuel investments by disclosers. They then explore heterogeneity, showing the response varies across insurers. The paper concludes that mandatory holdings disclosure can meaningfully reallocate institutional capital, with implications for climate-related disclosure policy.",
  "research_design": "A difference-in-differences design exploiting a 2016 California mandate requiring some U.S. insurers to publicly disclose fossil-fuel investments, comparing disclosing insurers to non-disclosers around the mandate. Outcomes are insurers' fossil-fuel holdings and investment policies; the unit of analysis is the insurer. Identification comes from which insurers were subject to the disclosure requirement.",
  "categories": [
   "Sustainable Finance & ESG",
   "Disclosure Regulation",
   "Institutional Investors"
  ],
  "datasets": [
   {
    "provider": "California / NAIC insurer disclosures",
    "product": "Fossil-fuel investment filings",
    "description": "Publicly disclosed fossil-fuel investment holdings of U.S. insurers under the California mandate, with insurers' broader investment portfolios.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "insurers; fossil fuel; disclosure"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "2016 California fossil-fuel disclosure mandate",
   "type": "Disclosure regulation",
   "what": "A 2016 California requirement that certain U.S. insurers publicly disclose their fossil-fuel investments, applying to a subset of insurers and used as a quasi-exogenous disclosure shock."
  },
  "missing_notes": null,
  "std_name": "Journal of Accounting and Economics - 2026 - Miller et al. - Mandatory disclosure of investors' fossil fuel holdings",
  "orig_filename": "1-s2.0-S0165410125000655-main.pdf"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Berger",
   "Tomy"
  ],
  "title": "Supply chain shocks and firm productivity: The role of reporting quality",
  "summary": "This paper studies how reporting quality shapes firms' responses to an adverse supply-chain shock, using the 1999 Taiwan earthquake that disrupted supply chains for certain U.S. high-technology manufacturers. Affected firms experience increases in total factor productivity relative to unaffected firms, and this effect is stronger when pre-shock reporting quality is higher. The findings suggest that better information helps firms restructure effectively after shocks, not just in routine decisions. The paper extends the literature on reporting quality from recurring operating choices to crisis-driven restructuring. It links accounting information quality to real productivity outcomes.",
  "logical_flow": "The paper notes that prior work links reporting quality to routine operating and investing decisions but says little about its role when firms must restructure after an adverse shock. It uses the 1999 Taiwan earthquake as an exogenous disruption to the supply chains of specific U.S. high-technology manufacturers, creating affected and unaffected groups. The authors hypothesize that firms with higher pre-shock reporting quality have better internal information to reallocate resources and adapt, so they should recover or even improve productivity more. Comparing affected to unaffected firms, they find total factor productivity rises for affected firms, and more so where pre-shock reporting quality was higher. This pattern indicates that information quality is especially valuable precisely when firms face large, non-routine disruptions. The paper concludes that reporting quality has real effects that show up in how firms weather shocks.",
  "research_design": "A difference-in-differences design using the 1999 Taiwan earthquake as an exogenous supply-chain shock, comparing exposed U.S. high-technology manufacturers to unexposed firms and interacting exposure with pre-shock reporting quality. The outcome is firm total factor productivity; the unit of analysis is the firm. Identification comes from the earthquake's disruption of specific supply chains.",
  "categories": [
   "Financial Reporting Quality",
   "Supply Chains",
   "Firm Productivity"
  ],
  "datasets": [
   {
    "provider": "Supply-chain relationship data",
    "product": null,
    "description": "Customer-supplier links used to identify U.S. manufacturers exposed to Taiwanese supply-chain disruption, combined with firm financials to compute productivity.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "supply chain; customer-supplier; productivity"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "1999 Taiwan (Chi-Chi) earthquake",
   "type": "Natural disaster",
   "what": "The 1999 Taiwan earthquake, which disrupted supply chains for certain U.S. high-technology manufacturers, used as an exogenous adverse shock whose productivity effects vary with firms' pre-shock reporting quality."
  },
  "missing_notes": null,
  "std_name": "Journal of Accounting and Economics - 2026 - Berger and Tomy - Supply chain shocks and firm productivity The role of reporting quality",
  "orig_filename": "1-s2.0-S0165410125000692-main.pdf"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Armstrong",
   "Glaeser"
  ],
  "title": "Does taxpayer assistance encourage entrepreneurship?",
  "summary": "This paper examines whether help with tax filing and compliance affects entrepreneurship, focusing on IRS Taxpayer Assistance Centers that help taxpayers navigate the tax system. The authors find that Taxpayer Assistance Centers are positively associated with local entrepreneur entry and with overall Schedule C business income. They interpret this as evidence that assistance encourages traditional business entrepreneurship by reducing the compliance costs that stem from tax complexity. The paper highlights an underappreciated channel—tax administration and support—through which the tax system affects business formation. It connects the accessibility of tax assistance to real entrepreneurial activity.",
  "logical_flow": "The paper starts from the idea that tax complexity imposes compliance costs that may deter people from starting businesses, and asks whether reducing those costs through assistance encourages entrepreneurship. It focuses on IRS Taxpayer Assistance Centers, which help taxpayers—including would-be entrepreneurs—file correctly and navigate tax rules. The authors reason that easier access to such help should lower the effective cost of entering self-employment and thus raise local business formation. They relate the presence of Taxpayer Assistance Centers to local entrepreneur entry and to Schedule C business income. Finding positive associations on both margins, they interpret assistance as lowering compliance barriers to entrepreneurship. The paper concludes that the administrative side of the tax system, not just tax rates, shapes business formation.",
  "research_design": "An empirical design relating the local presence of IRS Taxpayer Assistance Centers to entrepreneurship outcomes, using IRS Statistics of Income Schedule C data and measures of local business entry. The analysis associates access to tax assistance with entrepreneur entry and business income across localities; the unit of analysis is the locality/area. The interpretation is that assistance reduces tax-compliance costs to entry.",
  "categories": [
   "Taxation",
   "Entrepreneurship",
   "Public Economics"
  ],
  "datasets": [
   {
    "provider": "IRS",
    "product": "Taxpayer Assistance Center locations",
    "description": "Locations and coverage of IRS Taxpayer Assistance Centers used to measure local access to tax help.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "IRS; tax assistance; access"
   },
   {
    "provider": "IRS",
    "product": "Statistics of Income (Schedule C)",
    "description": "Local self-employment/Schedule C business income and entry data used to measure entrepreneurship.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "Schedule C; entrepreneurship; income"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Accounting and Economics - 2026 - Armstrong and Glaeser - Does taxpayer assistance encourage entrepreneurship",
  "orig_filename": "1-s2.0-S0165410125000758-main.pdf"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Kim",
   "Valentine"
  ],
  "title": "Earnings targets, strategic patent sales, and patent trolls",
  "summary": "This paper shows that innovative public firms sell significantly more patents in the last month of their fiscal year than earlier, consistent with reporting incentives driving these sales. Final-month sales are disproportionately internally generated, non-core patents with high potential accounting gains, and are more pronounced among firms with incentives to meet earnings expectations—especially when firms narrowly beat expectations and when executive incentives dominate. Managers also increase insider equity sales following these strategic patent sales. The paper documents earnings-motivated asset sales in the specific context of patents, some of which flow to patent-assertion entities. It links financial-reporting incentives to real decisions about intellectual property.",
  "logical_flow": "The paper begins from the observation that firms can realize accounting gains by selling assets, and asks whether earnings incentives drive the timing and composition of patent sales. It documents a sharp increase in patent sales in the final month of the fiscal year, a pattern hard to reconcile with pure operating motives but consistent with meeting reporting targets. Examining which patents are sold, it finds they are disproportionately internally generated, non-core, and carry high potential accounting gains—exactly the assets useful for boosting reported earnings. The authors then show these sales concentrate among firms with strong incentives to hit earnings benchmarks, particularly narrow beaters and firms where executive incentives are salient. They further link the sales to subsequent insider equity sales, reinforcing the reporting-and-incentive interpretation. The paper concludes that financial-reporting incentives shape patent divestitures, with some patents ending up in the hands of patent-assertion entities.",
  "research_design": "An empirical design documenting within-fiscal-year timing of patent sales and relating it to earnings-management incentives, using USPTO patent assignment (reassignment) records linked to firm financials and earnings-expectation data. Tests compare final-month versus earlier sales, characterize the patents sold, and split by benchmark-beating and executive-incentive proxies. The unit of analysis is the firm (and patent sale).",
  "categories": [
   "Earnings Management",
   "Innovation & Intellectual Property",
   "Financial Reporting"
  ],
  "datasets": [
   {
    "provider": "USPTO",
    "product": "Patent assignment (reassignment) records",
    "description": "Records of patent ownership transfers used to identify the timing and characteristics of firms' patent sales.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "patents; assignments; sales"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Accounting and Economics - 2026 - Kim and Valentine - Earnings targets, strategic patent sales, and patent trolls",
  "orig_filename": "1-s2.0-S0165410125000898-main.pdf"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Weber",
   "Xu",
   "Zhang"
  ],
  "title": "SEC scrutiny and corporate risk-taking",
  "summary": "This paper examines whether heightened SEC scrutiny of financial reporting affects corporate risk-taking. Tighter oversight could raise risk-taking by strengthening accounting's governance role, or reduce it by constraining managers' ability to shield themselves through earnings management. Using the SEC's 2007 elevation of six district offices to regional status—which expanded their authority and enforcement resources—in a difference-in-differences design, the authors identify the effect on firms in the affected jurisdictions. The results speak to how regulatory monitoring shapes managerial risk choices. The paper connects the geography of SEC enforcement to real corporate decisions.",
  "logical_flow": "The paper poses a genuinely two-sided question: because managers are risk-averse and can hide poor outcomes through earnings management, stronger SEC oversight might either encourage risk-taking (by improving governance and reducing the payoff to concealment) or discourage it (by limiting managers' ability to manage earnings and thus their willingness to take risks). To adjudicate, it needs exogenous variation in scrutiny, which it finds in the SEC's 2007 reorganization that elevated six district offices to regional level, expanding their enforcement reach. Firms headquartered in the newly elevated offices' jurisdictions become treated, while others serve as controls. The authors compare risk-taking outcomes for treated versus control firms around the reorganization in a difference-in-differences framework. The estimated effect reveals which force dominates in practice. The paper interprets the result as evidence on how the intensity and reach of regulatory monitoring shapes corporate risk-taking.",
  "research_design": "A difference-in-differences design exploiting the SEC's 2007 elevation of six district offices to regional status, which expanded their authority and enforcement resources, as a shock to local SEC scrutiny. Firms in the elevated offices' jurisdictions are treated and compared to firms elsewhere; outcomes are measures of corporate risk-taking. The unit of analysis is the firm; identification comes from the office reorganization's geographic variation.",
  "categories": [
   "Financial Regulation",
   "Corporate Risk-Taking",
   "Enforcement"
  ],
  "datasets": [
   {
    "provider": "SEC",
    "product": "Office jurisdictions / enforcement geography",
    "description": "Assignment of firms to SEC district/regional offices and the 2007 reorganization, used to define treated jurisdictions.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "SEC; jurisdictions; enforcement"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "2007 SEC elevation of six district offices to regional status",
   "type": "Regulatory reorganization",
   "what": "A 2007 SEC reorganization that raised six district offices to regional level, expanding their authority and enforcement resources, used as a difference-in-differences shock to local financial-reporting scrutiny."
  },
  "missing_notes": null,
  "std_name": "Journal of Accounting and Economics - 2026 - Weber et al. - SEC scrutiny and corporate risk-taking",
  "orig_filename": "1-s2.0-S0165410125000916-main.pdf"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Hanlon",
   "Shroff",
   "Yoon"
  ],
  "title": "Taxes and Competition: Evidence from the airline industry",
  "summary": "This paper examines whether corporate tax cuts alter product-market competition by differentially affecting firms with high versus low tax burdens. Because tax cuts raise after-tax cash flows for profitable firms but do little for loss-making firms, they can shift competitive positions. Studying the 1986 Tax Reform Act, which cut the top corporate rate by 12 percentage points, and route-level price and quantity data from U.S. airlines, the authors find that profitable airlines cut ticket prices by 4.2% relative to loss-making rivals and gain 3.3 percentage points of market share. The results show taxes affect not only firm value but also competitive dynamics. The paper documents a competition channel of corporate taxation.",
  "logical_flow": "The paper argues that corporate taxes may reshape competition because a tax cut benefits profitable firms—who owe taxes—far more than loss-making firms, changing the two groups' relative cost positions. It uses the 1986 Tax Reform Act as a large, discrete cut in the top corporate tax rate, generating a differential shock across profitable and loss-making firms. The airline industry provides an ideal testing ground because route-level prices and quantities are observable, allowing precise measurement of competitive responses. The authors predict that after the Act, profitable airlines will use their improved after-tax cash flows to cut prices and gain share relative to loss-making rivals. The data confirm this: profitable airlines lower fares by 4.2% and gain 3.3 percentage points of market share relative to loss makers. The paper concludes that corporate taxation has real competitive consequences beyond its direct effect on firm value.",
  "research_design": "A difference-in-differences design using the 1986 Tax Reform Act as a differential shock to profitable versus loss-making firms, applied to route-level price and quantity data from the U.S. airline industry. Profitable airlines (which benefit from the rate cut) are compared to loss-making rivals on prices and market share around the reform; the unit of analysis is the airline-route. Identification comes from the differential tax impact across firms.",
  "categories": [
   "Taxation",
   "Industrial Organization",
   "Competition"
  ],
  "datasets": [
   {
    "provider": "U.S. Department of Transportation",
    "product": "DB1B (Airline Origin and Destination Survey)",
    "description": "Route-level airline ticket prices and quantities used to measure competitive responses to the tax reform.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "airlines; routes; prices; market share"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Tax Reform Act of 1986",
   "type": "Tax reform",
   "what": "The 1986 Tax Reform Act, which cut the top corporate tax rate by about 12 percentage points, benefiting profitable firms much more than loss-making ones and used to study competition between differentially taxed airlines."
  },
  "missing_notes": null,
  "std_name": "Journal of Accounting and Economics - 2026 - Hanlon et al. - Taxes and Competition Evidence from the airline industry",
  "orig_filename": "1-s2.0-S0165410126000169-main.pdf"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Huber",
   "Watts",
   "Zhu"
  ],
  "title": "Information flows in trading networks",
  "summary": "This paper studies the informational value of trading networks in over-the-counter markets using transaction-level corporate bond data. Investors with larger dealer networks make superior trading decisions before changes in credit fundamentals, earning better risk-adjusted performance. The authors trace this outperformance to trading connections where dealers are most likely to have access to novel credit-relevant information, supporting an interpretation of private information flowing through networks. The results show that whom an investor trades with, not just what they trade, carries information. The paper provides evidence on how information diffuses through OTC market structure.",
  "logical_flow": "The paper begins from the feature of OTC markets that trades occur through dealers, so an investor's set of dealer relationships forms a network through which information can travel. It asks whether investors with more or better-connected dealer networks obtain an informational advantage. Using detailed transaction-level bond data that reveal who trades with whom, the authors measure investors' dealer networks and their trading performance around changes in credit fundamentals. They hypothesize that investors with larger networks—especially connections to dealers likely to see novel credit information—should trade profitably ahead of fundamental changes. The data show these investors do make superior, better-risk-adjusted trades before credit news, and the advantage is concentrated in information-rich connections. The paper concludes that trading networks are conduits for private information in OTC markets.",
  "research_design": "An empirical network design using transaction-level corporate bond data (with dealer identities) to construct investors' dealer networks and relate network size/position to trading performance around credit-fundamental changes. The analysis links network characteristics to risk-adjusted returns and isolates information-rich connections; the unit of analysis is the investor (and trade). Identification of information flows comes from performance concentrated in connections where dealers likely hold novel information.",
  "categories": [
   "Market Microstructure",
   "Information & Networks",
   "Fixed Income"
  ],
  "datasets": [
   {
    "provider": "FINRA",
    "product": "TRACE (enhanced, with dealer identifiers)",
    "description": "Transaction-level corporate bond trades including masked dealer identities, used to reconstruct trading networks and measure performance.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "corporate bonds; dealers; networks"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Accounting and Economics - 2026 - Huber et al. - Information flows in trading networks",
  "orig_filename": "1-s2.0-S0165410126000194-main.pdf"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Liu"
  ],
  "title": "Website cookies and voluntary disclosure",
  "summary": "This paper uses website cookie information from U.S. firms' websites to measure the consumer data firms collect and examines how such data affect their accounting information environments. Cookies provide granular, real-time consumer data that can improve the quality of firms' internal information for external reporting. Firms with more cookies issue management sales forecasts more frequently and devote more of their 10-K disclosures to customer, marketing, and product topics. The effects are stronger when cookie-collected data are more relevant to the firm. The paper links a novel, technology-based measure of consumer-data collection to firms' disclosure behavior.",
  "logical_flow": "The paper starts from the idea that firms increasingly collect granular consumer data online, and that such data may improve the internal information managers use for external reporting. It proposes website cookies as a measurable proxy for the intensity of consumer-data collection, since cookies capture real-time behavior of a firm's website visitors. The author reasons that firms with richer cookie-based data have better internal information about demand and customers, which should show up in more frequent and more customer-oriented disclosure. Measuring cookies across U.S. firms' websites, the paper relates cookie intensity to the frequency of management sales forecasts and to the content of 10-K filings. It finds firms with more cookies forecast sales more often and devote more disclosure to customer, marketing, and product topics, with stronger effects where the data are more relevant. The paper concludes that consumer-data collection technology shapes firms' voluntary disclosure and information environment.",
  "research_design": "An empirical design using website cookie data as a proxy for firms' consumer-data collection, related to disclosure outcomes (management forecast frequency and 10-K content). Cross-sectional and within-firm variation in cookie intensity is linked to disclosure behavior, with heterogeneity by data relevance; the unit of analysis is the firm. The novelty is the cookie-based measure of internal information inputs.",
  "categories": [
   "Voluntary Disclosure",
   "Data & Technology",
   "Information Environment"
  ],
  "datasets": [
   {
    "provider": "Website cookie / web-tracking data",
    "product": null,
    "description": "Data on cookies and trackers deployed on U.S. firms' websites, used to proxy the intensity of consumer-data collection.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "cookies; web tracking; consumer data"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Accounting and Economics - 2026 - Liu - Website cookies and voluntary disclosure",
  "orig_filename": "1-s2.0-S0165410126000285-main.pdf"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Guay",
   "Kim",
   "Timmermans"
  ],
  "title": "Internal information quality and performance metric selection",
  "summary": "This paper examines how firms' internal information quality (IIQ) shapes the design of executive incentive contracts. Higher IIQ is associated with using more performance metrics and with greater dissimilarity from peers' contracts, especially along non-financial dimensions. These relations hold using changes in IIQ likely induced by plausibly exogenous shifts in two financial accounting standards. Incorporating more numerous and more dissimilar non-financial metrics is positively associated with future profitability, but only when IIQ is high. The paper links the quality of firms' internal information to how they structure managerial incentives and to the payoff from doing so.",
  "logical_flow": "The paper argues that designing executive incentives around multiple, tailored performance metrics requires firms to have high-quality internal information to measure those metrics reliably. It hypothesizes that firms with better internal information quality will use more performance metrics and design contracts that differ more from peers', particularly on harder-to-measure non-financial dimensions. Because IIQ is endogenous, the authors use changes induced by plausibly exogenous shifts in two financial accounting standards to sharpen identification. They then relate IIQ to the number and dissimilarity of performance metrics in executive contracts. They further test whether these design choices pay off, finding that more numerous and dissimilar non-financial metrics are associated with higher future profitability only when IIQ is high. The paper concludes that internal information quality is a key determinant of, and complement to, sophisticated incentive-contract design.",
  "research_design": "An empirical design relating firms' internal information quality to executive performance-metric choices, using changes in information quality induced by two plausibly exogenous accounting-standard changes for identification. Outcomes include the number and peer-dissimilarity of performance metrics and their association with future profitability; the unit of analysis is the firm. Identification leans on the standard-change-induced shifts in information quality.",
  "categories": [
   "Executive Compensation",
   "Internal Information Quality",
   "Corporate Governance"
  ],
  "datasets": [
   {
    "provider": "ISS Incentive Lab",
    "product": "Executive incentive-plan data",
    "description": "Detailed executive compensation and performance-metric data used to measure the number and dissimilarity of metrics in incentive contracts.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "executive pay; performance metrics"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Two financial accounting standard changes",
   "type": "Accounting-standard changes",
   "what": "Plausibly exogenous shifts in two financial accounting standards that changed firms' internal information quality, used to identify the effect of information quality on executive performance-metric design."
  },
  "missing_notes": null,
  "std_name": "Journal of Accounting and Economics - 2026 - Guay et al. - Internal information quality and performance metric selection",
  "orig_filename": "1-s2.0-S0165410126000376-main.pdf"
 },
 {
  "journal": "Journal of Accounting Research",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Boulland",
   "Bourveau",
   "Breuer"
  ],
  "title": "Company Websites: A New Measure of Disclosure",
  "summary": "This paper proposes a new measure of firms' disclosure based on the content of their company websites, which are widely available and information-rich. Using historical website data for U.S. public firms, the authors construct the measure, validate it against existing disclosure and information-asymmetry proxies, and study its determinants. They then apply it to settings where standard measures are unavailable—U.S. private firms and French firms' compliance with a nonfinancial disclosure mandate. The applications show the website-based measure captures disclosure where other data cannot. The paper offers a broadly applicable, low-cost measure of corporate disclosure.",
  "logical_flow": "The paper notes that existing disclosure measures rely on filings or databases that are unavailable for many firms, such as private companies, limiting research. It proposes that company websites, which nearly all firms maintain, contain rich disclosure that can be measured from archived web data. The authors build the measure from historical website snapshots and validate it by showing it correlates with established disclosure and information-asymmetry proxies for public firms. Having established credibility, they apply it where standard data are missing—characterizing U.S. private firms' disclosure and testing French firms' compliance with a nonfinancial disclosure mandate. These applications demonstrate the measure's reach. The paper concludes that website content is a practical, general source for measuring disclosure across firm types and jurisdictions.",
  "research_design": "A measurement paper that constructs a website-based disclosure metric from archived (historical) website data for U.S. public firms, validates it against existing disclosure and information-asymmetry measures, and applies it to U.S. private firms and to French firms subject to a nonfinancial disclosure mandate. The unit of analysis is the firm; validation is cross-sectional, and the French application uses the mandate as a setting.",
  "categories": [
   "Disclosure Measurement",
   "Text & Web Data",
   "Financial Reporting"
  ],
  "datasets": [
   {
    "provider": "Internet Archive",
    "product": "Historical website (Wayback) data",
    "description": "Archived snapshots of firms' company websites used to construct the disclosure measure.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "websites; disclosure; web archive"
   },
   {
    "provider": "French nonfinancial disclosure filings",
    "product": null,
    "description": "French firms' nonfinancial disclosures used to test compliance with a mandate as an application of the measure.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "France; nonfinancial disclosure"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "French nonfinancial disclosure mandate",
   "type": "Disclosure regulation",
   "what": "A French mandate requiring firms to provide nonfinancial disclosure, used as a setting to validate the website-based measure through firms' compliance."
  },
  "missing_notes": null,
  "std_name": "Journal of Accounting Research - 2026 - Boulland et al. - Company Websites A New Measure of Disclosure",
  "orig_filename": "J of Accounting Research - 2025 - BOULLAND - Company Websites A New Measure of Disclosure.pdf"
 },
 {
  "journal": "Journal of Accounting Research",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Caskey",
   "Huang",
   "Saavedra"
  ],
  "title": "Amendment Thresholds and Voting Rules in Debt Contracts",
  "summary": "This paper studies the optimal voting rule that lets lenders waive a covenant violation. When lenders have heterogeneous preferences, lenient voting rules raise the chance of waivers that permit inefficient investment, while stringent rules can allocate the pivotal vote to lenders who deny waivers after false alarms in order to renegotiate and extract value, incurring deadweight costs. The optimal rule balances these forces to improve contracting efficiency. The authors derive comparative statics on how the optimal voting rule varies with contract and firm characteristics and test them empirically. The paper explains why debt contracts use the amendment thresholds and voting rules they do.",
  "logical_flow": "The paper begins from the observation that debt contracts specify voting rules governing how lenders can waive covenant violations, and asks what makes a voting rule optimal. It sets up a model with lenders who have heterogeneous preferences over waiving, so the choice of threshold determines whose vote is pivotal. Lenient rules make waivers easy but risk approving inefficient investments; stringent rules can hand the marginal vote to holdout lenders who deny waivers to renegotiate and extract value, which is costly. The optimal rule trades off these two inefficiencies. The authors derive comparative statics predicting how the optimal threshold shifts with contract and firm characteristics. They then test these predictions in data on debt-contract voting provisions, linking theory to observed contract design.",
  "research_design": "A contract-theory model of lender voting rules for covenant waivers, yielding comparative statics that are then tested empirically on debt-contract provisions. The model characterizes the optimal amendment threshold as a balance between approving inefficient investment and enabling value-extracting renegotiation; the empirical tests relate observed voting rules to firm and contract characteristics. The unit of analysis is the debt contract.",
  "categories": [
   "Debt Contracts & Covenants",
   "Contract Theory",
   "Corporate Finance"
  ],
  "datasets": [
   {
    "provider": "Debt-contract voting/amendment provisions (EDGAR)",
    "product": null,
    "description": "Hand-collected voting-rule and amendment-threshold terms from credit agreements filed with the SEC, used to test the model's comparative statics.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "credit agreements; voting rules; covenants"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Accounting Research - 2026 - Caskey et al. - Amendment Thresholds and Voting Rules in Debt Contracts",
  "orig_filename": "J of Accounting Research - 2025 - CASKEY - Amendment Thresholds and Voting Rules in Debt Contracts.pdf"
 },
 {
  "journal": "Journal of Accounting Research",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Ferracuti",
   "Lind"
  ],
  "title": "Macroeconomic Information Acquisition Around Earnings Clusters",
  "summary": "This paper shows that investors gather more macroeconomic information during earnings clusters—periods when many firms report earnings. This behavior is amplified during negative economic shocks and when macroeconomic announcements occur concurrently. The authors show these information-acquisition patterns have implications for equity valuations. The evidence suggests that clustered earnings news prompts investors to update their macro views, not just firm-specific ones. The paper links the timing of earnings news to macroeconomic learning and asset prices.",
  "logical_flow": "The paper starts from the idea that earnings announcements convey not only firm-specific but also macroeconomic information, especially when many firms report at once. It hypothesizes that during earnings clusters investors have more incentive to acquire macroeconomic information, because clustered reports are informative about the aggregate economy. Using proxies for information acquisition, the authors show macro information gathering rises during earnings clusters. They further show this rises more during negative economic shocks and alongside macroeconomic announcements, when aggregate uncertainty is higher. They then connect these acquisition patterns to equity valuations, indicating the learning affects prices. The paper concludes that clustered earnings news is a trigger for macroeconomic learning with asset-pricing consequences.",
  "research_design": "An empirical design relating measures of investor information acquisition (e.g., access to macro-relevant information) to the timing of earnings clusters, with amplification during negative economic shocks and concurrent macroeconomic announcements, and links to equity valuations. The unit of analysis is the time period/firm; identification relies on variation in earnings clustering and macro conditions.",
  "categories": [
   "Information Acquisition",
   "Macroeconomic Information",
   "Capital Markets"
  ],
  "datasets": [
   {
    "provider": "SEC EDGAR log files",
    "product": null,
    "description": "Records of investors' access to firms' filings, used to proxy information acquisition around earnings clusters.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "EDGAR logs; information acquisition"
   },
   {
    "provider": "RavenPack",
    "product": null,
    "description": "Macroeconomic news and announcement data used to identify concurrent macro information.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "macro news; announcements"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Accounting Research - 2026 - Ferracuti and Lind - Macroeconomic Information Acquisition Around Earnings Clusters",
  "orig_filename": "J of Accounting Research - 2025 - FERRACUTI - Macroeconomic Information Acquisition Around Earnings Clusters.pdf"
 },
 {
  "journal": "Journal of Accounting Research",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Ferrell",
   "Manconi",
   "Neretina",
   "Powley",
   "Renneboog"
  ],
  "title": "Corporate Litigation, Governance, and the Role of Law Firms",
  "summary": "This paper studies plaintiff law firms in corporate litigation, focusing on 'star' firms that dominate settlement outcomes. Star firms are associated with larger settlements, but much of this is predicted by the defendant's litigation-insurance coverage, suggesting assortative matching of stars with lawsuits that have larger expected payoffs. Stars also charge higher fees for a given settlement size. Additional tests suggest visibility and information asymmetry toward less sophisticated plaintiffs sustain the stars' market position. The paper illuminates how the market for plaintiff legal services shapes corporate-litigation outcomes.",
  "logical_flow": "The paper focuses on plaintiff law firms as key but understudied players in corporate litigation, asking why a few 'star' firms dominate. It first documents that star firms obtain larger settlements, which could reflect skill or selection. To distinguish these, it shows much of the star-settlement relation is explained by the defendant's litigation-insurance coverage, implying stars match with cases that already have larger expected payoffs. This points to assortative matching rather than pure skill. The authors then show stars charge higher fees for a given settlement, indicating market power. Tests on visibility and information asymmetry toward less sophisticated plaintiffs suggest how stars sustain that power. The paper concludes that the structure of the plaintiff-lawyer market, including matching and information frictions, shapes litigation outcomes.",
  "research_design": "An empirical study of plaintiff law firms in corporate litigation, relating 'star' firm involvement to settlement size, fees, and defendant litigation-insurance coverage to separate skill from assortative matching. Additional tests examine visibility and plaintiff sophistication as sources of star firms' market power. The unit of analysis is the lawsuit/law-firm.",
  "categories": [
   "Corporate Litigation",
   "Governance",
   "Legal Services"
  ],
  "datasets": [
   {
    "provider": "Corporate litigation and settlement data",
    "product": null,
    "description": "Data on corporate/securities lawsuits, settlements, plaintiff law firms, and defendants' litigation-insurance coverage, used to study star law firms.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "litigation; settlements; law firms"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Accounting Research - 2026 - Ferrell et al. - Corporate Litigation, Governance, and the Role of Law Firms",
  "orig_filename": "J of Accounting Research - 2025 - FERRELL - Corporate Litigation Governance and the Role of Law Firms.pdf"
 },
 {
  "journal": "Journal of Accounting Research",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Gao",
   "Lu"
  ],
  "title": "Aggregated Compensation Peer Group Disclosure and Managerial Labor Market Competition: A Network Analysis",
  "summary": "This paper uses network analysis of firms' disclosed compensation peer groups to study managerial labor-market competition. By aggregating peer-group disclosures across firms, the authors map the network of which firms benchmark against which, and use its structure to characterize competition for executive talent. The analysis links a firm's position in the peer network to compensation and labor-market outcomes. The paper shows that compensation peer-group disclosures, taken together, reveal the structure of the managerial labor market. It provides a network-based lens on executive-pay benchmarking.",
  "logical_flow": "The paper starts from the fact that firms disclose the peer groups they use to benchmark executive pay, and argues these disclosures, aggregated across firms, form a network revealing labor-market linkages. It constructs this network by connecting firms that name each other (or common peers) in compensation benchmarking. It then uses network structure—centrality, clustering, and connections—to characterize competition for executive talent. Relating network position to pay and outcomes shows that where a firm sits in the network matters for compensation. The authors interpret the network as a map of managerial labor-market competition that individual disclosures cannot reveal in isolation. The paper concludes that aggregated peer-group disclosures are informative about the structure of the executive labor market.",
  "research_design": "A network-analysis design that aggregates firms' disclosed compensation peer groups into a network and relates network structure and firm position to compensation and labor-market outcomes. The unit of analysis is the firm within the peer network; identification is descriptive/associational from network structure.",
  "categories": [
   "Executive Compensation",
   "Peer Benchmarking",
   "Networks"
  ],
  "datasets": [
   {
    "provider": "Compensation peer-group disclosures (SEC proxy filings)",
    "product": null,
    "description": "Firms' disclosed executive-compensation peer groups from proxy statements, aggregated into a benchmarking network.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "peer groups; executive pay; proxy"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Accounting Research - 2026 - Gao and Lu - Aggregated Compensation Peer Group Disclosure and Managerial Labor Market Competition A Network Analysis",
  "orig_filename": "J of Accounting Research - 2025 - GAO - Aggregated Compensation Peer Group Disclosure and Managerial Labor Market.pdf"
 },
 {
  "journal": "Journal of Accounting Research",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Liu",
   "Madsen",
   "Zhou"
  ],
  "title": "Strategic (Inconsistent) Disclosures and Sophisticated Investors: Evidence from Hedge Funds",
  "summary": "This paper studies strategic, inconsistent disclosures by hedge-fund advisers. The authors find that 40% of these disclosures omit or de-emphasize information about advisers' operational and investment risks relative to other public sources. Funds with such inconsistencies subsequently deliver predictably lower performance, yet do not differ in fund flows, the flow-performance relation, ownership structure, or management. The results suggest even sophisticated investors do not fully penalize inconsistent disclosure. The paper documents strategic disclosure in a setting with ostensibly sophisticated investors.",
  "logical_flow": "The paper examines whether hedge-fund advisers strategically shade their disclosures, comparing what they tell investors against other public information about their risks. It finds a large share of disclosures omit or downplay operational and investment risks, indicating strategic inconsistency. The authors then ask whether these inconsistencies are informative about outcomes, and find funds with inconsistent disclosures earn predictably lower returns. Strikingly, investors do not appear to react—flows, the flow-performance relation, ownership, and management look similar—suggesting the inconsistency is not fully priced. This is notable because hedge-fund investors are typically considered sophisticated. The paper concludes that strategic disclosure persists and is under-penalized even among sophisticated investors.",
  "research_design": "An empirical design comparing hedge-fund advisers' disclosures (from regulatory filings) against other public information to flag 'inconsistent' disclosures, then relating inconsistency to future fund performance and to investor responses (flows, flow-performance, ownership). The unit of analysis is the fund/adviser; identification is cross-sectional.",
  "categories": [
   "Disclosure & Hedge Funds",
   "Investor Sophistication",
   "Financial Reporting"
  ],
  "datasets": [
   {
    "provider": "SEC Form ADV (hedge-fund adviser filings)",
    "product": null,
    "description": "Investment-adviser regulatory disclosures of operational and investment risks, compared with other public information to identify inconsistencies.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "Form ADV; hedge funds; disclosure"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Accounting Research - 2026 - Liu et al. - Strategic (Inconsistent) Disclosures and Sophisticated Investors Evidence from Hedge Funds",
  "orig_filename": "J of Accounting Research - 2025 - LIU - Strategic Inconsistent Disclosures and Sophisticated Investors Evidence from.pdf"
 },
 {
  "journal": "Journal of Accounting Research",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Armstrong",
   "Glaeser",
   "Park"
  ],
  "title": "The Assignment of Intellectual Property Rights and Innovation",
  "summary": "This paper studies how the assignment of intellectual-property rights between inventors and their employers affects innovation. Incomplete-contracting theory predicts that stronger employer property rights reduce the threat that employee inventors hold up their employers, affecting inventor and invention outcomes. The authors test this using a U.S. appellate court ruling that shifted the assignment of patent rights from inventors to employers. Within-employer-year analyses show affected inventors are less likely to retain patent rights or assign patents to new employers. The paper provides causal evidence on how IP-rights allocation shapes innovation.",
  "logical_flow": "The paper draws on incomplete-contracting theory, which holds that who owns the rights to an invention affects bargaining and hold-up between inventors and employers, and thus innovation. It predicts that strengthening employer property rights reduces inventors' ability to hold up employers, changing inventor mobility and invention outcomes. To test this causally, the authors exploit a U.S. appellate court ruling that shifted the default assignment of patent rights from inventors toward employers, providing an exogenous change in property rights. Using within-employer-year comparisons, they examine how affected inventors' retention and reassignment of patents change. Finding affected inventors less able to retain or move patents, they confirm the theory's predictions. The paper concludes that the legal allocation of IP rights materially shapes inventor behavior and innovation.",
  "research_design": "A natural-experiment design exploiting a U.S. appellate court ruling that shifted the default assignment of patent rights from employee inventors to employers, with within-employer-year comparisons of affected versus unaffected inventors. Outcomes include inventors' retention of patent rights and reassignment to new employers. The unit of analysis is the inventor (within employer-year); identification comes from the court ruling.",
  "categories": [
   "Intellectual Property",
   "Innovation",
   "Law and Finance"
  ],
  "datasets": [
   {
    "provider": "USPTO patent and inventor data",
    "product": null,
    "description": "Patents, inventors, and assignment records used to track inventors' retention and reassignment of rights around the ruling.",
    "access_type": "Public",
    "delivery": "Bulk",
    "topic_tags": "patents; inventors; assignments"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "U.S. appellate court ruling on IP-rights assignment",
   "type": "Court decision",
   "what": "An appellate court ruling that shifted the default assignment of patent rights from employee inventors to their employers, used as a natural experiment on employer property rights."
  },
  "missing_notes": null,
  "std_name": "Journal of Accounting Research - 2026 - Armstrong et al. - The Assignment of Intellectual Property Rights and Innovation",
  "orig_filename": "J of Accounting Research - 2026 - Armstrong - The Assignment of Intellectual Property Rights and Innovation.pdf"
 },
 {
  "journal": "Journal of Accounting Research",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Billings",
   "Holthausen",
   "Petrovits",
   "Wang"
  ],
  "title": "Incidence, Risk, and Disclosure of Corporate Litigation: Insights from Federal Court Filings",
  "summary": "This paper assembles and describes a comprehensive sample of 174,782 lawsuits filed against 218,437 public-company defendants in U.S. federal district court from 2006 to 2021. The suits span product liability, civil-rights discrimination, contract breaches, labor practices, antitrust, corruption, securities violations, pollution, and IP infringement. The sample shows rich variation across firms, industries, time, suit types, plaintiffs, and outcomes. The authors use it to characterize the incidence, risk, and disclosure of corporate litigation. The paper provides a broad, systematic picture of public firms' exposure to federal litigation.",
  "logical_flow": "The paper is motivated by the fact that research on corporate litigation often focuses narrowly on securities class actions, missing the full breadth of lawsuits firms face. To address this, the authors build a large sample of federal district court filings against public companies over 2006-2021, covering many allegation types. They document the sample's rich variation across firms, industries, suit types, plaintiffs, and outcomes, establishing basic facts about corporate litigation. They then examine the incidence and risk of litigation across firms and how firms disclose it. The descriptive breadth allows comparisons that narrower samples cannot support. The paper concludes by providing a systematic, comprehensive characterization of public firms' federal-litigation exposure and its disclosure.",
  "research_design": "A descriptive, data-construction study assembling 174,782 federal district court lawsuits against 218,437 public-company defendants (2006-2021) and characterizing the incidence, risk, and disclosure of corporate litigation across firms, industries, suit types, and outcomes. The unit of analysis is the lawsuit/firm; the contribution is comprehensive measurement rather than a single identification strategy.",
  "categories": [
   "Corporate Litigation",
   "Disclosure",
   "Law and Finance"
  ],
  "datasets": [
   {
    "provider": "U.S. federal district court filings (PACER)",
    "product": null,
    "description": "Comprehensive civil case filings against public-company defendants (2006-2021), spanning many allegation types, used to measure litigation incidence and outcomes.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "federal courts; lawsuits; PACER"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Journal of Accounting Research - 2026 - Billings et al. - Incidence, Risk, and Disclosure of Corporate Litigation Insights from Federal Court Filings",
  "orig_filename": "J of Accounting Research - 2026 - BILLINGS - Incidence Risk and Disclosure of Corporate Litigation Insights from Federal.pdf"
 },
 {
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Bernstein"
  ],
  "title": "The Value of Ratings: Evidence from Their Introduction in Securities Markets",
  "summary": "This paper studies the first-ever ratings for corporate securities, using John Moody's 1909 publication that assigned letter grades to most listed railroad bonds. These ratings had no regulatory implications and were largely explainable from public information, isolating their pure information role. The author finds that bonds rated lower than the market implied saw rising secondary-market yields, and that rated bonds experienced a substantial decline in bid-ask spreads, consistent with reduced information asymmetry and improved liquidity. The results show ratings can improve information transmission even without regulatory force. The paper provides clean historical evidence on the value of ratings.",
  "logical_flow": "The paper addresses the difficulty of isolating the informational value of credit ratings, since modern ratings are entangled with regulation. It turns to a historical natural experiment: John Moody's 1909 introduction of the first letter ratings for railroad bonds, which carried no regulatory implications and drew on public information. Because the ratings had no mechanical regulatory effect, any market reaction reflects their information content. The author examines how bond yields and liquidity responded to being rated and to receiving ratings below market-implied levels. Finding that lower-than-expected ratings raised yields and that rated bonds saw narrower bid-ask spreads, the paper attributes these to reduced information asymmetry. It concludes that ratings can improve information transmission and liquidity on their own informational merits.",
  "research_design": "A historical natural-experiment design using John Moody's 1909 introduction of the first-ever corporate (railroad) bond ratings, which had no regulatory implications, to isolate the information value of ratings. The author compares secondary-market yields and bid-ask spreads for rated versus unrated bonds and for bonds rated below market-implied levels. The unit of analysis is the bond; identification comes from the rating introduction.",
  "categories": [
   "Credit Ratings",
   "Liquidity & Microstructure",
   "Financial History"
  ],
  "datasets": [
   {
    "provider": "Moody's (1909)",
    "product": "First-ever railroad bond ratings",
    "description": "Hand-collected letter ratings from John Moody's 1909 publication covering most listed railroad bonds.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "ratings; railroad bonds; 1909"
   },
   {
    "provider": "Historical railroad bond prices/yields",
    "product": null,
    "description": "Secondary-market prices, yields, and bid-ask spreads for early-1900s railroad bonds, used to measure market reactions to ratings.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "bond yields; bid-ask; historical"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "1909 introduction of Moody's bond ratings",
   "type": "Historical natural experiment",
   "what": "John Moody's 1909 publication of the first-ever letter ratings for railroad bonds, which had no regulatory implications, used as a natural experiment isolating the information value of ratings."
  },
  "missing_notes": null,
  "std_name": "Working Paper - 2026 - Bernstein - The Value of Ratings Evidence from Their Introduction in Securities Markets",
  "orig_filename": "739333.pdf"
 },
 {
  "journal": null,
  "is_working_paper": true,
  "year": 2025,
  "authors": [
   "Cabral",
   "Dillender"
  ],
  "title": "Doctor Discretion in Medical Evaluations",
  "summary": "This paper analyzes how the discretion of doctors conducting independent medical exams affects injured workers' outcomes. Exploiting the quasi-random assignment of claimants to more- or less-generous independent examiners, the authors trace downstream outcomes over three years—duration out of work, cash disability benefits, and subsequent medical care. Reduced-form estimates show that claimants assigned a more generous examiner experience longer durations out of work and receive more benefits and care. The results reveal that examiner discretion has large, persistent effects on disability outcomes. The paper documents the consequences of doctor discretion in a high-stakes evaluation setting.",
  "logical_flow": "The paper studies independent medical exams, where a doctor's assessment shapes an injured worker's eligibility for disability benefits, and asks how much the specific examiner matters. It exploits the fact that claimants are effectively randomly assigned to examiners who differ in generosity, creating an examiner-leniency design. Because assignment is quasi-random, differences in outcomes across examiners of different generosity can be attributed to the examiner rather than the claimant. The authors follow claimants for three years, measuring duration out of work, cash benefits, and medical care. They find that being assigned a more generous examiner leads to longer time out of work and more benefits and care, though these also depend on later actions by claimants, insurers, and treating doctors. The paper concludes that doctor discretion in evaluations has substantial and lasting effects on disability outcomes.",
  "research_design": "An examiner-assignment (leniency) design exploiting the quasi-random assignment of injured workers to independent medical examiners who differ in generosity, using administrative data. Reduced-form estimates relate assigned-examiner generosity to downstream outcomes—duration out of work, cash disability benefits, and subsequent medical care—over three years. The unit of analysis is the claimant; identification comes from quasi-random examiner assignment.",
  "categories": [
   "Health Economics",
   "Disability Insurance",
   "Public Economics"
  ],
  "datasets": [
   {
    "provider": "Administrative workers' compensation / disability claims",
    "product": null,
    "description": "Administrative records of injured workers, their independent medical exam assignments, benefits, and medical care, used in the examiner-assignment design.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "workers' compensation; disability; medical exams"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Quasi-random assignment of independent medical examiners",
   "type": "Examiner-leniency (random assignment)",
   "what": "The quasi-random assignment of injured workers to more- or less-generous independent medical examiners, used as an examiner-leniency design to identify the effect of medical evaluations on disability outcomes."
  },
  "missing_notes": null,
  "std_name": "Working Paper - 2025 - Cabral and Dillender - Doctor Discretion in Medical Evaluations",
  "orig_filename": "CabralDillender_DoctorDiscretion.pdf"
 },
 {
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Ormazabal",
   "Shi"
  ],
  "title": "Industry (Re)Classifications and Corporate Governance",
  "summary": "This paper examines whether industry classifications affect how firms select peers for executive-compensation benchmarking, using Global Industry Classification Standard (GICS) reclassifications for identification. Firms become significantly more likely to include peers whose GICS codes converge with their own and to drop peers whose codes diverge, with stronger effects when proxy advisors and passive investors provide greater oversight. The authors find no evidence of strategic peer selection. The results show that formal industry classifications shape compensation benchmarking and are disciplined by governance. The paper links classification systems to executive-pay peer selection.",
  "logical_flow": "The paper asks whether the industry classification a firm is assigned influences which peers it chooses for executive-pay benchmarking. Because classifications and peer choices could be jointly determined, it needs exogenous variation, which it finds in GICS reclassifications that reassign firms to different industry codes for reasons outside their compensation decisions. The authors predict that after a reclassification, firms will add peers whose codes now converge with theirs and drop peers whose codes diverge. They confirm this and show the effect is stronger where proxy advisors and passive investors exert more oversight, indicating governance disciplines the process. Finding no evidence of strategic selection, they interpret the peer changes as governance-driven rather than opportunistic. The paper concludes that formal industry classifications causally shape compensation peer selection.",
  "research_design": "A quasi-experimental design exploiting GICS industry reclassifications as exogenous changes in firms' industry membership, examining how firms adjust their executive-compensation peer groups (adding converging, dropping diverging peers). Heterogeneity by proxy-advisor and passive-investor oversight tests the governance channel. The unit of analysis is the firm (and peer pair); identification comes from reclassification events.",
  "categories": [
   "Executive Compensation",
   "Corporate Governance",
   "Industry Classification"
  ],
  "datasets": [
   {
    "provider": "GICS reclassification data (MSCI/S&P)",
    "product": null,
    "description": "Changes in firms' Global Industry Classification Standard codes, used as exogenous variation in industry membership.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "GICS; reclassification; industry"
   },
   {
    "provider": "Compensation peer-group disclosures",
    "product": null,
    "description": "Firms' disclosed executive-compensation peer groups (proxy statements) used to measure peer additions and drops.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "peer groups; executive pay"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "GICS industry reclassifications",
   "type": "Classification change",
   "what": "Reassignments of firms' Global Industry Classification Standard codes, occurring for reasons outside firms' compensation decisions, used as exogenous variation in industry membership."
  },
  "missing_notes": null,
  "std_name": "Working Paper - 2026 - Ormazabal and Shi - Industry (Re)Classifications and Corporate Governance",
  "orig_filename": "OS-April-2026.pdf"
 },
 {
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Da",
   "Wang",
   "Zeng"
  ],
  "title": "Presidential Cycles in PEAD",
  "summary": "This paper documents that post-earnings announcement drift (PEAD) follows presidential cycles: it earns about 4.1% per year during Democratic presidencies but rises to 14.9% during Republican presidencies. Survey evidence indicates greater investor underreaction to earnings news under Republican presidents. The authors argue the tax component of earnings is more volatile during Republican periods, implying larger tax-policy uncertainty that amplifies information uncertainty and investor underreaction. Consistent with this, firms reference Republican tax laws more frequently in filings. The paper links political cycles and tax-policy uncertainty to a classic asset-pricing anomaly.",
  "logical_flow": "The paper starts from the well-known PEAD anomaly—prices drift in the direction of earnings surprises—and asks whether its strength varies with the political cycle. Documenting that PEAD is much larger under Republican than Democratic presidents, it seeks a mechanism rather than treating the pattern as coincidence. It proposes that tax-policy uncertainty, which affects the tax component of earnings, is greater during Republican periods, making earnings news harder to interpret and amplifying underreaction. Survey evidence of stronger underreaction under Republican presidents supports the behavioral channel. Showing that the tax component of earnings is more volatile and that firms reference Republican tax laws more often reinforces the tax-uncertainty explanation. The paper concludes that political cycles, operating through tax-policy uncertainty, systematically shape the magnitude of PEAD.",
  "research_design": "An empirical asset-pricing design documenting variation in post-earnings announcement drift across presidential (Republican vs. Democratic) periods, combined with survey evidence on underreaction and tests of the tax-policy-uncertainty mechanism (volatility of the tax component of earnings, references to tax laws in filings). The unit of analysis is the firm-earnings announcement over time; identification is associational across political regimes.",
  "categories": [
   "Market Anomalies",
   "Behavioral Finance",
   "Political Economy & Finance"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": null,
  "missing_notes": "Uses standard earnings and analyst (I/B/E/S) data plus survey-based expectations; no distinctive datasets identified.",
  "std_name": "Working Paper - 2026 - Da et al. - Presidential Cycles in PEAD",
  "orig_filename": "Presidential_PEAD-Mar-2026.pdf"
 },
 {
  "journal": "Review of Financial Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Zhong",
   "Zhou"
  ],
  "title": "Dynamic Coordination and Bankruptcy Regulations",
  "summary": "This paper studies how bankruptcy regulations that restrict creditors' ability to exit a distressed firm affect coordination among creditors. Such ex post restrictions can protect coordination once distress hits, but may harm creditors' ex ante incentives to stay invested, worsening outcomes. The authors build a dynamic coordination model to show how this tension shapes creditor runs, bankruptcy filings, and the design of regulation. A striking result is that filing for bankruptcy early, preserving more assets for latecomers, can prolong firm life, and that regulators' clawbacks on pre-bankruptcy repayments can dominate firms' commitment to early filing. The paper derives implications for automatic stay and avoidance provisions.",
  "logical_flow": "The paper starts from the classic coordination problem among creditors of a distressed firm, where each creditor may 'run' by exiting, hastening the firm's collapse. It observes that bankruptcy regulations often restrict exit ex post to preserve value, but points out this changes creditors' ex ante incentives to remain invested in the first place. Building a dynamic coordination model, it traces how the anticipation of exit restrictions feeds back into creditor runs and the timing of bankruptcy filings. The model yields the counterintuitive result that filing early, which conserves assets for later creditors, can extend firm life by improving coordination. It further shows that regulators using clawbacks on pre-bankruptcy repayments can achieve better outcomes than relying on firms to commit to early filing. The paper then draws out design implications for automatic stay and avoidance rules.",
  "research_design": "A dynamic coordination (theoretical) model of creditors deciding whether to remain invested in or exit a distressed firm, used to analyze creditor runs, the timing of bankruptcy filings, and the design of bankruptcy regulations. There is no external dataset; results are analytical, with implications derived for automatic stay, clawbacks, and avoidance provisions.",
  "categories": [
   "Bankruptcy & Restructuring",
   "Coordination & Contract Theory",
   "Financial Regulation"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": null,
  "missing_notes": "Theoretical model; no datasets.",
  "std_name": "Review of Financial Studies - 2026 - Zhong and Zhou - Dynamic Coordination and Bankruptcy Regulations",
  "orig_filename": "hhaf039.pdf"
 },
 {
  "journal": "Review of Financial Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Dávila",
   "Parlatore"
  ],
  "title": "Identifying Price Informativeness",
  "summary": "This paper shows how to identify and estimate price informativeness, a necessary step in testing theories of information aggregation. Starting from a pricing equation and a stochastic process for payoffs, the authors derive how to recover relative price informativeness from regressions of asset-price changes on payoff changes. Applying the method, they estimate stock-specific informativeness measures for U.S. stocks. In the cross-section, large stocks with high turnover, idiosyncratic volatility, institutional ownership, and analyst coverage are more informative. The paper provides a tractable, theory-grounded way to measure how much prices reveal about fundamentals.",
  "logical_flow": "The paper addresses a gap between theories of information aggregation, which hinge on how informative prices are, and the empirical difficulty of measuring price informativeness. It begins from a general pricing equation and an assumed stochastic process for payoffs, and shows analytically how informativeness maps into the relationship between price changes and payoff changes. This yields an identification result: informativeness can be recovered from regressions of asset-price changes on changes in fundamentals. The authors implement this to estimate a panel of stock-specific informativeness measures for U.S. stocks. They then characterize the cross-section, finding informativeness rises with size, turnover, idiosyncratic volatility, institutional ownership, and analyst coverage. The paper concludes by offering the method as a general tool for measuring and testing price informativeness.",
  "research_design": "A methodological/econometric contribution deriving identification of price informativeness from a pricing equation and payoff process, then estimating stock-specific informativeness for U.S. stocks via regressions of price changes on payoff changes. The estimates are characterized in the cross-section; the unit of analysis is the stock over time.",
  "categories": [
   "Market Efficiency & Information",
   "Asset Pricing",
   "Econometric Methods"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": null,
  "missing_notes": "Estimated from standard U.S. stock price/payoff data; no distinctive datasets.",
  "std_name": "Review of Financial Studies - 2026 - Dávila and Parlatore - Identifying Price Informativeness",
  "orig_filename": "hhaf051.pdf"
 },
 {
  "journal": "Review of Financial Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Mayordomo",
   "Pavanini",
   "Tarantino"
  ],
  "title": "The Impact of Bank Consolidation on Credit Supply and Performance",
  "summary": "This paper studies how bank mergers affect credit supply and bank performance, combining quasi-experimental evidence with a structural model. Reduced-form analysis shows merged banks restrict credit supply and set higher interest rates, but also reject fewer applicants and report fewer nonperforming loans. To interpret these offsetting effects, the authors estimate a structural model in which banks set both interest rates and lending standards. They find that despite relaxing lending standards, merged banks' credit performance improved because of a significant drop in screening costs. The paper shows that consolidation can raise prices yet improve screening efficiency.",
  "logical_flow": "The paper examines a central policy question—whether bank consolidation helps or harms borrowers—that is hard to answer because mergers change both prices and lending standards at once. It first uses a quasi-experimental design around bank mergers to document a mixed pattern: merged banks lend less and charge more, yet reject fewer applicants and have fewer nonperforming loans. Because these facts pull in different directions, a purely reduced-form reading is ambiguous. The authors therefore build a structural model in which banks jointly choose interest rates and lending standards, allowing them to separate pricing from screening. Estimating the model, they find merged banks relaxed standards but nonetheless improved credit performance thanks to lower screening costs. The paper concludes that consolidation's welfare effects hinge on efficiency gains in screening, not just on prices.",
  "research_design": "A two-part design pairing a quasi-experimental analysis of bank mergers with a structural model of bank credit in which banks set interest rates and lending standards. The reduced form compares merged and non-merged banks on credit supply, rates, rejections, and nonperforming loans; the structural model identifies screening costs and disentangles pricing from standards. The unit of analysis is the bank/loan, using a credit register.",
  "categories": [
   "Banking",
   "Credit Markets",
   "Industrial Organization"
  ],
  "datasets": [
   {
    "provider": "Banco de España",
    "product": "Spanish Credit Register (CIRBE)",
    "description": "Supervisory loan-level credit-register data on Spanish banks' lending, rates, and performance, used to study merger effects.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "credit register; Spain; bank mergers"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Bank mergers (consolidation events)",
   "type": "Quasi-experiment",
   "what": "Bank merger events used as quasi-experimental variation in local bank consolidation to identify effects on credit supply, pricing, and performance."
  },
  "missing_notes": null,
  "std_name": "Review of Financial Studies - 2026 - Mayordomo et al. - The Impact of Bank Consolidation on Credit Supply and Performance",
  "orig_filename": "hhaf107.pdf"
 },
 {
  "journal": "Review of Financial Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Burkart",
   "Lee",
   "Petri"
  ],
  "title": "The Structure of Leveraged Buyouts and the Free-Rider Problem",
  "summary": "This paper studies the structure of public-firm buyouts in a model featuring both the Berle-Means problem (weak incentives) and the Grossman-Hart problem (shareholder holdout/free-riding). The authors show that bootstrapping, debt in excess of funding needs, and upfront fees to bidders are socially optimal and increase buyout premiums. These elements make LBO financing akin to a 'management contract' arranged by an outside manager who receives cash and incentives to run the firm, funded by excess debt imposed on the target. The model also rationalizes why private-equity firms collect fees from their equity partnerships. The paper offers a unified explanation for common but puzzling LBO financing features.",
  "logical_flow": "The paper starts from two classic frictions in taking a public firm private: shareholders lack incentives to monitor (Berle-Means) and can hold out to free-ride on a bidder's value creation (Grossman-Hart). It asks whether the peculiar financing structures seen in leveraged buyouts—bootstrapping, debt beyond funding needs, and upfront fees to bidders—can be understood as solutions to these frictions. Building a model of buyouts, it shows these features are socially optimal and raise premiums, because they help overcome holdout and align incentives. The authors reinterpret LBO financing as effectively a management contract, where an outside manager receives cash and incentives to run the firm, funded by excess debt loaded onto the target. The model further explains why PE firms extract fees from their equity partnerships. The paper concludes that seemingly extractive LBO features can be efficiency-enhancing responses to free-riding.",
  "research_design": "A theoretical model of public-firm leveraged buyouts incorporating the Berle-Means (incentive) and Grossman-Hart (holdout/free-rider) problems, used to show that bootstrapping, excess debt, and upfront bidder fees are socially optimal and raise premiums. There is no external dataset; the analysis is analytical and reinterprets LBO financing as a management contract.",
  "categories": [
   "Private Equity & Buyouts",
   "Corporate Governance",
   "Financial Theory"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": null,
  "missing_notes": "Theoretical model; no datasets.",
  "std_name": "Review of Financial Studies - 2026 - Burkart et al. - The Structure of Leveraged Buyouts and the Free-Rider Problem",
  "orig_filename": "hhaf111.pdf"
 },
 {
  "journal": "Review of Financial Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Filipović",
   "Wagner"
  ],
  "title": "The Intangibles Song in Takeover Announcements: Good Tempo, Hollow Tune",
  "summary": "This paper develops a word list of intangibles and applies it to takeover announcements. One standard deviation more intangible-related language ('intangibles talk') lowers the acquirer's announcement return by 0.53 percentage points and predicts worse operating performance. Yet bidder managers appear to believe in these deals, as shown by their insider trades, payment choices, and higher completion probabilities and speed. The authors conclude that takeover-announcement texts reveal important information about hard-to-measure aspects of deal quality. The paper shows that how firms talk about intangibles in deal announcements is informative—often negatively—about deal outcomes.",
  "logical_flow": "The paper begins from the difficulty of assessing deal quality when much of the rationale for a takeover rests on hard-to-measure intangibles. It proposes measuring 'intangibles talk' by building a word list of intangibles and applying it to the text of takeover announcements. The authors find that more intangibles language is associated with lower acquirer announcement returns and predicts worse subsequent operating performance, suggesting such talk often signals weaker deals. Interestingly, bidder managers seem to genuinely believe in these deals, as revealed by their insider trading, choice of payment, and the higher probability and speed of completion. This gap between market skepticism and managerial conviction is itself informative. The paper concludes that announcement texts, and intangibles language in particular, carry real information about deal quality that markets partly read.",
  "research_design": "A text-as-data design that builds an intangibles word list and applies it to takeover-announcement texts, relating 'intangibles talk' to acquirer announcement returns and future operating performance. Managerial belief is examined through insider trades, payment choices, and completion probability/speed. The unit of analysis is the deal/acquirer; identification is cross-sectional using textual intensity.",
  "categories": [
   "Mergers & Acquisitions",
   "Text-as-Data",
   "Intangibles"
  ],
  "datasets": [
   {
    "provider": "Takeover announcement texts (intangibles word list)",
    "product": null,
    "description": "Text of takeover announcements scored with a purpose-built intangibles word list, used to measure 'intangibles talk'.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "M&A; text; intangibles"
   },
   {
    "provider": "Insider trading filings (Form 4)",
    "product": null,
    "description": "Bidder managers' insider trades used to gauge managerial belief in deals.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "insider trades; managers"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "Review of Financial Studies - 2026 - Filipović and Wagner - The Intangibles Song in Takeover Announcements Good Tempo, Hollow Tune",
  "orig_filename": "hhaf116.pdf"
 },
 {
  "journal": "American Economic Review",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Gertler",
   "Huckfeldt",
   "Trigari"
  ],
  "title": "Temporary Layoffs, Loss-of-Recall, and Cyclical Unemployment Dynamics",
  "summary": "This paper revisits the role of temporary layoffs in the business cycle. While recall hiring from temporary layoff can stabilize employment, the authors quantify an important countervailing, destabilizing force—'loss-of-recall'—whereby workers on temporary layoff end up losing their jobs permanently. They develop a quantitative model with endogenous flows across employment, temporary-layoff unemployment, and jobless unemployment. The model matches both pre- and post-pandemic unemployment dynamics, including the contractionary role of loss-of-recall. The paper shows temporary layoffs are not purely stabilizing and that loss-of-recall shapes cyclical unemployment.",
  "logical_flow": "The paper starts from the view that temporary layoffs can stabilize the labor market because laid-off workers are often recalled, dampening cyclical swings. It challenges this by highlighting 'loss-of-recall,' where workers expecting recall instead become permanently unemployed, which is destabilizing and countercyclical. To weigh these opposing forces, the authors build a quantitative model with endogenous worker flows among employment, temporary-layoff unemployment, and jobless unemployment. They discipline the model with data on these flows and show it reproduces unemployment dynamics both before and during the pandemic. The model reveals that loss-of-recall plays a contractionary role that offsets the stabilizing effect of recall hiring. The paper concludes that accounting for loss-of-recall is essential for understanding cyclical unemployment.",
  "research_design": "A quantitative macro-labor model with endogenous flows across employment, temporary-layoff unemployment, and jobless unemployment, calibrated to worker-flow data (e.g., from the Current Population Survey). The model is used to quantify the stabilizing role of recall versus the destabilizing 'loss-of-recall' and to match pre- and post-pandemic unemployment dynamics. The unit of analysis is the aggregate labor market; identification comes from matching flow moments.",
  "categories": [
   "Macroeconomics",
   "Labor Markets",
   "Business Cycles"
  ],
  "datasets": [
   {
    "provider": "Current Population Survey (CPS)",
    "product": null,
    "description": "Worker-flow data on temporary layoffs, recall, and unemployment transitions used to calibrate and test the model.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "CPS; layoffs; unemployment flows"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": null,
  "std_name": "American Economic Review - 2026 - Gertler et al. - Temporary Layoffs, Loss-of-Recall, and Cyclical Unemployment Dynamics",
  "orig_filename": "gertler-et-al-2026-temporary-layoffs-loss-of-recall-and-cyclical-unemployment-dynamics.pdf"
 },
 {
  "journal": "American Economic Review",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Hankins",
   "Momeni",
   "Sovich"
  ],
  "title": "Consumer Credit and the Incidence of Tariffs: Evidence from the Auto Industry",
  "summary": "This paper examines how tariffs affecting the auto industry pass through to consumers via the credit market. Using tariffs as a shock to vehicle costs, the authors study effects on auto loans and purchases, showing how much of the tariff burden falls on consumers through financing. The setting isolates the incidence of trade policy at the point of consumer borrowing. The results reveal that consumer credit is an important channel through which tariff costs reach households. The paper connects trade policy to household finance in a concrete, high-volume market.",
  "logical_flow": "The paper asks who ultimately bears the cost of tariffs, focusing on a channel usually overlooked: consumer credit. It uses tariffs affecting the auto industry as a shock that raises the cost of vehicles, and studies how this propagates to auto loans and purchases. Because cars are typically financed, the credit market is where much of the tariff incidence is realized for households. The authors trace how loan terms, borrowing, and purchases respond to the tariff-driven cost increase. Finding meaningful pass-through to consumers through financing, they quantify the household incidence of the tariffs. The paper concludes that consumer credit is a key conduit through which trade-policy costs reach households, enriching the standard analysis of tariff incidence.",
  "research_design": "A quasi-experimental design using tariffs affecting the auto industry as a shock to vehicle costs, tracing effects on auto loans, borrowing, and purchases to measure the consumer incidence of tariffs through the credit market. The unit of analysis is the loan/consumer; identification comes from the tariff event and cross-sectional exposure.",
  "categories": [
   "Trade Policy",
   "Household & Consumer Credit",
   "Industrial Organization"
  ],
  "datasets": [
   {
    "provider": "Auto loan / consumer credit data",
    "product": null,
    "description": "Loan-level auto financing and purchase data used to measure the pass-through of tariffs to consumers.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "auto loans; consumer credit; incidence"
   },
   {
    "provider": "Vehicle tariff exposure data",
    "product": null,
    "description": "Data on which vehicles/components were affected by tariffs, used to construct exposure to the trade-policy shock.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "tariffs; autos; exposure"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Auto-industry tariffs",
   "type": "Trade policy shock",
   "what": "The imposition of tariffs affecting the auto industry, which raised vehicle costs, used as a shock to identify the incidence of tariffs on consumer credit and auto purchases."
  },
  "missing_notes": null,
  "std_name": "American Economic Review - 2026 - Hankins et al. - Consumer Credit and the Incidence of Tariffs Evidence from the Auto Industry",
  "orig_filename": "hankins-et-al-2026-consumer-credit-and-the-incidence-of-tariffs-evidence-from-the-auto-industry.pdf"
 },
 {
  "journal": "American Economic Review: Insights",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Meyer-ter-Vehn",
   "Board"
  ],
  "title": "Breaking Bad News",
  "summary": "This paper studies how information disclosure shapes social learning about a potentially harmful product. Increased transparency helps early agents avoid harm, but this may undermine learning by later agents who no longer observe harmful outcomes. Despite this conflict of interest, the authors show that full transparency is uniquely optimal for all agents when they learn only by observing neighbors' harm. They then investigate whether full transparency about harm continues to benefit everyone when agents also learn from additional, imperfect signals. The paper clarifies when disclosing bad news helps versus hinders collective learning.",
  "logical_flow": "The paper considers a setting where agents adopt a product that may be harmful and learn about its danger from others' experiences, so disclosure of harm affects social learning. It identifies a tension: making harm transparent protects early agents but can slow later agents' learning, because avoided harm generates less observable information. The authors ask whether, given this conflict, more transparency is good for everyone. When agents learn only by observing neighbors' harm, they prove full transparency is uniquely optimal for all agents, resolving the tension in favor of disclosure. They then extend the analysis to allow additional imperfect signals and examine whether full transparency about harm still benefits all agents. The paper thus delineates the conditions under which breaking bad news aids collective welfare.",
  "research_design": "A theoretical model of social learning about a potentially harmful product, analyzing how information disclosure (transparency about harm) affects early versus later agents. The analysis derives conditions—learning only from neighbors' harm versus also from additional imperfect signals—under which full transparency is optimal for all agents. There is no external dataset.",
  "categories": [
   "Information Economics",
   "Social Learning",
   "Financial Theory"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": null,
  "missing_notes": "Theoretical model; no datasets.",
  "std_name": "American Economic Review Insights - 2026 - Meyer-ter-Vehn and Board - Breaking Bad News",
  "orig_filename": "meyer-ter-vehn-board-2026-breaking-bad-news.pdf"
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
