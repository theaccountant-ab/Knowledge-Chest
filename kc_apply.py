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
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Ghosh",
   "Jacob",
   "Kang",
   "Zhang"
  ],
  "title": "Consumption tax and corporate product mix decisions",
  "summary": "The paper examines how consumption taxes, by taxing input goods, shape firms' product mix decisions. It studies India's staggered transition from a sales tax—under which intermediate inputs could be taxed repeatedly along the supply chain—to a value-added tax (VAT) that offers credits for taxes paid on inputs. Using detailed product-level data for Indian manufacturing firms and a stacked difference-in-differences design, the authors find that firms respond to VAT adoption by narrowing their product scope and reducing vertical integration, shifting away from tax-induced internal input production toward specialization in their most profitable outputs. This vertical disintegration is associated with lower manufacturing costs, higher profitability, improved investment efficiency, and greater firm value. The paper shows that consumption-tax design affects the allocation of resources within firms and that lowering tax burdens on inputs enhances operational efficiency.",
  "logical_flow": "The paper begins from the observation that consumption taxes are the largest source of government revenue worldwide yet little is known about how they reshape corporate behavior and resource allocation. It focuses on a specific distortion of turnover-style sales taxes: because inputs are taxed each time they change hands, firms have a tax incentive to vertically integrate and produce inputs internally to avoid tax cascading. A move to a VAT, which credits taxes paid on inputs, removes this cascading and therefore removes the tax motive for vertical integration, leading the authors to predict that firms will disintegrate and narrow their product scope after VAT adoption. They exploit India's staggered transition from sales tax to VAT across states as the source of variation, using firm-level product data to measure how many distinct products each firm makes. The logic is that once inputs are no longer double-taxed, firms can profitably buy inputs from specialized suppliers rather than make them, concentrating on the outputs where they are most efficient. This reallocation should reduce production costs and raise profitability, investment efficiency, and value, because resources are no longer tied up in tax-motivated in-house input production. The authors trace these predicted effects through cost behavior, vertical integration measures, and firm performance. The argument connects a feature of consumption-tax design directly to the internal boundaries of the firm, concluding that reducing input tax burdens improves within-firm resource allocation.",
  "research_design": "A stacked difference-in-differences design exploiting India's staggered, state-level transition from a cascading sales tax to a value-added tax (VAT) that grants credits for taxes paid on inputs. The unit of analysis is the firm-year, using detailed product-level data on Indian listed manufacturing firms from the Prowess database (1995-2016), with firms assigned to states by registered office address. The main comparison contrasts changes in product scope, vertical integration, and performance for firms in states that adopted VAT earlier versus later, stacking cohorts to avoid contamination from already-treated units. Outcomes include the number of unique products, measures of vertical integration, manufacturing costs, profitability, investment efficiency, and firm value.",
  "categories": [
   "Corporate Taxation",
   "Accounting",
   "Industrial Organization"
  ],
  "datasets": [
   {
    "provider": "Centre for Monitoring Indian Economy",
    "product": "Prowess database",
    "description": "Firm-level identity, financial-statement, stock-price, and detailed product-level data for listed and unlisted Indian companies, covering 2971 product types across 139 industries; used at the firm-year level to measure product mix and vertical integration for manufacturing firms over 1995-2016.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "India; product-level; manufacturing; vertical integration"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "India's staggered sales-tax-to-VAT transition",
   "type": "tax policy reform (staggered difference-in-differences)",
   "what": "India's phased, state-by-state replacement of a cascading sales tax with a value-added tax that credits input taxes, which removed the tax incentive for vertical integration at different times across states, providing plausibly exogenous variation in firms' input tax treatment."
  },
  "missing_notes": null,
  "orig_filename": "1-s2.0-S0165410126000418-main.pdf",
  "std_name": "Journal of Accounting and Economics - 2026 - Ghosh et al. - Consumption tax and corporate product mix decisions"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Li"
  ],
  "title": "Ideology-driven social media opinions and capital markets: Evidence from polarizing boycotts",
  "summary": "The paper studies whether ideology-driven opinions on social media affect how stock markets respond to polarizing firm news, using consumer boycotts as the empirical setting. On average, polarizing boycotts are associated with a roughly 1 percent equity-value decline over the seven trading days after they gain social media traction, and about 2.3 percent over 60 days. The immediate price reaction is more negative when the social media conversation is dominated by users ideologically aligned with the boycotters, especially when their opinions are visible and financially relevant, and a mediation analysis suggests the effect operates through overall social media sentiment toward the firm. Trading volume and return volatility after boycotts rise with the ideological diversity of participating users. The paper concludes that ideology-driven social media opinions can shape investor responses to polarizing corporate news.",
  "logical_flow": "The paper starts from the fact that investors increasingly gather information from social media even though those platforms are saturated with ideological opinion, and asks whether such ideology-driven opinion influences market reactions to firm news. It defines polarizing news as news that generates disagreement among people with different ideological beliefs, and argues boycotts over firms' stances on contested issues are a clean empirical proxy for such news. The author faces three measurement problems—identifying polarizing firm news, inferring social media users' ideologies, and determining those users' views of firms—and addresses them by combining trending-topic boycott data, an ideology-classification of users, and text analysis of posts about boycotted firms. The conceptual mechanism is that when users aligned with the boycotters dominate the conversation, the visible sentiment turns more negative and financially relevant, pushing prices down more sharply. Because different investors weight these signals differently, greater ideological diversity among participants should raise disagreement and hence trading volume and volatility rather than just the price level. A mediation analysis is used to show the price effect runs through aggregate social media sentiment rather than the boycott event alone. The argument links political ideology, the composition of online opinion, and asset-price responses. It concludes that the ideological makeup of social media commentary is itself a determinant of how markets price polarizing corporate events.",
  "research_design": "An empirical event-study-style design at the boycott (firm-event) level, measuring cumulative abnormal returns, trading volume, and volatility in windows around when boycotts gain social media traction. Boycotts are identified from archived daily X (Twitter) trending topics and hand-classified by ideology; social media posts referencing boycotted firms' cashtags are collected via the Twitter Academic API and scored for sentiment and user ideology. The main analyses relate abnormal returns to the ideological alignment and diversity of participating users, and a mediation analysis tests whether the effect operates through overall social media sentiment. Identifying variation comes from cross-boycott differences in the ideological composition and visibility of the online conversation.",
  "categories": [
   "Capital Markets",
   "Behavioral Finance",
   "Political Economy"
  ],
  "datasets": [
   {
    "provider": "TrendCalendar.com",
    "product": "archived daily X (Twitter) trending topics",
    "description": "Archive of the daily top-50 trending topics on X (Twitter), used to identify social media boycotts of US public companies from April 2016 through October 2022.",
    "access_type": "Public",
    "delivery": "Web",
    "topic_tags": "social media; trending topics; boycotts"
   },
   {
    "provider": "Twitter (X)",
    "product": "Twitter Academic API",
    "description": "Post-level social media data on tweets containing boycotted firms' cashtags in the days after boycotts trended, yielding about 135,943 posts by 59,934 unique users, used to measure social media sentiment and user ideology.",
    "access_type": "Restricted",
    "delivery": "API",
    "topic_tags": "tweets; cashtags; sentiment; user ideology"
   },
   {
    "provider": "USC Center for Public Relations",
    "product": "USC Polarization Index",
    "description": "Index classifying issues by how politically polarizing they are, used to retain boycotts tied to polarizing corporate stances (e.g., abortion, climate change, LGBTQ rights, gun policy).",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "political polarization; issue classification"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Polarizing consumer boycotts of public firms",
   "type": "event study (social media boycott events)",
   "what": "Discrete boycotts triggered by firms' perceived stances on ideologically polarizing issues, whose timing (when they trend on social media) provides news events used to measure market reactions; treated as plausibly exogenous shocks to the firm's information environment."
  },
  "missing_notes": null,
  "orig_filename": "1-s2.0-S0165410126000522-main.pdf",
  "std_name": "Journal of Accounting and Economics - 2026 - Li - Ideology-driven social media opinions and capital markets Evidence from polarizing boycotts"
 },
 {
  "journal": "Journal of Accounting and Economics",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Chen",
   "Hutchens",
   "Xia"
  ],
  "title": "The effects of tax clienteles on disclosure: Evidence from the municipal bond market",
  "summary": "The paper offers a tax-clientele explanation for the municipal bond market's well-known lack of disclosure, contrasting it with the standard weak-regulation explanation. Because most municipal bonds are tax-exempt, they primarily attract high-tax retail investors who have little incentive or ability to demand and monitor continuing disclosures, whereas taxable municipal bonds draw a broader base including institutional investors who can demand disclosure. The authors find that issuers provide more continuing disclosures in years when they have taxable bonds outstanding, increase disclosure after issuing their first taxable bond, and reduce it after calling their last taxable bond. Exploiting a tax law change that raised taxable bond issuance for plausibly exogenous reasons, they confirm that disclosures rise when taxable issuance increases. The results imply that the investor tax clientele, not just regulation, shapes disclosure practices in the municipal market.",
  "logical_flow": "The paper motivates itself with the puzzle that the municipal bond market—over 50,000 issuers and more than $4 trillion outstanding—is persistently criticized for weak disclosure despite SEC Rule 15c2-12, with prior work attributing non-compliance mainly to weak regulatory oversight. The authors propose an alternative, demand-side explanation rooted in the tax status of the bonds and hence the type of investor who holds them. Tax-exempt municipal bonds are attractive chiefly to high-tax individual investors, a clientele that is dispersed and ill-equipped to monitor issuers, so issuers face little pressure to disclose. Taxable municipal bonds, by contrast, appeal to institutions and tax-exempt investors who have both the incentive and the sophistication to demand information, so an issuer with taxable bonds outstanding should disclose more. This yields within-issuer predictions: disclosure should rise when taxable bonds are present, increase after a first taxable issuance, and fall after the last taxable bond is called. To rule out that unobserved issuer characteristics drive both taxable issuance and disclosure, the authors exploit a tax law change that pushed issuers toward taxable bonds for reasons plausibly unrelated to their disclosure preferences. Finding that disclosure responds to this exogenous variation supports a causal tax-clientele channel. The paper concludes that the composition of an issuer's investor base, driven by bond tax status, is a first-order determinant of municipal disclosure.",
  "research_design": "A panel design at the issuer-year level relating municipal issuers' continuing disclosure activity to the presence and share of taxable bonds outstanding, using within-issuer variation around first taxable issuances and last taxable calls. Bond characteristics come from the Mergent Municipal Bond Securities Database and continuing-disclosure filings from the MSRB's EMMA system, aggregated to the issuer-year. For causal identification, the authors use a difference-in-differences-style analysis exploiting the Tax Cuts and Jobs Act's elimination of tax-exempt advance refunding, which forced issuers to use taxable advance refunding bonds, generating plausibly exogenous increases in taxable issuance. Outcomes are measures of the frequency and completeness of continuing disclosures.",
  "categories": [
   "Financial Accounting",
   "Public Economics",
   "Fixed Income"
  ],
  "datasets": [
   {
    "provider": "Mergent",
    "product": "Municipal Bond Securities Database",
    "description": "Security-level data on US municipal bonds—issuer identity, tax status, amounts outstanding, calls, and other characteristics—aggregated to the issuer-year to construct measures of taxable bonds and controls for issuers in the 50 states and DC, fiscal years 2009-2022.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "municipal bonds; issuer-year; tax status"
   },
   {
    "provider": "Municipal Securities Rulemaking Board (MSRB)",
    "product": "EMMA continuing disclosures",
    "description": "Continuing disclosure filings (financial filings and event notices) submitted by municipal issuers to the MSRB's Electronic Municipal Market Access (EMMA) repository, matched to issuers by CUSIP and filing date to measure post-issuance disclosure.",
    "access_type": "Public",
    "delivery": "Web",
    "topic_tags": "continuing disclosure; EMMA; municipal issuers"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "TCJA elimination of tax-exempt advance refunding",
   "type": "tax policy reform (difference-in-differences)",
   "what": "The 2017 Tax Cuts and Jobs Act ended tax-exempt advance refunding, forcing municipalities that wanted to advance-refund debt to issue taxable bonds, which raised taxable issuance for reasons plausibly unrelated to issuers' disclosure preferences."
  },
  "missing_notes": null,
  "orig_filename": "1-s2.0-S0165410126000534-main.pdf",
  "std_name": "Journal of Accounting and Economics - 2026 - Chen et al. - The effects of tax clienteles on disclosure Evidence from the municipal bond market"
 },
 {
  "journal": "Journal of Political Economy",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Demirer",
   "Karaduman"
  ],
  "title": "Do Mergers and Acquisitions Improve Efficiency? Evidence from Power Plants",
  "summary": "The paper studies whether mergers and acquisitions raise productive efficiency, using US power plants where hourly output and input data allow precise productivity measurement. Analyzing thousands of ownership changes, the authors find a roughly 2 percent average increase in efficiency at acquired plants, emerging about five months after acquisition. Efficiency gains reach about 5 percent when direct plant ownership changes, but there is no significant change when only the parent (ultimate) owner changes. About three-quarters of the gain comes from higher productive efficiency and the remainder from dynamic efficiency through reallocation of production across units. The evidence suggests that high-productivity firms buy underperforming assets from low-productivity firms and raise them toward their own productivity through operational improvements.",
  "logical_flow": "The paper is motivated by the central antitrust trade-off between the market-power harms and the potential efficiency gains of mergers, noting that efficiency effects are hard to measure because output and inputs are usually observed only coarsely. Power generation offers an unusually clean laboratory: plants produce a homogeneous good (electricity) with hourly data on fuel inputs and output, so physical productivity can be measured directly and tracked over time. The authors assemble thousands of ownership changes and compare plants' productivity before and after acquisition, using the timing of ownership change as the key source of variation. They distinguish changes in direct ownership of the plant from changes only in the ultimate parent, reasoning that operational control—and thus scope for efficiency improvement—accompanies direct ownership rather than distant parent restructuring. Finding gains only for direct ownership changes supports the interpretation that acquirers actively improve operations rather than merely relabeling assets. They then decompose the improvement into productive efficiency (getting more output from given inputs) and dynamic efficiency (reallocating production toward more efficient units), attributing about three-quarters to the former. The pattern that acquirers are high-productivity firms buying low-productivity assets supports a 'buy low, improve' mechanism of asset reallocation. The paper concludes that M&A can generate real, measurable efficiency gains through operational improvements at acquired assets.",
  "research_design": "An event-study / difference-in-differences design at the plant level exploiting the timing of thousands of ownership changes at US power plants, comparing acquired plants' productivity before and after acquisition against non-acquired plants. Productivity is measured from high-frequency (hourly) plant operating data on fuel inputs and electricity output, allowing direct estimation of physical/productive efficiency. The design separates direct plant-ownership changes from ultimate-parent-only changes to isolate operational control, and decomposes gains into productive efficiency and dynamic (reallocation) efficiency. Identifying variation comes from the staggered timing and type of ownership changes.",
  "categories": [
   "Industrial Organization",
   "Energy Economics",
   "Corporate Finance"
  ],
  "datasets": [
   {
    "provider": "US Environmental Protection Agency",
    "product": "Continuous Emissions Monitoring System (CEMS)",
    "description": "Hourly, unit-level data on US fossil-fuel power plants' fuel inputs, gross generation, and emissions, used to measure physical plant productivity/efficiency at high frequency around ownership changes.",
    "access_type": "Public",
    "delivery": "Bulk",
    "topic_tags": "power plants; hourly; emissions; productivity"
   },
   {
    "provider": "US Energy Information Administration",
    "product": "power plant operations and generation data",
    "description": "Plant- and generator-level data on US power plants (capacity, fuel, generation, and characteristics) used together with hourly operating data to construct productivity measures and plant panels.",
    "access_type": "Public",
    "delivery": "Bulk",
    "topic_tags": "power plants; generation; capacity"
   },
   {
    "provider": "S&P Global (SNL)",
    "product": "Velocity Suite",
    "description": "Data on US power-plant ownership and thousands of ownership changes over time, distinguishing direct plant ownership from ultimate parent ownership, used to date and classify acquisitions.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "ownership changes; power plants; M&A"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Power-plant ownership changes (acquisitions)",
   "type": "event study / natural experiment (M&A timing)",
   "what": "Thousands of staggered changes in ownership of US power plants, whose timing and type (direct plant vs. parent-only) provide variation used to identify post-acquisition changes in plant productivity relative to non-acquired plants."
  },
  "missing_notes": null,
  "orig_filename": "740222.pdf",
  "std_name": "Journal of Political Economy - 2026 - Demirer and Karaduman - Do Mergers and Acquisitions Improve Efficiency Evidence from Power Plants"
 },
 {
  "journal": "The Review of Financial Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Trebbi",
   "Zhang",
   "Simkovic"
  ],
  "title": "The Cost of Regulatory Compliance in the United States",
  "summary": "The paper measures the cost of regulatory compliance across US firms and asks whether it falls evenly on small and large businesses. Using comprehensive establishment-occupation microdata combined with occupation-level task information, the authors quantify a firm's compliance cost as the share of its wage bill devoted to regulatory-compliance tasks, a measure they call RegIndex. They document an inverted-U relationship between RegIndex and firm size: mid-sized firms (around 500 employees) spend about 47 percent more on compliance, as a share of the wage bill, than the smallest firms and about 18 percent more than the largest firms. They develop a shift-share methodology to separate the roles of regulatory requirements and enforcement in driving compliance costs. The findings speak to how regulatory costs affect business dynamism and the relative burden on mid-sized firms.",
  "logical_flow": "The paper begins from the debate over declining US business dynamism and the concern that regulatory compliance costs may fall unevenly across firm sizes, potentially disadvantaging particular firms. Because compliance costs are not directly reported, the authors need a way to measure them from observable data, and they build one from the labor devoted to compliance tasks. Their key idea is to combine detailed occupation-by-establishment employment data with task descriptions that identify which occupations perform regulatory-compliance work, so a firm's compliance intensity is the share of its wage bill spent on such tasks—the RegIndex. Aggregating to the firm level, they uncover an inverted-U in RegIndex against size, meaning mid-sized firms bear the heaviest relative compliance burden while both the smallest and largest firms bear less. They interpret this shape through economies of scale in compliance and the differential incidence of requirements and enforcement across firm sizes. To separate these forces, they design a shift-share (Bartik-style) methodology that decomposes compliance costs into variation coming from regulatory requirements versus enforcement intensity. This lets them attribute the size gradient to specific regulatory channels rather than to unobserved firm heterogeneity. The paper concludes that regulatory compliance imposes a quantitatively large and size-dependent burden, with mid-sized firms most affected.",
  "research_design": "A measurement-and-decomposition study that constructs a firm-level compliance-cost index (RegIndex) from establishment-occupation microdata and occupation task information, defined as the share of the wage bill spent on regulatory-compliance tasks. The unit of analysis is the firm (and establishment), built up from occupation-by-establishment employment and wages. The central empirical object is the relationship between RegIndex and firm size, and a shift-share methodology is developed to disentangle regulatory requirements from enforcement as drivers of compliance costs. Identification of the requirement-versus-enforcement components relies on the shift-share structure rather than a single natural experiment.",
  "categories": [
   "Regulation",
   "Labor Economics",
   "Industrial Organization"
  ],
  "datasets": [
   {
    "provider": "US Bureau of Labor Statistics",
    "product": "Occupational Employment and Wage Statistics (OEWS) microdata",
    "description": "Establishment-occupation-level microdata (2002-2014) tracking employment and wage rates for over 800 occupations across roughly 1.2 million establishments, with sampling weights, NAICS 6-digit industry, county, government-ownership indicators, and parent-firm EIN; used to build firm- and establishment-level compliance-cost measures.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "establishment-occupation; employment; wages; United States"
   },
   {
    "provider": "US Department of Labor",
    "product": "O*NET occupational task information",
    "description": "Occupation-level task and work-activity descriptions used to identify which occupations perform regulatory-compliance tasks, forming the basis of the RegIndex compliance-cost measure.",
    "access_type": "Public",
    "delivery": "Web",
    "topic_tags": "occupations; tasks; compliance"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": "Primarily a measurement paper: it constructs the RegIndex and uses a shift-share design to decompose requirements versus enforcement rather than exploiting a single quasi-exogenous shock.",
  "orig_filename": "hhag046.pdf",
  "std_name": "The Review of Financial Studies - 2026 - Trebbi et al. - The Cost of Regulatory Compliance in the United States"
 },
 {
  "journal": "The Review of Financial Studies",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Bauer",
   "Huber",
   "Offner",
   "Wilms"
  ],
  "title": "Corporate Green Pledges",
  "summary": "The paper builds a novel dataset of time-stamped corporate decarbonization commitments—'green pledges'—for US public firms by classifying news articles with large language models and human validation. Firms that announce green pledges tend to be larger and 'browner' (more carbon-intensive) than other firms, both within and across industries. Announcing a green pledge significantly raises a firm's stock price, consistent with a reduction in its carbon premium, and predicts sizable subsequent declines in carbon emissions and emission intensity, with the strongest effects among firms in brown industries. The authors interpret these results as evidence that green pledges are credible, convey new information to investors, and create meaningful financial incentives to decarbonize. The paper contributes both a new measurement of corporate climate commitments and evidence on their asset-pricing and real effects.",
  "logical_flow": "The paper is motivated by the difficulty of knowing whether corporate climate commitments are credible and material, given that pledges are announced in unstructured news rather than standardized filings. The authors' first move is a measurement innovation: they use large language models, validated by humans, to read news articles and extract time-stamped green pledges for US public firms, creating a panel of who pledged and when. Describing the sample, they find that pledging firms are larger and more carbon-intensive, suggesting pledges come disproportionately from firms with the most to abate. They then treat pledge announcements as information events and examine stock-price reactions, reasoning that if pledges are credible and reduce expected future carbon risk, prices should rise as the carbon premium falls. Observing positive announcement returns supports the view that investors update toward lower carbon risk. To test whether pledges are more than cheap talk, they examine subsequent realized emissions and find meaningful declines, especially for brown firms, indicating real decarbonization follows. The rising stock price also implies a financial reward that strengthens firms' incentives to follow through. Linking measurement, asset prices, and realized emissions, the argument builds from 'can we observe pledges' to 'do markets believe them' to 'do firms act on them.' The paper concludes that green pledges are credible, informative, and consequential for both valuation and emissions.",
  "research_design": "An empirical study combining a novel text-based dataset with an event-study analysis of announcement returns. The authors construct the green-pledge dataset by classifying Dow Jones news articles with large language models plus human validation to produce time-stamped, firm-level pledge events for US public firms. They then estimate stock-price reactions around pledge announcements (interpreting positive returns as a reduced carbon premium) and use panel regressions to relate pledges to subsequent carbon emissions and emission intensity, with heterogeneity by industry 'brownness.' Identifying variation comes from the timing of pledge announcements across firms; realized-emissions data provide an ex post validation of credibility.",
  "categories": [
   "Climate Finance",
   "Asset Pricing",
   "Environmental Economics"
  ],
  "datasets": [
   {
    "provider": "Bauer, Huber, Offner, and Wilms",
    "product": "corporate green-pledges dataset",
    "description": "A novel hand-validated, LLM-assisted dataset of time-stamped corporate decarbonization commitments ('green pledges') for US public firms, constructed by classifying news articles; records the identity and announcement date of each pledging firm.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "green pledges; decarbonization; text classification; US firms"
   },
   {
    "provider": "Dow Jones",
    "product": "news article archive",
    "description": "Archive of business news articles classified with large language models to detect and time-stamp corporate green pledges for US public firms.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "news; text; corporate announcements"
   },
   {
    "provider": "S&P Global",
    "product": "Trucost carbon emissions data",
    "description": "Firm-level carbon-emissions and emission-intensity data used to test whether green pledges predict subsequent declines in emissions.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "carbon emissions; emission intensity; firm-level"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Corporate green-pledge announcements",
   "type": "event study (information events)",
   "what": "Time-stamped announcements of corporate decarbonization commitments, whose timing is used as firm-level information events to measure stock-price reactions and subsequent emission changes."
  },
  "missing_notes": null,
  "orig_filename": "hhag063.pdf",
  "std_name": "The Review of Financial Studies - 2026 - Bauer et al. - Corporate Green Pledges"
 },
 {
  "journal": "American Economic Review",
  "is_working_paper": false,
  "year": 2026,
  "authors": [
   "Kennedy",
   "Dobridge",
   "Landefeld",
   "Mortenson"
  ],
  "title": "Corporate Tax Cuts, Firm Growth, and Workers' Earnings",
  "summary": "The paper studies the effects of the 2017 Tax Cuts and Jobs Act (TCJA)—the largest corporate income tax cut in US history—on firms and their workers. Using employer-employee matched federal tax records, the authors run event studies comparing similarly sized firms in the same industry that faced divergent tax changes because of their preexisting legal status (C versus S corporations). They find that corporate tax cuts cause increases in firms' investment, sales, profits, employment, and payrolls. Earnings gains are concentrated among highly paid workers, with 87 percent of the short-run private income gains flowing to the top 10 percent of the income distribution. The paper provides causal, administrative-data evidence on how corporate tax cuts are distributed between firms and different groups of workers.",
  "logical_flow": "The paper is motivated by long-standing debate over who benefits from corporate tax cuts—shareholders, firms, or workers—and by the 2017 TCJA, which sharply cut the corporate rate and changed investment incentives. The empirical challenge is that firm outcomes are driven by many forces, so the authors need variation in tax treatment that is not confounded with a firm's own prospects. Their key insight is that TCJA changed taxes very differently for C corporations and pass-through S corporations, so otherwise similar firms experienced divergent tax shocks purely because of their preexisting legal status. Comparing similarly sized firms in the same industry that differ in this status isolates the causal effect of the corporate tax change. Using matched employer-employee tax records lets them trace effects from the firm (investment, sales, profits, employment, payroll) down to individual workers' earnings. They find that tax cuts raise firm activity and total payrolls, but that the earnings gains accrue mainly to highly paid workers rather than being broadly shared. Quantifying the distribution, they show most short-run private income gains go to the top of the distribution. The argument moves from a policy shock, to a design that exploits legal-status-based tax differences, to firm responses, and finally to the incidence of those responses across the earnings distribution. The paper concludes that corporate tax cuts stimulate firm growth but distribute the gains regressively in the short run.",
  "research_design": "An event-study / difference-in-differences design using US employer-employee matched federal tax records (2013-2019). Identification exploits that the 2017 TCJA changed tax treatment differentially for C corporations versus S corporations, comparing similarly sized firms in the same industry that faced divergent tax changes solely due to preexisting legal status. Firm-level outcomes (investment, sales, profits, employment, payroll) come from IRS corporate Statistics of Income files, and worker outcomes come from matched individual earnings records, allowing analysis of incidence across the earnings distribution. The comparison is between C- and S-corp firms before and after TCJA.",
  "categories": [
   "Public Economics",
   "Corporate Taxation",
   "Labor Economics"
  ],
  "datasets": [
   {
    "provider": "US Internal Revenue Service",
    "product": "corporate Statistics of Income (SOI) files",
    "description": "Stratified random samples of US corporate tax returns for C corps (Form 1120) and S corps (Form 1120-S), providing firm-level investment, sales, costs, profits, taxes paid, incorporation year, and industry; the analysis sample covers 49,235 firms and 231,360 firm-year observations, tax years 2013-2019.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "corporate tax returns; firm-year; SOI; United States"
   },
   {
    "provider": "US Internal Revenue Service",
    "product": "employer-employee matched tax records",
    "description": "Administrative individual and firm tax records linked into an employer-employee panel, used to measure workers' earnings and the incidence of corporate tax changes across the earnings distribution.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "matched employer-employee; earnings; administrative tax data"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "2017 Tax Cuts and Jobs Act (C- vs. S-corp differential)",
   "type": "tax policy reform (difference-in-differences by legal status)",
   "what": "The 2017 TCJA cut corporate taxes sharply but changed C-corporation and S-corporation taxation differently, so otherwise similar firms faced divergent tax shocks purely because of their preexisting legal status, providing plausibly exogenous variation in tax treatment."
  },
  "missing_notes": null,
  "orig_filename": "kennedy-et-al-2026-corporate-tax-cuts-firm-growth-and-workers-earnings.pdf",
  "std_name": "American Economic Review - 2026 - Kennedy et al. - Corporate Tax Cuts, Firm Growth, and Workers' Earnings"
 },
 {
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Drechsler",
   "Jung",
   "Peng",
   "Supera",
   "Zhou"
  ],
  "title": "Credit Card Banking",
  "summary": "The paper asks why credit card interest rates are so high—averaging about 22 percent, an 18 percentage-point spread over the short rate that exceeds any other loan or bond—even though nearly half of households borrow on cards. Using regulatory account-level data covering the lifetime cash flows of 550 million monthly accounts (about 90 percent of the US credit card market), the authors decompose the economics of credit card banking. Charge-offs, though high at around 6 percent, explain only part of the spread; rewards and non-interest expenses are more than offset by interchange and other non-interest income, while operating expenses—especially marketing—are very large and are used to build pricing power. After all costs, card lending still earns a 6.8 percent return on assets, more than four times the banking sector's, and, accounting for a roughly 4.3 percent default-risk premium, earns an alpha of about 1.5 percent relative to the aggregate bank sector. The paper explains high card rates through a combination of default risk and marketing-generated pricing power.",
  "logical_flow": "The paper opens with a puzzle: credit card rates carry an 18-point spread over the short rate, larger than on any other loan or bond, yet borrowing on cards is nearly universal, so simple risk-based explanations seem incomplete. To answer why, the authors argue one must account for the full economics of the credit card business rather than interest income alone. They use granular regulatory account-level data to reconstruct the lifetime cash flows of card accounts—interest, fees, interchange, rewards, charge-offs, and operating costs—so every component of profitability can be weighed. They first show that charge-offs, while high, account for only a fraction of the spread, ruling out default losses as the main driver. They then net rewards and non-interest expenses against interchange and other non-interest income, finding these roughly offset, which shifts attention to operating costs. Operating expenses, particularly heavy marketing spending, turn out to be very large, and the authors argue this marketing is not pure cost but an investment that generates pricing power over borrowers. Even after subtracting these costs, card lending earns an outsized return on assets, so the authors ask how much reflects compensation for risk. Using the cross section of accounts they estimate a default-risk premium comparable to high-yield bonds and show that, adjusting for it, card lending still earns a positive alpha. The paper concludes that high card rates reflect both substantial default risk and marketing-driven market power rather than default losses alone.",
  "research_design": "A descriptive and asset-pricing decomposition using comprehensive regulatory account-level credit card data (the Federal Reserve's Y-14M reports) covering roughly 550 million monthly accounts and about 90 percent of the US market. The authors reconstruct lifetime account cash flows and decompose card profitability into interest income, interchange, fees, rewards, charge-offs, and operating (including marketing) expenses to compute returns on assets. They exploit the cross section of accounts to estimate a default-risk premium and compute a risk-adjusted alpha relative to the aggregate bank sector. The analysis is primarily accounting/measurement and cross-sectional estimation rather than a single natural experiment.",
  "categories": [
   "Banking",
   "Household Finance",
   "Financial Intermediation"
  ],
  "datasets": [
   {
    "provider": "Federal Reserve",
    "product": "Y-14M reports (credit card schedule)",
    "description": "Comprehensive supervisory account-level monthly data on US credit card loans collected for stress testing and capital assessment, covering about 550 million monthly accounts and roughly 90 percent of the US credit card market, used to reconstruct lifetime account cash flows and profitability.",
    "access_type": "Restricted",
    "delivery": null,
    "topic_tags": "credit cards; account-level; supervisory; stress testing"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": null,
  "missing_notes": "NBER working paper (No. 35607, August 2026). Primarily a measurement/decomposition and cross-sectional asset-pricing study; it does not exploit a single quasi-exogenous shock.",
  "orig_filename": "w35607.pdf",
  "std_name": "Working Paper - 2026 - Drechsler et al. - Credit Card Banking"
 },
 {
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Prilmeier",
   "Stulz"
  ],
  "title": "The Cov-Lite Liquidity Advantage, Regulatory Pressures, and the Evolution of the Leveraged Loan Market",
  "summary": "The paper examines why leveraged loans increasingly lack maintenance covenants ('cov-lite' loans) and links this shift to post-crisis bank regulation. After the global financial crisis, regulators made it harder for banks to retain leveraged-loan exposure, and the authors conjecture this raised the share of cov-lite issuance because such loans are easier to sell to institutional investors. Consistent with this, the post-GFC rise in cov-lite issuance was larger for banks facing stricter regulation, banks that outright failed a stress test, or banks more vulnerable to the severely adverse stress-test scenario. The authors also show, as theory predicts, that cov-lite loans carry a liquidity advantage that lowers their credit spread, and that this advantage is larger for private firms. The paper connects regulatory pressure on banks to a structural transformation of the leveraged-loan market.",
  "logical_flow": "The paper starts from the dramatic post-GFC transformation of the leveraged-loan market, in which cov-lite loans—loans without maintenance covenants—went from a small minority to the dominant form. The authors note that the typical leveraged loan splits into a bank-held revolver and an institutionally held term-B tranche, so the salability of the term loan to non-bank investors matters for how it is structured. Their central conjecture is that post-crisis regulation, by making it costlier for banks to retain leveraged-loan exposure, pushed banks to originate loans they could more easily distribute, and cov-lite loans are easier to sell because they impose fewer ongoing constraints prized by institutional buyers. This yields a cross-sectional prediction: banks under greater regulatory pressure—those facing stricter rules, failing stress tests, or more exposed to adverse stress scenarios—should shift more toward cov-lite issuance. Confirming this pattern ties the market's evolution to regulation rather than to borrower demand alone. The authors then turn to pricing, arguing that if cov-lite loans are more liquid because they are easier to trade and hold, they should command lower credit spreads, a liquidity advantage. They further reason this advantage is larger for private firms, where covenants and monitoring would otherwise be more valuable, so the tradeoff tilts more toward liquidity. Linking regulation, loan structure, and pricing, the paper explains both the rise of cov-lite loans and their spread differential. It concludes that regulatory pressure on banks reshaped leveraged-loan contracting and pricing through a liquidity channel.",
  "research_design": "An empirical study of the US leveraged-loan market combining cross-sectional and difference-in-differences-style analysis. Loan-level data on issuance and covenant structure (from DealScan) are linked to bank-level regulatory measures—stress-test outcomes (CCAR/DFAST), including outright failures and vulnerability to the severely adverse scenario—to test whether banks under greater regulatory pressure shifted more toward cov-lite issuance after the GFC. Secondary-market loan prices (LSTA) are used to estimate the credit-spread (liquidity) advantage of cov-lite loans, with heterogeneity by borrower public/private status. Identifying variation comes from cross-bank differences in regulatory pressure and the pre/post-GFC timing.",
  "categories": [
   "Credit & Lending",
   "Financial Regulation",
   "Banking"
  ],
  "datasets": [
   {
    "provider": "LSEG (Refinitiv)",
    "product": "LPC DealScan",
    "description": "Loan-level data on syndicated and leveraged loans—issuance, tranche structure, and covenant terms (including cov-lite status)—used to measure the share of cov-lite issuance across banks and over time.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "leveraged loans; covenants; syndicated lending; cov-lite"
   },
   {
    "provider": "Federal Reserve",
    "product": "CCAR/DFAST stress test results",
    "description": "Supervisory stress-test outcomes for large US banks (CCAR/DFAST), including pass/fail results and post-stress capital under the severely adverse scenario, used to measure banks' regulatory pressure and vulnerability.",
    "access_type": "Public",
    "delivery": null,
    "topic_tags": "stress tests; bank regulation; capital; CCAR; DFAST"
   },
   {
    "provider": "Loan Syndications and Trading Association (LSTA)",
    "product": "secondary-market loan prices",
    "description": "Secondary-market bid/ask prices for syndicated leveraged loans, used to estimate credit spreads and the liquidity advantage of cov-lite loans.",
    "access_type": "Proprietary",
    "delivery": null,
    "topic_tags": "loan prices; liquidity; secondary market"
   }
  ],
  "no_nonstandard_datasets": false,
  "shock": {
   "name": "Post-GFC bank regulation and stress tests",
   "type": "regulatory reform (cross-bank exposure / difference-in-differences)",
   "what": "Post-crisis regulatory measures (leveraged-lending guidance and CCAR/DFAST stress testing) that made it costlier for banks to retain leveraged-loan exposure, with cross-bank differences in regulatory pressure providing variation used to identify the shift toward cov-lite issuance."
  },
  "missing_notes": "NBER working paper (No. 35617, August 2026).",
  "orig_filename": "w35617.pdf",
  "std_name": "Working Paper - 2026 - Prilmeier and Stulz - The Cov-Lite Liquidity Advantage, Regulatory Pressures, and the Evolution of the Leveraged Loan Market"
 },
 {
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Chen",
   "Jin"
  ],
  "title": "When Public Signals Backfire: Strategic Disclosure and Information Crowd-Out",
  "summary": "The paper studies voluntary disclosure when receivers are only partially sophisticated ('partially naive') and observe a public signal that is correlated with the sender's private information. The authors develop a theoretical framework in which the public signal changes how receivers interpret a sender's silence, and they derive a closed-form disclosure threshold that nests the classical unraveling result as a special case. A laboratory experiment that varies the correlation between the public signal and the sender's information confirms the predictions: disclosure falls as the public signal becomes more favorable, and receivers' guesses following nondisclosure rise toward the signal. Crucially, the public signal can crowd out disclosure, leaving receivers less informed than they would have been without it, with the informational loss peaking at intermediate correlation where the signal moves beliefs but remains too noisy to substitute for disclosure. The paper shows that more public information can, through strategic disclosure incentives, reduce how much receivers ultimately learn.",
  "logical_flow": "The paper is motivated by settings where buyers learn about quality from two sources: a public signal and a seller's private information that can be disclosed or withheld but not fabricated. Classical unraveling predicts full disclosure because silence is interpreted as bad news, but the authors ask how an informative public signal changes this logic. Their key modeling step is to let receivers be partially naive, so they do not fully back out the sender's incentives, which means the interpretation of silence depends on the public signal rather than adjusting to restore full revelation. When the public signal is favorable, silence looks less damaging, so the sender discloses less; the authors formalize this with a closed-form disclosure threshold that collapses to standard unraveling in the limiting rational case. This generates the counterintuitive possibility that a more informative or favorable public signal reduces disclosure enough to leave receivers worse informed overall—information crowd-out. The framework predicts this loss is largest at intermediate correlation, where the signal is influential enough to shift beliefs but too noisy to replace the withheld private information. To test the mechanism cleanly, the authors run a laboratory experiment that exogenously varies the correlation between the public signal and the sender's information. The experiment confirms that disclosure decreases and post-nondisclosure guesses rise as the signal becomes more favorable, matching the theory. The paper concludes that public information interacts with strategic disclosure so that more of it can reduce what receivers learn.",
  "research_design": "A theory paper paired with a laboratory experiment. The core is a game-theoretic voluntary-disclosure model with a sender holding verifiable private information, partially naive receivers, and a correlated public signal, yielding a closed-form disclosure threshold that nests classical unraveling. The predictions are tested in a controlled laboratory experiment that exogenously varies the correlation between the public signal and the sender's information, measuring disclosure rates and receivers' beliefs upon nondisclosure. Identification comes from the experimental manipulation of the signal's correlation/favorableness rather than from field data.",
  "categories": [
   "Information Economics",
   "Disclosure",
   "Experimental Economics"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": null,
  "missing_notes": "NBER working paper (No. 35625, August 2026). Theory plus a laboratory experiment; the only data are the authors' own experimental observations, so there is no external distinctive dataset and no field quasi-experiment.",
  "orig_filename": "w35625.pdf",
  "std_name": "Working Paper - 2026 - Chen and Jin - When Public Signals Backfire Strategic Disclosure and Information Crowd-Out"
 },
 {
  "journal": null,
  "is_working_paper": true,
  "year": 2026,
  "authors": [
   "Ferrante",
   "Prestipino",
   "Raffo",
   "Waugh"
  ],
  "title": "Tariffs, Investment, and the Missing Trade Collapse",
  "summary": "The paper asks why US imports rose in 2025 even though tariff rates climbed to levels not seen since the Great Depression—the 'missing trade collapse.' The authors build an open-economy New Keynesian model with tariff heterogeneity across goods, inventories, and shocks to investment that capture an AI-driven investment boom. The model matches the untargeted paths of imports, output, and inflation, and is used to decompose the separate effects of tariffs and the investment boom: absent the boom, imports would have fallen about 10 percent and activity contracted about 0.7 percent. The effects of tariffs depend on which goods are taxed—tariffs on consumption and intermediate goods act like supply shocks, while tariffs on capital goods act like demand shocks. Because the 2025 tariff increases were concentrated on consumption goods and largely spared capital goods, they limited the damage to output while amplifying the inflationary impulse.",
  "logical_flow": "The paper is motivated by a puzzle in the 2025 US data: tariffs rose to near-Depression-era highs, which standard intuition says should collapse imports, yet imports actually increased. The authors argue that resolving this requires a model rich enough to separate the contractionary force of tariffs from an offsetting boom in investment associated with the AI build-out. They construct an open-economy New Keynesian model whose key ingredients are tariff heterogeneity across consumption, intermediate, and capital goods, inventories that let importers front-run or smooth tariff changes, and investment shocks that proxy the AI-driven demand surge. The central conceptual point is that tariffs are not uniform in their macro effects: taxing consumption or intermediate goods raises costs and behaves like a supply shock, whereas taxing capital goods depresses investment demand and behaves like a demand shock. Because the 2025 tariffs fell mostly on consumption goods and spared capital goods, their contractionary bite on output was limited while their inflationary impulse was amplified. Meanwhile, the AI-linked investment boom raised demand for imported capital and intermediates, buoying imports and masking the trade decline tariffs would otherwise have caused. The model is disciplined by matching untargeted paths of imports, output, and inflation, lending credibility to its decomposition. Counterfactually, without the investment boom imports would have fallen sharply and activity contracted. The paper concludes that the composition of tariffs and a coincident investment boom together explain the missing trade collapse.",
  "research_design": "A quantitative theory / structural macro-modeling paper. The authors develop a two-country open-economy New Keynesian model with heterogeneous tariffs across consumption, intermediate, and capital goods, imported-input inventories, and investment shocks representing an AI-driven boom, building on standard small-open-economy frameworks. The model is calibrated/estimated and validated by its ability to match untargeted empirical paths of imports, output, and inflation, then used for counterfactual decompositions isolating the effects of tariffs versus the investment boom and of tariffs by good type. The analysis is model-based rather than reduced-form estimation on a distinctive dataset.",
  "categories": [
   "International Trade",
   "Macroeconomics",
   "International Finance"
  ],
  "datasets": [],
  "no_nonstandard_datasets": true,
  "shock": {
   "name": "2025 US tariff increases",
   "type": "trade policy shock (structural model input)",
   "what": "The 2025 rise in US tariff rates to levels not seen since the Great Depression, concentrated on consumption goods and largely sparing capital goods, used as the policy shock whose macro effects the model decomposes."
  },
  "missing_notes": "NBER working paper (No. 35630, August 2026). Structural open-economy model calibrated to aggregate data; no distinctive proprietary dataset.",
  "orig_filename": "w35630.pdf",
  "std_name": "Working Paper - 2026 - Ferrante et al. - Tariffs, Investment, and the Missing Trade Collapse"
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
