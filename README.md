# NST DVA Capstone 2 - Project Repository

> **Newton School of Technology | Data Visualization & Analytics**
> A 2-week industry simulation capstone using Python, GitHub, and Tableau to convert raw data into actionable business intelligence.

---

## Before You Start

1. Rename the repository using the format `SectionName_TeamID_ProjectName`.
2. Fill in the project details and team table below.
3. Add the raw dataset to `data/raw/`.
4. Complete the notebooks in order from `01` to `05`.
5. Publish the final dashboard and add the public link in `tableau/dashboard_links.md`.
6. Export the final report and presentation as PDFs into `reports/`.

### Quick Start

If you are working locally:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
jupyter notebook
```

If you are working in Google Colab:

- Upload or sync the notebooks from `notebooks/`
- Keep the final `.ipynb` files committed to GitHub
- Export any cleaned datasets into `data/processed/`

---

### Project Overview

| Field | Details |
|---|---|
| **Project Title** | Bank Term Deposit Campaign Intelligence |
| **Sector** | Finance / Bank Marketing |
| **Team ID** | Sec-B G-2 |
| **Section** | Sec-B |
| **Faculty Mentor** | To be updated |
| **Institute** | Newton School of Technology, Rishihood University |
| **Submission Date** | 29 April 2026 |

### Team Members

| Role | Name | GitHub Username |
|---|---|---|
| Project Lead | Harshvardhan Gupta | `Harsh-gupta-07` |
| Data Lead | Harshvardhan Gupta | `Harsh-gupta-07` |
| ETL Lead | Kartik Maddan | `madaankartik` |
| Analysis Lead | Dev Tyagi | `devtyagi108` |
| Visualization Lead | Harshvardhan Gupta | `Harsh-gupta-07` |
| Strategy Lead | Shitanshu Tiwari | `sitanshu1018` |
| PPT and Quality Lead | Kartik Madaan | `madaankartik` |

---

## Business Problem

Retail banking campaigns often contact many clients, but only a small share subscribe to a term deposit. The bank needs a better way to decide which customers should be contacted first, which channel should be used, and when repeated outreach becomes wasteful. This project helps a campaign manager identify the customer segments and campaign patterns that lead to higher subscription conversion.

**Core Business Question**

> Which clients are most likely to subscribe to a term deposit, and how should the bank prioritize outreach to improve conversion?

**Decision Supported**

> Prioritize high-propensity client segments, use the best contact channel, and reduce repeated low-yield calling to improve campaign efficiency.

---

## Dataset

| Attribute | Details |
|---|---|
| **Source Name** | UCI Machine Learning Repository |
| **Direct Access Link** | https://archive.ics.uci.edu/ml/datasets/bank+marketing |
| **Row Count** | 45,211 |
| **Column Count** | 17 raw columns, 22 cleaned / derived columns |
| **Time Period Covered** | Historical bank marketing campaign data |
| **Format** | CSV |

**Key Columns Used**

| Column Name | Description | Role in Analysis |
|---|---|---|
| `age` | Client age in years | Segmentation |
| `job` | Occupation category | KPI / filter / segmentation |
| `marital` | Marital status | Segmentation |
| `education` | Education level | Segmentation |
| `balance` | Average yearly balance | Financial segmentation |
| `contact` | Contact communication type | Channel analysis |
| `campaign` | Number of contacts in current campaign | Touch intensity analysis |
| `previous` | Number of prior contacts | Campaign history |
| `pdays` | Days since last contact | Prior-contact flag |
| `poutcome` | Outcome of previous campaign | Root-cause analysis |
| `y` | Subscription outcome | North-star KPI |

For full column definitions, see [`docs/data_dictionary.md`](docs/data_dictionary.md).

---

## KPI Framework

| KPI | Definition | Formula / Computation |
|---|---|---|
| Subscription Rate | Share of clients who subscribed | `sum(y) / count(rows)` |
| Total Subscriptions | Total number of term-deposit subscriptions | `sum(y)` |
| Avg. Contacts per Client | Average number of contacts in the current campaign | `mean(campaign)` |
| Previously Contacted Rate | Share of clients contacted before | `mean(previously_contacted)` |
| Avg. Yearly Balance | Mean account balance | `mean(balance)` |
| Default Rate | Share of clients with default = yes | `count(default == yes) / count(rows)` |
| Loan Rate | Share of clients with personal loan = yes | `count(loan == yes) / count(rows)` |

Document KPI logic clearly in `notebooks/04_statistical_analysis.ipynb` and `notebooks/05_final_load_prep.ipynb`.

---

## Tableau Dashboard

| Item | Details |
|---|---|
| **Dashboard URL** | https://public.tableau.com/views/Bank_Marketing_Analysis_17773884323250/CUSTOMERSEGMENTATIONSUBSCRIPTIONPROFILE?:language=en-US&publish=yes&:sid=&:redirect=auth&:display_count=n&:origin=viz_share_link |
| **Executive View** | A high-level summary of overall subscription performance, customer segments, and balance distribution. |
| **Operational View** | Drill-down views showing contact strategy, campaign touch bands, prior-contact lift, and risk/balance segmentation. |
| **Main Filters** | Job, age group, education, marital status, balance band, contact type, campaign band, prior contact, previous outcome |

Store dashboard screenshots in [`tableau/screenshots/`](tableau/screenshots/) and document the public links in [`tableau/dashboard_links.md`](tableau/dashboard_links.md).

---

## Key Insights

1. Students and retirees convert best, so they should be prioritized early in the call queue.
2. Clients aged 65+ show the highest conversion, making age a strong targeting signal.
3. Previously contacted clients convert far better than cold leads, so warm leads should be routed first.
4. The cellular channel outperforms telephone, so outreach should be cellular-first wherever possible.
5. Conversion drops as campaign touches increase, which means repeated calling becomes inefficient after a point.
6. March, December, September, and October are the strongest months for conversion, so campaign timing matters.
7. Higher balance bands convert better than low or negative balance bands, so financial strength is a useful filter.
8. Prior campaign success is the strongest predictor of conversion, so previous outcome should be a premium routing rule.
9. Blue-collar and service-oriented segments sit near the lower end of conversion, so they should not dominate premium outreach.
10. Education and marital status matter, but they are secondary signals compared with prior contact and channel choice.

---

## Recommendations

| # | Insight | Recommendation | Expected Impact |
|---|---|---|---|
| 1 | Students, retirees, and 65+ clients convert best | Prioritize these segments in the first call wave | Higher conversion rate |
| 2 | Previously contacted clients perform much better | Build a warm-lead routing rule in CRM | Better lead efficiency |
| 3 | Conversion drops with repeated touches | Cap repeated outreach after 3 calls | Lower wasted effort |
| 4 | Cellular performs better than telephone | Make cellular the default contact channel | Improved response rate |
| 5 | Certain months perform strongly | Schedule heavier campaigns in peak months | Better campaign yield |


## Repository Structure

```text
Sec-B_G-2_BankMarketing/
|
|-- README.md
|
|-- data/
|   |-- raw/                         # Original dataset (never edited)
|   `-- processed/                   # Cleaned output from ETL pipeline
|
|-- notebooks/
|   |-- 01_extraction.ipynb
|   |-- 02_cleaning.ipynb
|   |-- 03_eda.ipynb
|   |-- 04_statistical_analysis.ipynb
|   `-- 05_final_load_prep.ipynb
|
|-- scripts/
|   `-- etl_pipeline.py
|
|-- tableau/
|   |-- screenshots/
|   `-- dashboard_links.md
|
|-- reports/
|   |-- README.md
|   |-- project_report_template.md
|   `-- presentation_outline.md
|
|-- docs/
|   `-- data_dictionary.md
|
|-- DVA-oriented-Resume/
`-- DVA-focused-Portfolio/
```

---

## Analytical Pipeline

The project follows a structured 7-step workflow:

1. **Define** - Sector selected, problem statement scoped, mentor approval obtained.
2. **Extract** - Raw dataset sourced and committed to `data/raw/`; data dictionary drafted.
3. **Clean and Transform** - Cleaning pipeline built in `notebooks/02_cleaning.ipynb` and optionally `scripts/etl_pipeline.py`.
4. **Analyze** - EDA and statistical analysis performed in notebooks `03` and `04`.
5. **Visualize** - Interactive Tableau dashboard built and published on Tableau Public.
6. **Recommend** - 3-5 data-backed business recommendations delivered.
7. **Report** - Final project report and presentation deck completed and exported to PDF in `reports/`.

---

## Tech Stack

| Tool | Status | Purpose |
|---|---|---|
| Python + Jupyter Notebooks | Mandatory | ETL, cleaning, analysis, and KPI computation |
| Google Colab | Supported | Cloud notebook execution environment |
| Tableau Public | Mandatory | Dashboard design, publishing, and sharing |
| GitHub | Mandatory | Version control, collaboration, contribution audit |
| SQL | Optional | Initial data extraction only, if documented |

**Recommended Python libraries:** `pandas`, `numpy`, `matplotlib`, `seaborn`, `scipy`, `statsmodels`

---

## Evaluation Rubric

| Area | Marks | Focus |
|---|---|---|
| Problem Framing | 10 | Is the business question clear and well-scoped? |
| Data Quality and ETL | 15 | Is the cleaning pipeline thorough and documented? |
| Analysis Depth | 25 | Are statistical methods applied correctly with insight? |
| Dashboard and Visualization | 20 | Is the Tableau dashboard interactive and decision-relevant? |
| Business Recommendations | 20 | Are insights actionable and well-reasoned? |
| Storytelling and Clarity | 10 | Is the presentation professional and coherent? |
| **Total** | **100** | |

> Marks are awarded for analytical thinking and decision relevance, not chart quantity, visual decoration, or code length.

---

## Submission Checklist

**GitHub Repository**

- [x] Public repository created with the correct naming convention (`Sec-B_G-2_BankMarketing`)
- [x] All notebooks committed in `.ipynb` format
- [x] `data/raw/` contains the original, unedited dataset
- [x] `data/processed/` contains the cleaned pipeline output
- [x] `tableau/screenshots/` contains dashboard screenshots
- [x] `tableau/dashboard_links.md` contains the Tableau Public URL
- [x] `docs/data_dictionary.md` is complete
- [x] `README.md` explains the project, dataset, and team
- [x] All members have visible commits and pull requests

**Tableau Dashboard**

- [x] Published on Tableau Public and accessible via public URL
- [x] At least one interactive filter included
- [x] Dashboard directly addresses the business problem

**Project Report**

- [x] Final report exported as PDF into `reports/`
- [x] Cover page, executive summary, sector context, problem statement
- [x] Data description, cleaning methodology, KPI framework
- [x] EDA with written insights, statistical analysis results
- [x] Dashboard screenshots and explanation
- [x] 8-12 key insights in decision language
- [x] 3-5 actionable recommendations with impact estimates
- [x] Contribution matrix matches GitHub history

**Presentation Deck**

- [x] Final presentation exported as PDF into `reports/`
- [x] Title slide through recommendations, impact, limitations, and next steps

**Individual Assets**

- [x] DVA-oriented resume updated to include this capstone
- [x] Portfolio link or project case study added

---

## Contribution Matrix

This table must match evidence in GitHub Insights, PR history, and committed files.

| Team Member | Dataset and Sourcing | ETL and Cleaning | EDA and Analysis | Statistical Analysis | Tableau Dashboard | Report Writing | PPT and Viva |
|---|---|---|---|---|---|---|---|
| Harshvardhan Gupta | ✅ | ✅ | ✅ | ✅ |  |  |  |
| Karthik Madaan |  |  |  | ✅ | ✅ | ✅ | ✅ |
| Dev Tyagi |  |  |  | ✅ |  | ✅ | ✅ |
| Sitanshu Tiwari |  |  |  |  | ✅ |  |  |
| Siddarth Dangi |  |  |  |  |  |  |  |
| Sathvik Mani Tripathi |  |  |  |  |  |  |  |

_Declaration: We confirm that the above contribution details are accurate and verifiable through GitHub Insights, PR history, and submitted artifacts._

**Team Lead Name:** Harshvardhan Gupta

**Date:** 26 April 2026

---

## Academic Integrity

All analysis, code, and recommendations in this repository must be the original work of the team listed above. Free-riding is tracked via GitHub Insights and pull request history. Any mismatch between the contribution matrix and actual commit history may result in individual grade adjustments.

---

*Newton School of Technology - Data Visualization & Analytics | Capstone 2*
