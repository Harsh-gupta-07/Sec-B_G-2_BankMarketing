# Data Dictionary

This dictionary documents the raw bank marketing dataset stored in `data/raw/bank-full.csv`.
Each row represents one client contact record from the UCI Bank Marketing dataset.

## Dataset Summary

| Item | Details |
|---|---|
| Dataset name | Bank Marketing Full |
| Source | UCI Machine Learning Repository |
| Raw file name | `bank-full.csv` |
| Last updated | 2026-04-20 |
| Granularity | One row per client contact record |

## Column Definitions

| Column Name | Data Type | Description | Example Value | Used In | Cleaning Notes |
|---|---|---|---|---|---|
| age | integer | Client age in years | 58 | EDA, statistical analysis, Tableau | Keep numeric |
| job | string | Occupation category | management | EDA, segmentation, Tableau | Strip whitespace; preserve `unknown` |
| marital | string | Marital status | married | EDA, segmentation, Tableau | Standardize casing; preserve `unknown` |
| education | string | Education level | tertiary | EDA, segmentation, Tableau | Keep source labels; preserve `unknown` |
| default | string | Whether the client has credit in default | no | EDA, segmentation, Tableau | Keep as categorical yes/no/unknown |
| balance | integer | Average yearly balance in euros | 2143 | EDA, statistical analysis, Tableau | Keep numeric; inspect outliers |
| housing | string | Whether the client has a housing loan | yes | EDA, segmentation, Tableau | Keep as categorical yes/no/unknown |
| loan | string | Whether the client has a personal loan | no | EDA, segmentation, Tableau | Keep as categorical yes/no/unknown |
| contact | string | Contact communication type | unknown | EDA, segmentation, Tableau | Preserve source categories |
| day | integer | Day of the month of the last contact | 5 | EDA, Tableau | Keep numeric |
| month | string | Month of the last contact | may | EDA, statistical analysis, Tableau | Keep as month code; derive month order where needed |
| duration | integer | Last contact duration in seconds | 261 | EDA, statistical analysis | Keep for descriptive analysis; do not use as a predictive feature |
| campaign | integer | Number of contacts performed during this campaign for this client | 1 | EDA, statistical analysis, Tableau | Keep numeric |
| pdays | integer | Days since the client was last contacted, or `-1` if never contacted before | -1 | EDA, statistical analysis, Tableau | Convert `-1` to a `previously_contacted` flag |
| previous | integer | Number of contacts performed before this campaign | 0 | EDA, statistical analysis, Tableau | Keep numeric |
| poutcome | string | Outcome of the previous campaign | unknown | EDA, statistical analysis, Tableau | Keep source categories |
| y | string | Whether the client subscribed to a term deposit | no | All notebooks |

## Derived Columns

| Derived Column | Logic | Business Meaning |
|---|---|---|
| y | `1` if `y == "yes"`, else `0` | Subscription outcome for analysis and KPI tracking |
| age_group | Bucket age into life-stage bands | Easier demographic segmentation for Tableau |
| balance_band | Bucket `balance` into low/medium/high ranges | Helps interpret client financial segments |
| campaign_band | Bucket `campaign` into low/medium/high touch groups | Helps identify over-contacted clients |
| previously_contacted | `1` if `pdays != -1`, else `0` | Whether the client had prior campaign exposure |
| duration_minutes | `duration / 60` | Easier interpretation of call duration in dashboards |
| month | Map month abbreviations to calendar order | Enables chronological sorting in visuals |

## Data Quality Notes

- The raw file is semicolon-delimited, not comma-delimited.
- The dataset contains no true missing values in the raw file, but several categorical fields use `unknown` as an explicit label.
- `pdays = -1` means the client was not contacted in a previous campaign.
- `duration` is informative for post-contact analysis but should not be used in a realistic pre-call prediction model because it leaks outcome information.
