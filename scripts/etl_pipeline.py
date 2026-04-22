"""ETL pipeline for the Bank Marketing dataset.

This script reads the raw `bank-full.csv` file, applies dataset-specific cleaning,
adds derived columns used in analysis/Tableau, and writes processed outputs.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

EXPECTED_RAW_COLUMNS = {
    "age",
    "job",
    "marital",
    "education",
    "default",
    "balance",
    "housing",
    "loan",
    "contact",
    "day",
    "month",
    "duration",
    "campaign",
    "pdays",
    "previous",
    "poutcome",
    "y",
}

TABLEAU_COLUMNS = [
    "age",
    "age_group",
    "job",
    "marital",
    "education",
    "default",
    "balance",
    "balance_band",
    "housing",
    "loan",
    "contact",
    "day",
    "month",
    "duration",
    "duration_minutes",
    "campaign",
    "campaign_band",
    "pdays",
    "previous",
    "previously_contacted",
    "poutcome",
    "y",
]


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result.columns = (
        result.columns.str.strip()
        .str.lower()
        .str.replace(r"[^a-z0-9]+", "_", regex=True)
        .str.strip("_")
    )
    return result


def validate_expected_columns(df: pd.DataFrame) -> None:
    missing = sorted(EXPECTED_RAW_COLUMNS.difference(df.columns))
    if missing:
        raise ValueError(f"Missing required columns: {missing}")


def standardize_string_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    for column in result.select_dtypes(include="object").columns:
        result[column] = result[column].astype("string").str.strip().str.lower()
    return result


def add_derived_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()

    target_map = {"no": 0, "yes": 1}
    if not set(result["y"].dropna().unique()).issubset(target_map.keys()):
        raise ValueError("Column 'y' has unexpected values; expected only 'yes'/'no'.")

    result["previously_contacted"] = (result["pdays"] != -1).astype("int8")
    result["duration_minutes"] = (result["duration"] / 60).round(2)

    result["age_group"] = pd.cut(
        result["age"],
        bins=[0, 24, 34, 44, 54, 64, 200],
        labels=["18-24", "25-34", "35-44", "45-54", "55-64", "65+"],
        include_lowest=True,
    )
    result["balance_band"] = pd.cut(
        result["balance"],
        bins=[-100000, 0, 1000, 5000, 1000000],
        labels=["negative_or_zero", "low", "medium", "high"],
    )
    result["campaign_band"] = pd.cut(
        result["campaign"],
        bins=[0, 1, 3, 5, 100],
        labels=["1", "2-3", "4-5", "6+"],
        include_lowest=True,
    )

    return result


def clean_bank_marketing(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = normalize_columns(df)
    validate_expected_columns(cleaned)
    cleaned = cleaned.drop_duplicates().reset_index(drop=True)
    cleaned = standardize_string_columns(cleaned)
    cleaned = add_derived_columns(cleaned)
    return cleaned


def build_clean_dataset(input_path: Path, sep: str = ";") -> pd.DataFrame:
    raw_df = pd.read_csv(input_path, sep=sep)
    return clean_bank_marketing(raw_df)


def build_tableau_ready_dataset(clean_df: pd.DataFrame) -> pd.DataFrame:
    return clean_df[TABLEAU_COLUMNS].copy()


def build_kpi_summary(clean_df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "kpi": [
                "total_contacts",
                "total_subscriptions",
                "subscription_rate",
                "average_age",
                "average_balance",
                "median_campaign_count",
                "previously_contacted_share",
            ],
            "value": [
                len(clean_df),
                int(clean_df["y"].sum()),
                clean_df["y"].mean(),
                clean_df["age"].mean(),
                clean_df["balance"].mean(),
                clean_df["campaign"].median(),
                clean_df["previously_contacted"].mean(),
            ],
        }
    )


def save_csv(df: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run ETL pipeline for the bank marketing dataset.")
    parser.add_argument(
        "--input",
        required=True,
        type=Path,
        help="Path to raw CSV (e.g., data/raw/bank-full.csv).",
    )
    parser.add_argument(
        "--output",
        required=True,
        type=Path,
        help="Path to cleaned CSV output (e.g., data/processed/cleaned_dataset.csv).",
    )
    parser.add_argument(
        "--sep",
        default=";",
        help="CSV delimiter for raw input. Defaults to ';'.",
    )
    parser.add_argument(
        "--tableau-output",
        type=Path,
        default=None,
        help="Optional path for Tableau-ready row-level CSV.",
    )
    parser.add_argument(
        "--kpi-output",
        type=Path,
        default=None,
        help="Optional path for KPI summary CSV.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    clean_df = build_clean_dataset(args.input, sep=args.sep)
    save_csv(clean_df, args.output)

    print(f"Cleaned dataset saved to: {args.output}")
    print(f"Rows: {len(clean_df):,} | Columns: {len(clean_df.columns):,}")
    print(f"Subscriptions: {int(clean_df['y'].sum()):,} | Rate: {clean_df['y'].mean():.2%}")

    if args.tableau_output is not None:
        tableau_df = build_tableau_ready_dataset(clean_df)
        save_csv(tableau_df, args.tableau_output)
        print(f"Tableau-ready dataset saved to: {args.tableau_output}")

    if args.kpi_output is not None:
        kpi_df = build_kpi_summary(clean_df)
        save_csv(kpi_df, args.kpi_output)
        print(f"KPI summary saved to: {args.kpi_output}")


if __name__ == "__main__":
    main()
