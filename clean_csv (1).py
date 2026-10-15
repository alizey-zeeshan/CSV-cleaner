"""
CSV Cleaner: turns a messy CSV file into a clean Excel file.

Usage:
    python clean_csv.py messy_customers.csv
    python clean_csv.py messy_customers.csv -o clean_customers.xlsx
"""
import argparse
import re

import pandas as pd

EMAIL_PATTERN = r"^[\w.+-]+@[\w-]+\.[\w.-]+$"


def clean_column_names(df):
    """' Customer Name ' -> 'customer_name'"""
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(r"[^a-z0-9]+", "_", regex=True)
        .str.strip("_")
    )
    return df


def clean_text(df):
    """Trim spaces, collapse double spaces, turn blanks into real missing values."""
    for col in df.columns:
        df[col] = (
            df[col]
            .astype("string")
            .str.strip()
            .str.replace(r"\s+", " ", regex=True)
            .replace("", pd.NA)
        )
    return df


def clean_email(series):
    """Lowercase emails and blank out invalid ones."""
    series = series.str.lower()
    return series.where(series.str.match(EMAIL_PATTERN, na=False), pd.NA)


def clean_phone(series):
    """Keep digits only. Format 10-digit numbers as 555-123-4567."""
    digits = series.str.replace(r"\D", "", regex=True)
    formatted = digits.str.replace(r"^(\d{3})(\d{3})(\d{4})$", r"\1-\2-\3", regex=True)
    return formatted


def clean_date(series):
    """Read many date styles; anything unreadable becomes missing."""
    return pd.to_datetime(series, format="mixed", errors="coerce", dayfirst=False).dt.date


def clean_money(series):
    """'$1,200.50' -> 1200.5"""
    return pd.to_numeric(series.str.replace(r"[^\d.\-]", "", regex=True), errors="coerce")


def clean_names(series):
    """'  maria  GARCIA ' -> 'Maria Garcia'"""
    return series.str.title()


def clean_dataframe(df):
    report = {"rows_in": len(df)}

    df = clean_column_names(df)
    df = clean_text(df)

    df = df.dropna(how="all")  # remove fully empty rows
    report["empty_rows_removed"] = report["rows_in"] - len(df)

    # Pick a cleaning step based on the column name
    for col in df.columns:
        if "email" in col:
            df[col] = clean_email(df[col])
        elif "phone" in col:
            df[col] = clean_phone(df[col])
        elif "date" in col:
            df[col] = clean_date(df[col])
        elif any(word in col for word in ("amount", "price", "total", "cost")):
            df[col] = clean_money(df[col])
        elif "name" in col or "city" in col:
            df[col] = clean_names(df[col])

    # Fill missing text values so the file has no confusing blanks
    for col in df.select_dtypes(include="string").columns:
        df[col] = df[col].fillna("Unknown")

    # Remove duplicate people (same name + same phone or email)
    before = len(df)
    keys = [c for c in df.columns if "email" in c or "phone" in c]
    df = df.drop_duplicates(subset=keys, keep="first") if keys else df.drop_duplicates()
    report["duplicates_removed"] = before - len(df)

    report["rows_out"] = len(df)
    return df.reset_index(drop=True), report


def main():
    parser = argparse.ArgumentParser(description="Clean a messy CSV file.")
    parser.add_argument("input", help="path to the messy CSV file")
    parser.add_argument("-o", "--output", default="cleaned_data.xlsx", help="output Excel file")
    args = parser.parse_args()

    df = pd.read_csv(args.input, dtype=str, skipinitialspace=True)
    clean, report = clean_dataframe(df)
    clean.to_excel(args.output, index=False)

    print(f"Rows in:               {report['rows_in']}")
    print(f"Empty rows removed:    {report['empty_rows_removed']}")
    print(f"Duplicates removed:    {report['duplicates_removed']}")
    print(f"Rows out:              {report['rows_out']}")
    print(f"Saved to:              {args.output}")


if __name__ == "__main__":
    main()
