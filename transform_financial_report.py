"""
Transform Financial Report Excel → 3 output CSVs.
Replicates the exact Power BI / Power Query M transformations.

Usage:
    python transform_financial_report.py
    python transform_financial_report.py --input "sample-data/input/Financial report YTD 2026.01.xlsx"
"""

import re
import argparse
import pandas as pd
from pathlib import Path

INPUT_DIR   = Path("sample-data/input")
OUTPUT_DIR  = Path("sample-data/output")

MONTH_ABBR  = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
BU_KEEP     = {"BFSI", "DDN", "DDX", "DHN", "JSV"}   # "Dev" row is loaded then dropped

# ── Helpers ──────────────────────────────────────────────────────────────────

def parse_year_month(filename: str) -> tuple[int, int]:
    """Extract (year, month) from 'Financial report YTD 2026.01.xlsx'."""
    m = re.search(r"YTD (\d{4})\.(\d{2})", filename)
    if not m:
        raise ValueError(f"Cannot parse year/month from: {filename}")
    return int(m.group(1)), int(m.group(2))


def month_abbr(month: int) -> str:
    return MONTH_ABBR[month - 1]

def month_full(month: int) -> str:
    from datetime import date
    return date(2020, month, 1).strftime("%B")


# ── Data exp ─────────────────────────────────────────────────────────────────

def transform_data_exp(file_path: Path) -> pd.DataFrame:
    """
    Sheet 'Data exp' — expense transaction ledger.

    Transform File function:
      - Skip 2 title rows, promote row 3 as header
      - Filter: (Diễn giải is null OR == 'Diễn giải') AND BU != '0'
      - Drop dimension code/name columns

    Main query (after stacking files):
      - Parse Report Year / Report Month from filename
      - Normalize Cost center: SUPPORT→Support, SALE→Sale
      - Add Software/JSV column
      - Drop rows where Cost center is null
    """
    filename = file_path.name
    report_year, report_month = parse_year_month(filename)

    # ── Transform File step ───────────────────────────────────────────────
    raw = pd.read_excel(file_path, sheet_name="Data exp", header=2, dtype=str)

    keep = [
        "Ngày hạch toán", "Ngày chứng từ", "Số chứng từ", "Diễn giải",
        "Tài khoản", "TK đối ứng",
        "Phát sinh Nợ (ROW)", "Phát sinh Có (ROW)", "Dư Nợ", "Dư Có",
        "Phát sinh Nợ", "Phát sinh Có",
        "BU", "Cost center", "Cost items",
    ]
    raw = raw[[c for c in keep if c in raw.columns]].copy()

    # Filter rows: (Diễn giải null or == "Diễn giải") AND BU != "0"
    diễn_giải = raw["Diễn giải"].fillna("")
    raw = raw[
        ((diễn_giải == "") | (diễn_giải == "Diễn giải")) &
        (raw["BU"].fillna("") != "0")
    ].copy()

    # ── Main query steps ──────────────────────────────────────────────────
    raw.insert(0, "Source.Name", filename)
    raw["Report Year"]       = report_year
    raw["Report Month"]      = report_month
    raw["Report Month Name"] = month_abbr(report_month)

    # Strip trailing/leading spaces from all text columns
    for col in ["Ngày hạch toán", "Ngày chứng từ", "Số chứng từ", "Diễn giải",
                "Tài khoản", "TK đối ứng", "BU", "Cost center", "Cost items"]:
        if col in raw.columns:
            raw[col] = raw[col].astype(str).str.strip().replace("nan", pd.NA)

    raw["Cost center"] = raw["Cost center"].replace({"SUPPORT": "Support", "SALE": "Sale"})
    raw["Software/JSV"] = raw["BU"].apply(lambda x: "JSV" if str(x).strip() == "JSV" else "Software")

    # Drop rows where Cost center is null
    raw = raw[raw["Cost center"].notna() & (raw["Cost center"].str.strip() != "")].copy()

    # Cast currency columns to float
    currency_cols = [
        "Phát sinh Nợ (ROW)", "Phát sinh Có (ROW)", "Dư Nợ", "Dư Có",
        "Phát sinh Nợ", "Phát sinh Có",
    ]
    for col in currency_cols:
        if col in raw.columns:
            raw[col] = pd.to_numeric(raw[col], errors="coerce")

    raw["Report Year"]  = raw["Report Year"].astype(int)
    raw["Report Month"] = raw["Report Month"].astype(int)

    col_order = [
        "Source.Name",
        "Ngày hạch toán", "Ngày chứng từ", "Số chứng từ", "Diễn giải",
        "Tài khoản", "TK đối ứng",
        "Phát sinh Nợ (ROW)", "Phát sinh Có (ROW)", "Dư Nợ", "Dư Có",
        "Phát sinh Nợ", "Phát sinh Có",
        "BU", "Cost center", "Cost items",
        "Report Year", "Report Month", "Report Month Name", "Software/JSV",
    ]
    raw = raw[[c for c in col_order if c in raw.columns]]
    return raw.reset_index(drop=True)


# ── Data rev ─────────────────────────────────────────────────────────────────

def transform_data_rev(file_path: Path) -> pd.DataFrame:
    """
    Sheet 'Data Rev' — revenue by BU in wide format (months as columns).

    Transform File (2) function:
      - Skip 2 rows, remove blank rows, promote headers
      - Keep rows: Items in {BFSI, DDN, DDX, Dev, DHN, JSV}
      - Remove last row
      - Fill null Unit with 'Tr.VND'
      - Drop Dev row, drop (%) and Y2026 columns
      - Unpivot Jan–Dec → Month name / Revenue
      - Add Month number (1–12)
      - Rename columns

    Main query:
      - Parse Report Year / Report Month from filename
      - Add Report Month Name
      - Add Is_month_active: 1 if Month Number <= Report Month Number, else 0
      - Replace DDX → DX in BU
    """
    filename = file_path.name
    report_year, report_month = parse_year_month(filename)

    # ── Transform File step ───────────────────────────────────────────────
    raw = pd.read_excel(file_path, sheet_name="Data Rev", header=None)

    # Skip 2 rows, remove fully blank rows, promote row 3 as header
    raw = raw.iloc[2:].reset_index(drop=True)
    raw = raw.dropna(how="all")
    raw.columns = raw.iloc[0]
    raw = raw.iloc[1:].reset_index(drop=True)

    # Keep only BU rows
    raw = raw[raw["Items"].isin(BU_KEEP | {"Dev"})].copy()

    # Remove last row (RemoveLastN 1)
    raw = raw.iloc[:-1].copy()

    # Drop "Note" and any unnamed/None columns (M code drops Note, Column18, Column19)
    drop_cols = [c for c in raw.columns if c == "Note" or not isinstance(c, str) or c.startswith("Column")]
    raw = raw.drop(columns=drop_cols, errors="ignore")

    # Fill null Unit with "Tr.VND"
    raw["Unit"] = raw["Unit"].fillna("Tr.VND")

    # Drop Dev row
    raw = raw[raw["Items"] != "Dev"].copy()

    # Drop (%) and Y2026 columns
    raw = raw.drop(columns=["(%)", "Y2026"], errors="ignore")

    # Unpivot month columns Jan–Dec
    month_cols = [c for c in MONTH_ABBR if c in raw.columns]
    df = raw.melt(id_vars=["Items", "Unit"], value_vars=month_cols,
                  var_name="Month name", value_name="Revenue")

    # Add Month number
    df["Month number"] = df["Month name"].apply(lambda x: MONTH_ABBR.index(x) + 1)

    # Rename
    df = df.rename(columns={"Items": "BU"})
    df["Revenue"] = pd.to_numeric(df["Revenue"], errors="coerce").fillna(0).astype(int)

    # ── Main query steps ──────────────────────────────────────────────────
    df.insert(0, "Source.Name", filename)
    df["Report Year"]         = report_year
    df["Report Month Number"] = report_month
    df["Report Month Name"]   = month_full(report_month)

    # Is_month_active: 1 if month number <= report month number
    df["Is_month_active"] = (df["Month number"] <= report_month).astype(int)

    # Rename month columns to final names
    df = df.rename(columns={
        "Month name":   "Month Name",
        "Month number": "Month Number",
    })

    # Replace DDX → DX
    df["BU"] = df["BU"].replace("DDX", "DX")

    col_order = [
        "Source.Name", "BU", "Month Name", "Month Number",
        "Revenue", "Unit",
        "Report Year", "Report Month Number", "Report Month Name",
        "Is_month_active",
    ]
    df = df[[c for c in col_order if c in df.columns]]
    return df.reset_index(drop=True)


# ── Data HC ──────────────────────────────────────────────────────────────────

def transform_data_hc(file_path: Path) -> pd.DataFrame:
    """
    Sheet 'Data HC' — headcount (man-months) pivot: rows=cost centers, cols=BUs.

    Transform File (3) function:
      - Promote row 1 as headers
      - Find column starting with 'T' (e.g. ' T01 ') → rename to 'Cost center'
      - Drop the ' Y2026 ' summary column
      - Unpivot BU columns → long format (BU, MM)
      - Cast MM to number; replace null/errors with 0
      - Replace BOD → BOM in BU

    Main query:
      - Parse Report Year / Report Month from filename
      - Add Report Month Name (3-char)
      - Filter out ' Total MM ' rows (before trim)
      - Trim & clean Cost center
      - Add HC Group: Delivery/Sale/Support → 'Total MM'; Intern → 'Intern'
      - Trim BU; replace HN→DHN, DN→DDN
    """
    filename = file_path.name
    report_year, report_month = parse_year_month(filename)

    # ── Transform File step ───────────────────────────────────────────────
    raw = pd.read_excel(file_path, sheet_name="Data HC", header=None, dtype=str)

    # Promote first row as headers
    raw.columns = raw.iloc[0]
    raw = raw.iloc[1:].reset_index(drop=True)

    # Find column starting with 'T' after trimming (e.g. ' T01 ') → Cost center
    month_col = next(
        (c for c in raw.columns if isinstance(c, str) and c.strip().startswith("T")),
        None,
    )
    if month_col is None:
        raise ValueError("Could not find month column (starting with 'T') in Data HC sheet")
    raw = raw.rename(columns={month_col: "Cost center"})

    # Drop the Year summary column (e.g. ' Y2026 ')
    year_col = next(
        (c for c in raw.columns if isinstance(c, str) and c.strip().startswith("Y")),
        None,
    )
    if year_col:
        raw = raw.drop(columns=[year_col])

    # Unpivot BU columns
    bu_cols = [c for c in raw.columns if c != "Cost center"]
    df = raw.melt(id_vars=["Cost center"], value_vars=bu_cols,
                  var_name="BU", value_name="MM")

    # Cast MM to numeric (errors → 0, null → 0)
    df["MM"] = pd.to_numeric(df["MM"], errors="coerce").fillna(0)

    # Replace BOD → BOM
    df["BU"] = df["BU"].str.strip().replace("BOD", "BOM")

    # ── Main query steps ──────────────────────────────────────────────────
    df.insert(0, "Source.Name", filename)
    df["Report Year"]       = report_year
    df["Report Month"]      = report_month
    df["Report Month Name"] = month_abbr(report_month)

    # Filter out ' Total MM ' rows (raw value before trim)
    df = df[df["Cost center"] != " Total MM "].copy()

    # Trim & clean Cost center
    df["Cost center"] = df["Cost center"].str.strip()

    # Add HC Group
    hc_group_map = {"Delivery": "Total MM", "Sale": "Total MM", "Support": "Total MM", "Intern": "Intern"}
    df["HC Group"] = df["Cost center"].map(hc_group_map)

    # Trim BU, then normalize codes
    df["BU"] = df["BU"].str.strip()
    df["BU"] = df["BU"].replace({"HN": "DHN", "DN": "DDN"})

    # Drop rows with 0 MM to keep output clean (optional — mirrors the filter null step)
    # df = df[df["MM"] != 0].copy()

    col_order = [
        "Source.Name", "Cost center", "BU", "MM",
        "Report Year", "Report Month", "Report Month Name", "HC Group",
    ]
    df = df[[c for c in col_order if c in df.columns]]
    return df.reset_index(drop=True)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", default=str(INPUT_DIR), help="Folder containing input Excel files")
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="Output directory for CSVs")
    args = parser.parse_args()

    input_dir  = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Collect all matching files, sorted ascending by name
    all_files = sorted(input_dir.glob("Financial report YTD *.xlsx"))
    if not all_files:
        print(f"No matching files found in {input_dir}")
        return

    print(f"Input dir:  {input_dir}")
    print(f"Files:      {[f.name for f in all_files]}")
    print(f"Output dir: {output_dir}\n")

    # data_exp — stack all files
    print("Processing data_exp ...")
    df_exp = pd.concat([transform_data_exp(f) for f in all_files], ignore_index=True)
    df_exp.to_csv(output_dir / "data_exp.csv", index=False, encoding="utf-8-sig")
    print(f"  -> {len(df_exp):,} rows x {len(df_exp.columns)} cols  =>  data_exp.csv")

    # data_rev — latest file only
    print("Processing data_rev ...")
    df_rev = transform_data_rev(all_files[-1])
    df_rev.to_csv(output_dir / "data_rev.csv", index=False, encoding="utf-8-sig")
    print(f"  -> {len(df_rev):,} rows x {len(df_rev.columns)} cols  =>  data_rev.csv")

    # data_hc — stack all files
    print("Processing data_hc ...")
    df_hc = pd.concat([transform_data_hc(f) for f in all_files], ignore_index=True)
    df_hc.to_csv(output_dir / "data_hc.csv", index=False, encoding="utf-8-sig")
    print(f"  -> {len(df_hc):,} rows x {len(df_hc.columns)} cols  =>  data_hc.csv")

    print("\nDone.")


if __name__ == "__main__":
    main()
