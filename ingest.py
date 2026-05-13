"""Ingest SharePoint Excel files (Financial report YTD YYYY.MM.xlsx) into Postgres.

Loads every matching workbook in $SP_LOCAL_DIR into 3 consolidated tables:
    data_exp, data_rev, data_hc

Each row carries a "Source.Name" column (e.g. "Financial report YTD 2026.01.xlsx")
identifying which month's snapshot it came from. Re-running a single file replaces
only that file's rows (snapshot-replace); adding a new month appends incrementally.
A single transaction wraps the whole run — anything failing rolls every load back.
"""

import io
import os
import re
from pathlib import Path

import psycopg
from dotenv import load_dotenv

from transform_financial_report import (
    transform_data_exp,
    transform_data_rev,
    transform_data_hc,
)

load_dotenv()

# ---------- per-sheet target schemas (match transform_financial_report output) ----------

DATA_EXP_COLUMNS = [
    "Source.Name",
    "Ngày hạch toán", "Ngày chứng từ", "Số chứng từ", "Diễn giải",
    "Tài khoản", "TK đối ứng",
    "Phát sinh Nợ (ROW)", "Phát sinh Có (ROW)", "Dư Nợ", "Dư Có",
    "Phát sinh Nợ", "Phát sinh Có",
    "BU", "Cost center", "Cost items",
    "Report Year", "Report Month", "Report Month Name", "Software/JSV",
]
DATA_EXP_TYPES = {
    "Phát sinh Nợ (ROW)": "numeric", "Phát sinh Có (ROW)": "numeric",
    "Dư Nợ": "numeric", "Dư Có": "numeric",
    "Phát sinh Nợ": "numeric", "Phát sinh Có": "numeric",
    "Report Year": "int", "Report Month": "int",
}

DATA_REV_COLUMNS = [
    "Source.Name", "BU", "Month Name", "Month Number",
    "Revenue", "Unit",
    "Report Year", "Report Month Number", "Report Month Name",
    "Is_month_active",
]
DATA_REV_TYPES = {
    "Month Number": "int", "Revenue": "bigint",
    "Report Year": "int", "Report Month Number": "int",
    "Is_month_active": "int",
}

DATA_HC_COLUMNS = [
    "Source.Name", "Cost center", "BU", "MM",
    "Report Year", "Report Month", "Report Month Name", "HC Group",
]
DATA_HC_TYPES = {
    "MM": "numeric",
    "Report Year": "int", "Report Month": "int",
}

SHEET_CONFIGS = [
    {"transform": transform_data_exp, "columns": DATA_EXP_COLUMNS,
     "column_types": DATA_EXP_TYPES, "table": "data_exp"},
    {"transform": transform_data_rev, "columns": DATA_REV_COLUMNS,
     "column_types": DATA_REV_TYPES, "table": "data_rev"},
    {"transform": transform_data_hc,  "columns": DATA_HC_COLUMNS,
     "column_types": DATA_HC_TYPES,  "table": "data_hc"},
]

# ---------- file discovery ----------

FILE_RE = re.compile(r"^Financial report YTD (\d{4})\.(0[1-9]|1[0-2])\.xlsx$", re.IGNORECASE)


def discover_files(base: Path) -> list[Path]:
    return sorted(
        (p for p in base.iterdir() if FILE_RE.match(p.name)),
        key=lambda p: p.name,
    )


def load_snapshot(cur, table, columns, df, column_types, src_file):
    df = df.reindex(columns=columns)
    csv_buf = io.StringIO()
    df.to_csv(csv_buf, index=False, header=False, na_rep="")
    cols_sql = ", ".join(f'"{c}"' for c in columns)  # quote: spaces/non-ASCII/dots
    create_cols = ", ".join(f'"{c}" {column_types.get(c, "TEXT")}' for c in columns)
    cur.execute(f"CREATE TABLE IF NOT EXISTS {table} ({create_cols})")
    cur.execute(f'CREATE INDEX IF NOT EXISTS {table}_source_name_idx ON {table} ("Source.Name")')
    cur.execute(f'DELETE FROM {table} WHERE "Source.Name" = %s', (src_file,))
    with cur.copy(f"COPY {table} ({cols_sql}) FROM STDIN WITH (FORMAT csv, NULL '')") as cp:
        cp.write(csv_buf.getvalue())
    print(f"loaded {table} <- {src_file} ({len(df)} rows)")


def main() -> None:
    base = Path(os.environ["SP_LOCAL_DIR"])
    files = discover_files(base)
    with psycopg.connect(os.environ["PG_DSN"]) as conn, conn.cursor() as cur:
        for filepath in files:
            for cfg in SHEET_CONFIGS:
                df = cfg["transform"](filepath)
                load_snapshot(cur, cfg["table"], cfg["columns"], df,
                              cfg["column_types"], filepath.name)
    print("done")


if __name__ == "__main__":
    main()
