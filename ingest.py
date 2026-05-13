"""Ingest SharePoint Excel files (Financial report YTD YYYY.MM.xlsx) into Postgres.

Discovers every matching workbook in $SP_LOCAL_DIR and loads each of its 3
sheets into `data_<year>_<mon>_<suffix>` tables. Applies the same transforms
as powerbi_cleaned.txt (Power Query M) before loading. Demo-quality: no error
handling, no logging, no tests. A single transaction wraps the whole run —
anything failing rolls every load back. Re-run after fixing.
"""

import io
import os
import re
from pathlib import Path

import pandas as pd
import psycopg

# ---------- per-sheet transforms (mirror powerbi_cleaned.txt) ----------

DATA_EXP_DROP = [
    "Mã đối tượng", "Tên đối tượng", "Mã khoản mục CP", "Tên khoản mục CP",
    "Mã đơn vị", "Tên đơn vị", "Mã đối tượng_1", "Tên đối tượng_2",
    "Mã công trình", "Tên công trình", "Mã thống kê", "Tên thống kê", "Column23",
]
DATA_EXP_FINAL = [
    "Ngày hạch toán", "Report Day", "Report Month", "Report Year",
    "Ngày chứng từ", "Số chứng từ", "Diễn giải", "Tài khoản", "TK đối ứng",
    "Phát sinh Nợ (ROW)", "Phát sinh Có (ROW)", "Dư Nợ", "Dư Có",
    "Phát sinh Nợ", "Phát sinh Có", "BU", "Cost center", "Cost items",
]


def transform_data_exp(df: pd.DataFrame) -> pd.DataFrame:
    df = df.drop(columns=[c for c in DATA_EXP_DROP if c in df.columns])

    df = df[(df["Phát sinh Nợ"] != "") & (df["BU"] != "") & (df["BU"] != "0")].copy()

    df["Cost center"] = df["Cost center"].replace({"SUPPORT": "Support", "SALE": "Sale"})

    # dd/mm/yyyy in the source. pd.to_datetime with dayfirst handles both
    # text-formatted dates and openpyxl-converted datetime strings.
    dates = pd.to_datetime(df["Ngày hạch toán"], dayfirst=True, errors="coerce")
    df["Report Day"]   = dates.dt.day.astype("Int64").astype(str).replace("<NA>", "")
    df["Report Month"] = dates.dt.month.astype("Int64").astype(str).replace("<NA>", "")
    df["Report Year"]  = dates.dt.year.astype("Int64").astype(str).replace("<NA>", "")
    df["Ngày hạch toán"] = dates.dt.strftime("%Y-%m-%d").fillna("")

    return df[DATA_EXP_FINAL]


DATA_REV_KEEP = {"BFSI", "DDN", "DDX", "Dev", "DHN", "JSV"}
DATA_REV_FINAL = ["BU", "Month name", "Revenue", "Unit", "Month number"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def transform_data_rev(df: pd.DataFrame) -> pd.DataFrame:
    # Read with header=None: skip top 2, drop fully-blank rows, promote next as header
    df = df.iloc[2:].reset_index(drop=True)
    df = df[~df.apply(lambda r: all(v == "" for v in r), axis=1)]
    df.columns = df.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)

    df = df[df["Items"].isin(DATA_REV_KEEP)]
    df = df.iloc[:-1]   # drop totals row at the bottom
    df = df.drop(columns=[c for c in ["Note", "Column18", "Column19"] if c in df.columns])
    df["Unit"] = df["Unit"].replace("", "Tr.VND")
    df = df[df["Items"] != "Dev"]
    df = df.drop(columns=[c for c in ["(%)", "Y2026"] if c in df.columns])

    df = df.melt(id_vars=["Items", "Unit"], var_name="Attribute", value_name="Value")
    df["MonthNum"] = df["Attribute"].map(
        lambda x: str(MONTHS.index(x) + 1) if x in MONTHS else ""
    )
    df = df.rename(columns={
        "Items": "BU", "Attribute": "Month name",
        "Value": "Revenue", "MonthNum": "Month number",
    })
    return df[DATA_REV_FINAL]


DATA_HC_FINAL = ["Cost center", "BU", "MM"]


def transform_data_hc(df: pd.DataFrame) -> pd.DataFrame:
    # First column whose trimmed header starts with "T" (e.g. "Tháng") → "Cost center"
    for col in list(df.columns):
        if str(col).strip().startswith("T"):
            df = df.rename(columns={col: "Cost center"})
            break

    df = df.drop(columns=[c for c in [" Y2026 "] if c in df.columns])
    df = df.melt(id_vars=["Cost center"], var_name="Attribute", value_name="Value")
    df = df.rename(columns={"Attribute": "BU", "Value": "MM"})

    df["MM"] = pd.to_numeric(df["MM"], errors="coerce").fillna(0).astype(str)
    df["BU"] = df["BU"].astype(str).str.replace("BOD", "BOM", regex=False)
    return df[DATA_HC_FINAL]


# ---------- per-sheet config ----------

SHEET_CONFIGS = {
    "Data exp": {
        "transform": transform_data_exp,
        "read_kwargs": {"header": 2, "dtype": str, "keep_default_na": False},
        "columns": DATA_EXP_FINAL,
        "suffix": "exp",
    },
    "Data Rev": {
        "transform": transform_data_rev,
        "read_kwargs": {"header": None, "dtype": str, "keep_default_na": False},
        "columns": DATA_REV_FINAL,
        "suffix": "rev",
    },
    "Data HC": {
        "transform": transform_data_hc,
        "read_kwargs": {"header": 0, "dtype": str, "keep_default_na": False},
        "columns": DATA_HC_FINAL,
        "suffix": "hc",
    },
}

MONTH_SUFFIX = ["jan", "feb", "mar", "apr", "may", "jun",
                "jul", "aug", "sep", "oct", "nov", "dec"]
FILE_RE = re.compile(r"^Financial report YTD (\d{4})\.(0[1-9]|1[0-2])\.xlsx$", re.IGNORECASE)


def discover_files(base: Path) -> list[tuple[Path, int, str]]:
    found = []
    for p in base.iterdir():
        m = FILE_RE.match(p.name)
        if m:
            year, month_num = int(m.group(1)), int(m.group(2))
            found.append((p, year, MONTH_SUFFIX[month_num - 1]))
    return sorted(found, key=lambda x: (x[1], MONTH_SUFFIX.index(x[2])))


def load_sheet(cur, table: str, columns: list[str], df: pd.DataFrame) -> None:
    df = df.reindex(columns=columns)
    csv_buf = io.StringIO()
    df.to_csv(csv_buf, index=False, header=False, na_rep="")
    cols_sql = ", ".join(f'"{c}"' for c in columns)  # quote: spaces/non-ASCII
    create_cols = ", ".join(f'"{c}" TEXT' for c in columns)
    cur.execute(f"CREATE TABLE IF NOT EXISTS {table} ({create_cols})")
    cur.execute(f"TRUNCATE {table}")
    with cur.copy(f"COPY {table} ({cols_sql}) FROM STDIN WITH (FORMAT csv, NULL '')") as cp:
        cp.write(csv_buf.getvalue())
    print(f"loaded {table} ({len(df)} rows)")


def main() -> None:
    base = Path(os.environ["SP_LOCAL_DIR"])
    files = discover_files(base)
    with psycopg.connect(os.environ["PG_DSN"]) as conn, conn.cursor() as cur:
        for filepath, year, month in files:
            for sheet_name, cfg in SHEET_CONFIGS.items():
                raw = pd.read_excel(filepath, sheet_name=sheet_name, **cfg["read_kwargs"])
                df = cfg["transform"](raw)
                table = f"data_{year}_{month}_{cfg['suffix']}"
                load_sheet(cur, table, cfg["columns"], df)
    print("done")


if __name__ == "__main__":
    main()
