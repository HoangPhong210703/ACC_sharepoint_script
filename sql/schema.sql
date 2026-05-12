-- All columns are text because ingest.py reads everything as strings.
-- Tighten to date/int/numeric later once the data quality is known.
-- Quoted identifiers are required because column names contain spaces and
-- non-ASCII (Vietnamese) characters.

-- =========================================================================
-- Data exp tables (one per file)
-- =========================================================================

DROP TABLE IF EXISTS data_2026_jan_exp;
DROP TABLE IF EXISTS data_2026_feb_exp;
DROP TABLE IF EXISTS data_2026_mar_exp;

CREATE TABLE data_2026_jan_exp (
    "Ngày hạch toán"     text,
    "Report Day"         text,
    "Report Month"       text,
    "Report Year"        text,
    "Ngày chứng từ"      text,
    "Số chứng từ"        text,
    "Diễn giải"          text,
    "Tài khoản"          text,
    "TK đối ứng"         text,
    "Phát sinh Nợ (ROW)" text,
    "Phát sinh Có (ROW)" text,
    "Dư Nợ"              text,
    "Dư Có"              text,
    "Phát sinh Nợ"       text,
    "Phát sinh Có"       text,
    "BU"                 text,
    "Cost center"        text,
    "Cost items"         text
);
CREATE TABLE data_2026_feb_exp (LIKE data_2026_jan_exp INCLUDING ALL);
CREATE TABLE data_2026_mar_exp (LIKE data_2026_jan_exp INCLUDING ALL);

-- =========================================================================
-- Data Rev tables (one per file)
-- =========================================================================

DROP TABLE IF EXISTS data_2026_jan_rev;
DROP TABLE IF EXISTS data_2026_feb_rev;
DROP TABLE IF EXISTS data_2026_mar_rev;

CREATE TABLE data_2026_jan_rev (
    "BU"           text,
    "Month name"   text,
    "Revenue"      text,
    "Unit"         text,
    "Month number" text
);
CREATE TABLE data_2026_feb_rev (LIKE data_2026_jan_rev INCLUDING ALL);
CREATE TABLE data_2026_mar_rev (LIKE data_2026_jan_rev INCLUDING ALL);

-- =========================================================================
-- Data HC tables (one per file)
-- =========================================================================

DROP TABLE IF EXISTS data_2026_jan_hc;
DROP TABLE IF EXISTS data_2026_feb_hc;
DROP TABLE IF EXISTS data_2026_mar_hc;

CREATE TABLE data_2026_jan_hc (
    "Cost center" text,
    "BU"          text,
    "MM"          text
);
CREATE TABLE data_2026_feb_hc (LIKE data_2026_jan_hc INCLUDING ALL);
CREATE TABLE data_2026_mar_hc (LIKE data_2026_jan_hc INCLUDING ALL);
