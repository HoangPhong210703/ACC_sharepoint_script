-- Reference schema for the 3 consolidated tables created by ingest.py.
-- ingest.py creates these tables on demand (CREATE TABLE IF NOT EXISTS) and
-- ingests each Excel file as a snapshot, identified by "Source.Name".
-- Re-running a file replaces only that file's rows (DELETE WHERE Source.Name = ...).
-- This file is documentation; you do not need to run it before ingest.py.
-- Quoted identifiers are required because column names contain spaces,
-- non-ASCII (Vietnamese) characters, dots, and slashes.

-- =========================================================================
-- data_exp — expense transactions, one row per ledger line per month snapshot
-- =========================================================================

CREATE TABLE IF NOT EXISTS data_exp (
    "Source.Name"        text,
    "Ngày hạch toán"     text,
    "Ngày chứng từ"      text,
    "Số chứng từ"        text,
    "Diễn giải"          text,
    "Tài khoản"          text,
    "TK đối ứng"         text,
    "Phát sinh Nợ (ROW)" numeric,
    "Phát sinh Có (ROW)" numeric,
    "Dư Nợ"              numeric,
    "Dư Có"              numeric,
    "Phát sinh Nợ"       numeric,
    "Phát sinh Có"       numeric,
    "BU"                 text,
    "Cost center"        text,
    "Cost items"         text,
    "Report Year"        int,
    "Report Month"       int,
    "Report Month Name"  text,
    "Software/JSV"       text
);
CREATE INDEX IF NOT EXISTS data_exp_source_name_idx ON data_exp ("Source.Name");

-- =========================================================================
-- data_rev — revenue by BU x month, one row per BU-month per snapshot
-- =========================================================================

CREATE TABLE IF NOT EXISTS data_rev (
    "Source.Name"         text,
    "BU"                  text,
    "Month Name"          text,
    "Month Number"        int,
    "Revenue"             bigint,
    "Unit"                text,
    "Report Year"         int,
    "Report Month Number" int,
    "Report Month Name"   text,
    "Is_month_active"     int
);
CREATE INDEX IF NOT EXISTS data_rev_source_name_idx ON data_rev ("Source.Name");

-- =========================================================================
-- data_hc — headcount (man-months) by cost center x BU per snapshot
-- =========================================================================

CREATE TABLE IF NOT EXISTS data_hc (
    "Source.Name"       text,
    "Cost center"       text,
    "BU"                text,
    "MM"                numeric,
    "Report Year"       int,
    "Report Month"      int,
    "Report Month Name" text,
    "HC Group"          text
);
CREATE INDEX IF NOT EXISTS data_hc_source_name_idx ON data_hc ("Source.Name");
