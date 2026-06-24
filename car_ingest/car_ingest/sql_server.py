from __future__ import annotations

import json
from dataclasses import asdict
from typing import Any

from .config import AppConfig
from .docx_parser import ParsedReport


MAIN_COLUMNS = [
    "source_file_name",
    "source_file_path",
    "source_file_hash",
    "source_last_modified_utc",
    "claim_number",
    "claim_manager",
    "report_date",
    "insured",
    "div_hos_notice",
    "date_of_loss",
    "date_of_report",
    "date_of_est",
    "date_of_suit",
    "plaintiff_name",
    "hca_named",
    "jurisdiction",
    "primary_insurance_limit",
    "indemnity_reserve",
    "lae_paid",
    "defense_counsel",
    "defense_firm",
    "plaintiff_counsel",
    "plaintiff_firm",
    "authority_required",
    "demand_offer",
    "trial_date",
    "mediation_date",
    "chance_dv",
    "verdict_value",
    "settlement_value",
    "executive_summary",
    "resolution_strategy",
    "facts",
    "injury",
    "plaintiff_section",
    "damages",
    "allegations",
    "defenses",
    "peer_review_remediation",
    "internal_review",
    "experts",
    "defense_section",
    "raw_json",
]


CREATE_TABLE_SQL = """
IF OBJECT_ID('{schema}.{main_table}', 'U') IS NULL
BEGIN
    CREATE TABLE {schema}.{main_table} (
        report_id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        source_file_name NVARCHAR(260) NOT NULL,
        source_file_path NVARCHAR(1000) NOT NULL UNIQUE,
        source_file_hash CHAR(64) NOT NULL,
        source_last_modified_utc DATETIME2(0) NULL,
        claim_number NVARCHAR(255) NULL,
        claim_manager NVARCHAR(255) NULL,
        report_date NVARCHAR(255) NULL,
        insured NVARCHAR(500) NULL,
        div_hos_notice NVARCHAR(255) NULL,
        date_of_loss NVARCHAR(255) NULL,
        date_of_report NVARCHAR(255) NULL,
        date_of_est NVARCHAR(255) NULL,
        date_of_suit NVARCHAR(255) NULL,
        plaintiff_name NVARCHAR(500) NULL,
        hca_named NVARCHAR(255) NULL,
        jurisdiction NVARCHAR(255) NULL,
        primary_insurance_limit NVARCHAR(255) NULL,
        indemnity_reserve NVARCHAR(255) NULL,
        lae_paid NVARCHAR(255) NULL,
        defense_counsel NVARCHAR(500) NULL,
        defense_firm NVARCHAR(500) NULL,
        plaintiff_counsel NVARCHAR(500) NULL,
        plaintiff_firm NVARCHAR(500) NULL,
        authority_required NVARCHAR(255) NULL,
        demand_offer NVARCHAR(255) NULL,
        trial_date NVARCHAR(255) NULL,
        mediation_date NVARCHAR(255) NULL,
        chance_dv NVARCHAR(255) NULL,
        verdict_value NVARCHAR(255) NULL,
        settlement_value NVARCHAR(255) NULL,
        executive_summary NVARCHAR(MAX) NULL,
        resolution_strategy NVARCHAR(MAX) NULL,
        facts NVARCHAR(MAX) NULL,
        injury NVARCHAR(MAX) NULL,
        plaintiff_section NVARCHAR(MAX) NULL,
        damages NVARCHAR(MAX) NULL,
        allegations NVARCHAR(MAX) NULL,
        defenses NVARCHAR(MAX) NULL,
        peer_review_remediation NVARCHAR(MAX) NULL,
        internal_review NVARCHAR(MAX) NULL,
        experts NVARCHAR(MAX) NULL,
        defense_section NVARCHAR(MAX) NULL,
        raw_json NVARCHAR(MAX) NULL,
        ingested_at_utc DATETIME2(0) NOT NULL CONSTRAINT DF_{main_table}_ingested_at DEFAULT SYSUTCDATETIME(),
        last_seen_at_utc DATETIME2(0) NOT NULL CONSTRAINT DF_{main_table}_last_seen_at DEFAULT SYSUTCDATETIME()
    );
END;

IF OBJECT_ID('{schema}.{provider_table}', 'U') IS NULL
BEGIN
    CREATE TABLE {schema}.{provider_table} (
        provider_id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        report_id INT NOT NULL,
        row_index INT NOT NULL,
        involved_provider_specialty NVARCHAR(500) NULL,
        relationship_to_facility NVARCHAR(500) NULL,
        carrier_limits NVARCHAR(500) NULL,
        prior_hci_claim_involvement NVARCHAR(500) NULL,
        hci_insured_status NVARCHAR(500) NULL,
        raw_text NVARCHAR(MAX) NULL,
        CONSTRAINT FK_{provider_table}_report FOREIGN KEY (report_id)
            REFERENCES {schema}.{main_table}(report_id)
            ON DELETE CASCADE
    );
END;
"""



def _sql_ident(name: str) -> str:
    return f"[{name.replace(']', ']]')}]"



def build_connection_string(cfg: AppConfig) -> str:
    parts = [f"DRIVER={{{cfg.sql_driver}}}", f"SERVER={cfg.sql_server}", f"DATABASE={cfg.sql_database}"]
    if cfg.trusted_connection:
        parts.append("Trusted_Connection=yes")
    else:
        if not cfg.sql_username or not cfg.sql_password:
            raise ValueError("SQL_USERNAME and SQL_PASSWORD are required unless SQL_TRUSTED_CONNECTION=true")
        parts.append(f"UID={cfg.sql_username}")
        parts.append(f"PWD={cfg.sql_password}")
    parts.append("TrustServerCertificate=yes")
    return ";".join(parts)



def ensure_schema_objects(conn, cfg: AppConfig) -> None:
    schema = _sql_ident(cfg.sql_schema)
    main_table = _sql_ident(cfg.main_table)
    provider_table = _sql_ident(cfg.provider_table)
    sql = CREATE_TABLE_SQL.format(schema=schema, main_table=main_table, provider_table=provider_table)
    cur = conn.cursor()
    cur.execute(sql)
    conn.commit()



def upsert_report(conn, cfg: AppConfig, report: ParsedReport) -> None:
    schema = _sql_ident(cfg.sql_schema)
    main_table = _sql_ident(cfg.main_table)
    provider_table = _sql_ident(cfg.provider_table)

    fields = report.fields.copy()
    raw_json = json.dumps(
        {
            "fields": report.fields,
            "providers": [asdict(p) for p in report.providers],
            "file_name": report.file_name,
            "file_path": report.file_path,
            "file_hash": report.file_hash,
            "last_modified_utc": report.last_modified_utc,
        },
        ensure_ascii=False,
        default=str,
    )

    field_values = {
        "source_file_name": report.file_name,
        "source_file_path": report.file_path,
        "source_file_hash": report.file_hash,
        "source_last_modified_utc": report.last_modified_utc,
        "claim_number": fields.get("claim_number"),
        "claim_manager": fields.get("claim_manager"),
        "report_date": fields.get("report_date"),
        "insured": fields.get("insured"),
        "div_hos_notice": fields.get("div_hos_notice"),
        "date_of_loss": fields.get("date_of_loss"),
        "date_of_report": fields.get("date_of_report"),
        "date_of_est": fields.get("date_of_est"),
        "date_of_suit": fields.get("date_of_suit"),
        "plaintiff_name": fields.get("plaintiff_name"),
        "hca_named": fields.get("hca_named"),
        "jurisdiction": fields.get("jurisdiction"),
        "primary_insurance_limit": fields.get("primary_insurance_limit"),
        "indemnity_reserve": fields.get("indemnity_reserve"),
        "lae_paid": fields.get("lae_paid"),
        "defense_counsel": fields.get("defense_counsel"),
        "defense_firm": fields.get("defense_firm"),
        "plaintiff_counsel": fields.get("plaintiff_counsel"),
        "plaintiff_firm": fields.get("plaintiff_firm"),
        "authority_required": fields.get("authority_required"),
        "demand_offer": fields.get("demand_offer"),
        "trial_date": fields.get("trial_date"),
        "mediation_date": fields.get("mediation_date"),
        "chance_dv": fields.get("chance_dv"),
        "verdict_value": fields.get("verdict_value"),
        "settlement_value": fields.get("settlement_value"),
        "executive_summary": fields.get("executive_summary"),
        "resolution_strategy": fields.get("resolution_strategy"),
        "facts": fields.get("facts"),
        "injury": fields.get("injury"),
        "plaintiff_section": fields.get("plaintiff_section"),
        "damages": fields.get("damages"),
        "allegations": fields.get("allegations"),
        "defenses": fields.get("defenses"),
        "peer_review_remediation": fields.get("peer_review_remediation"),
        "internal_review": fields.get("internal_review"),
        "experts": fields.get("experts"),
        "defense_section": fields.get("defense_section"),
        "raw_json": raw_json,
    }

    cols = list(field_values.keys())
    params = [field_values[c] for c in cols]
    placeholders = ", ".join(["?"] * len(cols))
    col_sql = ", ".join(_sql_ident(c) for c in cols)

    cur = conn.cursor()
    cur.execute(f"SELECT report_id, source_file_hash FROM {schema}.{main_table} WHERE source_file_path = ?", report.file_path)
    row = cur.fetchone()
    if row:
        report_id, existing_hash = row
        if existing_hash == report.file_hash:
            cur.execute(
                f"UPDATE {schema}.{main_table} SET last_seen_at_utc = SYSUTCDATETIME() WHERE report_id = ?",
                report_id,
            )
            conn.commit()
            return
        set_sql = ", ".join(f"{_sql_ident(c)} = ?" for c in cols if c != "source_file_path")
        upd_params = [field_values[c] for c in cols if c != "source_file_path"] + [report_id]
        cur.execute(f"UPDATE {schema}.{main_table} SET {set_sql}, last_seen_at_utc = SYSUTCDATETIME() WHERE report_id = ?", upd_params)
    else:
        cur.execute(f"INSERT INTO {schema}.{main_table} ({col_sql}) VALUES ({placeholders})", params)
        report_id = cur.execute("SELECT SCOPE_IDENTITY()").fetchone()[0]

    cur.execute(f"DELETE FROM {schema}.{provider_table} WHERE report_id = ?", report_id)
    for p in report.providers:
        cur.execute(
            f"INSERT INTO {schema}.{provider_table} (report_id, row_index, involved_provider_specialty, relationship_to_facility, carrier_limits, prior_hci_claim_involvement, hci_insured_status, raw_text) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            report_id,
            p.row_index,
            p.involved_provider_specialty,
            p.relationship_to_facility,
            p.carrier_limits,
            p.prior_hci_claim_involvement,
            p.hci_insured_status,
            p.raw_text,
        )
    conn.commit()
