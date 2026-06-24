from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class AppConfig:
    watch_folder: Path
    sql_server: str
    sql_database: str
    sql_username: str | None
    sql_password: str | None
    sql_driver: str
    trusted_connection: bool
    sql_schema: str
    main_table: str
    provider_table: str
    poll_interval_seconds: int



def load_config() -> AppConfig:
    watch_folder = Path(os.getenv("CAR_WATCH_FOLDER", r"C:\Users\QKU4847\OneDrive - HCA Healthcare\CARProject"))
    return AppConfig(
        watch_folder=watch_folder,
        sql_server=os.getenv("SQL_SERVER", "localhost"),
        sql_database=os.getenv("SQL_DATABASE", "CAR"),
        sql_username=os.getenv("SQL_USERNAME") or None,
        sql_password=os.getenv("SQL_PASSWORD") or None,
        sql_driver=os.getenv("SQL_DRIVER", "ODBC Driver 17 for SQL Server"),
        trusted_connection=os.getenv("SQL_TRUSTED_CONNECTION", "false").strip().lower() in {"1", "true", "yes", "y"},
        sql_schema=os.getenv("SQL_SCHEMA", "dbo"),
        main_table=os.getenv("SQL_MAIN_TABLE", "CaseAssessmentReport"),
        provider_table=os.getenv("SQL_PROVIDER_TABLE", "CaseAssessmentReportProvider"),
        poll_interval_seconds=int(os.getenv("POLL_INTERVAL_SECONDS", "30")),
    )
