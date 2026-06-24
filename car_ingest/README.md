# Case Assessment Report to SQL Server

This package watches a folder of `.docx` reports, extracts the report fields, and upserts them into Microsoft SQL Server.

## Folder watched
Default:

`C:\Users\QKU4847\OneDrive - HCA Healthcare\CARProject`

Override with `CAR_WATCH_FOLDER`.

## Install

```bash
pip install -r requirements.txt
```

## Configure environment

Create environment variables:

```bash
set SQL_SERVER=YOURSERVER
set SQL_DATABASE=YOURDB
set SQL_USERNAME=YOURUSER
set SQL_PASSWORD=YOURPASSWORD
set SQL_DRIVER=ODBC Driver 17 for SQL Server
set SQL_TRUSTED_CONNECTION=false
```

For Windows integrated auth:

```bash
set SQL_TRUSTED_CONNECTION=true
```

Optional table names:

```bash
set SQL_SCHEMA=dbo
set SQL_MAIN_TABLE=CaseAssessmentReport
set SQL_PROVIDER_TABLE=CaseAssessmentReportProvider
```

## Run

Create tables:

```bash
python -m car_ingest setup
```

Ingest all `.docx` files once:

```bash
python -m car_ingest ingest
```

Watch the folder continuously:

```bash
python -m car_ingest watch --interval 30
```

Ingest one file:

```bash
python -m car_ingest ingest-one "C:\path\to\file.docx"
```

## How updates work

Each file is keyed by full file path. If the file changes, the row is updated and the provider rows are replaced.

## Notes

The narrative sections are stored as NVARCHAR(MAX). The repeated provider section is also captured in a child table, with a raw-text fallback so later tuning is straightforward.
