"""CAR report ingestion for Word .docx files into Microsoft SQL Server.

What this script does:
- Scans a folder (recursively) for new or modified .docx files.
- Reads the Word XML directly.
- Extracts the Case Assessment Report fields.
- Upserts one main row per report into SQL Server.
- Replaces child provider rows on each refresh so revised reports stay in sync.

Dependencies:
- pyodbc
- Microsoft ODBC Driver 17 or 18 for SQL Server

Usage examples:
    python car_report_ingest.py --once \
        --folder "C:\\Users\\QKU4847\\OneDrive - HCA Healthcare\\CARProject" \
        --conn "Driver={ODBC Driver 18 for SQL Server};Server=;Database=;UID=USER;PWD=PASS;TrustServerCertificate=yes"

    python car_report_ingest.py --watch \
        --folder "C:\\Users\\QKU4847\\OneDrive - HCA Healthcare\\CARProject" \
        --conn "Driver={ODBC Driver 18 for SQL Server};Server=;Database=;UID=USER;PWD=PASS;TrustServerCertificate=yes"

If no connection string is supplied, the script runs in dry-run mode and prints JSON.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import time
import traceback
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Tuple

try:
    import pyodbc
except ImportError:  # pragma: no cover
    pyodbc = None

from xml.etree import ElementTree as ET


WATCH_FOLDER_DEFAULT = r"C:\Users\QKU4847\OneDrive - HCA Healthcare\CARProject"
SQL_TABLE_DEFAULT = "dbo.CaseAssessmentReport"
SQL_PROVIDER_TABLE_DEFAULT = "dbo.CaseAssessmentReportProvider"
POLL_SECONDS_DEFAULT = 15

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
PLACEHOLDERS = {
    "",
    "Click to enter text.",
    "Click or tap here to enter text.",
    "Click or tap here to enter text",
}


@dataclass
class ProviderRow:
    provider_order: int
    provider_specialty: str = ""
    relationship_to_facility: str = ""
    carrier_limits: str = ""
    prior_hci_claim_involvement: str = ""
    hci_insured_status: str = ""


@dataclass
class CaseAssessmentRecord:
    # Identification / lineage
    report_id: Optional[int] = None
    report_key: str = ""
    claim_number: str = ""
    source_file_name: str = ""
    source_file_path: str = ""
    document_hash: str = ""
    file_modified_utc: str = ""
    last_processed_utc: str = ""

    # Top of report
    claim_manager: str = ""
    report_date: str = ""
    insured: str = ""
    div_hos_notice: str = ""
    date_of_loss: str = ""
    date_of_report: str = ""
    date_of_est: str = ""
    date_of_suit: str = ""
    plaintiff: str = ""
    hca_named: str = ""
    jurisdiction: str = ""
    prim_ins_limit: str = ""
    indemnity_reserve: str = ""
    lae_paid: str = ""
    defense_counsel: str = ""
    defense_firm: str = ""
    plaintiff_counsel: str = ""
    plaintiff_firm: str = ""

    # Co-defendants / providers
    provider_1_specialty: str = ""
    provider_1_relationship: str = ""
    provider_1_carrier_limits: str = ""
    provider_1_prior_hci_claim_involvement: str = ""
    provider_1_hci_insured_status: str = ""
    provider_2_specialty: str = ""
    provider_2_relationship: str = ""
    provider_2_carrier_limits: str = ""
    provider_2_prior_hci_claim_involvement: str = ""
    provider_2_hci_insured_status: str = ""
    provider_3_specialty: str = ""
    provider_3_relationship: str = ""
    provider_3_carrier_limits: str = ""
    provider_3_prior_hci_claim_involvement: str = ""
    provider_3_hci_insured_status: str = ""

    # Strategy / case economics
    authority_req: str = ""
    demand_offer: str = ""
    trial_date: str = ""
    mediation_date: str = ""
    chance_dv: str = ""
    verdict_value: str = ""
    settlement_value: str = ""

    # Narrative / section text
    executive_summary: str = ""
    resolution_strategy: str = ""
    facts_text: str = ""
    injury: str = ""
    plaintiff_text: str = ""
    damages: str = ""
    damages_summary_text: str = ""
    allegations_text: str = ""
    defenses_text: str = ""
    peer_review_remediation_text: str = ""
    internal_review_text: str = ""
    experts_text: str = ""
    experts_defense_text: str = ""

    def as_sql_params(self) -> Dict[str, str]:
        return asdict(self)


# ---------------------------------------------------------------------------
# Text normalization
# ---------------------------------------------------------------------------

def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def normalize_for_match(text: str) -> str:
    """Collapse all whitespace to a single line."""
    return " ".join((text or "").replace("\r", "\n").replace("\n", " ").split()).strip()


def normalize_preserve_newlines(text: str) -> str:
    """Normalize each line but preserve line breaks."""
    if text is None:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines: List[str] = []
    for raw_line in text.split("\n"):
        line = " ".join(raw_line.split()).strip()
        if not line:
            continue
        if line in PLACEHOLDERS:
            continue
        lines.append(line)
    return "\n".join(lines).strip()


def clean_scalar(text: str) -> str:
    return normalize_for_match(text)


def strip_label_prefix(text: str, label: str) -> str:
    value = normalize_for_match(text)
    label_norm = normalize_for_match(label)
    if value.upper().startswith(label_norm.upper()):
        value = value[len(label_norm):].strip()
    return normalize_preserve_newlines(value)


# ---------------------------------------------------------------------------
# DOCX XML helpers
# ---------------------------------------------------------------------------

def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def parse_docx_xml(docx_path: Path) -> ET.Element:
    with zipfile.ZipFile(docx_path) as zf:
        xml = zf.read("word/document.xml")
    return ET.fromstring(xml)


def get_body_table(root: ET.Element) -> ET.Element:
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("word/document.xml does not contain w:body")
    for child in body:
        if child.tag == f"{{{W_NS}}}tbl":
            return child
    raise ValueError("Could not find report table in document body")


def row_children(row: ET.Element) -> List[ET.Element]:
    return [child for child in list(row) if child.tag in {f"{{{W_NS}}}tc", f"{{{W_NS}}}sdt", f"{{{W_NS}}}trPr"}]


def _element_lines(elem: ET.Element) -> List[str]:
    paragraphs = []
    for p in elem.findall(".//w:p", NS):
        txt = normalize_for_match("".join(p.itertext()))
        if txt and txt not in PLACEHOLDERS:
            paragraphs.append(txt)
    if not paragraphs:
        txt = normalize_for_match("".join(elem.itertext()))
        if txt and txt not in PLACEHOLDERS:
            paragraphs.append(txt)
    return paragraphs


def cell_text(cell: ET.Element) -> str:
    return normalize_preserve_newlines("\n".join(_element_lines(cell)))


def sdt_text(sdt: ET.Element) -> str:
    return normalize_preserve_newlines("\n".join(_element_lines(sdt)))


def split_labels_values(row: ET.Element) -> List[str]:
    values: List[str] = []
    for child in row_children(row):
        if child.tag == f"{{{W_NS}}}trPr":
            continue
        if child.tag == f"{{{W_NS}}}tc":
            values.append(cell_text(child))
        elif child.tag == f"{{{W_NS}}}sdt":
            values.append(sdt_text(child))
    return values


def first_descendant_sdt_text(row: ET.Element) -> str:
    sdt = row.find(".//w:sdt", NS)
    return sdt_text(sdt) if sdt is not None else ""


def row_text(row: ET.Element) -> str:
    parts: List[str] = []
    for child in row_children(row):
        if child.tag == f"{{{W_NS}}}trPr":
            continue
        if child.tag == f"{{{W_NS}}}tc":
            txt = cell_text(child)
        elif child.tag == f"{{{W_NS}}}sdt":
            txt = sdt_text(child)
        else:
            txt = ""
        if txt:
            parts.append(txt)
    return normalize_preserve_newlines("\n".join(parts))


def row_body_text(row: ET.Element) -> str:
    """Prefer SDT text when present; otherwise fall back to the whole row."""
    sdt = first_descendant_sdt_text(row)
    return sdt if sdt else row_text(row)


def row_matches_heading(row: ET.Element, heading: str) -> bool:
    txt = normalize_for_match(row_text(row)).upper()
    return txt.startswith(normalize_for_match(heading).upper())


def find_heading_row(rows: Sequence[ET.Element], heading: str, start_at: int = 0) -> Optional[int]:
    for idx in range(start_at, len(rows)):
        if row_matches_heading(rows[idx], heading):
            return idx
    return None


def collect_rows_text(rows: Sequence[ET.Element], start_idx: int, end_idx: int, use_row_text: bool = False) -> str:
    pieces: List[str] = []
    for idx in range(start_idx, end_idx):
        txt = row_text(rows[idx]) if use_row_text else row_body_text(rows[idx])
        if txt:
            pieces.append(txt)
    return normalize_preserve_newlines("\n".join(pieces))




def row_looks_like_labels_only(values: Sequence[str]) -> bool:
    cleaned = [clean_scalar(v) for v in values if clean_scalar(v)]
    if not cleaned:
        return False
    return all(v.endswith(":") for v in cleaned)


def label_value_row_values(rows: Sequence[ET.Element], row_idx: int) -> List[str]:
    """Return the value row for a label/value block.

    Some report sections use a header row followed by a value row. If the
    current row is labels only and the next row contains actual values, return
    the next row. Otherwise return the current row.
    """
    current = split_labels_values(rows[row_idx]) if row_idx < len(rows) else []
    if row_idx + 1 < len(rows):
        nxt = split_labels_values(rows[row_idx + 1])
    else:
        nxt = []

    if row_looks_like_labels_only(current) and nxt and not row_looks_like_labels_only(nxt):
        return nxt
    return current

def section_from_heading(
    rows: Sequence[ET.Element],
    heading: str,
    next_heading: Optional[str] = None,
    include_same_row_remainder: bool = False,
) -> str:
    idx = find_heading_row(rows, heading)
    if idx is None:
        return ""

    current = row_body_text(rows[idx])
    remainder = strip_label_prefix(current, heading)
    if include_same_row_remainder and remainder:
        return remainder

    if next_heading is not None:
        next_idx = find_heading_row(rows, next_heading, start_at=idx + 1)
        if next_idx is None:
            next_idx = len(rows)
        start_idx = idx + 1
        end_idx = next_idx
    else:
        start_idx = idx + 1
        end_idx = len(rows)

    if start_idx >= end_idx:
        return remainder if remainder else ""

    body = collect_rows_text(rows, start_idx, end_idx)
    if remainder and body:
        return normalize_preserve_newlines(remainder + "\n" + body)
    return body or remainder


# ---------------------------------------------------------------------------
# Extraction logic
# ---------------------------------------------------------------------------

def extract_case_assessment(docx_path: Path) -> Tuple[CaseAssessmentRecord, List[ProviderRow]]:
    root = parse_docx_xml(docx_path)
    table = get_body_table(root)
    rows = table.findall("w:tr", NS)
    record = CaseAssessmentRecord()

    # Top rows
    row0 = split_labels_values(rows[0]) if len(rows) > 0 else []
    if len(row0) >= 6:
        record.claim_number = row0[1]
        record.claim_manager = row0[3]
        record.report_date = row0[5]

    row2 = split_labels_values(rows[2]) if len(rows) > 2 else []
    if len(row2) >= 6:
        record.insured = row2[0]
        record.div_hos_notice = row2[1]
        record.date_of_loss = row2[2]
        record.date_of_report = row2[3]
        record.date_of_est = row2[4]
        record.date_of_suit = row2[5]

    row4 = split_labels_values(rows[4]) if len(rows) > 4 else []
    if len(row4) >= 6:
        record.plaintiff = row4[0]
        record.hca_named = row4[1]
        record.jurisdiction = row4[2]
        record.prim_ins_limit = row4[3]
        record.indemnity_reserve = row4[4]
        record.lae_paid = row4[5]

    row6 = split_labels_values(rows[6]) if len(rows) > 6 else []
    if len(row6) >= 4:
        record.defense_counsel = row6[0]
        record.defense_firm = row6[1]
        record.plaintiff_counsel = row6[2]
        record.plaintiff_firm = row6[3]

    # Provider rows (repeating section)
    providers: List[ProviderRow] = []
    for provider_order, row_idx in enumerate([9, 10, 11], start=1):
        if row_idx >= len(rows):
            break
        vals = split_labels_values(rows[row_idx])
        provider = ProviderRow(provider_order=provider_order)
        if len(vals) >= 1:
            provider.provider_specialty = vals[0]
        if len(vals) >= 2:
            provider.relationship_to_facility = vals[1]
        if len(vals) >= 3:
            provider.carrier_limits = vals[2]
        if len(vals) >= 4:
            provider.prior_hci_claim_involvement = vals[3]
        provider.hci_insured_status = provider.prior_hci_claim_involvement
        providers.append(provider)

    # Strategy / economics
    # Strategy / economics
    # These sections are stored as two stacked rows in the template:
    #   Authority row:  Authority Req / Demand-Offer / Trial Date / Mediation Date
    #   Chance row:     % Chance DV / Verdict Value / Settlement Value
    auth_idx = find_heading_row(rows, "Authority Req:")
    if auth_idx is not None:
        auth_vals = label_value_row_values(rows, auth_idx)
        if len(auth_vals) >= 4:
            record.authority_req = auth_vals[0]
            record.demand_offer = auth_vals[1]
            record.trial_date = auth_vals[2]
            record.mediation_date = auth_vals[3]

    chance_idx = find_heading_row(rows, "% Chance DV:")
    if chance_idx is not None:
        chance_vals = label_value_row_values(rows, chance_idx)
        if len(chance_vals) >= 3:
            record.chance_dv = chance_vals[0]
            record.verdict_value = chance_vals[1]
            record.settlement_value = chance_vals[2]

    # Narrative sections
    record.executive_summary = section_from_heading(rows, "Executive Summary:", next_heading="Resolution Strategy:", include_same_row_remainder=True)
    record.resolution_strategy = section_from_heading(rows, "Resolution Strategy:", next_heading="FACTS:", include_same_row_remainder=True)
    record.facts_text = section_from_heading(rows, "FACTS:", next_heading="INJURY:")
    record.injury = section_from_heading(rows, "INJURY:", next_heading="PLAINTIFF:", include_same_row_remainder=True)
    record.plaintiff_text = section_from_heading(rows, "PLAINTIFF:", next_heading="DAMAGES:", include_same_row_remainder=True)
    record.damages = section_from_heading(rows, "DAMAGES:", next_heading="ALLEGATIONS:", include_same_row_remainder=True)
    damage_summary_start = find_heading_row(rows, "A summary of the Plaintiff and Defense cases on damages is provided below:")
    allegations_idx = find_heading_row(rows, "ALLEGATIONS:")
    if damage_summary_start is not None:
        damage_summary_end = allegations_idx if allegations_idx is not None else len(rows)
        record.damages_summary_text = collect_rows_text(rows, damage_summary_start, damage_summary_end, use_row_text=True)
    else:
        record.damages_summary_text = ""
    record.allegations_text = section_from_heading(rows, "ALLEGATIONS:", next_heading="DEFENSES:", include_same_row_remainder=True)
    record.defenses_text = section_from_heading(rows, "DEFENSES:", next_heading="PEER REVIEW/REMEDIATION:", include_same_row_remainder=True)
    record.peer_review_remediation_text = section_from_heading(rows, "PEER REVIEW/REMEDIATION:", next_heading="INTERNAL REVIEW:", include_same_row_remainder=True)
    record.internal_review_text = section_from_heading(rows, "INTERNAL REVIEW:", next_heading="EXPERTS:", include_same_row_remainder=True)

    # Experts are all in one long control in this template.
    experts_blob = section_from_heading(rows, "EXPERTS:", next_heading=None, include_same_row_remainder=False)
    experts_blob = normalize_preserve_newlines(experts_blob)
    if experts_blob:
        # Split into plaintiff vs defense expert narratives if possible.
        defense_match = re.search(r"(?i)\bDefense:\s*", experts_blob)
        if defense_match:
            plaintiff_blob = experts_blob[: defense_match.start()].strip()
            defense_blob = experts_blob[defense_match.end():].strip()
            record.experts_text = normalize_preserve_newlines(strip_label_prefix(plaintiff_blob, "Plaintiff:"))
            record.experts_defense_text = normalize_preserve_newlines(defense_blob)
        else:
            record.experts_text = normalize_preserve_newlines(strip_label_prefix(experts_blob, "Plaintiff:"))
            record.experts_defense_text = ""

    # Final cleanup with newline preservation.
    for key, value in record.as_sql_params().items():
        if isinstance(value, str):
            setattr(record, key, normalize_preserve_newlines(value))

    return record, providers


# ---------------------------------------------------------------------------
# SQL helpers
# ---------------------------------------------------------------------------

def choose_report_key(record: CaseAssessmentRecord) -> str:
    claim = clean_scalar(record.claim_number)
    if claim:
        return claim
    if record.source_file_name:
        return record.source_file_name.lower()
    return record.document_hash


def sql_safe_column(name: str) -> str:
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
        raise ValueError(f"Unsafe SQL identifier: {name}")
    return name


MAIN_COLUMNS = [
    "report_key",
    "claim_number",
    "source_file_name",
    "source_file_path",
    "document_hash",
    "file_modified_utc",
    "last_processed_utc",
    "claim_manager",
    "report_date",
    "insured",
    "div_hos_notice",
    "date_of_loss",
    "date_of_report",
    "date_of_est",
    "date_of_suit",
    "plaintiff",
    "hca_named",
    "jurisdiction",
    "prim_ins_limit",
    "indemnity_reserve",
    "lae_paid",
    "defense_counsel",
    "defense_firm",
    "plaintiff_counsel",
    "plaintiff_firm",
    "authority_req",
    "demand_offer",
    "trial_date",
    "mediation_date",
    "chance_dv",
    "verdict_value",
    "settlement_value",
    "executive_summary",
    "resolution_strategy",
    "facts_text",
    "injury",
    "plaintiff_text",
    "damages",
    "damages_summary_text",
    "allegations_text",
    "defenses_text",
    "peer_review_remediation_text",
    "internal_review_text",
    "experts_text",
    "experts_defense_text",
]


def get_existing_report_id(conn, main_table: str, record: CaseAssessmentRecord) -> Optional[int]:
    cursor = conn.cursor()
    if record.claim_number:
        sql = f"SELECT TOP 1 report_id FROM {main_table} WHERE claim_number = ? ORDER BY report_id DESC"
        row = cursor.execute(sql, record.claim_number).fetchone()
        if row:
            return int(row[0])
    if record.source_file_path:
        sql = f"SELECT TOP 1 report_id FROM {main_table} WHERE source_file_path = ? ORDER BY report_id DESC"
        row = cursor.execute(sql, record.source_file_path).fetchone()
        if row:
            return int(row[0])
    return None


def upsert_report(conn, main_table: str, provider_table: str, record: CaseAssessmentRecord, providers: List[ProviderRow]) -> int:
    record.report_key = choose_report_key(record)
    params = record.as_sql_params()
    existing_id = get_existing_report_id(conn, main_table, record)
    cursor = conn.cursor()

    if existing_id is None:
        cols = [sql_safe_column(c) for c in MAIN_COLUMNS]
        placeholders = ", ".join(["?"] * len(cols))
        sql = f"INSERT INTO {main_table} ({', '.join(cols)}) OUTPUT INSERTED.report_id VALUES ({placeholders})"
        values = [params.get(col) for col in cols]
        row = cursor.execute(sql, values).fetchone()
        report_id = int(row[0])
    else:
        assignments = ", ".join([f"{sql_safe_column(col)} = ?" for col in MAIN_COLUMNS if col != "report_key"])
        sql = f"UPDATE {main_table} SET report_key = ?, {assignments} WHERE report_id = ?"
        values = [params.get("report_key")] + [params.get(col) for col in MAIN_COLUMNS if col != "report_key"] + [existing_id]
        cursor.execute(sql, values)
        report_id = existing_id

    # Refresh child rows so revised reports replace the previous provider list.
    cursor.execute(f"DELETE FROM {provider_table} WHERE report_id = ?", report_id)
    for p in providers:
        cursor.execute(
            f"""
            INSERT INTO {provider_table}
                (report_id, provider_order, provider_specialty, relationship_to_facility, carrier_limits,
                 prior_hci_claim_involvement, hci_insured_status)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            report_id,
            p.provider_order,
            p.provider_specialty,
            p.relationship_to_facility,
            p.carrier_limits,
            p.prior_hci_claim_involvement,
            p.hci_insured_status,
        )

    conn.commit()
    return report_id


# ---------------------------------------------------------------------------
# File scanning / runtime
# ---------------------------------------------------------------------------

def list_docx_files(folder: Path) -> List[Path]:
    return sorted(
        [p for p in folder.rglob("*.docx") if p.is_file() and not p.name.startswith("~$")],
        key=lambda p: str(p).lower(),
    )


def is_stable_file(path: Path, wait_seconds: float = 1.0) -> bool:
    try:
        size1 = path.stat().st_size
        time.sleep(wait_seconds)
        size2 = path.stat().st_size
        return size1 == size2
    except FileNotFoundError:
        return False


def print_dry_run(record: CaseAssessmentRecord, providers: List[ProviderRow]) -> None:
    payload = {
        "record": record.as_sql_params(),
        "providers": [asdict(p) for p in providers],
    }
    print(json.dumps(payload, indent=2, ensure_ascii=False))


def process_file(path: Path, args: argparse.Namespace, conn) -> None:
    if not path.exists() or not path.is_file():
        return
    if not is_stable_file(path, wait_seconds=0.5):
        return

    record, providers = extract_case_assessment(path)
    record.source_file_name = path.name
    record.source_file_path = str(path)
    record.document_hash = sha256_file(path)
    record.file_modified_utc = datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc).replace(microsecond=0).isoformat()
    record.last_processed_utc = utc_now_iso()
    record.report_key = choose_report_key(record)

    if conn is None:
        print(f"[DRY RUN] {path}")
        print_dry_run(record, providers)
        return

    report_id = upsert_report(conn, args.main_table, args.provider_table, record, providers)
    print(f"[SQL] Upserted report_id={report_id} from {path.name}")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Ingest Case Assessment Report .docx files into SQL Server")
    parser.add_argument("--folder", default=os.environ.get("CAR_FOLDER", WATCH_FOLDER_DEFAULT), help="Folder to scan")
    parser.add_argument("--file", default=os.environ.get("CAR_FILE", ""), help="Optional single .docx file to process")
    parser.add_argument("--conn", default=os.environ.get("SQLSERVER_CONN_STR", ""), help="SQL Server ODBC connection string")
    parser.add_argument("--main-table", default=os.environ.get("CAR_MAIN_TABLE", SQL_TABLE_DEFAULT), help="Main SQL table")
    parser.add_argument("--provider-table", default=os.environ.get("CAR_PROVIDER_TABLE", SQL_PROVIDER_TABLE_DEFAULT), help="Provider SQL table")
    parser.add_argument("--watch", action="store_true", help="Keep polling the folder for changes")
    parser.add_argument("--poll-seconds", type=int, default=POLL_SECONDS_DEFAULT, help="Polling interval when --watch is set")
    parser.add_argument("--once", action="store_true", help="Run one scan and exit")
    return parser


def target_paths(args: argparse.Namespace) -> List[Path]:
    if args.file:
        return [Path(args.file)]
    return list_docx_files(Path(args.folder))


def scan_and_process(args: argparse.Namespace, conn, seen_hashes: Dict[str, str]) -> None:
    for path in target_paths(args):
        try:
            current_hash = sha256_file(path)
        except Exception as exc:
            print(f"[WARN] Could not hash {path.name if path.name else path}: {exc}")
            continue

        cache_key = str(path).lower()
        if seen_hashes.get(cache_key) == current_hash:
            continue

        try:
            process_file(path, args, conn)
            seen_hashes[cache_key] = current_hash
        except Exception as exc:
            print(f"[ERROR] Failed to process {path}: {exc}")
            traceback.print_exc()


def main() -> int:
    args = build_arg_parser().parse_args()

    if args.file:
        file_path = Path(args.file)
        if not file_path.exists():
            print(f"File does not exist: {file_path}")
            return 1
    else:
        folder = Path(args.folder)
        if not folder.exists():
            print(f"Folder does not exist: {folder}")
            return 1

    conn = None
    if args.conn:
        if pyodbc is None:
            print("pyodbc is not installed. Run: pip install pyodbc")
            return 1
        conn = pyodbc.connect(args.conn)
        print(f"Connected to SQL Server: {args.main_table}, {args.provider_table}")
    else:
        print("No SQL connection string supplied. Running in dry-run mode.")

    seen_hashes: Dict[str, str] = {}
    scan_and_process(args, conn, seen_hashes)

    if args.watch and not args.once:
        watch_target = args.file if args.file else args.folder
        print(f"Watching {watch_target} every {args.poll_seconds} seconds. Press Ctrl+C to stop.")
        while True:
            time.sleep(args.poll_seconds)
            scan_and_process(args, conn, seen_hashes)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
