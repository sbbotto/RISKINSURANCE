from __future__ import annotations

import hashlib
import json
import re
import zipfile
from datetime import datetime, timezone
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Iterable
import xml.etree.ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

LABEL_MAP: dict[str, str] = {
    "Claim #:": "claim_number",
    "Claim Mgr:": "claim_manager",
    "Date:": "report_date",
    "Insured": "insured",
    "Div/Hos Notice": "div_hos_notice",
    "Date of Loss": "date_of_loss",
    "Date of Report": "date_of_report",
    "Date of Est.": "date_of_est",
    "Date of Suit": "date_of_suit",
    "Plaintiff": "plaintiff_name",
    "HCA Named": "hca_named",
    "Jurisdiction": "jurisdiction",
    "Prim Ins Limit": "primary_insurance_limit",
    "Indemnity Res.": "indemnity_reserve",
    "LAE Paid": "lae_paid",
    "Defense Counsel:": "defense_counsel",
    "Defense Firm:": "defense_firm",
    "Plaintiff Counsel:": "plaintiff_counsel",
    "Plaintiff Firm:": "plaintiff_firm",
    "Authority Req:": "authority_required",
    "Demand/Offer:": "demand_offer",
    "Trial Date:": "trial_date",
    "Mediation Date:": "mediation_date",
    "% Chance DV:": "chance_dv",
    "Verdict Value:": "verdict_value",
    "Settlement Value:": "settlement_value",
    "Executive Summary:": "executive_summary",
    "Resolution Strategy:": "resolution_strategy",
    "FACTS:": "facts",
    "INJURY:": "injury",
    "DAMAGES:": "damages",
    "ALLEGATIONS:": "allegations",
    "DEFENSES:": "defenses",
    "PEER REVIEW/REMEDIATION:": "peer_review_remediation",
    "INTERNAL REVIEW:": "internal_review",
    "EXPERTS:": "experts",
    "Defense:": "defense_section",
}

PROVIDER_HEADERS = [
    "involved_provider_specialty",
    "relationship_to_facility",
    "carrier_limits",
    "prior_hci_claim_involvement",
    "hci_insured_status",
]

SECTION_ORDER = [
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
]

SECTION_LABEL_TO_FIELD = {
    "Executive Summary:": "executive_summary",
    "Resolution Strategy:": "resolution_strategy",
    "FACTS:": "facts",
    "INJURY:": "injury",
    "PLAINTIFF:": "plaintiff_section",
    "DAMAGES:": "damages",
    "ALLEGATIONS:": "allegations",
    "DEFENSES:": "defenses",
    "PEER REVIEW/REMEDIATION:": "peer_review_remediation",
    "INTERNAL REVIEW:": "internal_review",
    "EXPERTS:": "experts",
    "Defense:": "defense_section",
}

SECTION_LABELS = list(SECTION_LABEL_TO_FIELD.keys())

TABLE_SCHEMAS: dict[int, list[tuple[str, str]]] = {
    0: [("claim_number", "Claim #:"), ("claim_manager", "Claim Mgr:"), ("report_date", "Date:")],
    1: [("insured", "Insured"), ("div_hos_notice", "Div/Hos Notice"), ("date_of_loss", "Date of Loss"), ("date_of_report", "Date of Report"), ("date_of_est", "Date of Est."), ("date_of_suit", "Date of Suit")],
    3: [("plaintiff_name", "Plaintiff"), ("hca_named", "HCA Named"), ("jurisdiction", "Jurisdiction"), ("primary_insurance_limit", "Prim Ins Limit"), ("indemnity_reserve", "Indemnity Res."), ("lae_paid", "LAE Paid")],
    5: [("defense_counsel", "Defense Counsel:"), ("defense_firm", "Defense Firm:"), ("plaintiff_counsel", "Plaintiff Counsel:"), ("plaintiff_firm", "Plaintiff Firm:")],
    12: [("authority_required", "Authority Req:"), ("demand_offer", "Demand/Offer:"), ("trial_date", "Trial Date:"), ("mediation_date", "Mediation Date:")],
    14: [("chance_dv", "% Chance DV:"), ("verdict_value", "Verdict Value:"), ("settlement_value", "Settlement Value:")],
}

@dataclass
class ProviderRow:
    row_index: int
    involved_provider_specialty: str | None = None
    relationship_to_facility: str | None = None
    carrier_limits: str | None = None
    prior_hci_claim_involvement: str | None = None
    hci_insured_status: str | None = None
    raw_text: str | None = None


@dataclass
class ParsedReport:
    fields: dict[str, Any]
    providers: list[ProviderRow]
    file_name: str
    file_path: str
    file_hash: str
    last_modified_utc: str



def normalize_text(text: str | None) -> str:
    if not text:
        return ""
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()



def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()



def _text_of(elem: ET.Element) -> str:
    return "".join(t.text or "" for t in elem.findall('.//w:t', NS))



def _iter_document_children(body: ET.Element) -> Iterable[ET.Element]:
    for child in list(body):
        yield child



def _extract_table_rows(tbl: ET.Element) -> list[list[str]]:
    rows: list[list[str]] = []
    for tr in tbl.findall("w:tr", NS):
        row: list[str] = []
        for tc in tr.findall("w:tc", NS):
            texts = []
            for p in tc.findall(".//w:p", NS):
                txt = normalize_text(_text_of(p))
                if txt:
                    texts.append(txt)
            row.append("\n".join(texts).strip())
        rows.append(row)
    return rows



def _extract_paragraphs(root: ET.Element) -> list[str]:
    body = root.find("w:body", NS)
    if body is None:
        return []
    paragraphs: list[str] = []

    def walk(elem: ET.Element) -> None:
        for child in list(elem):
            tag = child.tag.split("}")[-1]
            if tag == "p":
                txt = normalize_text(_text_of(child))
                if txt:
                    paragraphs.append(txt)
            elif tag == "tbl":
                for tr in child.findall("w:tr", NS):
                    for tc in tr.findall("w:tc", NS):
                        for p in tc.findall(".//w:p", NS):
                            txt = normalize_text(_text_of(p))
                            if txt:
                                paragraphs.append(txt)
            else:
                walk(child)

    walk(body)
    return paragraphs



def _extract_inline_value(text: str, label: str) -> str | None:
    pattern = re.escape(label)
    m = re.search(pattern, text, flags=re.IGNORECASE)
    if not m:
        return None
    remainder = text[m.end():].strip()
    remainder = remainder.lstrip("-:–—")
    remainder = remainder.strip()
    return remainder or None



def _extract_labeled_fields(text: str, labels: dict[str, str]) -> dict[str, str]:
    found: dict[str, str] = {}
    if not text:
        return found

    normalized = normalize_text(text)
    # Prefer labels at the beginning of the text block; this avoids accidental
    # matches inside narrative sentences such as "The plaintiff ...".
    for label, field in sorted(labels.items(), key=lambda kv: len(kv[0]), reverse=True):
        if re.match(rf"^{re.escape(label)}(?=\s|$)", normalized, flags=re.IGNORECASE):
            value = normalized[len(label):].lstrip(" :	-–—").strip()
            if value:
                found[field] = value

    return found



def _strip_heading_prefix(text: str, label: str) -> str:
    if not text:
        return ""
    m = re.search(re.escape(label), text, flags=re.IGNORECASE)
    if not m:
        return text.strip()
    return text[m.end():].lstrip(" :\t-–—").strip()



def _extract_sections(paragraphs: list[str]) -> dict[str, str]:
    section_text: dict[str, list[str]] = {field: [] for field in SECTION_ORDER}
    current_field: str | None = None

    def is_heading(text: str) -> tuple[str | None, str]:
        for label, field in SECTION_LABEL_TO_FIELD.items():
            if re.match(rf"^{re.escape(label)}\s*.*$", text, flags=re.IGNORECASE):
                return field, _strip_heading_prefix(text, label)
        return None, text

    for text in paragraphs:
        field, inline = is_heading(text)
        if field:
            current_field = field
            if inline:
                section_text[current_field].append(inline)
            continue
        if current_field:
            # Stop when we hit obvious top-level table-like text elsewhere? Not necessary.
            section_text[current_field].append(text)

    return {k: normalize_text("\n".join(v)) for k, v in section_text.items() if normalize_text("\n".join(v))}



def _parse_provider_rows(tbl_rows: list[list[str]]) -> list[ProviderRow]:
    header_idx = None
    for i, row in enumerate(tbl_rows):
        joined = " ".join(c for c in row if c)
        if "Involved Provider(s) and Specialty" in joined and "Relationship to Facility" in joined:
            header_idx = i
            break
    if header_idx is None:
        return []

    providers: list[ProviderRow] = []
    row_number = 0
    for row in tbl_rows[header_idx + 1 :]:
        cleaned = [normalize_text(c) for c in row if normalize_text(c)]
        if not cleaned:
            continue
        if any(label in " ".join(cleaned) for label in ["Authority Req:", "Executive Summary:", "FACTS:"]):
            break

        row_number += 1
        provider = ProviderRow(row_index=row_number, raw_text=" | ".join(cleaned))
        if len(cleaned) >= 1:
            provider.involved_provider_specialty = cleaned[0]
        if len(cleaned) >= 2:
            provider.relationship_to_facility = cleaned[1]
        if len(cleaned) >= 3:
            provider.carrier_limits = cleaned[2]
        if len(cleaned) >= 4:
            provider.prior_hci_claim_involvement = cleaned[3]
        if len(cleaned) >= 5:
            provider.hci_insured_status = cleaned[4]
        providers.append(provider)
    return providers



def parse_docx(path: str | Path) -> ParsedReport:
    path = Path(path)
    file_hash = sha256_file(path)
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    paragraphs = _extract_paragraphs(root)

    # Extract table rows for the structured/provider section.
    body = root.find("w:body", NS)
    tables = [child for child in list(body or []) if child.tag.split("}")[-1] == "tbl"]
    table_rows: list[list[str]] = _extract_table_rows(tables[0]) if tables else []

    fields: dict[str, Any] = {}

    # Table-based extraction using the known report layout.
    for row_index, schema in TABLE_SCHEMAS.items():
        if row_index >= len(table_rows):
            continue
        row = table_rows[row_index]
        for cell_index, (field, label) in enumerate(schema):
            if cell_index >= len(row):
                continue
            cell_text = normalize_text(row[cell_index])
            if not cell_text:
                continue
            if cell_text.lower().startswith(label.lower()):
                inline = _extract_inline_value(cell_text, label)
                if inline:
                    fields[field] = inline
            else:
                # If the layout is edited and the label is removed, capture any text in the expected cell.
                fields[field] = cell_text

    # Paragraph-based extraction for the narrative sections.
    for para in paragraphs:
        fields.update(_extract_labeled_fields(para, SECTION_LABEL_TO_FIELD))
        for label, field in SECTION_LABEL_TO_FIELD.items():
            if para.lower().startswith(label.lower()):
                inline = _strip_heading_prefix(para, label)
                if inline:
                    fields[field] = inline

    sections = _extract_sections(paragraphs)
    fields.update(sections)

    providers = _parse_provider_rows(table_rows)
    return ParsedReport(
        fields=fields,
        providers=providers,
        file_name=path.name,
        file_path=str(path.resolve()),
        file_hash=file_hash,
        last_modified_utc=datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc).strftime('%Y-%m-%d %H:%M:%S'),
    )



def parsed_report_to_json(report: ParsedReport) -> str:
    payload = {
        "fields": report.fields,
        "providers": [asdict(p) for p in report.providers],
        "file_name": report.file_name,
        "file_path": report.file_path,
        "file_hash": report.file_hash,
        "last_modified_utc": report.last_modified_utc,
    }
    return json.dumps(payload, indent=2, ensure_ascii=False)
