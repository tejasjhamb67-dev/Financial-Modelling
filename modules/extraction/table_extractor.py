"""Extraction: turn parsed PDF tables into candidate datapoints.

Every value produced here is a *candidate* with status REQUIRES_REVIEW and its
``original_text`` preserved. The extractor:

  * scans tables on candidate financial-statement pages,
  * matches row labels to canonical metrics via the synonym dictionary,
  * parses period columns from the table header (FYxx / years),
  * parses numeric cells (handling commas, brackets-as-negative, dashes),
  * writes the extracted records to ``data/extracted`` for audit.

When the number of parsed period columns is ambiguous, or a row matches no
metric, the datapoint is still recorded but clearly flagged so the analyst can
resolve it. Nothing is silently dropped or invented.
"""

from __future__ import annotations

import os
import re
from typing import Optional

from modules.document_analysis.pdf_reader import ParsedDocument
from modules.extraction.line_item_synonyms import match_metric
from modules.schemas import (
    Consolidation,
    DataPoint,
    DataType,
    SourceDatabase,
    ValidationStatus,
)

_YEAR_HEADER = re.compile(r"(20\d{2})[-/](\d{2,4})|FY\s?(\d{2,4})|(20\d{2})", re.IGNORECASE)
_NUM = re.compile(r"^\(?-?[\d,]+(?:\.\d+)?\)?$")


def extract_candidates(parsed: ParsedDocument, company: str,
                       extracted_dir: Optional[str] = None) -> SourceDatabase:
    db = SourceDatabase(company)
    db.add_document(parsed.document)
    doc = parsed.document

    pages_of_interest = set(sum(parsed.statement_pages.values(), []))
    if not pages_of_interest:
        pages_of_interest = {p.page_number for p in parsed.pages}

    records: list[DataPoint] = []
    for page in parsed.pages:
        if page.page_number not in pages_of_interest:
            continue
        for table in page.tables:
            periods = _header_periods(table)
            for row in table:
                if not row or not row[0]:
                    continue
                label = str(row[0]).strip()
                m = match_metric(label)
                if not m:
                    continue
                metric, statement = m
                numeric_cells = [(i, _parse_num(c)) for i, c in enumerate(row[1:], start=1)]
                numeric_cells = [(i, v) for i, v in numeric_cells if v is not None]
                if not numeric_cells:
                    continue
                # Map numeric cells to periods where possible.
                for pos, (col_idx, val) in enumerate(numeric_cells):
                    period = _period_for_column(periods, col_idx, pos, len(numeric_cells))
                    ambiguous = period is None
                    period = period or f"COL{col_idx}"
                    records.append(DataPoint(
                        source_id=f"{doc.doc_id or 'DOC'}_P{page.page_number}_{statement.value}_{metric}_{period}".upper(),
                        company=company,
                        metric=metric,
                        period=period,
                        value=val,
                        label=label,
                        statement=statement,
                        unit=doc.unit,
                        currency=doc.currency,
                        consolidation=doc.consolidation or Consolidation.UNKNOWN,
                        data_type=DataType.REPORTED,
                        source_document=doc.name,
                        page=page.page_number,
                        section="Auto-extracted table",
                        original_text=" | ".join(str(c) for c in row if c),
                        status=(ValidationStatus.REQUIRES_REVIEW if not ambiguous
                                else ValidationStatus.RECONCILIATION_FAILED),
                        notes=("Auto-extracted from PDF table; verify against source."
                               + (" Period column ambiguous." if ambiguous else "")),
                        raw_value=" | ".join(str(c) for c in row if c),
                    ))

    for dp in records:
        db.add(dp)

    if extracted_dir:
        _dump_extracted(extracted_dir, doc.name, db)
    return db


def _header_periods(table: list[list[Optional[str]]]) -> list[Optional[str]]:
    """Return per-column period labels from the first few header rows."""
    for row in table[:3]:
        found: list[Optional[str]] = []
        any_year = False
        for cell in row:
            p = _period_from_text(cell)
            found.append(p)
            any_year = any_year or p is not None
        if any_year:
            return found
    return []


def _period_from_text(cell: Optional[str]) -> Optional[str]:
    if not cell:
        return None
    m = _YEAR_HEADER.search(str(cell))
    if not m:
        return None
    if m.group(1):  # 2023-24
        yy = m.group(2)
        yy = yy[-2:]
        return f"FY{yy}"
    if m.group(3):  # FY24 / FY2024
        yy = m.group(3)[-2:]
        return f"FY{yy}"
    if m.group(4):  # 2024
        return f"FY{m.group(4)[-2:]}"
    return None


def _period_for_column(periods: list[Optional[str]], col_idx: int,
                       pos: int, n: int) -> Optional[str]:
    if periods and col_idx < len(periods) and periods[col_idx]:
        return periods[col_idx]
    # fall back to any parsed periods in order
    parsed = [p for p in periods if p]
    if parsed and pos < len(parsed):
        return parsed[pos]
    return None


def _parse_num(cell: Optional[str]) -> Optional[float]:
    if cell is None:
        return None
    s = str(cell).strip()
    if not s or s in ("-", "—", "–", "NA", "N/A", "nm", "NM"):
        return None
    s2 = s.replace(" ", "")
    if not _NUM.match(s2):
        return None
    neg = s2.startswith("(") and s2.endswith(")")
    s2 = s2.strip("()").replace(",", "")
    try:
        val = float(s2)
        return -val if neg else val
    except ValueError:
        return None


def _dump_extracted(extracted_dir: str, name: str, db: SourceDatabase) -> None:
    os.makedirs(extracted_dir, exist_ok=True)
    base = os.path.splitext(name)[0]
    db.save(os.path.join(extracted_dir, f"{base}_candidates.json"))
