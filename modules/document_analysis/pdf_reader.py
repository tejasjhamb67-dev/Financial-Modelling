"""Document understanding for PDF source files.

Parses PDFs with ``pdfplumber`` and produces:
  * a :class:`Document` inventory record (fiscal year, units, currency,
    consolidation basis, candidate financial-statement pages), and
  * a per-page dump of text and tables written to ``data/raw`` for audit.

This module never *interprets* numbers into the model. It only reads the
document structure. Extraction (turning table cells into candidate datapoints)
happens in :mod:`modules.extraction.table_extractor`, and every candidate is
flagged for analyst review - successful PDF parsing is not successful
financial-data extraction.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Any, Optional

from modules.schemas import Consolidation, Document

try:
    import pdfplumber
    _HAVE_PDFPLUMBER = True
except Exception:  # pragma: no cover - optional dependency
    _HAVE_PDFPLUMBER = False


_UNIT_PATTERNS = [
    (r"in\s+(?:rs\.?|inr|₹)\s*crore", "INR crore"),
    (r"₹\s*crore|rs\.?\s*crore|inr\s*crore|in\s*crores?", "INR crore"),
    (r"₹\s*lakh|rs\.?\s*lakh|in\s*lakhs?", "INR lakh"),
    (r"₹\s*million|rs\.?\s*million|inr\s*million", "INR million"),
    (r"usd\s*million|\$\s*million|us\$\s*million", "USD million"),
    (r"usd\s*thousand|\$\s*thousand", "USD thousand"),
    (r"in\s*millions?", "million"),
    (r"in\s*thousands?", "thousand"),
]

_CURRENCY_PATTERNS = [
    (r"₹|inr|rupee", "INR"),
    (r"\bus\$|usd|\$\b|dollar", "USD"),
    (r"€|eur|euro", "EUR"),
    (r"£|gbp|pound", "GBP"),
]

_FY_PATTERN = re.compile(r"\b(?:FY\s?)?(20\d{2})[-/]?(?:\s?(\d{2,4}))?\b", re.IGNORECASE)

_STATEMENT_KEYWORDS = {
    "IS": ["statement of profit and loss", "profit and loss", "income statement",
           "statement of operations", "consolidated statement of profit"],
    "BS": ["balance sheet", "statement of financial position"],
    "CF": ["cash flow statement", "statement of cash flows", "cash flow"],
}


@dataclass
class ParsedPage:
    page_number: int
    text: str
    tables: list[list[list[Optional[str]]]] = field(default_factory=list)


@dataclass
class ParsedDocument:
    document: Document
    pages: list[ParsedPage]
    statement_pages: dict[str, list[int]]


def parse_pdf(path: str, raw_dir: Optional[str] = None,
              max_pages: Optional[int] = None) -> ParsedDocument:
    """Parse a PDF into pages + a Document inventory record."""
    if not _HAVE_PDFPLUMBER:
        raise RuntimeError("pdfplumber is required for PDF ingestion. `pip install pdfplumber`.")

    name = os.path.basename(path)
    pages: list[ParsedPage] = []
    statement_pages: dict[str, list[int]] = {"IS": [], "BS": [], "CF": []}
    cover_text = ""

    with pdfplumber.open(path) as pdf:
        n = len(pdf.pages) if max_pages is None else min(max_pages, len(pdf.pages))
        for i in range(n):
            page = pdf.pages[i]
            text = page.extract_text() or ""
            tables = []
            try:
                tables = page.extract_tables() or []
            except Exception:
                tables = []
            pages.append(ParsedPage(page_number=i + 1, text=text, tables=tables))
            if i < 3:
                cover_text += "\n" + text
            low = text.lower()
            for key, kws in _STATEMENT_KEYWORDS.items():
                if any(kw in low for kw in kws):
                    statement_pages[key].append(i + 1)

    all_text = cover_text + "\n" + "\n".join(p.text for p in pages[:40])
    doc = Document(
        name=name,
        doc_type=_guess_doc_type(name, cover_text),
        fiscal_year=_guess_fiscal_year(cover_text),
        consolidation=_guess_consolidation(all_text),
        currency=_guess_currency(all_text),
        unit=_guess_unit(all_text),
        relevant_pages=sorted(set(sum(statement_pages.values(), []))),
        source_path=path,
    )

    if raw_dir:
        _dump_raw(raw_dir, name, pages)

    return ParsedDocument(document=doc, pages=pages, statement_pages=statement_pages)


def _guess_doc_type(name: str, cover: str) -> str:
    low = (name + " " + cover).lower()
    if "investor" in low or "presentation" in low:
        return "Investor Presentation"
    if "result" in low and "present" in low:
        return "Results Presentation"
    if "annual report" in low or "integrated report" in low:
        return "Annual Report"
    return "Annual Report"


def _guess_fiscal_year(cover: str) -> Optional[str]:
    matches = _FY_PATTERN.findall(cover)
    years = []
    for a, b in matches:
        years.append(int(a))
        if b and len(b) == 4:
            years.append(int(b))
    if not years:
        return None
    yr = max(y for y in years if 2000 <= y <= 2099)
    return f"FY{str(yr)[-2:]}"


def _guess_consolidation(text: str) -> Consolidation:
    low = text.lower()
    has_consol = "consolidated" in low
    has_standalone = "standalone" in low
    if has_consol:
        return Consolidation.CONSOLIDATED
    if has_standalone:
        return Consolidation.STANDALONE
    return Consolidation.UNKNOWN


def _guess_currency(text: str) -> Optional[str]:
    low = text.lower()
    for pat, cur in _CURRENCY_PATTERNS:
        if re.search(pat, low):
            return cur
    return None


def _guess_unit(text: str) -> Optional[str]:
    low = text.lower()
    for pat, unit in _UNIT_PATTERNS:
        if re.search(pat, low):
            return unit
    return None


def _dump_raw(raw_dir: str, name: str, pages: list[ParsedPage]) -> None:
    os.makedirs(raw_dir, exist_ok=True)
    base = os.path.splitext(name)[0]
    with open(os.path.join(raw_dir, f"{base}.txt"), "w", encoding="utf-8") as fh:
        for p in pages:
            fh.write(f"\n===== PAGE {p.page_number} =====\n")
            fh.write(p.text or "")
