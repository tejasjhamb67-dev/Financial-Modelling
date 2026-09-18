"""Ingestion: turn a structured *company package* into a SourceDatabase.

A company package (YAML/JSON) is the reliable, provenance-carrying input path.
It mirrors what a careful analyst (or the PDF extraction module) would pull out
of the source documents, with every value tagged to a document/page. The
package deliberately does NOT contain constructed totals unless the source only
disclosed a total - the financial-statement engine decides whether to build a
total by formula (components present) or keep it as a reported input.

Package shape (all statement blocks optional)::

    company: "Acme Chemicals Ltd"
    currency: INR
    unit: "INR crore"
    consolidation: consolidated
    historical_periods: [FY22, FY23, FY24, FY25, FY26]
    documents:
      - {name: "FY26 Annual Report", doc_type: Annual Report, fiscal_year: FY26,
         page: 0, consolidation: consolidated}
    provenance:                     # optional per-metric page / section hints
      revenue: {document: "FY26 Annual Report", page: 142, section: "Statement of P&L"}
    income_statement:
      revenue:   {FY22: 8000, FY23: 8800, ...}
      cogs:      {FY22: 4800, ...}
    balance_sheet: {...}
    cash_flow: {...}
    segments:
      revenue: {"Specialty": {FY22: 5000, ...}, "Commodity": {FY22: 3000, ...}}
    operating_drivers:
      capacity_tonnes: {unit: "'000 tonnes", values: {FY22: 250, ...}}
    restatements:
      - {metric: revenue, period: FY24, old_value: 9600, old_document: "FY24 AR",
         new_value: 9500, new_document: "FY26 AR", reason: "Discontinued op reclass"}
"""

from __future__ import annotations

import json
import os
from typing import Any

import yaml

from modules.schemas import (
    Consolidation,
    DataPoint,
    DataType,
    Document,
    Restatement,
    SourceDatabase,
    Statement,
    ValidationStatus,
)

_STATEMENT_MAP = {
    "income_statement": Statement.INCOME_STATEMENT,
    "balance_sheet": Statement.BALANCE_SHEET,
    "cash_flow": Statement.CASH_FLOW,
    "operating_drivers": Statement.OPERATING,
    "segments": Statement.SEGMENT,
}


def load_package(path: str) -> SourceDatabase:
    """Load a company package file (``.yaml``/``.yml``/``.json``)."""
    with open(path, encoding="utf-8") as fh:
        if path.endswith(".json"):
            pkg = json.load(fh)
        else:
            pkg = yaml.safe_load(fh)
    return build_database(pkg)


def build_database(pkg: dict[str, Any]) -> SourceDatabase:
    company = pkg.get("company", "Unknown Company")
    db = SourceDatabase(company)

    currency = pkg.get("currency")
    unit = pkg.get("unit")
    consol = Consolidation(pkg.get("consolidation", "consolidated"))

    # -- documents -------------------------------------------------------
    for i, d in enumerate(pkg.get("documents", []) or []):
        doc = Document(
            name=d["name"],
            doc_type=d.get("doc_type", "Annual Report"),
            fiscal_year=d.get("fiscal_year"),
            publication_date=d.get("publication_date"),
            reporting_period=d.get("reporting_period"),
            consolidation=Consolidation(d.get("consolidation", consol.value)),
            currency=d.get("currency", currency),
            unit=d.get("unit", unit),
            relevant_pages=d.get("relevant_pages", []) or [],
            source_path=d.get("source_path"),
            doc_id=d.get("doc_id", f"DOC{i+1:02d}"),
            notes=d.get("notes", ""),
        )
        db.add_document(doc)

    default_doc = db.documents[0].name if db.documents else pkg.get("default_document", "Source Document")
    provenance = pkg.get("provenance", {}) or {}

    def _src_id(metric: str, period: str, stmt: Statement) -> str:
        return f"{stmt.value}_{metric}_{period}".upper()

    # -- statement / driver / segment datapoints -------------------------
    for block, statement in _STATEMENT_MAP.items():
        data = pkg.get(block)
        if not data:
            continue
        if block == "operating_drivers":
            _load_drivers(db, data, company, currency, consol, provenance, default_doc)
            continue
        if block == "segments":
            _load_segments(db, data, company, currency, unit, consol, provenance, default_doc)
            continue
        for metric, series in data.items():
            prov = provenance.get(metric, {})
            doc = prov.get("document", default_doc)
            page = prov.get("page")
            section = prov.get("section")
            for period, value in series.items():
                dp = DataPoint(
                    source_id=_src_id(metric, period, statement),
                    company=company,
                    metric=metric,
                    period=period,
                    value=_num(value),
                    label=prov.get("label"),
                    statement=statement,
                    unit=unit,
                    currency=currency,
                    consolidation=consol,
                    data_type=DataType.REPORTED,
                    source_document=doc,
                    page=page,
                    section=section,
                    status=(ValidationStatus.VERIFIED if _num(value) is not None
                            else ValidationStatus.NOT_FOUND),
                    raw_value=value,
                )
                db.add(dp)

    # -- restatements ----------------------------------------------------
    for r in pkg.get("restatements", []) or []:
        old_v, new_v = _num(r.get("old_value")), _num(r.get("new_value"))
        pct = None
        if old_v not in (None, 0) and new_v is not None:
            pct = new_v / old_v - 1.0
        db.add_restatement(Restatement(
            metric=r["metric"],
            period=r["period"],
            old_value=old_v,
            new_value=new_v,
            old_document=r.get("old_document", "Older report"),
            new_document=r.get("new_document", "Newer report"),
            pct_change=pct,
            reason=r.get("reason", "Undetermined - requires analyst review"),
            treatment=r.get("treatment",
                            "Newer disclosure retained; older value preserved in notes"),
        ))
    return db


def _load_drivers(db, data, company, currency, consol, provenance, default_doc):
    for metric, spec in data.items():
        unit = spec.get("unit")
        series = spec.get("values", {})
        prov = provenance.get(metric, {})
        for period, value in series.items():
            db.add(DataPoint(
                source_id=f"OP_{metric}_{period}".upper(),
                company=company,
                metric=metric,
                period=period,
                value=_num(value),
                label=spec.get("label", metric.replace("_", " ").title()),
                statement=Statement.OPERATING,
                category="operating_driver",
                unit=unit,
                currency=currency if spec.get("is_monetary") else None,
                consolidation=consol,
                data_type=DataType.REPORTED,
                source_document=prov.get("document", default_doc),
                page=prov.get("page"),
                status=(ValidationStatus.VERIFIED if _num(value) is not None
                        else ValidationStatus.NOT_FOUND),
                raw_value=value,
            ))


def _load_segments(db, data, company, currency, unit, consol, provenance, default_doc):
    # data: {dimension: {segment_name: {period: value}}}
    for dimension, segs in data.items():
        for seg_name, series in segs.items():
            metric = f"segment_{dimension}_{_slug(seg_name)}"
            for period, value in series.items():
                db.add(DataPoint(
                    source_id=f"SEG_{metric}_{period}".upper(),
                    company=company,
                    metric=metric,
                    period=period,
                    value=_num(value),
                    label=f"{seg_name} ({dimension})",
                    statement=Statement.SEGMENT,
                    category=dimension,
                    unit=unit,
                    currency=currency,
                    consolidation=consol,
                    data_type=DataType.REPORTED,
                    source_document=default_doc,
                    status=(ValidationStatus.VERIFIED if _num(value) is not None
                            else ValidationStatus.NOT_FOUND),
                    raw_value=value,
                ))


def _num(v: Any):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "N/A", "NA", "Not Disclosed", "Not Found", "-", "nm", "NM"):
        return None
    # bracketed negatives
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        val = float(s)
        return -val if neg else val
    except ValueError:
        return None


def _slug(name: str) -> str:
    return "".join(c if c.isalnum() else "_" for c in name).strip("_").lower()
