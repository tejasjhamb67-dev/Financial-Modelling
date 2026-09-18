"""
Core data schemas for the Financial Model Construction Agent.

These schemas are the backbone of the whole system. Every number that ends up
in the Excel model must be able to trace back to a ``DataPoint`` here, which
in turn records where it came from, how it was obtained and how confident we
are in it. The guiding rule of the whole agent lives in this module:

    A visible missing number is always preferable to a fabricated number.

Nothing in this file ever invents a value. Missing information is represented
explicitly (value is ``None`` and status is ``NOT_FOUND`` / ``REQUIRES_REVIEW``)
so it survives all the way to the workbook instead of being silently filled.
"""

from __future__ import annotations

import enum
import json
from dataclasses import dataclass, field, asdict
from typing import Any, Optional


# ---------------------------------------------------------------------------
# Enumerations
# ---------------------------------------------------------------------------


class DataType(str, enum.Enum):
    """How a number came to exist.

    These classifications must stay distinct throughout the system and drive
    the Excel colour coding (see ``excel.formatting``).
    """

    REPORTED = "REPORTED"          # Directly disclosed in a source document
    DERIVED = "DERIVED"            # Calculated from reported values
    LINKED = "LINKED"              # Pulled from another worksheet
    ASSUMPTION = "USER ASSUMPTION" # Entered / to be entered by the analyst


class Consolidation(str, enum.Enum):
    CONSOLIDATED = "consolidated"
    STANDALONE = "standalone"
    UNKNOWN = "unknown"


class Statement(str, enum.Enum):
    """Which part of the model a datapoint belongs to."""

    INCOME_STATEMENT = "IS"
    BALANCE_SHEET = "BS"
    CASH_FLOW = "CF"
    OPERATING = "OP"       # Operating / KPI drivers
    SEGMENT = "SEG"        # Segment / geography splits
    OTHER = "OTHER"


class ValidationStatus(str, enum.Enum):
    VERIFIED = "VERIFIED"
    DERIVED_OK = "DERIVED_OK"
    REQUIRES_REVIEW = "REQUIRES_REVIEW"
    NOT_FOUND = "NOT_FOUND"
    NOT_DISCLOSED = "NOT_DISCLOSED"
    RESTATED = "RESTATED"
    RECONCILIATION_FAILED = "RECONCILIATION_FAILED"


# Sentinel strings for cells that legitimately have no number. Never guess.
NA = "N/A"
NOT_DISCLOSED = "Not Disclosed"
NOT_FOUND = "Not Found"
REQUIRES_REVIEW = "Requires Review"


# ---------------------------------------------------------------------------
# Documents
# ---------------------------------------------------------------------------


@dataclass
class Document:
    """A source document supplied by the user (annual report, presentation...)."""

    name: str
    doc_type: str = "Annual Report"
    fiscal_year: Optional[str] = None
    publication_date: Optional[str] = None
    reporting_period: Optional[str] = None
    consolidation: Consolidation = Consolidation.UNKNOWN
    currency: Optional[str] = None
    unit: Optional[str] = None
    relevant_pages: list[int] = field(default_factory=list)
    source_path: Optional[str] = None
    doc_id: Optional[str] = None
    notes: str = ""

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["consolidation"] = self.consolidation.value
        return d


# ---------------------------------------------------------------------------
# Datapoints (the source database)
# ---------------------------------------------------------------------------


@dataclass
class DataPoint:
    """A single, provenance-carrying financial or operating number.

    ``value`` may be ``None`` when the number is genuinely missing. The rest of
    the metadata is preserved regardless so the analyst can always answer:
    what / when / where / how / why for every cell in the model.
    """

    source_id: str
    company: str
    metric: str                       # canonical metric key, e.g. "revenue"
    period: str                       # e.g. "FY24"
    value: Optional[float] = None
    label: Optional[str] = None       # human label, e.g. "Revenue from operations"
    statement: Statement = Statement.OTHER
    category: Optional[str] = None
    unit: Optional[str] = None
    currency: Optional[str] = None
    consolidation: Consolidation = Consolidation.UNKNOWN
    data_type: DataType = DataType.REPORTED
    source_document: Optional[str] = None
    page: Optional[int] = None
    section: Optional[str] = None
    original_text: Optional[str] = None
    status: ValidationStatus = ValidationStatus.REQUIRES_REVIEW
    notes: str = ""
    # provenance chain: raw -> normalized -> validated -> model
    raw_value: Optional[Any] = None

    def is_missing(self) -> bool:
        return self.value is None

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["statement"] = self.statement.value
        d["consolidation"] = self.consolidation.value
        d["data_type"] = self.data_type.value
        d["status"] = self.status.value
        return d

    @staticmethod
    def from_dict(d: dict[str, Any]) -> "DataPoint":
        d = dict(d)
        if "statement" in d and d["statement"] is not None:
            d["statement"] = Statement(d["statement"])
        if "consolidation" in d and d["consolidation"] is not None:
            d["consolidation"] = Consolidation(d["consolidation"])
        if "data_type" in d and d["data_type"] is not None:
            d["data_type"] = DataType(d["data_type"])
        if "status" in d and d["status"] is not None:
            d["status"] = ValidationStatus(d["status"])
        return DataPoint(**d)


@dataclass
class Restatement:
    """A detected difference in the same period across two documents."""

    metric: str
    period: str
    old_value: Optional[float]
    new_value: Optional[float]
    old_document: str
    new_document: str
    pct_change: Optional[float] = None
    reason: str = "Undetermined - requires analyst review"
    treatment: str = "Newer disclosure retained; older value preserved in notes"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


# ---------------------------------------------------------------------------
# The source database
# ---------------------------------------------------------------------------


class SourceDatabase:
    """In-memory registry of every datapoint plus documents and restatements.

    This is the single source of truth that the Excel builder consumes. It is
    persisted to ``sources/source_registry`` for auditability.
    """

    def __init__(self, company: str = "Unknown Company"):
        self.company = company
        self.documents: list[Document] = []
        self.datapoints: list[DataPoint] = []
        self.restatements: list[Restatement] = []
        # Verbatim "as-reported" statement blocks preserved exactly as disclosed
        # (the RAW layer, surfaced undistorted on the Reported Financials sheet).
        # Each block: {name, document, pages, lines:[{label, note, bold, indent,
        # values:{period: value}}]}
        self.reported_statements: list[dict] = []

    # -- mutation --------------------------------------------------------
    def add_document(self, doc: Document) -> None:
        self.documents.append(doc)

    def add(self, dp: DataPoint) -> None:
        self.datapoints.append(dp)

    def add_restatement(self, r: Restatement) -> None:
        self.restatements.append(r)

    # -- querying --------------------------------------------------------
    def get(self, metric: str, period: str,
            consolidation: Optional[Consolidation] = None) -> Optional[DataPoint]:
        """Return the best datapoint for a metric/period.

        Prefers the requested consolidation basis, then VERIFIED status, then
        the most recently added (newest disclosure).
        """
        candidates = [dp for dp in self.datapoints
                      if dp.metric == metric and dp.period == period]
        if consolidation is not None:
            pref = [dp for dp in candidates if dp.consolidation == consolidation]
            if pref:
                candidates = pref
        if not candidates:
            return None
        candidates.sort(key=lambda dp: (
            dp.status == ValidationStatus.VERIFIED,
            dp.value is not None,
        ))
        return candidates[-1]

    def value(self, metric: str, period: str,
              consolidation: Optional[Consolidation] = None) -> Optional[float]:
        dp = self.get(metric, period, consolidation)
        return dp.value if dp else None

    def metrics(self) -> list[str]:
        seen: list[str] = []
        for dp in self.datapoints:
            if dp.metric not in seen:
                seen.append(dp.metric)
        return seen

    def periods(self) -> list[str]:
        seen: list[str] = []
        for dp in self.datapoints:
            if dp.period not in seen:
                seen.append(dp.period)
        return seen

    def by_statement(self, statement: Statement) -> list[DataPoint]:
        return [dp for dp in self.datapoints if dp.statement == statement]

    # -- persistence -----------------------------------------------------
    def to_dict(self) -> dict[str, Any]:
        return {
            "company": self.company,
            "documents": [d.to_dict() for d in self.documents],
            "datapoints": [dp.to_dict() for dp in self.datapoints],
            "restatements": [r.to_dict() for r in self.restatements],
            "reported_statements": self.reported_statements,
        }

    def save(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(self.to_dict(), fh, indent=2, default=str)

    @staticmethod
    def load(path: str) -> "SourceDatabase":
        with open(path, encoding="utf-8") as fh:
            d = json.load(fh)
        db = SourceDatabase(d.get("company", "Unknown Company"))
        for doc in d.get("documents", []):
            doc = dict(doc)
            doc["consolidation"] = Consolidation(doc.get("consolidation", "unknown"))
            db.documents.append(Document(**doc))
        for dp in d.get("datapoints", []):
            db.datapoints.append(DataPoint.from_dict(dp))
        for r in d.get("restatements", []):
            db.restatements.append(Restatement(**r))
        db.reported_statements = d.get("reported_statements", []) or []
        return db

    # -- summary counts (for human-in-the-loop review) -------------------
    def summary(self) -> dict[str, int]:
        reported = sum(1 for dp in self.datapoints if dp.data_type == DataType.REPORTED and dp.value is not None)
        derived = sum(1 for dp in self.datapoints if dp.data_type == DataType.DERIVED and dp.value is not None)
        unresolved = sum(1 for dp in self.datapoints if dp.value is None)
        review = sum(1 for dp in self.datapoints if dp.status in (
            ValidationStatus.REQUIRES_REVIEW, ValidationStatus.RECONCILIATION_FAILED))
        return {
            "documents": len(self.documents),
            "datapoints": len(self.datapoints),
            "reported": reported,
            "derived": derived,
            "unresolved": unresolved,
            "requires_review": review,
            "restatements": len(self.restatements),
        }
