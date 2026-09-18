"""Disclosure mapping.

Before building the model we understand how the company actually reports itself,
and let that drive which sheets/schedules the workbook will contain. We never
force a company into a generic template - a sheet is only built when the source
data supports it.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from modules.schemas import SourceDatabase, Statement


@dataclass
class DisclosureMap:
    company: str
    has_income_statement: bool = False
    has_balance_sheet: bool = False
    has_cash_flow: bool = False
    has_revenue: bool = False
    has_segments: bool = False
    segment_dimensions: list[str] = field(default_factory=list)
    segments: dict[str, list[str]] = field(default_factory=dict)  # dimension -> names
    has_operating_drivers: bool = False
    operating_drivers: list[str] = field(default_factory=list)
    has_working_capital: bool = False
    has_capex_or_dna: bool = False
    has_debt: bool = False
    notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "company": self.company,
            "has_income_statement": self.has_income_statement,
            "has_balance_sheet": self.has_balance_sheet,
            "has_cash_flow": self.has_cash_flow,
            "has_revenue": self.has_revenue,
            "has_segments": self.has_segments,
            "segment_dimensions": self.segment_dimensions,
            "segments": self.segments,
            "has_operating_drivers": self.has_operating_drivers,
            "operating_drivers": self.operating_drivers,
            "has_working_capital": self.has_working_capital,
            "has_capex_or_dna": self.has_capex_or_dna,
            "has_debt": self.has_debt,
            "notes": self.notes,
        }


def _present(db: SourceDatabase, metric: str) -> bool:
    return any(dp.metric == metric and dp.value is not None for dp in db.datapoints)


def map_disclosures(db: SourceDatabase) -> DisclosureMap:
    dm = DisclosureMap(company=db.company)

    is_metrics = {dp.metric for dp in db.by_statement(Statement.INCOME_STATEMENT) if dp.value is not None}
    bs_metrics = {dp.metric for dp in db.by_statement(Statement.BALANCE_SHEET) if dp.value is not None}
    cf_metrics = {dp.metric for dp in db.by_statement(Statement.CASH_FLOW) if dp.value is not None}

    dm.has_income_statement = len(is_metrics) >= 2
    dm.has_balance_sheet = len(bs_metrics) >= 2
    dm.has_cash_flow = len(cf_metrics) >= 1
    dm.has_revenue = "revenue" in is_metrics

    # segments
    seg_dps = [dp for dp in db.by_statement(Statement.SEGMENT) if dp.value is not None]
    if seg_dps:
        dm.has_segments = True
        for dp in seg_dps:
            dim = dp.category or "segment"
            dm.segments.setdefault(dim, [])
            name = (dp.label or dp.metric).split(" (")[0]
            if name not in dm.segments[dim]:
                dm.segments[dim].append(name)
        dm.segment_dimensions = list(dm.segments.keys())

    # operating drivers
    drivers = sorted({dp.metric for dp in db.by_statement(Statement.OPERATING) if dp.value is not None})
    if drivers:
        dm.has_operating_drivers = True
        dm.operating_drivers = drivers

    # working capital requires receivables/inventory/payables + revenue
    wc_ok = ("revenue" in is_metrics) and (
        {"receivables", "inventory", "payables"} & bs_metrics)
    dm.has_working_capital = bool(wc_ok)

    # capex / D&A
    dm.has_capex_or_dna = ("capex" in cf_metrics) or ("depreciation" in is_metrics) or ("ppe" in bs_metrics)

    # debt
    dm.has_debt = bool({"short_term_debt", "long_term_debt"} & bs_metrics)

    if not dm.has_cash_flow:
        dm.notes.append("Historical cash flow not fully disclosed; CF sheet will be partly derived.")
    if not dm.has_segments:
        dm.notes.append("No segment disclosure found; revenue build kept single-line.")
    return dm
