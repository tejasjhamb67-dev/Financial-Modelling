"""Model specification (the output of Gate 2).

An internal, explicit description of the model to be generated: periods,
currency/units, consolidation basis, which statement line items exist, which
schedules and valuation methods to build, and known exceptions. The Excel
engine consumes ONLY this spec + the source database, so the workbook is fully
determined and reproducible.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from modules.config_loader import AccountingTaxonomy, model_config
from modules.disclosure_mapping.mapper import DisclosureMap
from modules.schemas import SourceDatabase, Statement


def _fy_num(period: str) -> int:
    digits = "".join(c for c in period if c.isdigit())
    return int(digits) if digits else 0


def order_historical_periods(periods: list[str]) -> list[str]:
    """Sort FY labels oldest -> newest (chronology is non-negotiable)."""
    return sorted(set(periods), key=_fy_num)


@dataclass
class ModelSpec:
    company: str
    currency: Optional[str]
    unit: Optional[str]
    consolidation: str
    historical_periods: list[str]
    forecast_periods: list[str]
    disclosure: DisclosureMap
    income_statement_lines: list[str] = field(default_factory=list)
    balance_sheet_lines: list[str] = field(default_factory=list)
    cash_flow_lines: list[str] = field(default_factory=list)
    schedules: list[str] = field(default_factory=list)
    valuation_methods: list[str] = field(default_factory=list)
    sheets: list[str] = field(default_factory=list)
    exceptions: list[str] = field(default_factory=list)

    @property
    def all_periods(self) -> list[str]:
        return self.historical_periods + self.forecast_periods

    def to_dict(self) -> dict:
        d = dict(self.__dict__)
        d["disclosure"] = self.disclosure.to_dict()
        return d


def _make_forecast_labels(hist: list[str], n: int, suffix: str) -> list[str]:
    if not hist:
        return []
    last = _fy_num(hist[-1])
    return [f"FY{str(last + i)[-2:]}{suffix}" for i in range(1, n + 1)]


def build_spec(db: SourceDatabase, disclosure: DisclosureMap) -> ModelSpec:
    cfg = model_config()
    tax = AccountingTaxonomy()

    hist = order_historical_periods([
        dp.period for dp in db.datapoints
        if dp.statement in (Statement.INCOME_STATEMENT, Statement.BALANCE_SHEET,
                            Statement.CASH_FLOW) and dp.period and dp.period[0:2] == "FY"
    ])
    # cap to target historical years, keep the most recent ones
    target = cfg["model"]["target_historical_years"]
    if len(hist) > target:
        hist = hist[-target:]

    n_forecast = cfg["model"]["forecast_years"]
    forecast = _make_forecast_labels(hist, n_forecast, cfg["model"]["forecast_suffix"])

    # reporting basis
    unit = currency = None
    consolidation = cfg["model"]["default_consolidation"]
    for dp in db.datapoints:
        if dp.unit and not unit:
            unit = dp.unit
        if dp.currency and not currency:
            currency = dp.currency
    if db.documents:
        consolidation = db.documents[-1].consolidation.value or consolidation

    # statement lines: keep taxonomy order, include a line if any component or the
    # line itself is present, OR it is a constructed total whose inputs exist.
    def _present_metrics(statement: Statement) -> set[str]:
        return {dp.metric for dp in db.by_statement(statement) if dp.value is not None}

    is_present = _present_metrics(Statement.INCOME_STATEMENT)
    bs_present = _present_metrics(Statement.BALANCE_SHEET)
    cf_present = _present_metrics(Statement.CASH_FLOW)

    is_lines = _select_lines("income_statement", tax, is_present)
    bs_lines = _select_lines("balance_sheet", tax, bs_present)
    cf_lines = _select_lines("cash_flow", tax, cf_present | {"pat", "depreciation"})

    # schedules
    schedules = []
    if disclosure.has_revenue:
        schedules.append("Revenue Build")
    if disclosure.has_operating_drivers:
        schedules.append("Operating Drivers")
    if disclosure.has_working_capital:
        schedules.append("Working Capital")
    if disclosure.has_capex_or_dna:
        schedules.append("Capex & D&A")
    if disclosure.has_debt:
        schedules.append("Debt")

    # valuation
    valuation = []
    vcfg = cfg.get("valuation", {})
    if vcfg.get("build_dcf", True):
        valuation.append("DCF")
    if vcfg.get("build_trading_comps", True):
        valuation.append("Trading Comparables")
    if vcfg.get("build_sotp_if_segments", True) and disclosure.has_segments:
        valuation.append("SOTP")

    # sheet order
    sheets = ["Cover", "Sources", "Assumptions"]
    if "Revenue Build" in schedules:
        sheets.append("Revenue Build")
    if "Operating Drivers" in schedules:
        sheets.append("Operating Drivers")
    sheets += ["Income Statement", "Balance Sheet", "Cash Flow", "3-Statement Model"]
    for s in ("Working Capital", "Capex & D&A", "Debt"):
        if s in schedules:
            sheets.append(s)
    sheets += ["Ratios", "Valuation"]
    if cfg.get("sensitivities", {}).get("build", True) and valuation:
        sheets.append("Sensitivities")
    sheets.append("Checks")

    exceptions = list(disclosure.notes)
    if not hist:
        exceptions.append("No historical FY periods identified - cannot build model.")
    if len(hist) < 3:
        exceptions.append(f"Only {len(hist)} historical year(s) identified; ideally ~5.")

    return ModelSpec(
        company=db.company,
        currency=currency,
        unit=unit,
        consolidation=consolidation,
        historical_periods=hist,
        forecast_periods=forecast,
        disclosure=disclosure,
        income_statement_lines=is_lines,
        balance_sheet_lines=bs_lines,
        cash_flow_lines=cf_lines,
        schedules=schedules,
        valuation_methods=valuation,
        sheets=sheets,
        exceptions=exceptions,
    )


def _available(statement_key: str, tax: AccountingTaxonomy, metric: str,
               present: set[str], seen: set[str] | None = None) -> bool:
    """A line is available if reported, or if every component of its construction
    is itself (recursively) available - so totals build from reported leaves even
    when the intermediate subtotals are not separately disclosed."""
    seen = seen or set()
    if metric in seen:
        return False
    seen = seen | {metric}
    construct = tax.construction(statement_key, metric)
    if not construct or not isinstance(construct, str) or construct in (
            "reported_or_segment_sum", "link_income_statement"):
        return metric in present
    comps = _tokens(construct)
    if not comps:
        return metric in present
    return all(_available(statement_key, tax, c, present, seen) for c in comps)


def _select_lines(statement_key: str, tax: AccountingTaxonomy, present: set[str]) -> list[str]:
    """Include a line if reported, or if it's a constructible total whose inputs
    are (recursively) available."""
    lines: list[str] = []
    for metric in tax.order(statement_key):
        if metric in present or _available(statement_key, tax, metric, present):
            lines.append(metric)
    return lines


def _tokens(expr: str) -> list[str]:
    for ch in "+-*/()":
        expr = expr.replace(ch, " ")
    return [t for t in expr.split() if t and not t.replace(".", "").isdigit()]
