"""Assumptions sheet - the analyst's control panel.

Every forecast in the model is driven by a yellow input cell here. The agent
seeds each input with a NEUTRAL placeholder (the latest historical actual /
ratio, i.e. a hold-flat default) purely so the workbook computes on open. These
are explicitly NOT the agent's forecast - each cell is yellow, carries a note,
and the analyst is expected to overwrite it. The agent never decides the
business outlook.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional

from excel.formatting import styles
from excel.workbook_builder.grid import CellValue, SheetBuilder
from modules.model_spec import ModelSpec
from modules.schemas import Consolidation, DataType, SourceDatabase


@dataclass
class Assumption:
    key: str
    label: str
    number_format: str
    seed: Callable[[SourceDatabase, ModelSpec], Optional[float]]
    note: str = ""


def _last_two_hist(spec: ModelSpec) -> tuple[Optional[str], Optional[str]]:
    h = spec.historical_periods
    if len(h) >= 2:
        return h[-2], h[-1]
    if len(h) == 1:
        return None, h[-1]
    return None, None


def _ratio(db, num, den, period, consol) -> Optional[float]:
    a = db.value(num, period, consol)
    b = db.value(den, period, consol)
    if a is None or b in (None, 0):
        return None
    return a / b


_CONSOL = Consolidation.CONSOLIDATED


def _seed_rev_growth(db, spec):
    p, c = _last_two_hist(spec)
    if not p or not c:
        return 0.0
    r0 = db.value("revenue", p, _CONSOL)
    r1 = db.value("revenue", c, _CONSOL)
    if r0 in (None, 0) or r1 is None:
        return 0.0
    return round(r1 / r0 - 1.0, 4)


def _seed_gross_margin(db, spec):
    _, c = _last_two_hist(spec)
    gp = _ratio(db, "gross_profit", "revenue", c, _CONSOL)
    if gp is not None:
        return round(gp, 4)
    cogs = db.value("cogs", c, _CONSOL)
    rev = db.value("revenue", c, _CONSOL)
    if cogs is not None and rev not in (None, 0):
        return round(1 - cogs / rev, 4)
    return 0.30


def _seed_pct_rev(metric):
    def _fn(db, spec):
        _, c = _last_two_hist(spec)
        r = _ratio(db, metric, "revenue", c, _CONSOL)
        return round(r, 4) if r is not None else 0.05
    return _fn


def _seed_tax_rate(db, spec):
    _, c = _last_two_hist(spec)
    r = _ratio(db, "tax", "pbt", c, _CONSOL)
    return round(r, 4) if r is not None else 0.25


def _seed_days(num, den):
    def _fn(db, spec):
        _, c = _last_two_hist(spec)
        r = _ratio(db, num, den, c, _CONSOL)
        return round(r * 365, 1) if r is not None else 45.0
    return _fn


def _seed_capex_pct(db, spec):
    _, c = _last_two_hist(spec)
    r = _ratio(db, "capex", "revenue", c, _CONSOL)
    return round(abs(r), 4) if r is not None else 0.05


def _seed_dep_rate(db, spec):
    _, c = _last_two_hist(spec)
    r = _ratio(db, "depreciation", "ppe", c, _CONSOL)
    return round(r, 4) if r is not None else 0.10


def _seed_interest_rate(db, spec):
    _, c = _last_two_hist(spec)
    debt = None
    std = db.value("short_term_debt", c, _CONSOL) or 0
    ltd = db.value("long_term_debt", c, _CONSOL) or 0
    debt = std + ltd
    intr = db.value("interest_expense", c, _CONSOL)
    if intr is not None and debt:
        return round(intr / debt, 4)
    return 0.08


ASSUMPTIONS: list[Assumption] = [
    Assumption("assum_revenue_growth", "Revenue growth (YoY)", styles.NF_PERCENT, _seed_rev_growth,
               "Seeded = last historical YoY growth. Analyst to revise."),
    Assumption("assum_gross_margin", "Gross margin", styles.NF_PERCENT, _seed_gross_margin,
               "Seeded = last historical gross margin."),
    Assumption("assum_employee_pct", "Employee expense (% revenue)", styles.NF_PERCENT,
               _seed_pct_rev("employee_expense")),
    Assumption("assum_other_opex_pct", "Other opex (% revenue)", styles.NF_PERCENT,
               _seed_pct_rev("other_opex")),
    Assumption("assum_other_income_growth", "Other income growth", styles.NF_PERCENT,
               lambda db, spec: 0.0, "Seeded flat. Analyst to revise."),
    Assumption("assum_tax_rate", "Effective tax rate", styles.NF_PERCENT, _seed_tax_rate),
    Assumption("assum_capex_pct", "Capex (% revenue)", styles.NF_PERCENT, _seed_capex_pct),
    Assumption("assum_dep_rate", "Depreciation rate (% opening PP&E)", styles.NF_PERCENT, _seed_dep_rate),
    Assumption("assum_dso", "DSO (days)", styles.NF_DAYS, _seed_days("receivables", "revenue")),
    Assumption("assum_dio", "DIO (days on COGS)", styles.NF_DAYS, _seed_days("inventory", "cogs")),
    Assumption("assum_dpo", "DPO (days on COGS)", styles.NF_DAYS, _seed_days("payables", "cogs")),
    Assumption("assum_interest_rate", "Interest rate on debt", styles.NF_PERCENT, _seed_interest_rate),
    Assumption("assum_debt_drawdown", "New debt drawdown", styles.NF_CURRENCY,
               lambda db, spec: 0.0, "Seeded 0. Analyst to revise."),
    Assumption("assum_debt_repayment", "Debt repayment", styles.NF_CURRENCY,
               lambda db, spec: 0.0, "Seeded 0. Analyst to revise."),
    Assumption("assum_dividend_payout", "Dividend payout (% PAT)", styles.NF_PERCENT,
               lambda db, spec: 0.0, "Seeded 0. Analyst to revise."),
]

# Single-value valuation assumptions (not per-year)
VALUATION_ASSUMPTIONS = [
    Assumption("assum_wacc", "WACC", styles.NF_PERCENT, lambda db, spec: 0.11,
               "Placeholder. Analyst decides discount rate."),
    Assumption("assum_terminal_growth", "Terminal growth", styles.NF_PERCENT,
               lambda db, spec: 0.04, "Placeholder. Analyst decides."),
    Assumption("assum_target_multiple", "Target EV/EBITDA (x)", styles.NF_MULTIPLE,
               lambda db, spec: 10.0, "Placeholder. Analyst decides."),
]


def build_assumptions(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sb = SheetBuilder("Assumptions", periods_row=True,
                      subtitle="Yellow cells are USER INPUTS. Seeded values are neutral "
                               "placeholders (hold-last-actual) - analyst must review.")
    sb.section("Forecast assumptions (analyst-controlled)")

    forecast = set(spec.forecast_periods)

    def make_writer(a: Assumption):
        seed = a.seed(db, spec)

        def writer(period, registry):
            if period in forecast:
                return CellValue(content=seed if seed is not None else 0.0,
                                 data_type=DataType.ASSUMPTION,
                                 number_format=a.number_format,
                                 comment=(a.note or "User assumption - overwrite as needed."))
            return None  # historical columns are blank on the assumptions sheet
        return writer

    for a in ASSUMPTIONS:
        sb.line(a.key, a.label, make_writer(a), number_format=a.number_format)

    sb.blank()
    sb.section("Valuation assumptions (analyst-controlled)")

    def make_val_writer(a: Assumption):
        seed = a.seed(db, spec)
        first_forecast = spec.forecast_periods[0] if spec.forecast_periods else None

        def writer(period, registry):
            if period == first_forecast:
                return CellValue(content=seed, data_type=DataType.ASSUMPTION,
                                 number_format=a.number_format, comment=a.note)
            return None
        return writer

    for a in VALUATION_ASSUMPTIONS:
        sb.line(a.key, a.label, make_val_writer(a), number_format=a.number_format)

    sb.blank()
    sb.note("Colour legend: BLUE = reported input | BLACK = formula | GREEN = link to "
            "another sheet | YELLOW = user assumption | RED = failed check.")
    return sb


def assum_addr(registry, key: str, period: str) -> str:
    """Absolute reference to a per-year assumption cell (single-value assumptions
    use the first forecast period)."""
    ref = registry.ref("Assumptions", key, period)
    return ref
