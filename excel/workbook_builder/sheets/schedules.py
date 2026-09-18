"""Operating and financing schedules.

Revenue Build, Operating Drivers, Working Capital, Capex & D&A and Debt. Each
schedule computes historical drivers by formula from the financial statements
and lets the analyst control the forecast via the Assumptions sheet. Closing
balances feed back into the Balance Sheet / Income Statement / Cash Flow so the
model is genuinely integrated.
"""

from __future__ import annotations

from typing import Optional

from excel.formulas import fw
from excel.formatting import styles
from excel.workbook_builder.grid import CellValue, SheetBuilder
from excel.workbook_builder.sheets.statements import A, _prev
from modules.model_spec import ModelSpec
from modules.schemas import Consolidation, DataType, SourceDatabase, Statement

_CONSOL = Consolidation.CONSOLIDATED


def _segment_metrics(db: SourceDatabase) -> list[tuple[str, str]]:
    seen: list[tuple[str, str]] = []
    for dp in db.by_statement(Statement.SEGMENT):
        if dp.category and "revenue" not in (dp.category or "").lower():
            continue
        key = dp.metric
        label = (dp.label or key).split(" (")[0]
        if (key, label) not in seen:
            seen.append((key, label))
    return seen


# ---------------------------------------------------------------------------
# Revenue build
# ---------------------------------------------------------------------------

def build_revenue(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sheet = "Revenue Build"
    sb = SheetBuilder(sheet, subtitle="Historical from disclosure; forecast off Assumptions "
                                      "(segments grow at the group revenue-growth assumption).")
    hist = set(spec.historical_periods)
    segments = _segment_metrics(db)

    sb.section("Revenue build")

    def seg_writer(metric):
        def w(period, registry):
            if period in hist:
                return fw.hist_input(spec, db, metric, period, registry, _CONSOL)
            p = _prev(spec, period)
            pa = registry.addr(sheet, metric, p)
            g = A(registry, "assum_revenue_growth", period)
            return CellValue(f"={pa}*(1+{g})", DataType.DERIVED) if pa else None
        return w

    for metric, label in segments:
        sb.line(metric, label, seg_writer(metric))

    def revenue_writer(period, registry):
        if segments:
            # explicit additive construction of segment total
            addrs = [registry.addr(sheet, m, period) for m, _ in segments]
            addrs = [a for a in addrs if a]
            if addrs:
                return CellValue("=" + "+".join(addrs), DataType.DERIVED,
                                 comment="Total revenue = sum of disclosed segments")
        if period in hist:
            return fw.hist_input(spec, db, "revenue", period, registry, _CONSOL)
        p = _prev(spec, period)
        pa = registry.addr(sheet, "revenue", p)
        g = A(registry, "assum_revenue_growth", period)
        return CellValue(f"={pa}*(1+{g})", DataType.DERIVED) if pa else None

    sb.total("revenue", "Total revenue", revenue_writer)

    if segments:
        sb.blank()
        sb.section("Segment mix (%)")

        def mix_writer(metric):
            def w(period, registry):
                seg = registry.addr(sheet, metric, period)
                tot = registry.addr(sheet, "revenue", period)
                return CellValue(f"={seg}/{tot}", DataType.DERIVED,
                                 number_format=styles.NF_PERCENT) if seg and tot else None
            return w
        for metric, label in segments:
            sb.line("mix_" + metric, label, mix_writer(metric), number_format=styles.NF_PERCENT)
    return sb


# ---------------------------------------------------------------------------
# Operating drivers
# ---------------------------------------------------------------------------

def build_operating_drivers(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sheet = "Operating Drivers"
    sb = SheetBuilder(sheet, subtitle="Company-specific KPIs. Historical as disclosed; "
                                      "forecast held flat as a placeholder (analyst to revise).")
    hist = set(spec.historical_periods)
    drivers: list[tuple[str, str, str]] = []
    for dp in db.by_statement(Statement.OPERATING):
        if (dp.metric, dp.label or dp.metric, dp.unit or "") not in drivers:
            drivers.append((dp.metric, dp.label or dp.metric, dp.unit or ""))

    sb.section("Operating drivers")

    def w_for(metric):
        def w(period, registry):
            if period in hist:
                return fw.hist_input(spec, db, metric, period, registry, None)
            p = _prev(spec, period)
            pa = registry.addr(sheet, metric, p)
            return CellValue(f"={pa}", DataType.DERIVED,
                             comment="Held flat - analyst forecast input") if pa else None
        return w

    for metric, label, unit in drivers:
        lbl = f"{label}" + (f" ({unit})" if unit else "")
        sb.line(metric, lbl, w_for(metric), number_format=styles.NF_RATIO)
    return sb


# ---------------------------------------------------------------------------
# Working capital
# ---------------------------------------------------------------------------

def build_working_capital(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sheet = "Working Capital"
    sb = SheetBuilder(sheet, subtitle="Historical ratios computed from statements; "
                                      "forecast balances driven by DSO/DIO/DPO assumptions.")
    hist = set(spec.historical_periods)

    def days_writer(bs_metric, is_metric, assum_key):
        def w(period, registry):
            if period in hist:
                bal = registry.ref("Balance Sheet", bs_metric, period)
                base = registry.ref("Income Statement", is_metric, period)
                if bal and base:
                    return CellValue(f"={bal}/{base}*365", DataType.DERIVED,
                                     number_format=styles.NF_DAYS)
                return None
            return fw.link("Assumptions", assum_key, period, registry)
        return w

    def bal_writer(days_key, is_metric):
        def w(period, registry):
            if period in hist:
                return fw.link("Balance Sheet", _bs_of(days_key), period, registry)
            days = registry.addr(sheet, days_key, period)
            base = registry.ref("Income Statement", is_metric, period)
            if days and base:
                return CellValue(f"={days}/365*{base}", DataType.DERIVED)
            return None
        return w

    sb.section("Receivables")
    sb.line("wc_dso", "DSO (days)", days_writer("receivables", "revenue", "assum_dso"),
            number_format=styles.NF_DAYS)
    sb.line("receivables", "Trade receivables", bal_writer("wc_dso", "revenue"))
    sb.blank()
    sb.section("Inventory")
    sb.line("wc_dio", "DIO (days on COGS)", days_writer("inventory", "cogs", "assum_dio"),
            number_format=styles.NF_DAYS)
    sb.line("inventory", "Inventories", bal_writer("wc_dio", "cogs"))
    sb.blank()
    sb.section("Payables")
    sb.line("wc_dpo", "DPO (days on COGS)", days_writer("payables", "cogs", "assum_dpo"),
            number_format=styles.NF_DAYS)
    sb.line("payables", "Trade payables", bal_writer("wc_dpo", "cogs"))
    sb.blank()

    def nwc_writer(period, registry):
        r = registry.addr(sheet, "receivables", period)
        i = registry.addr(sheet, "inventory", period)
        p = registry.addr(sheet, "payables", period)
        if r and i and p:
            return CellValue(f"={r}+{i}-{p}", DataType.DERIVED,
                             comment="Net working capital = receivables + inventory - payables")
        return None
    sb.total("net_working_capital", "Net working capital", nwc_writer)
    return sb


_BS_OF = {"wc_dso": "receivables", "wc_dio": "inventory", "wc_dpo": "payables"}


def _bs_of(days_key: str) -> str:
    return _BS_OF[days_key]


# ---------------------------------------------------------------------------
# Capex & D&A
# ---------------------------------------------------------------------------

def build_capex_dna(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sheet = "Capex & D&A"
    sb = SheetBuilder(sheet, subtitle="PP&E roll-forward. Forecast capex = % of revenue; "
                                      "D&A = rate on opening PP&E (Assumptions).")
    hist = set(spec.historical_periods)
    first_hist = spec.historical_periods[0] if spec.historical_periods else None

    def opening_writer(period, registry):
        p = _prev(spec, period)
        pc = registry.addr(sheet, "closing_ppe", p)
        if pc:
            return CellValue(f"={pc}", DataType.LINKED)
        if period == first_hist:
            # anchor opening to reported prior PP&E if available, else same-year closing
            return fw.link("Balance Sheet", "ppe", period, registry,
                           comment="Opening PP&E anchored to reported balance (first year)")
        return None

    def capex_writer(period, registry):
        if period in hist:
            cv = fw.hist_input(spec, db, "capex", period, registry, _CONSOL)
            if cv:
                return cv
            return fw.link("Cash Flow", "capex", period, registry)
        rev = registry.ref("Income Statement", "revenue", period)
        return CellValue(f"={rev}*{A(registry, 'assum_capex_pct', period)}", DataType.DERIVED) if rev else None

    def dep_writer(period, registry):
        if period in hist:
            return fw.link("Income Statement", "depreciation", period, registry)
        opening = registry.addr(sheet, "opening_ppe", period)
        return CellValue(f"={opening}*{A(registry, 'assum_dep_rate', period)}", DataType.DERIVED) if opening else None

    def closing_writer(period, registry):
        if period in hist:
            return fw.link("Balance Sheet", "ppe", period, registry,
                           comment="Historical closing PP&E = reported balance")
        o = registry.addr(sheet, "opening_ppe", period)
        c = registry.addr(sheet, "capex", period)
        d = registry.addr(sheet, "depreciation", period)
        if o and c and d:
            return CellValue(f"={o}+{c}-{d}", DataType.DERIVED,
                             comment="Closing PP&E = opening + capex - D&A")
        return None

    sb.section("PP&E roll-forward")
    sb.line("opening_ppe", "Opening PP&E", opening_writer)
    sb.line("capex", "Capital expenditure", capex_writer)
    sb.line("depreciation", "Depreciation & amortisation", dep_writer)
    sb.total("closing_ppe", "Closing PP&E", closing_writer)
    return sb


# ---------------------------------------------------------------------------
# Debt
# ---------------------------------------------------------------------------

def build_debt(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sheet = "Debt"
    sb = SheetBuilder(sheet, subtitle="Debt roll-forward. Forecast drawdown/repayment from "
                                      "Assumptions; interest = rate on average debt.")
    hist = set(spec.historical_periods)
    first_hist = spec.historical_periods[0] if spec.historical_periods else None

    def opening_writer(period, registry):
        p = _prev(spec, period)
        pc = registry.addr(sheet, "closing_debt", p)
        if pc:
            return CellValue(f"={pc}", DataType.LINKED)
        if period == first_hist:
            return fw.link("Debt", "closing_debt", period, registry)
        return None

    def closing_writer(period, registry):
        if period in hist:
            std = registry.ref("Balance Sheet", "short_term_debt", period)
            ltd = registry.ref("Balance Sheet", "long_term_debt", period)
            terms = [t for t in (std, ltd) if t]
            if terms:
                return CellValue("=" + "+".join(terms), DataType.LINKED,
                                 comment="Total debt = short-term + long-term (Balance Sheet)")
            return None
        o = registry.addr(sheet, "opening_debt", period)
        dr = registry.addr(sheet, "drawdown", period)
        rp = registry.addr(sheet, "repayment", period)
        if o and dr and rp:
            return CellValue(f"={o}+{dr}-{rp}", DataType.DERIVED,
                             comment="Closing debt = opening + drawdown - repayment")
        return None

    def drawdown_writer(period, registry):
        if period in hist:
            cv = fw.hist_input(spec, db, "debt_raised", period, registry, _CONSOL)
            return cv if cv else CellValue(0, DataType.REPORTED)
        return fw.link("Assumptions", "assum_debt_drawdown", period, registry)

    def repayment_writer(period, registry):
        if period in hist:
            cv = fw.hist_input(spec, db, "debt_repaid", period, registry, _CONSOL)
            return cv if cv else CellValue(0, DataType.REPORTED)
        return fw.link("Assumptions", "assum_debt_repayment", period, registry)

    def interest_writer(period, registry):
        if period in hist:
            return fw.link("Income Statement", "interest_expense", period, registry)
        o = registry.addr(sheet, "opening_debt", period)
        c = registry.addr(sheet, "closing_debt", period)
        rate = A(registry, "assum_interest_rate", period)
        if o and c:
            return CellValue(f"=AVERAGE({o},{c})*{rate}", DataType.DERIVED,
                             comment="Interest = rate x average debt")
        return None

    sb.section("Debt roll-forward")
    sb.line("opening_debt", "Opening debt", opening_writer)
    sb.line("drawdown", "New drawdown", drawdown_writer)
    sb.line("repayment", "Repayment", repayment_writer)
    sb.total("closing_debt", "Closing debt", closing_writer)
    sb.blank()
    sb.line("interest", "Interest expense", interest_writer)
    return sb
