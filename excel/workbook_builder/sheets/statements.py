"""Income Statement, Balance Sheet and Cash Flow sheets.

Historical columns are built from source inputs (blue) and explicit additive
construction formulas (black) - never SUM() for constructive totals, and only
constructed when the components are actually disclosed (else the reported total
is used as a blue input, per the accuracy-over-construction rule).

Forecast columns are entirely formula-driven off the Assumptions sheet and the
operating schedules, and are internally integrated so the balance sheet balances
by construction (cash is the plug fed from the cash-flow statement).
"""

from __future__ import annotations

from typing import Optional

from excel.formulas import fw
from excel.formatting import styles
from excel.workbook_builder.grid import CellValue, SheetBuilder
from modules.config_loader import AccountingTaxonomy
from modules.model_spec import ModelSpec
from modules.schemas import Consolidation, DataType, SourceDatabase

_CONSOL = Consolidation.CONSOLIDATED
_TAX = AccountingTaxonomy()


def _prev(spec: ModelSpec, period: str) -> Optional[str]:
    ap = spec.all_periods
    i = ap.index(period)
    return ap[i - 1] if i > 0 else None


def A(registry, key: str, period: str) -> Optional[str]:
    return registry.ref("Assumptions", key, period)


def _components(statement_key, metric) -> list[str]:
    expr = _TAX.construction(statement_key, metric)
    if not expr or not isinstance(expr, str) or expr in (
            "reported_or_segment_sum", "link_income_statement"):
        return []
    return [t for t in fw._tokenize(expr) if t not in "+-*/()" and not fw._is_number(t)]


def _resolvable(db, statement_key, metric, period, _seen=None) -> bool:
    """Can this metric be established for the period from disclosed data?

    A leaf metric is resolvable if it is reported (present in the source DB).
    A constructed metric is resolvable if all of its components are resolvable.
    This lets totals (EBITDA, EBIT, PBT, PAT, subtotals) be built by formula
    from reported leaves even though the intermediate sub-totals are not in the
    source DB themselves.
    """
    _seen = _seen or set()
    if metric in _seen:
        return False
    _seen = _seen | {metric}
    comps = _components(statement_key, metric)
    if not comps:
        return db.value(metric, period, _CONSOL) is not None
    return all(_resolvable(db, statement_key, c, period, _seen) for c in comps)


def _construct_or_report(db, statement_key, metric, sheet, period, registry) -> Optional[CellValue]:
    comps = _components(statement_key, metric)
    if comps and all(_resolvable(db, statement_key, c, period) for c in comps):
        cv = fw.construct_cv(_TAX.construction(statement_key, metric), sheet, period, registry)
        if cv:
            return cv
    return fw.reported(db, metric, period, _CONSOL)


# ---------------------------------------------------------------------------
# Income statement
# ---------------------------------------------------------------------------

def build_income_statement(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sheet = "Income Statement"
    sb = SheetBuilder(sheet, subtitle=f"{spec.currency or ''} {spec.unit or ''}".strip())
    hist = set(spec.historical_periods)
    has_revbuild = "Revenue Build" in spec.schedules
    has_capex = "Capex & D&A" in spec.schedules
    has_debt = "Debt" in spec.schedules
    lines = spec.income_statement_lines

    sb.section("Income statement")

    def writer_for(metric):
        def writer(period, registry):
            if period in hist:
                if metric == "revenue" and has_revbuild:
                    return fw.link("Revenue Build", "revenue", period, registry,
                                   comment="Revenue linked from Revenue Build")
                return _construct_or_report(db, "income_statement", metric, sheet, period, registry)
            # ---- forecast ----
            p = _prev(spec, period)
            rev = registry.addr(sheet, "revenue", period)
            if metric == "revenue":
                if has_revbuild:
                    return fw.link("Revenue Build", "revenue", period, registry)
                return fw.growth_formula(sheet, "revenue", p, A(registry, "assum_revenue_growth", period), registry)
            if metric == "cogs" and rev:
                return CellValue(f"={rev}*(1-{A(registry, 'assum_gross_margin', period)})", DataType.DERIVED)
            if metric == "employee_expense" and rev:
                return CellValue(f"={rev}*{A(registry, 'assum_employee_pct', period)}", DataType.DERIVED)
            if metric == "other_opex" and rev:
                return CellValue(f"={rev}*{A(registry, 'assum_other_opex_pct', period)}", DataType.DERIVED)
            if metric == "depreciation":
                if has_capex:
                    return fw.link("Capex & D&A", "depreciation", period, registry)
                pa = registry.addr(sheet, "depreciation", p)
                return CellValue(f"={pa}", DataType.DERIVED) if pa else None
            if metric == "other_income":
                pa = registry.addr(sheet, "other_income", p)
                return CellValue(f"={pa}*(1+{A(registry, 'assum_other_income_growth', period)})",
                                 DataType.DERIVED) if pa else None
            if metric == "interest_expense":
                if has_debt:
                    return fw.link("Debt", "interest", period, registry)
                pa = registry.addr(sheet, "interest_expense", p)
                return CellValue(f"={pa}", DataType.DERIVED) if pa else None
            if metric == "tax":
                pbt = registry.addr(sheet, "pbt", period)
                return CellValue(f"={pbt}*{A(registry, 'assum_tax_rate', period)}", DataType.DERIVED) if pbt else None
            if metric == "eps":
                pat_t = registry.addr(sheet, "pat", period)
                pat_p = registry.addr(sheet, "pat", p)
                eps_p = registry.addr(sheet, "eps", p)
                if pat_t and pat_p and eps_p:
                    return CellValue(f"={pat_t}*({eps_p}/{pat_p})", DataType.DERIVED)
                return None
            # constructed totals (gross_profit, ebitda, ebit, pbt, pat)
            expr = _TAX.construction("income_statement", metric)
            if expr and isinstance(expr, str):
                cv = fw.construct_cv(expr, sheet, period, registry)
                if cv:
                    return cv
            return None
        return writer

    for metric in lines:
        label = _TAX.label("income_statement", metric)
        is_total = metric in ("gross_profit", "ebitda", "ebit", "pbt", "pat")
        nf = styles.NF_RATIO if metric == "eps" else styles.NF_CURRENCY
        sb.line(metric, label, writer_for(metric), number_format=nf,
                kind="total" if is_total else "line", bold=is_total)
    return sb


# ---------------------------------------------------------------------------
# Balance sheet
# ---------------------------------------------------------------------------

def build_balance_sheet(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sheet = "Balance Sheet"
    sb = SheetBuilder(sheet, subtitle=f"{spec.currency or ''} {spec.unit or ''}".strip())
    hist = set(spec.historical_periods)
    has_wc = "Working Capital" in spec.schedules
    has_capex = "Capex & D&A" in spec.schedules
    has_debt = "Debt" in spec.schedules
    lines = spec.balance_sheet_lines

    totals = {"total_current_assets", "total_non_current_assets", "total_assets",
              "total_current_liabilities", "total_non_current_liabilities",
              "total_liabilities", "total_equity", "total_equity_and_liabilities"}

    def flat(metric):
        def w(period, registry):
            p = _prev(spec, period)
            pa = registry.addr(sheet, metric, p)
            return CellValue(f"={pa}", DataType.DERIVED) if pa else None
        return w

    def writer_for(metric):
        def writer(period, registry):
            if period in hist:
                return _construct_or_report(db, "balance_sheet", metric, sheet, period, registry)
            p = _prev(spec, period)
            # forecast
            if metric in totals:
                expr = _TAX.construction("balance_sheet", metric)
                cv = fw.construct_cv(expr, sheet, period, registry) if expr else None
                return cv
            if metric == "cash":
                # plug: closing cash from the cash flow statement
                return fw.link("Cash Flow", "closing_cash", period, registry,
                               comment="Cash plug linked from Cash Flow")
            if metric == "receivables" and has_wc:
                return fw.link("Working Capital", "receivables", period, registry)
            if metric == "inventory" and has_wc:
                return fw.link("Working Capital", "inventory", period, registry)
            if metric == "payables" and has_wc:
                return fw.link("Working Capital", "payables", period, registry)
            if metric == "ppe" and has_capex:
                return fw.link("Capex & D&A", "closing_ppe", period, registry)
            if metric == "long_term_debt" and has_debt:
                closing = registry.ref("Debt", "closing_debt", period)
                std = registry.addr(sheet, "short_term_debt", period)
                if closing and std:
                    return CellValue(f"={closing}-{std}", DataType.LINKED,
                                     comment="Long-term debt = Debt schedule closing less short-term")
                return fw.link("Debt", "closing_debt", period, registry)
            if metric == "reserves":
                pa = registry.addr(sheet, "reserves", p)
                pat = registry.ref("Income Statement", "pat", period)
                div = registry.ref("Cash Flow", "dividends_paid", period)
                if pa and pat:
                    f = f"={pa}+{pat}"
                    if div:
                        f += f"-{div}"
                    return CellValue(f, DataType.DERIVED,
                                     comment="Reserves roll-forward: opening + PAT - dividends")
                return flat("reserves")(period, registry)
            # default: hold flat (change captured in CF as zero)
            return flat(metric)(period, registry)
        return writer

    sb.section("Assets")
    _emit(sb, ["cash", "receivables", "inventory", "other_current_assets",
               "total_current_assets", "ppe", "intangibles", "other_non_current_assets",
               "total_non_current_assets", "total_assets"], lines, totals, writer_for, "balance_sheet")
    sb.blank()
    sb.section("Liabilities")
    _emit(sb, ["payables", "short_term_debt", "other_current_liabilities",
               "total_current_liabilities", "long_term_debt", "other_non_current_liabilities",
               "total_non_current_liabilities", "total_liabilities"], lines, totals, writer_for, "balance_sheet")
    sb.blank()
    sb.section("Equity")
    _emit(sb, ["share_capital", "reserves", "total_equity",
               "total_equity_and_liabilities"], lines, totals, writer_for, "balance_sheet")
    return sb


def _emit(sb, group, lines, totals, writer_for, statement_key):
    for metric in group:
        if metric not in lines:
            continue
        label = _TAX.label(statement_key, metric)
        is_total = metric in totals
        sb.line(metric, label, writer_for(metric),
                kind="total" if is_total else "line", bold=is_total)


# ---------------------------------------------------------------------------
# Cash flow
# ---------------------------------------------------------------------------

def build_cash_flow(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sheet = "Cash Flow"
    sb = SheetBuilder(sheet, subtitle=f"{spec.currency or ''} {spec.unit or ''}".strip())
    hist = set(spec.historical_periods)
    has_wc = "Working Capital" in spec.schedules
    has_capex = "Capex & D&A" in spec.schedules
    has_debt = "Debt" in spec.schedules
    lines = spec.cash_flow_lines

    totals = {"cfo", "cfi", "cff", "net_change_in_cash", "closing_cash"}

    def bs_delta(metric, period, registry, sign):
        """sign * (prior - current) for asset WC change to CFO (asset up = cash down)."""
        p = _prev(spec, period)
        cur = registry.ref("Balance Sheet", metric, period)
        pri = registry.ref("Balance Sheet", metric, p)
        if cur and pri:
            return f"({pri}-{cur})" if sign > 0 else f"({cur}-{pri})"
        return None

    def writer_for(metric):
        def writer(period, registry):
            p = _prev(spec, period)
            if period in hist:
                # PAT and D&A always link to the income statement
                if metric in ("pat", "depreciation"):
                    return fw.link("Income Statement", metric, period, registry)
                # Opening cash: prior-year closing, else the disclosed opening / first-year cash
                if metric == "opening_cash":
                    pc = registry.addr(sheet, "closing_cash", p)
                    if pc:
                        return CellValue(f"={pc}", DataType.LINKED)
                    if db.get("opening_cash", period, _CONSOL) and \
                            db.value("opening_cash", period, _CONSOL) is not None:
                        return fw.reported(db, "opening_cash", period, _CONSOL)
                    return fw.reported(db, "cash", period, _CONSOL)
                # Closing cash: opening + net change (roll-forward)
                if metric == "closing_cash":
                    oc = registry.addr(sheet, "opening_cash", period)
                    nc = registry.addr(sheet, "net_change_in_cash", period)
                    if oc and nc:
                        return CellValue(f"={oc}+{nc}", DataType.DERIVED)
                    return _construct_or_report(db, "cash_flow", metric, sheet, period, registry)
                # Net change: reported if disclosed, else CFO+CFI+CFF (subtotals below)
                if metric == "net_change_in_cash":
                    if db.value("net_change_in_cash", period, _CONSOL) is not None:
                        return fw.reported(db, "net_change_in_cash", period, _CONSOL)
                    return fw.construct_cv("cfo + cfi + cff", sheet, period, registry)
                # Subtotals cfo/cfi/cff and every line: prefer the reported total,
                # construct only when all components are disclosed (accuracy first).
                return _construct_or_report(db, "cash_flow", metric, sheet, period, registry)
            # ---- forecast ----
            if metric == "pat":
                return fw.link("Income Statement", "pat", period, registry)
            if metric == "depreciation":
                return fw.link("Income Statement", "depreciation", period, registry)
            if metric == "cf_working_capital_change":
                terms = []
                if has_wc or True:
                    r = bs_delta("receivables", period, registry, +1)
                    i = bs_delta("inventory", period, registry, +1)
                    pay = bs_delta("payables", period, registry, -1)
                    for t in (r, i, pay):
                        if t:
                            terms.append(t)
                if terms:
                    return CellValue("=" + "+".join(terms), DataType.DERIVED,
                                     comment="Change in core working capital (recv, inv, payables)")
                return CellValue(0, DataType.DERIVED)
            if metric == "cf_other_operating":
                return CellValue(0, DataType.ASSUMPTION,
                                 comment="Other operating adjustments (analyst input)")
            if metric == "cfo":
                return fw.construct_cv(_TAX.construction("cash_flow", "cfo"), sheet, period, registry)
            if metric == "capex":
                if has_capex:
                    return fw.link("Capex & D&A", "capex", period, registry)
                rev = registry.ref("Income Statement", "revenue", period)
                return CellValue(f"={rev}*{A(registry, 'assum_capex_pct', period)}", DataType.DERIVED) if rev else None
            if metric == "cf_other_investing":
                return CellValue(0, DataType.ASSUMPTION, comment="Other investing (analyst input)")
            if metric == "cfi":
                return fw.construct_cv(_TAX.construction("cash_flow", "cfi"), sheet, period, registry)
            if metric == "debt_raised":
                return CellValue(f"={A(registry, 'assum_debt_drawdown', period)}", DataType.DERIVED)
            if metric == "debt_repaid":
                return CellValue(f"={A(registry, 'assum_debt_repayment', period)}", DataType.DERIVED)
            if metric == "dividends_paid":
                pat = registry.addr("Income Statement", "pat", period)
                pat = registry.ref("Income Statement", "pat", period)
                return CellValue(f"={pat}*{A(registry, 'assum_dividend_payout', period)}", DataType.DERIVED) if pat else None
            if metric == "cf_other_financing":
                return CellValue(0, DataType.ASSUMPTION, comment="Other financing (analyst input)")
            if metric == "cff":
                return fw.construct_cv(_TAX.construction("cash_flow", "cff"), sheet, period, registry)
            if metric == "net_change_in_cash":
                return fw.construct_cv(_TAX.construction("cash_flow", "net_change_in_cash"), sheet, period, registry)
            if metric == "opening_cash":
                # Forecast opening cash anchors to the prior period's reported /
                # forecast balance-sheet cash, so the forecast ties exactly even
                # when historical CF carries disclosed non-cash reconciling items.
                bs_cash_prev = registry.ref("Balance Sheet", "cash", p)
                if bs_cash_prev:
                    return CellValue(f"={bs_cash_prev}", DataType.LINKED,
                                     comment="Opening cash = prior-year balance-sheet cash")
                pc = registry.addr(sheet, "closing_cash", p)
                return CellValue(f"={pc}", DataType.LINKED) if pc else None
            if metric == "closing_cash":
                oc = registry.addr(sheet, "opening_cash", period)
                nc = registry.addr(sheet, "net_change_in_cash", period)
                return CellValue(f"={oc}+{nc}", DataType.DERIVED) if oc and nc else None
            return None
        return writer

    sb.section("Operating activities")
    for metric in ["pat", "depreciation", "cf_working_capital_change", "cf_other_operating", "cfo"]:
        _cf_line(sb, metric, lines, totals, writer_for)
    sb.blank()
    sb.section("Investing activities")
    for metric in ["capex", "cf_other_investing", "cfi"]:
        _cf_line(sb, metric, lines, totals, writer_for)
    sb.blank()
    sb.section("Financing activities")
    for metric in ["debt_raised", "debt_repaid", "dividends_paid", "cf_other_financing", "cff"]:
        _cf_line(sb, metric, lines, totals, writer_for)
    sb.blank()
    sb.section("Reconciliation")
    for metric in ["net_change_in_cash", "opening_cash", "closing_cash"]:
        _cf_line(sb, metric, lines, totals, writer_for, force=True)
    return sb


def _cf_line(sb, metric, lines, totals, writer_for, force=False):
    if metric not in lines and metric not in ("pat", "depreciation", "cf_working_capital_change",
                                              "cf_other_operating", "cfo", "cf_other_investing",
                                              "cfi", "debt_raised", "debt_repaid", "dividends_paid",
                                              "cf_other_financing", "cff", "net_change_in_cash",
                                              "opening_cash", "closing_cash"):
        if not force:
            return
    label = _TAX.label("cash_flow", metric)
    is_total = metric in totals
    sb.line(metric, label, writer_for(metric),
            kind="total" if is_total else "line", bold=is_total)
