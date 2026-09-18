"""3-Statement Model (integrated layer) and Ratios sheet.

The 3-Statement Model is not a duplicate of the three statements: every line is
a green link into the statements/schedules, arranged to show the integration
(P&L -> cash flow -> balance sheet, capex/D&A -> PP&E, debt -> interest). The
Ratios sheet is entirely formula-driven off the statements.
"""

from __future__ import annotations

from typing import Optional

from excel.formulas import fw
from excel.formatting import styles
from excel.workbook_builder.grid import CellValue, SheetBuilder
from excel.workbook_builder.sheets.statements import _prev
from modules.model_spec import ModelSpec
from modules.schemas import DataType, SourceDatabase


def build_three_statement(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sheet = "3-Statement Model"
    sb = SheetBuilder(sheet, subtitle="Integrated layer - every line links to the statements "
                                      "and schedules (no duplicated hardcodes).")
    is_lines = set(spec.income_statement_lines)
    bs_lines = set(spec.balance_sheet_lines)
    has_debt = "Debt" in spec.schedules

    def link_is(m):
        return lambda period, registry: fw.link("Income Statement", m, period, registry)

    def link_bs(m):
        return lambda period, registry: fw.link("Balance Sheet", m, period, registry)

    sb.section("Profit & loss")
    for m, lbl, tot in [("revenue", "Revenue", False), ("ebitda", "EBITDA", True),
                        ("ebit", "EBIT", True), ("depreciation", "D&A", False),
                        ("pat", "PAT", True)]:
        if m in is_lines:
            sb.line(m, lbl, link_is(m), kind="total" if tot else "line", bold=tot)

    sb.blank()
    sb.section("Cash flow")
    sb.line("cfo", "Cash flow from operations", lambda p, r: fw.link("Cash Flow", "cfo", p, r), bold=True)
    sb.line("capex", "Capital expenditure", lambda p, r: fw.link("Cash Flow", "capex", p, r))

    def fcf_writer(period, registry):
        cfo = registry.addr(sheet, "cfo", period)
        capex = registry.addr(sheet, "capex", period)
        if cfo and capex:
            return CellValue(f"={cfo}-{capex}", DataType.DERIVED,
                             comment="Free cash flow = CFO - capex")
        return None
    sb.total("fcf", "Free cash flow", fcf_writer)

    sb.blank()
    sb.section("Balance sheet")
    if "cash" in bs_lines:
        sb.line("cash", "Cash & equivalents", link_bs("cash"))
    if "ppe" in bs_lines:
        sb.line("ppe", "PP&E", link_bs("ppe"))

    def debt_writer(period, registry):
        if has_debt:
            return fw.link("Debt", "closing_debt", period, registry)
        std = registry.ref("Balance Sheet", "short_term_debt", period)
        ltd = registry.ref("Balance Sheet", "long_term_debt", period)
        terms = [t for t in (std, ltd) if t]
        return CellValue("=" + "+".join(terms), DataType.LINKED) if terms else None
    sb.line("total_debt", "Total debt", debt_writer)

    def netdebt_writer(period, registry):
        d = registry.addr(sheet, "total_debt", period)
        c = registry.addr(sheet, "cash", period)
        if d and c:
            return CellValue(f"={d}-{c}", DataType.DERIVED, comment="Net debt = total debt - cash")
        return None
    sb.line("net_debt", "Net debt", netdebt_writer)
    if "total_equity" in bs_lines:
        sb.line("total_equity", "Total equity", link_bs("total_equity"), bold=True)

    sb.blank()
    sb.section("Integration check")

    def bscheck_writer(period, registry):
        ta = registry.ref("Balance Sheet", "total_assets", period)
        tel = registry.ref("Balance Sheet", "total_equity_and_liabilities", period)
        if ta and tel:
            return CellValue(f"={ta}-{tel}", DataType.DERIVED,
                             comment="Assets - (Equity + Liabilities); should be ~0")
        return None
    sb.line("bs_check", "Balance check (A - E&L)", bscheck_writer)
    return sb


# ---------------------------------------------------------------------------
# Ratios
# ---------------------------------------------------------------------------

def build_ratios(spec: ModelSpec, db: SourceDatabase) -> SheetBuilder:
    sheet = "Ratios"
    sb = SheetBuilder(sheet, subtitle="All ratios are formula-driven off the financial statements.")
    is_lines = set(spec.income_statement_lines)
    bs_lines = set(spec.balance_sheet_lines)

    def growth(m):
        def w(period, registry):
            p = _prev(spec, period)
            cur = registry.ref("Income Statement", m, period)
            pri = registry.ref("Income Statement", m, p)
            if cur and pri:
                return CellValue(f"={cur}/{pri}-1", DataType.DERIVED, number_format=styles.NF_PERCENT)
            return None
        return w

    def margin(m):
        def w(period, registry):
            num = registry.ref("Income Statement", m, period)
            den = registry.ref("Income Statement", "revenue", period)
            if num and den:
                return CellValue(f"={num}/{den}", DataType.DERIVED, number_format=styles.NF_PERCENT)
            return None
        return w

    sb.section("Growth")
    for m, lbl in [("revenue", "Revenue growth"), ("ebitda", "EBITDA growth"),
                   ("ebit", "EBIT growth"), ("pat", "PAT growth")]:
        if m in is_lines:
            sb.line("g_" + m, lbl, growth(m), number_format=styles.NF_PERCENT)

    sb.blank()
    sb.section("Margins")
    for m, lbl in [("gross_profit", "Gross margin"), ("ebitda", "EBITDA margin"),
                   ("ebit", "EBIT margin"), ("pat", "PAT margin")]:
        if m in is_lines:
            sb.line("m_" + m, lbl, margin(m), number_format=styles.NF_PERCENT)

    sb.blank()
    sb.section("Returns")

    def roe(period, registry):
        pat = registry.ref("Income Statement", "pat", period)
        eq = registry.ref("Balance Sheet", "total_equity", period)
        if pat and eq:
            return CellValue(f"={pat}/{eq}", DataType.DERIVED, number_format=styles.NF_PERCENT)
        return None

    def roce(period, registry):
        ebit = registry.ref("Income Statement", "ebit", period)
        eq = registry.ref("Balance Sheet", "total_equity", period)
        std = registry.ref("Balance Sheet", "short_term_debt", period)
        ltd = registry.ref("Balance Sheet", "long_term_debt", period)
        cap_terms = [t for t in (eq, std, ltd) if t]
        if ebit and eq and len(cap_terms) >= 1:
            return CellValue(f"={ebit}/({'+'.join(cap_terms)})", DataType.DERIVED,
                             number_format=styles.NF_PERCENT)
        return None
    if "pat" in is_lines and "total_equity" in bs_lines:
        sb.line("roe", "ROE", roe, number_format=styles.NF_PERCENT)
    if "ebit" in is_lines and "total_equity" in bs_lines:
        sb.line("roce", "ROCE", roce, number_format=styles.NF_PERCENT)

    sb.blank()
    sb.section("Leverage")

    def de(period, registry):
        eq = registry.ref("Balance Sheet", "total_equity", period)
        std = registry.ref("Balance Sheet", "short_term_debt", period)
        ltd = registry.ref("Balance Sheet", "long_term_debt", period)
        terms = [t for t in (std, ltd) if t]
        if eq and terms:
            return CellValue(f"=({'+'.join(terms)})/{eq}", DataType.DERIVED, number_format=styles.NF_RATIO)
        return None

    def nd_ebitda(period, registry):
        nd = registry.ref("3-Statement Model", "net_debt", period)
        ebitda = registry.ref("Income Statement", "ebitda", period)
        if nd and ebitda:
            return CellValue(f"={nd}/{ebitda}", DataType.DERIVED, number_format=styles.NF_MULTIPLE)
        return None
    if "total_equity" in bs_lines:
        sb.line("de", "Debt / Equity", de, number_format=styles.NF_RATIO)
    if "ebitda" in is_lines:
        sb.line("nd_ebitda", "Net debt / EBITDA", nd_ebitda, number_format=styles.NF_MULTIPLE)

    sb.blank()
    sb.section("Working capital")
    for key, lbl in [("wc_dso", "DSO (days)"), ("wc_dio", "DIO (days)"), ("wc_dpo", "DPO (days)")]:
        if "Working Capital" in spec.schedules:
            sb.line("r_" + key, lbl, _wc_link(key), number_format=styles.NF_DAYS)

    def ccc(period, registry):
        dso = registry.ref("Working Capital", "wc_dso", period)
        dio = registry.ref("Working Capital", "wc_dio", period)
        dpo = registry.ref("Working Capital", "wc_dpo", period)
        if dso and dio and dpo:
            return CellValue(f"={dso}+{dio}-{dpo}", DataType.DERIVED, number_format=styles.NF_DAYS)
        return None
    if "Working Capital" in spec.schedules:
        sb.line("ccc", "Cash conversion cycle", ccc, number_format=styles.NF_DAYS)

    sb.blank()
    sb.section("Cash flow")

    def cfo_ebitda(period, registry):
        cfo = registry.ref("Cash Flow", "cfo", period)
        ebitda = registry.ref("Income Statement", "ebitda", period)
        if cfo and ebitda:
            return CellValue(f"={cfo}/{ebitda}", DataType.DERIVED, number_format=styles.NF_PERCENT)
        return None

    def fcf_margin(period, registry):
        fcf = registry.ref("3-Statement Model", "fcf", period)
        rev = registry.ref("Income Statement", "revenue", period)
        if fcf and rev:
            return CellValue(f"={fcf}/{rev}", DataType.DERIVED, number_format=styles.NF_PERCENT)
        return None

    def capex_rev(period, registry):
        capex = registry.ref("Cash Flow", "capex", period)
        rev = registry.ref("Income Statement", "revenue", period)
        if capex and rev:
            return CellValue(f"={capex}/{rev}", DataType.DERIVED, number_format=styles.NF_PERCENT)
        return None
    if "ebitda" in is_lines:
        sb.line("cfo_ebitda", "CFO / EBITDA", cfo_ebitda, number_format=styles.NF_PERCENT)
    sb.line("fcf_margin", "FCF margin", fcf_margin, number_format=styles.NF_PERCENT)
    sb.line("capex_rev", "Capex / Revenue", capex_rev, number_format=styles.NF_PERCENT)
    return sb


def _wc_link(key):
    def w(period, registry):
        ref = registry.ref("Working Capital", key, period)
        return CellValue("=" + ref, DataType.LINKED, number_format=styles.NF_DAYS) if ref else None
    return w
