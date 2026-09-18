"""Canonical line-item synonym dictionary for heuristic extraction.

Maps messy annual-report row labels to canonical metric keys used across the
model. Matching is deliberately conservative: a candidate is only proposed when
a synonym matches, and every proposed datapoint is flagged REQUIRES_REVIEW.
"""

from __future__ import annotations

from modules.schemas import Statement

# metric_key -> (statement, [synonym substrings, lowercased])
SYNONYMS: dict[str, tuple[Statement, list[str]]] = {
    # Income statement
    "revenue": (Statement.INCOME_STATEMENT, [
        "revenue from operations", "total income", "total revenue", "net sales",
        "revenue from contracts", "turnover", "net revenue"]),
    "cogs": (Statement.INCOME_STATEMENT, [
        "cost of goods sold", "cost of materials consumed", "cost of sales",
        "cost of revenue"]),
    "employee_expense": (Statement.INCOME_STATEMENT, [
        "employee benefit", "employee expense", "staff cost", "personnel cost"]),
    "other_opex": (Statement.INCOME_STATEMENT, [
        "other expenses", "other operating expense"]),
    "ebitda": (Statement.INCOME_STATEMENT, ["ebitda", "operating profit before"]),
    "depreciation": (Statement.INCOME_STATEMENT, [
        "depreciation and amortis", "depreciation & amortis", "depreciation expense",
        "depreciation, amortis"]),
    "ebit": (Statement.INCOME_STATEMENT, ["operating profit", "ebit "]),
    "other_income": (Statement.INCOME_STATEMENT, ["other income"]),
    "interest_expense": (Statement.INCOME_STATEMENT, [
        "finance cost", "finance costs", "interest expense", "interest and finance"]),
    "pbt": (Statement.INCOME_STATEMENT, [
        "profit before tax", "profit/(loss) before tax", "pbt"]),
    "tax": (Statement.INCOME_STATEMENT, ["tax expense", "total tax", "income tax"]),
    "pat": (Statement.INCOME_STATEMENT, [
        "profit for the year", "profit after tax", "net profit", "profit/(loss) for the",
        "pat"]),
    "eps": (Statement.INCOME_STATEMENT, ["earnings per share", "basic eps", "eps"]),
    # Balance sheet
    "cash": (Statement.BALANCE_SHEET, [
        "cash and cash equivalent", "cash & cash equivalent", "cash and bank"]),
    "receivables": (Statement.BALANCE_SHEET, ["trade receivable", "sundry debtor"]),
    "inventory": (Statement.BALANCE_SHEET, ["inventor", "stock-in-trade"]),
    "ppe": (Statement.BALANCE_SHEET, [
        "property, plant and equipment", "property, plant & equipment",
        "fixed assets", "tangible assets"]),
    "intangibles": (Statement.BALANCE_SHEET, ["intangible asset", "goodwill"]),
    "total_current_assets": (Statement.BALANCE_SHEET, ["total current assets"]),
    "total_assets": (Statement.BALANCE_SHEET, ["total assets"]),
    "payables": (Statement.BALANCE_SHEET, ["trade payable", "sundry creditor"]),
    "short_term_debt": (Statement.BALANCE_SHEET, [
        "short-term borrowing", "short term borrowing", "current borrowing"]),
    "long_term_debt": (Statement.BALANCE_SHEET, [
        "long-term borrowing", "long term borrowing", "non-current borrowing"]),
    "total_current_liabilities": (Statement.BALANCE_SHEET, ["total current liabilities"]),
    "total_liabilities": (Statement.BALANCE_SHEET, ["total liabilities"]),
    "share_capital": (Statement.BALANCE_SHEET, ["equity share capital", "share capital"]),
    "reserves": (Statement.BALANCE_SHEET, ["other equity", "reserves and surplus",
                                           "reserves & surplus"]),
    "total_equity": (Statement.BALANCE_SHEET, ["total equity", "shareholders' fund",
                                               "shareholders funds", "net worth"]),
    # Cash flow
    "cfo": (Statement.CASH_FLOW, [
        "net cash from operating", "cash flow from operating",
        "net cash generated from operating", "cash generated from operations"]),
    "cfi": (Statement.CASH_FLOW, [
        "net cash from investing", "cash flow from investing",
        "net cash used in investing"]),
    "cff": (Statement.CASH_FLOW, [
        "net cash from financing", "cash flow from financing",
        "net cash used in financing"]),
    "capex": (Statement.CASH_FLOW, [
        "purchase of property", "purchase of fixed assets", "capital expenditure",
        "additions to property"]),
    "depreciation_cf": (Statement.CASH_FLOW, ["depreciation and amortis"]),
}


def match_metric(label: str) -> tuple[str, Statement] | None:
    """Return (metric_key, statement) for a row label, or None if no match.

    Longer synonym matches win to avoid e.g. matching 'ebit ' inside 'ebitda'.
    """
    low = " ".join(label.lower().split())
    best: tuple[str, Statement] | None = None
    best_len = 0
    for metric, (stmt, syns) in SYNONYMS.items():
        for syn in syns:
            if syn in low and len(syn) > best_len:
                best = (metric, stmt)
                best_len = len(syn)
    return best
