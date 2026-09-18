"""Gate 3 - final workbook QC.

Programmatically inspects the generated ``.xlsx``: structure, chronology,
forecast horizon, formula integrity, cross-sheet reference validity, presence of
the balance-sheet / cash-flow checks, and accidental forecast hardcodes. Numeric
reconciliation itself is performed live inside the Checks sheet (Excel evaluates
those formulas on open); this QC verifies that the machinery is correctly wired.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from openpyxl import load_workbook

from modules.model_spec import ModelSpec
from modules.validation.gates import GateResult

_SHEET_REF = re.compile(r"'([^']+)'!|([A-Za-z0-9_][A-Za-z0-9 _&\-]*)!")


@dataclass
class QCResult(GateResult):
    pass


def run_qc(path: str, spec: ModelSpec) -> GateResult:
    res = GateResult("Gate 3 - Final Workbook QC")
    try:
        wb = load_workbook(path)
    except Exception as e:  # pragma: no cover
        res.add("ERROR", "OPEN", f"Workbook failed to open: {e}")
        return res

    names = wb.sheetnames
    # 1. sheets exist
    for expected in spec.sheets:
        if expected[:31] not in names:
            res.add("ERROR", "SHEET_MISSING", f"Expected sheet '{expected}' missing.")

    # 2. chronology oldest -> newest
    def fy(p):
        d = "".join(c for c in p if c.isdigit())
        return int(d) if d else 0
    hist = spec.historical_periods
    if hist != sorted(hist, key=fy):
        res.add("ERROR", "CHRONOLOGY", "Historical periods are not oldest->newest.")
    # forecast follow historical
    if hist and spec.forecast_periods:
        if fy(spec.forecast_periods[0]) <= fy(hist[-1]):
            res.add("ERROR", "FORECAST_ORDER", "Forecast periods do not follow historical.")
    if len(spec.forecast_periods) != 3:
        res.add("WARNING", "FORECAST_HORIZON",
                f"Forecast horizon is {len(spec.forecast_periods)} (expected 3).")

    # 3. formula integrity + broken references + forecast hardcodes
    forecast_cols = set()  # (col letters) computed per sheet below
    broken = 0
    formula_cells = 0
    statement_sheets = {"Income Statement", "Balance Sheet", "Cash Flow"}
    forecast_hardcodes = 0

    # map period -> column letter (consistent across grid sheets)
    from openpyxl.utils import get_column_letter
    period_col = {p: get_column_letter(2 + i) for i, p in enumerate(spec.all_periods)}
    fc_cols = {period_col[p] for p in spec.forecast_periods}

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    formula_cells += 1
                    for m in _SHEET_REF.finditer(v):
                        sheet = m.group(1) or m.group(2)
                        if sheet and sheet not in names and not _is_function_token(sheet):
                            broken += 1
                            res.add("WARNING", "BROKEN_REF",
                                    f"{ws.title}!{cell.coordinate} references missing sheet '{sheet}'.")
                elif isinstance(v, (int, float)) and ws.title in statement_sheets:
                    col = cell.column_letter
                    if col in fc_cols and v not in (0,):
                        forecast_hardcodes += 1

    if formula_cells == 0:
        res.add("ERROR", "NO_FORMULAS", "Workbook contains no formulas - not a live model.")
    if forecast_hardcodes:
        res.add("WARNING", "FORECAST_HARDCODE",
                f"{forecast_hardcodes} forecast statement cell(s) are hardcoded numbers "
                "(should be formulas/assumptions).")

    # 4. cross-sheet linkage present on 3-statement
    if "3-Statement Model" in names:
        ws = wb["3-Statement Model"]
        links = sum(1 for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith("=") and "!" in c.value)
        if links == 0:
            res.add("ERROR", "STMT_NOT_LINKED", "3-Statement Model has no cross-sheet links.")

    # 5. checks wired
    if "Checks" in names:
        ws = wb["Checks"]
        text = " ".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
        if "Balance Sheet" not in text:
            res.add("WARNING", "NO_BS_CHECK", "Balance-sheet check not found on Checks sheet.")
        if "Cash Flow" not in text and "closing cash" not in text.lower():
            res.add("WARNING", "NO_CF_CHECK", "Cash-flow check not found on Checks sheet.")
    else:
        res.add("ERROR", "NO_CHECKS", "Checks sheet missing.")

    res.add("INFO", "SUMMARY",
            f"{formula_cells} formula cells, {broken} broken refs, {forecast_hardcodes} forecast hardcodes.")
    return res


_FUNCTIONS = {"IF", "MAX", "MIN", "ABS", "AVERAGE", "MEDIAN", "SUMPRODUCT", "IFERROR",
              "COUNTIF", "SUM", "ROUND"}


def _is_function_token(tok: str) -> bool:
    return tok.upper() in _FUNCTIONS
