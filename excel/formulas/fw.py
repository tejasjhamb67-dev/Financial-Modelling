"""Reusable cell-writer helpers (formula construction).

These build :class:`CellValue` objects for the common modelling patterns:
reported historical inputs, explicit additive construction, cross-sheet links,
assumption references and prior-period growth. Formula construction deliberately
uses explicit ``+``/``-`` of cell addresses (never ``SUM(range)`` for
constructive totals) so the build stays visible and auditable.
"""

from __future__ import annotations

from typing import Optional

from excel.linking.cell_registry import CellRegistry
from excel.workbook_builder.grid import CellValue
from modules.schemas import Consolidation, DataPoint, DataType, SourceDatabase


def source_comment(dp: DataPoint) -> str:
    parts = [f"{dp.label or dp.metric}",
             f"Source ID: {dp.source_id}",
             f"Document: {dp.source_document or 'n/a'}"]
    if dp.page:
        parts.append(f"Page: {dp.page}")
    if dp.section:
        parts.append(f"Section: {dp.section}")
    parts.append(f"Type: {dp.data_type.value}  Status: {dp.status.value}")
    if dp.notes:
        parts.append(f"Notes: {dp.notes}")
    return "\n".join(parts)


def reported(db: SourceDatabase, metric: str, period: str,
             consol: Optional[Consolidation] = None) -> Optional[CellValue]:
    """A directly disclosed historical number (blue input) with source comment."""
    dp = db.get(metric, period, consol)
    if dp is None or dp.value is None:
        return None
    return CellValue(
        content=dp.value, data_type=DataType.REPORTED,
        comment=source_comment(dp), source_id=dp.source_id,
        source_document=dp.source_document, page=dp.page,
    )


def construct(expr: str, sheet: str, period: str, registry: CellRegistry,
              data_type: DataType = DataType.DERIVED) -> Optional[str]:
    """Translate an expression of metric keys into an explicit local formula.

    ``"revenue - cogs"`` -> ``"=B10-B11"`` using this sheet's registered rows.
    Returns None if any referenced metric is not present on the sheet.
    """
    tokens = _tokenize(expr)
    out = []
    for tok in tokens:
        if tok in "+-*/()":
            out.append(tok)
        elif _is_number(tok):
            out.append(tok)
        else:
            addr = registry.addr(sheet, tok, period)
            if addr is None:
                return None
            out.append(addr)
    return "=" + "".join(out)


def construct_cv(expr: str, sheet: str, period: str, registry: CellRegistry,
                 data_type: DataType = DataType.DERIVED) -> Optional[CellValue]:
    f = construct(expr, sheet, period, registry, data_type)
    if f is None:
        return None
    return CellValue(content=f, data_type=data_type)


def link(sheet_from: str, metric: str, period: str, registry: CellRegistry,
         comment: Optional[str] = None) -> Optional[CellValue]:
    """A green cross-sheet link."""
    ref = registry.ref(sheet_from, metric, period)
    if ref is None:
        return None
    return CellValue(content="=" + ref, data_type=DataType.LINKED,
                     comment=comment or f"Linked from {sheet_from}")


def local_ref(sheet: str, metric: str, period: str, registry: CellRegistry,
              data_type: DataType = DataType.DERIVED) -> Optional[CellValue]:
    addr = registry.addr(sheet, metric, period)
    if addr is None:
        return None
    return CellValue(content="=" + addr, data_type=data_type)


def growth_formula(sheet: str, metric: str, prev_period: str, growth_addr: str,
                   registry: CellRegistry) -> Optional[CellValue]:
    """``prev * (1 + growth_cell)`` - forecast off an assumption."""
    prev = registry.addr(sheet, metric, prev_period)
    if prev is None:
        return None
    return CellValue(content=f"={prev}*(1+{growth_addr})", data_type=DataType.DERIVED)


def _tokenize(expr: str) -> list[str]:
    out = []
    cur = ""
    for ch in expr:
        if ch in "+-*/()":
            if cur.strip():
                out.append(cur.strip())
            out.append(ch)
            cur = ""
        else:
            cur += ch
    if cur.strip():
        out.append(cur.strip())
    return out


def _is_number(tok: str) -> bool:
    try:
        float(tok)
        return True
    except ValueError:
        return False
