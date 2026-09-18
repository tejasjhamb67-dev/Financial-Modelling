"""Master workbook builder (two-pass: layout then fill)."""

from __future__ import annotations

from typing import Callable

from openpyxl import Workbook

from excel.linking.cell_registry import CellRegistry
from excel.workbook_builder.grid import SheetBuilder
from excel.workbook_builder.sheets import (
    assumptions,
    meta,
    model_ratios,
    schedules,
    statements,
    valuation,
)
from modules.model_spec import ModelSpec
from modules.schemas import SourceDatabase

# grid sheet builders -> functions returning a SheetBuilder
_GRID: dict[str, Callable[[ModelSpec, SourceDatabase], SheetBuilder]] = {
    "Assumptions": assumptions.build_assumptions,
    "Revenue Build": schedules.build_revenue,
    "Operating Drivers": schedules.build_operating_drivers,
    "Income Statement": statements.build_income_statement,
    "Balance Sheet": statements.build_balance_sheet,
    "Cash Flow": statements.build_cash_flow,
    "3-Statement Model": model_ratios.build_three_statement,
    "Working Capital": schedules.build_working_capital,
    "Capex & D&A": schedules.build_capex_dna,
    "Debt": schedules.build_debt,
    "Ratios": model_ratios.build_ratios,
}

_CUSTOM = {"Cover", "Sources", "Valuation", "Sensitivities", "Checks"}


class WorkbookBuilder:
    def __init__(self, spec: ModelSpec, db: SourceDatabase, build_date: str = ""):
        self.spec = spec
        self.db = db
        self.build_date = build_date
        self.registry = CellRegistry(spec.all_periods)
        self.wb = Workbook()
        self.wb.remove(self.wb.active)

    def build(self, path: str) -> CellRegistry:
        spec = self.spec
        grid_builders: dict[str, SheetBuilder] = {}

        # create worksheets in the specified order
        for name in spec.sheets:
            self.wb.create_sheet(title=name[:31])

        # phase 1: layout all grid sheets (registers every cell address)
        for name in spec.sheets:
            if name in _GRID:
                sb = _GRID[name](spec, self.db)
                sb.layout(self.registry)
                grid_builders[name] = sb

        # phase 2: fill (all addresses now known -> cross-sheet links resolve)
        unit_str = f"{spec.currency or ''} {spec.unit or ''}".strip() or "Reporting unit"
        for name in spec.sheets:
            ws = self.wb[name[:31]]
            if name in grid_builders:
                grid_builders[name].fill(ws, self.registry, spec.forecast_periods,
                                         title=name, currency_unit=unit_str)
            elif name == "Cover":
                meta.fill_cover(ws, self.registry, spec, self.db, self.build_date)
            elif name == "Sources":
                meta.fill_sources(ws, self.registry, spec, self.db)
            elif name == "Valuation":
                valuation.fill_valuation(ws, self.registry, spec, self.db)
            elif name == "Sensitivities":
                valuation.fill_sensitivities(ws, self.registry, spec, self.db)
            elif name == "Checks":
                meta.fill_checks(ws, self.registry, spec, self.db)

        self.wb.save(path)
        return self.registry
