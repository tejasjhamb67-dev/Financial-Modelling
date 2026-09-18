"""Two-pass sheet framework.

A :class:`SheetBuilder` collects ordered :class:`RowSpec` rows. The workbook is
built in two passes:

  1. LAYOUT - every sheet reserves its rows and registers the address of each
     (sheet, metric, period) cell in the :class:`CellRegistry`.
  2. FILL   - every sheet writes its cells. Because *all* addresses were
     registered in pass 1, a writer on any sheet can emit a real cross-sheet
     reference (``='Income Statement'!F10``) that is guaranteed to resolve.

Row writers return a :class:`CellValue` describing the content (a number, or a
formula string beginning with ``=``) and its :class:`DataType`, which drives the
font colour and the lineage record.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Optional, Union

from openpyxl.worksheet.worksheet import Worksheet

from excel.formatting import styles
from excel.linking.cell_registry import CellRegistry, Lineage
from modules.schemas import DataType

Content = Union[float, int, str, None]


@dataclass
class CellValue:
    content: Content
    data_type: DataType = DataType.DERIVED
    comment: Optional[str] = None
    number_format: Optional[str] = None
    source_id: Optional[str] = None
    source_document: Optional[str] = None
    page: Optional[int] = None


# writer(period, registry) -> CellValue | None
Writer = Callable[[str, CellRegistry], Optional[CellValue]]


@dataclass
class RowSpec:
    key: Optional[str] = None
    label: str = ""
    kind: str = "line"            # title|header|section|line|total|blank|note
    number_format: str = styles.NF_CURRENCY
    bold: bool = False
    writer: Optional[Writer] = None
    row: int = 0                  # assigned during layout


class SheetBuilder:
    def __init__(self, name: str, periods_row: bool = True, subtitle: str = ""):
        self.name = name
        self.periods_row = periods_row
        self.subtitle = subtitle
        self.rows: list[RowSpec] = []

    # -- authoring -------------------------------------------------------
    def section(self, label: str) -> None:
        self.rows.append(RowSpec(label=label, kind="section", bold=True))

    def blank(self) -> None:
        self.rows.append(RowSpec(kind="blank"))

    def note(self, label: str) -> None:
        self.rows.append(RowSpec(label=label, kind="note"))

    def line(self, key: Optional[str], label: str, writer: Optional[Writer] = None,
             number_format: str = styles.NF_CURRENCY, bold: bool = False,
             kind: str = "line") -> None:
        self.rows.append(RowSpec(key=key, label=label, kind=kind,
                                 number_format=number_format, bold=bold, writer=writer))

    def total(self, key: Optional[str], label: str, writer: Optional[Writer] = None,
              number_format: str = styles.NF_CURRENCY) -> None:
        self.line(key, label, writer, number_format, bold=True, kind="total")

    # -- pass 1: layout --------------------------------------------------
    def layout(self, registry: CellRegistry) -> None:
        r = 1  # title
        r += 1  # subtitle
        if self.periods_row:
            r += 1  # header row with period labels
        for spec in self.rows:
            if spec.kind == "blank":
                spec.row = r
                r += 1
                continue
            spec.row = r
            if spec.key and spec.kind in ("line", "total"):
                registry.set_row(self.name, spec.key, r)
                for period in registry.all_periods:
                    registry.register(self.name, spec.key, period, r)
            r += 1

    # -- pass 2: fill ----------------------------------------------------
    def fill(self, ws: Worksheet, registry: CellRegistry, forecast_periods: list[str],
             title: str, currency_unit: str) -> None:
        label_w, data_w = styles.col_widths()
        ws.column_dimensions[registry.label_col_letter()].width = label_w
        for period in registry.all_periods:
            ws.column_dimensions[registry.period_col_letter(period)].width = data_w

        # title
        tcell = ws.cell(row=1, column=1, value=title)
        tcell.font = styles.header_font()
        tcell.fill = styles.header_fill()
        last_col = registry.first_data_col + max(len(registry.all_periods) - 1, 0)
        for c in range(1, last_col + 1):
            ws.cell(row=1, column=c).fill = styles.header_fill()
        # subtitle
        sub = self.subtitle or currency_unit
        scell = ws.cell(row=2, column=1, value=sub)
        scell.font = styles.base_font(styles._C["subtle_grey"])

        header_row = 3
        if self.periods_row:
            hc = ws.cell(row=header_row, column=1, value="")
            for period in registry.all_periods:
                cell = ws.cell(row=header_row, column=registry.period_col_idx(period), value=period)
                cell.font = styles.section_font()
                cell.alignment = styles.CENTER
                cell.border = styles.BORDER_BOTTOM
                if period in forecast_periods:
                    cell.fill = styles.forecast_fill()

        for spec in self.rows:
            self._write_row(ws, registry, spec, forecast_periods)

        # freeze label col + header
        if self.periods_row:
            from openpyxl.utils import get_column_letter
            ws.freeze_panes = f"{get_column_letter(registry.first_data_col)}{header_row + 1}"
        ws.sheet_view.showGridLines = False

    def _write_row(self, ws, registry, spec: RowSpec, forecast_periods) -> None:
        row = spec.row
        if spec.kind == "blank":
            return
        # label
        lcell = ws.cell(row=row, column=1, value=spec.label)
        if spec.kind == "section":
            lcell.font = styles.section_font()
            lcell.fill = styles.subheader_fill()
            last_col = registry.first_data_col + max(len(registry.all_periods) - 1, 0)
            for c in range(2, last_col + 1):
                ws.cell(row=row, column=c).fill = styles.subheader_fill()
            return
        if spec.kind == "note":
            lcell.font = styles.base_font(styles._C["subtle_grey"])
            return
        lcell.font = styles.base_font(bold=spec.bold)

        if not spec.writer or not spec.key:
            return
        for period in registry.all_periods:
            cv = spec.writer(period, registry)
            if cv is None:
                continue
            self._write_cell(ws, registry, spec, period, cv, forecast_periods)

    def _write_cell(self, ws, registry, spec, period, cv: CellValue, forecast_periods) -> None:
        col = registry.period_col_idx(period)
        cell = ws.cell(row=spec.row, column=col)
        content = cv.content
        is_formula = isinstance(content, str) and content.startswith("=")
        if content is None:
            cell.value = None
        else:
            cell.value = content
        # font colour by data type / link detection
        is_link = is_formula and "!" in content
        cell.font = styles.base_font(styles.font_color_for(cv.data_type, is_link=is_link),
                                     bold=spec.bold)
        cell.number_format = cv.number_format or spec.number_format
        cell.alignment = styles.RIGHT
        if spec.kind == "total":
            cell.border = styles.BORDER_TOP
        if period in forecast_periods:
            cell.fill = styles.forecast_fill()
            if cv.data_type == DataType.ASSUMPTION:
                cell.fill = styles.assumption_fill()
        elif cv.data_type == DataType.ASSUMPTION:
            cell.fill = styles.assumption_fill()
        # comment
        if cv.comment:
            from openpyxl.comments import Comment
            cell.comment = Comment(cv.comment, "FM Agent")
        # lineage
        registry.add_lineage(Lineage(
            sheet=self.name, metric=spec.key, period=period,
            address=f"{registry.period_col_letter(period)}{spec.row}",
            data_type=cv.data_type.value,
            source_id=cv.source_id, source_document=cv.source_document, page=cv.page,
            formula=content if is_formula else None,
        ))
