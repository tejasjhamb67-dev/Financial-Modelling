"""Styling helpers driven by ``config/formatting_config.yaml``.

Colour coding is meaningful (sell-side convention):
  BLUE   - hardcoded historical / external inputs
  BLACK  - formulas within the same worksheet
  GREEN  - links to another worksheet
  PURPLE - external / reference information
  RED    - errors / failed checks
  YELLOW fill - user assumption cells
"""

from __future__ import annotations

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from modules.config_loader import formatting_config
from modules.schemas import DataType

_CFG = formatting_config()
_C = _CFG["colors"]
_F = _CFG["fills"]
_NF = _CFG["number_formats"]
_FONT = _CFG["fonts"]["name"]
_SIZE = _CFG["fonts"]["size"]


def font_color_for(data_type: DataType, is_link: bool = False) -> str:
    if is_link or data_type == DataType.LINKED:
        return _C["link_green"]
    if data_type == DataType.REPORTED:
        return _C["input_blue"]
    if data_type == DataType.ASSUMPTION:
        return _C["formula_black"]
    return _C["formula_black"]  # DERIVED / formula


def base_font(color: str | None = None, bold: bool = False, size: int | None = None) -> Font:
    return Font(name=_FONT, size=size or _SIZE, bold=bold,
                color=color or _C["formula_black"])


def input_font() -> Font:
    return base_font(_C["input_blue"])


def formula_font() -> Font:
    return base_font(_C["formula_black"])


def link_font() -> Font:
    return base_font(_C["link_green"])


def reference_font() -> Font:
    return base_font(_C["reference_purple"])


def error_font() -> Font:
    return base_font(_C["error_red"], bold=True)


def title_font() -> Font:
    return Font(name=_FONT, size=_CFG["fonts"]["title_size"], bold=True, color="FFFFFFFF")


def header_font() -> Font:
    return Font(name=_FONT, size=_CFG["fonts"]["header_size"], bold=True, color="FFFFFFFF")


def section_font() -> Font:
    return Font(name=_FONT, size=_SIZE, bold=True, color=_C["header_navy"])


def assumption_fill() -> PatternFill:
    return PatternFill("solid", fgColor=_F["assumption_yellow"])


def header_fill() -> PatternFill:
    return PatternFill("solid", fgColor=_F["header_fill"])


def subheader_fill() -> PatternFill:
    return PatternFill("solid", fgColor=_F["subheader_fill"])


def forecast_fill() -> PatternFill:
    return PatternFill("solid", fgColor=_F["forecast_band"])


def check_fill(status: str) -> PatternFill:
    key = {"PASS": "check_pass", "WARNING": "check_warn",
           "ERROR": "check_error"}.get(status, "check_warn")
    return PatternFill("solid", fgColor=_F[key])


# number formats
NF_CURRENCY = _NF["currency"]
NF_CURRENCY_INT = _NF["currency_int"]
NF_RATIO = _NF["ratio"]
NF_PERCENT = _NF["percent"]
NF_MULTIPLE = _NF["multiple"]
NF_DAYS = _NF["days"]

RIGHT = Alignment(horizontal="right")
LEFT = Alignment(horizontal="left")
CENTER = Alignment(horizontal="center")

_thin = Side(style="thin", color="FFBFBFBF")
BORDER_BOTTOM = Border(bottom=_thin)
BORDER_TOP = Border(top=Side(style="thin", color="FF808080"))


def col_widths() -> tuple[int, int]:
    lay = _CFG["layout"]
    return lay["label_col_width"], lay["data_col_width"]
