"""Explicit validation gates.

Gate 1 - data validation (before model construction).
Gate 2 - model specification validation (before Excel generation).
Gate 3 - final workbook QC (before delivery, see excel.qc).

Each gate returns a :class:`GateResult` with pass/fail plus itemised
findings. A gate that fails on a *material* error stops the pipeline unless the
caller explicitly overrides - we never silently continue past material errors.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from modules.config_loader import validation_config
from modules.model_spec import ModelSpec
from modules.schemas import SourceDatabase, Statement, ValidationStatus


@dataclass
class Finding:
    level: str      # INFO / WARNING / ERROR
    code: str
    message: str


@dataclass
class GateResult:
    name: str
    passed: bool = True
    findings: list[Finding] = field(default_factory=list)

    def add(self, level: str, code: str, message: str) -> None:
        self.findings.append(Finding(level, code, message))
        if level == "ERROR":
            self.passed = False

    @property
    def errors(self) -> list[Finding]:
        return [f for f in self.findings if f.level == "ERROR"]

    @property
    def warnings(self) -> list[Finding]:
        return [f for f in self.findings if f.level == "WARNING"]

    def report_lines(self) -> list[str]:
        out = [f"GATE: {self.name} -> {'PASS' if self.passed else 'FAIL'}"]
        for f in self.findings:
            out.append(f"  [{f.level}] {f.code}: {f.message}")
        return out


def gate1_data_validation(db: SourceDatabase) -> GateResult:
    cfg = validation_config()["gate1_data_validation"]
    res = GateResult("Gate 1 - Data Validation")

    def _has(stmt: Statement, n: int = 2) -> bool:
        return len({dp.metric for dp in db.by_statement(stmt) if dp.value is not None}) >= n

    if cfg.get("require_income_statement") and not _has(Statement.INCOME_STATEMENT):
        res.add("ERROR", "IS_MISSING", "Income statement not sufficiently extracted.")
    if cfg.get("require_balance_sheet") and not _has(Statement.BALANCE_SHEET):
        res.add("ERROR", "BS_MISSING", "Balance sheet not sufficiently extracted.")
    if cfg.get("require_cash_flow") and not _has(Statement.CASH_FLOW, 1):
        res.add("WARNING", "CF_MISSING", "Cash flow not disclosed; will be partly derived.")

    # units known
    if cfg.get("require_units_known"):
        unknown_units = [dp for dp in db.datapoints
                         if dp.value is not None and dp.currency and not dp.unit]
        if unknown_units:
            res.add("WARNING", "UNIT_UNKNOWN",
                    f"{len(unknown_units)} monetary datapoints have unknown units.")

    # consolidation known
    if cfg.get("require_consolidation_known"):
        if all(dp.consolidation.value == "unknown" for dp in db.datapoints):
            res.add("WARNING", "CONSOL_UNKNOWN", "Consolidation basis not established.")

    # historical periods
    fy = {dp.period for dp in db.datapoints if dp.period.startswith("FY")}
    if len(fy) < cfg.get("require_min_historical_periods", 2):
        res.add("ERROR", "FEW_PERIODS",
                f"Only {len(fy)} historical period(s); need >= "
                f"{cfg.get('require_min_historical_periods', 2)}.")

    # material extraction errors
    failed = [dp for dp in db.datapoints
              if dp.status == ValidationStatus.RECONCILIATION_FAILED]
    if failed:
        res.add("WARNING", "EXTRACTION_AMBIGUOUS",
                f"{len(failed)} datapoints have ambiguous extraction; review required.")

    if not cfg.get("fail_on_material_error", True):
        res.passed = True
    return res


def gate2_model_spec(spec: ModelSpec) -> GateResult:
    cfg = validation_config()["gate2_model_specification"]
    res = GateResult("Gate 2 - Model Specification")

    if cfg.get("require_currency") and not spec.currency:
        res.add("WARNING", "NO_CURRENCY", "Reporting currency not established.")
    if cfg.get("require_units") and not spec.unit:
        res.add("WARNING", "NO_UNITS", "Reporting unit not established.")
    if cfg.get("require_period_labels") and not spec.historical_periods:
        res.add("ERROR", "NO_PERIODS", "No historical periods in spec.")

    if len(spec.forecast_periods) != 3:
        res.add("WARNING", "FORECAST_HORIZON",
                f"Forecast horizon is {len(spec.forecast_periods)} years (expected 3).")

    if not spec.income_statement_lines:
        res.add("ERROR", "NO_IS_LINES", "Income statement has no lines to build.")

    for exc in spec.exceptions:
        res.add("INFO", "SPEC_NOTE", exc)
    return res
