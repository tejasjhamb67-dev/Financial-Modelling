"""MODEL_BUILD_REPORT.md generator."""

from __future__ import annotations

from modules.model_spec import ModelSpec
from modules.schemas import DataType, SourceDatabase, ValidationStatus


def generate(db: SourceDatabase, spec: ModelSpec, gate_results, qc_result,
             modules_created: list[str], build_date: str) -> str:
    s = db.summary()
    linked = "many (cross-sheet references throughout)"

    checks_line = "See Checks sheet (evaluated live in Excel)."
    lines = [
        f"# Model Build Report - {spec.company}",
        "",
        f"*Generated {build_date}*",
        "",
        "## Overview",
        f"- Company: **{spec.company}**",
        f"- Reporting currency / unit: {spec.currency or 'N/A'} / {spec.unit or 'N/A'}",
        f"- Consolidation basis: {spec.consolidation}",
        f"- Historical periods: {', '.join(spec.historical_periods) or 'None'}",
        f"- Forecast periods: {', '.join(spec.forecast_periods) or 'None'}",
        f"- Documents processed: {s['documents']}",
        "",
        "## Sheets created",
    ]
    for sheet in spec.sheets:
        lines.append(f"- {sheet}")
    lines += [
        "",
        "## Modules",
    ]
    for m in modules_created:
        lines.append(f"- {m}")

    lines += [
        "",
        "## Datapoints",
        f"- Total datapoints: {s['datapoints']}",
        f"- Reported: {s['reported']}",
        f"- Derived (constructed in-model by formula): built from components where disclosed",
        f"- Linked: {linked}",
        f"- Unresolved (missing / Not Found): {s['unresolved']}",
        f"- Requires review: {s['requires_review']}",
        f"- Restatements detected: {s['restatements']}",
        "",
        "## Validation gates",
    ]
    for g in gate_results:
        lines.append(f"### {g.name} - {'PASS' if g.passed else 'FAIL'}")
        for f in g.findings:
            lines.append(f"- [{f.level}] {f.code}: {f.message}")
    lines += [
        "",
        "## Final QC (Gate 3)",
        f"Result: {'PASS' if qc_result.passed else 'FAIL'}",
    ]
    for f in qc_result.findings:
        lines.append(f"- [{f.level}] {f.code}: {f.message}")

    lines += [
        "",
        "## Checks",
        checks_line,
        "",
        "## Data type classification",
        f"- REPORTED cells: {s['reported']} (blue)",
        "- DERIVED cells: constructed by explicit additive formulas (black)",
        "- LINKED cells: cross-sheet references (green)",
        "- USER ASSUMPTION cells: forecast/valuation inputs (yellow)",
    ]
    return "\n".join(lines) + "\n"
