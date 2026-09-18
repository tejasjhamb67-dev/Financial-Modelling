"""MODEL_DILIGENCE.md generator - unresolved items must be highly visible."""

from __future__ import annotations

from modules.model_spec import ModelSpec
from modules.schemas import DataType, SourceDatabase, ValidationStatus


def generate(db: SourceDatabase, spec: ModelSpec, build_date: str) -> str:
    L = [f"# Model Diligence Report - {spec.company}", "", f"*Generated {build_date}*", ""]

    L += ["## Documents used"]
    if db.documents:
        for d in db.documents:
            L.append(f"- **{d.name}** ({d.doc_type}, {d.fiscal_year or 'FY n/a'}, "
                     f"{d.consolidation.value}, {d.currency or ''} {d.unit or ''})")
    else:
        L.append("- None recorded.")

    L += ["", "## Historical coverage",
          f"- Periods: {', '.join(spec.historical_periods) or 'None'}",
          f"- Target was ~5 years; {len(spec.historical_periods)} identified."]

    L += ["", "## Consolidation treatment",
          f"- Primary model basis: **{spec.consolidation}**.",
          "- Standalone figures, if supplied, are preserved separately in the source registry."]

    # restatements
    L += ["", "## Restatements & reclassifications"]
    if db.restatements:
        for r in db.restatements:
            pct = f"{r.pct_change*100:.1f}%" if r.pct_change is not None else "n/a"
            L.append(f"- **{r.metric} {r.period}**: {r.old_value} ({r.old_document}) -> "
                     f"{r.new_value} ({r.new_document}), change {pct}. Reason: {r.reason} "
                     f"Treatment: {r.treatment}")
    else:
        L.append("- None detected across supplied documents.")

    # missing data
    missing = [dp for dp in db.datapoints if dp.value is None]
    L += ["", "## Missing data (NOT fabricated - shown as Not Found)"]
    if missing:
        for dp in missing[:50]:
            L.append(f"- {dp.metric} {dp.period} ({dp.statement.value}) - {dp.status.value}")
        if len(missing) > 50:
            L.append(f"- ...and {len(missing) - 50} more.")
    else:
        L.append("- No explicitly missing datapoints.")

    # requires review
    review = [dp for dp in db.datapoints if dp.status in (
        ValidationStatus.REQUIRES_REVIEW, ValidationStatus.RECONCILIATION_FAILED)]
    L += ["", "## ⚠ Analyst review required"]
    if review:
        for dp in review[:50]:
            L.append(f"- {dp.metric} {dp.period}: {dp.status.value} - {dp.notes or 'verify against source'}")
        if len(review) > 50:
            L.append(f"- ...and {len(review) - 50} more (see Sources sheet).")
    else:
        L.append("- No items flagged for review.")

    # derived metrics
    L += ["", "## Derived metrics",
          "- Statement totals (gross profit, EBITDA, EBIT, PBT, PAT, current/non-current totals,",
          "  CFO/CFI/CFF) are constructed by explicit additive formulas where components are",
          "  disclosed; otherwise the reported total is used as a blue input.",
          "- All ratios and forecast lines are formula-driven."]

    # checks
    L += ["", "## Accounting & cash-flow checks",
          "- Balance sheet identity (Assets - Equity - Liabilities), cash-flow closing-cash vs",
          "  balance-sheet cash, revenue segment reconciliation and debt reconciliation are wired",
          "  on the **Checks** sheet and evaluate live in Excel.",
          "- Checks are never forced to pass; genuine discrepancies are displayed."]

    # exceptions
    L += ["", "## Other exceptions"]
    if spec.exceptions:
        for e in spec.exceptions:
            L.append(f"- {e}")
    else:
        L.append("- None.")

    L += ["", "## Analyst responsibilities (not automated)",
          "- Forecast assumptions, business outlook, valuation assumptions (WACC, terminal growth,",
          "  multiples, peers, share count), scenarios, investment thesis, target price and",
          "  conclusion remain the analyst's decisions."]
    return "\n".join(L) + "\n"
