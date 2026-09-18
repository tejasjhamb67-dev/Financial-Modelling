"""Restatement engine.

Compares overlapping periods for the same metric across different source
documents. When the same FYxx value differs materially between an older and a
newer report, we record a :class:`Restatement` for analyst review. We never
silently overwrite the older number - both are preserved and the newer
disclosure is retained by default (documented treatment).
"""

from __future__ import annotations

from modules.config_loader import validation_config
from modules.schemas import Restatement, SourceDatabase, Statement


def detect_restatements(db: SourceDatabase) -> list[Restatement]:
    """Detect cross-document restatements not already declared in the package."""
    tol = validation_config()["tolerances"]["restatement_flag_rel"]
    declared = {(r.metric, r.period) for r in db.restatements}
    found: list[Restatement] = []

    # group datapoints by (metric, period, consolidation)
    groups: dict[tuple, list] = {}
    for dp in db.datapoints:
        if dp.statement not in (Statement.INCOME_STATEMENT, Statement.BALANCE_SHEET,
                                Statement.CASH_FLOW):
            continue
        if dp.value is None or not dp.source_document:
            continue
        groups.setdefault((dp.metric, dp.period, dp.consolidation.value), []).append(dp)

    for (metric, period, _), dps in groups.items():
        if len(dps) < 2:
            continue
        docs = {dp.source_document: dp for dp in dps}
        if len(docs) < 2:
            continue
        ordered = sorted(dps, key=lambda d: (d.source_document or ""))
        old, new = ordered[0], ordered[-1]
        if old.value in (None, 0):
            continue
        rel = abs(new.value - old.value) / abs(old.value)
        if rel > tol and (metric, period) not in declared:
            found.append(Restatement(
                metric=metric,
                period=period,
                old_value=old.value,
                new_value=new.value,
                old_document=old.source_document,
                new_document=new.source_document,
                pct_change=new.value / old.value - 1.0,
                reason="Cross-report difference detected automatically; investigate "
                       "(restatement / reclassification / policy change).",
            ))
    return found


def apply(db: SourceDatabase) -> None:
    for r in detect_restatements(db):
        db.add_restatement(r)
