"""Normalization stage.

Takes a raw SourceDatabase (possibly merged from several documents) and returns
a normalized copy where all monetary datapoints share one reporting unit and
currency. The raw value is preserved on each datapoint (``raw_value``) so the
RAW -> NORMALIZED -> VALIDATED -> MODEL chain stays auditable.

Normalization never fabricates: if a unit is unknown it leaves the value as-is
and marks the datapoint for review rather than guessing a scale.
"""

from __future__ import annotations

import copy
from dataclasses import replace
from typing import Optional

from modules.normalization.units import same_currency, scale_factor
from modules.schemas import DataPoint, SourceDatabase, Statement, ValidationStatus


def choose_reporting_basis(db: SourceDatabase) -> tuple[Optional[str], Optional[str]]:
    """Pick the model reporting (unit, currency) from the newest document."""
    unit = currency = None
    if db.documents:
        # newest document = last added / highest fiscal year
        docs = sorted(db.documents, key=lambda d: (d.fiscal_year or ""))
        newest = docs[-1]
        unit = newest.unit
        currency = newest.currency
    if unit is None:
        for dp in db.datapoints:
            if dp.unit:
                unit = dp.unit
                break
    if currency is None:
        for dp in db.datapoints:
            if dp.currency:
                currency = dp.currency
                break
    return unit, currency


def normalize(db: SourceDatabase) -> SourceDatabase:
    target_unit, target_currency = choose_reporting_basis(db)

    out = SourceDatabase(db.company)
    out.documents = copy.deepcopy(db.documents)
    out.restatements = copy.deepcopy(db.restatements)
    # Preserve the verbatim as-reported layer unchanged (it is the RAW record).
    out.reported_statements = copy.deepcopy(db.reported_statements)

    for dp in db.datapoints:
        ndp = copy.deepcopy(dp)
        # only scale monetary datapoints (skip pure operating drivers w/o currency)
        is_monetary = dp.statement in (
            Statement.INCOME_STATEMENT, Statement.BALANCE_SHEET,
            Statement.CASH_FLOW, Statement.SEGMENT) or dp.currency is not None
        if is_monetary and dp.value is not None and dp.unit and target_unit:
            if dp.unit.strip().lower() != (target_unit or "").strip().lower():
                if not same_currency(dp.unit, target_unit):
                    ndp.status = ValidationStatus.REQUIRES_REVIEW
                    ndp.notes = (ndp.notes + " | Currency mismatch vs model basis; not rescaled.").strip(" |")
                else:
                    factor = scale_factor(dp.unit, target_unit)
                    if factor != 1.0:
                        ndp.raw_value = dp.value
                        ndp.value = dp.value * factor
                        ndp.notes = (ndp.notes + f" | Rescaled {dp.unit} -> {target_unit} (x{factor:g})").strip(" |")
                    ndp.unit = target_unit
            else:
                ndp.unit = target_unit
        if is_monetary and target_currency and ndp.currency is None:
            ndp.currency = target_currency
        out.add(ndp)
    return out
