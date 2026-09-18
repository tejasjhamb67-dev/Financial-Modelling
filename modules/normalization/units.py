"""Unit handling for normalization.

Monetary values across documents may be disclosed in different units (crore,
lakh, million, thousand). We convert everything to a single model reporting
unit so statements reconcile. The absolute scale factor is relative to the base
currency unit (1.0 = one unit of currency).
"""

from __future__ import annotations

from typing import Optional

# scale in absolute currency units
_UNIT_SCALE = {
    "unit": 1.0,
    "thousand": 1e3,
    "lakh": 1e5,
    "million": 1e6,
    "crore": 1e7,
    "billion": 1e9,
}


def _canon(unit: Optional[str]) -> tuple[Optional[str], Optional[str]]:
    """Split 'INR crore' -> ('INR', 'crore'). Currency may be None."""
    if not unit:
        return None, None
    parts = unit.strip().lower().split()
    currency = None
    scale = None
    for p in parts:
        if p in _UNIT_SCALE:
            scale = p
        elif p in ("inr", "usd", "eur", "gbp", "rs", "rs.", "₹", "$"):
            currency = {"rs": "INR", "rs.": "INR", "₹": "INR", "$": "USD"}.get(p, p.upper())
    return currency, scale


def scale_factor(from_unit: Optional[str], to_unit: Optional[str]) -> float:
    """Multiplicative factor to convert a value from ``from_unit`` to ``to_unit``.

    Returns 1.0 when either unit is unknown (we never silently rescale on a
    guess - unknown units are surfaced by validation instead).
    """
    _, from_scale = _canon(from_unit)
    _, to_scale = _canon(to_unit)
    if not from_scale or not to_scale:
        return 1.0
    return _UNIT_SCALE[from_scale] / _UNIT_SCALE[to_scale]


def same_currency(a: Optional[str], b: Optional[str]) -> bool:
    ca, _ = _canon(a)
    cb, _ = _canon(b)
    if ca is None or cb is None:
        return True  # unknown -> don't block; validation flags it
    return ca == cb
