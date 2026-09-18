"""Central configuration loader.

All tunable behaviour lives in ``config/*.yaml`` so the code stays free of
hardcoded company specifics and magic numbers.
"""

from __future__ import annotations

import os
from functools import lru_cache
from typing import Any

import yaml

CONFIG_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config")


def _load(name: str) -> dict[str, Any]:
    path = os.path.join(CONFIG_DIR, name)
    with open(path, encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


@lru_cache(maxsize=None)
def model_config() -> dict[str, Any]:
    return _load("model_config.yaml")


@lru_cache(maxsize=None)
def formatting_config() -> dict[str, Any]:
    return _load("formatting_config.yaml")


@lru_cache(maxsize=None)
def validation_config() -> dict[str, Any]:
    return _load("validation_config.yaml")


@lru_cache(maxsize=None)
def accounting_rules() -> dict[str, Any]:
    return _load("accounting_rules.yaml")


class AccountingTaxonomy:
    """Convenience wrapper over ``accounting_rules.yaml``.

    Exposes ordered line items and construction rules per statement so the
    financial-statement engine can build explicit additive formulas.
    """

    def __init__(self, rules: dict[str, Any] | None = None):
        self.rules = rules or accounting_rules()

    def order(self, statement_key: str) -> list[str]:
        return list(self.rules.get(statement_key, {}).get("order", []))

    def metrics(self, statement_key: str) -> dict[str, Any]:
        return dict(self.rules.get(statement_key, {}).get("metrics", {}))

    def label(self, statement_key: str, metric: str) -> str:
        m = self.metrics(statement_key).get(metric, {})
        return m.get("label", metric.replace("_", " ").title())

    def construction(self, statement_key: str, metric: str) -> str | None:
        m = self.metrics(statement_key).get(metric, {})
        return m.get("construct")

    def is_cost(self, statement_key: str, metric: str) -> bool:
        m = self.metrics(statement_key).get(metric, {})
        return m.get("sign") == "cost"
