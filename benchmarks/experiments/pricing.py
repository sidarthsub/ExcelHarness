"""Per-model $/token rates, keyed by the model IDs the Agent SDK accepts.

Rates are in USD per token (not per million). Keep these in sync with
https://www.anthropic.com/pricing. The loss function uses cold-equivalent
cost — every input-class token (regular, cache_creation, cache_read) is
priced at the plain input rate. This removes cache-state dependence from
the signal the outer loop optimizes.

`resolve(model)` accepts both short aliases the Agent SDK uses
(`"sonnet"`, `"haiku"`, `"opus"`) and fully-qualified model IDs. Add new
models here as they're introduced.
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Rates:
    input_per_token: float
    output_per_token: float


# USD per token. Update when pricing changes or new models are added.
# Source: https://www.anthropic.com/pricing (2026-04).
_RATES: dict[str, Rates] = {
    "claude-opus-4-7":     Rates(15.0e-6, 75.0e-6),
    "claude-sonnet-4-6":   Rates(3.0e-6,  15.0e-6),
    "claude-haiku-4-5":    Rates(1.0e-6,  5.0e-6),
}

_ALIAS: dict[str, str] = {
    "opus":   "claude-opus-4-7",
    "sonnet": "claude-sonnet-4-6",
    "haiku":  "claude-haiku-4-5",
}


def resolve(model: str) -> Rates:
    """Look up rates by alias or full model ID."""
    key = _ALIAS.get(model, model)
    if key not in _RATES:
        raise KeyError(
            f"no pricing entry for model {model!r}. "
            f"Add it to benchmarks/experiments/pricing.py._RATES."
        )
    return _RATES[key]


def known_models() -> list[str]:
    return list(_ALIAS.keys()) + list(_RATES.keys())
