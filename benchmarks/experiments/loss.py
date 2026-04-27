"""Scalar loss function the autoresearch outer loop minimizes.

REWRITE 2026-04-27. Old loss had three brittleness problems that pushed
the Researcher toward routing-style optimizations rather than real
architectural improvements:

  1. Cost/time were normalized against per-task budgets from task.yaml
     and capped at 2× budget. Tight t0 budgets ($0.10) saturated
     constantly — no gradient on the expensive side.
  2. Budgets are arbitrary anchors that don't exist in production.
     Every routing change exploited a budget threshold that won't be
     present when deployed.
  3. Accuracy and cost+time had equal max contribution (~1.0 each), so
     a 50% accuracy run could score the same as a slightly-over-budget
     100% accuracy run.

New formula:

    acc_loss   = 2.0 * (1 - accuracy)                           # [0, 2]
    cost_loss  = max(0, 0.05 * log(max(cost, 0.01) / 0.20))     # log-scale, no cap
    time_loss  = max(0, 0.05 * log(max(wall, 1.0)  / 60.0))     # log-scale, no cap

    loss = acc_loss + cost_loss + time_loss

  - Accuracy dominates: 1% accuracy drop = 0.02 loss; 5x cost increase
    = ~0.08 loss. Researcher cannot trade accuracy for cost cheaply.
  - No saturation: $0.30 vs $5 differ by log(5/0.3)*0.05 = 0.14, still
    a real signal. Old loss capped at 0.6 for both.
  - No per-task budget. Reference points (0.20 USD, 60s) are session-
    global, not benchmark-specific. Production safe.
  - No fail_bonus. Failed runs already get accuracy=0 → contribute 2.0
    on the accuracy axis. Adding extra was double-counting.

Cost is *cold-equivalent* — computed from token counts alone so the
Anthropic prompt cache state doesn't leak into the signal.

Corpus loss = arithmetic mean of per-run loss across (task × seed).
"""
from __future__ import annotations

import math
from typing import Any

from benchmarks.experiments.pricing import resolve as resolve_rates


# Loss weights and references.
ACC_WEIGHT = 2.0           # accuracy contribution to loss is in [0, 2.0]
COST_LOG_WEIGHT = 0.05     # cost_loss = max(0, 0.05 * ln(cost / 0.20))
COST_REFERENCE_USD = 0.20  # log anchor — runs cheaper than this contribute zero cost loss
TIME_LOG_WEIGHT = 0.05     # time_loss = max(0, 0.05 * ln(wall / 60))
TIME_REFERENCE_S = 60.0    # log anchor — runs faster than this contribute zero time loss

# Per-tier corpus weighting. A 0.1 reduction on t2 is worth ~2.5× a 0.1
# reduction on t0 in the corpus mean. Pulls the Researcher's gradient
# toward architectural fixes that help t2 (pipeline-level: Planner/Oracle/
# Builder fragility) over micro-tweaks on already-easy t0/t1 tasks.
TIER_WEIGHTS = {0: 1.0, 1: 1.5, 2: 2.5}


def _tier_from_task_id(task_id: str | None) -> int:
    """Extract tier from a task_id like 't0_npv', 't1_returns_table',
    't2_lbo_mini', 't0ref_*'. Falls back to 1 (the median) if unknown."""
    if not task_id:
        return 1
    if task_id.startswith("t0"):
        return 0
    if task_id.startswith("t1"):
        return 1
    if task_id.startswith("t2"):
        return 2
    return 1


def tier_weight(task_id: str | None) -> float:
    return TIER_WEIGHTS.get(_tier_from_task_id(task_id), 1.0)


def cold_cost_usd(usage: dict[str, Any], model: str) -> float:
    """Cold-equivalent USD cost: every input-class token priced at the plain input rate.

    Deterministic given token counts — independent of whether Anthropic's
    prompt cache happened to be warm when this run executed. The three
    token buckets (regular input, cache_creation, cache_read) all roll
    into a single input-rate sum.
    """
    rates = resolve_rates(model)
    total_input = (
        (usage.get("input_tokens") or 0)
        + (usage.get("cache_creation_input_tokens") or 0)
        + (usage.get("cache_read_input_tokens") or 0)
    )
    total_output = usage.get("output_tokens") or 0
    return total_input * rates.input_per_token + total_output * rates.output_per_token


def cold_cost_from_result(result: dict) -> float:
    """Sum cold cost across builder + planner from a `result.json` blob."""
    builder_model = result.get("builder_model") or "sonnet"
    builder_usage = result.get("builder_usage") or {}
    planner_stats = result.get("planner_stats") or {}
    planner_model = planner_stats.get("planner_model") or builder_model
    # planner_stats aggregates usage into flat keys; map to usage dict.
    planner_usage = {
        "input_tokens": planner_stats.get("planner_input_tokens") or 0,
        "output_tokens": planner_stats.get("planner_output_tokens") or 0,
        # The current headless_builder doesn't record planner cache
        # buckets separately. Treat absent as zero — will under-report
        # slightly vs. true cold cost, but consistently across runs.
        "cache_creation_input_tokens": planner_stats.get("planner_cache_creation_input_tokens") or 0,
        "cache_read_input_tokens": planner_stats.get("planner_cache_read_input_tokens") or 0,
    }

    cost = cold_cost_usd(builder_usage, builder_model)
    if any(planner_usage.values()):
        cost += cold_cost_usd(planner_usage, planner_model)

    # Oracle LLM fallback cost, already in real dollars (not token-based).
    oracle = planner_stats.get("oracle") or {}
    cost += oracle.get("cost_dollars") or 0.0
    return cost


def per_run_loss(result: dict) -> dict[str, float]:
    """Compute the per-run loss and its components. Returns a dict for logging."""
    accuracy = result.get("accuracy") or 0.0
    wall = result.get("wall_seconds") or 0.0
    cost_cold = cold_cost_from_result(result)

    acc_loss = ACC_WEIGHT * (1.0 - accuracy)
    # Log-scale, monotonically increasing past the reference point. Cheaper
    # than reference contributes zero (max(0, ...)) so we don't reward
    # implausibly tiny costs that come from instrumentation gaps.
    cost_loss = max(0.0, COST_LOG_WEIGHT * math.log(max(cost_cold, 0.01) / COST_REFERENCE_USD))
    time_loss = max(0.0, TIME_LOG_WEIGHT * math.log(max(wall, 1.0) / TIME_REFERENCE_S))

    return {
        "loss": acc_loss + cost_loss + time_loss,
        "acc_loss": acc_loss,
        "cost_loss": cost_loss,
        "time_loss": time_loss,
        "cost_cold_usd": cost_cold,
        "cost_actual_usd": result.get("dollars") or 0.0,
    }


def corpus_loss(results: list[dict]) -> dict[str, Any]:
    """Aggregate per-run losses into a tier-weighted mean plus diagnostics.

    corpus_loss = Σ(w_i × loss_i) / Σ(w_i), where w_i = TIER_WEIGHTS[tier_i].
    Without weighting, t2 fragility (acc=0.3, loss=1.6) gets averaged
    with t0 perfection (acc=1.0, loss=0) and the corpus number understates
    where the real headroom is.
    """
    if not results:
        return {"corpus_loss": float("inf"), "n": 0, "per_run": []}
    per = []
    for r in results:
        row = per_run_loss(r)
        row["task_id"] = r.get("task_id")
        row["seed"] = r.get("seed")
        row["accuracy"] = r.get("accuracy")
        row["wall_seconds"] = r.get("wall_seconds")
        row["completed"] = r.get("completed")
        row["tier_weight"] = tier_weight(r.get("task_id"))
        per.append(row)
    weight_sum = sum(p["tier_weight"] for p in per) or 1.0
    weighted_loss = sum(p["loss"] * p["tier_weight"] for p in per) / weight_sum
    return {
        "corpus_loss": weighted_loss,
        "n": len(per),
        "mean_accuracy": sum(p["accuracy"] or 0.0 for p in per) / len(per),
        "mean_cost_cold_usd": sum(p["cost_cold_usd"] for p in per) / len(per),
        "mean_cost_actual_usd": sum(p["cost_actual_usd"] for p in per) / len(per),
        "mean_wall_seconds": sum(p["wall_seconds"] or 0.0 for p in per) / len(per),
        "completion_rate": sum(1 for p in per if p["completed"]) / len(per),
        "per_run": per,
    }
