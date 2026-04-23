"""Scalar loss function the autoresearch outer loop minimizes.

Per-run loss blends three signals, each normalized against that task's
own budget (from `task.yaml`) so t0 and t2 contribute on the same scale:

    acc_loss    = 1 - accuracy                            # [0, 1]
    cost_loss   = 0.3 * min(cost_cold / cost_budget,  2)  # [0, 0.6]
    time_loss   = 0.2 * min(wall     / time_budget,  2)   # [0, 0.4]
    fail_bonus  = 0.5 if not completed else 0.0

    loss = acc_loss + cost_loss + time_loss + fail_bonus

Cost is *cold-equivalent* — computed from token counts alone so the
cache state of the Anthropic server doesn't leak into the signal. See
`cold_cost_usd`.

Corpus loss = arithmetic mean of per-run loss across (task × seed).

Weights (0.3 / 0.2 / 0.5) are intentional first-draft numbers. Revisit
once the Researcher has a few runs on the board and the distribution
across axes is visible.
"""
from __future__ import annotations

from typing import Any

from benchmarks.experiments.pricing import resolve as resolve_rates


ACC_WEIGHT = 1.0
COST_WEIGHT = 0.3
TIME_WEIGHT = 0.2
FAIL_BONUS = 0.5
BUDGET_CAP = 2.0  # cap cost/time loss at 2× budget so a single runaway doesn't dominate


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
    completed = bool(result.get("completed"))
    time_budget = result.get("time_budget_seconds") or 1.0

    # The result.json doesn't always carry cost_budget_dollars — fall back
    # to the task's tier default if missing.
    cost_budget = result.get("cost_budget_dollars")
    if cost_budget is None or cost_budget <= 0:
        cost_budget = _default_cost_budget_for_tier(result.get("tier"))

    cost_cold = cold_cost_from_result(result)

    acc_loss = ACC_WEIGHT * (1.0 - accuracy)
    cost_loss = COST_WEIGHT * min(cost_cold / cost_budget, BUDGET_CAP)
    time_loss = TIME_WEIGHT * min(wall / time_budget, BUDGET_CAP)
    fail = FAIL_BONUS if not completed else 0.0

    return {
        "loss": acc_loss + cost_loss + time_loss + fail,
        "acc_loss": acc_loss,
        "cost_loss": cost_loss,
        "time_loss": time_loss,
        "fail_bonus": fail,
        "cost_cold_usd": cost_cold,
        "cost_actual_usd": result.get("dollars") or 0.0,
    }


def corpus_loss(results: list[dict]) -> dict[str, Any]:
    """Aggregate per-run losses into a single scalar plus diagnostics."""
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
        per.append(row)
    mean_loss = sum(p["loss"] for p in per) / len(per)
    return {
        "corpus_loss": mean_loss,
        "n": len(per),
        "mean_accuracy": sum(p["accuracy"] or 0.0 for p in per) / len(per),
        "mean_cost_cold_usd": sum(p["cost_cold_usd"] for p in per) / len(per),
        "mean_cost_actual_usd": sum(p["cost_actual_usd"] for p in per) / len(per),
        "mean_wall_seconds": sum(p["wall_seconds"] or 0.0 for p in per) / len(per),
        "completion_rate": sum(1 for p in per if p["completed"]) / len(per),
        "per_run": per,
    }


def _default_cost_budget_for_tier(tier: int | None) -> float:
    # Matches the README's tier budgets. Defensive only — every task.yaml
    # should specify cost_budget_dollars.
    return {0: 0.10, 1: 1.00, 2: 3.00, 3: 8.00}.get(tier or 0, 1.00)
