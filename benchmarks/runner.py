"""Benchmark runner.

Modes:
    grade       — given a candidate xlsx, grade it against a task's rubric
                  and emit a one-line JSON result.
    oracle      — given a questions JSON and a task, produce a CLARIFICATIONS
                  block (stdout) and a stats object.
    list        — list discovered tasks and their tier/budget.

The runner is intentionally builder-agnostic: it does not spawn a Builder.
The autoresearch outer loop hands it a candidate xlsx (produced by whatever
builder variant is being evaluated) and the runner scores it.
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from pathlib import Path

import yaml

from benchmarks.grader import grade
from benchmarks.oracle import OracleConfig, format_clarifications, resolve_questions


BENCH_ROOT = Path(__file__).resolve().parent
TASKS_ROOT = BENCH_ROOT / "tasks"


def load_task(task_id: str) -> dict:
    task_dir = TASKS_ROOT / task_id
    if not task_dir.exists():
        raise FileNotFoundError(f"task {task_id} not found in {TASKS_ROOT}")
    meta = yaml.safe_load((task_dir / "task.yaml").read_text())
    meta["_dir"] = task_dir
    return meta


def list_tasks() -> list[dict]:
    out = []
    if not TASKS_ROOT.exists():
        return out
    for d in sorted(TASKS_ROOT.iterdir()):
        yml = d / "task.yaml"
        if yml.exists():
            meta = yaml.safe_load(yml.read_text())
            meta["_dir"] = str(d)
            out.append(meta)
    return out


# ---- grade mode ------------------------------------------------------------


def cmd_grade(args: argparse.Namespace) -> int:
    task = load_task(args.task)
    task_dir = task["_dir"]
    rubric = task_dir / "gold" / "grading.yaml"
    candidate = args.candidate or (task_dir / "candidate" / "model.xlsx")
    if not Path(candidate).exists():
        print(json.dumps({"error": f"candidate not found: {candidate}"}))
        return 2

    t0 = time.time()
    result = grade(
        candidate_xlsx=Path(candidate),
        grading_yaml=rubric,
        task_dir=task_dir,
        recalc=None if args.recalc is None else args.recalc,
    )
    wall = time.time() - t0

    out = {
        "task_id": task["id"],
        "tier": task["tier"],
        "mode": "grade",
        "accuracy": result["accuracy"],
        "passed": result["passed"],
        "total": result["total"],
        "weighted_score": result["weighted_score"],
        "total_weight": result["total_weight"],
        "grade_wall_seconds": round(wall, 2),
        "time_budget_seconds": task["time_budget_seconds"],
        "cost_budget_dollars": task["cost_budget_dollars"],
        "candidate": str(candidate),
        "checks": result["checks"],
    }
    if args.pretty:
        print(json.dumps(out, indent=2, default=str))
    else:
        print(json.dumps(out, default=str))
    return 0 if result["accuracy"] >= (args.pass_threshold or 0.0) else 1


# ---- oracle mode -----------------------------------------------------------


def cmd_oracle(args: argparse.Namespace) -> int:
    import asyncio
    task = load_task(args.task)
    task_dir = task["_dir"]
    ctx = (task_dir / "context.md").read_text()
    seed = task_dir / "clarifications.seed.yaml"
    budget = (task.get("clarification_budget") or {})
    cfg = OracleConfig(
        task_id=task["id"],
        context_md=ctx,
        seed_path=seed if seed.exists() else None,
        max_questions=budget.get("max_questions", 10),
    )
    questions = json.loads(Path(args.questions).read_text())
    answers, stats = asyncio.run(resolve_questions(questions, cfg))
    block = format_clarifications(answers)
    if args.json:
        print(json.dumps({"clarifications": answers, "stats": stats, "block": block}, indent=2))
    else:
        print(block)
        print("\n# stats")
        print(json.dumps(stats, indent=2))
    return 0


# ---- list mode -------------------------------------------------------------


def cmd_list(args: argparse.Namespace) -> int:
    tasks = list_tasks()
    if args.json:
        print(json.dumps(tasks, indent=2, default=str))
        return 0
    if not tasks:
        print("(no tasks)")
        return 0
    print(f"{'id':40}  {'tier':>4}  {'time_s':>7}  {'$_max':>6}  tags")
    print("-" * 90)
    for t in tasks:
        print(f"{t['id']:40}  {t['tier']:>4}  {t['time_budget_seconds']:>7}  "
              f"{t['cost_budget_dollars']:>6.2f}  {','.join(t.get('tags', []))}")
    return 0


# ---- CLI -------------------------------------------------------------------


def cmd_run(args: argparse.Namespace) -> int:
    """End-to-end: pseudo-bridge + headless Builder + grade."""
    import asyncio
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )
    from datetime import datetime
    from harness import run_session
    task_dir = Path(__file__).resolve().parent / "tasks" / args.task
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = Path(__file__).resolve().parent / "runs" / f"{ts}_{args.task}"
    result = asyncio.run(run_session(
        task_dir=task_dir,
        run_dir=run_dir,
        model=args.model,
        time_budget_seconds=args.time_budget,
        max_turns=args.max_turns,
    ))
    if args.pretty:
        print(json.dumps(result, indent=2, default=str))
    else:
        summary = {k: v for k, v in result.items() if k != "checks"}
        print(json.dumps(summary, default=str))
    acc = result.get("accuracy") or 0.0
    return 0 if acc >= (args.pass_threshold or 0.0) else 1


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="benchmarks.runner")
    sub = p.add_subparsers(dest="cmd", required=True)

    g = sub.add_parser("grade", help="grade a candidate xlsx against a task rubric")
    g.add_argument("--task", required=True)
    g.add_argument("--candidate", type=Path, default=None,
                   help="Path to candidate xlsx. Default: <task>/candidate/model.xlsx")
    g.add_argument("--pretty", action="store_true")
    g.add_argument("--pass-threshold", type=float, default=None)
    g.add_argument("--recalc", dest="recalc", action="store_true", default=None)
    g.add_argument("--no-recalc", dest="recalc", action="store_false")
    g.set_defaults(func=cmd_grade)

    o = sub.add_parser("oracle", help="resolve Planner clarifications via Oracle")
    o.add_argument("--task", required=True)
    o.add_argument("--questions", required=True, type=Path)
    o.add_argument("--json", action="store_true")
    o.set_defaults(func=cmd_oracle)

    l = sub.add_parser("list", help="list discovered tasks")
    l.add_argument("--json", action="store_true")
    l.set_defaults(func=cmd_list)

    r = sub.add_parser("run", help="end-to-end: spawn pseudo-bridge + Builder, then grade")
    r.add_argument("--task", required=True)
    r.add_argument("--model", default="sonnet")
    r.add_argument("--time-budget", type=float, default=None)
    r.add_argument("--max-turns", type=int, default=40)
    r.add_argument("--pretty", action="store_true")
    r.add_argument("--pass-threshold", type=float, default=None)
    r.add_argument("--skip-planner", action="store_true",
                   help="Skip the Planner+Oracle phase; feed the brief to the Builder directly.")
    r.set_defaults(func=cmd_run)

    return p


def main() -> int:
    args = build_parser().parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
