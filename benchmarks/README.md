# ExcelHarness benchmarks

Task suite + grading + automated Q&A for evaluating Excel-building agents in a
Karpathy-style autoresearch loop.

## Layout

```
benchmarks/
  oracle.py           # simulated-user agent for Planner clarifications
  grader.py           # rubric executor against a candidate .xlsx
  runner.py           # CLI: grade | oracle | list | run
  recalc.py           # soffice shim — forces formula recalc before grading
  pseudo_bridge.py    # xlwings-backed HTTP bridge (drop-in for live bridge_server)
  headless_builder.py # orchestrator: pseudo-bridge + Claude Agent SDK Builder
  schemas/            # JSON schemas for task.yaml and grading.yaml
  cache/              # persisted oracle answers
  runs/               # per-run artifacts: scripts, candidate xlsx, result.json
  tasks/
    t0_option_pool_issuance/      # tier 0 — atomic formula, <30s, <$0.10
    t1_inputs_from_term_sheet/    # tier 1 — single-sheet build, <3m, <$1
    t2_safe_convert_series_a/     # tier 2 — multi-sheet mini, <10m, <$3
    t3_three_statement_model/     # tier 3 — integrated 3-stmt, <30m, <$8
```

Each task folder contains:

| file | who reads it | purpose |
|------|-------------|---------|
| `task.yaml` | runner | tier, budgets, tags |
| `brief.md` | Planner/Builder | the user prompt |
| `context.md` | **Oracle only** | what the simulated user knows |
| `inputs/` | Planner/Builder | reference files (xlsx, md, pdf) |
| `stubs/starter.xlsx` | Builder | pre-built scaffold (Tier 0/1) |
| `clarifications.seed.yaml` | Oracle | pre-baked Q→A for fast path |
| `gold/model.xlsx` | **Grader only** | canonical correct answer |
| `gold/grading.yaml` | Grader | rubric (list of checks with weights) |
| `gen.py` | maintainer | deterministically regenerates stub + gold |

## Running

List tasks:
```bash
python3 -m benchmarks.runner list
```

**End-to-end autoresearch loop iteration** (pseudo-bridge + Builder + grade):
```bash
python3 -m benchmarks.runner run --task t0_option_pool_issuance
```
Spawns the xlwings-backed pseudo-bridge, launches a Sonnet Builder that
writes scripts against `bridge.py` (identical to the live harness), waits
for the "Model complete." sentinel or a time-budget abort, then grades
and emits a single-line JSON result.

Grade an existing candidate workbook:
```bash
python3 -m benchmarks.runner grade \
    --task t2_safe_convert_series_a \
    --candidate runs/<session>/models/model.xlsx
```

Regenerate a task's gold artifacts:
```bash
python3 -m benchmarks.tasks.t3_three_statement_model.gen
```

## Oracle — automated Q&A

The Oracle replaces the human during the Planner's clarification pass.
Resolution order:

1. **Seed match** in `clarifications.seed.yaml` — by `id` or by fuzzy
   phrase in `matches`. No LLM call.
2. **Persistent cache** keyed by `hash(task_id + question_text)`. No LLM call.
3. **LLM fallback** — Sonnet, system-prompted to impersonate the user with
   access to `context.md` only. Cached for future runs.

Safety: the Oracle never sees `gold/`. `context.md` is intent-level — never
cell addresses or formula text. The seed file covers the common questions
so the Oracle is deterministic and free on most runs.

Standalone usage:
```bash
python3 -m benchmarks.oracle \
    --task-dir benchmarks/tasks/t2_safe_convert_series_a \
    --questions /tmp/planner_questions.json
```

The runner also exposes it:
```bash
python3 -m benchmarks.runner oracle --task t2_safe_convert_series_a \
    --questions /tmp/planner_questions.json --json
```

## Grading check types

| type | what it does |
|------|--------------|
| `numeric` | cell value matches expected with `tolerance_abs` / `tolerance_rel` |
| `sheet_exists` | named sheet is present |
| `label_exists` | string appears on a sheet (or anywhere in the workbook) |
| `formula_driven` | target cell contains a formula, not a literal |
| `bs_balances` | Total assets = Total L+E across listed columns (for 3-statement) |
| `hardcode_penalty` | **negative-weight** — triggers if the candidate pastes a literal matching a source cell |
| `value_range` | cell value is in `[min, max]` |
| `range_sum` | sum of a range matches expected |
| `llm_judge` | stub — pass-through until wired up |

Positive weights count toward the accuracy denominator; negative weights
are penalties that subtract from the weighted score when they trigger.
Accuracy = max(0, weighted_score / sum_of_positive_weights).

## Pseudo-bridge + headless Builder

The autoresearch loop uses `pseudo_bridge.py` — an HTTP server with the
**identical API surface** as `bridge_server.py`. The Builder's code is
unchanged; it still calls `bridge.Bridge()` and issues the same commands.
The Builder-facing semantics (`bridge.py`) has been extended to honor a
`BRIDGE_URL` env var so the orchestrator can point it at the local port.

Under the hood, pseudo_bridge drives a **hidden real Excel instance via
xlwings**. The xlsx it produces is a real Excel file — cached formula
values, proper formatting, iterative-calc cache populated — so the
existing downstream pipeline (`snapshot_renderer.py`, `live_dump.py`,
evaluator) works on it without modification.

**Important platform quirks encoded in the bridge:**
- xlwings appscript is not thread-safe on macOS → all Excel ops run on a
  dedicated worker thread with a command queue.
- Excel silently hangs when saving to paths under the user's home /
  `~/Documents/...` (sandbox block). The bridge writes to a flat
  `/private/tmp/*.xlsx` and copies to the caller-requested output path on
  every snapshot + on shutdown.
- Subdirectories of `/private/tmp` also fail (only a `~$*.xlsx` lock
  lands). Flat path required.

## Integrating with the live harness

The Oracle is compatible with the Planner's Pass-1 JSON output. Wire it
into `harness_v3.run_planner` by replacing the `ask_user` call in the
clarification loop with a call to `oracle.resolve_questions`, gated by
an env var like `EXCEL_HARNESS_AUTO_QA=1`.

For independent grading, run:
```bash
python3 -m benchmarks.runner grade --task <id> --candidate <xlsx>
```
against the live harness's `runs/<session>/models/model.xlsx`.

## Design principles

- **Gold is generated by Python**, not curated. `gen.py` is the
  specification — re-running it reproduces identical bytes.
- **Grader trusts `soffice` for recalc**. openpyxl cannot compute
  formulas; we shell out to LibreOffice headless before loading values.
- **Context is intent-level**, not implementation-level. The Oracle's
  boundary is "what would the user realistically know."
- **Cache by question hash** so iteration on the Planner doesn't re-pay
  Oracle LLM cost on already-seen questions.
- **No builder coupled in**. The runner grades whatever xlsx you point
  at; the autoresearch loop chooses the builder.

## Adding a new task

```bash
mkdir -p benchmarks/tasks/t0_new_task/{inputs,stubs,gold}
```

Minimum files:
- `task.yaml` — see `schemas/task.schema.json`
- `brief.md`
- `context.md` — intent only; no implementation leaks
- `clarifications.seed.yaml` — pre-bake the obvious Qs
- `gold/model.xlsx` + `gold/grading.yaml` — ideally emitted by a `gen.py`
- any reference files under `inputs/`

Validate:
```bash
python3 -m benchmarks.runner grade --task t0_new_task \
    --candidate benchmarks/tasks/t0_new_task/gold/model.xlsx
```
The gold model must grade at 1.0. If it doesn't, fix either the rubric
or the gen.
