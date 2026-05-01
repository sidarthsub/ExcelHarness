# ExcelHarness

A multi-agent harness that builds Excel workbooks from a natural-language brief. The pipeline:

```
brief.md  ─►  Planner (3 passes, asks clarifying Qs)
                │
                ▼
        spec.json (validated)
                │
                ▼
            Builder (writes Python scripts that drive a hidden Excel
                     instance via xlwings; iterates with Evaluator
                     feedback until done)
                │
                ▼
          model.xlsx
```

Three Claude agents share the loop:

- **Planner** — converts brief + reference files into a validated `model_spec.json`. Asks clarification questions when ambiguous; the user (or an injected `UserChannel` adapter) answers.
- **Builder** — writes one Python build script per sheet, runs it against a hidden Excel instance through `bridge.py`, and emits a `b.checkpoint()` after each sheet.
- **Evaluator** — runs in the background on every checkpoint, dumps the workbook + screenshots, and either passes (no fail file) or writes `eval_fail_<N>.json` with concrete findings the builder must address before declaring done.

The pseudo-bridge drives **real Excel** under the hood (via xlwings on macOS / Windows). The output is a real `.xlsx` file with cached formula values, formatting, and screenshots — not a synthetic flat dump.

## Requirements

- macOS or Windows with Microsoft Excel installed (xlwings drives it)
- Python 3.13
- LibreOffice (`soffice` on `PATH`) for formula recalc
- An Anthropic auth path: either the [Claude Code CLI](https://claude.com/claude-code) signed in, or `ANTHROPIC_API_KEY` in your environment

## Install

```bash
git clone https://github.com/<you>/ExcelHarness
cd ExcelHarness
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## Quickstart

The fastest way to see the loop end-to-end is to give it a deliberately sparse brief — the planner will ask you clarifying questions via stdin, you answer, then the builder runs:

```bash
python3 -m harness --brief-text "Build me a simple NPV calculator with a discount rate, 5 years of cash flows, and the NPV in cell B13."
```

When the planner asks clarifications, you'll see prompts like:

```
  ❓ How should the discount rate be applied — to Y0 or excluded?
     (Excel's NPV function excludes Y0 by convention)
     1. Y1-Y5 only (NPV(rate, B5:B9) + B10)
     2. Y0 included (NPV(rate, B5:B10))
  >
```

Type `1`, `2`, or your own free-text answer. **This is the `InteractiveCLIChannel` in action** — defined in [`harness.py`](harness.py); the planner's Pass 1 questions get routed to it; your answers feed back into Pass 2 of the spec generation.

For a real run with input files:

```bash
python3 -m harness \
  --brief path/to/brief.md \
  --inputs path/to/inputs/ \
  --stub  path/to/starter.xlsx \
  --out   runs/my_first_run
```

### What goes in `--inputs`

A flat directory (no required structure) of reference files the planner and builder should consider. Every file is staged into the run's `input/` folder and processed by type:

| File type | What happens |
|---|---|
| `*.xlsx`, `*.xls` | Dumped via `dump.py`: per-sheet values, formulas, styles, screenshot JPGs. The builder sees the dumps and can call `b.copy_sheet_from_input("/abs/path.xlsx", "SourceSheet", "TargetSheet")` to pull entire sheets into the candidate workbook. |
| `*.pdf` | Extracted to a companion `.txt` so the planner can read it (term sheets, 10-Ks, brief addenda). |
| `*.txt`, `*.md` | Included verbatim in the planner's context. |

Example layout for a DCF run:

```
inputs/
  term_sheet.pdf              # planner reads the extracted text
  historical_financials.xlsx  # builder can copy_sheet_from_input from this
  brief_addendum.md           # extra context the planner sees
```

### What `--stub` does (optional)

A starting workbook the builder copies into the candidate as its **first step**, before building anything else. Use this when the brief assumes a pre-populated state ("inputs are in B3:B6 of the Calc sheet, write the NPV formula in B13"). The harness stages the stub into `input/` (so `dump.py` indexes its values + formulas) and adds an explicit directive in the builder prompt:

> *The workbook starts BLANK. The file `<abs path>` is the stub — use `b.copy_sheet_from_input(...)` to load each of its sheets first, then proceed.*

Skip this flag if you want the workbook built from scratch.

After Q&A the builder runs (~1–10 min depending on workbook complexity), and you get:

```
runs/my_first_run/
  ├── candidate/model.xlsx        ← the built workbook
  ├── model_spec.json             ← what the planner decided
  ├── chat_log.jsonl              ← Q&A transcript
  ├── eval_verdicts/              ← evaluator findings per checkpoint
  └── result.json                 ← summary (cost, wall, status)
```

## Architecture

### File layout

```
harness.py              Entry point. Owns run_session(), CLI, watchdog,
                        UserChannel protocol, InteractiveCLIChannel default.
pseudo_bridge.py        HTTP server that drives a hidden Excel via xlwings.
bridge.py               Client lib that builder scripts import.
recalc.py               LibreOffice-backed xlsx formula recalc.
session.py              Per-run dir layout (runs/<ts>/{input,scripts,...}).
dump.py                 xlsx → text dumps used by the planner context.
live_dump.py            xlsx → JSON dumps used by the evaluator.
snapshot_renderer.py    xlsx → PDF/JPG screenshots for the evaluator.
timing.py               Lightweight mark/log helper.

agents/
  planner_v3.md         Planner system prompt
  builder_v3.md         Builder system prompt
  evaluator_v3.md       Evaluator system prompt

schemas/
  model_spec.schema.json    JSON Schema the planner's spec must validate against

```

(Eval substrate — graders, simulated-user adapters, batch runners — is
intentionally not shipped in this repo. The harness exposes
`grading_yaml` and `user_channel` as extension points; implement them
in your own module if you need them.)

### The UserChannel seam

The planner's Pass 1 emits clarification questions as JSON. The harness routes them through a `UserChannel`:

```python
class UserChannel(Protocol):
    async def __call__(self, questions: list[dict]) -> dict[str, str]: ...
```

The default is `InteractiveCLIChannel` (defined in `harness.py`) — it prompts via stdin/stdout. To wire a different channel (web UI, Slack bot, IDE pane, LLM-simulated user for tests), implement the Protocol and pass it into `run_session(user_channel=...)`.

### Builder ↔ Evaluator loop

After each `b.checkpoint("desc")`:
1. The bridge auto-saves the workbook
2. A background task spawns the Evaluator agent (Sonnet) reading the workbook dumps + screenshots + spec
3. If status=fail, an `eval_fail_<N>.json` lands in the run dir
4. Builder must `Glob` for fail files between sheets and before declaring done; if any exist, it edits the offending script, re-runs, deletes the fail file, and continues
5. If the builder declares done while fail files are pending, the harness defers the sentinel until evals land (capped at 3 retries to prevent infinite loops)

### Watchdog

Three layers of timeout enforcement, fail-safe ordered:
1. Inline turn-loop check (cooperative)
2. `asyncio.wait_for` soft cancel at `time_budget + 30s`
3. OS-level SIGKILL via `_CellWatchdog` at `time_budget + 60s`, targeting subprocesses spawned by this cell only (computed at fire-time, not snapshot-time)

## Programmatic use

```python
import asyncio
from pathlib import Path
from harness import run_session, InteractiveCLIChannel

result = asyncio.run(run_session(
    brief="Build a 5-year DCF for a SaaS company...",
    inputs_dir=Path("./inputs"),       # optional
    run_dir=Path("./runs/dcf_run"),
    user_channel=InteractiveCLIChannel(),  # or your own UserChannel
    stub_xlsx=Path("./template.xlsx"), # optional starting workbook
    model="sonnet",
    time_budget_seconds=1800,
))

print(result["candidate"])  # path to the built model.xlsx
```

## License

(specify)
