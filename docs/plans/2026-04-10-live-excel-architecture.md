# ExcelHarness v3 — Live Excel Architecture

**Date:** 2026-04-10
**Status:** Design approved, not yet implemented
**Branch target:** `v3-live-excel` (off `main`)

## Overview

The harness moves from a headless `openpyxl` pipeline to an always-live Excel session coordinated through an Office.js bridge. The user watches the model build in real time inside Excel, steers via a chat panel in the add-in's taskpane, and never edits cells directly during a build. The workbook is effectively read-only to the user while the Builder is running.

## Guiding principles

1. **Live Excel is the source of truth.** The `.xlsx` on disk is a serialization of the live workbook, not a separate state. Git versions it at every checkpoint.
2. **Simplicity over cleverness.** No parallel agents, no in-memory daemons, no state files that could drift. The workbook is the state; the harness is stateless.
3. **Single interaction surface.** Everything the user types — planning clarifications, build steering, post-build tweaks — goes through the taskpane chat. No terminal-vs-Excel context switching.
4. **Adversarial review preserved.** The Evaluator survives the rewrite. Its backend changes (live queries instead of dumped artifacts) but its role — fresh-context review that catches things the Builder misses — is load-bearing.
5. **The Planner is a product manager, not an engineer.** Its output schema allows intent and constraints, not implementation details. Overprescription is prevented structurally, not by prompting harder.

## What gets deprecated

- `openpyxl` as the build backend
- `dump.py` as a post-build artifact extraction step (replaced by a live `dumpSheet` command and on-demand snapshot rendering)
- Hardcoded sprints as an execution unit (replaced by Builder-declared checkpoints)
- Per-sprint scripts on disk as the authoritative build record (replaced by the git history of the workbook)
- `LibreOffice` as the recalc engine; it remains only as a PNG renderer for evaluator screenshots
- The separate Scoper agent, merged into the Planner

## System components

Two components. Everything else is code inside them.

### 1. The add-in (`taskpane.html`)

JavaScript loaded inside Excel's task pane via sideloaded manifest. Holds one WebSocket to the harness. Responsibilities:

- Dispatch bridge commands received over WSS to `Excel.run` handlers (writeValues, writeFormulas, formatRange, charts, etc.)
- Render a chat UI: scrollable message log, input box, send button, status line
- Send user chat messages to the harness over the same WSS
- Receive agent-emitted chat messages from the harness and render them
- Protect/unprotect the workbook on command so users can't edit during builds

### 2. The harness (`harness.py`)

Single Python process that does everything:

- Serves HTTPS on `localhost:3000` for add-in static files
- Accepts and holds the WSS connection from Excel
- Exposes `POST /api/command` for Builder scripts to send bridge commands in-process (no second hop)
- Maintains an in-memory FIFO of inbound chat messages
- Runs the Claude Agent SDK sessions for Planner, Builder, and Evaluator
- Coordinates the turn-based message delivery loop
- Handles checkpoints: snapshot → LibreOffice render → Evaluator → git commit → release long-poll
- Remains stateless across restarts — recovery works by re-reading the current workbook via `dumpSheet`

### `bridge.py` — supporting utility

A ~50-line HTTP client library that Builder-written scripts import. Not a "component" architecturally — a shared utility like `requests`. One function per bridge command, plus `checkpoint()` and `emit()`.

## Data flow

### Startup
1. User runs `python harness.py` from a terminal. Harness starts HTTPS+WSS on localhost and waits for the add-in.
2. User opens Excel with the sideloaded add-in. Taskpane connects over WSS.
3. User types their brief into the taskpane chat.
4. Harness receives the brief and kicks off the flow.

### Planning phase
5. Planner runs. First pass emits ambiguity questions; each streams to the taskpane chat one at a time and waits for a reply.
6. With answers folded in, Planner produces `model_spec.json`, then does a self-review pass against a critic prompt.
7. Spec is validated against the strict schema. Harness posts a short summary to chat: "Planned 6 sheets: … Starting build."

### Build phase
8. Harness protects the workbook via the bridge. Workbook is read-only to the user.
9. Builder agent session starts. System prompt includes the spec, `conventions.md`, and the `bridge.py` API. First turn: "Begin building per the spec."
10. Builder writes a Python script to a temp file, imports `bridge`, runs it with Bash. Bridge commands POST to `localhost:3000/api/command` (the harness's own endpoint). Cells appear in Excel as the script runs.
11. Script finishes. Builder reads stdout (errors, self-check `readValues` output).
12. Harness drains the chat queue. Any user messages since the last turn get injected as the next user turn.
13. Builder continues. At natural stopping points, the script calls `bridge.checkpoint("Revenue sheet complete")`. That function blocks via HTTP long-poll.

### Checkpoint → Evaluator → continue
14. Harness receives the checkpoint. Saves a snapshot `.xlsx`, runs LibreOffice to render per-sheet PNGs.
15. Evaluator runs in a fresh Claude context with the snapshot artifacts + checkpoint description. Returns PASS or FAIL+findings.
16. On PASS: harness commits the `.xlsx` to git, releases the long-poll with `{status: "pass"}`, and emits a chat status ("✓ Revenue sheet complete, committed").
17. On FAIL: harness releases the long-poll with `{status: "fail", findings: [...]}`. Builder sees findings in its next turn and fixes them.

### End of session
18. Builder emits a final chat message: "Model complete. Ready for review."
19. Harness unprotects the workbook. User can edit cells. Session ends when the user closes the taskpane or types "end session" in chat.

### Key flow properties
- The Builder never reads chat mid-script. Messages only arrive at turn boundaries, matching the SDK's turn model.
- `checkpoint()` is both the commit gate and the review gate — one mechanism.
- Recovery from a mid-build disconnect is automatic: the long-poll errors, the script fails, the Builder sees the error in its next turn, re-reads the workbook via `dumpSheet`, and resumes.

## Key contracts

### Bridge command surface (add-in ↔ harness)
Already implemented in the prototype. Expands with:
- `dumpSheet(sheet)` — returns `{cells: [{address, formula, value, format}], dimensions}` for the Evaluator and for post-disconnect recovery
- `saveSnapshot(path)` — persists the workbook for LibreOffice rendering
- `protectWorkbook()` / `unprotectWorkbook()` — toggle user cell-editing
- `setStatus(text)` — updates the status line in the taskpane
- Chat messages travel as a separate WSS message type (`{type: "chat", direction: "in"|"out", text}`), no HTTP endpoint

### Python bridge client (`bridge.py`)
Thin HTTP wrapper Builder scripts import. One function per command, plus:
- `checkpoint(description)` — blocking call that triggers snapshot → Evaluator → commit; returns `{status: "pass"}` or `{status: "fail", findings: [...]}`
- `emit(text)` — sends an agent-initiated chat message to the user
- No higher-level helpers, no convention encoding. Opus writes the cell-level code.

### Planner's spec schema (`model_spec.json`)
Strict contract. Only allows:
- `intent` — one paragraph, what this model answers
- `sheets[]` — each with `name`, `purpose` (one sentence), `must_contain[]`, `data_sources[]`
- `constraints[]` — hard rules ("use 8% discount rate", "fiscal year ends June 30")
- `out_of_scope[]` — explicit non-goals

Disallowed (will fail schema validation): cell addresses, specific formulas, formatting decisions, chart types, layout directives, row/column counts, anything already in `conventions.md`. The schema has no fields for these, so the Planner physically cannot emit them.

### Session artifacts on disk
```
ExcelHarness/
├── harness.py                    # Single-process orchestrator
├── bridge.py                     # Thin HTTP client for Builder scripts
├── conventions.md                # Unchanged
├── agents/
│   ├── planner.md                # Merged scoper+planner prompt
│   ├── builder.md                # Updated for bridge + long-running
│   └── evaluator.md              # Updated for live queries
├── officejs-prototype/
│   └── addin/
│       └── taskpane.html         # Extended with chat UI + new commands
├── models/
│   └── model.xlsx                # Git-versioned, committed at each checkpoint
├── runs/
│   └── YYYYMMDD-HHMMSS/          # Per-session directory
│       ├── brief.md
│       ├── clarifications.md     # Q&A answers from planning phase
│       ├── model_spec.json       # Validated spec
│       ├── chat_log.jsonl        # Full chat transcript
│       ├── snapshots/            # Per-checkpoint .xlsx snapshots
│       └── screenshots/          # Per-checkpoint PNGs (LibreOffice render)
```

### Files that go away
- `dump.py` (~1000 lines) — replaced by `dumpSheet`
- `deterministic_evaluator.py` — replaced by Evaluator agent querying live state
- `scripts/` — Builder writes to tmp, git history is the build record
- `evals/` — artifacts live per-session under `runs/` or are thrown away
- `progress.txt` — replaced by `chat_log.jsonl` + git commit messages
- `agents/scoper.md` — merged into `agents/planner.md`

## Migration plan

### Phase 0 — Branch setup
- Create `v3-live-excel` off `main`. `v2-three-agent` stays untouched as a fallback.
- Move `officejs-prototype/` out of untracked state and commit it.

### Phase 1 — Harness owns the bridge
- Fold `relay.py` into `harness.py`. Harness serves HTTPS for add-in static files, holds the WSS to Excel, exposes `/api/command` in-process.
- Write `bridge.py` — one function per bridge command, no checkpoint or chat yet.
- Verify end-to-end: a hand-written Python script using `bridge.py` can build a test sheet. This is the current prototype test sequence, translated JS→Python.

### Phase 2 — Chat channel
- Extend `taskpane.html` with chat UI (input, log, status line).
- Add chat message types to WSS. Harness maintains the FIFO.
- At this phase the harness is still a script proxy — no agents. Verify round-trip: type in taskpane, see in harness stdout, and vice versa.

### Phase 3 — Planner + Q&A loop
- Write `agents/planner.md` (merged scoper+planner). Write the strict `model_spec.json` schema + validator.
- Implement planning phase: receive brief over chat, run Planner first pass, stream questions, collect answers, run second pass, self-review, validate, write spec to `runs/<session>/model_spec.json`.
- Verify with a real brief. Output is a validated spec file. No Excel writing yet.

### Phase 4 — Builder loop
- Write `agents/builder.md`. System prompt emphasizes: long-running, writes Python scripts, uses `bridge.py`, calls `bridge.checkpoint()` at natural stopping points, self-checks via `readValues` and `dumpSheet`.
- Implement the loop: run turn → agent writes+runs script → drain chat queue → inject as next turn → repeat.
- Checkpoint mechanism: `bridge.checkpoint()` blocks via long-poll; harness snapshots, runs LibreOffice render, **stubs the evaluator (returns PASS)**, commits, releases. Validates the flow without the evaluator in the loop.
- Build a small model (2-3 sheets) manually verified. Workbook protection enabled during build.

### Phase 5 — Evaluator
- Write `agents/evaluator.md`. Uses live `dumpSheet` queries + rendered PNGs. Fresh context per invocation.
- Wire into checkpoint handler. FAIL responses return findings to the Builder's `checkpoint()` call.
- Test FAIL→retry with a contrived failure.

### Phase 6 — Deletion
- Delete `dump.py`, `deterministic_evaluator.py`, `scripts/`, `evals/`, `progress.txt`, `agents/scoper.md`, old planner/builder versions.
- Merge `v3-live-excel` into `main`.

## Risks

**Phase 1 complexity is the biggest unknown.** Folding the relay into the harness touches asyncio + HTTP + WSS + subprocess management. If it turns into a mess, the fallback is keeping `relay.py` as a separate process and having `harness.py` POST to it. That's still simpler than v2 because there are only two processes and one clear data flow.

**Phase 4 has a subtle error-handling issue.** When a Builder script errors (e.g., bridge call fails mid-build because Excel disconnected), the harness needs to decide whether the agent sees the stderr and retries or the harness kills and restarts the Builder. Default is "agent sees stderr and decides." Watch for infinite retry loops in early testing.

**Phase 5's live-query pattern is a new cognitive task for the Evaluator.** The old `dump.py` gave it a pre-computed grep-friendly corpus. Asking it to drive `dumpSheet` calls and reason is different. May need prompt iteration. Fallback: a `dumpAllSheets()` helper produces a text blob mimicking the old dump format on demand.

**Schema validation is load-bearing.** If the Planner ships without strict schema validation, overprescription creeps back in. Don't skip it.

## Out of scope

- New bridge commands beyond what Section 4 lists — migration uses only the proven surface
- Replacing LibreOffice with native Excel rendering — still blocked on Office.js image export shipping
- Any UI beyond the chat panel — no dashboards, no progress bars beyond a status line, no visual diffs
- Multi-user or multi-workbook sessions
- Any attempt to make the system run headlessly — "always live" is the explicit operating model
