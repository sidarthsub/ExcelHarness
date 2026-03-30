#!/usr/bin/env python3
"""Spreadsheet Harness v2 — three-agent architecture: Planner, Builder, Evaluator."""

import asyncio
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from claude_agent_sdk import query
from claude_agent_sdk.types import (
    AssistantMessage,
    ClaudeAgentOptions,
    PermissionResultAllow,
    PermissionResultDeny,
    ResultMessage,
    StreamEvent,
    TextBlock,
    ThinkingBlock,
    ToolPermissionContext,
    ToolResultBlock,
    ToolUseBlock,
)

HARNESS_ROOT = Path(__file__).parent
AGENTS_DIR = HARNESS_ROOT / "agents"

MAX_EVAL_ROUNDS = 3

INGESTABLE_TEXT = {".txt", ".csv", ".md", ".json", ".tsv"}
INGESTABLE_EXCEL = {".xlsx", ".xls"}
INGESTABLE_IMAGE = {".png", ".jpg", ".jpeg"}
INGESTABLE_PDF = {".pdf"}

# --- Permissions per agent ---
PLANNER_ALLOWED = [
    "Read", "Glob", "Grep",
    "Write(model_spec.json)", "Edit(model_spec.json)",
    "Bash(ls*)", "Bash(cat*)", "Bash(wc*)",
]

SCOPER_ALLOWED = [
    "Read", "Glob", "Grep",
    "Write(scope.json)", "Edit(scope.json)",
    "Bash(ls*)", "Bash(cat*)", "Bash(wc*)",
]

BUILDER_ALLOWED = [
    "Read", "Glob", "Grep",
    "Write(scripts/*)", "Write(models/*)",
    "Edit(scripts/*)", "Edit(models/*)",
    "Bash(python*)", "Bash(ls*)", "Bash(cat*)",
]
BUILDER_DISALLOWED = ["Agent"]

EVALUATOR_ALLOWED = [
    "Read", "Glob", "Grep",
    "Bash(cat*)", "Bash(grep*)", "Bash(ls*)",
]
EVALUATOR_DISALLOWED = ["Agent"]


def read_file(path: Path) -> str:
    return path.read_text() if path.exists() else ""


def load_prompt(agent_name: str) -> str:
    return read_file(AGENTS_DIR / f"{agent_name}.md")


class Run:
    """Encapsulates a single harness run with its own isolated directory."""

    def __init__(self, brief: str, run_dir: Path | None = None):
        self.brief = brief

        if run_dir:
            # Resume an existing run directory
            self.run_dir = Path(run_dir).resolve()
            self.timestamp = self.run_dir.name
        else:
            # Create a new run directory
            self.timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            self.run_dir = HARNESS_ROOT / "runs" / self.timestamp

        self.models_dir = self.run_dir / "models"
        self.evals_dir = self.run_dir / "evals"
        self.input_dir = self.run_dir / "input"
        self.status_path = self.run_dir / "status.json"
        self.progress_path = self.run_dir / "progress.txt"

        # Create/ensure directory structure
        for d in [self.models_dir,
                  self.evals_dir / "formulas", self.evals_dir / "styles",
                  self.evals_dir / "screenshots",
                  self.input_dir]:
            d.mkdir(parents=True, exist_ok=True)

        # Copy shared config into run dir (only if not already there)
        for fname in ["conventions.md", "dump.py"]:
            dst = self.run_dir / fname
            if not dst.exists():
                shutil.copy2(HARNESS_ROOT / fname, dst)

        if not self.progress_path.exists():
            self.progress_path.write_text("")

        # Symlink dashboard for convenience
        dash_src = HARNESS_ROOT / "dashboard.html"
        dash_dst = self.run_dir / "dashboard.html"
        if dash_src.exists() and not dash_dst.exists():
            dash_dst.symlink_to(dash_src)

    def update_status(self, phase: str, detail: str, **extra) -> None:
        status = {
            "updated_at": datetime.now().isoformat(),
            "phase": phase,
            "detail": detail,
            "brief": self.brief,
            "run": self.timestamp,
            "activity": getattr(self, "_activity", []),
            **extra,
        }
        self.status_path.write_text(json.dumps(status, indent=2))
        icon = {"planning": "📋", "generating": "🔨", "dumping": "📸",
                "evaluating": "🔍", "passed": "✅", "failed": "❌",
                "complete": "🏁", "ingesting": "📥", "scoping": "🧭"}.get(phase, "⏳")
        print(f"\n{icon}  [{phase.upper()}] {detail}")

    def log_activity(self, agent: str, kind: str, text: str) -> None:
        """Append an activity entry and flush to status.json."""
        if not hasattr(self, "_activity"):
            self._activity = []
        entry = {"t": datetime.now().isoformat(), "agent": agent, "kind": kind, "text": text}
        self._activity.append(entry)
        # Keep last 50 entries
        self._activity = self._activity[-200:]
        # Live-update status.json
        if self.status_path.exists():
            try:
                status = json.loads(self.status_path.read_text())
                status["activity"] = self._activity
                status["updated_at"] = datetime.now().isoformat()
                self.status_path.write_text(json.dumps(status, indent=2))
            except (json.JSONDecodeError, OSError):
                pass

    def append_progress(self, message: str) -> None:
        with open(self.progress_path, "a") as f:
            f.write(f"[{datetime.now().isoformat()}] {message}\n")

    def run_dump(self) -> None:
        model_path = self.models_dir / "model.xlsx"
        if not model_path.exists():
            sys.exit(f"No model found at {model_path}")
        subprocess.run(
            [sys.executable, str(self.run_dir / "dump.py"), str(model_path)],
            check=True, cwd=str(self.run_dir),
        )


def ingest_inputs(run: Run, input_paths: list[str]) -> dict:
    """Copy provided files into the run's input/ dir and pre-process Excel files.
    Returns a manifest — Planner browses the files itself."""
    manifest = {"files": [], "has_excel": False}

    # Copy any explicitly provided files into the run's input/
    for p in input_paths:
        src = Path(p).resolve()
        if not src.exists():
            print(f"Warning: input file not found: {src}", file=sys.stderr)
            continue
        if src.is_dir():
            for f in sorted(src.iterdir()):
                if not f.name.startswith("."):
                    shutil.copy2(f, run.input_dir / f.name)
        else:
            shutil.copy2(src, run.input_dir / src.name)

    # Now scan run's input/ for everything that landed there
    for f in sorted(run.input_dir.iterdir()):
        if f.name.startswith(".") or f.name.startswith("~$"):
            continue
        suffix = f.suffix.lower()
        manifest["files"].append(f.name)

        if suffix in INGESTABLE_EXCEL:
            manifest["has_excel"] = True
            is_first_xlsx = not any((run.evals_dir / "formulas").iterdir())
            subprocess.run(
                [sys.executable, str(run.run_dir / "dump.py"), str(f)],
                check=True, cwd=str(run.run_dir),
            )
            if is_first_xlsx:
                # Preserve first xlsx's dumps as reference for side-by-side comparison
                ref_dir = run.evals_dir / "reference"
                ref_dir.mkdir(exist_ok=True)
                for subdir in ["screenshots", "styles", "formulas"]:
                    src = run.evals_dir / subdir
                    dst = ref_dir / subdir
                    if src.exists() and not dst.exists():
                        shutil.copytree(src, dst)
            if not (run.models_dir / "model.xlsx").exists():
                shutil.copy2(f, run.models_dir / "model.xlsx")
        elif suffix in INGESTABLE_PDF:
            # Convert PDF to text so agents can read it
            try:
                import pdfplumber
                txt_path = f.with_suffix(".txt")
                with pdfplumber.open(f) as pdf:
                    text = "\n\n".join(page.extract_text() or "" for page in pdf.pages)
                txt_path.write_text(text)
                manifest["files"].append(txt_path.name)
                print(f"  PDF → {txt_path.name} ({len(pdf.pages)} pages)")
            except ImportError:
                print(f"  Warning: pdfplumber not installed, skipping {f.name}", file=sys.stderr)

    # --- Generate planner-friendly concat files ---
    # 1. Sheet manifest: one line per sheet with name + dimensions (for triage)
    manifest_path = run.evals_dir / "sheets.txt"
    if not manifest_path.exists():
        formulas_dir = run.evals_dir / "formulas"
        if formulas_dir.exists():
            lines = []
            for f in sorted(formulas_dir.iterdir()):
                if f.suffix == ".txt":
                    content = f.read_text()
                    rows = content.count("\n")
                    header = content.split("\n")[0] if content else ""
                    cols = header.count("\t")
                    lines.append(f"{f.stem}\t{rows} rows\t{cols} cols")
            manifest_path.write_text("\n".join(lines))

    # Clean up any stale concat files so agents don't accidentally read them
    for stale in ["formulas_all.txt", "styles_all.txt"]:
        p = run.evals_dir / stale
        if p.exists():
            p.unlink()

    return manifest




async def run_agent(prompt: str, system_prompt: str, allowed_tools: list[str],
                    cwd: str, run: "Run", agent_name: str,
                    max_turns: int | None = None,
                    model: str | None = None,
                    disallowed_tools: list[str] | None = None,
                    can_use_tool=None) -> str:
    """Run a Claude agent with scoped permissions and return its text output."""
    def on_stderr(line: str) -> None:
        stripped = line.rstrip()
        if stripped:
            run.log_activity(agent_name, "stderr", stripped)
            # Surface rate-limit and error info to console
            lowered = stripped.lower()
            if any(kw in lowered for kw in ("rate", "limit", "429", "retry", "error", "timeout", "waiting")):
                print(f"  [{agent_name}] ⚠ {stripped}", flush=True)

    opts = ClaudeAgentOptions(
        system_prompt=system_prompt,
        permission_mode="bypassPermissions",
        allowed_tools=allowed_tools,
        disallowed_tools=disallowed_tools or [],
        cwd=cwd,
        max_turns=max_turns,
        model=model,
        stderr=on_stderr,
        include_partial_messages=False,
        extra_args={},
    )
    text_parts = []
    async for msg in query(prompt=prompt, options=opts):
        if isinstance(msg, AssistantMessage):
            for block in msg.content:
                if isinstance(block, TextBlock):
                    text_parts.append(block.text)
                    run.log_activity(agent_name, "text", block.text)
                    print(f"  [{agent_name}] {block.text[:120]}", flush=True)
                elif isinstance(block, ThinkingBlock):
                    # Log truncated thinking for dashboard visibility
                    thinking = block.thinking[:300]
                    run.log_activity(agent_name, "thinking", thinking)
                    print(f"  [{agent_name}] 💭 {thinking[:100]}", flush=True)
                elif isinstance(block, ToolUseBlock):
                    # Summarize tool call
                    inp = block.input
                    if block.name in ("Read", "Glob"):
                        detail = inp.get("file_path") or inp.get("pattern", "")
                    elif block.name == "Grep":
                        detail = f"/{inp.get('pattern', '')}/ in {inp.get('path', '')}"
                    elif block.name == "Bash":
                        cmd = inp.get("command", "")
                        detail = cmd[:120]
                    elif block.name in ("Write", "Edit"):
                        detail = inp.get("file_path", "")
                    else:
                        detail = str(inp)[:120]
                    run.log_activity(agent_name, "tool", f"{block.name}: {detail}")
                    print(f"  [{agent_name}] 🔧 {block.name}: {detail[:100]}", flush=True)
                elif isinstance(block, ToolResultBlock):
                    # Log truncated result
                    content = block.content
                    if isinstance(content, str):
                        preview = content[:150]
                    elif isinstance(content, list):
                        preview = str(content[0])[:150] if content else "(empty)"
                    else:
                        preview = "(no content)"
                    is_err = " ⚠️" if block.is_error else ""
                    run.log_activity(agent_name, "result", f"{preview}{is_err}")
        elif isinstance(msg, ResultMessage):
            if msg.is_error:
                run.log_activity(agent_name, "error", msg.result or "unknown error")
                print(f"  [{agent_name} error: {msg.result}]", file=sys.stderr)
            else:
                cost = f", ${msg.total_cost_usd:.3f}" if msg.total_cost_usd else ""
                run.log_activity(agent_name, "done", f"{msg.num_turns} turns, {msg.duration_ms/1000:.1f}s{cost}")
                print(f"  [{agent_name} done: {msg.num_turns} turns, {msg.duration_ms/1000:.1f}s{cost}]", flush=True)
    return "\n".join(text_parts)


AGENT_COOLDOWN_SECS = int(os.environ.get("HARNESS_COOLDOWN", "10"))


async def _cooldown(run: Run, next_agent: str) -> None:
    """Brief pause between agent calls to avoid rate-limit stalls on Max plans."""
    if AGENT_COOLDOWN_SECS > 0:
        print(f"  ⏸ {AGENT_COOLDOWN_SECS}s cooldown before {next_agent}...", flush=True)
        await asyncio.sleep(AGENT_COOLDOWN_SECS)


# ============================================================
# v2 Three-Agent Architecture: Planner → Builder → Evaluator
# ============================================================


def _parse_eval(result: str) -> dict:
    """Parse evaluator JSON output into logic pass/fail and visual grade."""

    def _extract_json_blob(text: str) -> dict | None:
        stripped = text.strip()
        if stripped.startswith("```"):
            lines = stripped.splitlines()
            if lines and lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            stripped = "\n".join(lines).strip()
        try:
            return json.loads(stripped)
        except json.JSONDecodeError:
            pass

        match = re.search(r"\{.*\}", text, re.DOTALL)
        if not match:
            return None
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            return None

    payload = _extract_json_blob(result) or {}
    if not payload:
        parse_issue = {
            "severity": "critical",
            "summary": "Evaluator did not return valid JSON",
            "details": "The evaluator response could not be parsed into the required JSON schema",
            "fix": "Return JSON only with logic and visual sections",
        }
        return {
            "logic_passed": False,
            "visual_grade": "?",
            "logic_feedback": _issues_to_feedback([parse_issue]),
            "visual_feedback": "",
            "logic_issues": [parse_issue],
            "visual_issues": [],
            "raw": result,
            "parsed": {},
        }
    logic = payload.get("logic") or {}
    visual = payload.get("visual") or {}
    logic_issues = logic.get("issues") if isinstance(logic.get("issues"), list) else []
    visual_issues = visual.get("issues") if isinstance(visual.get("issues"), list) else []

    return {
        "logic_passed": bool(logic.get("passed")),
        "visual_grade": str(visual.get("grade") or "?")[:1].upper(),
        "logic_feedback": _issues_to_feedback(logic_issues),
        "visual_feedback": _issues_to_feedback(visual_issues),
        "logic_issues": logic_issues,
        "visual_issues": visual_issues,
        "raw": result,
        "parsed": payload,
    }


def _issues_to_feedback(issues: list[dict]) -> str:
    lines = []
    for issue in issues:
        severity = str(issue.get("severity") or "warning").upper()
        summary = str(issue.get("summary") or "").strip()
        details = str(issue.get("details") or "").strip()
        fix = str(issue.get("fix") or "").strip()
        parts = [f"[{severity}]"]
        sheet = str(issue.get("sheet") or "").strip()
        if sheet:
            parts.append(f"{sheet}:")
        if summary:
            parts.append(summary)
        line = " ".join(parts).strip()
        if details:
            line += f" | {details}"
        if fix:
            line += f" | Fix: {fix}"
        lines.append(f"- {line}")
    return "\n".join(lines)


MAX_VISUAL_FIX_ROUNDS = 4


async def scope_inputs(run: Run, manifest: dict) -> dict:
    """Scoper: triages input/reference files into a structural scope map."""
    system_prompt = load_prompt("scoper")
    editing = manifest["has_excel"]
    files_list = "\n".join(f"- {f}" for f in manifest["files"])

    prompt = f"""## User Brief
{run.brief}

## Mode: {"EDIT EXISTING MODEL" if editing else "NEW MODEL"}
{"An existing .xlsx was provided and reference data has been extracted to evals/." if editing else "Build from scratch."}

## Input Files (in input/)
{files_list}

## Reference Data (in evals/)
{"Formula dumps: evals/formulas/" if editing else "No reference model provided."}
{"Style dumps: evals/styles/" if editing else ""}
{"Screenshots: evals/screenshots/" if editing else ""}
{"Sheet manifest: evals/sheets.txt" if editing else ""}

Write scope.json to {run.run_dir / "scope.json"}.
Keep it structural and concise. Do not write the build spec.
"""

    run.update_status("scoping", "Triaging inputs and reference scope")
    await _cooldown(run, "scoper")
    await run_agent(prompt, system_prompt, SCOPER_ALLOWED,
                    cwd=str(run.run_dir), run=run, agent_name="scoper",
                    model="claude-haiku-3-5")

    scope_path = run.run_dir / "scope.json"
    if not scope_path.exists():
        print("Scoper failed to generate scope.json", file=sys.stderr)
        sys.exit(1)

    scope = json.loads(scope_path.read_text())
    run.append_progress(f"Scoper done: {len(scope.get('requested_outputs', []))} outputs scoped")
    return scope



async def plan(run: Run, manifest: dict, scope: dict | None = None) -> dict:
    """Planner: brief + inputs → high-level spec."""
    conventions = read_file(run.run_dir / "conventions.md")
    system_prompt = load_prompt("planner_v2")

    files_list = "\n".join(f"- {f}" for f in manifest["files"])
    editing = manifest["has_excel"]
    scope_section = ""
    if scope:
        scope_section = f"""
## Precomputed Scope (in scope.json)
```json
{json.dumps(scope, indent=2)}
```

Read `scope.json` first and use it to guide which files to inspect. Still verify the important files yourself before writing the spec.
"""

    prompt = f"""## User Brief
{run.brief}

{scope_section}

## Mode: {"EDIT EXISTING MODEL" if editing else "NEW MODEL"}
{"An existing .xlsx was provided and is at models/model.xlsx. Reference data has been extracted to evals/." if editing else "Build from scratch."}

## Input Files (in input/)
{files_list}

## Reference Data (in evals/)
{"Formula dumps: evals/formulas/" if editing else "No reference model provided."}
{"Style dumps: evals/styles/" if editing else ""}
{"Screenshots: evals/screenshots/" if editing else ""}
{"Sheet manifest: evals/sheets.txt" if editing else ""}

Browse these files with your tools to understand the reference structure. Use formula/style dumps first. Use screenshots only if the dumps leave layout or visual intent ambiguous. Do NOT open .xlsx files directly.

## Conventions
{conventions}

## Instructions
Write model_spec.json to {run.run_dir / "model_spec.json"}.
Keep it high-level — describe what each sheet should contain, not cell-by-cell formulas.
"""

    run.update_status("planning", "Generating build spec")
    await _cooldown(run, "planner")
    result = await run_agent(prompt, system_prompt, PLANNER_ALLOWED,
                             cwd=str(run.run_dir), run=run, agent_name="planner",
                             model="claude-sonnet-4-6")

    # Load and return the spec
    spec_path = run.run_dir / "model_spec.json"
    if not spec_path.exists():
        print("Planner failed to generate model_spec.json", file=sys.stderr)
        sys.exit(1)

    spec = json.loads(spec_path.read_text())
    run.append_progress(f"Planner done: {len(spec.get('sheets', []))} sheets planned")
    return spec


async def build(run: Run, spec: dict, feedback: str | None = None, model: str | None = None) -> None:
    """Builder: reads spec + reference data, builds all sheets in one session."""
    conventions = read_file(run.run_dir / "conventions.md")
    system_prompt = load_prompt("builder")
    spec_json = json.dumps(spec, indent=2)

    # Ensure scripts dir exists
    scripts_dir = run.run_dir / "scripts"
    scripts_dir.mkdir(exist_ok=True)

    feedback_section = ""
    if feedback:
        feedback_section = f"""
## Evaluator Feedback (fix these issues)
{feedback}

Previous build scripts are in scripts/. Read them to understand what was built, then fix the issues.
"""

    prompt = f"""## Build Spec
```json
{spec_json}
```

## Conventions
{conventions}

## Working Directory
You are in: {run.run_dir}
- Model file: models/model.xlsx
- Reference data: evals/ (formulas, styles, screenshots from input xlsx files)
- Input data: input/ (text files, CSVs, Excel dumps)
- Save scripts to: scripts/ (e.g., scripts/01_SheetName.py)
- Dump tool: `python3 dump.py models/model.xlsx` (extracts formulas/styles/screenshots to evals/)

## Instructions
Build every sheet listed in the spec, in order. For each sheet:
1. Read the reference formula/style dumps if a reference sheet is specified
1a. Use reference screenshots only if formula/style dumps leave layout or visual intent ambiguous
2. Write and run a Python script to build the sheet
3. Run dump.py to extract your output
4. Read your own dump to verify correctness before moving on

After all sheets are built, run dump.py one final time and do a quick self-check across all sheets.
{feedback_section}"""

    phase = "building (fix)" if feedback else "building"
    run.update_status("generating", phase)
    await _cooldown(run, "builder")
    result = await run_agent(prompt, system_prompt, BUILDER_ALLOWED,
                             cwd=str(run.run_dir), run=run, agent_name="builder",
                             disallowed_tools=BUILDER_DISALLOWED, model=model)
    run.append_progress(f"Builder done{' (fix pass)' if feedback else ''}")


async def evaluate(run: Run, spec: dict) -> dict:
    """Evaluator: checks completed model against spec and returns a parsed verdict."""
    system_prompt = load_prompt("evaluator_v2")
    conventions = read_file(run.run_dir / "conventions.md")
    spec_json = json.dumps(spec, indent=2)

    prompt = f"""## Build Spec
```json
{spec_json}
```

## Conventions
{conventions}

## Working Directory
You are in: {run.run_dir}
- Formula dumps: evals/formulas/ (one .txt per sheet — the builder's output)
- Style dumps: evals/styles/
- Screenshots: evals/screenshots/
- Reference data: evals/reference/ (original input model dumps, if applicable)
- Input data: input/

Review the model thoroughly. Check formulas, formatting, cross-sheet references, and completeness against the spec. Use formula/style dumps as the primary evidence. Use screenshots only when layout or visual intent is ambiguous or when you need to confirm a visual concern that the dumps do not settle.
"""

    run.update_status("evaluating", "Final QA review")
    await _cooldown(run, "evaluator")
    result = await run_agent(prompt, system_prompt, EVALUATOR_ALLOWED,
                             cwd=str(run.run_dir), run=run, agent_name="evaluator",
                             disallowed_tools=EVALUATOR_DISALLOWED,
                             model="claude-sonnet-4-6")

    return _parse_eval(result)


async def main():
    import argparse
    parser = argparse.ArgumentParser(description="Spreadsheet Harness v2")
    parser.add_argument("brief", nargs="?", help="User brief describing what to build")
    parser.add_argument("inputs", nargs="*", help="Input file/directory paths")
    parser.add_argument("--run", help="Resume an existing run directory (skip ingest)")
    parser.add_argument("--start-from", choices=["plan", "build", "eval"],
                        default="plan", help="Skip earlier phases (requires --run)")
    args = parser.parse_args()

    if not args.brief and not args.run:
        parser.error("brief is required unless resuming with --run")

    # --- Initialize or resume run ---
    if args.run:
        # Load brief from existing run's status.json if not provided
        brief = args.brief
        if not brief:
            status_path = Path(args.run) / "status.json"
            if status_path.exists():
                brief = json.loads(status_path.read_text()).get("brief", "")
        run = Run(brief, run_dir=args.run)
        print(f"Resuming: {run.run_dir}")
    else:
        run = Run(args.brief)
        print(f"Run: {run.run_dir}")

    start = args.start_from

    # --- Phase 0: Ingest ---
    if start == "plan" and not args.run:
        run.update_status("ingesting", "Processing inputs")
        manifest = ingest_inputs(run, args.inputs)
        run.append_progress(f"Ingested {len(manifest['files'])} files")
    else:
        manifest = {"files": [f.name for f in run.input_dir.iterdir() if not f.name.startswith(".")],
                     "has_excel": any(f.suffix.lower() in INGESTABLE_EXCEL for f in run.input_dir.iterdir())}

    # --- Phase 1: Plan ---
    if start in ("plan",):
        scope = await scope_inputs(run, manifest)
        spec = await plan(run, manifest, scope)
    else:
        spec = json.loads((run.run_dir / "model_spec.json").read_text())
        print(f"Loaded existing spec: {len(spec.get('sheets', []))} sheets")

    # --- Phase 2: Build ---
    if start in ("plan", "build"):
        await build(run, spec)
        run.update_status("dumping", "Extracting final model data")
        try:
            run.run_dump()
        except Exception as e:
            print(f"  dump.py failed: {e}", file=sys.stderr)

    # --- Phase 3: Logic eval → Fix loop ---
    visual_feedback = ""
    for round_num in range(1, MAX_EVAL_ROUNDS + 1):
        verdict = await evaluate(run, spec)
        grade = verdict["visual_grade"]
        visual_feedback = verdict["visual_feedback"]

        run.append_progress(f"Eval round {round_num}: logic={'PASS' if verdict['logic_passed'] else 'FAIL'}, visual={grade}")
        print(f"  Logic: {'PASS' if verdict['logic_passed'] else 'FAIL'} | Visual: {grade}")

        if verdict["logic_passed"]:
            break

        if round_num < MAX_EVAL_ROUNDS:
            await build(run, spec, feedback=verdict["logic_feedback"])
            run.update_status("dumping", f"Re-extracting after logic fix (round {round_num})")
            try:
                run.run_dump()
            except Exception as e:
                print(f"  dump.py failed: {e}", file=sys.stderr)
    else:
        # All logic rounds exhausted without passing
        run.update_status("failed", f"Logic evaluation failed after {MAX_EVAL_ROUNDS} rounds")
        run.append_progress(f"Model failed after {MAX_EVAL_ROUNDS} logic evaluation rounds")
        print(f"\nFailed. Output: {run.run_dir}")
        sys.exit(1)

    # --- Phase 4: Visual fix loop (logic already passed) ---
    if visual_feedback and grade not in ("A",):
        for v_round in range(1, MAX_VISUAL_FIX_ROUNDS + 1):
            run.update_status("generating", f"Visual fix (round {v_round})")
            visual_prompt = f"""## Visual Issues Only — DO NOT change any formulas or data

{visual_feedback}

Read the reference style dumps and screenshots, then write a fix script that ONLY changes formatting:
borders, fills, fonts, column widths, alignment, number formats. Do NOT touch cell values or formulas.
Save the script to scripts/fix_visual_{v_round}.py and run it."""

            await build(run, spec, feedback=visual_prompt, model="claude-sonnet-4-6")
            run.update_status("dumping", f"Re-extracting after visual fix (round {v_round})")
            try:
                run.run_dump()
            except Exception as e:
                print(f"  dump.py failed: {e}", file=sys.stderr)

            # Re-evaluate visuals only
            verdict = await evaluate(run, spec)
            grade = verdict["visual_grade"]
            visual_feedback = verdict["visual_feedback"]
            run.append_progress(f"Visual fix round {v_round}: grade={grade}")
            print(f"  Visual after fix: {grade}")

            if grade in ("A", "B"):
                break

    run.append_progress(f"Model complete (visual grade: {grade})")
    run.update_status("complete", f"Done (visual: {grade}) — model at {run.models_dir / 'model.xlsx'}")
    print(f"\nOutput: {run.run_dir}")


if __name__ == "__main__":
    asyncio.run(main())
