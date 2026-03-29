#!/usr/bin/env python3
"""Adversarial Spreadsheet Harness — orchestrates Planner, Generator, Evaluator agents."""

import asyncio
import json
import os
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

MAX_RETRIES = 5

INGESTABLE_TEXT = {".txt", ".csv", ".md", ".json", ".tsv"}
INGESTABLE_EXCEL = {".xlsx", ".xls"}
INGESTABLE_IMAGE = {".png", ".jpg", ".jpeg"}
INGESTABLE_PDF = {".pdf"}

# --- File-write permissions per agent ---
PLANNER_ALLOWED = [
    "Read", "Glob", "Grep",
    "Write(model_spec.json)", "Write(AGENTS.md)",
    "Edit(model_spec.json)", "Edit(AGENTS.md)",
]

GENERATOR_ALLOWED = [
    "Write(models/*)", "Edit(models/*)",
    "Bash(python*)",
]
GENERATOR_RETRY_ALLOWED = [
    "Read(scripts/*)", "Read(evals/*)",
    "Write(models/*)", "Edit(models/*)",
    "Bash(python*)",
]
GENERATOR_DISALLOWED = ["Agent"]

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
                "complete": "🏁", "ingesting": "📥"}.get(phase, "⏳")
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


async def _generator_tool_filter(tool_name: str, tool_input: dict, ctx: ToolPermissionContext):
    """Only allow python execution and file writes in models/. Block ls, cat, grep, cd, etc."""
    if tool_name == "Bash":
        cmd = tool_input.get("command", "")
        if cmd.strip().startswith("python"):
            return PermissionResultAllow()
        return PermissionResultDeny(reason="Generator can only run python scripts. All data is in the prompt.")
    if tool_name in ("Write", "Edit"):
        return PermissionResultAllow()
    return PermissionResultDeny(reason=f"Tool {tool_name} not allowed for generator.")


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


async def scope(run: Run) -> str:
    """Phase 1: Quick Sonnet agent reads the sheet manifest and brief, returns relevant sheet names."""
    sheets_manifest = read_file(run.evals_dir / "sheets.txt")
    if not sheets_manifest:
        return ""  # No manifest, planner reads everything

    prompt = f"""## User Brief
{run.brief}

## Available Sheets (from reference model)
```
{sheets_manifest}
```

## Task
Based on the brief, identify which sheets from the reference model to include as reference data for the planner.

Rules:
1. Include sheets the user explicitly mentioned or wants to replicate
2. Include sheets that contain data needed for the output
3. **Deduplicate structural variants**: If multiple sheets follow the same structure but with different data (e.g., "Series A" and "Series B" are the same round layout, "Post A Returns" and "Post B Returns" are the same scenario layout), include ONLY the first/simplest one as the template. The planner can extrapolate variants from one example.
4. Ignore separator sheets ("Rounds -->", "Scenarios -->", "From Company -->")

Reply in this exact format — one line per sheet, marking templates:
```
Series A (template for: Series B, Series C)
Post A Returns Scenarios (template for: Post B Returns Scenarios)
Fig Cap Table - November 2025
```

If a sheet is not a template for others, just list its name with no parenthetical."""

    run.update_status("planning", "Phase 1: scoping relevant sheets")
    result = await run_agent(prompt,
                             "You identify which sheets are relevant and deduplicate structural variants. Reply in the specified format.",
                             ["Read", "Glob", "Grep"],
                             cwd=str(run.run_dir), run=run, agent_name="scoper",
                             max_turns=None, model="claude-sonnet-4-6")
    return result.strip()


def _parse_scoper_output(relevant_sheets: str) -> tuple[set, str]:
    """Parse scoper output into template sheet names and the full scoper text.
    Handles both old comma-separated format and new template format.
    Returns (set of sheet names to include in dumps, full scoper text for planner)."""
    if not relevant_sheets:
        return set(), ""

    sheet_names = set()
    for line in relevant_sheets.strip().split("\n"):
        line = line.strip().strip("```").strip("- ")
        if not line:
            continue
        # Extract sheet name (before any parenthetical)
        name = line.split("(")[0].strip().rstrip(",")
        if name:
            sheet_names.add(name)

    # Fallback: if no newlines, try comma-separated
    if not sheet_names:
        sheet_names = {s.strip() for s in relevant_sheets.split(",") if s.strip()}

    return sheet_names, relevant_sheets


def _concat_dump_dir(dump_dir: Path) -> str:
    """Concatenate all .txt files in a dump directory."""
    if not dump_dir.exists():
        return ""
    parts = []
    for f in sorted(dump_dir.iterdir()):
        if f.suffix == ".txt":
            parts.append(f"{'='*60}\nSHEET: {f.stem}\n{'='*60}")
            parts.append(f.read_text())
    return "\n\n".join(parts)


def _build_filtered_dumps(run: "Run", relevant_sheets: str) -> tuple[str, str]:
    """Build filtered formulas and styles dumps containing only template sheets.
    Returns (formulas_text, styles_text)."""
    if not relevant_sheets:
        # No scoping — include all individual dump files
        return _concat_dump_dir(run.evals_dir / "formulas"), _concat_dump_dir(run.evals_dir / "styles")

    sheet_names, _ = _parse_scoper_output(relevant_sheets)

    def filter_dump_dir(subdir: str) -> str:
        parts = []
        dump_dir = run.evals_dir / subdir
        if not dump_dir.exists():
            return ""
        for f in sorted(dump_dir.iterdir()):
            if f.suffix != ".txt":
                continue
            # Match by stem (filename without .txt)
            if f.stem in sheet_names:
                parts.append(f"{'='*60}\nSHEET: {f.stem}\n{'='*60}")
                parts.append(f.read_text())
        return "\n\n".join(parts)

    return filter_dump_dir("formulas"), filter_dump_dir("styles")


def _strip_irrelevant_sheets(run: "Run", relevant_sheets: str) -> None:
    """Remove sheets from models/model.xlsx that aren't in the relevant set."""
    model_path = run.models_dir / "model.xlsx"
    if not model_path.exists() or not relevant_sheets:
        return

    import openpyxl
    sheet_names, _ = _parse_scoper_output(relevant_sheets)
    wb = openpyxl.load_workbook(model_path)

    to_remove = [s for s in wb.sheetnames if s not in sheet_names]
    if not to_remove:
        return

    for name in to_remove:
        del wb[name]

    wb.save(model_path)
    print(f"  Stripped {len(to_remove)} irrelevant sheets from model.xlsx, kept: {wb.sheetnames}", flush=True)

    # Also strip irrelevant screenshots, formulas, and styles dumps
    removed_files = 0
    for subdir in ["screenshots", "formulas", "styles"]:
        dump_dir = run.evals_dir / subdir
        if not dump_dir.exists():
            continue
        for f in dump_dir.iterdir():
            if f.stem not in sheet_names and not f.name.startswith("."):
                f.unlink()
                removed_files += 1
    if removed_files:
        print(f"  Stripped {removed_files} irrelevant dump files", flush=True)


MAX_CONTRACT_REVIEW_ROUNDS = 2


async def review_contracts(run: Run, spec: dict) -> dict:
    """Adversarial contract review — Opus agent checks for contradictions, then planner fixes."""
    conventions = read_file(run.run_dir / "conventions.md")
    system_prompt = load_prompt("contract_reviewer")
    spec_json = json.dumps(spec, indent=2)

    for round_num in range(1, MAX_CONTRACT_REVIEW_ROUNDS + 1):
        run.update_status("planning", f"Phase 3: contract review (round {round_num})")

        prompt = f"""## Model Spec
```json
{spec_json}
```

## Conventions
{conventions}

## User Brief
{run.brief}

Review every sprint contract for errors. Be thorough and adversarial."""

        await _cooldown(run, "contract_reviewer")
        result = await run_agent(prompt, system_prompt, [],
                                 cwd=str(run.run_dir), run=run,
                                 agent_name="contract_reviewer",
                                 max_turns=None, model="claude-opus-4-6")
        print(result)

        if _check_passed(result):
            run.append_progress(f"Contract review PASSED (round {round_num})")
            return spec

        # Contracts have issues — send feedback to planner to fix
        run.append_progress(f"Contract review FAILED (round {round_num}), sending to planner for fixes")
        run.update_status("planning", f"Phase 3b: planner fixing contracts (round {round_num})")

        fix_prompt = f"""## Contract Review Feedback
The adversarial reviewer found these issues with your sprint contracts:

{result}

## Current model_spec.json
```json
{spec_json}
```

## Reference Data
If the reviewer asks you to verify against the reference model, the Fig reference data is at:
- `{run.evals_dir}/formulas/` — per-sheet formula dumps
- `{run.evals_dir}/styles/` — per-sheet style dumps
Read these with your tools to resolve any ambiguities.

Fix ONLY the issues identified above. Do not change anything else.
Write the corrected model_spec.json to {run.run_dir / "model_spec.json"}."""

        await _cooldown(run, "planner (fix)")
        await run_agent(fix_prompt, load_prompt("planner"),
                        PLANNER_ALLOWED,
                        cwd=str(run.run_dir), run=run, agent_name="planner",
                        max_turns=None, model="claude-sonnet-4-6")

        spec_path = run.run_dir / "model_spec.json"
        spec = json.loads(spec_path.read_text())
        spec_json = json.dumps(spec, indent=2)

    run.append_progress(f"Contract review completed after {MAX_CONTRACT_REVIEW_ROUNDS} rounds")
    return spec


async def plan(run: Run, input_manifest: dict) -> dict:
    conventions = read_file(run.run_dir / "conventions.md")
    system_prompt = load_prompt("planner")

    editing = input_manifest["has_excel"]
    files = input_manifest["files"]

    # Phase 1: Quick scope to identify relevant sheets
    relevant_sheets = await scope(run)

    # Phase 1.5: Strip irrelevant sheets from model.xlsx
    if editing:
        _strip_irrelevant_sheets(run, relevant_sheets)

    # Phase 1.6: Build filtered dumps with only relevant sheets
    filtered_formulas, filtered_styles = _build_filtered_dumps(run, relevant_sheets)
    formulas_size = len(filtered_formulas.encode())
    styles_size = len(filtered_styles.encode())
    print(f"  Filtered dumps: formulas={formulas_size//1024}KB, styles={styles_size//1024}KB "
          f"(from {relevant_sheets or 'all sheets'})", flush=True)

    # Inline dumps if they fit, otherwise tell planner to browse filtered files
    INLINE_MAX = 30_000  # ~30KB — inline small dumps, let planner browse larger ones
    if not filtered_formulas and not filtered_styles:
        data_section = """
## Reference Model
No reference model was provided. Design the sheet structure, formatting, and formulas from scratch based on the brief and conventions. Use professional financial model defaults (Calibri 11, gridlines off, accounting number formats, section borders)."""
    elif formulas_size + styles_size <= INLINE_MAX:
        data_section = f"""
## Reference Model — Formulas (relevant sheets only)
```
{filtered_formulas}
```

## Reference Model — Styles (relevant sheets only)
```
{filtered_styles}
```

All reference data is inlined above. Do NOT read evals/ files — everything you need is here."""
    else:
        # Write filtered files for planner to browse
        (run.evals_dir / "formulas_filtered.txt").write_text(filtered_formulas)
        (run.evals_dir / "styles_filtered.txt").write_text(filtered_styles)
        data_section = f"""
## Reference Model Data
Filtered dumps (relevant sheets only) are at:
- `{run.evals_dir}/formulas_filtered.txt` ({formulas_size//1024}KB)
- `{run.evals_dir}/styles_filtered.txt` ({styles_size//1024}KB)
Read these files. Do NOT read formulas_all.txt or styles_all.txt — they contain irrelevant sheets."""

    # Inline text files from input/
    input_texts = []
    for f in sorted(run.input_dir.iterdir()):
        if f.suffix in INGESTABLE_TEXT and not f.name.startswith("."):
            content = f.read_text()
            if len(content.encode()) < 20_000:
                input_texts.append(f"### {f.name}\n```\n{content}\n```")

    input_section = ""
    if files:
        input_section = f"""
## Input Materials
The user provided these files in `{run.input_dir}/`:
{chr(10).join(f"- {f}" for f in files)}
{"- `evals/screenshots/` — rendered PNGs per sheet" if editing else ""}"""
        if input_texts:
            input_section += "\n\n" + "\n\n".join(input_texts)

    scope_section = ""
    if relevant_sheets:
        scope_section = f"""
## Relevant Sheets (identified by pre-scan)
{relevant_sheets}

Note: Where a sheet is marked as a "template for" variants, only the template's dumps are included below. Use the template as a structural guide for variants but adapt to the brief — variants may have different investors, columns, terms, or structure. The template is a starting point, not an exact spec."""

    prompt = f"""## User Brief
{run.brief}

{"## Mode: EDIT EXISTING MODEL" if editing else "## Mode: NEW MODEL"}
{"The user provided an existing .xlsx. Plan targeted sprints to modify it — do NOT rebuild from scratch. The existing model is already at models/model.xlsx." if editing else "Build from scratch."}
{scope_section}{data_section}{input_section}
## Conventions
{conventions}

## Instructions
1. Analyze the reference data above to understand the model structure and formatting
2. Generate model_spec.json and AGENTS.md
Write model_spec.json to {run.run_dir / "model_spec.json"} and AGENTS.md to {run.run_dir / "AGENTS.md"}.
Follow the output format specified in your instructions exactly."""

    run.update_status("planning", f"Phase 2: deep planning", mode="edit" if editing else "new")
    await _cooldown(run, "planner")
    result = await run_agent(prompt, system_prompt, PLANNER_ALLOWED,
                             cwd=str(run.run_dir), run=run, agent_name="planner",
                             max_turns=None)
    print(result)

    spec_path = run.run_dir / "model_spec.json"
    if not spec_path.exists():
        sys.exit("Planner failed to generate model_spec.json")
    spec = json.loads(spec_path.read_text())
    run.append_progress(f"Planner complete: {len(spec.get('sprints', []))} sprints planned")

    # Phase 3: Adversarial contract review
    spec = await review_contracts(run, spec)

    run.update_status("planning", "Done", sprints=len(spec.get("sprints", [])),
                      sheets=[s["sheet"] for s in spec.get("sprints", [])])
    return spec


def _trim_spec_for_sprint(spec: dict, sprint: dict) -> dict:
    """Return a trimmed spec with only the current sprint's sheet + its dependencies."""
    current_sheet = sprint["sheet"]
    # Find dependencies for current sprint
    dep_names = set()
    for s in spec.get("sheets", []):
        if s["name"] == current_sheet:
            dep_names = set(s.get("dependencies", []))
            break

    relevant_names = {current_sheet} | dep_names
    trimmed = {
        "project_name": spec.get("project_name", ""),
        "formatting": spec.get("formatting", {}),
        "sheets": [s for s in spec.get("sheets", []) if s["name"] in relevant_names],
        "current_sprint": sprint,
    }
    return trimmed


GENERATOR_TEMPLATE = '''#!/usr/bin/env python3
"""Sprint {sprint_num}: {sheet_name}"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Load existing workbook (or create new for sprint 1)
wb = openpyxl.load_workbook("{model_path}")
# wb = openpyxl.Workbook()  # Uncomment for sprint 1 of a NEW model

# Get or create the target sheet
if "{sheet_name}" in wb.sheetnames:
    ws = wb["{sheet_name}"]
else:
    ws = wb.create_sheet("{sheet_name}")

# --- Workbook settings ---
wb.calculation.iterate = True
wb.calculation.iterateCount = 100
wb.calculation.iterateDelta = 0.001
ws.sheet_view.showGridLines = False

# --- YOUR CODE HERE ---
# Set column widths, write data/formulas, apply formatting
# Remember: formulas as strings, e.g. ws["C5"].value = "=\'Detailed Cap Table\'!B5"

# --- Save ---
wb.save("{model_path}")
print("Saved to {model_path}")
'''


async def generate(run: Run, spec: dict, sprint: dict, feedback: str | None = None) -> None:
    conventions = read_file(run.run_dir / "conventions.md")
    system_prompt = load_prompt("generator")
    is_retry = feedback is not None
    trimmed_spec = _trim_spec_for_sprint(spec, sprint)

    feedback_section = ""
    if feedback:
        current_formulas = _read_sprint_dump(run, sprint["sheet"], "formulas") or ""
        feedback_section = f"""
## Evaluator Feedback (fix ONLY these issues)
{feedback}

## Current State of {sprint['sheet']}
```
{current_formulas}
```
""" if current_formulas else f"""
## Evaluator Feedback (fix ONLY these issues)
{feedback}
"""

    # Find the template sprint's script if this sprint has dependencies
    scripts_dir = run.run_dir / "scripts"
    prior_scripts = ""
    if scripts_dir.exists():
        script_files = sorted(scripts_dir.glob("sprint_*.py"))
        if script_files:
            # Find the most relevant prior script (dependency or same-structure template)
            dep_names = set()
            for s in spec.get("sheets", []):
                if s["name"] == sprint["sheet"]:
                    dep_names = set(s.get("dependencies", []))
                    break

            # Inline the most recent dependency's script as a template
            template_script = None
            for sf in reversed(script_files):
                for dep in dep_names:
                    if dep.replace(" ", "_") in sf.name or dep.replace(" ", "") in sf.name:
                        template_script = sf
                        break
                if template_script:
                    break

            if template_script and template_script.stat().st_size < 30_000:
                content = template_script.read_text()
                prior_scripts = f"""
## Template Script ({template_script.name})
This sprint's structure is based on a prior sprint. Here is the working script that passed evaluation — adapt it, do not rewrite from scratch:
```python
{content}
```
"""
            else:
                names = [f.name for f in script_files]
                prior_scripts = f"\n## Prior sprint scripts (in {scripts_dir}/ — read with tools if helpful)\n{chr(10).join(f'- {n}' for n in names)}\n"

    existing_section = f"""
## Existing Sheets
Formulas and styles for completed sheets are in evals/formulas/ and evals/styles/. Read with tools if needed for cross-sheet references. Do NOT open .xlsx files."""

    prompt = f"""## Spec
```json
{json.dumps(trimmed_spec, indent=2)}
```

## Conventions
{conventions}

## Sprint {sprint['number']}: {sprint['sheet']}
{chr(10).join(f"- {item}" for item in sprint['contract'])}
{feedback_section}{prior_scripts}{existing_section}
Model path: {run.models_dir / 'model.xlsx'} (ABSOLUTE path — use this exact path, do not cd elsewhere)
Input files are in: {run.input_dir}/
{"Write one script. Execute it. Do not read files first." if not is_retry else "Fix ONLY the issues above. Read the current formulas dump in evals/formulas/ to understand what exists, then write a targeted patch script. Do NOT rewrite from scratch — load the workbook and fix only the broken cells."}"""

    run.update_status("generating", f"Sprint {sprint['number']}: {sprint['sheet']}{' (retry)' if is_retry else ''}",
                      sprint=sprint["number"], retry=is_retry)
    await _cooldown(run, f"generator (sprint {sprint['number']})")
    result = await run_agent(prompt, system_prompt, GENERATOR_ALLOWED,
                             cwd=str(run.run_dir), run=run, agent_name="generator",
                             max_turns=None,
                             model="claude-sonnet-4-6",
                             disallowed_tools=GENERATOR_DISALLOWED)
    print(result)


INLINE_DUMP_MAX_BYTES = 30_000  # If dump exceeds this, don't inline — let agent browse


def _check_passed(result: str) -> bool:
    """Check if evaluator result indicates PASS. Checks first and last 5 lines."""
    lines = result.strip().split("\n")
    check_lines = lines[:5] + lines[-5:]
    for line in check_lines:
        stripped = line.strip().strip("#").strip("*").strip("-").strip()
        if stripped.startswith("PASS"):
            return True
        if stripped.startswith("FAIL"):
            return False
    return False  # No clear verdict = fail


def _read_sprint_dump(run: "Run", sheet_name: str, subdir: str) -> str | None:
    """Read a dump file and return its content if small enough to inline."""
    dump_path = run.evals_dir / subdir / f"{sheet_name}.txt"
    if not dump_path.exists():
        return None
    content = dump_path.read_text()
    if len(content.encode()) > INLINE_DUMP_MAX_BYTES:
        return None  # Too large — evaluator should browse with tools
    return content


async def evaluate_formula(run: Run, spec: dict, sprint: dict) -> tuple[bool, str]:
    """Run the formula evaluator — checks logic, hardcodes, cross-refs, errors."""
    conventions = read_file(run.run_dir / "conventions.md")
    system_prompt = load_prompt("evaluator")

    # Inline the sprint's formulas dump if small enough
    formulas = _read_sprint_dump(run, sprint["sheet"], "formulas")
    inline_section = ""
    if formulas:
        inline_section = f"""
## Formulas Dump (inlined for {sprint['sheet']})
```
{formulas}
```
"""
    else:
        inline_section = f"""
## Formulas
Browse evals/formulas/ at {run.evals_dir}/formulas/ using your tools."""

    prompt = f"""## Sprint Under Review: Sprint {sprint['number']} — {sprint['sheet']}

### Sprint Contract
{chr(10).join(f"- {item}" for item in sprint['contract'])}

## Conventions
{conventions}
{inline_section}
Focus on formula correctness ONLY. Do not check formatting.
{"The formulas dump is inlined above — verify every contract item against it." if formulas else "Browse the files using your tools. Be thorough and adversarial."}
You may still use tools to read other sheets' formulas for cross-sheet verification.
Your response MUST begin with the word PASS or FAIL on the very first line, before any other text."""

    await _cooldown(run, "eval_formula")
    result = await run_agent(prompt, system_prompt, EVALUATOR_ALLOWED,
                             cwd=str(run.run_dir), run=run, agent_name="eval_formula",
                             max_turns=None, model="claude-haiku-4-5-20251001",
                             disallowed_tools=EVALUATOR_DISALLOWED)
    print(result)
    passed = _check_passed(result)
    return passed, result


async def evaluate_visual(run: Run, spec: dict, sprint: dict,
                          prior_result: str | None = None) -> tuple[bool, str]:
    """Run the visual evaluator — checks formatting, styles, aesthetics."""
    conventions = read_file(run.run_dir / "conventions.md")
    system_prompt = load_prompt("evaluator_visual")
    spec_json = json.dumps(spec.get("formatting", {}), indent=2)

    # Inline the sprint's styles dump if small enough
    styles = _read_sprint_dump(run, sprint["sheet"], "styles")
    inline_section = ""
    if styles:
        inline_section = f"""
## Styles Dump (inlined for {sprint['sheet']})
```
{styles}
```
"""
    else:
        inline_section = f"""
## Styles
Browse evals/styles/ at {run.evals_dir}/styles/ using your tools."""

    prior_section = ""
    if prior_result:
        prior_section = f"""
## Prior Visual Review (from last attempt)
The generator was given this feedback and claims to have fixed the issues. Verify the fixes landed and check for regressions.
```
{prior_result[:3000]}
```
"""

    prompt = f"""## Sprint Under Review: Sprint {sprint['number']} — {sprint['sheet']}

### Sprint Contract
{chr(10).join(f"- {item}" for item in sprint['contract'])}

## Target Formatting Spec
```json
{spec_json}
```

## Conventions
{conventions}
{inline_section}{prior_section}
Focus on visual/formatting correctness ONLY. Do not check formula logic.
{"The styles dump is inlined above — check every formatting requirement against it." if styles else "Browse the files using your tools. Be thorough and adversarial."}
You MUST also read the screenshots: evals/screenshots/{sprint['sheet']}.png (generated) and evals/reference/screenshots/ (reference model).
Your response MUST begin with the word PASS or FAIL on the very first line, before any other text."""

    await _cooldown(run, "eval_visual")
    result = await run_agent(prompt, system_prompt, EVALUATOR_ALLOWED,
                             cwd=str(run.run_dir), run=run, agent_name="eval_visual",
                             max_turns=None, model="claude-sonnet-4-6",
                             disallowed_tools=EVALUATOR_DISALLOWED)
    print(result)
    passed = _check_passed(result)
    return passed, result


async def run_sprint(run: Run, spec: dict, sprint: dict) -> tuple[bool, asyncio.Task | None]:
    """Run a sprint. Returns (formula_passed, pending_visual_task).
    Formula eval loops until pass. Once formula passes, kicks off visual async
    and returns immediately so caller can start the next sprint."""

    for attempt in range(1, MAX_RETRIES + 1):
        feedback = None if attempt == 1 else last_feedback

        await generate(run, spec, sprint, feedback)

        if not (run.models_dir / "model.xlsx").exists():
            run.update_status("failed", f"Sprint {sprint['number']}: no model.xlsx produced",
                              sprint=sprint["number"], attempt=attempt)
            last_feedback = "Generator did not produce models/model.xlsx. You must write and execute the openpyxl script."
            continue

        run.update_status("dumping", f"Sprint {sprint['number']}: extracting formulas/styles/screenshots",
                          sprint=sprint["number"], attempt=attempt)
        try:
            run.run_dump()
        except subprocess.CalledProcessError as e:
            print(f"  dump.py failed: {e}", file=sys.stderr)
            last_feedback = f"dump.py failed with error: {e}"
            continue

        # Formula eval only — no visual during this loop
        run.update_status("evaluating",
                          f"Sprint {sprint['number']}: {sprint['sheet']} (formula)",
                          sprint=sprint["number"])

        formula_passed, formula_result = await evaluate_formula(run, spec, sprint)

        if not formula_passed:
            run.update_status("failed",
                              f"Sprint {sprint['number']}: {sprint['sheet']} (attempt {attempt}/{MAX_RETRIES})",
                              sprint=sprint["number"], attempt=attempt)
            last_feedback = f"## Formula Evaluator: FAIL\n{formula_result}"
            continue

        # Formula passed — save the generator's script for future sprints
        scripts_dir = run.run_dir / "scripts"
        scripts_dir.mkdir(exist_ok=True)
        # Find the last python script from generator activity
        for entry in reversed(run._activity):
            if entry["agent"] == "generator" and entry["kind"] == "tool" and "Bash:" in entry["text"]:
                cmd = entry["text"].replace("Bash: ", "", 1)
                if "openpyxl" in cmd or "import openpyxl" in cmd.replace("\\n", "\n"):
                    script_path = scripts_dir / f"sprint_{sprint['number']}_{sprint['sheet']}.py"
                    # Extract the actual script from the command
                    if cmd.startswith("python3 << 'EOF'"):
                        script_content = cmd.replace("python3 << 'EOF'\n", "").rsplit("\nEOF", 1)[0]
                    elif cmd.startswith("python3 -c"):
                        script_content = cmd  # save the whole command
                    else:
                        script_content = cmd
                    script_path.write_text(script_content)
                    run.append_progress(f"Saved script: {script_path.name}")
                    break

        run.append_progress(f"Sprint {sprint['number']} ({sprint['sheet']}) formula PASSED on attempt {attempt}")
        run.update_status("passed", f"Sprint {sprint['number']}: {sprint['sheet']} (formula ✓, visual pending)",
                          sprint=sprint["number"], attempt=attempt)
        visual_task = asyncio.create_task(
            evaluate_visual(run, spec, sprint)
        )
        return True, visual_task

    run.append_progress(f"Sprint {sprint['number']} ({sprint['sheet']}) FAILED after {MAX_RETRIES} attempts")
    return False, None


async def main():
    import argparse
    parser = argparse.ArgumentParser(description="Adversarial Spreadsheet Harness")
    parser.add_argument("brief", help="What to build or change")
    parser.add_argument("inputs", nargs="*", help="Input files or directories to ingest")
    parser.add_argument("--run", dest="run_dir", help="Resume an existing run directory")
    parser.add_argument("--skip-planning", action="store_true", help="Reuse existing model_spec.json and AGENTS.md")
    parser.add_argument("--start-sprint", type=int, default=1, help="Start from this sprint number (skip earlier ones)")
    args = parser.parse_args()

    run_dir = Path(args.run_dir) if args.run_dir else None
    run = Run(args.brief, run_dir=run_dir)
    print(f"Run: {run.run_dir}")
    print(f"  input/  → {run.input_dir}")
    print(f"  models/ → {run.models_dir}")
    print()

    # Ingest inputs (skips if files already in run's input/)
    run.update_status("ingesting", "Processing input files")
    input_manifest = ingest_inputs(run, args.inputs)
    if input_manifest["files"]:
        run.update_status("ingesting", f"Found: {', '.join(input_manifest['files'])}")
    else:
        run.update_status("ingesting", "No input files (building from scratch)")

    # Plan (or reuse existing)
    if args.skip_planning:
        spec_path = run.run_dir / "model_spec.json"
        if not spec_path.exists():
            sys.exit("--skip-planning requires an existing model_spec.json in the run directory")
        spec = json.loads(spec_path.read_text())
        run.update_status("planning", f"Reusing existing plan ({len(spec.get('sprints', []))} sprints)")
    else:
        spec = await plan(run, input_manifest)

    # Execute sprints
    sprints = sorted(spec["sprints"], key=lambda s: s["number"])
    total = len(sprints)

    # Skip sprints if --start-sprint specified
    if args.start_sprint > 1:
        skipped = [s for s in sprints if s["number"] < args.start_sprint]
        sprints = [s for s in sprints if s["number"] >= args.start_sprint]
        if skipped:
            run.update_status("generating", f"Skipping sprints {', '.join(str(s['number']) for s in skipped)}")

    MAX_VISUAL_RETRIES = 3

    def _merge_formatting(run: Run, visual_xlsx: Path, sheet_name: str) -> None:
        """Merge formatting from visual copy back into main model.xlsx.
        Copies only cell styles (fonts, fills, borders, number formats, alignment)
        and column widths/row heights. Does NOT touch values or formulas."""
        import openpyxl
        from copy import copy

        main_path = run.models_dir / "model.xlsx"
        wb_main = openpyxl.load_workbook(main_path)
        wb_vis = openpyxl.load_workbook(visual_xlsx)

        if sheet_name not in wb_main.sheetnames or sheet_name not in wb_vis.sheetnames:
            return

        ws_main = wb_main[sheet_name]
        ws_vis = wb_vis[sheet_name]

        # Copy column widths
        for col_letter, dim in ws_vis.column_dimensions.items():
            if dim.width:
                ws_main.column_dimensions[col_letter].width = dim.width

        # Copy row heights
        for row_num, dim in ws_vis.row_dimensions.items():
            if dim.height:
                ws_main.row_dimensions[row_num].height = dim.height

        # Copy gridline setting
        ws_main.sheet_view.showGridLines = ws_vis.sheet_view.showGridLines

        # Copy cell formatting (not values)
        for row in ws_vis.iter_rows(min_row=1, max_row=ws_vis.max_row, max_col=ws_vis.max_column):
            for cell in row:
                main_cell = ws_main[cell.coordinate]
                if cell.font:
                    main_cell.font = copy(cell.font)
                if cell.fill:
                    main_cell.fill = copy(cell.fill)
                if cell.border:
                    main_cell.border = copy(cell.border)
                if cell.alignment:
                    main_cell.alignment = copy(cell.alignment)
                if cell.number_format:
                    main_cell.number_format = cell.number_format

        wb_main.save(main_path)

    async def resolve_visual(run: Run, spec: dict, sprint: dict, visual_task: asyncio.Task) -> None:
        """Wait for visual eval, run fix passes on a copy, merge back. Never fatal."""
        visual_passed, visual_result = await visual_task
        if visual_passed:
            run.append_progress(f"Sprint {sprint['number']} ({sprint['sheet']}) visual PASSED")
            return

        # Create an isolated working directory for visual fixes
        visual_dir = run.run_dir / f"visual_s{sprint['number']}"
        visual_dir.mkdir(exist_ok=True)
        (visual_dir / "models").mkdir(exist_ok=True)
        visual_model = visual_dir / "models" / "model.xlsx"
        shutil.copy2(run.models_dir / "model.xlsx", visual_model)

        # Create a mini Run that points at the visual dir
        # The generator will write to visual_dir/models/model.xlsx
        visual_run = Run.__new__(Run)
        visual_run.brief = run.brief
        visual_run.run_dir = run.run_dir
        visual_run.timestamp = run.timestamp
        visual_run.models_dir = visual_dir / "models"
        visual_run.evals_dir = run.evals_dir
        visual_run.input_dir = run.input_dir
        visual_run.status_path = run.status_path
        visual_run.progress_path = run.progress_path
        visual_run._activity = run._activity

        for v_attempt in range(1, MAX_VISUAL_RETRIES + 1):
            run.update_status("generating",
                              f"Visual fix {v_attempt}/{MAX_VISUAL_RETRIES}: Sprint {sprint['number']}: {sprint['sheet']} (isolated)")

            feedback = f"## Visual Evaluator: FAIL\n{visual_result}\n\n## IMPORTANT: Do NOT add or remove rows/columns. Only modify cell formatting (fonts, fills, borders, number formats, column widths, alignment). Structural changes will break formulas."
            await generate(visual_run, spec, sprint, feedback)

            # Dump from the visual copy for eval
            subprocess.run(
                [sys.executable, str(run.run_dir / "dump.py"), str(visual_model)],
                check=True, cwd=str(run.run_dir),
            )
            visual_passed, visual_result = await evaluate_visual(run, spec, sprint, prior_result=visual_result)
            if visual_passed:
                _merge_formatting(run, visual_model, sprint["sheet"])
                run.run_dump()
                run.append_progress(f"Sprint {sprint['number']} ({sprint['sheet']}) visual fix PASSED on attempt {v_attempt}, merged")
                shutil.rmtree(visual_dir, ignore_errors=True)
                return

        # Merge best effort even if not fully passing
        _merge_formatting(run, visual_model, sprint["sheet"])
        run.run_dump()
        run.append_progress(f"Sprint {sprint['number']} ({sprint['sheet']}) visual fix incomplete after {MAX_VISUAL_RETRIES} attempts — merged best effort")
        shutil.rmtree(visual_dir, ignore_errors=True)

    visual_tasks: list[asyncio.Task] = []

    for i, sprint in enumerate(sprints, 1):
        run.update_status("generating", f"Sprint {sprint['number']}/{total}: {sprint['sheet']}",
                          progress=f"{i}/{total}")
        success, visual_task = await run_sprint(run, spec, sprint)
        if not success:
            # Wait for pending visual work before exiting
            for vt in visual_tasks:
                if not vt.done():
                    await vt
            run.update_status("failed", f"Sprint {sprint['number']} failed after {MAX_RETRIES} retries")
            sys.exit(1)
        if visual_task:
            visual_tasks.append(
                asyncio.create_task(resolve_visual(run, spec, sprint, visual_task))
            )

    # Wait for all visual fix tasks to complete
    if visual_tasks:
        run.update_status("evaluating", "Waiting for visual fixes to complete")
        await asyncio.gather(*[vt for vt in visual_tasks if not vt.done()])

    # Holistic evaluation (if multi-sprint)
    if len(sprints) > 1:
        run.update_status("evaluating", "Holistic review of complete model")
        run.run_dump()
        all_contracts = []
        for s in sprints:
            all_contracts.append(f"Sprint {s['number']} ({s['sheet']}):")
            all_contracts.extend(f"  - {item}" for item in s["contract"])

        cross_checks = spec.get("cross_sheet_checks", [])
        holistic_sprint = {
            "number": "final",
            "sheet": "All Sheets",
            "contract": [
                "All sprint contracts are satisfied (see below)",
                "Cross-sheet references are consistent",
                "No errors (#REF!, #VALUE!, #NAME?, #DIV/0!) anywhere in the model",
                *cross_checks,
                "",
                "Individual sprint contracts:",
                *all_contracts,
            ],
        }
        formula_passed, _ = await evaluate_formula(run, spec, holistic_sprint)
        if formula_passed:
            run.append_progress("Holistic evaluation PASSED")
            run.update_status("complete", "All sprints passed + holistic review")
        else:
            run.append_progress("Holistic evaluation FAILED")
            run.update_status("failed", "Holistic evaluation failed")
            sys.exit(1)

    run.append_progress(f"Model complete: {run.brief}")
    run.update_status("complete", f"Done — model at {run.models_dir / 'model.xlsx'}")
    print(f"\nOutput: {run.run_dir}")


if __name__ == "__main__":
    asyncio.run(main())
