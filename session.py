"""Per-run session state: paths, artifacts, chat log."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path


class Session:
    def __init__(self, root: Path, timestamp: str | None = None,
                 run_dir: Path | None = None):
        """Either pass `timestamp` (path becomes root/runs/timestamp) or
        pass an explicit `run_dir` (used as-is). The latter is what the
        headless harness uses so per-cell run dirs can live under
        benchmarks/runs/ instead of root/runs/.
        """
        self.root = Path(root)
        if run_dir is not None:
            self.run_dir = Path(run_dir)
            self.timestamp = self.run_dir.name
        else:
            self.timestamp = timestamp or datetime.now().strftime("%Y%m%d-%H%M%S")
            self.run_dir = self.root / "runs" / self.timestamp
        self.input_dir = self.run_dir / "input"
        self.snapshots_dir = self.run_dir / "snapshots"
        self.screenshots_dir = self.run_dir / "screenshots"
        for d in (self.run_dir, self.input_dir, self.snapshots_dir, self.screenshots_dir):
            d.mkdir(parents=True, exist_ok=True)

    def save_brief(self, text: str) -> None:
        (self.run_dir / "brief.md").write_text(text)

    def append_clarification(self, question_id: str, answer: str) -> None:
        path = self.run_dir / "clarifications.md"
        with path.open("a") as f:
            f.write(f"- {question_id}: {answer}\n")

    def save_spec(self, spec: dict) -> None:
        (self.run_dir / "model_spec.json").write_text(json.dumps(spec, indent=2))

    def load_spec(self) -> dict:
        return json.loads((self.run_dir / "model_spec.json").read_text())

    def append_chat(self, role: str, text: str) -> None:
        with (self.run_dir / "chat_log.jsonl").open("a") as f:
            f.write(json.dumps({
                "ts": datetime.now().isoformat(),
                "role": role,
                "text": text,
            }) + "\n")
