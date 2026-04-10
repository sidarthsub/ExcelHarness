"""Tests for the Session helper that manages per-run state (paths, artifacts)."""
import json
from pathlib import Path

from session import Session


def test_session_creates_directory_structure(tmp_path):
    s = Session(root=tmp_path)
    assert s.run_dir.exists()
    assert s.run_dir.parent == tmp_path / "runs"
    assert s.snapshots_dir.exists()
    assert s.screenshots_dir.exists()


def test_session_writes_brief(tmp_path):
    s = Session(root=tmp_path)
    s.save_brief("build me a model")
    assert (s.run_dir / "brief.md").read_text() == "build me a model"


def test_session_appends_clarifications(tmp_path):
    s = Session(root=tmp_path)
    s.append_clarification("discount_rate", "8%")
    s.append_clarification("fiscal_year_end", "Dec 31")
    body = (s.run_dir / "clarifications.md").read_text()
    assert "discount_rate: 8%" in body
    assert "fiscal_year_end: Dec 31" in body


def test_session_saves_spec(tmp_path):
    s = Session(root=tmp_path)
    spec = {"intent": "x", "sheets": [], "constraints": [], "out_of_scope": []}
    s.save_spec(spec)
    loaded = json.loads((s.run_dir / "model_spec.json").read_text())
    assert loaded == spec


def test_session_append_chat_log(tmp_path):
    s = Session(root=tmp_path)
    s.append_chat("user", "hi")
    s.append_chat("agent", "hello")
    lines = (s.run_dir / "chat_log.jsonl").read_text().strip().split("\n")
    assert len(lines) == 2
    assert json.loads(lines[0])["role"] == "user"
    assert json.loads(lines[0])["text"] == "hi"
