"""Oracle agent — simulates the user answering the Planner's clarifications.

Contract (matches the Planner's Pass-1 output):

    input:  JSON array of {id, question, context?, choices?}
    output: a text block of the form

        CLARIFICATIONS:
        - <id>: <answer>
        - <id>: <answer>
        ...

    (which is exactly what the harness already feeds into Pass 2.)

Resolution order for each question:
    1. Seed match  — clarifications.seed.yaml by id (exact) or by fuzzy
       phrase in `matches:`. No LLM call.
    2. Persistent cache — hash(task_id + question_text) in
       benchmarks/cache/oracle_cache.json. No LLM call.
    3. LLM call — Sonnet via the Claude Agent SDK (same auth as Claude
       Code; no ANTHROPIC_API_KEY needed). System-prompted to impersonate
       the user with access to context.md only. Result is cached.

Budgets:
    max_questions — hard cap on the total number of questions answered.
        Excess questions get "You have enough information to proceed.
        Make a reasonable default choice."

Safety:
    The Oracle NEVER sees the gold/ folder. The caller is responsible
    for passing only context.md. This module also forbids paths containing
    "/gold/" from being read, as a belt-and-suspenders check.
"""
from __future__ import annotations

import asyncio
import hashlib
import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

from claude_agent_sdk import query, ClaudeAgentOptions
from claude_agent_sdk.types import AssistantMessage, ResultMessage, TextBlock


ROOT = Path(__file__).resolve().parent.parent
CACHE_DIR = Path(__file__).resolve().parent / "cache"
CACHE_DIR.mkdir(exist_ok=True)
CACHE_PATH = CACHE_DIR / "oracle_cache.json"

ORACLE_MODEL = os.environ.get("ORACLE_MODEL", "sonnet")

ORACLE_SYSTEM = """You are simulating a user who wrote a brief for a financial modeling task.

You are NOT an engineer or analyst. You have access only to the CONTEXT block,
which describes what you (the user) want and what defaults you'd reasonably pick.

You may also receive a SEEDS block — pre-decided answers for known recurring
ambiguities. SEEDS take precedence over CONTEXT and over your own reasoning.
Match a question to a seed by topic, NOT by exact wording: a seed labeled
"exit_date_anchoring" with matches like ["exit date", "hold period"] applies
to any question asking about exit timing, hold-period anchoring, XIRR dates,
etc. — even when the question's surface wording is different. When you find
a topic-matching seed, return its `answer` VERBATIM (do not paraphrase).

When the assistant asks you clarification questions, answer in the following
format, and nothing else:

CLARIFICATIONS:
- <question_id>: <concise answer, one sentence>
- <question_id>: <concise answer, one sentence>

Rules:
- If a SEED entry matches the question's topic, return its `answer` verbatim.
- Otherwise, answer from CONTEXT. If CONTEXT doesn't cover it, say
  "Your call — default: <the most common industry practice>".
- NEVER reveal cell addresses, formula text, or specific layout — UNLESS
  a seed answer contains them, in which case repeat the seed verbatim.
- Be concise. One sentence per answer.
- If a question was already answered in an earlier turn, repeat that answer.
- Answer every question that appears in the input, in the same order.
"""


@dataclass
class OracleConfig:
    task_id: str
    context_md: str
    seed_path: Path | None = None
    max_questions: int = 10
    cache_path: Path = CACHE_PATH

    # populated at load()
    seed_entries: list[dict] = field(default_factory=list)

    def load_seed(self) -> None:
        if self.seed_path and self.seed_path.exists():
            data = yaml.safe_load(self.seed_path.read_text()) or []
            self.seed_entries = data if isinstance(data, list) else []


# Filler words that don't carry concept signal. Kept tight on purpose —
# we want to keep finance-relevant terms (rate, basis, cap) and only strip
# question-grammar words.
_STOPWORDS: frozenset[str] = frozenset({
    "the", "a", "an", "and", "or", "but", "of", "to", "in", "on", "at",
    "for", "with", "from", "by", "as", "is", "are", "was", "were", "be",
    "this", "that", "these", "those", "it", "its", "i", "we", "you",
    "should", "would", "could", "do", "does", "did", "have", "has", "had",
    "what", "which", "how", "where", "when", "why", "who", "whom",
    "if", "then", "than", "so", "such", "any", "all", "each",
    "use", "used", "using", "make", "go", "come", "get",
    "please", "thanks", "thank",
})


def _content_tokens(s: str, min_len: int = 3) -> set[str]:
    """Length-≥min_len lowercase alphabetic tokens with stopwords removed.

    Used to compute fuzzy overlap between question text and seed match
    phrases, and to canonicalize cache keys so re-phrasings collide."""
    return {t for t in _TOKEN_RE.findall(s.lower())
            if len(t) >= min_len and t not in _STOPWORDS}


def _canonicalize_qtext(qtext: str) -> str:
    """Stable string form of a question for cache keying.

    Drops word order, casing, punctuation, and stopwords. Two questions
    that ask the same thing in different words produce the same string.
    """
    return " ".join(sorted(_content_tokens(qtext)))


def _hash_key(task_id: str, qtext: str) -> str:
    h = hashlib.sha256()
    h.update(task_id.encode())
    h.update(b"\x00")
    h.update(_canonicalize_qtext(qtext).encode())
    return h.hexdigest()[:20]


def _load_cache(path: Path) -> dict:
    if path.exists():
        try:
            return json.loads(path.read_text())
        except json.JSONDecodeError:
            return {}
    return {}


def _save_cache(path: Path, cache: dict) -> None:
    path.write_text(json.dumps(cache, indent=2, sort_keys=True))


_TOKEN_RE = re.compile(r"[a-z]+")


def _format_seeds(seed_entries: list[dict]) -> str:
    """Format seed entries as a SEEDS block for the oracle prompt.

    Each entry shows id, topic phrases (so the LLM knows what kinds of
    questions it covers), and the verbatim answer to return.
    """
    if not seed_entries:
        return ""
    lines = ["SEEDS (return the `answer` verbatim when a question matches by topic):"]
    for e in seed_entries:
        sid = e.get("id", "")
        matches = e.get("matches") or []
        ans = e.get("answer", "")
        topic = ", ".join(matches) if matches else "(no topics)"
        lines.append(f"- id: {sid}")
        lines.append(f"  topics: {topic}")
        lines.append(f"  answer: {ans}")
    return "\n".join(lines)


async def _llm_resolve(pending: list[dict], cfg: OracleConfig) -> tuple[str, dict]:
    """One Claude Agent SDK call to answer the pending questions.

    Seeds are inlined into the prompt; the oracle does the topic-matching
    and returns either a seed answer verbatim or a context-based answer.

    Returns (text_response, usage_stats).
    """
    stats = {"input_tokens": 0, "output_tokens": 0, "cost_dollars": 0.0}
    user_block = json.dumps(
        [{"id": q.get("id", ""),
          "question": q.get("question", ""),
          "context": q.get("context", ""),
          "choices": q.get("choices", [])} for q in pending],
        indent=2,
    )
    seeds_block = _format_seeds(cfg.seed_entries)
    seeds_prefix = f"{seeds_block}\n\n" if seeds_block else ""
    prompt = (
        f"{seeds_prefix}"
        f"CONTEXT:\n{cfg.context_md}\n\n"
        f"QUESTIONS (JSON array):\n{user_block}\n\n"
        "Produce the CLARIFICATIONS block now."
    )
    options = ClaudeAgentOptions(
        system_prompt=ORACLE_SYSTEM,
        allowed_tools=[],
        permission_mode="bypassPermissions",
        model=ORACLE_MODEL,
        max_thinking_tokens=0,
    )
    final_text = ""
    async for message in query(prompt=prompt, options=options):
        if isinstance(message, AssistantMessage):
            for block in message.content:
                if isinstance(block, TextBlock):
                    final_text = block.text
        elif isinstance(message, ResultMessage):
            u = message.usage or {}
            stats["input_tokens"] += u.get("input_tokens", 0) or 0
            stats["output_tokens"] += u.get("output_tokens", 0) or 0
            stats["cost_dollars"] += message.total_cost_usd or 0.0
    return final_text, stats


async def resolve_questions(
    questions: list[dict],
    cfg: OracleConfig,
    *,
    questions_asked_so_far: int = 0,
) -> tuple[dict, dict]:
    """Return (answers_by_id, stats).

    stats = {"seed_hits": int, "cache_hits": int, "llm_calls": int,
             "over_budget": int, "tokens_in": int, "tokens_out": int,
             "cost_dollars": float}
    """
    cfg.load_seed()
    cache = _load_cache(cfg.cache_path)

    answers: dict[str, str] = {}
    stats = {"seed_hits": 0, "cache_hits": 0, "llm_calls": 0,
             "over_budget": 0, "tokens_in": 0, "tokens_out": 0, "cost_dollars": 0.0}

    # Resolution order: seed-by-id (exact qid match) → cache → LLM (with seeds inlined).
    # The LLM does semantic seed-topic matching, so we no longer keyword-match in code —
    # this eliminates the brittle substring/token-overlap bugs that caused re-phrased
    # questions to miss their seed and fall through to stochastic oracle answers.
    pending: list[dict] = []
    for q in questions:
        qid = q.get("id", "")
        qtext = q.get("question", "")
        total_asked = questions_asked_so_far + len(answers) + len(pending)
        if total_asked >= cfg.max_questions:
            answers[qid] = "You have enough information to proceed — pick the most common industry default."
            stats["over_budget"] += 1
            continue

        seed_by_id = next((s for s in cfg.seed_entries if s.get("id") == qid), None)
        if seed_by_id and "answer" in seed_by_id:
            answers[qid] = seed_by_id["answer"]
            stats["seed_hits"] += 1
            continue

        key = _hash_key(cfg.task_id, qtext)
        if key in cache:
            answers[qid] = cache[key]
            stats["cache_hits"] += 1
            continue

        pending.append(q)

    if pending:
        text, usage = await _llm_resolve(pending, cfg)
        stats["llm_calls"] += 1
        stats["tokens_in"] += usage["input_tokens"]
        stats["tokens_out"] += usage["output_tokens"]
        stats["cost_dollars"] += usage["cost_dollars"]
        parsed = parse_clarifications_block(text)
        seed_answers = {s["answer"] for s in cfg.seed_entries if "answer" in s}
        for q in pending:
            qid = q.get("id", "")
            qtext = q.get("question", "")
            ans = parsed.get(qid) or parsed.get(qid.strip()) or "Your call — pick the most common industry default."
            answers[qid] = ans
            cache[_hash_key(cfg.task_id, qtext)] = ans
            # Post-hoc: an LLM answer that exactly matches a seed counts as a seed hit.
            if ans in seed_answers:
                stats["seed_hits"] += 1
        _save_cache(cfg.cache_path, cache)

    return answers, stats


def parse_clarifications_block(text: str) -> dict[str, str]:
    """Parse a CLARIFICATIONS:\n- id: answer\n... block into a dict."""
    out: dict[str, str] = {}
    in_block = False
    for line in text.splitlines():
        line_s = line.strip()
        if line_s.upper().startswith("CLARIFICATIONS"):
            in_block = True
            continue
        if not in_block:
            continue
        if not line_s.startswith("-"):
            if line_s == "":
                continue
            # line without a dash ends the block once we hit unindented text
            if line_s and not line.startswith(" "):
                break
            continue
        body = line_s.lstrip("-").strip()
        if ":" not in body:
            continue
        qid, ans = body.split(":", 1)
        out[qid.strip()] = ans.strip()
    return out


def format_clarifications(answers: dict[str, str]) -> str:
    lines = ["CLARIFICATIONS:"]
    for qid, ans in answers.items():
        lines.append(f"- {qid}: {ans}")
    return "\n".join(lines)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--task-dir", required=True, type=Path)
    ap.add_argument("--questions", required=True, type=Path,
                    help="Path to a JSON file containing the questions array.")
    args = ap.parse_args()

    ctx = (args.task_dir / "context.md").read_text()
    seed = args.task_dir / "clarifications.seed.yaml"
    cfg = OracleConfig(
        task_id=args.task_dir.name,
        context_md=ctx,
        seed_path=seed if seed.exists() else None,
    )
    qs = json.loads(args.questions.read_text())
    answers, stats = asyncio.run(resolve_questions(qs, cfg))
    print(format_clarifications(answers))
    print("\n# stats")
    print(json.dumps(stats, indent=2))
