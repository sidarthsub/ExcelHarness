"""Local dashboard for the autoresearch loop.

Run:
    python -m benchmarks.experiments.dashboard --port 8765
    open http://127.0.0.1:8765

Serves:
    /               → the HTML dashboard
    /api/state      → live JSON: status, history, recent runs, baseline
    /api/proposal   → ?iter=N returns the iter_N.md proposal text (markdown)

The dashboard polls /api/state every 3s. No build step, no deps beyond
the stdlib — Chart.js is loaded from a CDN at runtime.
"""
from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from benchmarks.experiments import store as store_mod


ROOT = Path(__file__).resolve().parents[2]
BENCH_ROOT = ROOT / "benchmarks"
EXP_DIR = BENCH_ROOT / "experiments"
PROPOSALS_DIR = EXP_DIR / "proposals"
HISTORY_PATH = EXP_DIR / "history.jsonl"
STATUS_PATH = EXP_DIR / "status.json"
LAST_EVAL_PATH = PROPOSALS_DIR / "last_eval.json"


# ---- data aggregation ------------------------------------------------------


def _read_history() -> list[dict]:
    if not HISTORY_PATH.exists():
        return []
    out = []
    for line in HISTORY_PATH.read_text().splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            out.append(json.loads(line))
        except Exception:
            pass
    return out


def _read_status() -> dict:
    if not STATUS_PATH.exists():
        return {"phase": "not_running"}
    try:
        return json.loads(STATUS_PATH.read_text())
    except Exception:
        return {"phase": "not_running"}


def _read_last_eval() -> dict | None:
    if not LAST_EVAL_PATH.exists():
        return None
    try:
        return json.loads(LAST_EVAL_PATH.read_text())
    except Exception:
        return None


def _recent_runs(conn: sqlite3.Connection, limit: int = 60) -> list[dict]:
    cur = conn.execute(
        """SELECT task_id, label, seed, accuracy, loss, cost_cold_usd, cost_actual_usd,
                  wall_seconds, completed, terminated_reason, ts, tier
             FROM runs ORDER BY id DESC LIMIT ?""",
        (limit,),
    )
    return [dict(r) for r in cur.fetchall()]


def _current_iter_runs(conn: sqlite3.Connection, iter_n: int | None) -> list[dict]:
    if iter_n is None:
        return []
    cur = conn.execute(
        """SELECT task_id, label, seed, accuracy, loss, cost_cold_usd, cost_actual_usd,
                  wall_seconds, completed, terminated_reason, ts, tier
             FROM runs
            WHERE label = ? OR label = ?
            ORDER BY id DESC""",
        (f"iter{iter_n}_canary", f"iter{iter_n}_visible"),
    )
    return [dict(r) for r in cur.fetchall()]


def _label_aggregates(conn: sqlite3.Connection) -> list[dict]:
    cur = conn.execute(
        """SELECT label, COUNT(*) n, AVG(loss) loss, AVG(accuracy) acc,
                  AVG(cost_cold_usd) cost_cold, AVG(wall_seconds) wall,
                  SUM(CASE WHEN completed=1 THEN 1 ELSE 0 END)*1.0/COUNT(*) completion
             FROM runs
            WHERE label IS NOT NULL
            GROUP BY label
            ORDER BY MAX(id) DESC
            LIMIT 40""",
    )
    return [dict(r) for r in cur.fetchall()]


def build_state() -> dict:
    status = _read_status()
    history = _read_history()
    baseline = _read_last_eval()
    conn = store_mod.connect()
    try:
        recent = _recent_runs(conn, limit=80)
        aggs = _label_aggregates(conn)
        current_runs = _current_iter_runs(conn, status.get("iter"))
    finally:
        conn.close()
    return {
        "status": status,
        "history": history,
        "baseline": baseline,
        "recent_runs": recent,
        "label_aggregates": aggs,
        "current_iter_runs": current_runs,
    }


def read_proposal(iter_n: int) -> str:
    p = PROPOSALS_DIR / f"iter_{iter_n}.md"
    if not p.exists():
        return ""
    return p.read_text()


# ---- HTTP handler ----------------------------------------------------------


HTML = """<!doctype html>
<html lang=en>
<meta charset=utf-8>
<title>autoresearch dashboard</title>
<style>
  :root {
    --bg: #0f1116;
    --panel: #171a21;
    --panel-2: #1f2430;
    --ink: #e6e8ee;
    --ink-dim: #9aa3b2;
    --ok: #5dd39e;
    --bad: #ff6b6b;
    --accent: #7ab7ff;
    --warn: #ffc857;
    --border: #262b36;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 24px;
    background: var(--bg); color: var(--ink);
    font: 13px/1.45 -apple-system, BlinkMacSystemFont, "SF Pro", system-ui, sans-serif;
  }
  h1, h2, h3 { margin: 0 0 8px; font-weight: 600; letter-spacing: -0.01em; }
  h1 { font-size: 18px; }
  h2 { font-size: 14px; color: var(--ink-dim); text-transform: uppercase; letter-spacing: .08em; }
  h3 { font-size: 13px; }
  .grid {
    display: grid; gap: 16px;
    grid-template-columns: 1.2fr 1fr;
  }
  .panel {
    background: var(--panel);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 16px;
  }
  .kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 16px; }
  .kpi {
    background: var(--panel); border: 1px solid var(--border);
    border-radius: 8px; padding: 12px 14px;
  }
  .kpi .l { font-size: 11px; color: var(--ink-dim); text-transform: uppercase; letter-spacing: .08em; }
  .kpi .v { font-size: 22px; font-weight: 600; margin-top: 4px; }
  .kpi .d { font-size: 11px; color: var(--ink-dim); margin-top: 2px; }
  .phase-pill {
    display: inline-block; padding: 2px 10px; border-radius: 999px;
    font-size: 11px; font-weight: 600; letter-spacing: .04em;
    text-transform: uppercase;
    background: var(--panel-2); color: var(--ink-dim);
    border: 1px solid var(--border);
  }
  .phase-pill.running { background: #1a2a1f; color: var(--ok); border-color: #2a4a38; }
  .phase-pill.evaluating { background: #2a231a; color: var(--warn); border-color: #4a3f2a; }
  .phase-pill.researching { background: #1a2236; color: var(--accent); border-color: #2a3a5a; }
  .phase-pill.done { background: #2a1a2a; color: #d48cff; border-color: #4a2a4a; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  th, td { text-align: left; padding: 6px 8px; border-bottom: 1px solid var(--border); }
  th { color: var(--ink-dim); font-weight: 500; font-size: 11px; text-transform: uppercase; letter-spacing: .05em; }
  td.num { text-align: right; font-variant-numeric: tabular-nums; }
  td.task { font-family: ui-monospace, "SF Mono", Menlo, monospace; color: var(--ink-dim); }
  .accept { color: var(--ok); }
  .reject { color: var(--bad); }
  .fail   { color: var(--bad); }
  .ok     { color: var(--ok); }
  pre.proposal {
    background: var(--panel-2); border: 1px solid var(--border);
    border-radius: 6px; padding: 12px; white-space: pre-wrap; font-size: 12px;
    max-height: 400px; overflow: auto; color: var(--ink-dim);
  }
  canvas { background: var(--panel-2); border-radius: 6px; padding: 8px; }
  .row { display: flex; justify-content: space-between; align-items: center; gap: 12px; margin-bottom: 12px; }
  .muted { color: var(--ink-dim); }
  .live-dot {
    display: inline-block; width: 8px; height: 8px; border-radius: 50%;
    background: var(--ok); margin-right: 6px;
    animation: pulse 1.8s ease-in-out infinite;
  }
  @keyframes pulse { 0%,100% { opacity: .4 } 50% { opacity: 1 } }
  .bars { display: flex; flex-direction: column; gap: 6px; }
  .bar {
    display: grid; grid-template-columns: 160px 1fr 60px; gap: 8px; align-items: center;
    font-size: 12px;
  }
  .bar .track { height: 8px; background: var(--panel-2); border-radius: 4px; overflow: hidden; }
  .bar .fill { height: 100%; background: var(--accent); }
  .bar .fill.bad { background: var(--bad); }
  .bar .fill.ok { background: var(--ok); }
  .bar .task { color: var(--ink-dim); font-family: ui-monospace, "SF Mono", Menlo, monospace; }
  .bar .v { text-align: right; font-variant-numeric: tabular-nums; }
  /* Hypothesis-forward presentation */
  .hypothesis-banner {
    background: linear-gradient(135deg, var(--panel) 0%, var(--panel-2) 100%);
    border: 1px solid var(--border);
    border-left: 3px solid var(--accent);
    border-radius: 8px;
    padding: 14px 18px;
    margin-bottom: 16px;
  }
  .hypothesis-banner.accept { border-left-color: var(--ok); }
  .hypothesis-banner.reject { border-left-color: var(--bad); }
  .hypothesis-banner.researching { border-left-color: var(--accent); }
  .hypothesis-banner .label {
    font-size: 11px; color: var(--ink-dim); text-transform: uppercase;
    letter-spacing: .08em; font-weight: 600;
  }
  .hypothesis-banner .hypothesis {
    font-size: 16px; font-weight: 500; color: var(--ink); margin: 6px 0 4px;
    line-height: 1.4;
  }
  .hypothesis-banner .meta {
    font-size: 12px; color: var(--ink-dim); margin-top: 4px;
  }
  td.hyp {
    max-width: 420px; font-size: 12px; color: var(--ink);
    line-height: 1.35;
  }
  td.hyp .truncated {
    overflow: hidden; text-overflow: ellipsis; display: -webkit-box;
    -webkit-line-clamp: 2; -webkit-box-orient: vertical;
  }
  details.proposal-details summary {
    cursor: pointer; color: var(--ink-dim); font-size: 12px;
    margin-top: 8px; user-select: none;
  }
  details.proposal-details summary:hover { color: var(--ink); }
  details[open] summary { color: var(--ink); }
</style>

<h1>autoresearch <span id=live></span></h1>
<div class=row>
  <div>
    <span class="phase-pill" id=phase>idle</span>
    <span class=muted id=phase-meta></span>
  </div>
  <div class=muted id=updated></div>
</div>

<div class=kpis>
  <div class=kpi><div class=l>corpus loss</div><div class=v id=kpi-loss>—</div><div class=d id=kpi-loss-d>baseline</div></div>
  <div class=kpi><div class=l>iteration</div><div class=v id=kpi-iter>—</div><div class=d id=kpi-iter-d>of N</div></div>
  <div class=kpi><div class=l>spent</div><div class=v id=kpi-spent>—</div><div class=d id=kpi-spent-d>of budget</div></div>
  <div class=kpi><div class=l>accuracy</div><div class=v id=kpi-acc>—</div><div class=d>mean across visible</div></div>
</div>

<div id=hyp-banner class=hypothesis-banner>
  <div class=label id=hyp-label>current hypothesis</div>
  <div class=hypothesis id=hyp-text>(no iteration in flight)</div>
  <div class=meta id=hyp-meta></div>
  <details class=proposal-details>
    <summary>show full proposal</summary>
    <pre class=proposal id=proposal>(no proposal yet)</pre>
  </details>
</div>

<div class=grid>
  <div class=panel>
    <h2>loss curve</h2>
    <canvas id=loss-chart height=200></canvas>
  </div>
  <div class=panel>
    <h2>per-task (baseline)</h2>
    <div class=bars id=task-bars></div>
  </div>
</div>

<div style="height:16px"></div>

<div class=panel>
  <h2>iteration history</h2>
  <table id=hist-table>
    <thead><tr>
      <th>iter</th><th>decision</th><th>hypothesis</th>
      <th class=num>Δ</th><th class=num>loss</th>
      <th class=num>wall</th><th class=num>spent</th>
    </tr></thead>
    <tbody></tbody>
  </table>
</div>

<div style="height:16px"></div>

<div class=panel>
  <h2>recent runs</h2>
  <table id=runs-table>
    <thead><tr>
      <th>label</th><th class=task>task</th><th class=num>seed</th>
      <th class=num>acc</th><th class=num>loss</th>
      <th class=num>cost cold</th><th class=num>wall</th>
      <th>term</th><th class=num>ts</th>
    </tr></thead>
    <tbody></tbody>
  </table>
</div>

<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script>
let chart = null;
const $ = id => document.getElementById(id);
const fmt = (n, d=3) => (n==null || Number.isNaN(n)) ? "—" : (+n).toFixed(d);
const fmtS = n => n==null ? "—" : (+n).toFixed(1) + "s";
const fmtD = n => n==null ? "—" : "$" + (+n).toFixed(3);

async function refresh() {
  let s;
  try {
    s = await (await fetch("/api/state")).json();
  } catch (e) {
    $("phase").textContent = "disconnected";
    return;
  }
  const status = s.status || {};
  const history = s.history || [];
  const baseline = s.baseline || {};
  const runs = s.recent_runs || [];

  // Phase pill
  const phase = status.phase || "idle";
  const p = $("phase");
  p.textContent = phase;
  p.className = "phase-pill " + phase.replace(/[^a-z]/g,'');
  $("updated").textContent = status.updated_at ? "updated " + status.updated_at : "";

  let meta = "";
  if (status.iter != null) meta += "iter " + status.iter;
  if (status.branch) meta += " · " + status.branch;
  $("phase-meta").textContent = meta;
  $("live").innerHTML = (phase === "researching" || phase === "evaluating" || phase === "baseline")
    ? '<span class="live-dot"></span>' : '';

  // Hypothesis banner FIRST — if anything below this throws (chart lib,
  // table render, etc.) the banner is still updated. The hypothesis is
  // the single most-watched element and shouldn't depend on the rest.
  try {
    const banner = $("hyp-banner");
    const curIter = status.iter;
    let bannerKind = "researching";
    let bannerLabel = "no iteration yet";
    let bannerHyp = "(waiting for first iter)";
    let bannerMeta = "";
    let bannerProposalIter = null;

    if (phase === "researching" || phase === "evaluating" || phase === "holdout_gate" || phase === "baseline" || phase === "holdout_baseline") {
      bannerKind = "researching";
      bannerLabel = `iter ${curIter ?? 0} · ${phase}`;
      bannerHyp = "(hypothesis in progress…)";
      if (curIter && curIter >= 1) bannerProposalIter = curIter;
    } else if (history.length) {
      const last = history[history.length - 1];
      bannerKind = last.accepted ? "accept" : "reject";
      bannerLabel = `iter ${last.iter} · ${last.accepted ? 'ACCEPTED' : 'rejected'}`;
      bannerHyp = last.hypothesis_summary || "(unrecorded)";
      const d = last.paired_delta;
      bannerMeta = (d != null ? `paired Δ=${d>=0?'+':''}${d.toFixed(4)}` : "")
                 + (last.reason ? ` · ${last.reason}` : "");
      bannerProposalIter = last.iter;
    }

    banner.className = "hypothesis-banner " + bannerKind;
    $("hyp-label").textContent = bannerLabel;
    $("hyp-text").textContent = bannerHyp;
    $("hyp-meta").textContent = bannerMeta;

    if (bannerProposalIter != null) {
      try {
        const txt = await (await fetch("/api/proposal?iter=" + bannerProposalIter)).text();
        $("proposal").textContent = txt || "(proposal file not written yet)";
        if (bannerHyp.startsWith("(hypothesis") && txt) {
          const firstLine = txt.split("\n").find(l => l.trim());
          if (firstLine) {
            $("hyp-text").textContent = firstLine.replace(/^#+\s*/, "").slice(0, 240);
          }
        }
      } catch (e) { /* proposal endpoint may not have content yet */ }
    } else {
      $("proposal").textContent = "(no proposal yet)";
    }
  } catch (e) {
    console.error("hypothesis banner update failed:", e);
  }

  // KPIs
  const lastAcceptedLoss = (history.filter(h => h.accepted).slice(-1)[0] || {}).new_loss;
  const currentLoss = (baseline.corpus_loss != null) ? baseline.corpus_loss : lastAcceptedLoss;
  $("kpi-loss").textContent = fmt(currentLoss, 4);
  const prev = history.length >= 2 ? history[history.length-2].baseline_loss : null;
  $("kpi-loss-d").textContent = prev != null && currentLoss != null
    ? "Δ " + ((currentLoss - prev) >= 0 ? "+" : "") + (currentLoss - prev).toFixed(4)
    : "baseline";

  $("kpi-iter").textContent = status.iter != null ? status.iter : "—";
  $("kpi-iter-d").textContent = status.max_iters ? "of " + status.max_iters : "";

  $("kpi-spent").textContent = status.spent_usd != null ? fmtD(status.spent_usd) : "—";
  $("kpi-spent-d").textContent = status.max_dollars != null ? "of $" + (+status.max_dollars).toFixed(0) : "";

  $("kpi-acc").textContent = (baseline.mean_accuracy != null)
    ? (baseline.mean_accuracy * 100).toFixed(0) + "%"
    : "—";

  // Loss chart
  const labels = ["baseline"].concat(history.map(h => "iter " + h.iter));
  const losses = [baseline.corpus_loss].concat(history.map(h => h.baseline_loss));
  const attempted = [null].concat(history.map(h => h.new_loss));
  drawChart(labels, losses, attempted);

  // Per-task bars (from baseline.per_task if present)
  const taskBars = $("task-bars");
  taskBars.innerHTML = "";
  const pt = baseline.per_task || {};
  const tasks = Object.keys(pt);
  if (tasks.length === 0) {
    taskBars.innerHTML = '<div class=muted>(no baseline eval yet)</div>';
  } else {
    // Find max loss for scaling
    const maxLoss = Math.max(0.01, ...tasks.map(t => (pt[t].loss ?? pt[t].mean_accuracy) || 0));
    for (const t of tasks) {
      const row = pt[t];
      const loss = row.loss ?? (1 - (row.mean_accuracy || 0));
      const acc = row.mean_accuracy;
      const pct = Math.min(100, (loss / maxLoss) * 100);
      const cls = loss < 0.1 ? "ok" : loss > 0.5 ? "bad" : "";
      const el = document.createElement("div");
      el.className = "bar";
      el.innerHTML = `
        <div class=task>${t}</div>
        <div class=track><div class="fill ${cls}" style="width:${pct}%"></div></div>
        <div class=v>${fmt(loss,3)}</div>`;
      taskBars.appendChild(el);
    }
  }

  // History table — hypothesis is the headline column.
  const hbody = document.querySelector("#hist-table tbody");
  hbody.innerHTML = "";
  for (const h of history.slice().reverse()) {
    const tr = document.createElement("tr");
    const delta = h.paired_delta != null ? h.paired_delta
                : ((h.new_loss != null && h.baseline_loss != null)
                    ? (h.new_loss - h.baseline_loss) : null);
    const hyp = h.hypothesis_summary || "(unrecorded)";
    const reason = (h.reason || "").replace(/"/g, '&quot;');
    tr.innerHTML = `
      <td>${h.iter}</td>
      <td class="${h.accepted ? 'accept' : 'reject'}">${h.accepted ? 'ACCEPT' : 'REJECT'}</td>
      <td class=hyp title="${reason}"><div class=truncated>${hyp.replace(/</g,'&lt;')}</div></td>
      <td class=num>${delta == null ? '—' : (delta >= 0 ? '+' : '') + delta.toFixed(4)}</td>
      <td class=num>${fmt(h.new_loss, 4)}</td>
      <td class=num>${fmtS(h.wall_seconds)}</td>
      <td class=num>${fmtD(h.spent_usd)}</td>`;
    hbody.appendChild(tr);
  }

  // (Hypothesis banner is updated above, before any potentially-failing
  // chart/table renders. Keeping it there ensures the most-watched UI
  // element doesn't get skipped on JS errors elsewhere.)

  // Recent runs
  const rbody = document.querySelector("#runs-table tbody");
  rbody.innerHTML = "";
  for (const r of runs) {
    const tr = document.createElement("tr");
    const accClass = (r.accuracy != null && r.accuracy >= 0.95) ? 'ok'
                   : (r.accuracy != null && r.accuracy < 0.5) ? 'bad' : '';
    const termClass = r.completed ? 'ok' : 'bad';
    tr.innerHTML = `
      <td>${r.label || ''}</td>
      <td class=task>${r.task_id}</td>
      <td class=num>${r.seed ?? ''}</td>
      <td class="num ${accClass}">${fmt(r.accuracy, 3)}</td>
      <td class=num>${fmt(r.loss, 3)}</td>
      <td class=num>${fmtD(r.cost_cold_usd)}</td>
      <td class=num>${fmtS(r.wall_seconds)}</td>
      <td class="${termClass}">${r.terminated_reason || ''}</td>
      <td class="num muted">${(r.ts || '').slice(11,19)}</td>`;
    rbody.appendChild(tr);
  }
}

function drawChart(labels, losses, attempted) {
  const ctx = $("loss-chart").getContext("2d");
  const data = {
    labels,
    datasets: [
      {
        label: 'baseline loss',
        data: losses,
        borderColor: '#7ab7ff',
        backgroundColor: 'rgba(122,183,255,0.12)',
        tension: 0.25,
        fill: true,
        pointRadius: 3,
      },
      {
        label: 'attempted loss',
        data: attempted,
        borderColor: '#ffc857',
        borderDash: [4,4],
        backgroundColor: 'transparent',
        tension: 0,
        pointRadius: 4,
        pointStyle: 'crossRot',
      }
    ]
  };
  const opts = {
    responsive: true,
    animation: false,
    plugins: {
      legend: { labels: { color: '#9aa3b2', font: { size: 11 } } },
    },
    scales: {
      x: { ticks: { color: '#9aa3b2', font: { size: 10 } }, grid: { color: '#262b36' } },
      y: { ticks: { color: '#9aa3b2', font: { size: 10 } }, grid: { color: '#262b36' }, beginAtZero: true },
    }
  };
  if (chart) { chart.data = data; chart.update(); }
  else chart = new Chart(ctx, { type: 'line', data, options: opts });
}

refresh();
setInterval(refresh, 3000);
</script>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    # Silence access log spam.
    def log_message(self, fmt: str, *args) -> None:  # noqa: N802
        return

    def _write(self, status: int, body: bytes, ctype: str = "text/plain; charset=utf-8") -> None:
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:  # noqa: N802
        u = urlparse(self.path)
        if u.path in ("/", "/index.html"):
            self._write(200, HTML.encode("utf-8"), "text/html; charset=utf-8")
            return
        if u.path == "/api/state":
            try:
                state = build_state()
                body = json.dumps(state, default=str).encode("utf-8")
            except Exception as e:
                body = json.dumps({"error": f"{type(e).__name__}: {e}"}).encode("utf-8")
                self._write(500, body, "application/json")
                return
            self._write(200, body, "application/json")
            return
        if u.path == "/api/proposal":
            q = parse_qs(u.query)
            try:
                iter_n = int(q.get("iter", ["-1"])[0])
            except ValueError:
                iter_n = -1
            txt = read_proposal(iter_n)
            self._write(200, txt.encode("utf-8"), "text/plain; charset=utf-8")
            return
        self._write(404, b"not found")


# ---- CLI -------------------------------------------------------------------


def _main() -> int:
    ap = argparse.ArgumentParser(prog="benchmarks.experiments.dashboard")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--host", default="127.0.0.1")
    args = ap.parse_args()
    srv = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"dashboard: http://{args.host}:{args.port}", file=sys.stderr)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nshutting down", file=sys.stderr)
    srv.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(_main())
