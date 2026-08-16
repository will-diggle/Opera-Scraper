"""
The Opera Jobs app.

A small website that runs on your own Mac. Start it, open the page, and you get
a search box, filters and a button that re-runs the scrapers. No spreadsheets
to hunt through.

Start it with:   .venv/bin/python app.py
Then open:       http://localhost:5055
Stop it with:    Ctrl-C in the terminal
"""

import csv
import os
import subprocess
import sys
import threading
from datetime import datetime

from flask import Flask, jsonify, render_template_string, request

HERE = os.path.dirname(os.path.abspath(__file__))
PY = os.path.join(HERE, ".venv", "bin", "python")

# The jobs files we know about, best first.
JOB_FILES = ["jobs_targets.csv", "opera_jobs.csv"]

app = Flask(__name__)

# --- running a scrape in the background -------------------------------------

state = {"running": False, "job": "", "log": [], "started": ""}
_lock = threading.Lock()

TASKS = {
    "curated": {
        "label": "Check my curated list (222 houses, ~10 min)",
        "cmd": [PY, "scraper.py"],
    },
    "full": {
        "label": "Full Operabase sweep (12,000 companies, ~90 min)",
        "cmd": [PY, "scraper.py", "--file", "targets.csv"],
    },
    "prioritise": {
        "label": "Re-grade results (instant)",
        "cmd": [PY, "stage5_prioritise.py"],
    },
    "refresh_directory": {
        "label": "Refresh company directory from Operabase (~90 min)",
        "cmd": [PY, "stage2_details.py"],
    },
}


def run_task(key):
    cmd = TASKS[key]["cmd"]
    with _lock:
        state.update(running=True, job=TASKS[key]["label"], log=[],
                     started=datetime.now().strftime("%H:%M"))
    try:
        proc = subprocess.Popen(cmd, cwd=HERE, stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT, text=True,
                                bufsize=1)
        for line in proc.stdout:
            with _lock:
                state["log"].append(line.rstrip())
                del state["log"][:-400]      # keep the last 400 lines
        proc.wait()
        with _lock:
            state["log"].append(
                "FINISHED" if proc.returncode == 0
                else f"STOPPED (exit code {proc.returncode})")
    except Exception as e:                    # noqa: BLE001 - show user anything
        with _lock:
            state["log"].append(f"ERROR: {e}")
    finally:
        with _lock:
            state["running"] = False


# --- reading the results ----------------------------------------------------

def load_jobs():
    """Read whichever results file exists, newest first."""
    rows = []
    for name in JOB_FILES:
        path = os.path.join(HERE, name)
        if not os.path.exists(path):
            continue
        with open(path, newline="", encoding="utf-8") as fh:
            for r in csv.DictReader(fh):
                r["_source"] = name
                rows.append(r)
    # add the Priority grade if stage 5 has been run
    meta = {}
    tpath = os.path.join(HERE, "targets.csv")
    if os.path.exists(tpath):
        with open(tpath, newline="", encoding="utf-8") as fh:
            meta = {r["name"]: r for r in csv.DictReader(fh)}
    for r in rows:
        m = meta.get(r.get("Company", ""), {})
        r["City"] = m.get("city", "")
    return rows


PAGE = """
<!doctype html><meta charset="utf-8"><title>Opera Jobs</title>
<style>
 :root{--bg:#f7f8fa;--ink:#1c1c1c;--line:#d3d8e0;--accent:#1b2a4e;--soft:#6b6b6b}
 *{box-sizing:border-box}
 body{margin:0;font:15px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
      background:var(--bg);color:var(--ink)}
 header{background:var(--accent);color:#fff;padding:18px 24px}
 header h1{margin:0;font-size:20px;font-weight:600}
 header p{margin:4px 0 0;opacity:.85;font-size:13px}
 .wrap{padding:20px 24px;max-width:1500px}
 .panel{background:#fff;border:1px solid var(--line);border-radius:8px;
        padding:16px;margin-bottom:18px}
 button{font:inherit;padding:9px 14px;border-radius:6px;border:1px solid var(--accent);
        background:var(--accent);color:#fff;cursor:pointer;margin:3px 4px 3px 0}
 button.ghost{background:#fff;color:var(--accent)}
 button:disabled{opacity:.45;cursor:not-allowed}
 input,select{font:inherit;padding:8px;border:1px solid var(--line);
              border-radius:6px;background:#fff}
 input[type=search]{width:320px}
 table{border-collapse:collapse;width:100%;background:#fff;font-size:13.5px}
 th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);
       vertical-align:top}
 th{background:#e8ecf3;position:sticky;top:0;cursor:pointer;white-space:nowrap}
 tr:hover td{background:#f4f7fc}
 a{color:var(--accent)}
 .tag{font-size:11px;padding:2px 7px;border-radius:99px;background:#eee;
      white-space:nowrap}
 .Singer{background:#d6e4f5}.Unclear{background:#eceff4}
 .log{background:#1c1c1c;color:#d6d6d6;font:12px/1.45 ui-monospace,Menlo,monospace;
      padding:12px;border-radius:6px;height:190px;overflow:auto;white-space:pre-wrap}
 .muted{color:var(--soft);font-size:13px}
</style>
<header>
  <h1>Opera Jobs</h1>
  <p>Auditions, chorus vacancies and young artist programmes, gathered from
     company websites.</p>
</header>
<div class="wrap">

<div class="panel">
  <strong>Re-scrape</strong>
  <div style="margin-top:8px">
    {% for key, t in tasks.items() %}
      <button onclick="startJob('{{key}}')" class="{{ '' if key=='curated' else 'ghost' }}">
        {{ t.label }}
      </button>
    {% endfor %}
  </div>
  <div id="status" class="muted" style="margin-top:10px"></div>
  <div class="log" id="log" style="display:none"></div>
</div>

<div class="panel">
  <input type="search" id="q" placeholder="Search role, company, city…"
         oninput="render()">
  <select id="type" onchange="render()">
    <option value="">Any type</option>
    <option value="Singer" selected>Singer only</option>
    <option value="Unclear">Unclear</option>
  </select>
  <select id="country" onchange="render()"></select>
  <select id="company" onchange="render()"></select>
  <label class="muted" style="margin-left:8px">
    <input type="checkbox" id="deadline" onchange="render()"> has a deadline
  </label>
  <span id="count" class="muted" style="margin-left:10px"></span>
</div>

<table id="tbl">
  <thead><tr>
    <th onclick="sortBy('Company')">Company</th>
    <th onclick="sortBy('Country')">Country</th>
    <th onclick="sortBy('Role / posting')">Role / posting</th>
    <th onclick="sortBy('Type')">Type</th>
    <th onclick="sortBy('Deadline found')">Deadline</th>
    <th>Links</th>
  </tr></thead>
  <tbody></tbody>
</table>
</div>

<script>
let ROWS = [], sortKey = "Company", sortDir = 1;

fetch("/api/jobs").then(r => r.json()).then(d => {
  ROWS = d;
  fill("country", [...new Set(d.map(r => r.Country))].sort(), "Any country");
  fill("company", [...new Set(d.map(r => r.Company))].sort(), "Any company");
  render();
});

function fill(id, vals, blank) {
  const s = document.getElementById(id);
  s.innerHTML = `<option value="">${blank}</option>` +
    vals.filter(Boolean).map(v => `<option>${esc(v)}</option>`).join("");
}
function esc(s){ return (s||"").replace(/[&<>"]/g, c =>
  ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }

function render() {
  const q = document.getElementById("q").value.toLowerCase();
  const type = document.getElementById("type").value;
  const country = document.getElementById("country").value;
  const company = document.getElementById("company").value;
  const dl = document.getElementById("deadline").checked;

  let rows = ROWS.filter(r =>
    (!type || r.Type === type) &&
    (!country || r.Country === country) &&
    (!company || r.Company === company) &&
    (!dl || r["Deadline found"]) &&
    (!q || (r.Company + " " + r["Role / posting"] + " " + (r.City||""))
             .toLowerCase().includes(q)));

  rows.sort((a,b) => ((a[sortKey]||"") > (b[sortKey]||"") ? 1 : -1) * sortDir);

  document.querySelector("#tbl tbody").innerHTML = rows.map(r => `
    <tr>
      <td>${esc(r.Company)}${r.City ? '<div class="muted">'+esc(r.City)+'</div>' : ''}</td>
      <td>${esc(r.Country)}</td>
      <td>${esc(r["Role / posting"])}</td>
      <td><span class="tag ${esc(r.Type)}">${esc(r.Type)}</span></td>
      <td>${esc(r["Deadline found"])}</td>
      <td><a href="${esc(r.Link)}" target="_blank" rel="noopener">open</a>
          ${r["Social accounts"] ? " · " + r["Social accounts"].split(" | ")
             .map(u => `<a href="${esc(u)}" target="_blank" rel="noopener">social</a>`)
             .join(" ") : ""}</td>
    </tr>`).join("");
  document.getElementById("count").textContent =
    rows.length + " of " + ROWS.length + " rows";
}
function sortBy(k){ sortDir = (k === sortKey) ? -sortDir : 1; sortKey = k; render(); }

function startJob(key) {
  fetch("/api/run/" + key, {method:"POST"}).then(r => r.json()).then(d => {
    if (d.error) alert(d.error);
    poll();
  });
}
function poll() {
  fetch("/api/status").then(r => r.json()).then(s => {
    const log = document.getElementById("log");
    document.querySelectorAll("button").forEach(b => b.disabled = s.running);
    document.getElementById("status").textContent = s.running
      ? `Running: ${s.job} (started ${s.started}) — you can leave this page open`
      : (s.job ? `Last run: ${s.job}` : "Nothing running.");
    if (s.log.length) {
      log.style.display = "block";
      log.textContent = s.log.join("\\n");
      log.scrollTop = log.scrollHeight;
    }
    if (s.running) setTimeout(poll, 1500);
    else if (s.job) setTimeout(() =>
      fetch("/api/jobs").then(r => r.json()).then(d => { ROWS = d; render(); }), 500);
  });
}
poll();
</script>
"""


@app.route("/")
def index():
    return render_template_string(PAGE, tasks=TASKS)


@app.route("/api/jobs")
def api_jobs():
    return jsonify(load_jobs())


@app.route("/api/status")
def api_status():
    with _lock:
        return jsonify(dict(state))


@app.route("/api/run/<key>", methods=["POST"])
def api_run(key):
    if key not in TASKS:
        return jsonify({"error": "unknown task"}), 400
    with _lock:
        if state["running"]:
            return jsonify({"error": "Something is already running."}), 409
    threading.Thread(target=run_task, args=(key,), daemon=True).start()
    return jsonify({"ok": True})


if __name__ == "__main__":
    print("\n  Opera Jobs app running.")
    print("  Open this in your browser:  http://localhost:5055")
    print("  Press Ctrl-C here to stop.\n")
    app.run(port=5055, debug=False)
