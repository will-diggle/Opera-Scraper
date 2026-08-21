"""
Build a single shareable web page from the latest results.

The app on your Mac only works on your Mac. This makes one self-contained HTML
file - search, filters and CSV export all baked in - that you can publish and
send to anyone as a link. It is a snapshot: rerun this after each scrape.

Run:  .venv/bin/python make_share_page.py
Out:  share/opera-jobs.html
"""

import csv
import json
import os
import re
from datetime import date

from freshness import assess
from stage5_prioritise import categorise, NOT_FOR_YOU

HERE = os.path.dirname(os.path.abspath(__file__))
JOB_FILES = ["jobs_targets.csv", "opera_jobs.csv"]
OUT_DIR = os.path.join(HERE, "share")
OUT = os.path.join(OUT_DIR, "opera-jobs.html")

# Only these go in the shared page - the rest is noise for a reader.
KEEP_PRIORITIES = ("A - Opera house / theatre", "A - Festival",
                   "B - Professional choir", "B - Orchestra / concert",
                   "B - Other")

# The two results files spell countries differently (DE vs Germany).
COUNTRY = {
    "AT": "Austria", "BE": "Belgium", "BG": "Bulgaria", "CH": "Switzerland",
    "CY": "Cyprus", "CZ": "Czechia", "DE": "Germany", "DK": "Denmark",
    "EE": "Estonia", "ES": "Spain", "FI": "Finland", "FR": "France",
    "GB": "UK", "GR": "Greece", "HR": "Croatia", "HU": "Hungary",
    "IE": "Ireland", "IS": "Iceland", "IT": "Italy", "LT": "Lithuania",
    "LU": "Luxembourg", "LV": "Latvia", "MC": "Monaco", "MT": "Malta",
    "NL": "Netherlands", "NO": "Norway", "PL": "Poland", "PT": "Portugal",
    "RO": "Romania", "RS": "Serbia", "SE": "Sweden", "SI": "Slovenia",
    "SK": "Slovakia", "UK": "UK",
}

# Photo credits and "read more" tails get swept up with the headline text.
JUNK = re.compile(
    r"^.{0,60}?\bon\s+Unsplash\b\s*"                       # photo credits
    r"|^\s*(?:©|\(c\))\s*[^|]{0,50}?\s(?=[A-ZÄÖÜ])"        # © Name Role…
    r"|\s*(?:Read more|Lire la suite|Weiterlesen|Leggi tutto|Mehr)\s*$",
    re.I)
INSTRUMENT_ONLY = re.compile(
    r"double bass|contrabass|kontrabass|violin|cello|viola|flute|oboe|"
    r"clarinet|bassoon|trumpet|trombone|percussion|harp|timpani", re.I)


def tidy(text):
    prev = None
    while prev != text:
        prev = text
        text = JUNK.sub("", text).strip(" ·-–—|")
    return text


def data_date(rows_dates):
    """When the listings were actually gathered.

    NOT the file's timestamp: the scraper resumes from a cache, so a file
    rewritten today can be almost entirely made of results fetched weeks ago.
    Each row records its own check date, so use those.
    """
    stamps = sorted({d for d in rows_dates if d})
    if not stamps:
        return date.today().strftime("%d %B %Y")
    first = date.fromisoformat(stamps[0])
    last = date.fromisoformat(stamps[-1])
    if first == last:
        return last.strftime("%d %B %Y")
    if first.month == last.month and first.year == last.year:
        return f"{first.day}\u2013{last.strftime('%d %B %Y')}"
    return f"{first.strftime('%d %b')}\u2013{last.strftime('%d %b %Y')}"


def load():
    meta = {}
    tpath = os.path.join(HERE, "targets.csv")
    if os.path.exists(tpath):
        with open(tpath, newline="", encoding="utf-8") as fh:
            meta = {r["name"]: r for r in csv.DictReader(fh)}

    rows, seen = [], set()
    for name in JOB_FILES:
        path = os.path.join(HERE, name)
        if not os.path.exists(path):
            continue
        with open(path, newline="", encoding="utf-8") as fh:
            for r in csv.DictReader(fh):
                key = (r.get("Company", ""), r.get("Role / posting", ""),
                       r.get("Link", ""))
                if key in seen:
                    continue
                seen.add(key)
                m = meta.get(r.get("Company", ""), {})
                posted, fresh = assess(r)
                prio = categorise(r.get("Company", ""), m.get("org_type", ""))
                typ = r.get("Type", "")
                if NOT_FOR_YOU.search(r.get("Role / posting", "")):
                    typ = "Not singing"
                if prio not in KEEP_PRIORITIES or fresh == "Out of date":
                    continue
                if typ != "Singer":
                    continue
                role = tidy(r.get("Role / posting", ""))
                if not role or INSTRUMENT_ONLY.search(role):
                    continue
                rows.append([
                    prio.split(" - ")[1] if " - " in prio else prio,
                    r.get("Company", ""), m.get("city", ""),
                    COUNTRY.get(r.get("Country", "").upper(),
                                r.get("Country", "")), role,
                    r.get("Deadline found", ""), posted.split(" (")[0], fresh,
                    r.get("Link", ""), r.get("Date checked", ""),
                ])
    rows.sort(key=lambda r: (r[7] != "Upcoming", r[3], r[1]))
    return rows


TEMPLATE = """<title>Opera Audition Watch</title>
<style>
:root{
  --ground:#f6f5f2; --panel:#ffffff; --ink:#191b1f; --ink-soft:#5d6472;
  --line:#dcdde2; --accent:#233a63; --accent-soft:#e8edf6;
  --live:#1d6b4f; --live-bg:#dcefe4; --quiet:#8a8f99;
  --serif:"Iowan Old Style","Palatino Linotype",Palatino,"Book Antiqua",
          Georgia,serif;
  --sans:var(--serif);
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --ground:#14161a; --panel:#1c1f25; --ink:#eceef2; --ink-soft:#a2a9b6;
    --line:#2e323b; --accent:#a8c0e8; --accent-soft:#232a38;
    --live:#7fd3ab; --live-bg:#1c3329; --quiet:#767d8a;
  }
}
:root[data-theme="dark"]{
  --ground:#14161a; --panel:#1c1f25; --ink:#eceef2; --ink-soft:#a2a9b6;
  --line:#2e323b; --accent:#a8c0e8; --accent-soft:#232a38;
  --live:#7fd3ab; --live-bg:#1c3329; --quiet:#767d8a;
}
*{box-sizing:border-box}
body{margin:0;background:var(--ground);color:var(--ink);
  font:16px/1.55 var(--sans)}
.masthead{padding:44px 28px 30px;max-width:1560px;margin:0 auto}
.eyebrow{font:600 11px/1 var(--serif);letter-spacing:.13em;
  text-transform:uppercase;color:var(--ink-soft)}
h1{font:400 clamp(31px,4.7vw,47px)/1.08 var(--serif);margin:14px 0 0;
  text-wrap:balance;letter-spacing:-.008em}
.sub{color:var(--ink-soft);margin:12px 0 0;max-width:60ch;
  font:16.5px/1.5 var(--serif)}
.stats{display:flex;flex-wrap:wrap;gap:26px;margin-top:26px;
  padding-top:22px;border-top:1px solid var(--line)}
.stat b{display:block;font:400 28px/1 var(--serif);
  font-variant-numeric:tabular-nums}
.stat span{font-size:12px;color:var(--ink-soft);letter-spacing:.04em;
  text-transform:uppercase}
.bar{position:sticky;top:0;z-index:5;background:var(--ground);
  border-bottom:1px solid var(--line);padding:12px 28px}
.bar .in{max-width:1560px;margin:0 auto;display:flex;flex-wrap:wrap;
  gap:9px;align-items:center}
input[type=search],select{font:inherit;font-size:14px;padding:8px 11px;
  border:1px solid var(--line);border-radius:7px;background:var(--panel);
  color:var(--ink)}
input[type=search]{min-width:250px;flex:1 1 250px}
button{font:inherit;font-size:14px;padding:8px 13px;border-radius:7px;
  border:1px solid var(--accent);background:var(--accent);
  color:var(--ground);cursor:pointer}
button.ghost{background:transparent;color:var(--accent)}
button:focus-visible,input:focus-visible,select:focus-visible{
  outline:2px solid var(--accent);outline-offset:2px}
.count{color:var(--ink-soft);font-size:13px;margin-left:auto;
  font-variant-numeric:tabular-nums}
main{max-width:1560px;margin:0 auto;padding:22px 28px 70px}
.tablewrap{overflow-x:auto;background:var(--panel);border:1px solid var(--line);
  border-radius:10px}
table{border-collapse:collapse;width:100%;font-size:14.5px;
  font-variant-numeric:tabular-nums}
th{text-align:left;padding:11px 13px;border-bottom:1px solid var(--line);
  font:600 11.5px/1 var(--serif);letter-spacing:.08em;
  text-transform:uppercase;color:var(--ink-soft);white-space:nowrap;cursor:pointer}
td{padding:12px 13px;border-bottom:1px solid var(--line);vertical-align:top}
tr:last-child td{border-bottom:0}
.chip{display:inline-block;font-size:11px;padding:2.5px 8px;border-radius:99px;
  background:var(--accent-soft);color:var(--accent);white-space:nowrap}
.chip.live{background:var(--live-bg);color:var(--live)}
.co{font:600 15px/1.3 var(--serif)}
.city{color:var(--ink-soft);font-size:12.5px}
.when{font-variant-numeric:tabular-nums;white-space:nowrap}
.q{color:var(--quiet);font-size:12px}
a{color:var(--accent)}
.empty{padding:50px;text-align:center;color:var(--ink-soft)}
th.pick,td.pick{width:34px;padding-right:0;text-align:center}
td.pick input,th.pick input{width:15px;height:15px;cursor:pointer;accent-color:var(--accent)}
tr.chosen td{background:var(--accent-soft)}
.selnote{color:var(--ink-soft);font-size:13px}
.selnote button{padding:4px 9px;font-size:12.5px;margin-left:6px}
footer{max-width:1560px;margin:0 auto;padding:0 28px 60px;
  color:var(--ink-soft);font-size:13px;line-height:1.6}
footer p{max-width:70ch}
</style>

<div class="masthead">
  <div class="eyebrow">Websites last checked &middot; __DATE__</div>
  <h1>Opera Audition Watch</h1>
  <p class="sub">Singer openings gathered from opera house, festival and
     professional choir websites across Europe. Every row links to the
     company's own page &mdash; always confirm details there before applying.</p>
  <div class="stats">
    <div class="stat"><b>__N__</b><span>openings listed</span></div>
    <div class="stat"><b>__UPCOMING__</b><span>with a future date</span></div>
    <div class="stat"><b>__COUNTRIES__</b><span>countries</span></div>
    <div class="stat"><b>__COMPANIES__</b><span>companies</span></div>
  </div>
</div>

<div class="bar"><div class="in">
  <input type="search" id="q" placeholder="Search role, company or city…">
  <select id="country"></select>
  <select id="prio">
    <option value="">All kinds</option>
    <option>Opera house / theatre</option>
    <option>Festival</option>
    <option>Professional choir</option>
    <option>Orchestra / concert</option>
  </select>
  <select id="when">
    <option value="">Any date</option>
    <option value="Upcoming">Dated ahead</option>
    <option value="Recent">Posted recently</option>
  </select>
  <select id="checked"></select>
  <button class="ghost" id="csv">Download CSV</button>
  <span class="selnote"><button class="ghost" id="clearsel" hidden>Clear selection</button></span>
  <span class="count" id="count"></span>
</div></div>

<main>
  <div class="tablewrap">
    <table>
      <thead><tr>
        <th class="pick"><input type="checkbox" id="all"
            title="Select everything matching the current filters"></th>
        <th data-k="1">Company</th><th data-k="3">Country</th>
        <th data-k="4">Role</th><th data-k="0">Kind</th>
        <th data-k="5">Deadline</th><th data-k="6">Posted</th>
        <th data-k="9">Checked</th><th>Link</th>
      </tr></thead>
      <tbody id="body"></tbody>
    </table>
  </div>
  <div class="empty" id="empty" hidden>Nothing matches those filters.</div>
</main>

<footer>
  <p><strong>How this was made.</strong> Company websites were checked
  automatically for audition, chorus and young-artist pages. Listings that
  could be dated and had clearly passed were removed, but many pages carry no
  date at all, so some entries here may already be closed. This is a research
  aid, not an official listing.</p>
</footer>

<script>
const COLS = ["Kind","Company","City","Country","Role","Deadline","Posted",
              "Freshness","Link","Checked"];
const ROWS = __DATA__;
ROWS.forEach((r, i) => r[10] = i);          // stable id for selection
const PICKED = new Set();
let view = ROWS, sortK = null, sortDir = 1;

const $ = id => document.getElementById(id);
const esc = s => (s||"").replace(/[&<>"]/g,
  c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));

const checked = [...new Set(ROWS.map(r => r[9]))].filter(Boolean).sort().reverse();
$("checked").innerHTML = '<option value="">Checked any time</option>' +
  checked.map(d => `<option value="${esc(d)}">Checked ${esc(d)}</option>`).join("");

const countries = [...new Set(ROWS.map(r => r[3]))].filter(Boolean).sort();
$("country").innerHTML = '<option value="">All countries</option>' +
  countries.map(c => `<option>${esc(c)}</option>`).join("");

function draw(){
  const q = $("q").value.toLowerCase().trim();
  const c = $("country").value, p = $("prio").value, w = $("when").value;
  const ch = $("checked").value;
  view = ROWS.filter(r =>
    (!c || r[3] === c) && (!p || r[0] === p) && (!w || r[7] === w) &&
    (!ch || r[9] === ch) &&
    (!q || (r[1] + " " + r[4] + " " + r[2]).toLowerCase().includes(q)));
  if (sortK !== null)
    view = [...view].sort((a,b) =>
      ((a[sortK]||"") > (b[sortK]||"") ? 1 : -1) * sortDir);

  $("body").innerHTML = view.map(r => `<tr class="${PICKED.has(r[10])?"chosen":""}">
    <td class="pick"><input type="checkbox" data-i="${r[10]}"
        ${PICKED.has(r[10])?"checked":""}></td>
    <td><div class="co">${esc(r[1])}</div>${r[2] ?
        `<div class="city">${esc(r[2])}</div>` : ""}</td>
    <td>${esc(r[3])}</td>
    <td>${esc(r[4])}</td>
    <td><span class="chip">${esc(r[0])}</span></td>
    <td class="when">${r[5] ? esc(r[5]) : '<span class="q">&mdash;</span>'}</td>
    <td class="when">${r[6] ? `<span class="chip ${r[7]==="Upcoming"?"live":""}">
        ${esc(r[6])}</span>` : '<span class="q">no date</span>'}</td>
    <td class="when q">${esc(r[9])}</td>
    <td><a href="${esc(r[8])}" target="_blank" rel="noopener">open</a></td>
  </tr>`).join("");
  $("empty").hidden = view.length > 0;
  $("count").textContent = PICKED.size
    ? `${PICKED.size} selected \u00b7 ${view.length} of ${ROWS.length}`
    : `${view.length} of ${ROWS.length}`;
  syncAll();
  $("csv").textContent = PICKED.size
    ? `Download ${PICKED.size} selected` : "Download CSV";
  $("clearsel").hidden = PICKED.size === 0;
}

["q","country","prio","when","checked"].forEach(id =>
  $(id).addEventListener("input", draw));

// one listener for every row box, now and after every redraw
$("body").addEventListener("change", e => {
  const box = e.target.closest("input[type=checkbox][data-i]");
  if (!box) return;
  const id = +box.dataset.i;
  box.checked ? PICKED.add(id) : PICKED.delete(id);
  box.closest("tr").classList.toggle("chosen", box.checked);
  draw();
});

function syncAll() {
  const box = $("all");
  const shown = view.length;
  const on = view.filter(r => PICKED.has(r[10])).length;
  box.checked = shown > 0 && on === shown;
  box.indeterminate = on > 0 && on < shown;
}

$("all").addEventListener("change", e => {
  // ticks or clears every row the current filters match, not just this page
  view.forEach(r => e.target.checked ? PICKED.add(r[10]) : PICKED.delete(r[10]));
  draw();
});

$("clearsel").addEventListener("click", () => { PICKED.clear(); draw(); });
document.querySelectorAll("th[data-k]").forEach(th =>
  th.addEventListener("click", () => {
    const k = +th.dataset.k;
    sortDir = (k === sortK) ? -sortDir : 1; sortK = k; draw();
  }));
let saver = null;
(async () => {
  // In the shared viewer this returns the save API; in a plain browser it is
  // null and we fall back to an ordinary download link.
  saver = (window.claude && claude.use) ? await claude.use("downloads") : null;
})();

$("csv").addEventListener("click", async () => {
  const quote = v => '"' + String(v==null?"":v).replace(/"/g,'""') + '"';
  const chosen = PICKED.size ? view.filter(r => PICKED.has(r[10])) : view;
  const csv = [COLS.join(",")]
    .concat(chosen.map(r => r.slice(0, COLS.length).map(quote).join(",")))
    .join("\\n");
  const body = "\\ufeff" + csv;
  const btn = $("csv"), original = btn.textContent;
  if (saver) {
    try {
      await saver.save({filename:"opera-auditions.csv", data:body});
      btn.textContent = "Saved";
    } catch (err) {
      btn.textContent = err && err.code === "declined"
        ? "Download cancelled" : "Download unavailable";
    }
    setTimeout(() => { btn.textContent = original; }, 2600);
    return;
  }
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([body], {type:"text/csv"}));
  a.download = "opera-auditions.csv";
  document.body.appendChild(a); a.click(); a.remove();
});
draw();
</script>
"""


def main():
    rows = load()
    os.makedirs(OUT_DIR, exist_ok=True)
    html = (TEMPLATE
            .replace("__DATA__", json.dumps(rows, ensure_ascii=False))
            .replace("__DATE__", data_date([r[9] for r in rows]))
            .replace("__N__", f"{len(rows):,}")
            .replace("__UPCOMING__", str(sum(1 for r in rows if r[7] == "Upcoming")))
            .replace("__COUNTRIES__", str(len({r[3] for r in rows if r[3]})))
            .replace("__COMPANIES__", str(len({r[1] for r in rows}))))
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write(html)
    kb = os.path.getsize(OUT) / 1024
    print(f"Wrote {OUT}  ({len(rows)} openings, {kb:.0f} KB)")


if __name__ == "__main__":
    main()
