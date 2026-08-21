"""
The public website: the same listings, for anyone with the link.

Deliberately different from app.py, which is yours alone:
  * no scraping controls - a stranger cannot start a 90-minute crawl
  * reads one small committed data file, not your raw scrape output
  * each visitor's shortlist lives in their own browser, so nobody sees
    anyone else's, and nothing they tick can change your data
  * the Excel download is built on the server, so it keeps the tick-box
    columns and the colours

Run locally:  .venv/bin/python public_site.py
Deployed:     gunicorn public_site:app
"""

import io
import json
import os

from flask import Flask, jsonify, render_template_string, request, send_file
from openpyxl import Workbook

from excel_style import dress

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "public_data.json")

app = Flask(__name__)


def payload():
    with open(DATA, encoding="utf-8") as fh:
        return json.load(fh)


@app.route("/api/rows")
def api_rows():
    return jsonify(payload())


@app.route("/api/excel", methods=["POST"])
def api_excel():
    rows = request.get_json(force=True) or []
    cols = ["Applied", "In touch", "Company", "City", "Country", "Role",
            "Kind", "Deadline", "Posted", "Checked", "Link"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Opera auditions"
    ws.append(cols)
    for r in rows:
        ws.append([r.get("Applied", "No"), r.get("In touch", "No")]
                  + [r.get(c, "") for c in cols[2:]])
    dress(ws, cols, [10, 10, 28, 15, 12, 54, 22, 14, 13, 12, 46],
          tick_cols=("Applied", "In touch"))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True, download_name="opera-auditions.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument."
                 "spreadsheetml.sheet")


@app.route("/")
def index():
    with open(os.path.join(HERE, "public_page.html"), encoding="utf-8") as fh:
        return render_template_string(fh.read())


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5056))
    print(f"\n  Public site preview: http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port)
