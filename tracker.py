"""
Your shortlist: the jobs you have actually decided to go for.

Everything else in this project is a firehose the scrapers refill. This file is
yours - what you saved, whether you have applied, whether you have emailed, and
your own notes. Nothing here is ever overwritten by a scrape.

Kept in saved_jobs.json for the app's own use; the file you actually open is
saved_jobs.xlsx, with Yes-No columns that turn green.
"""

import json
import os
from datetime import date

from openpyxl import Workbook

from excel_style import dress

HERE = os.path.dirname(os.path.abspath(__file__))
STORE = os.path.join(HERE, "saved_jobs.json")
XLSX = os.path.join(HERE, "saved_jobs.xlsx")

FIELDS = ["id", "Company", "Country", "City", "Role", "Voice type",
          "Deadline", "Link", "Saved on", "Applied", "Emailed", "Replied",
          "Notes"]


def row_id(row):
    """Same posting from two scrapes must land on the same shortlist entry."""
    return "|".join([
        (row.get("Company") or "").strip().lower(),
        (row.get("Role / posting") or row.get("Role") or "").strip().lower()[:60],
        (row.get("Link") or "").strip().lower(),
    ])


def load():
    if not os.path.exists(STORE):
        return []
    try:
        with open(STORE, encoding="utf-8") as fh:
            return json.load(fh)
    except (json.JSONDecodeError, OSError):
        return []


def save_all(rows):
    with open(STORE, "w", encoding="utf-8") as fh:
        json.dump([{k: r.get(k, "") for k in FIELDS} for r in rows],
                  fh, ensure_ascii=False, indent=1)


def add(row):
    """Add one posting to the shortlist. Returns True if it was new."""
    rows = load()
    rid = row_id(row)
    if any(r["id"] == rid for r in rows):
        return False
    rows.append({
        "id": rid,
        "Company": row.get("Company", ""),
        "Country": row.get("Country", ""),
        "City": row.get("City", ""),
        "Role": row.get("Role / posting") or row.get("Role", ""),
        "Voice type": row.get("Voice type", ""),
        "Deadline": row.get("Deadline found") or row.get("Deadline", ""),
        "Link": row.get("Link", ""),
        "Saved on": date.today().isoformat(),
        "Applied": "No",
        "Emailed": "No",
        "Replied": "No",
        "Notes": "",
    })
    save_all(rows)
    return True


def remove(rid):
    rows = load()
    kept = [r for r in rows if r["id"] != rid]
    save_all(kept)
    return len(kept) != len(rows)


def update(rid, field, value):
    """Tick a box or edit a note."""
    if field not in ("Applied", "Emailed", "Replied", "Notes"):
        return False
    rows = load()
    for r in rows:
        if r["id"] == rid:
            r[field] = value
            save_all(rows)
            return True
    return False


def export_xlsx():
    """Write the shortlist as a working Excel tracker.

    Applied / Emailed / Replied are Yes-No dropdowns. A cell set to Yes turns
    green, and once you have applied the whole row is tinted, so the sheet
    shows progress at a glance.
    """
    rows = load()
    cols = [c for c in FIELDS if c != "id"]

    wb = Workbook()
    ws = wb.active
    ws.title = "My applications"
    ws.append(cols)
    for r in rows:
        ws.append([r.get(c, "") for c in cols])

    dress(ws, cols, [26, 10, 15, 42, 26, 14, 46, 11, 11, 11, 11, 40],
          tick_cols=("Applied", "Emailed", "Replied"))

    wb.save(XLSX)
    return XLSX, len(rows)


if __name__ == "__main__":
    path, n = export_xlsx()
    print(f"Wrote {path} ({n} saved jobs)")
