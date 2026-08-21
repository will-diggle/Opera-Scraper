"""
Build the data file the public website serves.

The app on your Mac reads the raw scrape output. A public site should not: it
would expose every half-matched row and the whole 12,000-company sweep. This
writes one small, filtered file - the same listings as the shared page - which
is safe to commit and deploy.

Run:  .venv/bin/python make_public_data.py
Out:  public_data.json
"""

import json
import os

from make_share_page import load, data_date

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "public_data.json")

KEYS = ["Kind", "Company", "City", "Country", "Role", "Deadline", "Posted",
        "Freshness", "Link", "Checked"]


def main():
    rows = [dict(zip(KEYS, r)) for r in load()]
    payload = {"checked": data_date([r["Checked"] for r in rows]),
               "rows": rows}
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False)
    print(f"Wrote {OUT}: {len(rows)} openings, "
          f"{os.path.getsize(OUT)/1024:.0f} KB")


if __name__ == "__main__":
    main()
