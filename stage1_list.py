"""
STAGE 1 - build the master list of organisations from Operabase.

Operabase publishes sitemaps listing every organisation page it has. We read
those (they are meant to be read by machines) and save the page addresses.
Nothing clever, nothing heavy: 17 requests total.

Run:  .venv/bin/python stage1_list.py
Out:  operabase_orgs.csv
"""

import csv
import re
import time

import requests

INDEX = "https://www.operabase.com/sitemap_organizations.xml"
OUT = "operabase_orgs.csv"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}
LOC = re.compile(r"<loc>(.*?)</loc>")
# .../london-festival-opera-o7/en  ->  slug "london-festival-opera", id "o7"
SLUG = re.compile(r"operabase\.com/(.+?)-(o\d+)/")


def main():
    s = requests.Session()
    index = s.get(INDEX, timeout=30, headers=HEADERS).text
    sub_maps = LOC.findall(index)
    print(f"{len(sub_maps)} sub-sitemaps to read")

    rows, seen = [], set()
    for i, sm in enumerate(sub_maps, 1):
        text = s.get(sm, timeout=60, headers=HEADERS).text
        urls = LOC.findall(text)
        for u in urls:
            m = SLUG.search(u)
            if not m or m.group(2) in seen:
                continue
            seen.add(m.group(2))
            rows.append({
                "operabase_id": m.group(2),
                "slug": m.group(1),
                "operabase_url": u,
            })
        print(f"  {i}/{len(sub_maps)}: {len(urls)} urls (running total {len(rows)})")
        time.sleep(1)

    with open(OUT, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["operabase_id", "slug", "operabase_url"])
        w.writeheader()
        w.writerows(rows)
    print(f"\nWrote {OUT}: {len(rows)} organisations")


if __name__ == "__main__":
    main()
