"""
STAGE 3 - turn the huge Operabase list into a sensible target list to scrape.

84,000 organisations is far too many to check every week, and most are festivals
or venues that never hire singers directly. This picks the ones worth checking
and writes them in the format the job scraper expects.

Run:  .venv/bin/python stage3_filter.py                 (UK/Ireland + Europe)
      .venv/bin/python stage3_filter.py GB IE DE AT CH  (just these countries)
      .venv/bin/python stage3_filter.py ALL             (everything, worldwide)
Out:  targets.csv
"""

import csv
import sys
from urllib.parse import urlparse

IN = "operabase_details.csv"
OUT = "targets.csv"

EUROPE = [
    "GB", "UK", "IE", "DE", "AT", "CH", "FR", "IT", "ES", "PT", "NL", "BE",
    "LU", "DK", "SE", "NO", "FI", "IS", "PL", "CZ", "SK", "HU", "RO", "BG",
    "GR", "HR", "SI", "RS", "EE", "LV", "LT", "MT", "CY", "MC", "LI", "AD",
]

# Organisation types on Operabase that can actually engage singers.
KEEP_TYPES = {"TheaterGroup", "PerformingGroup", "MusicGroup",
              "Organization", "TheaterEvent", "Festival", ""}


def domain(url):
    try:
        d = urlparse(url).netloc.lower()
        return d[4:] if d.startswith("www.") else d
    except ValueError:
        return ""


def main():
    args = [a.upper() for a in sys.argv[1:]]
    if args == ["ALL"]:
        wanted, label = None, "worldwide"
    elif args:
        wanted, label = set(args), ", ".join(args)
    else:
        wanted, label = set(EUROPE), "UK/Ireland + Europe"

    with open(IN, newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))

    kept, seen_domains, skipped = [], set(), {"no website": 0,
                                              "other country": 0,
                                              "duplicate site": 0,
                                              "other type": 0}
    for r in rows:
        if r["error"] or not r["website"].startswith("http"):
            skipped["no website"] += 1
            continue
        if wanted is not None and r["country"].upper() not in wanted:
            skipped["other country"] += 1
            continue
        if r["org_type"] not in KEEP_TYPES:
            skipped["other type"] += 1
            continue
        d = domain(r["website"])
        if not d or d in seen_domains:
            skipped["duplicate site"] += 1
            continue
        seen_domains.add(d)
        kept.append(r)

    kept.sort(key=lambda r: (r["country"], r["name"]))
    cols = ["name", "country", "website", "city", "operabase_id",
            "facebook", "instagram", "twitter_x", "youtube", "linkedin"]
    with open(OUT, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for r in kept:
            w.writerow({c: r.get(c, "") for c in cols})

    print(f"Filter: {label}")
    print(f"  read    {len(rows)} organisations")
    for k, v in skipped.items():
        print(f"  skipped {v} ({k})")
    print(f"\nWrote {OUT}: {len(kept)} companies to check")
    print("Now run:  .venv/bin/python scraper.py --file targets.csv")


if __name__ == "__main__":
    main()
