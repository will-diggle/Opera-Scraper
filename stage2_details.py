"""
STAGE 2 - get each organisation's real details from its Operabase page.

Every Operabase organisation page carries a machine-readable block giving the
organisation's name, type, city, country, official website and social media
accounts. We read that block - we do not scrape the visible page.

There are ~84,000 organisations, so this is the slow stage. It is designed to
be stopped and restarted: everything fetched is appended to a cache file and
never fetched twice. Press Ctrl-C any time; run it again to carry on.

Run:  .venv/bin/python stage2_details.py            (keeps going until done)
      .venv/bin/python stage2_details.py 2000       (do 2000 more, then stop)
Out:  operabase_details.jsonl  (the cache)
      operabase_details.csv    (tidy version, rewritten each run)
"""

import csv
import json
import os
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor

import requests

IN = "operabase_orgs.csv"
CACHE = "operabase_details.jsonl"
OUT = "operabase_details.csv"

WORKERS = 8
PAUSE = 0.15          # per worker, between requests - keeps us gentle
TIMEOUT = 25

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "en",
}

LDJSON = re.compile(r'<script type="application/ld\+json">(.*?)</script>', re.S)

SOCIALS = {
    "facebook": "facebook.com",
    "instagram": "instagram.com",
    "twitter_x": ("twitter.com", "x.com"),
    "youtube": "youtube.com",
    "linkedin": "linkedin.com",
    "tiktok": "tiktok.com",
}

_lock = threading.Lock()
_local = threading.local()


def session():
    if not hasattr(_local, "s"):
        _local.s = requests.Session()
    return _local.s


def parse_profile(html):
    """Pull the organisation block out of the page's structured data."""
    for m in LDJSON.finditer(html):
        try:
            d = json.loads(m.group(1))
        except json.JSONDecodeError:
            continue
        if d.get("@type") != "ProfilePage":
            continue
        ent = d.get("mainEntity") or {}
        addr = ent.get("address") or {}
        same = ent.get("sameAs") or []
        if isinstance(same, str):
            same = [same]

        out = {
            "name": ent.get("name", ""),
            "org_type": ent.get("@type", ""),
            "city": addr.get("addressLocality", ""),
            "country": addr.get("addressCountry", ""),
            "street": (addr.get("streetAddress", "") or "").replace("\n", ", "),
            "phone": ent.get("telephone", ""),
            "email": ent.get("email", ""),
        }
        for key, domains in SOCIALS.items():
            domains = (domains,) if isinstance(domains, str) else domains
            out[key] = next((u for u in same
                             if any(d in u.lower() for d in domains)), "")
        # whatever is left over and is not a social account is the real website
        social_domains = [d for v in SOCIALS.values()
                          for d in ((v,) if isinstance(v, str) else v)]
        out["website"] = next(
            (u for u in same
             if not any(d in u.lower() for d in social_domains)
             and "wikipedia.org" not in u and "wikidata.org" not in u), "")
        return out
    return None


def load_done():
    done = set()
    if os.path.exists(CACHE):
        with open(CACHE, encoding="utf-8") as fh:
            for line in fh:
                try:
                    done.add(json.loads(line)["operabase_id"])
                except Exception:
                    pass
    return done


def fetch_one(row):
    try:
        r = session().get(row["operabase_url"], timeout=TIMEOUT, headers=HEADERS)
        time.sleep(PAUSE)
        if r.status_code != 200:
            return {**row, "error": f"HTTP {r.status_code}"}
        info = parse_profile(r.text)
        if info is None:
            return {**row, "error": "no profile data"}
        return {**row, **info, "error": ""}
    except requests.RequestException as e:
        return {**row, "error": type(e).__name__}


COLS = ["operabase_id", "name", "org_type", "city", "country", "website",
        "facebook", "instagram", "twitter_x", "youtube", "linkedin", "tiktok",
        "phone", "email", "street", "operabase_url", "error"]


def write_csv():
    rows = []
    with open(CACHE, encoding="utf-8") as fh:
        for line in fh:
            try:
                rows.append(json.loads(line))
            except Exception:
                pass
    rows.sort(key=lambda r: (r.get("country", ""), r.get("name", "")))
    with open(OUT, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=COLS, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c, "") for c in COLS})
    return len(rows)


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else None

    with open(IN, newline="", encoding="utf-8") as fh:
        todo = list(csv.DictReader(fh))
    done = load_done()
    todo = [r for r in todo if r["operabase_id"] not in done]
    if limit:
        todo = todo[:limit]

    print(f"{len(done)} already cached, {len(todo)} to fetch this run")
    if not todo:
        print(f"Nothing to do. {write_csv()} rows in {OUT}")
        return

    start, n = time.time(), 0
    with open(CACHE, "a", encoding="utf-8") as cache:
        with ThreadPoolExecutor(max_workers=WORKERS) as pool:
            for res in pool.map(fetch_one, todo):
                with _lock:
                    cache.write(json.dumps(res, ensure_ascii=False) + "\n")
                    n += 1
                    if n % 250 == 0:
                        cache.flush()
                        rate = n / (time.time() - start)
                        left = (len(todo) - n) / rate / 60
                        print(f"  {n}/{len(todo)}  "
                              f"{rate:.1f}/sec  ~{left:.0f} min left", flush=True)

    print(f"\nDone this run: {n}. Total cached: {len(done) + n}")
    print(f"Wrote {OUT}: {write_csv()} rows")


if __name__ == "__main__":
    main()
