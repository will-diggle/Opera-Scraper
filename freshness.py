"""
Work out how old a posting is, so past auditions stop looking current.

Two clues:
  1. Many sites put the date in the web address, e.g.
     bathopera.com/2025/12/22/bath-opera-open-auditions-...
  2. Failing that, a date written in the posting text.

Used by the app and the spreadsheet. Nothing here fetches anything.
"""

import re
from datetime import date

# /2025/12/22/ or /2025/12/ in a web address
URL_DATE = re.compile(r"/(20[0-2]\d)/(\d{1,2})(?:/(\d{1,2}))?(?:/|$)")

MONTHS = ("january february march april may june july august september "
          "october november december").split()
MONTHS_DE = ("januar februar märz april mai juni juli august september "
             "oktober november dezember").split()
MONTHS_IT = ("gennaio febbraio marzo aprile maggio giugno luglio agosto "
             "settembre ottobre novembre dicembre").split()

TEXT_DATE = re.compile(
    r"\b(\d{1,2})[./-](\d{1,2})[./-](20\d{2})\b"          # 03.07.2026
    r"|\b(20\d{2})-(\d{1,2})-(\d{1,2})\b"                  # 2026-07-03
    r"|\b(\d{1,2})\s+([a-zäöüà]+)\s+(20\d{2})\b", re.I)    # 3 July 2026


def _month_number(word):
    w = word.lower()
    for names in (MONTHS, MONTHS_DE, MONTHS_IT):
        for i, m in enumerate(names, start=1):
            if w.startswith(m[:3]) and len(w) >= 3:
                return i
    return None


def date_from_url(url):
    m = URL_DATE.search(url or "")
    if not m:
        return None
    y, mo = int(m.group(1)), int(m.group(2))
    d = int(m.group(3) or 1)
    try:
        return date(y, min(max(mo, 1), 12), min(max(d, 1), 28))
    except ValueError:
        return None


def date_from_text(*texts):
    for t in texts:
        m = TEXT_DATE.search(t or "")
        if not m:
            continue
        g = m.groups()
        try:
            if g[0]:
                return date(int(g[2]), int(g[1]), min(int(g[0]), 28))
            if g[3]:
                return date(int(g[3]), int(g[4]), min(int(g[5]), 28))
            mo = _month_number(g[7])
            if mo:
                return date(int(g[8]), mo, min(int(g[6]), 28))
        except (ValueError, TypeError):
            continue
    return None


def assess(row, today=None):
    """Return (posted_date_or_blank, freshness_label)."""
    today = today or date.today()
    d = date_from_url(row.get("Link", "")) or \
        date_from_url(row.get("Found on page", ""))
    source = "from web address"
    if d is None:
        d = date_from_text(row.get("Deadline found", ""),
                           row.get("Role / posting", ""))
        source = "from the text"

    if d is None:
        return "", "Unknown date"

    months = (today.year - d.year) * 12 + (today.month - d.month)
    if d > today:
        label = "Upcoming"          # a future date - probably a deadline
    elif months <= 2:
        label = "Recent"
    elif months <= 12:
        label = "Older"
    else:
        label = "Out of date"
    return f"{d.isoformat()} ({source})", label
