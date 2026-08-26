"""
What each scan actually covered, kept in one file the website can read.

The point is that the numbers on the site are produced by the scan itself.
Nobody has to remember them, and they cannot drift from reality.

A long scan often runs in more than one go (it resumes after a dropped
connection), so figures accumulate within a sweep and reset when a sweep
starts fresh.
"""

import json
import os
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
FILE = os.path.join(HERE, "scan_stats.json")


def read():
    if not os.path.exists(FILE):
        return {}
    try:
        with open(FILE, encoding="utf-8") as fh:
            return json.load(fh)
    except (json.JSONDecodeError, OSError):
        return {}


def record(section, numbers, fresh_sweep, extra=None):
    """Store one scan's figures.

    fresh_sweep=True starts the totals again (a new sweep); False adds to
    them (finishing a sweep that stopped part way).
    """
    all_stats = read()
    prev = {} if fresh_sweep else all_stats.get(section, {})
    merged = dict(prev)
    for key, value in numbers.items():
        merged[key] = (0 if fresh_sweep else prev.get(key, 0)) + value
    merged.update(extra or {})
    merged["last_run"] = date.today().isoformat()
    all_stats[section] = merged
    with open(FILE, "w", encoding="utf-8") as fh:
        json.dump(all_stats, fh, indent=1)
    return merged
