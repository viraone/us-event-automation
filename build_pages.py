"""Build the static GitHub Pages dashboard from the monthly gig CSVs.

Usage: python build_pages.py
Writes docs/data.json consumed by docs/index.html.
Re-run whenever the CSVs are updated, then commit and push.
"""

import csv
import glob
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(HERE, "docs")


def build():
    gigs = []
    for path in sorted(glob.glob(os.path.join(HERE, "US_Event_*_Gigs.csv"))):
        with open(path, newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                gigs.append(
                    {
                        "date": row.get("Date", ""),
                        "day": row.get("Day", ""),
                        "time": row.get("Start Time", ""),
                        "timeSort": row.get("Start Time Sort", ""),
                        "state": row.get("State", ""),
                        "store": row.get("Store / Account", ""),
                        "storeNum": row.get("Store Number", ""),
                        "city": row.get("City", ""),
                        "preferred": row.get("Preferred Area", ""),
                        "priority": row.get("Location Priority", ""),
                        "color": row.get("Calendar Color", ""),
                        "status": row.get("Status", ""),
                        "lastSeen": row.get("Last Seen", ""),
                    }
                )
    gigs.sort(key=lambda g: (g["date"], g["timeSort"]))
    os.makedirs(DOCS, exist_ok=True)
    out = os.path.join(DOCS, "data.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"gigs": gigs}, f, separators=(",", ":"))
    print(f"Wrote {len(gigs)} gigs -> {out}")


if __name__ == "__main__":
    build()
