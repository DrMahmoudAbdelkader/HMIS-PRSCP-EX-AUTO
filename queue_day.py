#!/usr/bin/env python3
"""
queue_day.py  -  collect ONE day's arrived-patient queue and save it as JSON.

Step 1 of the bulk pipeline. The workflow runs one copy of this per day, all
in parallel. Nothing is extracted from patient charts here - it only records
which MR codes arrived at which clinic on that day.

    python queue_day.py 01-09-2026 run-output/queue_01-09-2026.json

Output JSON:
    {"date": "01-09-2026",
     "records": [{"MR": "...", "Clinic": "...", "National ID": "..."}, ...]}

A day with no arrived patients (Friday/Saturday, holiday) is NOT an error:
it writes "records": [] and exits 0. A real failure (login rejected, dashboard
unreachable, wrong day shown) is retried once with a fresh login and then
exits 1, so the day shows red and the later steps don't silently miss it.
The audit workbook <date>_queue_direct.xlsx is written next to the JSON by
queue_mr_extractor itself.
"""

import json
import os
import re
import sys
import time
import traceback

import queue_mr_extractor

ATTEMPTS = int(os.environ.get("QUEUE_DAY_ATTEMPTS", "2"))


def main():
    if len(sys.argv) < 3 or not re.fullmatch(r"\d{2}-\d{2}-\d{4}", sys.argv[1]):
        print("usage: python queue_day.py dd-mm-yyyy out.json")
        sys.exit(2)
    day, out_path = sys.argv[1], sys.argv[2]
    out_dir = os.path.dirname(out_path) or "."
    os.makedirs(out_dir, exist_ok=True)

    rows, last_err = None, None
    for attempt in range(1, ATTEMPTS + 1):
        print(f"── Queue {day}  (attempt {attempt}/{ATTEMPTS})")
        try:
            rows = queue_mr_extractor.get_queue_records(day, output_dir=out_dir)
            break
        except Exception as e:
            last_err = e
            print(f"❌  Queue pull failed for {day}: {e}")
            traceback.print_exc()
            if attempt < ATTEMPTS:
                time.sleep(20)

    if rows is None:
        print(f"::error::Could not pull the queue for {day}: {last_err}")
        sys.exit(1)

    records = []
    for r in rows:
        mr = str(r.get("MR", "")).strip()
        if not mr:
            continue
        records.append({"MR": mr,
                        "Clinic": r.get("Clinic", "") or "",
                        "National ID": str(r.get("National ID", "") or "").strip()})

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"date": day, "records": records}, f, ensure_ascii=False)

    n_mr = len({r["MR"] for r in records})
    if not records:
        print(f"::notice::{day}: no arrived patients (weekend/holiday or empty queue)")
    print(f"✅  {day}: {len(records)} arrived row(s), {n_mr} unique MR -> {out_path}")


if __name__ == "__main__":
    main()
