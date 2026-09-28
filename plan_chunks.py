#!/usr/bin/env python3
"""
plan_chunks.py  -  step 2 of the bulk pipeline (stdlib only).

Reads every queue_<date>.json produced by queue_day.py and:

  1. De-duplicates MR codes WITHIN each day (first clinic seen wins, exactly
     like the single-day script).
  2. De-duplicates ACROSS the period: an MR that arrived on 01-09, 05-09 and
     13-09 becomes ONE patient carrying its three dates (and the clinic it was
     queued under on each date).
  3. Splits the unique patients into chunks of about --per-job patients,
     EVENLY (101 patients at 50/job -> 3 chunks of 34, not 50/50/1), because
     each patient page is opened once whatever number of days they have.
  4. Writes chunk_<n>.json files + an audit CSV and sets the GitHub outputs
     that drive the parallel extraction matrix.

    python plan_chunks.py --queue-dir queue-in --out-dir chunk-plan \
        --per-job 50 --max-chunks 250 [--fail-if-empty]
"""

import argparse
import csv
import glob
import json
import math
import os
import re
import sys
from datetime import datetime


def _key(d):
    return datetime.strptime(d, "%d-%m-%Y")


def set_output(name, value):
    path = os.environ.get("GITHUB_OUTPUT")
    if path:
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"{name}={value}\n")
    print(f"[output] {name}={value}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--queue-dir", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--per-job", type=int, default=50)
    ap.add_argument("--max-chunks", type=int, default=250,
                    help="GitHub matrices are limited to 256 jobs")
    ap.add_argument("--expected-days", default="",
                    help="JSON list of dd-mm-yyyy days the period should contain; "
                         "days without a queue file are reported as missing")
    ap.add_argument("--fail-if-empty", action="store_true",
                    help="exit 1 when no patient arrived at all (single-day runs)")
    a = ap.parse_args()

    per_job = max(1, a.per_job)
    os.makedirs(a.out_dir, exist_ok=True)

    files = sorted(glob.glob(os.path.join(a.queue_dir, "**", "queue_*.json"), recursive=True))
    if not files:
        print("::error::No queue_<date>.json files found - nothing to plan.")
        sys.exit(1)

    # mr -> {"dates": {date: clinic}, "national_id": str}
    patients = {}
    per_day_counts = {}
    for fp in files:
        with open(fp, encoding="utf-8") as f:
            day = json.load(f)
        d = day["date"]
        seen_today = set()
        for r in day["records"]:
            mr = str(r["MR"]).strip()
            if not mr or mr in seen_today:
                continue                       # step 1: dedup within the day
            seen_today.add(mr)
            p = patients.setdefault(mr, {"dates": {}, "national_id": ""})
            p["dates"][d] = r.get("Clinic", "") or ""
            if r.get("National ID") and not p["national_id"]:
                p["national_id"] = r["National ID"]
        per_day_counts[d] = len(seen_today)

    days = sorted(per_day_counts, key=_key)

    # A failed queue job leaves no file. Don't stop the pipeline for it (other
    # days are still valid): report it and let the final "verify" job go red.
    missing_days = []
    if a.expected_days.strip():
        expected = json.loads(a.expected_days)
        missing_days = [d for d in expected if d not in per_day_counts]
    if missing_days:
        print(f"::error::Queue collection FAILED for {len(missing_days)} day(s): "
              f"{', '.join(missing_days)}. Their patients are NOT in this run - "
              f"re-run those dates (safe: Supabase de-duplicates).")
    set_output("missing_days", ",".join(missing_days))
    with open(os.path.join(a.out_dir, "missing_days.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(missing_days))

    total_day_rows = sum(per_day_counts.values())
    n_patients = len(patients)
    saved = total_day_rows - n_patients

    print(f"Days read            : {len(days)}  ({days[0]} .. {days[-1]})")
    for d in days:
        print(f"   {d}: {per_day_counts[d]} unique MR")
    print(f"Patient-days (sum)   : {total_day_rows}")
    print(f"Unique MR in period  : {n_patients}")
    print(f"Chart openings saved : {saved} (repeat visits share one patient page)")

    # audit CSV: which MR came on which dates
    with open(os.path.join(a.out_dir, "mr_dates_map.csv"), "w",
              newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["MR", "Days queued", "Dates", "Clinics (per date)"])
        for mr in sorted(patients):
            ds = sorted(patients[mr]["dates"], key=_key)
            w.writerow([mr, len(ds), " | ".join(ds),
                        " | ".join(f"{d}: {patients[mr]['dates'][d]}" for d in ds)])

    if n_patients == 0:
        msg = "No arrived patients on any day of the period."
        if a.fail_if_empty:
            print(f"::error::{msg}")
            sys.exit(1)
        print(f"::notice::{msg} Nothing to extract.")
        set_output("has_chunks", "false")
        set_output("chunks", "[]")
        set_output("n_chunks", "0")
        return

    # size the chunks: as close to per_job as possible, all the same size
    n_chunks = math.ceil(n_patients / per_job)
    if n_chunks > a.max_chunks:
        n_chunks = a.max_chunks
        print(f"::warning::{n_patients} patients would need more than {a.max_chunks} "
              f"parallel jobs at {per_job}/job; using {n_chunks} bigger chunks "
              f"(~{math.ceil(n_patients / n_chunks)} patients each).")

    # round-robin over MR-sorted patients => equal sizes (differ by at most 1)
    buckets = [[] for _ in range(n_chunks)]
    for i, mr in enumerate(sorted(patients)):
        buckets[i % n_chunks].append(mr)

    for n, mrs in enumerate(buckets):
        payload = {"chunk": n, "patients": [
            {"mr": mr, "national_id": patients[mr]["national_id"],
             "dates": {d: patients[mr]["dates"][d]
                       for d in sorted(patients[mr]["dates"], key=_key)}}
            for mr in mrs]}
        with open(os.path.join(a.out_dir, f"chunk_{n}.json"), "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        n_days = sum(len(patients[m]["dates"]) for m in mrs)
        print(f"   chunk {n:>3}: {len(mrs)} patients, {n_days} patient-days")

    set_output("has_chunks", "true")
    set_output("chunks", json.dumps(list(range(n_chunks))))
    set_output("n_chunks", str(n_chunks))


if __name__ == "__main__":
    main()
