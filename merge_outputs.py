#!/usr/bin/env python3
"""
merge_outputs.py  -  final step of the bulk pipeline.

Merges every chunk_<n>_result.json (one per parallel extraction job) into ONE
workbook with the same sheets as the normal output: Patient Summary,
Investigations, Medical Reports (033), MDT (02), Statistics, Run Info.

    python merge_outputs.py --results-dir results --out run-output/Combined.xlsx \
        --label "01-09-2026 to 28-09-2026" --expected-chunks 40

Exit code 1 if any expected chunk result is missing, or any chunk reported a
Supabase push failure - the workbook is still written from whatever exists,
so nothing already extracted is lost.
"""

import argparse
import glob
import json
import os
import re
import sys
from datetime import datetime

# Reuse the exact workbook/statistics writers of the extraction script.
# (Importing it is safe: its CLI parsing only runs when executed as a script.)
import Extract_DMS_Patients_Prescriptioned_Services_Data_modified as ex


def _date_key(d):
    try:
        return datetime.strptime(str(d), "%d-%m-%Y")
    except Exception:
        return datetime.max


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--results-dir", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--label", default="")
    ap.add_argument("--expected-chunks", type=int, default=0)
    ap.add_argument("--missing-days", default="",
                    help="comma-separated days whose queue collection failed (info only)")
    a = ap.parse_args()

    files = sorted(glob.glob(os.path.join(a.results_dir, "**", "chunk_*_result.json"),
                             recursive=True))
    results = []
    for fp in files:
        with open(fp, encoding="utf-8") as f:
            results.append(json.load(f))
    got = {r["chunk"] for r in results}
    missing = sorted(set(range(a.expected_chunks)) - got) if a.expected_chunks else []

    merged, added = {}, {}
    for key in ex._SHEET_KEYS:
        rows = [row for r in results for row in r["rows"][key]]
        merged[key], added[key] = ex.merge_rows([], rows)
    # chronological order (Run Date column: index 2 for summary, last for the rest)
    idx = {"summary": 2, "investigations": -1, "medical_reports": -1, "mdt": -1}
    for key in ex._SHEET_KEYS:
        merged[key].sort(key=lambda row: (_date_key(row[idx[key]]), str(row[0])))

    errors = [(r["chunk"], m, why) for r in results for m, why in r["errors"]]
    push_failed = [r["chunk"] for r in results
                   if str(r.get("supabase_status", "")).startswith("partial")]
    statuses = {str(r.get("supabase_status", "")) for r in results} or {"n/a"}

    processed = len({row[0] for row in merged["summary"]})
    queued = sum(r["patients_in_chunk"] for r in results)
    run_info = {
        "run_date":            a.label or "bulk period",
        "timestamp":           datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "mr_from_queue":       queued,
        "patients_processed":  processed,
        "new_investigations":  len(merged["investigations"]),
        "new_medical_reports": len(merged["medical_reports"]),
        "new_mdt":             len(merged["mdt"]),
        "errors":              len(errors),
        "supabase_status":     ", ".join(sorted(statuses)),
    }
    ex.write_excel(merged, a.out, run_info)

    lines = [
        f"### Bulk extraction {a.label}".strip(),
        f"- Chunks with results: **{len(got)}**"
        + (f" of {a.expected_chunks}" if a.expected_chunks else ""),
        f"- Unique patients extracted: **{processed}** of {queued} queued",
        f"- Investigations: {len(merged['investigations'])} | "
        f"Medical reports: {len(merged['medical_reports'])} | "
        f"MDT forms: {len(merged['mdt'])}",
        f"- Patients skipped / errored: **{len(errors)}**",
    ]
    if a.missing_days.strip():
        lines.append(f"- ⚠️ Days whose QUEUE collection failed (not in this workbook): "
                     f"{a.missing_days}")
    if missing:
        lines.append(f"- ⚠️ Missing chunk results: {missing} (that job failed - re-run "
                     f"is safe, Supabase de-duplicates by row_hash)")
    if push_failed:
        lines.append(f"- ⚠️ Supabase push failed in chunks: {push_failed}")
    if errors:
        lines.append("\n| Chunk | MR | Reason |\n|---|---|---|")
        lines += [f"| {c} | {m} | {str(w)[:120]} |" for c, m, w in errors[:100]]
        if len(errors) > 100:
            lines.append(f"\n…and {len(errors) - 100} more (see the job log).")
    summary = "\n".join(lines)
    print("\n" + summary)
    sp = os.environ.get("GITHUB_STEP_SUMMARY")
    if sp:
        with open(sp, "a", encoding="utf-8") as f:
            f.write(summary + "\n")

    if missing or push_failed:
        print("::error::Some chunks are missing or failed to push to Supabase.")
        sys.exit(1)


if __name__ == "__main__":
    main()
