"""
queue_mr_extractor  —  v12  (THIN ADAPTER over the decree repo's queue logic)
===========================================================================
This module no longer contains its own report fetcher or parser. It
imports the two files that make up the queue logic in the decree-lookup
repo, copied here UNCHANGED (so a fix in one repo can simply be re-copied
into the other):

    queue_extractor.py  ->  fetch_report(), verify_report_date_range(),
                            parse_clinic_report()      (HMIS WebReport export)
    queue_parser.py     ->  extract_records_from_workbook_bytes()
                            (fixed-column parser: Clinic, Serial No.,
                             National ID Number, Patient File No.,
                             Appointment Number/Date, User,
                             Old Medical No., Status)

PUBLIC API (used by Extract_DMS_Patients_Prescriptioned_Services_Data_modified.py)
------------------------------------------------------------------------------
    get_queue_records(run_date_ddmmyyyy, output_dir=None) -> list[dict]
        One dict per queued patient: MR, Clinic, National ID,
        Appointment Number, Appointment Date, Serial No., Status, User,
        Old Medical No., Source.
    get_queue_mr_clinic_map(run_date_ddmmyyyy, output_dir=None) -> {MR: Clinic}
        Same contract as before (dict; {} on an empty day; RAISES on a real
        failure, never sys.exit).

WHICH REPORT / WHICH PARSER
---------------------------
Reports are tried in this order (first one that yields rows wins):
    1. queue_extractor.REPORT_CODE        (whatever the decree repo is set to)
    2. the other of outpat_clnc_lst_det_j / outpat_clnc_lst_det_sts_j
Each downloaded file is parsed with queue_parser first (National-ID layout)
and, if that finds nothing, with queue_extractor.parse_clinic_report()
(labelled Clinic/Doctor layout). This is needed because the two reports
have different sheet layouts and the decree repo's own files disagree on
which one is live (queue_extractor.py says det_j; queue_parser.py parses
the _sts_ layout). Set QUEUE_REPORT_CODES="code1,code2" to force an order.

MR CODE
-------
  * queue_parser layout : MR = "Patient File No." (column M). queue_parser.py
    labels it a guess; earlier work confirmed it holds MR-shaped values
    ("6891") while "Old Medical No." (AI) holds case numbers like 1132/2022.
  * clinic-list layout  : MR = "Medical No."

If a report comes back with zero rows, the first rows of the sheet are
printed so the cause (blank report vs. layout change) is visible in the
GitHub Actions log without downloading the artifact.
"""

import io
import os
import sys

import queue_extractor as qx
import queue_parser as qp
from openpyxl import load_workbook

ALT_CODES = {
    "outpat_clnc_lst_det_j": "outpat_clnc_lst_det_sts_j",
    "outpat_clnc_lst_det_sts_j": "outpat_clnc_lst_det_j",
}
OUTPUT_DIR = os.environ.get("QUEUE_OUTPUT_DIR", r"D:\Queue_DMS_Data")


def _warn(msg):
    print(f"   !! {msg}")
    if os.environ.get("GITHUB_ACTIONS") == "true":
        print("::warning::" + str(msg).replace("\r", " ").replace("\n", " "))


def _report_codes():
    forced = os.environ.get("QUEUE_REPORT_CODES", "").strip()
    if forced:
        return [c.strip() for c in forced.split(",") if c.strip()]
    first = qx.REPORT_CODE
    return [first, ALT_CODES.get(first, first)]


def _s(v):
    return "" if v is None else str(v).strip()


def _dump_head(content, rows=12):
    """Print the first non-empty rows of a report that parsed to 0 rows."""
    try:
        ws = load_workbook(io.BytesIO(content), data_only=True).active
        print(f"      sheet size: {ws.max_row} rows x {ws.max_column} cols")
        shown = 0
        for r in range(1, ws.max_row + 1):
            vals = [_s(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
            vals = [v for v in vals if v]
            if vals:
                print(f"      r{r}: {vals[:8]}")
                shown += 1
                if shown >= rows:
                    break
        if not shown:
            print("      (sheet is completely empty)")
    except Exception as e:
        print(f"      (could not open the file as a workbook: {e})")


def _from_parser(rec):
    return {
        "MR": _s(rec.get("Patient File No.")),
        "Clinic": _s(rec.get("Clinic")),
        "National ID": _s(rec.get("National ID Number")),
        "Appointment Number": _s(rec.get("Appointment Number")),
        "Appointment Date": _s(rec.get("Appointment Date")),
        "Serial No.": _s(rec.get("Serial No.")),
        "Status": _s(rec.get("Status")),
        "User": _s(rec.get("User")),
        "Old Medical No.": _s(rec.get("Old Medical No.")),
    }


def _from_clinic_list(rec):
    return {
        "MR": _s(rec.get("Medical No.")),
        "Clinic": _s(rec.get("Clinic")),
        "National ID": "",
        "Appointment Number": "",
        "Appointment Date": _s(rec.get("Date")),
        "Serial No.": "",
        "Status": "",
        "User": "",
        "Old Medical No.": "",
    }


def _pull_one(code, date_dash, out_dir):
    """Fetch -> save raw -> verify range -> parse. Returns normalized rows."""
    date_slash = qx.to_slash_date(date_dash)
    print(f"   -> Fetching {code} for {date_slash} (single day)")
    filename, content = qx.fetch_report(requests_session(), code, date_slash, date_slash)

    os.makedirs(out_dir, exist_ok=True)
    raw = os.path.join(out_dir, f"{date_dash}_to_{date_dash}_{filename}")
    with open(raw, "wb") as f:
        f.write(content)
    print(f"   -> Saved raw report ({len(content):,} bytes) -> {raw}")

    qx.verify_report_date_range(content, date_slash, date_slash)   # raises on mismatch

    rows = [_from_parser(r) for r in qp.extract_records_from_workbook_bytes(content)]
    layout = "queue_parser"
    if not rows:
        recs, warns = qx.parse_clinic_report(content)
        rows = [_from_clinic_list(r) for r in recs]
        layout = "clinic-list"
        for w in warns:
            _warn(w)
    rows = [r for r in rows if r["MR"]]
    for r in rows:
        r["Source"] = f"{code} ({layout})"
    if not rows:
        _warn(f"{code} produced 0 usable rows for {date_dash}. Head of the report:")
        _dump_head(content)
    return rows


def requests_session():
    import requests
    return requests.Session()


def get_queue_records(run_date_ddmmyyyy, output_dir=None):
    out_dir = output_dir if output_dir is not None else OUTPUT_DIR
    errors = []
    for i, code in enumerate(_report_codes()):
        try:
            rows = _pull_one(code, run_date_ddmmyyyy, out_dir)
        except Exception as e:
            _warn(f"{code} failed: {e}")
            errors.append(f"{code}: {e}")
            continue
        if rows:
            if i:
                _warn(f"Primary report gave nothing; using {code} ({len(rows)} rows).")
            return rows
    if errors and len(errors) == len(_report_codes()):
        # every report errored (not merely empty) -> a real failure, not a quiet day
        raise RuntimeError("All queue reports failed for "
                           f"{run_date_ddmmyyyy}: " + " | ".join(errors))
    return []   # reachable reports were genuinely empty


def get_queue_mr_clinic_map(run_date_ddmmyyyy, output_dir=None):
    rows = get_queue_records(run_date_ddmmyyyy, output_dir)
    out, conflicts = {}, []
    for r in rows:
        mr = r["MR"]
        if mr in out:
            if out[mr] != r["Clinic"]:
                conflicts.append((mr, out[mr], r["Clinic"]))
            continue
        out[mr] = r["Clinic"]
    for mr, kept, dropped in conflicts[:10]:
        print(f"   !! MR {mr} under several clinics: kept {kept!r}, also saw {dropped!r}")
    return out


if __name__ == "__main__":
    day = sys.argv[1] if len(sys.argv) > 1 else __import__("datetime").datetime.now().strftime("%d-%m-%Y")
    recs = get_queue_records(day)
    print(f"{len(recs)} queue rows, {len({r['MR'] for r in recs})} unique MR")
