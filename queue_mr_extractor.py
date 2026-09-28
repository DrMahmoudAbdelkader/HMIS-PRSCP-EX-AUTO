"""
queue_mr_extractor  —  v12  (THIN ADAPTER over the decree repo's queue logic)
===========================================================================
This module no longer contains its own report fetcher or parser. It
imports the two files that make up the queue logic in the decree-lookup
repo, copied here UNCHANGED (so a fix in one repo can simply be re-copied
into the other):

    queue_extractor.py  ->  fetch_report(), verify_report_date_range(),
                            parse_clinic_report()      (HMIS WebReport export)
    (queue_parser.py stays in the repo for reference; see the note on its
     column map below.)

PUBLIC API (used by Extract_DMS_Patients_Prescriptioned_Services_Data_modified.py)
------------------------------------------------------------------------------
    get_queue_records(run_date_ddmmyyyy, output_dir=None) -> list[dict]
        One dict per queued patient: MR, Clinic, National ID,
        Appointment Number, Appointment Date, Serial No., Status, User,
        Old Medical No., Source.
    get_queue_mr_clinic_map(run_date_ddmmyyyy, output_dir=None) -> {MR: Clinic}
        Same contract as before (dict; {} on an empty day; RAISES on a real
        failure, never sys.exit).

THE DATE-RANGE QUIRK (why D -> D+1)
-----------------------------------
HMIS treats the report's "To" date as an EXCLUSIVE midnight cut-off (each
row's Appn_date carries a time, e.g. "01/09/2026 04.15.43"). A single-day
request From=To=01/09/2026 therefore returns "No Data Found", while
From=01/09/2026 To=02/09/2026 returns every appointment dated 01/09
(verified on a real export: 534 rows, all dated 01/09/2026, none of 02/09).
So for a run date D this module requests D -> D+1, then keeps only the rows
whose own appointment date is D. If that comes back empty it also tries
D -> D once, in case the server behaviour changes.

WHICH REPORT / WHICH PARSER
---------------------------
Reports are tried in this order (first one that yields rows for D wins):
    1. outpat_clnc_lst_det_sts_j   ("by Status" - verified layout below)
    2. outpat_clnc_lst_det_j       (parsed with queue_extractor.parse_clinic_report)
Override with QUEUE_REPORT_CODES="code1,code2".

Verified layout of the "by Status" export (one patient = TWO sheet rows):
    row 1:  D serial | H National ID | M MR code | U Appointment No. |
            Z Appn_date+time | AN Old medical (case no. such as 3707/2020)
    row 2:  AG Arrive_date+time | AK User
Column "Status" (AM) is printed in the header but is empty in every data
row of the sample. The fixed columns in queue_parser.py (Appn date=Y,
User=AC, Old medical=AI) do NOT match this export, which is why this module
reads the columns itself; queue_parser.py's Serial/ID/MR/Appointment-No.
columns (D/H/M/U) are correct.

MR CODE: "by Status" -> column M ("Patient File No." in queue_parser);
clinic-list report -> "Medical No.".

If a report comes back with zero rows, the first rows of the sheet are
printed so the cause (blank report vs. layout change) is visible in the
GitHub Actions log without downloading the artifact.
"""

import io
import os
import sys

import queue_extractor as qx
import re
from datetime import datetime, timedelta
from openpyxl import load_workbook

OUTPUT_DIR = os.environ.get("QUEUE_OUTPUT_DIR", r"D:\Queue_DMS_Data")


def _warn(msg):
    print(f"   !! {msg}")
    if os.environ.get("GITHUB_ACTIONS") == "true":
        print("::warning::" + str(msg).replace("\r", " ").replace("\n", " "))


def _report_codes():
    forced = os.environ.get("QUEUE_REPORT_CODES", "").strip()
    if forced:
        return [c.strip() for c in forced.split(",") if c.strip()]
    return ["outpat_clnc_lst_det_sts_j", "outpat_clnc_lst_det_j"]


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


def _is_no_data(content):
    """True when the server generated its blank 'No Data Found' sheet."""
    try:
        ws = load_workbook(io.BytesIO(content), data_only=True).active
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 30), values_only=True):
            for v in row:
                if isinstance(v, str) and "no data found" in v.lower():
                    return True
    except Exception:
        pass
    return False


_DT_RE = re.compile(r"^(\d{2}/\d{2}/\d{4})[ .T]?(\d{2}[.:]\d{2}(?:[.:]\d{2})?)?")


def _cell(ws, r, col, window=1):
    """Value at (r, col) or the nearest populated column within +/-window."""
    for d in [0] + [x for k in range(1, window + 1) for x in (-k, k)]:
        c = col + d
        if c >= 1:
            v = ws.cell(row=r, column=c).value
            if v not in (None, ""):
                return v.strip() if isinstance(v, str) else v
    return None


def _clean_num(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _parse_status_export(content):
    """Parse the verified 'by Status' layout -> list of raw dicts."""
    ws = load_workbook(io.BytesIO(content), data_only=True).active
    recs, clinic = [], None
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(row=r, column=5).value or "").strip() == "Clinic":
            v = _cell(ws, r, 10, window=2)
            if v:
                clinic = str(v).strip()
            continue
        serial = ws.cell(row=r, column=4).value
        if not (isinstance(serial, (int, float)) or
                (isinstance(serial, str) and serial.strip().isdigit())):
            continue
        appn = _cell(ws, r, 26, window=2)
        m = _DT_RE.match(str(appn or ""))
        arrive = _cell(ws, r + 1, 33, window=2)
        recs.append({
            "MR": _clean_num(_cell(ws, r, 13)),
            "Clinic": clinic or "",
            "National ID": _clean_num(_cell(ws, r, 8)),
            "Appointment Number": _clean_num(_cell(ws, r, 21)),
            "Appointment Date": m.group(1) if m else "",
            "Appointment Time": (m.group(2) or "").replace(".", ":") if m else "",
            "Serial No.": _clean_num(serial),
            "Status": "",
            "User": _clean_num(_cell(ws, r + 1, 37, window=1)),
            "Old Medical No.": _clean_num(_cell(ws, r, 40, window=1)),
            "Arrival": _s(arrive),
        })
    return recs


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


def _next_day(date_dash):
    d = datetime.strptime(date_dash, "%d-%m-%Y") + timedelta(days=1)
    return d.strftime("%d-%m-%Y")


def _norm_date(v):
    """dd/mm/yyyy (or dd-mm-yyyy [hh...]) -> 'dd-mm-yyyy', else ''."""
    m = re.match(r"^(\d{2})[/-](\d{2})[/-](\d{4})", str(v or "").strip())
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else ""


def _fetch_range(code, from_dash, to_dash, out_dir):
    """Fetch one (from, to) range -> (content, was_no_data). Saves the raw file."""
    f_slash, t_slash = qx.to_slash_date(from_dash), qx.to_slash_date(to_dash)
    print(f"   -> Fetching {code}  From={f_slash}  To={t_slash}")
    filename, content = qx.fetch_report(requests_session(), code, f_slash, t_slash)
    os.makedirs(out_dir, exist_ok=True)
    raw = os.path.join(out_dir, f"{from_dash}_to_{to_dash}_{filename}")
    with open(raw, "wb") as f:
        f.write(content)
    print(f"   -> Saved raw report ({len(content):,} bytes) -> {raw}")
    if _is_no_data(content):
        # the server's blank sheet has no From/To header -> check this BEFORE
        # the date-range verification, or an empty range looks like a failure
        return content, True
    qx.verify_report_date_range(content, f_slash, t_slash)   # raises on mismatch
    return content, False


def _rows_from(code, content, day):
    if code == "outpat_clnc_lst_det_sts_j":
        rows = _parse_status_export(content)
        layout = "by-status"
    else:
        recs, warns = qx.parse_clinic_report(content)
        rows = [_from_clinic_list(r) for r in recs]
        layout = "clinic-list"
        for w in warns:
            _warn(w)
    total = len(rows)
    rows = [r for r in rows if r["MR"] and _norm_date(r["Appointment Date"]) == day]
    dropped = total - len(rows)
    if dropped:
        print(f"      kept {len(rows)} rows dated {day} (dropped {dropped} "
              f"with another date or no MR)")
    for r in rows:
        r["Source"] = f"{code} ({layout})"
    return rows


def _pull_one(code, day, out_dir):
    """One report for one day: try D->D+1 (server's To is exclusive), then D->D."""
    for from_d, to_d in ((day, _next_day(day)), (day, day)):
        content, no_data = _fetch_range(code, from_d, to_d, out_dir)
        if no_data:
            _warn(f"{code}: server returned 'No Data Found' for {from_d} -> {to_d}.")
            continue
        rows = _rows_from(code, content, day)
        if rows:
            return rows
        _warn(f"{code}: {from_d} -> {to_d} gave no rows dated {day}. Head of the report:")
        _dump_head(content)
    return []


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
