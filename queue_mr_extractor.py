"""
queue_mr_extractor  —  v11  (IMPORTABLE MODULE for the coupled pipeline)
====================================================================
PUBLIC API
-----------
    queue_mr_extractor.get_queue_mr_clinic_map(run_date_ddmmyyyy,
                                                output_dir=None)
        -> {MR_code: Clinic_name, ...}

This is what Extract_DMS_Patients_Prescriptioned_Services_Data_
modified.py imports and calls at startup to get the day's queued MR
codes and which clinic each one belongs to. The contract is unchanged
from v10:
  * returns a plain {MR: Clinic} dict ({} on a genuinely empty day);
  * RAISES on any real failure (never sys.exit) so the caller's
    `except Exception` block handles it;
  * if one MR shows up under several clinics, the first one seen is
    kept and every conflict is printed.

v11 — ALIGNED WITH THE DECREE-LOOKUP REPO'S QUEUE LOGIC
---------------------------------------------------------
The decree-lookup repo (queue_extractor.py v6b + queue_parser.py) is the
queue logic that is currently working in production. This module now
follows it:

  1. PRIMARY REPORT = "outpat_clnc_lst_det_j" ("Clinic List Detail"),
     exactly like queue_extractor.py. The "by Status" variant
     (outpat_clnc_lst_det_sts_j) started coming back blank on the live
     HMIS site (a server-side problem), which is why the decree repo
     reverted; this module had still been pointed at it.
  2. PARSER = queue_extractor.parse_clinic_report(), copied verbatim.
     It reads the labelled Clinic/Resource/Doctor blocks, so Medical No.
     is a real labelled column (no positional guessing), and it
     cross-checks itself against the report's own per-block "Total No.
     of Patient" and "Grand Total" figures (mismatches are printed as
     loud warnings, and as GitHub Actions ::warning:: annotations when
     run in CI).
  3. FALLBACK (on by default, QUEUE_STATUS_FALLBACK=0 to disable): if the
     primary report parses to ZERO rows, the "by Status" report is tried
     once with the offset-based parser this module used in v10
     (parse_status_report()). It reads the same fixed columns that
     queue_parser.py's COLUMNS dict uses (ID Number = H, MR-shaped
     "Patient File No." = M, Appointment Date = Y, Old Medical No. = AI),
     which is why M is treated as the MR code here — see below. Either
     report can be blank on a given day, so trying both avoids a
     spurious "nothing to extract" on the first zero-row result.
  4. Same fetch/AJAX-dateSelect/verify engine as queue_extractor.py
     (it was already identical apart from a looser ViewState regex, kept
     here because it also matches the real "j_id1:javax.faces.ViewState:0"
     update id), and the same CLEAN-file naming: the CLEAN file is named
     from the server's real filename, like the decree repo does.
  5. OUTPUT_DIR is still overridable via QUEUE_OUTPUT_DIR (needed on the
     Linux CI runner).

About the "by Status" report's MR column (fallback path only)
--------------------------------------------------------------
queue_parser.py labels column M "Patient File No." (a guess). This
module's v10 verified that same column against real records: it holds
MR-shaped values ("6891", "9334"), while the labelled "Old medical"
column (AI) holds case/file numbers such as "1132/2022" and must never
be used as an MR. So M -> "Medical No." here.

Output (both when imported and when run standalone)
-------------------------------------------------------
  1. <daterange>_<server filename>.xlsx      <- raw file as generated
  2. <daterange>_<report>_CLEAN.xlsx         <- flat 15-column table

Standalone:  python queue_mr_extractor.py [dd-mm-yyyy [dd-mm-yyyy]]
             (defaults to today; one date = single day)

pip install requests openpyxl
"""

import os
import re
import sys
import html as html_module
from datetime import datetime
import requests
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

# ═════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═════════════════════════════════════════════════════════════════

HOST            = "41.33.24.254:8080"
WEBREPORT_BASE  = f"http://{HOST}/WebReport-JWEB"

REPORT_CODE        = "outpat_clnc_lst_det_j"       # "Clinic List Detail" — primary, same as the decree repo
STATUS_REPORT_CODE = "outpat_clnc_lst_det_sts_j"   # "Clinic List Detail - by Status" — fallback only
LANG            = "L"
HSCD            = "01"                             # hospital/branch code

# Try the status report when the primary one parses to zero rows.
ENABLE_STATUS_FALLBACK = os.environ.get("QUEUE_STATUS_FALLBACK", "1") != "0"

# ---- Optional: restrict to one physician/resource. Leave both blank for ALL. ----
RESOURCE_ID   = ""
RESOURCE_NAME = ""

TIMEOUT = 60
# Overridable via env so this same file runs unchanged locally on Windows
# (defaults to your D: drive) and on a Linux CI runner (the workflow sets
# QUEUE_OUTPUT_DIR).
OUTPUT_DIR = os.environ.get("QUEUE_OUTPUT_DIR", r"D:\Queue_DMS_Data")

# ═════════════════════════════════════════════════════════════════
# HTTP CONSTANTS
# ═════════════════════════════════════════════════════════════════

_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
       "AppleWebKit/537.36 (KHTML, like Gecko) "
       "Chrome/150.0.0.0 Safari/537.36")

NAV_GET_HEADERS = {
    "Accept": ("text/html,application/xhtml+xml,application/xml;q=0.9,"
               "image/avif,image/webp,image/apng,*/*;q=0.8,"
               "application/signed-exchange;v=b3;q=0.7"),
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": _UA,
}

FORM_POST_HEADERS = {
    "Accept": ("text/html,application/xhtml+xml,application/xml;q=0.9,"
               "image/avif,image/webp,image/apng,*/*;q=0.8,"
               "application/signed-exchange;v=b3;q=0.7"),
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Cache-Control": "max-age=0",
    "Content-Type": "application/x-www-form-urlencoded",
    "Origin": f"http://{HOST}",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": _UA,
}

# Headers for the JSF/PrimeFaces partial-AJAX "dateSelect" replay (see
# sync_date_field() below for why this round trip is required).
AJAX_POST_HEADERS = {
    "Accept": "application/xml, text/xml, */*; q=0.01",
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Faces-Request": "partial/ajax",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": f"http://{HOST}",
    "User-Agent": _UA,
}

INPUT_RE  = re.compile(r'<input\b([^>]*?)/?>', re.I)
SELECT_RE = re.compile(r'<select\b([^>]*)>(.*?)</select>', re.I | re.S)
NAME_RE   = re.compile(r'name\s*=\s*"([^"]*)"')
VALUE_RE  = re.compile(r'value\s*=\s*"([^"]*)"')
TYPE_RE   = re.compile(r'type\s*=\s*"([^"]*)"', re.I)
CHECKED_RE = re.compile(r'\bchecked\s*=\s*"checked"', re.I)
OPTION_SELECTED_RE = re.compile(r'<option\s+value="([^"]*)"\s+selected="selected"', re.I)
OPTION_FIRST_RE    = re.compile(r'<option\s+value="([^"]*)"', re.I)

# Matches: mojarra.jsfcljs(document.getElementById('form'),{'ID':'ID'},'')
#          ...<input type="button" ... value="Export Excel" />
EXPORT_BTN_RE = re.compile(
    r"jsfcljs\(document\.getElementById\('(\w+)'\),\{'([^']+)':'[^']+'\}"
    r"[^)]*\)[^<]*<input[^>]*value=\"Export (\w+)\"",
    re.I,
)


# ═════════════════════════════════════════════════════════════════
# FORM-FIELD REPLAY ENGINE
# ═════════════════════════════════════════════════════════════════

def extract_form_fields(html_text):
    """
    Pull every submittable <input>/<select> field out of the page, in
    document order, preserving duplicates (JSF renders the same
    javax.faces.ViewState hidden field once per naming-container form,
    and the real browser submits it twice — we replicate that exactly).

    Buttons are skipped (they're added explicitly via find_export_button).
    Unchecked checkboxes/radios are skipped (browsers never submit them).
    """
    fields = []

    for m in INPUT_RE.finditer(html_text):
        attrs = m.group(1)
        name_m = NAME_RE.search(attrs)
        if not name_m:
            continue
        name = html_module.unescape(name_m.group(1))

        typ_m = TYPE_RE.search(attrs)
        typ = typ_m.group(1).lower() if typ_m else "text"
        if typ == "button":
            continue
        if typ in ("checkbox", "radio") and not CHECKED_RE.search(attrs):
            continue

        val_m = VALUE_RE.search(attrs)
        value = html_module.unescape(val_m.group(1)) if val_m else ""
        fields.append([name, value])

    for m in SELECT_RE.finditer(html_text):
        attrs, body = m.groups()
        name_m = NAME_RE.search(attrs)
        if not name_m:
            continue
        name = html_module.unescape(name_m.group(1))
        opt_m = OPTION_SELECTED_RE.search(body) or OPTION_FIRST_RE.search(body)
        value = html_module.unescape(opt_m.group(1)) if opt_m else ""
        fields.append([name, value])

    return fields


_DDMMYYYY_RE = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}$')


def find_date_range_fields(fields):
    """
    Locate the report's OWN From/To date inputs — NOT the unrelated
    fromDate_input/toDate_input that live inside the generic patient-
    search panel on the same page (that panel's date fields are always
    blank on page load; the report's own date fields come pre-populated
    with a real dd/mm/yyyy default, which is what we key off of here).
    """
    value_by_name = {}
    for n, v in fields:
        value_by_name.setdefault(n, v)  # first occurrence wins

    from_names = [n for n in value_by_name if n.endswith(':fromDate_input')]
    to_names   = [n for n in value_by_name if n.endswith(':toDate_input')]

    candidates = []
    for fn in from_names:
        m = re.match(r'^(.+):(\d+):fromDate_input$', fn)
        if not m:
            continue
        prefix, idx = m.group(1), int(m.group(2))
        tn = f"{prefix}:{idx + 1}:toDate_input"
        if tn in to_names:
            fv, tv = value_by_name[fn], value_by_name[tn]
            populated = bool(_DDMMYYYY_RE.match(fv or "")) and bool(_DDMMYYYY_RE.match(tv or ""))
            candidates.append((populated, prefix, fn, tn))

    candidates.sort(key=lambda c: (not c[0]), reverse=False)
    for populated, prefix, fn, tn in candidates:
        if populated:
            print(f"   -> Date fields detected: {fn}={value_by_name[fn]!r}  "
                  f"{tn}={value_by_name[tn]!r}")
            return fn, tn, f"{prefix}:0:physcResourceId", f"{prefix}:0:physcResourceName"

    if candidates:
        populated, prefix, fn, tn = candidates[0]
        print(f"   !! No pre-populated From/To pair found; falling back to "
              f"{fn} / {tn} (currently blank). Verify the result's own "
              f"From/To header after this run.")
        return fn, tn, f"{prefix}:0:physcResourceId", f"{prefix}:0:physcResourceName"

    raise RuntimeError(
        "Could not find ANY fromDate_input/toDate_input field pair with "
        "consecutive indices on the report page. The page structure has "
        "changed — send an updated HAR capture of the manual export."
    )


def find_export_button(html_text, label="Excel"):
    """Find the (form_id, field_name) for the 'Export <label>' button."""
    for form_id, field_name, btn_label in EXPORT_BTN_RE.findall(html_text):
        if btn_label.lower() == label.lower():
            return form_id, field_name
    return None, None


def set_field(fields, name, value):
    """Set value for every occurrence of `name` in the fields list (in place)."""
    hit = False
    for pair in fields:
        if pair[0] == name:
            pair[1] = value
            hit = True
    return hit


# ═════════════════════════════════════════════════════════════════
# REPORT FETCH
# ═════════════════════════════════════════════════════════════════

VIEWSTATE_UPDATE_RE = re.compile(
    r'<update\s+id="[^"]*ViewState[^"]*"[^>]*>\s*<!\[CDATA\[(.*?)\]\]>\s*</update>',
    re.S,
)


def sync_date_field(session, post_url, referer, fields, input_field_name, value):
    """
    Replays the browser's AJAX "dateSelect" event for a calendar field.

    The date fields are PrimeFaces/JSF calendar widgets: the date you
    see typed in the box is only a client-side display value. It does
    NOT get bound to the server-side report parameters through an
    ordinary form submit. It's only bound when the calendar's JS widget
    fires a dedicated partial-AJAX request for the "dateSelect" event.
    Skip that round trip, and the server just keeps whatever date was
    already in its ViewState no matter what raw value is submitted
    afterwards.

    This function fires that AJAX event for one date field, then reads
    the fresh javax.faces.ViewState token out of the server's partial-
    response XML and folds it back into `fields` (mutated in place) so
    the next request carries the server's up-to-date state.

    Returns True if the server acknowledged with an updated ViewState,
    False otherwise (a warning is also printed in that case).
    """
    component_id = (input_field_name[:-len("_input")]
                     if input_field_name.endswith("_input") else input_field_name)

    set_field(fields, input_field_name, value)

    ajax_fields = [
        ["javax.faces.partial.ajax", "true"],
        ["javax.faces.source", component_id],
        ["javax.faces.partial.execute", component_id],
        ["javax.faces.behavior.event", "dateSelect"],
        ["javax.faces.partial.event", "dateSelect"],
    ] + list(fields)

    headers = dict(AJAX_POST_HEADERS)
    headers["Referer"] = referer

    resp = session.post(post_url, headers=headers, data=ajax_fields, timeout=TIMEOUT)
    resp.raise_for_status()

    m = VIEWSTATE_UPDATE_RE.search(resp.text)
    if not m:
        print(f"   !! AJAX dateSelect for {input_field_name}={value} did not "
              f"return an updated ViewState (response started with: "
              f"{resp.text[:150]!r}). The date change may not register "
              f"server-side — check the final report's own From/To header.")
        return False

    new_viewstate = m.group(1)
    for pair in fields:
        if pair[0] == "javax.faces.ViewState":
            pair[1] = new_viewstate
    return True


def looks_like_report_form(html_text):
    return ('javax.faces.ViewState' in html_text
            and 'fromDate_input' in html_text
            and 'txtUsername' not in html_text)  # not a login page


def fetch_report(session, report_code, date_from_ddmmyyyy, date_to_ddmmyyyy,
                  resource_id="", resource_name=""):
    """
    Runs the two-step report-export flow and returns (filename, content_bytes).
    Raises RuntimeError with a diagnostic snippet on failure.
    """
    get_url = f"{WEBREPORT_BASE}/?repCode={report_code}&&lang={LANG}&&hscd={HSCD}"
    print(f"   -> GET  {get_url}")
    r = session.get(get_url, headers=NAV_GET_HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    page_html = r.text

    if not looks_like_report_form(page_html):
        snippet = page_html[:1000]
        raise RuntimeError(
            "Response doesn't look like the report parameter page "
            "(no ViewState/date fields found, or it looks like a login "
            "page). First 1000 chars:\n" + snippet
        )

    fields = extract_form_fields(page_html)
    from_key, to_key, res_id_key, res_name_key = find_date_range_fields(fields)

    post_url = f"{WEBREPORT_BASE}/faces/report/HMISReports.xhtml"

    print(f"   -> Replaying dateSelect for {from_key} = {date_from_ddmmyyyy}")
    sync_date_field(session, post_url, get_url, fields, from_key, date_from_ddmmyyyy)

    print(f"   -> Replaying dateSelect for {to_key} = {date_to_ddmmyyyy}")
    sync_date_field(session, post_url, get_url, fields, to_key, date_to_ddmmyyyy)

    if resource_id:
        set_field(fields, res_id_key, resource_id)
    if resource_name:
        set_field(fields, res_name_key, resource_name)

    form_id, export_field = find_export_button(page_html, label="Excel")
    if not export_field:
        raise RuntimeError(
            "Could not find the 'Export Excel' button on the report page. "
            "The page layout may differ from the captured session."
        )
    fields.append([export_field, export_field])

    post_headers = dict(FORM_POST_HEADERS)
    post_headers["Referer"] = get_url

    print(f"   -> POST {post_url}  ({from_key}={date_from_ddmmyyyy}  "
          f"{to_key}={date_to_ddmmyyyy})")
    resp = session.post(post_url, headers=post_headers, data=fields, timeout=TIMEOUT)
    resp.raise_for_status()

    content_disp = resp.headers.get("Content-Disposition", "")
    is_binary_xlsx = resp.content[:2] == b"PK"

    if "attachment" not in content_disp.lower() or not is_binary_xlsx:
        snippet = resp.text[:1000] if not is_binary_xlsx else "<binary, but no attachment header>"
        raise RuntimeError(
            "Report submission didn't return a file download as expected "
            f"(Content-Disposition={content_disp!r}). First part of response:\n"
            + snippet
        )

    fname_m = re.search(r'filename=([^;]+)', content_disp)
    filename = fname_m.group(1).strip() if fname_m else f"{report_code}.xlsx"
    return filename, resp.content


def verify_report_date_range(xlsx_bytes, expected_from_ddmmyyyy, expected_to_ddmmyyyy):
    """
    Reads the "From ... To ..." header the SERVER printed on the returned
    report and compares it to what was actually requested. Raises
    RuntimeError (raw file still saved) on a mismatch rather than
    silently saving data for the wrong range.
    """
    import io
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active

    actual_from = actual_to = None
    for r in range(1, min(20, ws.max_row) + 1):
        row_cells = {c: ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)}
        row_cells = {c: v for c, v in row_cells.items() if v is not None and str(v).strip() != ""}
        cols = sorted(row_cells)
        for c in cols:
            label = str(row_cells[c]).strip()
            if label == "From":
                later = [c2 for c2 in cols if c2 > c]
                if later:
                    actual_from = str(row_cells[later[0]]).strip()
            if label == "To":
                later = [c2 for c2 in cols if c2 > c]
                if later:
                    actual_to = str(row_cells[later[0]]).strip()
        if actual_from and actual_to:
            break

    if actual_from is None or actual_to is None:
        raise RuntimeError(
            "Could not find the 'From ... To ...' header on the returned "
            "report to verify the date range — check the raw file manually."
        )

    if actual_from != expected_from_ddmmyyyy or actual_to != expected_to_ddmmyyyy:
        raise RuntimeError(
            f"DATE RANGE MISMATCH — you requested From={expected_from_ddmmyyyy} "
            f"To={expected_to_ddmmyyyy}, but the report the server generated "
            f"is actually From={actual_from} To={actual_to}. The raw file has "
            f"still been saved so you can inspect it. Do NOT use the CLEAN "
            f"output from this run — it would silently describe the wrong "
            f"date range."
        )

    return actual_from, actual_to


# ═════════════════════════════════════════════════════════════════
# CLEAN-TABLE PARSERS — shared helpers
# ═════════════════════════════════════════════════════════════════

CLEAN_HEADERS = [
    "Date", "Day", "Clinic",
    "Resource ID", "Resource Name",
    "Doctor ID", "Doctor Name",
    "Medical No.", "Time", "Slots", "Patient Name",
    "Financial Cat. Code", "Financial Category",
    "Sex", "Birth Date",
]

HEADER_LABELS = ["Medical No.", "Time", "Slots", "Patient Name",
                  "Financial Cat.", "Sex", "Birth Date"]

_DATE_DDMMYYYY_RE = re.compile(r'^\d{2}/\d{2}/\d{4}$')
_IDNUM_RE = re.compile(r'^\d{10,15}$')
_NUMERIC_RE = re.compile(r'^-?\d+(\.\d+)?$')


def _row_values(ws, r, max_col):
    """Return a dict {col_index: value} of non-empty cells in row r."""
    out = {}
    for c in range(1, max_col + 1):
        v = ws.cell(row=r, column=c).value
        if v is not None and str(v).strip() != "":
            out[c] = v.strip() if isinstance(v, str) else v
    return out


def _label_value(cells, label, search_from=1):
    """
    Given a {col: value} row dict, find the column holding exactly
    `label` (or `label` with a trailing ':' stripped) at/after
    `search_from`, then return the next non-empty cell value to its
    right (survives merged cells shifting the value's own anchor
    column by 1-2 cells).
    """
    cols = sorted(c for c in cells if c >= search_from)
    for c in cols:
        val = str(cells[c]).strip().rstrip(":").strip()
        if val == label:
            for c2 in cols:
                if c2 > c:
                    return cells[c2]
            return None
    return None


def _as_number(v):
    """
    The real report stores ALL cell values as text, including totals
    (e.g. the 'Grand Total' cell holds the string '547', not the int
    547). Treat any int/float OR numeric-looking string as a number;
    checking only isinstance(v, (int, float)) silently finds nothing on
    the real file and turns the consistency checks into a no-op.
    """
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str) and _NUMERIC_RE.match(v.strip()):
        return float(v.strip()) if "." in v else int(v.strip())
    return None


def _nearby_value(cells, target_col, window=3):
    """Return the value at target_col, or the closest populated column
    within `window` columns either side (absorbs merged-cell offsets)."""
    if target_col is None:
        return None
    if target_col in cells:
        return cells[target_col]
    for d in range(1, window + 1):
        if target_col - d in cells:
            return cells[target_col - d]
        if target_col + d in cells:
            return cells[target_col + d]
    return None


# ═════════════════════════════════════════════════════════════════
# PRIMARY PARSER — "Clinic List Detail" (outpat_clnc_lst_det_j)
# Copied verbatim from the decree-lookup repo's queue_extractor.py v6b
# (only the docstring was corrected).
# ═════════════════════════════════════════════════════════════════

def parse_clinic_report(xlsx_bytes):
    """
    Parses the raw "Clinic List Detail" (outpat_clnc_lst_det_j) report
    bytes into a flat list of patient-record dicts (one per patient
    booking, keyed by CLEAN_HEADERS), plus a list of warning strings for
    any consistency check that failed (per-block "Total No. of Patient"
    and the report's "Grand Total").
    """

    import io
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active
    max_col = ws.max_column
    max_row = ws.max_row

    records = []
    warnings = []

    ctx = {
        "clinic": None, "date": None,
        "resource_id": None, "resource_name": None, "day": None,
        "doctor_id": None, "doctor_name": None,
    }

    grand_total_reported = None
    block_totals_reported = []
    block_totals_actual = []

    r = 1
    while r <= max_row:
        cells = _row_values(ws, r, max_col)
        if not cells:
            r += 1
            continue

        values_set = set(str(v).strip() for v in cells.values() if isinstance(v, str))

        # --- context lines -------------------------------------------------
        if "Clinic" in values_set:
            ctx["clinic"] = _label_value(cells, "Clinic")
            d = _label_value(cells, "Date")
            if d is not None:
                ctx["date"] = d
            r += 1
            continue

        if "Resource" in values_set:
            cols = sorted(cells)
            # Resource id is the first value after the "Resource" label;
            # Resource name is the next populated value after that.
            resource_id = _label_value(cells, "Resource")
            ctx["resource_id"] = resource_id
            # name = value after the id column
            id_col = None
            for c, v in cells.items():
                if v == resource_id:
                    id_col = c
                    break
            if id_col is not None:
                later = [c for c in cols if c > id_col]
                ctx["resource_name"] = cells[later[0]] if later else None
            day = _label_value(cells, "Day")
            if day is not None:
                ctx["day"] = day
            r += 1
            continue

        if "Doctor" in values_set:
            doctor_id = _label_value(cells, "Doctor")
            ctx["doctor_id"] = doctor_id
            cols = sorted(cells)
            id_col = None
            for c, v in cells.items():
                if v == doctor_id:
                    id_col = c
                    break
            if id_col is not None:
                later = [c for c in cols if c > id_col]
                ctx["doctor_name"] = cells[later[0]] if later else None
            r += 1
            continue

        # --- Grand Total footer ---------------------------------------------
        if "Grand Total" in values_set:
            nums = [n for n in (_as_number(v) for v in cells.values()) if n is not None]
            if nums:
                grand_total_reported = nums[0]
            r += 1
            continue

        # --- per-block patient total -----------------------------------------
        if "Total No. of Patient" in values_set:
            # the numeric total sits on the row ABOVE this label (seen in
            # the captured file: slots-total row, then the label row)
            prev_cells = _row_values(ws, r - 1, max_col)
            nums = [n for n in (_as_number(v) for v in prev_cells.values()) if n is not None]
            if nums:
                block_totals_reported.append((ctx["clinic"], ctx["date"], ctx["doctor_name"], nums[0]))
            r += 1
            continue

        # --- header row: "Medical No." + "Time" + "Patient Name" ... ---------
        if "Medical No." in values_set and "Patient Name" in values_set:
            header_cols = {}
            for c, v in cells.items():
                if isinstance(v, str) and v.strip() in HEADER_LABELS:
                    header_cols[v.strip()] = c

            # There is a sub-header row right below ("Financial Category")
            # that we skip.
            r += 2

            block_patient_count = 0
            while r <= max_row:
                pcells = _row_values(ws, r, max_col)
                pvals = set(str(v).strip() for v in pcells.values() if isinstance(v, str))

                if not pcells:
                    r += 1
                    continue
                if "Total No. of Patient" in pvals:
                    break  # let the outer loop handle the total-line
                if "Clinic" in pvals or "Grand Total" in pvals:
                    break  # malformed / unexpected — bail to outer loop

                # Detect a patient DATA row: must have a numeric Medical No.
                # and an H:MM-shaped Time. This also rejects the repeated
                # print-pagination header block that lands MID-BLOCK on a
                # multi-page report ("Clinic List Detail" / "Date :" /
                # "Time :" / "From ... To ..." / "Page X of 40") — without
                # this, the "Date :"/"Time :" lines' own values (e.g.
                # '27/07/2026', '02.44') fell inside the nearby-column
                # search window for Medical No./Time and got miscounted
                # as extra patients, inflating every block that happened
                # to span a page break.
                med_no_raw = _nearby_value(pcells, header_cols.get("Medical No.", 3))
                time_raw = _nearby_value(pcells, header_cols.get("Time", 6))
                patient_name = _nearby_value(pcells, header_cols.get("Patient Name", 20))

                med_no = str(med_no_raw).strip() if med_no_raw is not None else None
                time_val = str(time_raw).strip() if time_raw is not None else None

                looks_like_data_row = (
                    med_no is not None and re.match(r'^\d+$', med_no)
                    and time_val is not None and re.match(r'^\d{1,2}:\d{2}$', time_val)
                )

                if looks_like_data_row:
                    rec = {
                        "Date": ctx["date"],
                        "Day": ctx["day"],
                        "Clinic": ctx["clinic"],
                        "Resource ID": ctx["resource_id"],
                        "Resource Name": ctx["resource_name"],
                        "Doctor ID": ctx["doctor_id"],
                        "Doctor Name": ctx["doctor_name"],
                        "Medical No.": med_no,
                        "Time": time_val,
                        "Slots": _nearby_value(pcells, header_cols.get("Slots", 14)),
                        "Patient Name": patient_name,
                        "Financial Cat. Code": _nearby_value(pcells, header_cols.get("Financial Cat.", 27)),
                        "Financial Category": None,   # filled from the next row, below
                        "Sex": _nearby_value(pcells, header_cols.get("Sex", 32)),
                        "Birth Date": _nearby_value(pcells, header_cols.get("Birth Date", 39)),
                    }

                    # The row immediately below holds the full-text financial
                    # category, in the same column as Patient Name.
                    fincat_cells = _row_values(ws, r + 1, max_col)
                    name_col = header_cols.get("Patient Name", 20)
                    fincat_val = _nearby_value(fincat_cells, name_col)
                    # guard: don't grab it if it's actually the next block's label row
                    if fincat_val is not None and not set(
                        str(v).strip() for v in fincat_cells.values() if isinstance(v, str)
                    ) & {"Clinic", "Resource", "Doctor", "Total No. of Patient"}:
                        rec["Financial Category"] = fincat_val
                        r += 1  # consume the fin-category row too

                    records.append(rec)
                    block_patient_count += 1

                r += 1

            block_totals_actual.append((ctx["clinic"], ctx["date"], ctx["doctor_name"], block_patient_count))
            continue

        r += 1

    # --- consistency checks --------------------------------------------------
    for (reported, actual) in zip(block_totals_reported, block_totals_actual):
        r_clinic, r_date, r_doc, r_total = reported
        a_clinic, a_date, a_doc, a_total = actual
        if r_total != a_total:
            warnings.append(
                f"Block mismatch — Clinic={r_clinic!r} Date={r_date!r} "
                f"Doctor={r_doc!r}: report says 'Total No. of Patient' = "
                f"{r_total}, but parser extracted {a_total} rows."
            )

    if grand_total_reported is not None:
        parsed_total = len(records)
        if grand_total_reported != parsed_total:
            warnings.append(
                f"GRAND TOTAL mismatch: report's 'Grand Total' cell = "
                f"{grand_total_reported}, but parser extracted "
                f"{parsed_total} patient rows in total. Something was "
                f"missed or double-counted — do not trust this run until "
                f"resolved."
            )
    else:
        warnings.append("Could not find a 'Grand Total' cell to cross-check against.")

    return records, warnings


# ═════════════════════════════════════════════════════════════════
# FALLBACK PARSER — "Clinic List Detail - by Status"
# (outpat_clnc_lst_det_sts_j) — offset-based, only used when the
# primary report comes back empty. Same fixed columns as the decree
# repo's queue_parser.py.
# ═════════════════════════════════════════════════════════════════

def parse_status_report(xlsx_bytes):
    """
    FALLBACK PARSER (v10 logic, unchanged). Parses the raw "Clinic List
    Detail - by Status" (outpat_clnc_lst_det_sts_j) report into records shaped for the
    ORIGINAL 15-column CLEAN_HEADERS schema, because the rest of this
    script (and the downstream script that consumes the CLEAN file /
    the MR->Clinic map) is built around that shape.

    v11 NOTE — replaces a label-driven header-row detector that FAILED
    on the real server file (a genuine 91KB report with real data came
    back as "0 rows" because the header row never matched the expected
    label set exactly). This version instead ports the OFFSET-BASED
    extraction logic from the validated investigation script
    (parse_status_report() in the v7 script, confirmed against a real
    manual export, DMS5.har, and confirmed again by you as "correctly
    able to extract days queue data correctly from second report").
    Column POSITIONS are used directly (with a small nearby-column
    search to absorb merged-cell offsets) rather than re-deriving them
    from header text on every run — less theoretically elegant, but
    it's the version that has actually been proven against the real
    file, which beats an untested-in-production label search.

    Source layout: one continuous list per Clinic block (no per-
    Resource/per-Doctor sub-blocks, no per-block subtotal, no Grand
    Total footer). Each row carries its OWN appointment date
    (Appn_date) rather than one date per block.

    COLUMN MAPPING into the CLEAN_HEADERS shape:

      Date                 <- Appn_date (dd/mm/yyyy, found near col 25)
      Day                  <- computed from Appn_date
      Clinic               <- Clinic label
      Resource ID/Name       ALWAYS None -- not present in this report.
      Doctor ID/Name          ALWAYS None -- not present in this report.
      Medical No.          <- the short numeric column sitting between
                              ID Number and Patient Name (cols ~12-14).
                              There is NO label on the sheet confirming
                              this column's identity, but it has now
                              been CONFIRMED against real patient
                              records (values like '6891', '9334',
                              '6040' -- matching the MR format from the
                              original "Clinic List Detail" report
                              exactly) to be the correct MR code.
                              CORRECTION from an earlier version of
                              this parser: the report's own labeled
                              "Old medical" column was tried first
                              (since it's the one field with an actual
                              header label), but real data proved that
                              column holds a case/file number instead
                              (values like '1132/2022' -- clearly not
                              an MR code). It's kept under
                              _source_extra for reference but is no
                              longer used for MR lookups.
      Time / Slots            ALWAYS None -- not present in this report.
      Patient Name          <- Patient Name column, when populated (it
                              came back blank on every row of the
                              sample export used to validate this
                              parser -- that's expected, not a bug).
      Financial Cat./Sex/
      Birth Date               ALWAYS None -- not present in this report.

    Also returned in each record under "_source_extra": "Serial No.",
    "ID Number", "Appoint Number", "User", "Status", and the
    positionally-inferred Medical No. candidate described above.

    There is no per-block or grand-total figure printed anywhere on
    this report to cross-check the row count against -- the row count
    itself is the only consistency signal available.
    """
    import io
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active
    max_col = ws.max_column
    max_row = ws.max_row

    records = []
    warnings = []
    current_clinic = None

    r = 1
    while r <= max_row:
        cells = _row_values(ws, r, max_col)
        if not cells:
            r += 1
            continue

        values_set = set(str(v).strip() for v in cells.values() if isinstance(v, str))

        if "Clinic" in values_set:
            current_clinic = _label_value(cells, "Clinic")
            r += 1
            continue

        # Skip page/report boilerplate: title, "Date :"/"Time :" stamp,
        # the "From ... To ..." range header, "Page X of N", and the
        # column-header block ("Serial No." / "ID Number" / etc.) that
        # repeats at the top of every printed page.
        if values_set & {"From", "Date :", "Time :", "Serial No.", "ID Number",
                          "Patient Name", "Appoint Number", "Appn_date",
                          "Old medical", "Status", "User"}:
            r += 1
            continue
        if any(str(v).strip().startswith("Page ") for v in cells.values() if isinstance(v, str)):
            r += 1
            continue
        if any("Clinic List Detail" in str(v) for v in cells.values() if isinstance(v, str)):
            r += 1
            continue

        # --- data row detection ---------------------------------------
        # A real row has a dd/mm/yyyy Appn_date (found near col 25) AND a
        # 10-15 digit ID Number (found near col 8). Both are required --
        # this is what excludes the "From ... To ..." range header row,
        # which has a lone date but no ID Number nearby.
        date_val = _nearby_value(cells, 25, window=4)
        date_val = date_val if isinstance(date_val, str) and _DATE_DDMMYYYY_RE.match(date_val) else None

        id_number = None
        for c in range(6, 13):
            v = cells.get(c)
            if isinstance(v, str) and _IDNUM_RE.match(v):
                id_number = v
                break

        if date_val is None or id_number is None:
            r += 1
            continue

        med_no_unconfirmed = cells.get(13) or cells.get(12) or cells.get(14)
        appoint_no = _nearby_value(cells, 21, window=3)
        old_medical = cells.get(35) or cells.get(37) or cells.get(36)
        user = cells.get(29) or cells.get(30)

        # Patient Name / Status: use a NARROW window and explicitly refuse
        # to fall back onto a column already claimed by another field --
        # both came back blank in every row of the sample export; with a
        # wide fallback window this would "steal" the Medical No. or User
        # value from a nearby column and silently duplicate it into the
        # wrong field whenever the real cell is empty, which is worse
        # than just leaving it blank.
        claimed = {13, 12, 14, 21, 22, 23, 20, 24}
        patient_name = None
        for d in (0, 1, 2, -1, 3):
            c = 16 + d
            if c in cells and c not in claimed:
                patient_name = cells[c]
                break
        claimed = {29, 30}
        status = None
        for d in (0, 1, 2, -1):
            c = 32 + d
            if c in cells and c not in claimed:
                status = cells[c]
                break

        appn_date_dt = datetime.strptime(date_val, "%d/%m/%Y")
        rec = {
            "Date": date_val,
            "Day": appn_date_dt.strftime("%A"),
            "Clinic": current_clinic,
            "Resource ID": None,
            "Resource Name": None,
            "Doctor ID": None,
            "Doctor Name": None,
            "Medical No.": med_no_unconfirmed,
            "Time": None,
            "Slots": None,
            "Patient Name": patient_name,
            "Financial Cat. Code": None,
            "Financial Category": None,
            "Sex": None,
            "Birth Date": None,
            "_source_extra": {
                "Serial No.": cells.get(4) or cells.get(3),
                "ID Number": id_number,
                "Appoint Number": appoint_no,
                "User": user,
                "Status": status,
                "Old Medical No. (NOT an MR code -- confirmed a case/file "
                "number like '1132/2022', kept only for reference)": old_medical,
            },
        }
        records.append(rec)
        r += 1

    if not records:
        warnings.append(
            "0 rows parsed from the status report for this range -- either "
            "it genuinely came back empty, or the report layout has "
            "changed since this parser was written."
        )

    warnings.append(
        "Resource ID/Name, Doctor ID/Name, Time, Slots, Financial Cat. "
        "Code/Category, Sex, and Birth Date are always blank in this "
        "output -- the status report this data now comes from does not "
        "carry that information at all (see parse_status_report()'s "
        "docstring)."
    )
    warnings.append(
        "'Medical No.' is populated from the short numeric column that "
        "sits between ID Number and Patient Name (no label on the sheet "
        "confirms this, but it's now CONFIRMED against real patient "
        "records to be the correct MR code format, e.g. '6891', '9334'). "
        "The report's own labeled 'Old medical' column is NOT an MR code "
        "-- it's a case/file number like '1132/2022' -- and is kept only "
        "under _source_extra for reference, not used for MR lookups."
    )

    return records, warnings


def parse_report(report_code, xlsx_bytes):
    """Pick the parser that matches the report the bytes came from."""
    if report_code == STATUS_REPORT_CODE:
        return parse_status_report(xlsx_bytes)
    return parse_clinic_report(xlsx_bytes)


def write_clean_excel(records, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Clean Queue Data"
    ws.append(CLEAN_HEADERS)
    for rec in records:
        ws.append([rec.get(h) for h in CLEAN_HEADERS])

    # basic readability: bold header, freeze it, autosize-ish column widths
    for c in range(1, len(CLEAN_HEADERS) + 1):
        ws.cell(row=1, column=c).font = ws.cell(row=1, column=c).font.copy(bold=True)
    ws.freeze_panes = "A2"
    for c, header in enumerate(CLEAN_HEADERS, start=1):
        max_len = len(str(header))
        for rec in records:
            v = rec.get(header)
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(c)].width = min(max_len + 2, 45)

    wb.save(out_path)


# ═════════════════════════════════════════════════════════════════
# ORCHESTRATION  (shared by the public API and the standalone main)
# ═════════════════════════════════════════════════════════════════

def to_slash_date(ddmmyyyy_dash):
    d = datetime.strptime(ddmmyyyy_dash, "%d-%m-%Y")
    return d.strftime("%d/%m/%Y")


def _warn(msg):
    """Print a warning; also emit a GitHub Actions annotation when in CI."""
    print(f"   !! {msg}")
    if os.environ.get("GITHUB_ACTIONS") == "true":
        print("::warning::" + str(msg).replace("\r", " ").replace("\n", " "))


def _fetch_verify_parse(report_code, date_from_dash, date_to_dash, out_dir,
                        resource_id="", resource_name=""):
    """
    One full pass for one report: fetch -> save raw -> verify the server's
    own From/To header -> parse -> save CLEAN. Raises on any real failure
    (the raw file is still saved first, so it can be inspected).
    Returns (records, warnings).
    """
    date_from = to_slash_date(date_from_dash)
    date_to = to_slash_date(date_to_dash)

    print(f"   -> Fetching {report_code} for {date_from} -> {date_to}")
    filename, content = fetch_report(
        requests.Session(), report_code, date_from, date_to,
        resource_id=resource_id, resource_name=resource_name,
    )

    os.makedirs(out_dir, exist_ok=True)
    raw_path = os.path.join(out_dir, f"{date_from_dash}_to_{date_to_dash}_{filename}")
    with open(raw_path, "wb") as f:
        f.write(content)
    print(f"   -> Saved raw report ({len(content):,} bytes) -> {raw_path}")

    actual_from, actual_to = verify_report_date_range(content, date_from, date_to)
    print(f"   OK: report header confirms From={actual_from} To={actual_to}")

    records, warnings = parse_report(report_code, content)

    clean_name = f"{date_from_dash}_to_{date_to_dash}_{os.path.splitext(filename)[0]}_CLEAN.xlsx"
    clean_path = os.path.join(out_dir, clean_name)
    write_clean_excel(records, clean_path)
    print(f"   -> Saved clean table ({len(records)} rows) -> {clean_path}")
    return records, warnings


def extract_queue(date_from_dash, date_to_dash, output_dir=None,
                  resource_id="", resource_name="", allow_fallback=None):
    """
    Pull the queue for a date range (dd-mm-yyyy, inclusive) and return
    (records, warnings, report_code_used).

    Uses REPORT_CODE first (same report as the decree repo). If that
    parses to zero rows and the fallback is enabled, the status report
    is tried once. Raises on real failures of the PRIMARY report; a
    failure of the fallback only produces a warning (the primary
    report's empty result is then returned).
    """
    out_dir = output_dir if output_dir is not None else OUTPUT_DIR
    if allow_fallback is None:
        allow_fallback = ENABLE_STATUS_FALLBACK

    records, warnings = _fetch_verify_parse(
        REPORT_CODE, date_from_dash, date_to_dash, out_dir,
        resource_id=resource_id, resource_name=resource_name)
    used = REPORT_CODE

    if not records and allow_fallback:
        _warn(f"{REPORT_CODE} returned 0 rows for {date_from_dash} -> "
              f"{date_to_dash}; trying {STATUS_REPORT_CODE} once before "
              f"treating this as an empty day.")
        try:
            f_records, f_warnings = _fetch_verify_parse(
                STATUS_REPORT_CODE, date_from_dash, date_to_dash, out_dir,
                resource_id=resource_id, resource_name=resource_name)
        except Exception as e:
            _warn(f"Fallback {STATUS_REPORT_CODE} failed ({e}); "
                  f"keeping the empty primary result.")
        else:
            if f_records:
                _warn(f"Using {STATUS_REPORT_CODE} data ({len(f_records)} rows) — "
                      f"the primary report was empty.")
                records, warnings, used = f_records, f_warnings, STATUS_REPORT_CODE
            else:
                warnings = list(warnings) + [
                    f"{STATUS_REPORT_CODE} was also empty for this range."]

    return records, warnings, used


# ═════════════════════════════════════════════════════════════════
# PUBLIC MODULE API — what Extract_DMS_Patients_Prescriptioned_
# Services_Data_modified.py actually imports and calls
# ═════════════════════════════════════════════════════════════════

def get_queue_mr_clinic_map(run_date_ddmmyyyy, output_dir=None,
                             resource_id="", resource_name=""):
    """
    THE FUNCTION THE CALLER SCRIPT IMPORTS. Call site (Extract_DMS_
    Patients_Prescriptioned_Services_Data_modified.py):

        raw_map = queue_mr_extractor.get_queue_mr_clinic_map(
            RUN_DATE, output_dir=QUEUE_OUTPUT_DIR
        )

    RUN_DATE is a single day as "dd-mm-yyyy"; QUEUE_OUTPUT_DIR is where
    the raw + CLEAN audit files are saved.

    Returns a plain {MR: Clinic} dict.

    Error contract (matches the caller's own handling):
      * Raises an exception (never sys.exit) on any real failure —
        network error, date-range mismatch, missing header row, parse
        error — because the caller wraps this in `except Exception`
        and SystemExit would slip past it.
      * Returns an EMPTY dict (no exception) when the day genuinely has
        zero queued patients; the caller checks that itself.

    A patient can legitimately appear more than once in a day's queue
    (e.g. at more than one clinic). The FIRST clinic seen for that MR is
    kept and every conflict is printed, rather than silently picking one.
    """
    out_dir = output_dir if output_dir is not None else OUTPUT_DIR

    records, warnings, used = extract_queue(
        run_date_ddmmyyyy, run_date_ddmmyyyy, output_dir=out_dir,
        resource_id=resource_id, resource_name=resource_name)

    for w in warnings:
        _warn(w)

    mr_clinic_map = {}
    conflicts = []
    for rec in records:
        mr = rec.get("Medical No.")
        clinic = rec.get("Clinic")
        if not mr:
            continue
        mr = str(mr).strip()
        if not mr:
            continue
        if mr in mr_clinic_map:
            if mr_clinic_map[mr] != clinic:
                conflicts.append((mr, mr_clinic_map[mr], clinic))
            continue  # keep the first clinic seen for this MR
        mr_clinic_map[mr] = clinic

    if conflicts:
        print(f"   !! {len(conflicts)} MR code(s) appeared under more than "
              f"one clinic on {run_date_ddmmyyyy} — kept the first clinic "
              f"seen for each:")
        for mr, kept, dropped in conflicts[:10]:
            print(f"      MR {mr}: kept {kept!r}, also saw {dropped!r}")
        if len(conflicts) > 10:
            print(f"      ... and {len(conflicts) - 10} more")

    print(f"   -> {len(records)} queue rows -> {len(mr_clinic_map)} unique MR "
          f"code(s) (source report: {used})")
    return mr_clinic_map


# ═════════════════════════════════════════════════════════════════
# MAIN  (standalone use — pulls a date or date RANGE into raw + CLEAN
# files; get_queue_mr_clinic_map() above is the entry point used when
# this module is imported by another script)
# ═════════════════════════════════════════════════════════════════

def main():
    # sys.argv is read here, NOT at import time: the prescription script
    # imports this module and uses its own sys.argv[1] as RUN_DATE.
    today = datetime.now().strftime("%d-%m-%Y")
    date_from = sys.argv[1] if len(sys.argv) > 1 else today
    date_to = sys.argv[2] if len(sys.argv) > 2 else date_from

    print("=" * 70)
    print("  queue_mr_extractor — standalone run (v11, decree-repo aligned)")
    print(f"  Target      : {WEBREPORT_BASE}")
    print(f"  Report code : {REPORT_CODE}  (fallback: "
          f"{STATUS_REPORT_CODE if ENABLE_STATUS_FALLBACK else 'disabled'})")
    print(f"  Date range  : {date_from} -> {date_to}")
    print("=" * 70)

    try:
        records, warnings, used = extract_queue(
            date_from, date_to, output_dir=OUTPUT_DIR,
            resource_id=RESOURCE_ID, resource_name=RESOURCE_NAME)
    except Exception as e:
        print(f"\nXX Extraction failed: {e}")
        sys.exit(1)

    print(f"\n-- {len(records)} patient rows from {used}")
    if warnings:
        print("\n!! NOTES / WARNINGS — review before trusting this run:")
        for w in warnings:
            print("   -", w)
    else:
        print("\n-- Consistency check passed: parsed row count matches the")
        print("   report's own per-block and Grand Total figures.")
    print("\nDone.")


if __name__ == "__main__":
    main()
