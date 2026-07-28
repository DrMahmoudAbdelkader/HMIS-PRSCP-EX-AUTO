"""
queue_mr_extractor  —  v10  (IMPORTABLE MODULE for the coupled pipeline)
====================================================================
PUBLIC API
-----------
    queue_mr_extractor.get_queue_mr_clinic_map(run_date_ddmmyyyy,
                                                output_dir=None)
        -> {MR_code: Clinic_name, ...}

This is what Extract_DMS_Patients_Prescriptioned_Services_Data_
modified.py imports and calls at startup to get today's queued MR
codes and which clinic each one belongs to. See that function's own
docstring below for the exact contract (return shape, error handling,
duplicate-MR behavior) — it was written to match that caller's actual
call site and its `except Exception` / `isinstance(raw_map, dict)`
handling exactly.

Everything else in this file (fetch_report, sync_date_field,
verify_report_date_range, parse_clinic_report, write_clean_excel, and
the standalone main()) is the same queue-extraction pipeline from the
prior standalone script, unchanged in behavior — this revision's only
job is to expose it as an importable function with the right name and
contract, instead of only running as `python script.py`.

WHY THE STATUS REPORT (outpat_clnc_lst_det_sts_j)
----------------------------------------------------
The original report code, outpat_clnc_lst_det_j ("Clinic List
Detail"), drops data constantly — several single-day requests in the
20-28 Jul 2026 range came back completely EMPTY from that report while
the same days returned real data from outpat_clnc_lst_det_sts_j
("Clinic List Detail — by Status"), which this module uses instead.

*** READ THIS BEFORE YOU RELY ON THE OUTPUT ***
This report's sheet does NOT expose a column literally labelled
"Medical No." Two candidate columns were tried:
  - "Old medical" — the one field that actually has a real label
    printed on the sheet, tried FIRST for that reason in an earlier
    version of this module.
  - The short, unlabeled numeric column sitting between ID Number and
    Patient Name (no label confirms it, but it's positioned exactly
    where the original "Clinic List Detail" report's own Medical No.
    column sits).
Checked against real patient records, "Old medical" turned out to
hold a case/file number (e.g. "1132/2022") — NOT an MR code — while
the unlabeled positional column holds real MR-shaped values (e.g.
"6891", "9334", "6040"). This module now uses that positional column
as "Medical No." / the MR key in get_queue_mr_clinic_map()'s returned
dict; "Old medical" is kept per-row under _source_extra for reference
only and is never used for MR lookups.

This report also does NOT contain a Resource/Doctor/Time/Slots/
Financial/Sex/Birth-Date breakdown at all. Those columns are still
written to the CLEAN audit file for schema compatibility, but they
are always blank (None) — there is no server-side data to put in them.

COLUMN DETECTION uses validated column POSITIONS (with a small nearby-
column search to absorb merged-cell offsets), not a header-text search.
An earlier version of this module tried locating columns by matching
the sheet's own header-row text ("Serial No.", "ID Number", etc.)
exactly — that failed in production (a real, non-empty 91KB report
came back as "0 rows" because the header row never matched the
expected text precisely). The offset-based approach below is the one
that has actually been proven against real exports. Known header
labels are still used, but only to recognize and skip page-break
boilerplate rows (the repeated "Date :"/"Time :"/"Page X of N" block
that appears mid-file on multi-page exports), not to derive positions.

ALSO FIXED IN THIS REVISION: the AJAX "dateSelect" replay
(sync_date_field()) was silently failing to pick up the server's
refreshed session token on every call, because the response's actual
XML update id is "j_id1:javax.faces.ViewState:0" (a form-prefixed,
indexed id), not the bare "javax.faces.ViewState" the regex was
looking for. This didn't surface as a wrong date range in testing only
because the date being requested happened to already be the field's
default — it would have broken for any other date. Fixed to match the
real format.

WHAT'S UNCHANGED FROM THE PRIOR WORKING VERSION
---------------------------------------------------
- sync_date_field(): the AJAX "dateSelect" replay. The date fields are
  PrimeFaces/JSF calendar widgets whose typed value is client-side
  only until a partial-AJAX dateSelect event binds it to the server's
  ViewState — this logic is untouched.
- verify_report_date_range(): still reads the server's own printed
  "From ... To ..." header and RAISES (raw file still saved) on a
  mismatch, rather than silently saving data for the wrong range.

Output (both when imported and when run standalone)
-------------------------------------------------------
  1. <daterange>_<report>.xlsx        <- the raw file exactly as the
                                          server generated it
  2. <daterange>_..._CLEAN.xlsx       <- the flat table in the
                                          15-column schema (several
                                          columns always blank — see
                                          the warning above)

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
# CONFIGURATION – edit these before running
# ═════════════════════════════════════════════════════════════════

HOST            = "41.33.24.254:8080"
WEBREPORT_BASE  = f"http://{HOST}/WebReport-JWEB"

OLD_REPORT_CODE = "outpat_clnc_lst_det_j"       # "Clinic List Detail" (kept only as a reference — drops days, see docstring)
REPORT_CODE     = "outpat_clnc_lst_det_sts_j"   # "Clinic List Detail - by Status" (v9 data source — confirmed to catch full-range data)
LANG            = "L"
HSCD            = "01"                          # hospital/branch code

# ---- DATE RANGE TO EXTRACT ----
DATE_FROM = "20-07-2026"   # dd-mm-yyyy, inclusive
DATE_TO   = "28-07-2026"   # dd-mm-yyyy, inclusive

# ---- Optional: restrict to one physician/resource. Leave both blank for ALL. ----
RESOURCE_ID   = ""
RESOURCE_NAME = ""

TIMEOUT = 60
# Overridable via env so this same file runs unchanged locally on Windows
# (defaults to your D: drive) and on a Linux CI runner (set QUEUE_OUTPUT_DIR
# there to something like /tmp/queue_dms_data).
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
# CLEAN-TABLE PARSER  (v9 — label-driven, status report)
# ═════════════════════════════════════════════════════════════════

CLEAN_HEADERS = [
    "Date", "Day", "Clinic",
    "Resource ID", "Resource Name",
    "Doctor ID", "Doctor Name",
    "Medical No.", "Time", "Slots", "Patient Name",
    "Financial Cat. Code", "Financial Category",
    "Sex", "Birth Date",
]

# Header labels actually printed on the status report's sheet, confirmed
# against a real manual export. This is the ENTIRE contract the parser
# relies on — if the server ever renames/reorders these, the header-row
# detector below will fail loudly (0 rows + a warning) rather than
# silently reading the wrong column, because it locates every field by
# this exact text rather than a fixed column number.
# NOTE: the header labels below ("Serial No.", "ID Number", etc.) ARE
# real text printed on the sheet (confirmed against a manual export) —
# they're still used to recognize and skip page-break boilerplate rows
# in parse_clinic_report() — but they are NOT used to derive column
# positions anymore (that label-driven approach failed against the real
# server file; see parse_clinic_report()'s docstring for what replaced it).

_DATE_DDMMYYYY_RE = re.compile(r'^\d{2}/\d{2}/\d{4}$')
_IDNUM_RE = re.compile(r'^\d{10,15}$')


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


def parse_clinic_report(xlsx_bytes):
    """
    Parses the raw "Clinic List Detail - by Status"
    (outpat_clnc_lst_det_sts_j) report into records shaped for the
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
        "carry that information at all (see parse_clinic_report()'s "
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
# PUBLIC MODULE API — what Extract_DMS_Patients_Prescriptioned_
# Services_Data_modified.py actually imports and calls
# ═════════════════════════════════════════════════════════════════

def get_queue_mr_clinic_map(run_date_ddmmyyyy, output_dir=None,
                             resource_id="", resource_name=""):
    """
    THE FUNCTION THE CALLER SCRIPT IMPORTS. Confirmed against its own
    call site (Extract_DMS_Patients_Prescriptioned_Services_Data_
    modified.py, line ~1518):

        raw_map = queue_mr_extractor.get_queue_mr_clinic_map(
            RUN_DATE, output_dir=QUEUE_OUTPUT_DIR
        )

    where RUN_DATE is a single day as "dd-mm-yyyy" (e.g. "28-07-2026",
    from sys.argv[1] or datetime.now()), and QUEUE_OUTPUT_DIR is
    "D:\\Queue_DMS_Data" — both formats match this module's own
    DATE_FROM/DATE_TO and OUTPUT_DIR conventions already, so no
    conversion is needed on the caller's side.

    Pulls the queue/status report for that ONE day (From == To ==
    run_date) and returns a plain {MR: Clinic} dict — the caller
    explicitly also tolerates a list of {"MR":..., "Clinic":...}
    dicts, but a dict is what's actually produced here, since it's the
    simpler, unambiguous shape and matches the "MR -> Clinic mapping"
    the caller's own docstring describes.

    IMPORTANT — matches the caller's own error-handling contract:
      * Raises an exception (does NOT call sys.exit) on any real
        failure — network error, date-range mismatch, missing header
        row, etc. — because the caller wraps this call in
        `except Exception as e: print(...); sys.exit(1)`. If this
        function called sys.exit() itself, that raises SystemExit,
        which is NOT a subclass of Exception, so it would blow past
        the caller's except block entirely and kill the whole pipeline
        with a raw traceback instead of the caller's friendly message.
      * Returns an EMPTY dict (not an exception) when the day
        genuinely has zero queued patients — the caller already checks
        `if not mr_codes: print(...); sys.exit(1)` itself, so this
        function shouldn't pre-empt that with an exception for a
        perfectly normal "quiet day" case.

    As a side effect (for audit trail / manual troubleshooting, same
    as running this module directly), the raw report and a CLEAN
    table for the day are saved into `output_dir` using this module's
    usual naming convention.

    A patient can legitimately appear more than once in a single day's
    queue (e.g. queued at more than one clinic). When that happens,
    the FIRST clinic seen for that MR (in report order) is kept, and
    every such MR is listed in a printed warning — silently picking
    one without saying so would quietly misfile whichever record uses
    that MR later, exactly the kind of silent mismatch this whole
    project has been trying to eliminate.
    """
    date_slash = to_slash_date(run_date_ddmmyyyy)
    out_dir = output_dir if output_dir is not None else OUTPUT_DIR

    session = requests.Session()

    print(f"   -> Fetching {REPORT_CODE} for {date_slash} (single day)")
    filename, content = fetch_report(
        session, REPORT_CODE, date_slash, date_slash,
        resource_id=resource_id, resource_name=resource_name,
    )

    os.makedirs(out_dir, exist_ok=True)
    raw_path = os.path.join(out_dir, f"{run_date_ddmmyyyy}_to_{run_date_ddmmyyyy}_{filename}")
    with open(raw_path, "wb") as f:
        f.write(content)

    # Raises RuntimeError on mismatch — deliberately NOT caught here, so
    # it propagates to the caller's own except-block as documented above.
    verify_report_date_range(content, date_slash, date_slash)

    records, warnings = parse_clinic_report(content)
    for w in warnings:
        print(f"   !! {w}")

    clean_path = os.path.join(
        out_dir, f"{run_date_ddmmyyyy}_to_{run_date_ddmmyyyy}_{OLD_REPORT_CODE}_CLEAN.xlsx"
    )
    write_clean_excel(records, clean_path)
    print(f"   -> Saved raw ({len(content):,} bytes) and clean "
          f"({len(records)} rows) queue files to {out_dir}")

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
              f"one clinic on {date_slash} — kept the first clinic seen "
              f"for each, in case that's not what you want:")
        for mr, kept, dropped in conflicts[:10]:
            print(f"      MR {mr}: kept {kept!r}, also saw {dropped!r}")
        if len(conflicts) > 10:
            print(f"      ... and {len(conflicts) - 10} more")

    return mr_clinic_map


# ═════════════════════════════════════════════════════════════════
# MAIN  (standalone use — pulls a date RANGE into raw+CLEAN files;
# get_queue_mr_clinic_map() above is the entry point used when this
# module is imported by another script)
# ═════════════════════════════════════════════════════════════════

def to_slash_date(ddmmyyyy_dash):
    d = datetime.strptime(ddmmyyyy_dash, "%d-%m-%Y")
    return d.strftime("%d/%m/%Y")


def main():
    print("=" * 70)
    print("  queue_mr_extractor — standalone run (status-report + clean-table, v10)")
    print(f"  Target      : {WEBREPORT_BASE}")
    print(f"  Report code : {REPORT_CODE}  (data source; was {OLD_REPORT_CODE})")
    print(f"  Date range  : {DATE_FROM} -> {DATE_TO}")
    print("=" * 70)

    date_from = to_slash_date(DATE_FROM)
    date_to   = to_slash_date(DATE_TO)

    session = requests.Session()

    try:
        filename, content = fetch_report(
            session, REPORT_CODE, date_from, date_to,
            resource_id=RESOURCE_ID, resource_name=RESOURCE_NAME,
        )
    except Exception as e:
        print(f"\nXX Extraction failed: {e}")
        sys.exit(1)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    raw_name = f"{DATE_FROM}_to_{DATE_TO}_{filename}"
    raw_path = os.path.join(OUTPUT_DIR, raw_name)
    with open(raw_path, "wb") as f:
        f.write(content)
    print(f"\n-- Saved raw report   ({len(content):,} bytes) -> {raw_path}")

    print("\n-- Verifying the server actually used the requested date range...")
    try:
        actual_from, actual_to = verify_report_date_range(content, date_from, date_to)
        print(f"   OK: report header confirms From={actual_from} To={actual_to}")
    except Exception as e:
        print(f"\nXX {e}")
        sys.exit(1)

    print("\n-- Parsing into a clean flat table...")
    try:
        records, warnings = parse_clinic_report(content)
    except Exception as e:
        print(f"XX Parsing failed: {e}")
        print("   The raw file was still saved above — you can inspect it")
        print("   manually, or send it back to fix the parser.")
        sys.exit(1)

    # NOTE: deliberately using OLD_REPORT_CODE here, NOT the actual server
    # filename (which is now outpat_clnc_lst_det_sts_j.xlsx). This keeps
    # the CLEAN filename pattern identical to what the original script
    # produced, since a downstream script may look for files by that
    # exact naming pattern. The RAW file above still uses the server's
    # real filename, so you can always tell the two apart.
    clean_name = f"{DATE_FROM}_to_{DATE_TO}_{OLD_REPORT_CODE}_CLEAN.xlsx"
    clean_path = os.path.join(OUTPUT_DIR, clean_name)
    write_clean_excel(records, clean_path)

    print(f"-- Saved clean table  ({len(records)} patient rows) -> {clean_path}")

    if warnings:
        print("\n!! NOTES / WARNINGS — review before trusting this run:")
        for w in warnings:
            print("   -", w)
    else:
        print("\n-- No warnings for this run.")

    print("\nDone.")


if __name__ == "__main__":
    main()
