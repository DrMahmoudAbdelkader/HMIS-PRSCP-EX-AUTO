"""
CMIS MRM Patient Clinical Data Extractor  —  COUPLED daily pipeline
======================================================================
CHANGES IN THIS VERSION (bulk / parallel pipeline)
--------------------------------------------------
* NEW  --chunk chunk_<n>.json --chunk-out result.json
       Used by the parallel GitHub jobs. Each patient carries ALL the dates
       it was queued on ({"mr","national_id","dates":{date: clinic}}); the
       patient page is opened ONCE and every investigation / medical report /
       MDT dated on ANY of those days is kept, tagged with that day and the
       clinic of that day.
* No queue pull in chunk mode - plan_chunks.py already did it.
* Old single-day mode still works:  python script.py [dd-mm-yyyy]

CHANGES IN THIS VERSION (coupling with queue_mr_extractor.py)
-----------------------------------------------------------------
1. MR codes are no longer read from a hand-prepared Excel file. This
   script now calls queue_mr_extractor.get_queue_records() (v12: a thin
   adapter over the decree repo's queue_extractor.py + queue_parser.py,
   which must sit next to this script) at
   startup, which pulls TODAY's (or a given RUN_DATE's) queue report
   and returns MR -> Clinic mapping directly (as either a dict, or a
   list of {"MR":..., "Clinic":...} dicts — both are handled, see
   step 1 in main()). That mapping is:
     - the source of which MRs to process, and
     - carried through to every output row as a "Clinic" column, so
       the output is categorized by which clinic each patient was
       served at (exactly what the queue script saw them queued
       under that day).

2. "Last visit data only, dated today": Investigations, Medical
   Reports (033) and MDT forms (02) are now filtered to only the
   records whose own date equals RUN_DATE. Fetch is still pulled
   fresh per patient (that's how the source system works), but this
   script performs the additional date filtering so the output only
   ever contains what actually happened at/around today's visit,
   not the patient's whole history.

3. Attending Notes extraction has been REMOVED entirely (the
   AttendingNote/Index API call, its parser, and the "Attending
   Notes" output sheet) — not needed for this pipeline.

4. ACCUMULATING OUTPUT: instead of writing one fresh, isolated
   workbook per run, this script now maintains a single cumulative
   workbook (MASTER_OUTPUT_PATH). Each run's newly-extracted rows are
   appended to whatever is already saved there (with a dedup key so
   re-running the same day twice doesn't create duplicate rows), so
   the file grows day over day.

5. SUPABASE PUSH: new rows from each run are also upserted into four
   Supabase tables (patient_summary, investigations, medical_reports,
   mdt_forms), de-duplicated by an md5 row_hash. Credentials are read
   ONLY from environment variables (SUPABASE_URL / SUPABASE_SERVICE_KEY)
   — set them in run_extractor.bat, never hardcode them here. Excel
   stays as the local backup/audit trail regardless of whether the
   Supabase push succeeds.

For each patient (MR) this script still extracts:
  • Patient basic info (from Admission API)
  • Investigations (lab/radiology orders) dated RUN_DATE
  • Medical Reports (SheetCode 033) dated RUN_DATE
  • MDT forms      (SheetCode 02)  dated RUN_DATE

REQUIREMENTS
------------
  pip install requests openpyxl beautifulsoup4 lxml supabase

RUN
---
  python Extract_DMS_Patients_Prescriptioned_Services_Data.py
  python Extract_DMS_Patients_Prescriptioned_Services_Data.py 27-07-2026
      (optional CLI arg: RUN_DATE in dd-mm-yyyy — defaults to today)

  In production, run this via run_extractor.bat, which sets
  SUPABASE_URL / SUPABASE_SERVICE_KEY as environment variables before
  calling python — see that file for the credentials.
"""

import requests
import json
import time
import re
import sys
import os
import traceback
import html as html_lib
import hashlib
from urllib.parse import urlencode
from datetime import datetime
from bs4 import BeautifulSoup, Tag

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("❌  Missing dependency.  Run:  pip install requests openpyxl beautifulsoup4 lxml")
    sys.exit(1)

try:
    from supabase import create_client
except ImportError:
    create_client = None  # handled gracefully wherever Supabase is used

# Coupled queue module — lives alongside this script.
import queue_mr_extractor


# ═══════════════════════════════════════════════════════════════════
# CONFIGURATION  —  Edit these before running
# ═══════════════════════════════════════════════════════════════════

USERNAME       = os.environ.get("HMIS_USERNAME", "557")   # CMIS login user
PASSWORD       = os.environ.get("HMIS_PASSWORD", "557")   # CMIS login password
CMIS_BASE      = "http://41.33.24.253/CMIS/MRM_5.2"
SHEET_API_BASE = "http://41.33.24.253/GeneralSheetManagerApi/api"
COMMON_API     = "http://41.33.24.253/CMIS/Common_API/api"
HOSPITAL_CODE  = "01"           # Default hospital code

# The day this run represents (dd-mm-yyyy). Defaults to today; pass a
# date as the first CLI argument to backfill/re-run a specific day.
# This is used BOTH as the date the queue is pulled for, AND as the
# filter applied to investigations/medical reports/MDT (only records
# dated this same day are kept — "today's visit only").
# Default = today. The CLI (bottom of this file) overrides it for single-day
# runs. NOT read from sys.argv at import time: merge_outputs.py imports this
# module and has its own arguments.
RUN_DATE = datetime.now().strftime("%d-%m-%Y")

# Where the queue script's own raw/clean/simplified audit files land.
QUEUE_OUTPUT_DIR = os.environ.get("QUEUE_OUTPUT_DIR", r"D:\Queue_DMS_Data")

# The cumulative output this script grows day over day — kept as a
# local Excel workbook regardless of Supabase, as an audit trail.
MASTER_OUTPUT_PATH = os.environ.get(
    "MASTER_OUTPUT_PATH", r"D:\Queue_DMS_Data\Cumulative_Patient_Clinical_Data.xlsx"
)

# Set True to push new rows to Supabase each run (in addition to the
# Excel accumulator). If credentials are missing or the connection
# test fails, the script automatically falls back to Excel-only for
# that run rather than crashing.
PUSH_TO_SUPABASE = True

# ── Supabase credentials — READ FROM ENVIRONMENT ONLY ───────────────
# Never hardcode these. Set them in run_extractor.bat (or your
# scheduler's environment) before this script runs:
#   set SUPABASE_URL=https://your-project-id.supabase.co
#   set SUPABASE_SERVICE_KEY=your-service-role-key
SUPABASE_URL         = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")

# Maps this script's internal data-section keys to the Supabase table
# names created by the SQL setup script.
_SUPABASE_TABLE_NAMES = {
    "summary":          "patient_summary",
    "investigations":   "investigations",
    "medical_reports":  "medical_reports",
    "mdt":              "mdt_forms",
}

# Snake_case DB column names, in the SAME ORDER as each row tuple is
# built in main() / _SHEET_COLUMNS below. Order matters — these are
# zipped positionally against each row.
_SUPABASE_COLUMNS = {
    "summary": ["mr_code", "clinic", "run_date", "name_en", "name_ar",
                "age", "sex", "id_no", "birth_date", "hospital_code", "visit_number"],
    "investigations": ["mr_code", "clinic", "status", "order_type", "test_name",
                        "physician_comment", "order_date", "requesting_physician",
                        "schedule_date", "order_number", "case_number", "run_date"],
    "medical_reports": ["mr_code", "clinic", "patient_name", "report_date", "description", "run_date"],
    "mdt": ["mr_code", "clinic", "patient_name", "mdt_date", "outcome", "run_date"],
}

DELAY        = 0.4   # seconds between API calls (be gentle on the server)
TIMEOUT      = 30    # seconds per request


# ═══════════════════════════════════════════════════════════════════
# SESSION & LOGIN
# ═══════════════════════════════════════════════════════════════════

def build_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "Accept-Language":  "en-US,en;q=0.9,ar;q=0.8",
        "Connection":       "keep-alive",
        "User-Agent":       ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                             "AppleWebKit/537.36 (KHTML, like Gecko) "
                             "Chrome/150.0.0.0 Safari/537.36"),
    })
    return s


def _extract_gettoken(html: str) -> str:
    """
    Extract the anti-forgery token CMIS embeds in a <script> block on the
    authenticated Home/Index page:

        function gettoken() {
            return'TOKEN_VALUE';
        }

    This is NOT the classic hidden-input __RequestVerificationToken — it's
    read by the page's own JS and sent as a "RequestVerificationToken"
    header on every AJAX call. Without it, calls like
    Home/AuthorizedToSearchPatient are rejected (which looks like "not
    authorized" even though the account/MR is fine).
    """
    m = re.search(r"function\s+gettoken\s*\(\s*\)\s*\{.*?return\s*'([^']+)'",
                  html, re.DOTALL)
    return m.group(1) if m else ""


def login(session: requests.Session) -> bool:
    """
    Login inherited from the working DMS extractor, extended with an extra
    step to fetch the authenticated Home/Index page and pull out the
    gettoken() anti-forgery token, which is required on every subsequent
    AJAX call:
      1. GET Home/Index          → obtains ASP.NET_SessionId cookie
      2. GET Login/LoginModel    → keeps session alive
      3. POST Login/ValidateUserPopupForSEC → authenticates
      4. GET Home/Index (again)  → authenticated page; extract gettoken()
    """
    print("  [1/4] Opening session …")
    session.get(f"{CMIS_BASE}/Home/Index", timeout=TIMEOUT)

    print("  [2/4] Loading login form …")
    session.get(f"{CMIS_BASE}/Login/LoginModel", timeout=TIMEOUT)

    print("  [3/4] Submitting credentials …")
    r = session.post(
        f"{CMIS_BASE}/Login/ValidateUserPopupForSEC",
        data={"UserID": USERNAME, "Password": PASSWORD},
        headers={
            "Accept":           "*/*",
            "Content-Type":     "application/x-www-form-urlencoded; charset=UTF-8",
            "Origin":           "http://41.33.24.253",
            "Referer":          f"{CMIS_BASE}/Home/IndexSessionExpire",
            "X-Requested-With": "XMLHttpRequest",
        },
        timeout=TIMEOUT,
    )
    body = r.text.strip()
    success = r.status_code == 200 and any(
        kw in body for kw in ("hmis-menu", "http", "Index", "Home")
    )
    if not success:
        print(f"  ❌ Login failed — HTTP {r.status_code} | {body[:200]}")
        return False

    print("  [4/4] Fetching authenticated session token …")
    r2 = session.get(f"{CMIS_BASE}/Home/Index", timeout=TIMEOUT)
    token = _extract_gettoken(r2.text)
    if not token:
        print("  ❌ Login looked OK but could not find the gettoken() "
              "verification token on Home/Index — API calls will likely "
              "be rejected as 'not authorized'.")
        return False

    # Attach it at the session level so EVERY subsequent request
    # (api_authorize_patient, api_draw_tree, api_admission, etc.)
    # automatically carries it, matching what the real browser does.
    session.headers["RequestVerificationToken"] = token

    print("  ✅ Login successful")
    return True


# ═══════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════

def _json_safe(r: requests.Response):
    try:
        return r.json()
    except Exception:
        return {}


def _extract_token(html: str) -> str:
    """Extract __RequestVerificationToken from an HTML page."""
    m = re.search(
        r'<input[^>]*name=["\']__RequestVerificationToken["\'][^>]*value=["\']([^"\']+)["\']',
        html,
    )
    if m:
        return m.group(1)
    # alternate attribute order
    m = re.search(
        r'<input[^>]*value=["\']([^"\']+)["\'][^>]*name=["\']__RequestVerificationToken["\']',
        html,
    )
    return m.group(1) if m else ""


def _clean_text(text: str) -> str:
    """Strip trailing slashes, extra whitespace and stray JS artefacts."""
    if not text:
        return ""
    text = text.replace("\\n", "\n").replace("\\/", "/")
    text = re.sub(r'[ \t]+', ' ', text)          # collapse horizontal whitespace
    text = re.sub(r'\n{3,}', '\n\n', text)        # collapse triple+ newlines
    return text.strip()


def _fmt_date_iso(iso_str: str) -> str:
    """Convert ISO datetime string → dd-mm-yyyy."""
    try:
        dt = datetime.fromisoformat(iso_str.replace("Z", "+00:00"))
        return dt.strftime("%d-%m-%Y")
    except Exception:
        pass
    m = re.search(r'(\d{4}-\d{2}-\d{2})', iso_str)
    if m:
        y, mo, d = m.group(1).split("-")
        return f"{d}-{mo}-{y}"
    return iso_str


def _fmt_date_dmy(dmy_str: str) -> str:
    """Convert dd/mm/yyyy or dd/mm/yyyy hh:mm … → dd-mm-yyyy.
    Handles 1- or 2-digit day/month (e.g. 27/7/2026 or 7/27/2026-style).
    """
    # dd/mm/yyyy with optional time suffix  (1- or 2-digit day & month)
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', dmy_str)
    if m:
        return f"{m.group(1).zfill(2)}-{m.group(2).zfill(2)}-{m.group(3)}"
    # already dd-mm-yyyy (or d-m-yyyy)?
    m = re.search(r'(\d{1,2})-(\d{1,2})-(\d{4})', dmy_str.strip())
    if m:
        return f"{m.group(1).zfill(2)}-{m.group(2).zfill(2)}-{m.group(3)}"
    return dmy_str.strip()


def _normalize_any_date(raw: str) -> str:
    """
    Best-effort normalisation of a date string coming straight off a CMIS
    grid cell (format not guaranteed) into dd-mm-yyyy, so it can be
    compared against RUN_DATE. Returns "" if nothing date-shaped is found
    (callers treat that as "doesn't match RUN_DATE" rather than crashing).

    Handles:
      • dd/mm/yyyy HH:MM:SS AM/PM  (long date with time — common in CMIS)
      • d/m/yyyy or dd/m/yyyy      (non-zero-padded day/month)
      • dd-mm-yyyy  /  d-m-yyyy
      • yyyy-mm-dd  (ISO, with or without time)
    """
    if not raw:
        return ""
    raw = str(raw).strip()
    if not raw:
        return ""

    # yyyy-mm-dd must be tried FIRST before dd/mm/yyyy so that "2026-07-27"
    # is not accidentally matched by the dd-mm-yyyy pattern as "20-26-07".
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', raw)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"

    # dd/mm/yyyy  (1- or 2-digit day & month, with optional time suffix)
    # e.g. "27/07/2026 02:49:56 PM"  or  "27/7/2026"
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', raw)
    if m:
        return f"{m.group(1).zfill(2)}-{m.group(2).zfill(2)}-{m.group(3)}"

    # dd-mm-yyyy  (1- or 2-digit day & month, with optional time suffix)
    m = re.search(r'(\d{1,2})-(\d{1,2})-(\d{4})', raw)
    if m:
        return f"{m.group(1).zfill(2)}-{m.group(2).zfill(2)}-{m.group(3)}"

    return ""


def _to_iso_date(dmy_str: str) -> str:
    """
    Convert a dd-mm-yyyy string (RUN_DATE's format, used everywhere in
    this script) into yyyy-mm-dd, which is what Postgres/Supabase's
    `date` column type expects. Used ONLY when building the Supabase
    payload — the Excel sheets keep dd-mm-yyyy as before.

    Without this, pushing "27-07-2026" straight into a `date` column
    is ambiguous/invalid to Postgres and the upsert will fail.
    """
    if not dmy_str:
        return None
    s = str(dmy_str).strip()
    m = re.match(r'^(\d{2})-(\d{2})-(\d{4})$', s)
    if m:
        d, mo, y = m.groups()
        return f"{y}-{mo}-{d}"
    if re.match(r'^\d{4}-\d{2}-\d{2}$', s):
        return s  # already ISO
    return None  # unrecognised — let it be NULL rather than break the insert


# ═══════════════════════════════════════════════════════════════════
# API CALLS  (in workflow order)
# ═══════════════════════════════════════════════════════════════════

_XHR_HEADERS = {
    "Accept":           "*/*",
    "Content-Type":     "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Referer":          f"{CMIS_BASE}/Home/Index",
}


# ─── Step 1 ──────────────────────────────────────────────────────
def api_patient_get_by_mr(session, mr) -> dict:
    r = session.post(
        f"{CMIS_BASE}/Admission/PatientGetByMR",
        data={"MR": mr},
        headers=_XHR_HEADERS,
        timeout=TIMEOUT,
    )
    return _json_safe(r)


# ─── Step 2 ──────────────────────────────────────────────────────
def api_authorize_patient(session, mr) -> bool:
    r = session.post(
        f"{CMIS_BASE}/Home/AuthorizedToSearchPatient",
        data={"MR": mr},
        headers=_XHR_HEADERS,
        timeout=TIMEOUT,
    )
    return r.text.strip().lower() == "true"


# ─── Step 3 ──────────────────────────────────────────────────────
def api_draw_tree(session, mr) -> tuple:
    """
    Returns (hospital_code: str, visit_number: int|None).

    The visit number is extracted from calls like:
        onclick=DisplayPatientData('01', 6246, 17000, 'O', ...)

    NOTE: earlier versions only matched type 'O' (Outpatient), which
    silently returned no visit number for MRs whose most recent (or
    only) visits are Inpatient/Emergency/other types. We now match
    DisplayPatientData calls tied to THIS MR regardless of type, and
    prefer the largest 'O' visit number when one exists, falling back
    to the largest visit number of any type otherwise.
    """
    r = session.post(
        f"{CMIS_BASE}/Home/DrawTreeByMR",
        data={"MR": mr, "OrderBySpeciality": "False", "DoctorGroup": "False"},
        headers=_XHR_HEADERS,
        timeout=TIMEOUT,
    )
    # The server sends the onclick="DisplayPatientData('01',6246,...)"
    # arguments as HTML-entity-encoded quotes (&#39;) rather than literal
    # apostrophes. Decode entities first or the regexes below never match.
    html = html_lib.unescape(r.text)

    # Hospital code — first argument of DisplayPatientData
    hosp_code = HOSPITAL_CODE
    m = re.search(r"DisplayPatientData\s*\(\s*'([^']+)'", html)
    if m and m.group(1):
        hosp_code = m.group(1)

    # Visit number — match hosp_code, MR, visit_num, type together, and
    # keep only rows belonging to this MR (defensive against any stray
    # entries in the returned tree).
    mr_str = str(mr).strip()
    best_o = None      # best (largest) visit number among type 'O' rows
    best_any = None     # best (largest) visit number among ANY type
    for match in re.finditer(
        r"DisplayPatientData\s*\(\s*'([^']*)'\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*'([A-Za-z]*)'",
        html, re.IGNORECASE,
    ):
        row_hosp, row_mr, row_visit, row_type = match.groups()
        if row_mr != mr_str:
            continue
        v = int(row_visit)
        if best_any is None or v > best_any:
            best_any = v
        if row_type.upper() == "O" and (best_o is None or v > best_o):
            best_o = v
        if row_hosp:
            hosp_code = row_hosp

    visit_num = best_o if best_o is not None else best_any

    return hosp_code, visit_num


# ─── Step 4 ──────────────────────────────────────────────────────
def api_admission(session, mr, hosp_code, visit_num) -> dict:
    r = session.post(
        f"{CMIS_BASE}/Admission/Admission",
        data={"HospitalCode": hosp_code, "MR": mr, "VisitNumber": str(visit_num)},
        headers=_XHR_HEADERS,
        timeout=TIMEOUT,
    )
    return _json_safe(r)


# ─── Step 5 ──────────────────────────────────────────────────────
def api_patient_note(session, mr, hosp_code, visit_num):
    """Called as per the workflow — we don't extract data from it, just call it."""
    session.get(
        f"{COMMON_API}/PatientNote/load",
        params={"MR": mr, "HospCode": hosp_code, "VisitNumber": str(visit_num)},
        headers={"Accept": "*/*", "X-Requested-With": "XMLHttpRequest",
                 "Referer": f"{CMIS_BASE}/Home/Index"},
        timeout=TIMEOUT,
    )


# ─── Step 6 ──────────────────────────────────────────────────────
def api_investigations(session, mr) -> list:
    """
    The _ parameter is the current Unix timestamp in milliseconds
    (same as JavaScript's Date.now()).
    Returns a parsed list of investigation dicts.
    """
    ts = int(time.time() * 1000)
    r = session.get(
        f"{CMIS_BASE}/Investigations/Investigation",
        params={"_": ts},
        headers={
            "Accept":   "text/html,application/xhtml+xml,*/*",
            "Referer":  f"{CMIS_BASE}/Home/Index",
        },
        timeout=TIMEOUT,
    )
    return _parse_investigations(r.text, mr)


def _parse_investigations(html: str, mr) -> list:
    soup = BeautifulSoup(html, "lxml")
    results = []

    table = soup.find("table", class_="grid")
    if not table:
        return results
    tbody = table.find("tbody")
    if not tbody:
        return results

    for row in tbody.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) < 10:
            continue

        def cell_label(idx, default=""):
            if idx >= len(cells):
                return default
            lbl = cells[idx].find("label")
            if lbl:
                # Use separator=' ' so that date and time in adjacent child
                # elements are joined as "27/07/2026 02:49:56 PM" rather
                # than fused into "27/07/202602:49:56 PM".
                return lbl.get_text(separator=" ", strip=True)
            spn = cells[idx].find("span")
            if spn:
                return spn.get_text(separator=" ", strip=True)
            return cells[idx].get_text(separator=" ", strip=True)

        # Status  (col 2: div > span)
        status = ""
        if len(cells) > 2:
            sp = cells[2].find("span")
            if sp:
                status = sp.get_text(strip=True)

        # Test name (col 7): text node between spans / after gridHint span
        test_name = ""
        if len(cells) > 7:
            tc = cells[7]
            parts = []
            for node in tc.children:
                if isinstance(node, Tag):
                    classes = node.get("class", [])
                    if "gridHint" in classes:
                        continue           # skip the "waiting/Approved" badge
                    if node.name == "a":
                        continue           # skip PACS image links
                    t = node.get_text(strip=True)
                    if t:
                        parts.append(t)
                else:
                    t = str(node).strip()
                    if t:
                        parts.append(t)
            test_name = " ".join(parts).strip()
            # Remove extra whitespace
            test_name = re.sub(r"\s+", " ", test_name).strip()
            # Remove trailing empty parenthetical artifact, e.g. "Calcium (Total) ( , )"
            test_name = re.sub(r"\(\s*,\s*\)\s*$", "", test_name).strip()

        results.append({
            "MR":                    mr,
            "Status":                status,
            "Order Type":            cell_label(5),
            "Order Category":        cell_label(6),
            "Test Name":             test_name,
            "Physician Comment":     cell_label(8),
            "Order Date":            cell_label(9),
            "Requesting Physician":  cell_label(10),
            "Schedule Date":         cell_label(11),
            "Order Number":          cell_label(12),
            "Case Number":           cell_label(21),
        })

    return results


# ─── Step 8 ──────────────────────────────────────────────────────
def api_medical_sheets(session, mr, hosp_code, visit_num) -> tuple:
    """
    Returns (sheets: list[dict], token: str).
    sheets are filtered to SheetCode in {033, 02}.
    Handles pagination via MedicalSheets/Paging.
    """
    r = session.get(
        f"{CMIS_BASE}/MedicalSheets/Index",
        headers={"Accept": "text/html,*/*", "Referer": f"{CMIS_BASE}/Home/Index"},
        timeout=TIMEOUT,
    )
    html = r.text
    token = _extract_token(html)

    all_sheets = _parse_medical_sheets(html)

    # Check pagination: pagingMedicalReports(pageCount, startPage, ...)
    m = re.search(r"pagingMedicalReports\s*\(\s*(\d+)\s*,", html)
    page_count = int(m.group(1)) if m else 1

    for page_no in range(2, page_count + 1):
        time.sleep(DELAY)
        rp = session.post(
            f"{CMIS_BASE}/MedicalSheets/Paging",
            data={"PageNo": page_no, "Filter": "false", "group": ""},
            headers={
                "Accept":                    "*/*",
                "Content-Type":              "application/x-www-form-urlencoded; charset=UTF-8",
                "X-Requested-With":          "XMLHttpRequest",
                "RequestVerificationToken":  token,
                "Referer":                   f"{CMIS_BASE}/MedicalSheets/Index",
            },
            timeout=TIMEOUT,
        )
        try:
            data = rp.json()
            grid_html = data.get("grid", "")
        except Exception:
            grid_html = rp.text
        all_sheets.extend(_parse_medical_sheets(grid_html))

    target_codes = {"033", "02"}
    filtered = [s for s in all_sheets if s.get("SheetCode", "").strip() in target_codes]
    return filtered, token


def _parse_medical_sheets(html: str) -> list:
    soup = BeautifulSoup(html, "lxml")
    sheets = []

    table = soup.find("table", class_="grid")
    if not table:
        return sheets
    tbody = table.find("tbody")
    if not tbody:
        return sheets

    for row in tbody.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) < 4:
            continue

        sheet_code = seq_num = visit_num_sheet = ""
        sheet_group = "MRS"

        # Scan all <a> tags for GetSheetUrlString with OpenMode=SD
        for a in row.find_all("a"):
            onclick = a.get("onclick", "")
            m = re.search(
                r"GetSheetUrlString\?"
                r"OpenMode=SD"
                r"&SequenceNumber=(\d+)"
                r"&SheetCode=([^&']+)"
                r"&SpecialSheetCode=([^&']*)"
                r"&VisitNumber=(\d+)"
                r"&SheetGroup=([^&'\"]+)",
                onclick,
            )
            if m:
                seq_num          = m.group(1)
                sheet_code       = m.group(2)
                visit_num_sheet  = m.group(4)
                sheet_group      = m.group(5).strip("'\"")
                break

        if not sheet_code:
            continue  # no parseable sheet info — skip row

        sheet_name = ""
        if len(cells) > 3:
            lbl = cells[3].find("label")
            sheet_name = lbl.get_text(strip=True) if lbl else cells[3].get_text(strip=True)

        creation_date = ""
        if len(cells) > 5:
            lbl = cells[5].find("label", id="Date")
            if not lbl:
                lbl = cells[5].find("label")
            # Use separator=' ' so date+time from separate child elements
            # don't fuse (e.g. avoids "27/07/202602:49:56 PM").
            creation_date = (lbl.get_text(separator=" ", strip=True)
                             if lbl else cells[5].get_text(separator=" ", strip=True))

        sheets.append({
            "SheetCode":      sheet_code.strip(),
            "SheetName":      sheet_name,
            "SequenceNumber": seq_num,
            "VisitNumber":    visit_num_sheet,
            "SheetGroup":     sheet_group,
            "CreationDate":   creation_date,
        })

    return sheets


# ─── Step 9 ──────────────────────────────────────────────────────
def api_get_sheet_url(session, sheet: dict, token: str) -> str:
    """
    POST to GetSheetUrlString; response body is the sheet URL.

    IMPORTANT: the browser's OpenSheetURL() sends NO request body at all —
    every parameter (OpenMode, SequenceNumber, SheetCode, SpecialSheetCode,
    VisitNumber, SheetGroup) is embedded directly in the URL's query string.
    Sending them as a form body (the old `data={...}` approach) means the
    server never receives them, so it 404s and returns an HTML error page
    instead of a sheet URL.
    """
    query = urlencode({
        "OpenMode":         "SD",
        "SequenceNumber":   sheet["SequenceNumber"],
        "SheetCode":        sheet["SheetCode"],
        "SpecialSheetCode": "",
        "VisitNumber":      sheet["VisitNumber"],
        "SheetGroup":       sheet.get("SheetGroup", "MRS"),
    })
    r = session.post(
        f"{CMIS_BASE}/MedicalSheets/GetSheetUrlString?{query}",
        headers={
            "Accept":                   "*/*",
            "X-Requested-With":         "XMLHttpRequest",
            "RequestVerificationToken": token,
            "Referer":                  f"{CMIS_BASE}/MedicalSheets/Index",
        },
        timeout=TIMEOUT,
    )
    url = r.text.strip().strip('"')

    # Safety guard: a bad/HTML response (404 page, login redirect, etc.)
    # must never be treated as a usable sheet URL — it would otherwise get
    # stuffed into a Referer header downstream and crash with an
    # InvalidHeader error (embedded \r\n / control characters).
    if not url or "<" in url or "\n" in url or "\r" in url or not url.lower().startswith("http"):
        return ""

    return url


# ─── Step 10 ─────────────────────────────────────────────────────
def api_get_user_control_data(
    session,
    sheet_url: str,
    mr,
    hosp_code: str,
    visit_num,          # current patient visit
    sheet: dict,
) -> list:
    """
    Navigate to the sheet URL to establish Angular session context,
    then call GetUserControlData on the GeneralSheetManagerApi.

    CaseNo  = visit number from the sheet row (NOT necessarily current visit)
    SeqNo   = sequence number from the sheet row
    """
    # Defensive guard: a sheet_url that isn't a clean http(s) URL (e.g. an
    # HTML error page returned by api_get_sheet_url) must never be used as
    # a header value — doing so raises requests.exceptions.InvalidHeader
    # because of embedded \r\n / control characters. Fall back to the
    # normal Index page instead.
    safe_referer = f"{CMIS_BASE}/MedicalSheets/Index"
    if sheet_url and "\n" not in sheet_url and "\r" not in sheet_url and "<" not in sheet_url \
            and sheet_url.lower().startswith("http"):
        safe_referer = sheet_url

    if sheet_url:
        try:
            session.get(
                sheet_url, timeout=TIMEOUT,
                headers={"Referer": f"{CMIS_BASE}/MedicalSheets/Index"},
            )
        except Exception:
            pass

    payload = {
        "Hospcode":     hosp_code,
        "MR":           str(mr),
        "SheetCode":    sheet["SheetCode"],
        "FieldName":    "",
        "CaseNo":       str(sheet["VisitNumber"] or visit_num),
        "SeqNo":        str(sheet["SequenceNumber"]),
        "ModuleCode":   "MRM",
        "RegisteredTo": "null",
    }

    r = session.post(
        f"{SHEET_API_BASE}/GetUserControlData",
        json=payload,
        headers={
            "Accept":       "application/json, text/plain, */*",
            "Content-Type": "application/json",
            "Origin":       "http://41.33.24.253",
            "Referer":      safe_referer,
        },
        timeout=TIMEOUT,
    )
    try:
        result = r.json()
        return result if isinstance(result, list) else []
    except Exception:
        return []


# ═══════════════════════════════════════════════════════════════════
# DATA EXTRACTION FROM SHEET FIELDS
# ═══════════════════════════════════════════════════════════════════

def extract_medical_report(fields: list, creation_date: str) -> tuple:
    """
    From Medical Report (SheetCode 033) field array:
      description ← fieldName == "text-10"
      date        ← earliest modifyDate in the array (= creation date)
                    fallback to creation_date from the sheets table
    Returns (date_str, description).
    """
    description = ""
    earliest_iso = None

    for f in fields:
        fname = f.get("fieldName", "")
        fval  = f.get("fieldValue", "")
        mdate = f.get("modifyDate", "")

        if fname == "text-10":
            description = fval

        if mdate:
            if earliest_iso is None or mdate < earliest_iso:
                earliest_iso = mdate

    date_str = _fmt_date_iso(earliest_iso) if earliest_iso else _fmt_date_dmy(creation_date)
    description = _clean_text(description)
    return date_str, description


def extract_mdt(fields: list, creation_date: str) -> tuple:
    """
    From MDT form (SheetCode 02) field array:
      outcome ← fieldName == "text-5"
      date    ← fieldName == "Input-21"  (format "yyyy-mm-dd")
                fallback to creation_date from sheets table
    Returns (date_str, outcome).
    """
    outcome  = ""
    date_str = ""

    for f in fields:
        fname = f.get("fieldName", "")
        fval  = f.get("fieldValue", "")

        if fname == "text-5":
            outcome = fval
        elif fname == "Input-21" and not date_str:
            date_str = fval  # e.g. "2026-07-08"

    if not date_str:
        date_str = _fmt_date_dmy(creation_date)
    else:
        # Convert yyyy-mm-dd → dd-mm-yyyy
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", date_str)
        if m:
            date_str = f"{m.group(3)}-{m.group(2)}-{m.group(1)}"

    outcome = _clean_text(outcome)
    return date_str, outcome


# ═══════════════════════════════════════════════════════════════════
# STATISTICS
# ═══════════════════════════════════════════════════════════════════
#
# This is a RULE-BASED classifier, not a trained medical-NLP model.
# It uses two reliable signals that are already in the extracted data:
#   1. "Order Type" from the CMIS grid — this is "LAB" or "RAD" and is
#      the system's own classification, not guessed from text.
#   2. For RAD orders, a keyword search against the test name to split
#      plain imaging/scans from interventional-radiology procedures
#      (biopsy, portacath, tapping/aspiration, drain/catheter, etc).
#
# Anything that doesn't confidently match (e.g. an unexpected Order
# Type, or a RAD test name we don't recognise) is bucketed as
# "Unclassified (manual review)" and EXCLUDED from the ratio/average
# calculations below, per the requirement that ambiguous data must not
# skew the stats. It is still counted and listed so nothing is silently
# dropped.

_INTERVENTION_KEYWORDS_EN = [
    "biopsy", "portacath", "port-a-cath", "port a cath", "port cath",
    "tapping", "aspiration", "drain", "catheter", "picc", "hook wire",
    "hookwire", "localization", "localisation", "port insertion",
    "line insertion", "central line",
]
_INTERVENTION_KEYWORDS_AR = [
    "خزعة", "بايوبسي", "بورت", "بذل", "قسطرة", "ماهوك", "تركيب بورت",
    "شفط",
]

CATEGORY_LAB          = "Blood / Lab Test"
CATEGORY_IMAGING      = "Imaging / Scan"
CATEGORY_INTERVENTION = "Interventional Radiology"
CATEGORY_UNCLASSIFIED = "Unclassified (manual review)"


def classify_investigation(order_type: str, test_name: str) -> str:
    """Rule-based category for one investigation row. See module notes above."""
    ot = (order_type or "").strip().upper()
    tn_lower = (test_name or "").lower()

    if ot == "LAB":
        return CATEGORY_LAB

    if ot == "RAD":
        for kw in _INTERVENTION_KEYWORDS_EN:
            if kw in tn_lower:
                return CATEGORY_INTERVENTION
        for kw in _INTERVENTION_KEYWORDS_AR:
            if kw in (test_name or ""):
                return CATEGORY_INTERVENTION
        return CATEGORY_IMAGING

    return CATEGORY_UNCLASSIFIED


def build_statistics(data: dict) -> dict:
    """
    Computes everything the Statistics sheet needs from the already-extracted
    data. Returns a dict of sections so write_excel can lay them out.
    investigations rows are: [MR, Clinic, Status, Order Type, Test Name,
        Physician Comment, Order Date, Requesting Physician, Schedule Date,
        Order Number, Case Number]
    """
    investigations = data["investigations"]
    summary        = data["summary"]
    med_reports    = data["medical_reports"]
    mdt            = data["mdt"]

    total_patients = len(summary)

    # ── Category breakdown ────────────────────────────────────────
    cat_counts   = {}
    cat_patients = {}
    for row in investigations:
        mr, _clinic, _status, order_type, test_name = row[0], row[1], row[2], row[3], row[4]
        cat = classify_investigation(order_type, test_name)
        cat_counts[cat] = cat_counts.get(cat, 0) + 1
        cat_patients.setdefault(cat, set()).add(mr)

    total_investigations = len(investigations)
    classified_total = sum(v for k, v in cat_counts.items() if k != CATEGORY_UNCLASSIFIED)

    category_rows = []
    for cat in (CATEGORY_LAB, CATEGORY_IMAGING, CATEGORY_INTERVENTION, CATEGORY_UNCLASSIFIED):
        count = cat_counts.get(cat, 0)
        pct_base = classified_total if cat != CATEGORY_UNCLASSIFIED else total_investigations
        pct = (count / pct_base * 100) if pct_base and cat != CATEGORY_UNCLASSIFIED else (
              count / total_investigations * 100 if total_investigations else 0)
        n_pts = len(cat_patients.get(cat, set()))
        category_rows.append([cat, count, f"{pct:.1f}%", n_pts])

    # ── Per-patient / per-visit averages (excludes Unclassified) ───
    case_numbers_per_mr = {}
    classified_count_per_mr = {}
    for row in investigations:
        mr, order_type, test_name, case_no = row[0], row[3], row[4], row[10]
        cat = classify_investigation(order_type, test_name)
        if cat == CATEGORY_UNCLASSIFIED:
            continue
        classified_count_per_mr[mr] = classified_count_per_mr.get(mr, 0) + 1
        if case_no:
            case_numbers_per_mr.setdefault(mr, set()).add(case_no)

    patients_with_classified = len(classified_count_per_mr)
    avg_per_patient = (classified_total / patients_with_classified) if patients_with_classified else 0

    total_visits = sum(len(v) for v in case_numbers_per_mr.values()) or 0
    avg_per_visit = (classified_total / total_visits) if total_visits else 0

    # ── By requesting physician ─────────────────────────────────────
    phys_stats = {}  # physician -> {lab, imaging, intervention, patients:set}
    for row in investigations:
        mr, order_type, test_name, physician = row[0], row[3], row[4], row[7]
        cat = classify_investigation(order_type, test_name)
        physician = (physician or "Unknown").strip() or "Unknown"
        d = phys_stats.setdefault(physician, {
            CATEGORY_LAB: 0, CATEGORY_IMAGING: 0,
            CATEGORY_INTERVENTION: 0, CATEGORY_UNCLASSIFIED: 0,
            "patients": set(),
        })
        d[cat] += 1
        d["patients"].add(mr)

    physician_rows = []
    for phys, d in sorted(phys_stats.items(), key=lambda kv: -sum(
            kv[1][c] for c in (CATEGORY_LAB, CATEGORY_IMAGING, CATEGORY_INTERVENTION))):
        classified = d[CATEGORY_LAB] + d[CATEGORY_IMAGING] + d[CATEGORY_INTERVENTION]
        n_pts = len(d["patients"])
        avg = (classified / n_pts) if n_pts else 0
        physician_rows.append([
            phys, classified, d[CATEGORY_LAB], d[CATEGORY_IMAGING],
            d[CATEGORY_INTERVENTION], n_pts, f"{avg:.1f}",
        ])

    # ── By clinic ──────────────────────────────────────────────────
    clinic_patients = {}
    for row in summary:
        mr, clinic = row[0], row[1]
        clinic_patients.setdefault(clinic or "(unknown)", set()).add(mr)
    clinic_rows = [[clinic, len(mrs)] for clinic, mrs in
                   sorted(clinic_patients.items(), key=lambda kv: -len(kv[1]))]

    # ── MDT presence ────────────────────────────────────────────────
    patients_with_mdt = len({row[0] for row in mdt})

    overview_rows = [
        ("Total patients processed",                str(total_patients)),
        ("Total investigations extracted",           str(total_investigations)),
        ("  — classified (Lab/Imaging/Intervention)", str(classified_total)),
        ("  — unclassified (excluded from ratios)",  str(cat_counts.get(CATEGORY_UNCLASSIFIED, 0))),
        ("Avg. classified investigations / patient", f"{avg_per_patient:.2f}"),
        ("Avg. classified investigations / visit",   f"{avg_per_visit:.2f}"),
        ("Total medical reports (033) extracted",    str(len(med_reports))),
        ("Total MDT forms (02) extracted",           str(len(mdt))),
        ("Patients with ≥1 MDT form",                str(patients_with_mdt)),
    ]

    limitations = [
        "Classification is RULE-BASED: Order Type ('LAB'/'RAD') from the system, plus keyword",
        "matching on the test name to split Imaging/Scan from Interventional Radiology",
        "(biopsy, portacath, tapping/aspiration, drains/catheters, etc). It does not use a",
        "trained medical-NLP model, so uncommon or newly-worded test names may fall into",
        "'Unclassified' — these are shown above but excluded from all averages/ratios so they",
        "don't distort the numbers. Review the Investigations sheet for the exact rows.",
        "",
        "These numbers describe the CUMULATIVE master workbook (every day accumulated so",
        "far), not just today's run — see the Run Info sheet for what today's run added.",
    ]

    return {
        "overview":     overview_rows,
        "categories":   category_rows,
        "physicians":   physician_rows,
        "clinics":      clinic_rows,
        "limitations":  limitations,
    }


# ═══════════════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ═══════════════════════════════════════════════════════════════════

_HDR_FILL   = PatternFill("solid", start_color="1F4E79")
_HDR_FONT   = Font(bold=True, color="FFFFFF", name="Arial", size=11)
_ALT_FILL   = PatternFill("solid", start_color="DCE6F1")
_NORM_FILL  = PatternFill("solid", start_color="FFFFFF")
_THIN_SIDE  = Side(style="thin", color="B8CCE4")
_BORDER     = Border(left=_THIN_SIDE, right=_THIN_SIDE,
                     top=_THIN_SIDE, bottom=_THIN_SIDE)


def _write_sheet(ws, headers: list, widths: list, rows: list, rtl: bool = False):
    if rtl:
        ws.sheet_view.rightToLeft = True

    for col_idx, (hdr, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=1, column=col_idx, value=hdr)
        c.font      = _HDR_FONT
        c.fill      = _HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = _BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    for row_num, row_data in enumerate(rows, 2):
        fill = _ALT_FILL if row_num % 2 == 0 else _NORM_FILL
        for col_idx, val in enumerate(row_data, 1):
            text = str(val) if val is not None else ""
            c = ws.cell(row=row_num, column=col_idx, value=text)
            c.font   = Font(name="Arial", size=10)
            c.fill   = fill
            c.border = _BORDER
            wrap = widths[col_idx - 1] >= 40
            c.alignment = Alignment(
                horizontal="left" if widths[col_idx - 1] >= 20 else "center",
                vertical="top",
                wrap_text=wrap,
            )
        ws.row_dimensions[row_num].height = 15


def _write_stats_sheet(ws, stats: dict):
    """Lays out the Statistics sheet as stacked mini-tables (not a uniform grid)."""
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 14

    section_font = Font(bold=True, color="FFFFFF", name="Arial", size=12)
    section_fill = PatternFill("solid", start_color="1F4E79")
    hdr_font  = Font(bold=True, name="Arial", size=10, color="FFFFFF")
    hdr_fill  = PatternFill("solid", start_color="4472C4")
    body_font = Font(name="Arial", size=10)
    note_font = Font(name="Arial", size=9, italic=True, color="595959")

    r = 1

    def section_title(title):
        nonlocal r
        c = ws.cell(r, 1, title)
        c.font = section_font
        c.fill = section_fill
        for col in range(2, 8):
            ws.cell(r, col).fill = section_fill
        r += 1

    def blank():
        nonlocal r
        r += 1

    section_title("Overview")
    for label, val in stats["overview"]:
        ws.cell(r, 1, label).font = body_font
        ws.cell(r, 2, val).font = body_font
        r += 1
    blank()

    section_title("Investigation Category Breakdown (blood labs vs. scans vs. interventional radiology)")
    headers = ["Category", "Count", "% of classified total", "Unique Patients"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(r, i, h)
        c.font = hdr_font
        c.fill = hdr_fill
    r += 1
    for row in stats["categories"]:
        for i, val in enumerate(row, 1):
            ws.cell(r, i, val).font = body_font
        r += 1
    blank()

    section_title("By Requesting Physician")
    headers = ["Physician", "Total Classified", "Blood/Lab", "Imaging/Scan",
               "Intervention", "Unique Patients", "Avg / Patient"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(r, i, h)
        c.font = hdr_font
        c.fill = hdr_fill
    r += 1
    for row in stats["physicians"]:
        for i, val in enumerate(row, 1):
            ws.cell(r, i, val).font = body_font
        r += 1
    blank()

    section_title("By Clinic (unique patients, cumulative)")
    headers = ["Clinic", "Unique Patients"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(r, i, h)
        c.font = hdr_font
        c.fill = hdr_fill
    r += 1
    for row in stats["clinics"]:
        for i, val in enumerate(row, 1):
            ws.cell(r, i, val).font = body_font
        r += 1
    blank()

    section_title("Notes & Limitations")
    for line in stats["limitations"]:
        ws.cell(r, 1, line).font = note_font
        r += 1


def write_excel(data: dict, output_path: str, run_info: dict):
    """
    Writes the FULL (cumulative) dataset to output_path. `data` already
    contains every row accumulated so far (today's new rows merged with
    whatever existed before) — see accumulate_and_save() below, which is
    what actually orchestrates load-existing → merge → write.
    """
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Patient Summary"
    _write_sheet(
        ws1,
        ["MR Code", "Clinic", "Run Date", "Name (English)", "Name (Arabic)",
         "Age", "Sex", "ID No.", "Birth Date", "Hospital Code", "Visit Number"],
        [12, 22, 13, 34, 34, 8, 10, 20, 14, 14, 14],
        data["summary"],
    )

    ws2 = wb.create_sheet("Investigations")
    _write_sheet(
        ws2,
        ["MR", "Clinic", "Status", "Order Type", "Test Name",
         "Physician Comment", "Order Date", "Requesting Physician",
         "Schedule Date", "Order Number", "Case Number", "Run Date"],
        [10, 20, 13, 12, 38, 32, 14, 24, 14, 13, 12, 12],
        data["investigations"],
    )

    ws3 = wb.create_sheet("Medical Reports (033)")
    _write_sheet(
        ws3,
        ["MR", "Clinic", "Patient Name", "Report Date", "Description", "Run Date"],
        [10, 20, 30, 14, 90, 12],
        data["medical_reports"],
    )

    ws4 = wb.create_sheet("MDT (02)")
    _write_sheet(
        ws4,
        ["MR", "Clinic", "Patient Name", "MDT Date", "Outcome / Plan", "Run Date"],
        [10, 20, 30, 14, 90, 12],
        data["mdt"],
    )

    stats = build_statistics(data)
    ws_stats = wb.create_sheet("Statistics")
    _write_stats_sheet(ws_stats, stats)

    ws6 = wb.create_sheet("Run Info")
    info_rows = [
        ("Last Run Date",                    run_info.get("run_date", "")),
        ("Last Run Timestamp",               run_info.get("timestamp", "")),
        ("MR Codes From Queue (this run)",   str(run_info.get("mr_from_queue", 0))),
        ("Patients Processed (this run)",    str(run_info.get("patients_processed", 0))),
        ("New Investigation Rows (this run)", str(run_info.get("new_investigations", 0))),
        ("New Medical Report Rows (this run)", str(run_info.get("new_medical_reports", 0))),
        ("New MDT Rows (this run)",          str(run_info.get("new_mdt", 0))),
        ("Errors / Skipped (this run)",      str(run_info.get("errors", 0))),
        ("Supabase push (this run)",         run_info.get("supabase_status", "disabled")),
        ("", ""),
        ("TOTAL Patient Rows (cumulative)",       str(len(data["summary"]))),
        ("TOTAL Investigation Rows (cumulative)", str(len(data["investigations"]))),
        ("TOTAL Medical Report Rows (cumulative)", str(len(data["medical_reports"]))),
        ("TOTAL MDT Rows (cumulative)",           str(len(data["mdt"]))),
    ]
    ws6.cell(1, 1, "Item").font  = Font(bold=True, name="Arial", size=11)
    ws6.cell(1, 2, "Value").font = Font(bold=True, name="Arial", size=11)
    for rn, (k, v) in enumerate(info_rows, 2):
        ws6.cell(rn, 1, k).font = Font(bold=True, name="Arial")
        ws6.cell(rn, 2, v).font = Font(name="Arial")
    ws6.column_dimensions["A"].width = 34
    ws6.column_dimensions["B"].width = 55

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    wb.save(output_path)
    print(f"  ✅  Excel saved → {output_path}")


# ═══════════════════════════════════════════════════════════════════
# DATA STORE  —  load / merge / persist the cumulative dataset,
#                and push new rows to Supabase
# ═══════════════════════════════════════════════════════════════════

_SHEET_KEYS = ("summary", "investigations", "medical_reports", "mdt")

_SHEET_COLUMNS = {
    "summary": [
        "MR Code", "Clinic", "Run Date", "Name (English)", "Name (Arabic)",
        "Age", "Sex", "ID No.", "Birth Date", "Hospital Code", "Visit Number",
    ],
    "investigations": [
        "MR", "Clinic", "Status", "Order Type", "Test Name",
        "Physician Comment", "Order Date", "Requesting Physician",
        "Schedule Date", "Order Number", "Case Number", "Run Date",
    ],
    "medical_reports": [
        "MR", "Clinic", "Patient Name", "Report Date", "Description", "Run Date",
    ],
    "mdt": [
        "MR", "Clinic", "Patient Name", "MDT Date", "Outcome / Plan", "Run Date",
    ],
}

_SHEET_TITLES = {
    "summary":          "Patient Summary",
    "investigations":   "Investigations",
    "medical_reports":  "Medical Reports (033)",
    "mdt":              "MDT (02)",
}


def load_master_data(path: str) -> dict:
    """
    Reads whatever is already saved at `path` and returns it in the same
    {summary, investigations, medical_reports, mdt} shape that main()
    builds fresh data in. Returns all-empty lists if the file doesn't
    exist yet (first-ever run).
    """
    empty = {k: [] for k in _SHEET_KEYS}
    if not os.path.exists(path):
        return empty

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        print(f"  ⚠  Could not open existing master workbook ({e}) — "
              f"treating as empty (a backup of the old file is recommended "
              f"before this run overwrites it).")
        return empty

    out = {}
    for key in _SHEET_KEYS:
        title = _SHEET_TITLES[key]
        if title not in wb.sheetnames:
            out[key] = []
            continue
        ws = wb[title]
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None or all(v is None or str(v).strip() == "" for v in row):
                continue
            rows.append(["" if v is None else v for v in row])
        out[key] = rows
    wb.close()
    return out


def merge_rows(existing_rows: list, new_rows: list) -> tuple:
    """
    Appends new_rows to existing_rows, skipping exact duplicates (same
    values in every column). Order is preserved: existing rows first,
    then genuinely-new rows.

    Returns (merged_rows, num_added).
    """
    seen = set(tuple(str(v) for v in row) for row in existing_rows)
    merged = list(existing_rows)
    added = 0
    for row in new_rows:
        key = tuple(str(v) for v in row)
        if key in seen:
            continue
        seen.add(key)
        merged.append(row)
        added += 1
    return merged, added


def _row_hash(row: list) -> str:
    """MD5 hash of every value in a row — the Supabase dedup key."""
    joined = "|".join("" if v is None else str(v) for v in row)
    return hashlib.md5(joined.encode("utf-8")).hexdigest()


def test_supabase_connection() -> bool:
    """
    Quick round-trip against patient_summary to confirm URL/key are
    valid and the table exists, before spending a whole run's worth of
    extraction only to find the push fails at the end.
    """
    if create_client is None:
        print("  ❌  'supabase' package not installed — run: pip install supabase")
        return False
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        print("  ❌  SUPABASE_URL / SUPABASE_SERVICE_KEY not set in the environment "
              "(check run_extractor.bat).")
        return False
    try:
        client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)
        client.table("patient_summary").select("id").limit(1).execute()
        print("  ✅  Supabase connection OK")
        return True
    except Exception as e:
        print(f"  ❌  Supabase connection failed: {e}")
        return False


def push_rows_to_supabase(table_key: str, columns: list, rows: list) -> bool:
    """
    Upsert `rows` (only the NEW rows from this run, in the same column
    order as _SHEET_COLUMNS[table_key]) into the matching Supabase
    table, de-duplicated by an md5 row_hash. Returns True on success,
    False on failure (caller decides whether that's fatal — it isn't;
    Excel already has the data either way).
    """
    if not rows:
        return True
    if create_client is None or not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        print(f"  ⚠  Supabase not configured — skipping push for '{table_key}'")
        return False

    db_cols = _SUPABASE_COLUMNS.get(table_key)
    table_name = _SUPABASE_TABLE_NAMES.get(table_key)
    if not db_cols or not table_name:
        print(f"  ⚠  Unknown table_key '{table_key}' — skipping push")
        return False

    try:
        client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)

        payload = []
        for row in rows:
            rec = dict(zip(db_cols, row))

            # All DB columns here are `text` EXCEPT run_date, which is
            # `date`. Stringify everything else, and convert run_date
            # from dd-mm-yyyy to yyyy-mm-dd so Postgres accepts it —
            # this mismatch was the main cause of failed upserts.
            for k in list(rec.keys()):
                if k == "run_date":
                    rec[k] = _to_iso_date(rec[k])
                else:
                    rec[k] = "" if rec[k] is None else str(rec[k])

            rec["row_hash"] = _row_hash(row)
            payload.append(rec)

        # Supabase/PostgREST upserts work in reasonably sized batches;
        # chunk to avoid oversized requests on a big backfill run.
        CHUNK = 500
        for i in range(0, len(payload), CHUNK):
            chunk = payload[i:i + CHUNK]
            client.table(table_name).upsert(chunk, on_conflict="row_hash").execute()

        print(f"  ✅  Supabase: upserted {len(payload)} row(s) into '{table_name}'")
        return True

    except Exception as e:
        print(f"  ❌  Supabase push failed for '{table_key}': {e}")
        return False


def accumulate_and_save(new_data: dict, master_path: str, run_info: dict,
                         push_enabled: bool = None) -> dict:
    """
    Loads whatever's already saved at master_path, merges today's
    new_data into it (de-duplicated), writes the full cumulative
    workbook back out, and — if enabled — pushes just today's new rows
    to Supabase. Returns the merged dataset.

    push_enabled defaults to the module-level PUSH_TO_SUPABASE flag,
    but main() passes an explicit value after testing the connection,
    so a failed connection test doesn't attempt the push again.
    """
    if push_enabled is None:
        push_enabled = PUSH_TO_SUPABASE

    existing = load_master_data(master_path)
    merged = {}
    added_counts = {}
    for key in _SHEET_KEYS:
        merged[key], added_counts[key] = merge_rows(existing[key], new_data[key])

    run_info["new_investigations"]   = added_counts["investigations"]
    run_info["new_medical_reports"]  = added_counts["medical_reports"]
    run_info["new_mdt"]              = added_counts["mdt"]

    supabase_ok_all = True
    if push_enabled:
        any_pushed = False
        for key in _SHEET_KEYS:
            if added_counts[key]:
                any_pushed = True
                new_slice = merged[key][len(existing[key]):]
                ok = push_rows_to_supabase(key, _SHEET_COLUMNS[key], new_slice)
                supabase_ok_all = supabase_ok_all and ok
        run_info["supabase_status"] = "ok" if (not any_pushed or supabase_ok_all) else "partial failure — see console"
    else:
        run_info["supabase_status"] = "disabled"

    write_excel(merged, master_path, run_info)

    print(f"\n  Merge summary (new rows added this run):")
    for key in _SHEET_KEYS:
        print(f"    {_SHEET_TITLES[key]:<24} +{added_counts[key]}  "
              f"(total now: {len(merged[key])})")

    return merged


# ═══════════════════════════════════════════════════════════════════
# PER-PATIENT EXTRACTION  (one patient page opened ONCE, any number of dates)
# ═══════════════════════════════════════════════════════════════════

def process_patient(session, mr: str, dates: dict, q_nid: str = "") -> tuple:
    """
    Opens ONE patient's chart and extracts everything dated on ANY of the
    days that patient was queued.

    dates : {"dd-mm-yyyy": clinic_on_that_day, ...}   (never empty)

    Returns (summary_rows, investigation_rows, med_report_rows, mdt_rows),
    or None when the patient is not authorised. Each row carries the queue
    date it belongs to as its "Run Date" and the clinic the patient was
    queued under ON THAT DAY. Raises on failure (caller records the error).
    """
    day_set = set(dates)
    summary, invs, reps, mdts = [], [], [], []

    time.sleep(DELAY)
    patient_info = api_patient_get_by_mr(session, mr)
    eng_name = patient_info.get("EngName", "")
    ar_name  = patient_info.get("ARName", "")

    time.sleep(DELAY)
    if not api_authorize_patient(session, mr):
        return None

    time.sleep(DELAY)
    hosp_code, visit_num = api_draw_tree(session, mr)
    if visit_num is None:
        print(f"  ⚠  Could not determine visit number for MR {mr}")
        visit_num = 0
    print(f"  HospCode={hosp_code}  VisitNumber={visit_num}")

    time.sleep(DELAY)
    admission = api_admission(session, mr, hosp_code, visit_num)
    age    = admission.get("PatientAge", patient_info.get("PatientAge", ""))
    sex    = admission.get("Sex",        patient_info.get("Sex", ""))
    id_no  = admission.get("SixFieldDisplayValue", "")
    if q_nid:
        if not str(id_no or "").strip():
            id_no = q_nid
        elif str(id_no).strip() != q_nid:
            print(f"  ⚠  National ID differs: CMIS={id_no}  queue={q_nid} "
                  f"(kept CMIS value) — check MR {mr}")
    birth_dt = admission.get("FirstFieldDisplayValue", "")

    # one summary row per queue day (same patient info, that day's clinic)
    for d in sorted(day_set, key=lambda x: datetime.strptime(x, "%d-%m-%Y")):
        summary.append([mr, dates[d], d, eng_name, ar_name, age, sex,
                        id_no, birth_dt, hosp_code, visit_num])

    time.sleep(DELAY)
    api_patient_note(session, mr, hosp_code, visit_num)      # mandatory call, no data kept

    # Investigations - keep rows whose Order Date is one of the queue days
    time.sleep(DELAY)
    inv_list = api_investigations(session, mr)
    inv_kept = 0
    for inv in inv_list:
        od = _normalize_any_date(inv["Order Date"])
        if od not in day_set:
            continue
        invs.append([mr, dates[od], inv["Status"], inv["Order Type"], inv["Test Name"],
                     inv["Physician Comment"], inv["Order Date"],
                     inv["Requesting Physician"], inv["Schedule Date"],
                     inv["Order Number"], inv["Case Number"], od])
        inv_kept += 1
    print(f"  Investigations   : {inv_kept} on queue day(s) (of {len(inv_list)} on file)")

    # Medical sheets (033 + 02) - same rule
    time.sleep(DELAY)
    target_sheets, token = api_medical_sheets(session, mr, hosp_code, visit_num)
    mr_count = mdt_count = 0
    for sheet in target_sheets:
        time.sleep(DELAY)
        sheet_url = api_get_sheet_url(session, sheet, token)
        time.sleep(DELAY)
        fields = api_get_user_control_data(session, sheet_url, mr, hosp_code, visit_num, sheet)
        creation_date = sheet.get("CreationDate", "")
        code = sheet["SheetCode"].strip()
        if code == "033":
            date_str, description = extract_medical_report(fields, creation_date)
            if date_str in day_set and (description or date_str):
                reps.append([mr, dates[date_str], eng_name, date_str, description, date_str])
                mr_count += 1
        elif code == "02":
            date_str, outcome = extract_mdt(fields, creation_date)
            if date_str in day_set and (outcome or date_str):
                mdts.append([mr, dates[date_str], eng_name, date_str, outcome, date_str])
                mdt_count += 1
    print(f"  Medical Reports  : {mr_count}   MDT Forms : {mdt_count}")
    return summary, invs, reps, mdts


def run_patients(patients: list) -> tuple:
    """patients: [{"mr", "national_id", "dates": {date: clinic}}]. Logs in once."""
    print("\n── Logging in to CMIS MRM …")
    session = build_session()
    try:
        if not login(session):
            sys.exit(1)
    except requests.exceptions.ConnectionError:
        print("\n❌  Cannot reach the server.  Make sure you are connected to the hospital network.")
        sys.exit(1)

    data = {k: [] for k in _SHEET_KEYS}
    errors = []
    total = len(patients)
    for idx, p in enumerate(patients, 1):
        mr, dates = p["mr"], p["dates"]
        print(f"\n{'─' * 65}")
        print(f"  [{idx}/{total}]  MR = {mr}   Days = {', '.join(sorted(dates))}")
        try:
            res = process_patient(session, mr, dates, p.get("national_id", ""))
            if res is None:
                print(f"  ⚠  Not authorized for MR {mr} — skipping")
                errors.append((mr, "Not authorized"))
                continue
            for key, rows in zip(_SHEET_KEYS, res):
                data[key].extend(rows)
        except requests.exceptions.ConnectionError as exc:
            print(f"  ❌  Connection error: {exc}")
            errors.append((mr, f"ConnectionError: {exc}"))
        except requests.exceptions.Timeout:
            print("  ❌  Request timed out")
            errors.append((mr, "Timeout"))
        except Exception as exc:
            print(f"  ❌  Unexpected error: {exc}")
            traceback.print_exc()
            errors.append((mr, str(exc)))
    return data, errors


def _push_new_rows(new_data: dict) -> str:
    """Upsert this run's rows to Supabase -> 'ok' | 'disabled' | 'partial failure ...'."""
    if not PUSH_TO_SUPABASE:
        return "disabled"
    print("\n── Testing Supabase connection …")
    if not test_supabase_connection():
        print("  ⚠  Excel/JSON only for this run.")
        return "disabled"
    ok_all, pushed = True, False
    for key in _SHEET_KEYS:
        rows, _ = merge_rows([], new_data[key])   # drop in-batch duplicates (an upsert batch can't contain the same row_hash twice)
        if rows:
            pushed = True
            ok_all = push_rows_to_supabase(key, _SHEET_COLUMNS[key], rows) and ok_all
    return "ok" if (ok_all or not pushed) else "partial failure — see console"


# ═══════════════════════════════════════════════════════════════════
# MODE A - CHUNK  (used by the parallel GitHub jobs)
#   python script.py --chunk chunk_3.json --chunk-out chunk_3_result.json
# ═══════════════════════════════════════════════════════════════════

def main_chunk(chunk_path: str, out_path: str):
    with open(chunk_path, encoding="utf-8") as f:
        chunk = json.load(f)
    patients = chunk["patients"]
    n_days = sum(len(p["dates"]) for p in patients)

    print("=" * 65)
    print(f"  CMIS MRM — chunk {chunk['chunk']}:  {len(patients)} patient(s), {n_days} patient-day(s)")
    print("=" * 65)

    data, errors = run_patients(patients)
    status = _push_new_rows(data)

    result = {
        "chunk": chunk["chunk"],
        "patients_in_chunk": len(patients),
        "rows": data,
        "errors": [[m, str(w)] for m, w in errors],
        "supabase_status": status,
    }
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, default=str)

    print(f"\n{'=' * 65}")
    print(f"  ✅  Chunk {chunk['chunk']} done: {len({r[0] for r in data['summary']})}/{len(patients)} patients, "
          f"{len(data['investigations'])} investigations, {len(data['medical_reports'])} reports, "
          f"{len(data['mdt'])} MDT, {len(errors)} error(s), supabase={status}")
    print(f"  Result → {out_path}")
    if status.startswith("partial"):
        print("::error::Supabase push failed for this chunk (rows are in the result JSON; re-run is safe).")
        sys.exit(1)


# ═══════════════════════════════════════════════════════════════════
# MODE B - SINGLE DAY, ONE SESSION  (local runs / old behaviour)
#   python script.py [dd-mm-yyyy]
# ═══════════════════════════════════════════════════════════════════

def main():
    ts_tag = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print("=" * 65)
    print("  CMIS MRM — Patient Clinical Data Extractor  (queue-coupled)")
    print(f"  Run date     : {RUN_DATE}")
    print(f"  Master output: {MASTER_OUTPUT_PATH}")
    print("=" * 65)

    print(f"\n── Pulling queue for {RUN_DATE} (source of MR codes + clinic) …")
    try:
        queue_rows = queue_mr_extractor.get_queue_records(RUN_DATE, output_dir=QUEUE_OUTPUT_DIR)
    except Exception as e:
        print(f"❌  Could not pull queue data for {RUN_DATE}: {e}")
        traceback.print_exc()
        sys.exit(1)

    mr_clinic_map, mr_queue_nid = {}, {}
    for qr in queue_rows:
        k = str(qr.get("MR", "")).strip()
        if not k:
            continue
        mr_clinic_map.setdefault(k, qr.get("Clinic", "") or "")
        if qr.get("National ID") and k not in mr_queue_nid:
            mr_queue_nid[k] = str(qr["National ID"]).strip()

    mr_codes = sorted(mr_clinic_map)
    if not mr_codes:
        print(f"❌  No arrived patients found for {RUN_DATE} — nothing to extract.")
        sys.exit(1)
    print(f"  Found {len(mr_codes)} arrived MR code(s) across "
          f"{len(set(mr_clinic_map.values()))} clinic(s)")

    patients = [{"mr": m, "national_id": mr_queue_nid.get(m, ""),
                 "dates": {RUN_DATE: mr_clinic_map[m]}} for m in mr_codes]
    new_data, errors = run_patients(patients)

    push_enabled = PUSH_TO_SUPABASE
    if push_enabled:
        print("\n── Testing Supabase connection …")
        if not test_supabase_connection():
            print("  ⚠  Continuing with Excel-only output for this run.")
            push_enabled = False

    print(f"\n{'=' * 65}\n  Merging into cumulative master workbook …")
    run_info = {"run_date": RUN_DATE, "timestamp": ts_tag,
                "mr_from_queue": len(mr_codes),
                "patients_processed": len({r[0] for r in new_data["summary"]}),
                "errors": len(errors)}
    accumulate_and_save(new_data, MASTER_OUTPUT_PATH, run_info, push_enabled=push_enabled)

    print(f"\n  ✅  Extraction complete for {RUN_DATE}!")
    print(f"  Patients processed : {len(new_data['summary'])}/{len(mr_codes)}")
    print(f"  Investigations     : {len(new_data['investigations'])}")
    print(f"  Medical Reports    : {len(new_data['medical_reports'])}")
    print(f"  MDT Forms          : {len(new_data['mdt'])}")
    for m, why in errors:
        print(f"    MR {m}: {why}")
    print(f"\n  Output → {MASTER_OUTPUT_PATH}")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="CMIS MRM patient data extractor")
    ap.add_argument("run_date", nargs="?", default=None,
                    help="single-day mode: dd-mm-yyyy (default today)")
    ap.add_argument("--chunk", help="chunk mode: path to chunk_<n>.json from plan_chunks.py")
    ap.add_argument("--chunk-out", help="chunk mode: where to write the result JSON")
    args = ap.parse_args()
    if args.chunk:
        if not args.chunk_out:
            ap.error("--chunk requires --chunk-out")
        main_chunk(args.chunk, args.chunk_out)
    else:
        if args.run_date:
            RUN_DATE = args.run_date
        main()
