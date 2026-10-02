"""
SAP Concur US SMB Client Sales — Live SFDC Dashboard
Connects to Salesforce, runs 4 source reports, applies the full calculation
engine, and renders an interactive monthly performance dashboard.

Run:  streamlit run dashboard_sfdc.py
"""

import io, json, re, threading, time as _time, warnings
from calendar import monthrange
from datetime import datetime
from pathlib import Path

import requests as _requests   # used for direct POST to SFDC Analytics API

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG  (must be first Streamlit call)
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SAP Concur | SMB Client Sales Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# BRAND
# ─────────────────────────────────────────────────────────────────────────────
C = dict(blue="#0070F2", light_blue="#4CB1FF", dark_blue="#00144A",
         green="#188918", red="#BB0000", amber="#E8A000", yellow="#F0AB00",
         grey="#E1E2E6", dark_grey="#6A6D73", bg="#FFFFFF", bg_subtle="#F5F6F7")

SEG_COLORS    = {"Key": C["blue"], "Premier": C["light_blue"], "Strategic": C["dark_blue"]}
BUCKET_ORDER  = ["<50%", "50-75%", "75-100%", "100-120%", "120%+"]
BUCKET_COLORS = {"<50%": C["red"], "50-75%": C["amber"], "75-100%": C["yellow"],
                 "100-120%": "#5DB533", "120%+": C["green"]}
MONTH_NAMES   = ["January","February","March","April","May","June",
                 "July","August","September","October","November","December"]

# ─────────────────────────────────────────────────────────────────────────────
# ORG STRUCTURE
# ─────────────────────────────────────────────────────────────────────────────
VP_MAP = {"Key": "Peter Gadd", "Premier": "Jason Rainey", "Strategic": "Andrew Hop"}
LEADER_SEGMENT = {
    "Marissa Mock": "Key",      "Jake Rutenbar": "Key",
    "Ronit Cohn":   "Key",      "Randi Kruger":  "Key",
    "Heather Lewis":"Key",
    "Kyle Loving":  "Premier",  "Christopher Smith": "Premier",
    "Brooke Nelson":"Premier",  "Angie Koplan":  "Premier",
    "Christian Larson":"Strategic", "Amanda Meek":   "Strategic",
    "Megan Frodge": "Strategic",    "Blake Karnes":  "Strategic",
    "Samantha Young":"Strategic",   "Dave Elinger":  "Strategic",
}
NAME_MAP = {
    "Joe Bellefeuille":   "Joseph Bellefeuille",
    "Nathaniel Heussner": "Nate Heussner",
    "Stephen Snediker":   "Steve Snediker",
    "Tom Wahl":           "Thomas Wahl",
    "Chris Smith":        "Christopher Smith",
    "joe bellefeuille":   "Joseph Bellefeuille",
}
DEFAULT_TARGETS = Path(__file__).parent / "targets_2026.xlsx"

# Pro-rated quota periods for mid-year hires and team transfers (FY 2026).
# Derived from GTM Ops Plan Summary Concur.xlsx cross-referenced with targets_2026.xlsx.
# Rule: start day 1–15 → quota starts same month; 16–31 → quota starts following month.
# Format: rep_name (un-normalized) -> [(segment, first_quota_month, last_quota_month), ...]
# Reps with a full Jan 1 assignment and no transfer are NOT listed (no override needed).
PRORATED_PERIODS = {
    # ── Same-team moves (were outside SMB CS before start date) ──────────────
    "Jacob Nickoloff":     [("Key",       2, 12)],   # Jan 20 → starts Feb
    "Adam Sala":           [("Premier",   2, 12)],   # Jan 20 → starts Feb
    "Lindsay Wilson":      [("Premier",   2, 12)],   # Feb  1 → starts Feb
    # ── Cross-segment transfers ───────────────────────────────────────────────
    "Ashley McCue":        [("Strategic", 1,  3),    # Jan–Mar in Strategic
                            ("Premier",   4, 12)],   # Apr–Dec in Premier (moved Apr 1)
    "Jill Desjardine":     [("Key",       1,  5),    # Jan–May in Key
                            ("Strategic", 6, 12)],   # Jun–Dec in Strategic (moved Jun 1)
    "Kylie Barrett":       [("Key",       1,  8),    # Jan–Aug in Key
                            ("Strategic", 9, 12)],   # Sep–Dec in Strategic (moved Sep 1)
    # ── New hires ─────────────────────────────────────────────────────────────
    "Joel Segall":         [("Strategic", 5, 12)],   # May  4 → starts May
    "Nicole Breitenstein": [("Strategic", 5, 12)],   # Apr 20 → starts May
    "Jon Salmon":          [("Premier",   8, 12)],   # Aug  1 → starts Aug
    "Tom Osterberg":       [("Strategic", 9, 12)],   # Sep  1 → starts Sep
}

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
html,body,[class*="css"]{{font-family:'72','Helvetica Neue',Arial,sans-serif;}}
[data-testid="stSidebar"]{{background:{C['bg_subtle']};border-right:1px solid {C['grey']};}}
.kpi-card{{background:{C['bg']};border:1px solid {C['grey']};border-top:4px solid {C['blue']};
  border-radius:8px;padding:18px 20px 14px;text-align:center;
  box-shadow:0 1px 4px rgba(0,0,0,0.06);}}
.kpi-card.green{{border-top-color:{C['green']};}}
.kpi-card.amber{{border-top-color:{C['amber']};}}
.kpi-card.dark{{border-top-color:{C['dark_blue']};}}
.kpi-value{{font-size:26px;font-weight:700;color:{C['blue']};line-height:1.15;}}
.kpi-value.green{{color:{C['green']};}} .kpi-value.amber{{color:{C['amber']};}}
.kpi-value.dark{{color:{C['dark_blue']};}}
.kpi-label{{font-size:12px;color:{C['dark_grey']};margin-top:5px;
  text-transform:uppercase;letter-spacing:.03em;}}
.sh{{font-size:15px;font-weight:600;color:{C['dark_blue']};
  border-bottom:2px solid {C['blue']};padding-bottom:5px;margin:6px 0 10px;}}
</style>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def normalize(name):
    if pd.isna(name): return name
    s = str(name).strip()
    return NAME_MAP.get(s, s)


def build_region_leader_map(cw_df):
    """Derive a live Region → Leader map from the CW ARR dataset.

    For each Owner Region, finds the most recent calendar month that has
    deals and takes the most common 'Opportunity Owner: Manager' in that
    month as the authoritative current leader for that region.

    WHY this works for mid-year transfers:
      - A deal's Owner Region is stamped at close time and never changes.
      - The reps closing deals in a region RIGHT NOW are its current
        members, so their current manager IS the current regional leader.
      - Historical deals from reps who have since moved still carry their
        original region, so they are attributed to whoever manages that
        region today — which is exactly the desired behaviour.

    Priority in run_calc:
      1. region_leader_map[Owner Region]   ← this function
      2. Opportunity Owner: Manager         ← SFDC hierarchy (fallback)
      3. rep_leader_map from targets file   ← start-of-year snapshot (last resort)
    """
    df = cw_df.copy()
    df["_dt"] = pd.to_datetime(df.get("Close Date"), errors="coerce")

    # Prefer 'Oppty Manager' if the bucket field is still present, otherwise
    # use the standard SFDC manager hierarchy column
    mgr_col = ("Oppty Manager" if "Oppty Manager" in df.columns
               else "Opportunity Owner: Manager")

    # Use "Oppty Region" (stable — tied to the opportunity, never changes when
    # the owner transfers).  Fall back to "Oppty Team" if Oppty Region is absent.
    region_col = ("Oppty Region" if "Oppty Region" in df.columns
                  else "Owner Region" if "Owner Region" in df.columns
                  else None)

    if mgr_col not in df.columns or region_col is None:
        return {}

    df = df.dropna(subset=[region_col, "_dt", mgr_col])
    df["_mgr"]    = df[mgr_col].apply(normalize)
    df["_region"] = df[region_col].str.strip()
    df["_period"] = df["_dt"].dt.to_period("M")

    region_leader = {}
    for region, grp in df.groupby("_region"):
        if region in ("", "Unknown", "nan"):
            continue
        # Most recent month with ≥1 deal in this region
        latest = grp["_period"].max()
        recent = grp[grp["_period"] == latest]
        # Most common manager in that month (mode handles multiple reps)
        mode_vals = recent["_mgr"].dropna()
        mode_vals = mode_vals[mode_vals.str.strip().ne("")]
        if len(mode_vals) > 0:
            region_leader[region] = mode_vals.mode().iloc[0]

    return region_leader

def seg_from_team(team):
    if isinstance(team, str):
        for s in ("Key","Premier","Strategic"):
            if s in team: return s
    return "Unknown"

def ltc_rate(term):
    try:
        t = int(float(term))
    except: return 0.0
    if t >= 36: return 0.40
    if t >= 24: return 0.30
    if t >= 12: return 0.20
    return 0.0

def bucket(pct):
    if pct is None or (isinstance(pct, float) and np.isnan(pct)): return "No Quota"
    if pct < 0.50:  return "<50%"
    if pct < 0.75:  return "50-75%"
    if pct < 1.00:  return "75-100%"
    if pct < 1.20:  return "100-120%"
    return "120%+"

def kpi_html(label, value, cls=""):
    return (f"<div class='kpi-card {cls}'>"
            f"<div class='kpi-value {cls}'>{value}</div>"
            f"<div class='kpi-label'>{label}</div></div>")

def fmt_m(v):   return f"${v/1e6:.2f}M"
def fmt_pct(v): return f"{v*100:.1f}%"

# ─────────────────────────────────────────────────────────────────────────────
# SALESFORCE  (lazy import so app loads even without credentials)
# ─────────────────────────────────────────────────────────────────────────────
def sf_connect(username, password, security_token, domain="login"):
    from simple_salesforce import Salesforce
    return Salesforce(username=username, password=password,
                      security_token=security_token, domain=domain)

def sf_connect_session(session_id, instance_url="https://sapconcur.my.salesforce.com"):
    from simple_salesforce import Salesforce
    if not instance_url.startswith("http"):
        instance_url = f"https://{instance_url}"
    return Salesforce(session_id=session_id, instance_url=instance_url)


_meta_cache: dict = {}   # module-level: report_id → full API response dict


def _sf_base_url(sf):
    """Return the REST API base URL from a simple_salesforce connection object."""
    # simple_salesforce v1.x stores it in base_url; fall back to building it
    for attr in ("base_url", "sf_url"):
        val = getattr(sf, attr, None)
        if val and "services/data" in str(val):
            return str(val).rstrip("/") + "/"
    instance = getattr(sf, "sf_instance", "sapconcur.my.salesforce.com")
    version  = getattr(sf, "sf_version",  "57.0")
    return f"https://{instance}/services/data/v{version}/"


def _post_report(sf, report_id, body_dict, params=None):
    """POST to the Salesforce Analytics Reports API using direct requests.

    CRITICAL: we bypass simple_salesforce's restful(data=string) here.
    Using requests.post(json=body_dict) guarantees:
      • body_dict is serialised to JSON by the requests library
      • Content-Type: application/json is set automatically
      • The body actually arrives at Salesforce (simple_salesforce's data=
        path has version-specific quirks that can silently drop the body)

    Returns the parsed JSON response dict, or raises on non-200 status.
    """
    url     = _sf_base_url(sf) + f"analytics/reports/{report_id}"
    headers = {
        "Authorization": f"Bearer {sf.session_id}",
        "Content-Type":  "application/json",
        "Accept":        "application/json",
    }
    resp = _requests.post(
        url,
        json=body_dict,          # requests serialises + sets Content-Type
        headers=headers,
        params=params or {"includeDetails": "true"},
        timeout=60,
    )
    if resp.status_code == 200:
        return resp.json()
    raise Exception(f"SFDC POST HTTP {resp.status_code}: {resp.text[:400]}")


def _post_report_async(sf, report_id, body_dict, params=None):
    """Create an async report instance, poll until complete, return (payload, inst_id).

    Uses /analytics/reports/{id}/instances (async) instead of the sync
    /analytics/reports/{id} endpoint.  Async instances draw from a separate
    1,200/hr rate limit so they do not compete with the 500/hr synchronous
    run quota — critical when multiple users run the app concurrently.

    The returned payload has IDENTICAL JSON structure to a synchronous run
    (same factMap / reportMetadata / allData keys), so all existing parse
    logic works without modification.

    Returns (payload_dict, instance_id).
    Raises RuntimeError on report failure, TimeoutError if > 120 s elapses.
    """
    base_url = _sf_base_url(sf)
    headers  = {
        "Authorization": f"Bearer {sf.session_id}",
        "Content-Type":  "application/json",
        "Accept":        "application/json",
    }
    _params = params or {"includeDetails": "true"}

    # ── Fire the async instance ───────────────────────────────────────────────
    inst_url  = base_url + f"analytics/reports/{report_id}/instances"
    post_resp = _requests.post(
        inst_url,
        json=body_dict or {},
        headers=headers,
        params=_params,
        timeout=30,
    )
    if post_resp.status_code not in (200, 201):
        raise Exception(
            f"SFDC async POST HTTP {post_resp.status_code}: {post_resp.text[:400]}")
    inst_id  = post_resp.json()["id"]
    poll_url = base_url + f"analytics/reports/{report_id}/instances/{inst_id}"

    # ── Poll until Success (up to 120 s) ─────────────────────────────────────
    for _ in range(60):
        _time.sleep(2)
        poll = _requests.get(poll_url, headers=headers, params=_params, timeout=30)
        poll.raise_for_status()
        payload = poll.json()
        status  = payload.get("attributes", {}).get("status", "")
        if status == "Success":
            return payload, inst_id
        if status in ("Error", "Cancelled"):
            err = payload.get("attributes", {}).get("errorCode", "unknown")
            raise RuntimeError(f"Report instance {inst_id} failed: {err}")

    raise TimeoutError(f"Report instance {inst_id} did not complete within 120 s")


def _fetch_all_pages_async(sf, report_id, inst_id, first_result, base_params):
    """Paginate through an async instance's rows using GET startRow.

    Reuses the already-completed instance — zero new instance creation, zero
    extra rate-limit cost.  Each GET fetches the next 2,000 rows from the
    cached instance result.

    Returns the combined list of raw row dicts (all pages).
    """
    combined = list(first_result.get("factMap", {}).get("T!T", {}).get("rows", []))
    current  = first_result
    page     = 1
    MAX_PAGES = 15      # 15 × 2,000 = 30,000 row ceiling

    base_url = _sf_base_url(sf)
    headers  = {"Authorization": f"Bearer {sf.session_id}", "Accept": "application/json"}
    poll_url = base_url + f"analytics/reports/{report_id}/instances/{inst_id}"

    while not current.get("allData", True) and page < MAX_PAGES:
        pg_params = dict(base_params)
        pg_params["startRow"] = page * 2000
        try:
            resp = _requests.get(poll_url, headers=headers, params=pg_params, timeout=30)
            resp.raise_for_status()
            current = resp.json()
            combined.extend(current.get("factMap", {}).get("T!T", {}).get("rows", []))
            page += 1
        except Exception:
            break

    return combined


def _sf_get_meta(sf, report_id):
    """Fetch full report metadata (cached per module lifetime).

    Caches the COMPLETE response so we can inspect both reportMetadata
    and reportExtendedMetadata (needed for column label → API name mapping
    and for the diagnostic display).
    """
    if report_id not in _meta_cache:
        try:
            r = sf.restful(
                path=f"analytics/reports/{report_id}",
                method="GET",
                params={"includeDetails": "false"},
            )
            _meta_cache[report_id] = r          # store full response
        except Exception:
            _meta_cache[report_id] = {}
    return _meta_cache[report_id]


def _fetch_all_pages(sf, report_id, first_result, body_dict, base_params):
    """Paginate through a capped Salesforce Analytics API response.

    The Analytics API returns at most 2,000 rows per call.  When allData is
    False in the response, additional rows are available via the startRow
    query parameter.  We fetch them in 2,000-row increments and merge into
    the first call's row list.

    body_dict — the POST body that succeeded (None → use GET for subsequent pages)
    base_params — query-string params used on the first call

    Returns the combined list of raw row dicts (all pages).
    """
    combined = list(first_result.get("factMap", {}).get("T!T", {}).get("rows", []))
    current  = first_result
    page     = 1
    MAX_PAGES = 15          # 15 × 2,000 = 30,000 row ceiling

    while not current.get("allData", True) and page < MAX_PAGES:
        pg_params = dict(base_params)
        pg_params["startRow"] = page * 2000
        try:
            if body_dict is not None:
                current = _post_report(sf, report_id, body_dict, pg_params)
            else:
                current = sf.restful(
                    path=f"analytics/reports/{report_id}",
                    method="GET",
                    params=pg_params,
                )
            combined.extend(current.get("factMap", {}).get("T!T", {}).get("rows", []))
            page += 1
        except Exception:
            break           # stop pagination on error; return what we have so far

    return combined


def sf_run_report(sf, report_id, start_date=None, end_date=None):
    """Run a Salesforce Analytics report and return (DataFrame, row_count).

    Two POST attempts followed by a GET fallback.

      Attempt 1 — minimal standardDateFilter on CLOSE_DATE
        Fast path.  Works for LTC, Complete, Retention (no BucketFields).
        Skipped immediately if the report contains a BucketField (HTTP 400).
        Accepted if response < 2,000 rows.

      Attempt 2 — FULL saved-metadata POST, date patched in-place  ← KEY FIX
        Fetches the complete saved reportMetadata (including BucketField
        definitions, bucketFields list, reportBooleanFilter, etc.) via GET,
        deep-copies it, then patches ONLY the standardDateFilter and strips
        any existing date entries from reportFilters.

        WHY this is necessary:
        The Salesforce Analytics API requires that any field present in the
        report's saved metadata also be present in the POST body.  Reports
        that contain Bucket Fields (custom grouping fields) fail with
        HTTP 400 "Invalid value specified: BucketField_XXXXXXX" when a
        minimal POST body is sent — because the BucketField definition is
        missing.  By deep-copying the full saved metadata and only changing
        the date, we preserve every required field and the POST succeeds.

        Accepted if response < 2,000 rows.

      GET fallback — only if both POSTs throw an exception.

    All attempt results are logged to st.session_state.sf_post_debug.
    """
    import copy as _copy

    params = {"includeDetails": "true"}
    result = None
    debug  = []
    _winning_body    = None   # POST body that produced result (None → GET)
    _winning_params  = params  # query params used for winning call
    _winning_inst_id = None   # async instance id (None → sync GET fallback)

    if "sf_post_debug" not in st.session_state:
        st.session_state.sf_post_debug = {}

    # Date-field API names to strip when rebuilding reportFilters
    _DATE_COLS = {
        "CLOSE_DATE", "CLOSEDATE", "CLOSED_DATE",
        "CREATEDDATE", "LASTMODIFIEDDATE",
        "CLOSE_MONTH", "CLOSEMONTH",
    }

    if start_date and end_date:

        # ── Attempt 1: minimal standardDateFilter (fast path, no metadata call)
        _a1_bucket_err = False
        try:
            _a1_body = {
                "reportMetadata": {
                    "standardDateFilter": {
                        "column":        "CLOSE_DATE",
                        "durationValue": "CUSTOM",
                        "startDate":     start_date,
                        "endDate":       end_date,
                    }
                }
            }
            r1, _a1_inst_id = _post_report_async(sf, report_id, _a1_body, params)
            n1 = len(r1.get("factMap", {}).get("T!T", {}).get("rows", []))
            all1 = r1.get("allData", True)
            result = r1
            _winning_body    = _a1_body
            _winning_params  = params
            _winning_inst_id = _a1_inst_id
            debug.append(
                f"A1(stdDateFilter/CLOSE_DATE) OK → {n1} rows"
                + ("" if all1 else " [paginating…]"))
        except Exception as e1:
            err_str = str(e1)
            if "BucketField" in err_str:
                _a1_bucket_err = True
                debug.append(
                    f"A1 ERR: BucketField in report — minimal POST rejected; "
                    f"switching to full-metadata A2")
            else:
                debug.append(f"A1 ERR: {err_str}")

        # ── Attempt 2: FULL saved-metadata POST, only date patched ───────────
        if result is None:
            try:
                full_resp  = _sf_get_meta(sf, report_id)
                saved_meta = full_resp.get("reportMetadata", {})
                std_info   = saved_meta.get("standardDateFilter") or {}
                std_col    = std_info.get("column", "CLOSE_DATE")

                # Deep-copy full saved metadata — preserve BucketFields etc.
                patched_meta = _copy.deepcopy(saved_meta)

                # Patch only the date scope
                patched_meta["standardDateFilter"] = {
                    "column":        std_col,
                    "durationValue": "CUSTOM",
                    "startDate":     start_date,
                    "endDate":       end_date,
                }

                # Strip existing date entries from reportFilters; add our range
                existing_filters = patched_meta.get("reportFilters") or []
                non_date = [
                    f for f in existing_filters
                    if f.get("column", "").upper() not in _DATE_COLS
                ]
                patched_meta["reportFilters"] = non_date + [
                    {"column": std_col, "operator": "greaterOrEqual",
                     "value": start_date},
                    {"column": std_col, "operator": "lessOrEqual",
                     "value": end_date},
                ]

                _a2_body = {"reportMetadata": patched_meta}
                r2, _a2_inst_id = _post_report_async(sf, report_id, _a2_body, params)
                n2 = len(r2.get("factMap", {}).get("T!T", {}).get("rows", []))
                all2 = r2.get("allData", True)
                result = r2
                _winning_body    = _a2_body
                _winning_params  = params
                _winning_inst_id = _a2_inst_id
                debug.append(
                    f"A2(full-meta/col={std_col}) OK → {n2} rows"
                    + ("" if all2 else " [paginating…]"))
            except Exception as e2:
                debug.append(f"A2 ERR: {e2}")

    # ── GET fallback — only if both async POSTs threw exceptions ────────────
    if result is None:
        result = sf.restful(
            path=f"analytics/reports/{report_id}",
            method="GET",
            params=params,
        )
        nG = len(result.get("factMap", {}).get("T!T", {}).get("rows", []))
        all_g = result.get("allData", True)
        _winning_body    = None
        _winning_params  = params
        _winning_inst_id = None
        debug.append(f"GET-fallback → {nG} rows" + ("" if all_g else " [paginating…]"))

    # ── Pagination: fetch remaining pages if allData is False ────────────────
    # Async path: reuse the completed instance via GET startRow (no new instance).
    # Sync fallback path: use the original _fetch_all_pages helper.
    if _winning_inst_id:
        all_rows = _fetch_all_pages_async(
            sf, report_id, _winning_inst_id, result, _winning_params)
    else:
        all_rows = _fetch_all_pages(
            sf, report_id, result, _winning_body, _winning_params)
    if len(all_rows) > len(result.get("factMap", {}).get("T!T", {}).get("rows", [])):
        debug.append(f"pagination complete → {len(all_rows)} total rows")

    st.session_state.sf_post_debug[report_id] = " | ".join(debug)

    meta   = result.get("reportMetadata", {})
    ext    = result.get("reportExtendedMetadata", {})
    cols   = meta.get("detailColumns", [])
    cinfo  = ext.get("detailColumnInfo", {})
    labels = [cinfo.get(c, {}).get("label", c) for c in cols]
    records = []
    for row in all_rows:
        cells = row.get("dataCells", [])
        records.append({labels[i]: _cell_val(cells[i])
                        for i in range(min(len(labels), len(cells)))})
    df = pd.DataFrame(records)
    return df, len(records)


def sf_run_report_multi(sf, report_id, month_nums, year, full_year=False):
    """Run a report one month at a time and concatenate results.

    Makes len(month_nums) POST calls, each scoped to a single calendar month.
    Falls back to GET per month if POST is rejected.

    full_year=True: makes a SINGLE call spanning the entire year (Jan 1 – Dec 31).
    Use this for reports whose standard date field is a text/formula picklist
    (e.g. Retention's "Final Month Closed") so that the API cannot filter by month.
    A full-year call ensures standalone deals (prior-month close / current-month
    recognition) are not missed, while excluding true prior-year rows if the API
    filter works.  Python's FMC filter then scopes to the correct months.

    IMPORTANT — deduplication:
    If the Salesforce report's date filter cannot be overridden via the API
    (e.g. CW ARR which ignores both standardDateFilter and reportFilters),
    every month call returns the same full-year dataset.  Concatenating N months
    would produce N identical copies, inflating YTD totals by N×.
    We deduplicate on all columns after concat to collapse those copies back to
    one unique dataset.  The calc engine's own date filter then scopes to the
    correct months.  This is safe for reports where the API filter DOES work
    (each month returns distinct rows — dedup changes nothing in that case).

    Returns (combined DataFrame, total row count).
    """
    if full_year:
        start = f"{year}-01-01"
        end   = f"{year}-12-31"
        df, _ = sf_run_report(sf, report_id, start, end)
        combined = df.drop_duplicates()
        return combined, len(combined)

    dfs = []
    for m in month_nums:
        start    = f"{year}-{m:02d}-01"
        last_day = monthrange(year, m)[1]
        end      = f"{year}-{m:02d}-{last_day:02d}"
        df, _    = sf_run_report(sf, report_id, start, end)
        if len(df) > 0:
            dfs.append(df)
    if not dfs:
        return pd.DataFrame(), 0
    combined = pd.concat(dfs, ignore_index=True)
    before   = len(combined)
    combined = combined.drop_duplicates()
    after    = len(combined)
    if before != after:
        # Store for diagnostic display
        if "sf_dedup_info" not in st.session_state:
            st.session_state.sf_dedup_info = {}
        st.session_state.sf_dedup_info[report_id] = (
            f"Deduped {before} → {after} rows "
            f"(API date filter override not working for this report — "
            f"each month call returned the same dataset)"
        )
    return combined, len(combined)


# ─────────────────────────────────────────────────────────────────────────────
# CONCURRENCY CONTROL + SHARED REPORT CACHE
# ─────────────────────────────────────────────────────────────────────────────
# _sfdc_cache     : in-memory result store, shared across all user sessions
# _sfdc_key_locks : one Lock per cache key; different periods don't block each
#                   other but concurrent requests for the same period serialize
# _sfdc_meta_lock : protects the key-lock registry itself
_sfdc_cache      = {}
_sfdc_key_locks  = {}
_sfdc_meta_lock  = threading.Lock()
_CACHE_TTL       = 1800              # seconds (30 minutes)


def _cache_get(key):
    """Return cached result if present and not expired, else None."""
    entry = _sfdc_cache.get(key)
    if entry and (_time.time() - entry[0]) < _CACHE_TTL:
        return entry[1]
    return None


def _key_lock(key):
    """Return (creating if needed) the per-key Lock for this cache entry."""
    with _sfdc_meta_lock:
        if key not in _sfdc_key_locks:
            _sfdc_key_locks[key] = threading.Lock()
        return _sfdc_key_locks[key]


def fetch_all_reports(_sf, rpt_cw, rpt_ltc, rpt_ret, rpt_comp,
                      month_nums_tuple, year):
    """Fetch all 4 SFDC reports using double-checked locking.

    Fast path  (cache warm):
        Returns in milliseconds.  No SFDC call.  No lock acquired.

    Slow path  (cache cold or expired):
        Acquires a per-key lock.  Any other thread requesting the same period
        blocks here.  When it unblocks, the second cache check finds the result
        already populated and returns it without touching Salesforce.

    Result: exactly one SFDC fetch per period per 30-minute window,
    regardless of how many users click Run Reports simultaneously or how
    closely together they click.
    """
    key = (rpt_cw, rpt_ltc, rpt_ret, rpt_comp, month_nums_tuple, year)

    # ── First check — no lock, instant ──────────────────────────────────────
    result = _cache_get(key)
    if result is not None:
        return result

    # ── Cache miss: serialize on the per-key lock ────────────────────────────
    with _key_lock(key):
        # Second check: the previous lock-holder may have just finished fetching.
        result = _cache_get(key)
        if result is not None:
            return result                   # warm — skip fetch entirely

        # This thread is the designated fetcher for this period.
        month_nums = list(month_nums_tuple)
        cw_df,  n1 = sf_run_report_multi(_sf, rpt_cw,   month_nums, year)
        ltc_df, n2 = sf_run_report_multi(_sf, rpt_ltc,  month_nums, year)
        # Retention: full-year call so standalone deals (prior-month close /
        # current-month FMC recognition) are not dropped by a tight range.
        ret_df, n3 = sf_run_report_multi(_sf, rpt_ret,  month_nums, year,
                                         full_year=True)
        comp_df,n4 = sf_run_report_multi(_sf, rpt_comp, month_nums, year)
        result = (cw_df, n1, ltc_df, n2, ret_df, n3, comp_df, n4)
        _sfdc_cache[key] = (_time.time(), result)

    return result


def _sfdc_error_msg(e):
    """Return a user-friendly string for common Salesforce API errors."""
    s = str(e)
    if any(x in s for x in ["INVALID_SESSION_ID", "INVALID_AUTH_HEADER",
                              "Session expired", "expired session"]):
        return ("session_expired",
                "Your Salesforce session has expired. "
                "Paste a fresh Session ID in the sidebar and click Connect.")
    if any(x in s for x in ["EXCEEDED_MAX_CONCURRENT", "REQUEST_LIMIT_EXCEEDED",
                              "TXN_SECURITY_METERING", "concurrent"]):
        return ("rate_limit",
                "Salesforce is handling too many requests right now. "
                "Wait 30 seconds and click Run Reports again.")
    return ("other", f"Salesforce error: {e}")


def _cell_val(cell):
    """Extract a plain Python value from a Salesforce Analytics API data cell.
    Handles plain values, currency dicts, and any nested OrderedDict.
    For User lookup fields the API returns the record ID in 'value' and the
    display name in 'label'; we detect that case and return the label instead."""
    val = cell.get("value")
    if isinstance(val, dict):
        # Currency / compound field: {"amount": 27467.28, "currency": "USD"}
        for key in ("amount", "value", "number"):
            if key in val:
                return val[key]
        return cell.get("label")
    # Salesforce User / record IDs are 15- or 18-char alphanumeric strings with
    # well-known key prefixes (005 = User, 003 = Contact, 001 = Account, etc.)
    if isinstance(val, str) and len(val) in (15, 18) and val[:3] in (
            "005", "003", "001", "006", "00T", "00U"):
        label = cell.get("label")
        if label:
            return label
    return val

# ─────────────────────────────────────────────────────────────────────────────
# QUOTA LOADING
# ─────────────────────────────────────────────────────────────────────────────
QUARTER_OF = {1:("Q1",0),2:("Q1",1),3:("Q1",2),
              4:("Q2",0),5:("Q2",1),6:("Q2",2),
              7:("Q3",0),8:("Q3",1),9:("Q3",2),
              10:("Q4",0),11:("Q4",1),12:("Q4",2)}
QLIN_COL   = {"Q1":2,"Q2":3,"Q3":4,"Q4":5}
MLIN_COL   = {0:6, 1:7, 2:8}

# ─────────────────────────────────────────────────────────────────────────────
# PRO-RATED QUOTA OVERRIDES
# ─────────────────────────────────────────────────────────────────────────────
_ALL_MONTHS = ["January","February","March","April","May","June",
               "July","August","September","October","November","December"]


def build_quota_overrides(tgt_bytes, month_nums):
    """Return rep_quota_override dict: rep_name (normalized) -> period quota $.

    Uses PRORATED_PERIODS (hardcoded above) to determine which segment monthly
    amounts to include for each affected rep.  For each rep in PRORATED_PERIODS
    the function sums the CSE monthly dollar quota for every (segment, month)
    pair where the month falls within one of the rep's active periods AND within
    month_nums.  All other reps continue to use IC_Quota * monthly_factor.
    """
    if isinstance(month_nums, int):
        month_nums = [month_nums]

    if tgt_bytes is None or not PRORATED_PERIODS:
        return {}

    # ── CSE per-segment monthly dollar quotas from targets file ──────────────
    cse_buf = io.BytesIO(tgt_bytes)
    cse_raw = pd.read_excel(cse_buf, sheet_name="CSE Quotas with Linearity", header=None)
    # Row 4 = header with month names; rows 5/6/7 = Key/Strategic/Premier
    hdr = cse_raw.iloc[4].tolist()
    try:
        mid = {m: hdr.index(m) for m in _ALL_MONTHS}
    except ValueError:
        return {}   # sheet layout changed — skip overrides gracefully

    seg_monthly = {}
    for seg, ridx in [("Key", 5), ("Strategic", 6), ("Premier", 7)]:
        row = cse_raw.iloc[ridx]
        seg_monthly[seg] = {m: float(row.iloc[mid[m]]) for m in _ALL_MONTHS}

    # ── Build override per rep ────────────────────────────────────────────────
    override = {}
    for raw_name, periods in PRORATED_PERIODS.items():
        nm = normalize(raw_name)
        period_q = 0.0
        for (seg, mo_start, mo_end) in periods:
            if seg not in seg_monthly:
                continue
            for m in month_nums:
                if mo_start <= m <= mo_end:
                    period_q += seg_monthly[seg][_ALL_MONTHS[m - 1]]
        override[nm] = period_q

    return override

def load_quotas(tgt_bytes, month_nums):
    """Return (ldr_quota_map, rep_quota_map, rep_leader_map, rep_team_map,
               seg_pl_map, monthly_factor).
    month_nums: int or list of ints — quotas are summed across all months."""
    if isinstance(month_nums, int):
        month_nums = [month_nums]
    buf = io.BytesIO(tgt_bytes)

    # ── Leader Quotas with Linearity sheet ──────────────────────────────────
    ldr_raw = pd.read_excel(buf, sheet_name="Leader Quotas with Linearity", header=None)

    # Sum monthly_factor across all months in the period
    monthly_factor = sum(
        float(ldr_raw.iloc[1, QLIN_COL[QUARTER_OF[m][0]]]) *
        float(ldr_raw.iloc[1, MLIN_COL[QUARTER_OF[m][1]]])
        for m in month_nums
    )

    # Leader quota: sum of monthly quota columns for each month (col 7=Jan … 18=Dec)
    leader_quota_map = {}
    for _, row in ldr_raw.iloc[5:24].iterrows():
        name = normalize(str(row.iloc[1]).strip()) if not pd.isna(row.iloc[1]) else ""
        try:
            q = sum(float(row.iloc[6 + m]) for m in month_nums
                    if not pd.isna(row.iloc[6 + m]))
        except: q = 0.0
        if name and name not in ("nan",""):
            leader_quota_map[name] = q

    # P&L rows (VP + Tim Downs total) — sum across months
    pl_by_name = {}
    for _, row in ldr_raw.iloc[31:35].iterrows():
        name = normalize(str(row.iloc[1]).strip()) if not pd.isna(row.iloc[1]) else ""
        try:
            q = sum(float(row.iloc[6 + m]) for m in month_nums
                    if not pd.isna(row.iloc[6 + m]))
        except: q = 0.0
        if name: pl_by_name[name] = q

    seg_pl = {
        "Key":       pl_by_name.get("Peter Gadd",  0),
        "Premier":   pl_by_name.get("Jason Rainey", 0),
        "Strategic": pl_by_name.get("Andrew Hop",   0),
        "Total":     pl_by_name.get("Tim Downs",    0),
    }

    # ── Individual Quotas sheet ──────────────────────────────────────────────
    buf.seek(0)
    indiv = pd.read_excel(buf, sheet_name="Individual Quotas")
    indiv["Rep/RSD Name"] = indiv["Rep/RSD Name"].apply(normalize)
    indiv["Leader Name"]  = indiv["Leader Name"].apply(normalize)

    rep_quota_map  = {}
    rep_leader_map = {}
    rep_team_map   = {}
    for _, row in indiv[indiv["IC Quota"].notna()].iterrows():
        rn = str(row["Rep/RSD Name"]).strip()
        if not rn or rn == "nan": continue
        rep_quota_map[rn]  = float(row["IC Quota"])
        rep_leader_map[rn] = str(row["Leader Name"]).strip()
        rep_team_map[rn]   = str(row["Team"]).strip()

    return leader_quota_map, rep_quota_map, rep_leader_map, rep_team_map, seg_pl, monthly_factor

# ─────────────────────────────────────────────────────────────────────────────
# CALCULATION ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def _assign_leader(df, region_leader_map, mgr_col, rep_leader_map):
    """Vectorized 4-level leader assignment for a deal-level DataFrame.

    Priority (highest to lowest):
      1. Oppty Region  → region_leader_map   (stable — tied to the opp)
      2. Oppty Team    → region_leader_map   (fallback if Oppty Region absent)
      3. mgr_col       → Opportunity Owner: Manager / Oppty Manager
      4. rep_leader_map from targets file    (start-of-year snapshot)

    By applying in reverse priority order (lowest first, higher overwrites),
    each deal ends up with the most reliable leader available.
    """
    _bad = {"", "nan", "none", "unknown", "n/a"}

    def _clean(s):
        v = normalize(str(s)) if pd.notna(s) else ""
        return "" if str(v).lower().strip() in _bad else str(v).strip()

    # Base: targets file (lowest priority)
    leader = df["Opportunity Owner"].apply(
        lambda r: rep_leader_map.get(_clean(r), "Unknown"))

    # Layer 3: SFDC manager hierarchy
    if mgr_col in df.columns:
        mgr = df[mgr_col].apply(_clean)
        mask = mgr.ne("")
        leader = leader.where(~mask, mgr)

    if region_leader_map:
        # Layer 2: Oppty Team → region_leader_map
        if "Oppty Team" in df.columns:
            team_ldr = df["Oppty Team"].apply(
                lambda t: _clean(region_leader_map.get(str(t).strip(), "")))
            mask = team_ldr.ne("")
            leader = leader.where(~mask, team_ldr)

        # Layer 1: Oppty Region → region_leader_map (highest priority)
        for rcol in ("Oppty Region", "Owner Region"):
            if rcol in df.columns:
                reg_ldr = df[rcol].apply(
                    lambda r: _clean(region_leader_map.get(str(r).strip(), "")))
                mask = reg_ldr.ne("")
                leader = leader.where(~mask, reg_ldr)
                break   # use the first region column found

    return leader


def run_calc(cw_raw, ltc_raw, ret_raw, comp_raw,
             ldr_quota_map, rep_quota_map, rep_leader_map, rep_team_map,
             seg_pl, month_nums, year, monthly_factor=0.066687,
             region_leader_map=None, rep_quota_override=None):
    if rep_quota_override is None:
        rep_quota_override = {}
    if isinstance(month_nums, int):
        month_nums = [month_nums]
    month_names = [MONTH_NAMES[m - 1] for m in month_nums]

    # ── Filter by period ─────────────────────────────────────────────────────
    cw = cw_raw.copy()
    # Detect 18-char opportunity ID column. SFDC reports may export both
    # "Opportunity ID" (15-char standard field) and "ID (18 Char)" (18-char).
    # Using "ID (18 Char)" when present avoids duplicate-column issues.
    _cw_id_col = ("ID (18 Char)" if "ID (18 Char)" in cw.columns
                  else "Opportunity ID" if "Opportunity ID" in cw.columns
                  else None)
    # Guard: if CW report returns 0 rows (column-less DataFrame), ensure
    # required columns exist so downstream master-dataset logic completes cleanly.
    if "Close Date" in cw.columns:
        cw["Close Date"] = pd.to_datetime(cw["Close Date"], errors="coerce")
        cw = cw[(cw["Close Date"].dt.month.isin(month_nums)) &
                (cw["Close Date"].dt.year  == year)].copy()
    if "Opportunity Owner" in cw.columns:
        cw["Opportunity Owner"] = cw["Opportunity Owner"].apply(normalize)
    mgr_col_cw = "Oppty Manager" if "Oppty Manager" in cw.columns else "Opportunity Owner: Manager"
    if mgr_col_cw in cw.columns:
        cw[mgr_col_cw] = cw[mgr_col_cw].apply(normalize)
    for _col in ["Opportunity Owner", "Opportunity Name", "Close Date"]:
        if _col not in cw.columns:
            cw[_col] = pd.Series(dtype=object)

    ltc = ltc_raw.copy()
    _ltc_id_col = ("ID (18 Char)" if "ID (18 Char)" in ltc.columns
                   else "Opportunity ID" if "Opportunity ID" in ltc.columns
                   else None)
    # Guard: LTC report may return 0 rows (e.g. SFDC date filter active), producing
    # a column-less DataFrame.  Accessing ltc["Close Date"] on that would raise
    # KeyError and crash the entire calc run.  Mirror the same guard used for comp.
    if "Close Date" in ltc.columns:
        ltc["Close Date"] = pd.to_datetime(ltc["Close Date"], errors="coerce")
        ltc = ltc[(ltc["Close Date"].dt.month.isin(month_nums)) &
                  (ltc["Close Date"].dt.year  == year)].copy()
    if "Opportunity Owner" in ltc.columns:
        ltc["Opportunity Owner"] = ltc["Opportunity Owner"].apply(normalize)
    mgr_col_ltc = "Oppty Manager" if "Oppty Manager" in ltc.columns else "Opportunity Owner: Manager"
    if mgr_col_ltc in ltc.columns:
        ltc[mgr_col_ltc] = ltc[mgr_col_ltc].apply(normalize)

    ret = ret_raw.copy()
    # Guard: if Retention report returns 0 rows (column-less DataFrame), pre-populate
    # all columns the downstream logic touches so no KeyError is raised.
    # The FMC filter, groupby, and merge will all produce empty results cleanly.
    for _col in ["Final Month Closed", "Close Date", "Opportunity Owner",
                 "Opportunity Name", "Opportunity ID", "BMI Sales ARR",
                 "Roll-up Sales Credit Calculation (converted)",
                 "Oppty Region", "Oppty Team", "Opportunity Owner: Manager", "ARR Disputes"]:
        if _col not in ret.columns:
            ret[_col] = pd.Series(dtype=object)
    # ── Year-aware Final Month Closed filter ─────────────────────────────────
    # The Retention SFDC report often has no standard date field (Final Month
    # Closed is a text picklist), so the API may return ALL-TIME records.
    # The field stores just the month name ("June", no year) on some orgs.
    # Without a year check, isin(["June"]) would match June 2025 AND June 2026,
    # inflating retention by ~$300k.  We accept four formats:
    #   1. "June 2026"  — explicit month-year (most reliable)
    #   2. "June"       — month only, validated against Close Date year
    #   3. "Jun 2026" / "Jun-2026" — short-month variants
    #   4. blank/null   — Fall back to Close Date month+year.  Some SFDC orgs
    #                     leave Final Month Closed unpopulated; without this
    #                     fallback those records are silently excluded.
    _fmc = ret["Final Month Closed"].fillna("").astype(str).str.strip()
    _month_year_explicit = [f"{mn} {year}" for mn in month_names]          # ["June 2026"]
    _month_year_short    = [f"{mn[:3]} {year}" for mn in month_names]       # ["Jun 2026"]
    _month_year_hyphen   = [f"{mn[:3]}-{year}" for mn in month_names]       # ["Jun-2026"]
    _month_only          = month_names                                        # ["June"]
    # For month-only values ("June"), cross-validate against Close Date year when
    # the column is present.  Using ret.get() on a DataFrame returns a COLUMN
    # (Series) if it exists, or a default.  Using a zero-length default Series
    # caused all rows to be dropped.  Guard with an explicit column check instead.
    _match_explicit = _fmc.isin(_month_year_explicit + _month_year_short + _month_year_hyphen)
    if "Close Date" in ret.columns:
        _close_dt  = pd.to_datetime(ret["Close Date"], errors="coerce")
        _close_yr  = _close_dt.dt.year
        _close_mo  = _close_dt.dt.month
        _match_month_only = _fmc.isin(_month_only) & _close_yr.eq(year).fillna(False)
        # Fallback: blank Final Month Closed → use Close Date month and year
        _match_blank_fmc  = (_fmc.eq("") &
                              _close_yr.eq(year).fillna(False) &
                              _close_mo.isin(month_nums).fillna(False))
    else:
        # Close Date not in the Retention report — cannot do year cross-validation.
        # The API was called with a full-year (Jan–Dec year) range, so any rows
        # returned with a prior-year close date should already be excluded by the
        # API filter.  Accept all FMC-matching records.
        # To enable strict year validation, add a "Close Date" column to the
        # Salesforce Retention report and re-run.
        _match_month_only = _fmc.isin(_month_only)
        _match_blank_fmc  = pd.Series(False, index=ret.index)
    ret = ret[_match_explicit | _match_month_only | _match_blank_fmc].copy()
    ret["Opportunity Owner"] = ret["Opportunity Owner"].apply(normalize)
    mgr_col_ret = "Oppty Manager" if "Oppty Manager" in ret.columns else "Opportunity Owner: Manager"
    ret[mgr_col_ret] = ret[mgr_col_ret].apply(normalize)
    ret["BMI_ARR"] = pd.to_numeric(ret["BMI Sales ARR"], errors="coerce").fillna(0)
    # Retention incentive formula (per comp team): MAX(0, BMI Sales ARR − Forecast Amount).
    # Resolve Forecast Amount: check the retention file first (it may already include the
    # column); if absent, join from cw_raw by Opportunity ID so standalone retention deals
    # that have no matching CW row still get a value where possible.
    _fa_ret = next((c for c in ret.columns if c.lower().startswith("forecast amount")), None)
    if _fa_ret:
        _ret_fa_vals = pd.to_numeric(ret[_fa_ret], errors="coerce").fillna(0)
    elif "Opportunity ID" in ret.columns and "Opportunity ID" in cw_raw.columns:
        _fa_cw_col = next(
            (c for c in cw_raw.columns if c.lower().startswith("forecast amount")), None)
        if _fa_cw_col:
            _fa_map = dict(zip(
                cw_raw["Opportunity ID"],
                pd.to_numeric(cw_raw[_fa_cw_col], errors="coerce")))
            _ret_fa_vals = ret["Opportunity ID"].map(_fa_map).fillna(0)
        else:
            _ret_fa_vals = pd.Series(0.0, index=ret.index)
    else:
        _ret_fa_vals = pd.Series(0.0, index=ret.index)
    ret["Roll_SC_Ret"] = _ret_fa_vals   # alias kept so groupby SC_SUM works unchanged
    split_mask = ret.get("ARR Disputes", pd.Series([""] * len(ret))).fillna("")
    split_mask = split_mask.str.contains("Split Opportunity", case=False, na=False)
    ret = ret[~split_mask].copy()

    comp = comp_raw.copy()
    # ── Close Month filter ──────────────────────────────────────────────────
    # Close Month is a date formula field; the Analytics API returns ISO format
    # ("2026-06-01") or parseable text ("June 2026").  Parse with pd.to_datetime
    # and filter on year + month.  The broken v16 approach used str.contains on
    # the digit "6" which matched "2026" itself, letting all 12 months of 2026
    # through and inflating the Complete total.
    #
    # Guard: if the report returned 0 rows the DataFrame has no columns at all.
    # Accessing comp["Close Month"] would raise KeyError and crash the entire
    # run.  Detect this early and substitute an empty frame so the engine
    # continues with Complete credit = $0 for the period.
    if comp.empty or "Close Month" not in comp.columns:
        comp = pd.DataFrame(columns=["Opportunity Owner", "Opportunity Name",
                                      "Complete_Credit_Val"])
        comp["Complete_Credit_Val"] = pd.Series(dtype=float)
    else:
        comp["_cm_parsed"] = pd.to_datetime(
            comp["Close Month"].astype(str).str.strip(), errors="coerce", dayfirst=False)
        comp = comp[(comp["_cm_parsed"].dt.year  == year) &
                    (comp["_cm_parsed"].dt.month.isin(month_nums))].copy()
        comp.drop(columns=["_cm_parsed"], inplace=True)
        comp["Opportunity Owner"] = comp["Opportunity Owner"].apply(normalize)
        comp["Complete_Credit_Val"] = pd.to_numeric(
            comp["Roll-up Sales Credit Calculation (converted)"], errors="coerce").fillna(0)

    # ── Lookup tables ────────────────────────────────────────────────────────
    complete_names  = set(comp["Opportunity Name"].str.strip())
    complete_lookup = dict(zip(comp["Opportunity Name"].str.strip(), comp["Complete_Credit_Val"]))

    _fa_ltc = ("Forecast Amount (converted)" if "Forecast Amount (converted)" in ltc.columns
               else "Forecast Amount")
    if _ltc_id_col:
        ltc["LTC_Uplift_Calc"] = ltc.apply(
            lambda r: pd.to_numeric(r[_fa_ltc], errors="coerce") * ltc_rate(r["Term (no. of months)"]), axis=1)
        ltc_lookup = dict(zip(ltc[_ltc_id_col].astype(str).str.strip(), ltc["LTC_Uplift_Calc"]))
        _ltc_id_based = True
    elif "Opportunity Name" in ltc.columns:
        # Fallback: no ID column in LTC report — match by name (v4 behaviour)
        ltc["LTC_Uplift_Calc"] = ltc.apply(
            lambda r: pd.to_numeric(r[_fa_ltc], errors="coerce") * ltc_rate(r["Term (no. of months)"]), axis=1)
        ltc_lookup = dict(zip(ltc["Opportunity Name"].str.strip(), ltc["LTC_Uplift_Calc"]))
        _ltc_id_based = False
    else:
        # LTC returned no rows — no LTC credit this period.
        ltc["LTC_Uplift_Calc"] = pd.Series(dtype=float)
        ltc_lookup = {}
        _ltc_id_based = False

    ret_grp = ret.groupby("Opportunity ID").agg(
        BMI_SUM  = ("BMI_ARR",     "sum"),
        SC_SUM   = ("Roll_SC_Ret", "sum"),
        Rep      = ("Opportunity Owner", "first"),
        OppName  = ("Opportunity Name",  "first"),
        Oppty_Region_ret = ("Oppty Region", "first") if "Oppty Region" in ret.columns
                           else ("Oppty Team", "first"),
        Team     = ("Oppty Team",        "first"),
    ).reset_index()
    ret_grp["Retention_Credit"] = (ret_grp["BMI_SUM"] - ret_grp["SC_SUM"]).clip(lower=0)

    # Assign leader to each retention deal using the same 4-level priority
    _ret_for_ldr = ret.copy()
    _ret_for_ldr["Opportunity Owner"] = _ret_for_ldr["Opportunity Owner"].apply(normalize)
    ret_grp["Leader"] = _assign_leader(
        _ret_for_ldr.loc[ret.groupby("Opportunity ID").head(1).index
                         if False else slice(None)],
        region_leader_map, mgr_col_ret, rep_leader_map,
    ).reindex(range(len(ret_grp))).fillna("Unknown")

    # Simpler: reassign per ret_grp row using the first row of each opp group
    # (already aggregated above — derive leader from ret_grp region columns)
    _rg = ret_grp.rename(columns={"Oppty_Region_ret": "Oppty Region",
                                   "Team": "Oppty Team"})
    _rg["Opportunity Owner"] = _rg["Rep"]
    if mgr_col_ret in ret.columns:
        _first_mgr = (ret.groupby("Opportunity ID")[mgr_col_ret]
                      .first().reset_index()
                      .rename(columns={mgr_col_ret: "_mgr_ret"}))
        _rg = _rg.merge(_first_mgr, on="Opportunity ID", how="left")
        _rg[mgr_col_ret] = _rg.get("_mgr_ret", "")
    else:
        _rg[mgr_col_ret] = ""
    ret_grp["Leader"] = _assign_leader(_rg, region_leader_map,
                                       mgr_col_ret, rep_leader_map).values

    ret_by_oppname = ret_grp.groupby("OppName")["Retention_Credit"].sum().to_dict()

    # ── Master dataset ───────────────────────────────────────────────────────
    master = cw.copy()
    master["Rep"] = master["Opportunity Owner"]   # normalized above

    # Deal-level leader: Oppty Region → region_leader_map is source of truth.
    # This correctly handles mid-year transfers: a deal's Oppty Region is
    # stamped at close and never changes, so it is attributed to whoever
    # manages that region TODAY regardless of where the owner is now.
    master["Leader"] = _assign_leader(master, region_leader_map,
                                      mgr_col_cw, rep_leader_map)

    # Region display: prefer stable Oppty Region; fall back to Oppty Team
    if "Oppty Region" in master.columns:
        master["Region"] = master["Oppty Region"].fillna(
            master.get("Oppty Team", "Unknown")).str.strip()
    elif "Oppty Team" in master.columns:
        master["Region"] = master["Oppty Team"]
    else:
        master["Region"] = "Unknown"

    master["Segment"] = master["Region"].apply(seg_from_team)
    master["VP"]      = master["Segment"].map(VP_MAP).fillna("N/A")
    _fa_cw = ("Forecast Amount (converted)" if "Forecast Amount (converted)" in master.columns
              else "Forecast Amount")
    master["Forecast_Amount_ARR"] = pd.to_numeric(master[_fa_cw], errors="coerce").fillna(0)
    master["_OppName"] = master["Opportunity Name"].str.strip()
    master["In_Complete"] = master["_OppName"].isin(complete_names).astype(int)
    master["Complete_Credit"]  = master["_OppName"].map(complete_lookup).fillna(0)

    # ── Endorsed App CW ARR override ─────────────────────────────────────────
    # Endorsed App deals (partner contracts via Motus, Blue Dot, etc.) have two
    # distinct values in the CW report:
    #   • Forecast Amount  = SAP Concur's take-rate portion only (~35%)
    #   • Roll-up Sales Credit Calculation = full partner contract ARR (credit base)
    # For these deals we credit the full Roll-up Sales Credit value, not just the
    # Forecast Amount.  All other deals continue to use Forecast Amount as the base.
    # Complete-exclusion still applies: if in the Complete file, CW_ARR_Adjusted = $0.
    master["Endorsed_App_Flag"] = (
        master["Opportunity Name"]
        .str.contains("Endorsed App", case=False, na=False)
        .astype(int)
    )
    _ru_sc_cw = next(
        (c for c in master.columns if c.lower().startswith("roll-up sales credit")), None)
    if _ru_sc_cw:
        _rollup_sc = pd.to_numeric(master[_ru_sc_cw], errors="coerce").fillna(0)
        _base_arr  = np.where(master["Endorsed_App_Flag"] == 1,
                               _rollup_sc, master["Forecast_Amount_ARR"])
    else:
        _base_arr  = master["Forecast_Amount_ARR"]

    master["CW_ARR_Adjusted"]  = np.where(master["In_Complete"] == 1, 0, _base_arr)
    if _ltc_id_based and _cw_id_col:
        master["LTC_Uplift"] = master[_cw_id_col].astype(str).str.strip().map(ltc_lookup).fillna(0)
    else:
        master["LTC_Uplift"] = master["_OppName"].map(ltc_lookup).fillna(0)
    master["Retention_Credit"] = master["_OppName"].map(ret_by_oppname).fillna(0)
    master["Total_Credited"]   = (master["CW_ARR_Adjusted"] + master["LTC_Uplift"]
                                  + master["Retention_Credit"] + master["Complete_Credit"])

    # ── Standalone retention (not matched to CW) ─────────────────────────────
    cw_names    = set(master["_OppName"])
    standalone  = ret_grp[~ret_grp["OppName"].isin(cw_names)]
    total_ret   = ret_grp["Retention_Credit"].sum()

    # ── Rep aggregation ──────────────────────────────────────────────────────
    # Sort by Close Date so .last() gives the most recent deal's leader/region.
    # Used for the rep table only — the leader table is now aggregated at deal
    # level (see below) so mid-year transfers are split correctly there.
    master_s = master.sort_values("Close Date", na_position="first")
    cw_by_rep = master_s.groupby("Rep").agg(
        Leader_cw       = ("Leader",          "last"),   # current assignment
        Region_cw       = ("Region",          "last"),
        Segment_cw      = ("Segment",         "last"),
        VP              = ("VP",              "last"),
        CW_ARR          = ("CW_ARR_Adjusted", "sum"),
        LTC_Credit      = ("LTC_Uplift",      "sum"),
        Complete_Credit = ("Complete_Credit", "sum"),
        CW_Units        = ("Opportunity Name","count"),
    ).reset_index()

    ret_by_rep = ret_grp.groupby("Rep").agg(
        Retention_Credit = ("Retention_Credit","sum"),
        Leader_ret       = ("Leader",          "first"),
        Team_ret         = ("Team",            "first"),
    ).reset_index()

    rep_grp = cw_by_rep.merge(ret_by_rep[["Rep","Retention_Credit","Leader_ret","Team_ret"]],
                               on="Rep", how="outer")
    for col in ["CW_ARR","LTC_Credit","Complete_Credit","CW_Units","Retention_Credit"]:
        rep_grp[col] = rep_grp[col].fillna(0)

    def resolve_leader(row):
        # Leader_cw already reflects deal-level Oppty Region attribution,
        # with .last() = current assignment for transferred reps.
        for src in [row.get("Leader_cw",""), row.get("Leader_ret","")]:
            if isinstance(src, str) and src.strip().lower() not in ("","nan","unknown","n/a"):
                return src.strip()
        return rep_leader_map.get(row["Rep"], "Unknown")

    def resolve_segment(row):
        for src in [row.get("Segment_cw",""), seg_from_team(row.get("Team_ret",""))]:
            if isinstance(src, str) and src not in ("","Unknown","nan"):
                return src
        return seg_from_team(rep_team_map.get(row["Rep"], ""))

    def resolve_region(row):
        for src in [row.get("Region_cw",""), row.get("Team_ret","")]:
            if isinstance(src, str) and src.strip().lower() not in ("","nan","unknown","n/a"):
                return src.strip()
        return rep_team_map.get(row["Rep"], "Unknown")

    rep_grp["Leader"]  = rep_grp.apply(resolve_leader, axis=1)
    rep_grp["Region"]  = rep_grp.apply(resolve_region, axis=1)
    rep_grp["Segment"] = rep_grp.apply(resolve_segment, axis=1)
    rep_grp["VP"]      = rep_grp["Segment"].map(VP_MAP).fillna("N/A")
    # Force numeric dtype before division — outer merge can leave object columns
    for _c in ["CW_ARR","LTC_Credit","Retention_Credit","Complete_Credit","CW_Units"]:
        rep_grp[_c] = pd.to_numeric(rep_grp[_c], errors="coerce").fillna(0)
    rep_grp["Total_Credited"] = (rep_grp["CW_ARR"] + rep_grp["LTC_Credit"]
                                 + rep_grp["Retention_Credit"] + rep_grp["Complete_Credit"])
    rep_grp["Monthly_Quota"] = pd.to_numeric(
        rep_grp["Rep"].map(lambda r: rep_quota_override.get(r,
            rep_quota_map.get(r, 0) * monthly_factor)),
        errors="coerce").fillna(0)
    rep_grp["Pct_to_Linearity"] = np.where(
        rep_grp["Monthly_Quota"] > 0,
        rep_grp["Total_Credited"] / rep_grp["Monthly_Quota"], np.nan)
    rep_grp = rep_grp[["Rep","Leader","Region","Segment","VP",
                        "CW_ARR","LTC_Credit","Retention_Credit","Complete_Credit",
                        "Total_Credited","Monthly_Quota","Pct_to_Linearity","CW_Units"]]

    # ── Leader aggregation (deal-level) ─────────────────────────────────────
    # Aggregate from master (one row per CW deal) rather than from rep_grp
    # (one row per rep).  This ensures mid-year transfers are attributed
    # correctly: each deal is owned by whoever manages the Oppty Region today,
    # so Ashley McCue's Strategic-era deals flow to the Strategic leader and
    # her Premier-era deals flow to the Premier leader regardless of which
    # team she is currently on.  The rep table is unaffected — reps continue
    # to show under their most-recent-deal leader in the rep view.
    _ldr_cw = master.groupby("Leader").agg(
        Segment         = ("Segment",          "first"),
        Region          = ("Region",           "first"),
        VP              = ("VP",               "first"),
        CW_ARR          = ("CW_ARR_Adjusted",  "sum"),
        LTC_Credit      = ("LTC_Uplift",       "sum"),
        Complete_Credit = ("Complete_Credit",  "sum"),
        Retention_CW    = ("Retention_Credit", "sum"),   # matched retention
        CW_Units        = ("Opportunity Name", "count"),
    ).reset_index()

    # Standalone retention deals (no matching CW deal) — group by deal-level leader
    _ldr_ret = standalone.groupby("Leader").agg(
        Retention_Standalone = ("Retention_Credit", "sum"),
    ).reset_index()

    ldr_grp = _ldr_cw.merge(_ldr_ret, on="Leader", how="outer")
    for _c in ["CW_ARR", "LTC_Credit", "Complete_Credit", "Retention_CW", "CW_Units"]:
        ldr_grp[_c] = ldr_grp[_c].fillna(0)
    ldr_grp["Retention_Standalone"] = ldr_grp["Retention_Standalone"].fillna(0)

    # For leaders who only appear in standalone retention (no CW deals),
    # Segment/Region/VP will be NaN from the outer merge — fill from static maps.
    ldr_grp["Segment"] = ldr_grp.apply(
        lambda r: r["Segment"]
        if (pd.notna(r.get("Segment")) and
            str(r.get("Segment", "")).strip() not in ("", "Unknown", "nan"))
        else LEADER_SEGMENT.get(r["Leader"], "Unknown"), axis=1)
    ldr_grp["VP"]     = ldr_grp["Segment"].map(VP_MAP).fillna("N/A")
    ldr_grp["Region"] = ldr_grp["Region"].fillna(ldr_grp["Segment"])

    ldr_grp["Retention_Credit"] = ldr_grp["Retention_CW"] + ldr_grp["Retention_Standalone"]
    ldr_grp.drop(columns=["Retention_CW", "Retention_Standalone"], inplace=True)

    ldr_grp["Total_Credited"] = (ldr_grp["CW_ARR"] + ldr_grp["LTC_Credit"]
                                 + ldr_grp["Retention_Credit"] + ldr_grp["Complete_Credit"])
    ldr_grp["Monthly_Quota"] = pd.to_numeric(
        ldr_grp["Leader"].map(lambda l: ldr_quota_map.get(l, 0)),
        errors="coerce").fillna(0)
    ldr_grp["Pct_to_Linearity"] = np.where(
        ldr_grp["Monthly_Quota"] > 0,
        ldr_grp["Total_Credited"] / ldr_grp["Monthly_Quota"], np.nan)

    # ── Segment aggregation ──────────────────────────────────────────────────
    seg_grp = rep_grp[rep_grp["Segment"].isin(["Key","Premier","Strategic"])].groupby("Segment").agg(
        CW_ARR          = ("CW_ARR",           "sum"),
        LTC_Credit      = ("LTC_Credit",       "sum"),
        Retention_Credit= ("Retention_Credit", "sum"),
        Complete_Credit = ("Complete_Credit",  "sum"),
        Total_Credited  = ("Total_Credited",   "sum"),
        CW_Units        = ("CW_Units",         "sum"),
    ).reset_index()
    seg_grp["PL_Quota"]   = pd.to_numeric(seg_grp["Segment"].map(seg_pl),   errors="coerce").fillna(0)
    seg_grp["Lin_Target"] = pd.to_numeric(seg_grp["Segment"].map(
        lambda s: sum(ldr_quota_map.get(l,0) for l,sg in LEADER_SEGMENT.items() if sg==s)),
        errors="coerce").fillna(0)
    seg_grp["Pct_to_PL"]  = np.where(seg_grp["PL_Quota"]>0,
        seg_grp["Total_Credited"]/seg_grp["PL_Quota"], np.nan)
    seg_grp["Pct_to_Lin"] = np.where(seg_grp["Lin_Target"]>0,
        seg_grp["Total_Credited"]/seg_grp["Lin_Target"], np.nan)

    # ── Attainment distribution ──────────────────────────────────────────────
    rq = rep_grp[rep_grp["Monthly_Quota"] > 0].copy()
    rq["Bucket"] = rq["Pct_to_Linearity"].apply(bucket)
    dist = rq.groupby("Bucket").size().reset_index(name="Count")

    # ── Org summary ──────────────────────────────────────────────────────────
    org_total = rep_grp["Total_Credited"].sum()
    org_pl    = seg_pl.get("Total", 0)
    org_lin   = sum(ldr_quota_map.get(l,0) for l in LEADER_SEGMENT)
    org = {
        "Total_Credited":     org_total,
        "CW_ARR":             rep_grp["CW_ARR"].sum(),
        "LTC_Credit":         rep_grp["LTC_Credit"].sum(),
        "Retention_Credit":   total_ret,
        "Complete_Credit":    rep_grp["Complete_Credit"].sum(),
        "PL_Quota":           org_pl,
        "Lin_Target":         org_lin,
        "Pct_to_PL":          org_total / org_pl   if org_pl   else 0,
        "Pct_to_Linearity":   org_total / org_lin  if org_lin  else 0,
        "CW_Units":           int(len(cw)),
        "Above_100":          int((rq["Pct_to_Linearity"] >= 1.0).sum()),
        "Above_120":          int((rq["Pct_to_Linearity"] >= 1.2).sum()),
        "Total_Reps":         int(len(rq)),
        "Row_Counts": {"CW": len(cw), "LTC": len(ltc), "Retention": len(ret), "Complete": len(comp)},
    }

    # ── Deal-level breakdown table ────────────────────────────────────────────
    _DEAL_COLS = ["Rep","Leader","Region","Segment","Opportunity Name",
                  "Close Date","CW_ARR","LTC_Credit",
                  "Retention_Credit","Complete_Credit","Total_Credited"]
    _opp_cw = master[["Rep","Leader","Region","Segment",
                       "Opportunity Name","Close Date",
                       "CW_ARR_Adjusted","LTC_Uplift",
                       "Retention_Credit","Complete_Credit",
                       "Total_Credited"]].copy().rename(columns={
        "CW_ARR_Adjusted": "CW_ARR",
        "LTC_Uplift":      "LTC_Credit",
    })
    # Standalone retention deals (no matching CW deal)
    _rep_region_map = rep_grp.drop_duplicates("Rep").set_index("Rep")["Region"].to_dict()
    _rep_seg_map    = rep_grp.drop_duplicates("Rep").set_index("Rep")["Segment"].to_dict()
    _opp_ret = standalone[["Rep","Leader","OppName","Retention_Credit"]].copy()
    _opp_ret = _opp_ret.rename(columns={"OppName": "Opportunity Name"})
    _opp_ret["Region"]          = _opp_ret["Rep"].map(_rep_region_map).fillna("Unknown")
    _opp_ret["Segment"]         = _opp_ret["Rep"].map(_rep_seg_map).fillna("Unknown")
    _opp_ret["CW_ARR"]          = 0.0
    _opp_ret["LTC_Credit"]      = 0.0
    _opp_ret["Complete_Credit"] = 0.0
    _opp_ret["Close Date"]      = pd.NaT
    _opp_ret["Total_Credited"]  = _opp_ret["Retention_Credit"]
    deals = pd.concat([_opp_cw[_DEAL_COLS], _opp_ret[_DEAL_COLS]], ignore_index=True)

    return {"org": org, "segment": seg_grp, "leader": ldr_grp, "rep": rep_grp,
            "dist": dist, "master": master, "deals": deals}


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────────────────────
for k, v in [("sf", None), ("data", None), ("last_run", None),
              ("raw_cw", None), ("raw_ltc", None), ("raw_ret", None), ("raw_comp", None),
              ("tgt_bytes", None),
              ("monthly_factor", None), ("region_leader_map", {})]:
    if k not in st.session_state:
        st.session_state[k] = v

# period_label / sel_region / sel_rep defaults (overwritten by sidebar widgets each run)
period_label = "—"
sel_region   = "All Regions"
sel_rep      = "All Reps"

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### SAP Concur")
    st.markdown("**US SMB Client Sales Dashboard**")
    st.markdown("---")

    # ── Credentials ──────────────────────────────────────────────────────────
    with st.expander("Salesforce Connection", expanded=st.session_state.sf is None):
        secrets_sf = st.secrets.get("salesforce", {}) if hasattr(st, "secrets") else {}

        auth_mode = st.radio("Auth Method",
                             ["Session ID (SSO / Microsoft login)", "Username + Password"],
                             index=0, horizontal=True)

        if auth_mode == "Session ID (SSO / Microsoft login)":
            session_id = st.text_input("Session ID", type="password",
                                       value=secrets_sf.get("session_id", ""),
                                       placeholder="00D…  (paste from browser — see instructions below)",
                                       help="Log into Salesforce in your browser → F12 → "
                                            "Application → Cookies → sapconcur.my.salesforce.com → "
                                            "copy the value of the 'sid' cookie")
            instance_url = st.text_input("Instance URL",
                                         value=secrets_sf.get("instance_url",
                                                               "https://sapconcur.my.salesforce.com"))
            col_a, col_b = st.columns(2)
            if col_a.button("Connect", use_container_width=True):
                if not session_id:
                    st.error("Paste your Session ID first.")
                else:
                    try:
                        st.session_state.sf = sf_connect_session(session_id, instance_url)
                        st.success("Connected")
                    except Exception as e:
                        st.error(str(e))
        else:
            username  = st.text_input("Username (email)",
                                      value=secrets_sf.get("username", ""),
                                      placeholder="you@company.com")
            password  = st.text_input("Password", type="password",
                                      value=secrets_sf.get("password", ""))
            sec_token = st.text_input("Security Token", type="password",
                                      value=secrets_sf.get("security_token", ""),
                                      help="Salesforce Settings → My Personal Information → Reset My Security Token")
            domain    = st.text_input("Domain",
                                      value=secrets_sf.get("domain", "login"),
                                      help="'login' for standard orgs, 'test' for sandboxes, "
                                           "or your My Domain prefix e.g. 'sapconcur.my'")
            col_a, col_b = st.columns(2)
            if col_a.button("Connect", use_container_width=True):
                try:
                    st.session_state.sf = sf_connect(username, password, sec_token, domain)
                    st.success("Connected")
                except Exception as e:
                    st.error(str(e))

        if col_b.button("Disconnect", use_container_width=True):
            st.session_state.sf = None
            st.session_state.data = None

    conn_ok = st.session_state.sf is not None
    st.markdown(
        f"<span style='color:{'#188918' if conn_ok else '#BB0000'};font-weight:600'>"
        f"{'● Connected' if conn_ok else '○ Not connected'}</span>",
        unsafe_allow_html=True)

    st.markdown("---")

    # ── Report IDs ────────────────────────────────────────────────────────────
    with st.expander("Report IDs", expanded=True):
        secrets_rpt = st.secrets.get("reports", {}) if hasattr(st, "secrets") else {}
        rpt_cw   = st.text_input("CW ARR Report ID",    value=secrets_rpt.get("cw_arr_report_id",""),
                                  placeholder="00O…")
        rpt_ltc  = st.text_input("LTC Report ID",       value=secrets_rpt.get("ltc_report_id",""),
                                  placeholder="00O…")
        rpt_ret  = st.text_input("Retention Report ID", value=secrets_rpt.get("retention_report_id",""),
                                  placeholder="00O…")
        rpt_comp = st.text_input("Complete Report ID",  value=secrets_rpt.get("complete_report_id",""),
                                  placeholder="00O…")

    st.markdown("---")

    # ── Reporting Period ──────────────────────────────────────────────────────
    st.markdown("**Reporting Period**")
    period_type = st.radio("Period Type", ["Monthly", "Quarterly", "YTD"],
                           horizontal=True)
    sel_year = st.number_input("Year", min_value=2020, max_value=2030,
                                value=2026, step=1)
    QUARTER_MONTHS = {"Q1":[1,2,3],"Q2":[4,5,6],"Q3":[7,8,9],"Q4":[10,11,12]}

    if period_type == "Monthly":
        sel_month = st.selectbox("Month", MONTH_NAMES,
                                  index=MONTH_NAMES.index("July"))
        month_nums    = [MONTH_NAMES.index(sel_month) + 1]
        period_label  = f"{sel_month} {int(sel_year)}"
    elif period_type == "Quarterly":
        sel_quarter = st.selectbox("Quarter", ["Q1","Q2","Q3","Q4"], index=2)
        month_nums   = QUARTER_MONTHS[sel_quarter]
        period_label = f"{sel_quarter} {int(sel_year)}"
    else:  # YTD
        sel_month = st.selectbox("Through Month", MONTH_NAMES,
                                  index=MONTH_NAMES.index("July"))
        month_nums   = list(range(1, MONTH_NAMES.index(sel_month) + 2))
        period_label = f"YTD through {sel_month} {int(sel_year)}"

    st.markdown("---")

    # ── Quota Targets file ────────────────────────────────────────────────────
    st.markdown("**Quota Targets File**")
    tgt_upload = st.file_uploader("Upload targets .xlsx",
                                   type=["xlsx"],
                                   help="Final Client Sales Targets … Leader File.xlsx")
    if tgt_upload:
        st.session_state.tgt_bytes = tgt_upload.read()
    elif DEFAULT_TARGETS.exists() and st.session_state.tgt_bytes is None:
        st.session_state.tgt_bytes = DEFAULT_TARGETS.read_bytes()
        st.caption(f"Using: {DEFAULT_TARGETS.name}")

    st.markdown("---")

    # ── Run Reports ───────────────────────────────────────────────────────────
    run_btn = st.button("Run Reports", type="primary", use_container_width=True,
                        disabled=not conn_ok)
    if run_btn:
        if not all([rpt_cw, rpt_ltc, rpt_ret, rpt_comp]):
            st.error("All 4 Report IDs are required.")
        elif st.session_state.tgt_bytes is None:
            st.error("Please upload the Quota Targets file.")
        else:
            n_months = len(month_nums)
            sf   = st.session_state.sf
            _yr  = int(sel_year)
            with st.spinner(
                f"Running Salesforce reports… ({n_months} month{'s' if n_months>1 else ''})"
                " — serving from cache or queuing a new fetch"
            ):
                try:
                    (cw_df, n1, ltc_df, n2,
                     ret_df, n3, comp_df, n4) = fetch_all_reports(
                        sf, rpt_cw, rpt_ltc, rpt_ret, rpt_comp,
                        tuple(month_nums), _yr)

                    st.session_state.raw_cw   = cw_df
                    st.session_state.raw_ltc  = ltc_df
                    st.session_state.raw_ret  = ret_df
                    st.session_state.raw_comp = comp_df
                    st.caption(f"CW ARR: {n1} rows | LTC: {n2} rows | Retention: {n3} rows | Complete: {n4} rows")

                    # ── Row count warnings ────────────────────────────────────
                    zero_reports = [name for name, n in
                                    [("CW ARR", n1),("LTC", n2),("Retention", n3),("Complete", n4)]
                                    if n == 0]
                    if zero_reports:
                        st.warning(
                            f"**{', '.join(zero_reports)} returned 0 rows.** "
                            f"The Salesforce report(s) likely have a date filter built in "
                            f"that is excluding data. Open each report in Salesforce, remove "
                            f"or widen any Close Date / Date filters, save the report, "
                            f"then click Run Reports again.")
                    # ── API row cap warning ───────────────────────────────────
                    capped = [name for name, n in
                              [("CW ARR", n1),("LTC", n2),("Retention", n3),("Complete", n4)]
                              if n >= 2000]
                    if capped:
                        st.warning(
                            f"**{', '.join(capped)} returned 2,000 rows** — the Salesforce "
                            f"Analytics API cap. Some records may be missing. Consider narrowing "
                            f"the report's date filter.")

                except Exception as e:
                    kind, msg = _sfdc_error_msg(e)
                    if kind == "rate_limit":
                        st.warning(f"**{msg}**")
                    else:
                        st.error(f"**{msg}**")
                    st.stop()

            with st.spinner("Calculating metrics…"):
                try:
                    lq, rq, rlm, rtm, sp, mf = load_quotas(
                        st.session_state.tgt_bytes, month_nums)
                    st.session_state.monthly_factor = mf

                    # Build pro-rated quota overrides (mid-year hires / transfers)
                    rq_override = {}
                    try:
                        rq_override = build_quota_overrides(
                            st.session_state.tgt_bytes, month_nums)
                    except Exception as _ov_err:
                        st.warning(f"Pro-rated quota calculation skipped: {_ov_err}")

                    # Build live Region → Leader map from the full CW ARR
                    # dataset just fetched.  Uses most-recent-month deals per
                    # region so it always reflects the current org structure.
                    rlm_region = build_region_leader_map(st.session_state.raw_cw)
                    st.session_state.region_leader_map = rlm_region

                    result = run_calc(
                        st.session_state.raw_cw,  st.session_state.raw_ltc,
                        st.session_state.raw_ret, st.session_state.raw_comp,
                        lq, rq, rlm, rtm, sp, month_nums, int(sel_year),
                        monthly_factor=mf,
                        region_leader_map=rlm_region,
                        rep_quota_override=rq_override)
                    st.session_state.data     = result
                    st.session_state.last_run = datetime.now()
                except Exception as e:
                    st.error(f"Calculation error: {e}\n\nCheck the column preview above to verify "
                             f"column names match what the calc engine expects.")
                    st.stop()
            st.rerun()

    # ── Filters (shown after data loads) ─────────────────────────────────────
    if st.session_state.data:
        st.markdown("---")
        st.markdown("**Filters**")
        rep_df_all  = st.session_state.data["rep"]
        ldr_df_all  = st.session_state.data["leader"]

        segs_avail = ["All Segments"] + sorted(
            rep_df_all[rep_df_all["Segment"].isin(["Key","Premier","Strategic"])]["Segment"].unique())
        sel_seg = st.selectbox("Segment", segs_avail)

        _rep_seg = rep_df_all[rep_df_all["Segment"]==sel_seg] if sel_seg != "All Segments" else rep_df_all
        regions_avail = ["All Regions"] + sorted([
            r for r in _rep_seg["Region"].unique()
            if isinstance(r, str) and r.strip() not in ("", "Unknown", "nan")
        ])
        sel_region = st.selectbox("Region", regions_avail)

        _rep_reg = _rep_seg[_rep_seg["Region"]==sel_region] if sel_region != "All Regions" else _rep_seg
        ldrs_avail = ["All Leaders"] + sorted(_rep_reg["Leader"].dropna().unique().tolist())
        sel_ldr = st.selectbox("Leader", ldrs_avail)

        _rep_ldr = _rep_reg[_rep_reg["Leader"]==sel_ldr] if sel_ldr != "All Leaders" else _rep_reg
        reps_avail = ["All Reps"] + sorted(_rep_ldr["Rep"].dropna().unique().tolist())
        sel_rep = st.selectbox("Rep", reps_avail)
    else:
        sel_seg    = "All Segments"
        sel_region = "All Regions"
        sel_ldr    = "All Leaders"
        sel_rep    = "All Reps"

    if st.session_state.last_run:
        st.caption(f"Last run: {st.session_state.last_run.strftime('%b %d %Y %H:%M')}")

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — SPLASH if no data
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.data is None:
    st.markdown(
        f"<div style='text-align:center;padding:80px 0'>"
        f"<div style='font-size:22px;font-weight:700;color:{C['dark_blue']}'>"
        f"SAP Concur &nbsp;|&nbsp; US SMB Client Sales Performance Dashboard</div>"
        f"<div style='color:{C['dark_grey']};margin-top:12px;font-size:15px'>"
        f"Connect to Salesforce and click <strong>Run Reports</strong> to load data.</div>"
        f"<div style='color:{C['dark_grey']};margin-top:8px;font-size:13px'>"
        f"Credentials and Report IDs can be pre-configured in "
        f"<code>.streamlit/secrets.toml</code>.</div>"
        f"</div>",
        unsafe_allow_html=True)
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────
data    = st.session_state.data
org     = data["org"]
seg_df  = data["segment"].copy()
ldr_df  = data["leader"].copy()
rep_df  = data["rep"].copy()
dist_df = data["dist"].copy()

# Apply filters
ldr_filtered = ldr_df.copy()
rep_filtered = rep_df.copy()
if sel_seg != "All Segments":
    ldr_filtered = ldr_filtered[ldr_filtered["Segment"]==sel_seg]
    rep_filtered = rep_filtered[rep_filtered["Segment"]==sel_seg]
if sel_region != "All Regions":
    rep_filtered = rep_filtered[rep_filtered["Region"]==sel_region]
    leaders_in_region = rep_filtered["Leader"].unique()
    ldr_filtered = ldr_filtered[ldr_filtered["Leader"].isin(leaders_in_region)]
if sel_ldr != "All Leaders":
    rep_filtered = rep_filtered[rep_filtered["Leader"]==sel_ldr]
    ldr_filtered = ldr_filtered[ldr_filtered["Leader"]==sel_ldr]
if sel_rep != "All Reps":
    rep_filtered = rep_filtered[rep_filtered["Rep"]==sel_rep]

# ── Deal-level filter ────────────────────────────────────────────────────────
deals_filtered = data.get("deals", pd.DataFrame()).copy()
if not deals_filtered.empty:
    if sel_seg != "All Segments":
        deals_filtered = deals_filtered[deals_filtered["Segment"]==sel_seg]
    if sel_region != "All Regions":
        deals_filtered = deals_filtered[deals_filtered["Region"]==sel_region]
    if sel_ldr != "All Leaders":
        deals_filtered = deals_filtered[deals_filtered["Leader"]==sel_ldr]
    if sel_rep != "All Reps":
        deals_filtered = deals_filtered[deals_filtered["Rep"]==sel_rep]
    deals_filtered = deals_filtered[deals_filtered["Total_Credited"] > 0].copy()

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(
    f"<div style='font-size:21px;font-weight:700;color:{C['dark_blue']};"
    f"border-bottom:3px solid {C['blue']};padding-bottom:7px;margin-bottom:18px'>"
    f"SAP Concur &nbsp;|&nbsp; US SMB Client Sales &nbsp;|&nbsp; "
    f"{period_label} Performance Recap"
    f"</div>",
    unsafe_allow_html=True)

# ── KPI Cards ─────────────────────────────────────────────────────────────────
tc  = org["Total_Credited"]
plp = org["Pct_to_PL"]
lip = org["Pct_to_Linearity"]

k1,k2,k3,k4,k5,k6 = st.columns(6)
k1.markdown(kpi_html("Total Credited",    fmt_m(tc)), unsafe_allow_html=True)
k2.markdown(kpi_html("% to P&L Plan",    fmt_pct(plp), "green" if plp>=1 else "amber"), unsafe_allow_html=True)
k3.markdown(kpi_html("% to Linearity",   fmt_pct(lip), "green" if lip>=1 else "amber"), unsafe_allow_html=True)
k4.markdown(kpi_html("CW Units",         str(org["CW_Units"]), "dark"), unsafe_allow_html=True)
k5.markdown(kpi_html("Reps ≥ 100%",      f"{org['Above_100']}/{org['Total_Reps']}", "green"), unsafe_allow_html=True)
k6.markdown(kpi_html("Reps ≥ 120%",      f"{org['Above_120']}/{org['Total_Reps']}", "green"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Data Diagnostic (persistent expander in main area) ────────────────────────
with st.expander("Report Data Diagnostic — expand to inspect raw data", expanded=False):
    # ── POST attempt trace (shows exactly which path was taken per report) ───
    _dbg = st.session_state.get("sf_post_debug", {})
    if _dbg:
        st.markdown("**POST Attempt Log** *(A1=direct standardDateFilter, A2=metadata-rebuild, A3=reportFilters only)*")
        _rpt_ids = {
            st.secrets.get("reports", {}).get("cw_arr_report_id",  "—"): "CW ARR",
            st.secrets.get("reports", {}).get("ltc_report_id",     "—"): "LTC",
            st.secrets.get("reports", {}).get("retention_report_id","—"): "Retention",
            st.secrets.get("reports", {}).get("complete_report_id", "—"): "Complete",
        }
        for _rid, _trace in _dbg.items():
            _name = _rpt_ids.get(_rid, _rid)
            st.caption(f"**{_name}**: {_trace}")
        st.markdown("---")

    # ── Salesforce report saved configuration ─────────────────────────────
    st.markdown("**Salesforce Report Configuration** *(from saved metadata — helps diagnose filter issues)*")
    _rpt_label_ids = {
        "CW ARR":    st.secrets.get("reports", {}).get("cw_arr_report_id",  None),
        "LTC":       st.secrets.get("reports", {}).get("ltc_report_id",     None),
        "Retention": st.secrets.get("reports", {}).get("retention_report_id",None),
        "Complete":  st.secrets.get("reports", {}).get("complete_report_id", None),
    }
    for _lbl, _rid in _rpt_label_ids.items():
        if _rid and _rid in _meta_cache:
            _cached = _meta_cache[_rid]
            _rm = _cached.get("reportMetadata", {})
            _sdf = _rm.get("standardDateFilter")
            _rf  = _rm.get("reportFilters", [])
            st.caption(
                f"**{_lbl}** | standardDateFilter: {json.dumps(_sdf)} "
                f"| reportFilters ({len(_rf)} rows): "
                + ", ".join(f.get("column","?") for f in _rf[:6])
            )
        elif _rid:
            st.caption(f"**{_lbl}**: metadata not yet loaded (run reports first)")
    st.markdown("---")

    # ── Dedup trace ──────────────────────────────────────────────────────────
    _dedup = st.session_state.get("sf_dedup_info", {})
    if _dedup:
        st.markdown("**Deduplication Log** *(appears when a report ignores the API date filter — "
                    "indicates the direct POST body is still not being applied)*")
        for _rid, _msg in _dedup.items():
            _name = _rpt_ids.get(_rid, _rid) if _dbg else _rid
            st.warning(f"**{_name}**: {_msg}")
        st.markdown("---")

    for _lbl, _raw, _dcol in [
        ("CW ARR",    st.session_state.raw_cw,   "Close Date"),
        ("LTC",       st.session_state.raw_ltc,  "Close Date"),
        ("Retention", st.session_state.raw_ret,  "Final Month Closed"),
        ("Complete",  st.session_state.raw_comp, "Close Month"),
    ]:
        st.markdown(f"**{_lbl}** — {len(_raw)} rows from API")
        st.caption(f"Columns: {list(_raw.columns)}")
        if _dcol in _raw.columns and len(_raw) > 0:
            _sample = _raw[_dcol].dropna().head(3).tolist()
            st.caption(f"'{_dcol}' sample values: {_sample}")
            _parsed = pd.to_datetime(_raw[_dcol], errors="coerce")
            _dist   = _parsed.dt.to_period("M").value_counts().sort_index()
            st.caption(f"Date distribution (year-month: count): {_dist.to_dict()}")
        elif _dcol not in _raw.columns:
            st.warning(f"Column '{_dcol}' NOT FOUND in report. Available: {list(_raw.columns)}")
        # Complete: show all unique Close Month values and raw credit sum before
        # any Python filtering so we can see exactly what the API returned.
        if _lbl == "Complete" and len(_raw) > 0:
            _cm_uniq = _raw["Close Month"].astype(str).str.strip().unique().tolist() if "Close Month" in _raw.columns else []
            st.caption(f"All unique 'Close Month' raw values ({len(_cm_uniq)}): {_cm_uniq[:20]}")
            _sc_col = "Roll-up Sales Credit Calculation (converted)"
            if _sc_col in _raw.columns:
                _raw_credit_sum = pd.to_numeric(_raw[_sc_col], errors="coerce").sum()
                st.caption(f"Raw credit sum (ALL {len(_raw)} rows, before Python filter): ${_raw_credit_sum:,.0f}")
        # Retention: show FMC unique values so we know if year is embedded
        if _lbl == "Retention" and len(_raw) > 0 and "Final Month Closed" in _raw.columns:
            _fmc_uniq = _raw["Final Month Closed"].dropna().unique().tolist()[:20]
            st.caption(f"All unique 'Final Month Closed' values: {_fmc_uniq}")
            _has_cd = "Close Date" in _raw.columns
            st.caption(f"'Close Date' column present in Retention report: {_has_cd}")
        st.markdown("---")

# ── Row 2: Segment chart + Component donut ─────────────────────────────────
col_seg, col_donut = st.columns([3, 2])

with col_seg:
    st.markdown("<div class='sh'>Segment Performance</div>", unsafe_allow_html=True)
    plot_s = seg_df[seg_df["Segment"].isin(["Key","Premier","Strategic"])].copy()
    fig_s = go.Figure()
    for _, r in plot_s.iterrows():
        sn = r["Segment"]
        lin_pct = r.get("Pct_to_Lin", 0) or 0
        fig_s.add_trace(go.Bar(
            name=sn, x=[sn], y=[r["Total_Credited"]],
            marker_color=SEG_COLORS.get(sn, C["blue"]),
            text=f"${r['Total_Credited']/1e6:.2f}M<br>{lin_pct*100:.1f}% lin",
            textposition="outside", textfont=dict(size=11),
        ))
        if r.get("Lin_Target",0):
            fig_s.add_shape(type="line", x0=sn, x1=sn,
                            y0=0, y1=r["Lin_Target"],
                            line=dict(color=C["red"], width=2, dash="dot"))
    fig_s.update_layout(
        showlegend=False, height=310,
        paper_bgcolor=C["bg"], plot_bgcolor=C["bg"],
        margin=dict(t=30,b=10,l=10,r=10),
        yaxis=dict(title="Total Credited ($)", tickformat="$,.0f", gridcolor=C["grey"]),
        xaxis=dict(title=""),
        font=dict(family="72, Helvetica Neue, Arial"),
        annotations=[dict(x=0.98,y=1.06,xref="paper",yref="paper",
            text=f"<span style='color:{C['red']}'>— Linearity Target</span>",
            showarrow=False, font=dict(size=11), align="right")],
    )
    st.plotly_chart(fig_s, use_container_width=True)

with col_donut:
    st.markdown("<div class='sh'>Credit Components</div>", unsafe_allow_html=True)
    # Use filtered rep data so the donut updates when segment/region/leader filters are applied
    tc_f        = rep_filtered["Total_Credited"].sum()
    comp_vals   = [rep_filtered["CW_ARR"].sum(), rep_filtered["LTC_Credit"].sum(),
                   rep_filtered["Retention_Credit"].sum(), rep_filtered["Complete_Credit"].sum()]
    comp_labels = ["CW ARR", "LTC Uplift", "Retention", "Complete"]
    comp_colors = [C["blue"], C["light_blue"], "#5DB533", C["amber"]]
    fig_d = go.Figure(go.Pie(
        labels=comp_labels, values=comp_vals,
        marker=dict(colors=comp_colors), hole=0.52,
        textinfo="label+percent", textfont=dict(size=11),
        hovertemplate="%{label}: $%{value:,.0f}<extra></extra>",
    ))
    fig_d.add_annotation(text=f"<b>{fmt_m(tc_f)}</b><br>Total",
                         x=0.5, y=0.5, showarrow=False,
                         font=dict(size=13, color=C["dark_blue"]))
    fig_d.update_layout(height=310, paper_bgcolor=C["bg"],
                        margin=dict(t=30,b=10,l=10,r=10),
                        showlegend=True,
                        legend=dict(orientation="h",y=-0.08),
                        font=dict(family="72, Helvetica Neue, Arial"))
    st.plotly_chart(fig_d, use_container_width=True)

# ── Row 3: Leader attainment + Distribution ────────────────────────────────
col_ldr, col_dist = st.columns([3, 2])

with col_ldr:
    st.markdown("<div class='sh'>Leader Attainment</div>", unsafe_allow_html=True)
    lp = ldr_filtered[ldr_filtered["Segment"].isin(["Key","Premier","Strategic"])].copy()
    lp["Pct"] = lp["Pct_to_Linearity"].fillna(0) * 100
    lp["Color"] = lp["Pct"].apply(
        lambda x: C["green"] if x>=100 else (C["amber"] if x>=75 else C["red"]))
    lp = lp.sort_values("Pct", ascending=True)
    fig_l = go.Figure(go.Bar(
        x=lp["Pct"], y=lp["Leader"], orientation="h",
        marker_color=lp["Color"].tolist(),
        text=lp["Pct"].apply(lambda x: f"{x:.1f}%"),
        textposition="outside", textfont=dict(size=10),
        customdata=lp[["Total_Credited","Monthly_Quota","Segment"]].values,
        hovertemplate="<b>%{y}</b><br>Attainment: %{x:.1f}%<br>"
                      "Credited: $%{customdata[0]:,.0f}<br>"
                      "Quota: $%{customdata[1]:,.0f}<br>"
                      "Segment: %{customdata[2]}<extra></extra>",
    ))
    fig_l.add_vline(x=100, line_dash="dash", line_color=C["blue"], line_width=2,
                    annotation_text="100%", annotation_position="top right")
    h = max(280, len(lp)*36+60)
    fig_l.update_layout(height=h, paper_bgcolor=C["bg"], plot_bgcolor=C["bg"],
                        margin=dict(t=30,b=10,l=10,r=80),
                        xaxis=dict(title="% to Linearity",ticksuffix="%",gridcolor=C["grey"]),
                        yaxis=dict(title=""),
                        font=dict(family="72, Helvetica Neue, Arial"))
    st.plotly_chart(fig_l, use_container_width=True)

with col_dist:
    st.markdown("<div class='sh'>Attainment Distribution</div>", unsafe_allow_html=True)
    # Recompute distribution from filtered reps so it updates with segment/region/leader filters
    dist_f = rep_filtered[rep_filtered["Monthly_Quota"] > 0].copy()
    dist_f["Bucket"] = dist_f["Pct_to_Linearity"].apply(bucket)
    dp_raw = dist_f.groupby("Bucket").size().reset_index(name="Count")
    all_b = pd.DataFrame({"Bucket": BUCKET_ORDER})
    dp = all_b.merge(dp_raw, on="Bucket", how="left").fillna(0)
    dp["Color"] = dp["Bucket"].map(BUCKET_COLORS)
    fig_dist = go.Figure(go.Bar(
        x=dp["Count"], y=dp["Bucket"], orientation="h",
        marker_color=dp["Color"].tolist(),
        text=dp["Count"].astype(int), textposition="outside",
        textfont=dict(size=12),
        hovertemplate="%{y}: %{x:.0f} reps<extra></extra>",
    ))
    fig_dist.update_layout(height=280, paper_bgcolor=C["bg"], plot_bgcolor=C["bg"],
                           margin=dict(t=30,b=10,l=10,r=40),
                           xaxis=dict(title="Number of Reps",gridcolor=C["grey"]),
                           yaxis=dict(title="",categoryorder="array",
                                      categoryarray=BUCKET_ORDER),
                           font=dict(family="72, Helvetica Neue, Arial"))
    st.plotly_chart(fig_dist, use_container_width=True)

# ── Row 4: Rep table ───────────────────────────────────────────────────────
st.markdown("<div class='sh'>Rep Performance Detail</div>", unsafe_allow_html=True)
rt = rep_filtered[["Rep","Leader","Region","Segment","Monthly_Quota","CW_ARR","LTC_Credit",
                    "Retention_Credit","Complete_Credit","Total_Credited",
                    "Pct_to_Linearity","CW_Units"]].copy()
rt["Pct_to_Linearity"] = (rt["Pct_to_Linearity"] * 100).round(1)
rt.rename(columns={"Monthly_Quota":"Monthly Quota","CW_ARR":"CW ARR",
                    "LTC_Credit":"LTC Uplift","Retention_Credit":"Retention",
                    "Complete_Credit":"Complete","Total_Credited":"Total Credited",
                    "Pct_to_Linearity":"% to Lin","CW_Units":"Units"}, inplace=True)

money_cols = ["Monthly Quota","CW ARR","LTC Uplift","Retention","Complete","Total Credited"]
st.dataframe(
    rt.style
      .format({**{c:"${:,.0f}" for c in money_cols}, "% to Lin":"{:.1f}%"})
      .map(lambda v: (f"color:{C['green']};font-weight:600" if isinstance(v,float) and v>=100
                      else f"color:{C['red']}" if isinstance(v,float) and v<75 else ""),
           subset=["% to Lin"]),
    use_container_width=True, height=420,
)

# ── Row 5: Opportunity Breakdown ─────────────────────────────────────────────
st.markdown("<div class='sh'>Opportunity Breakdown</div>", unsafe_allow_html=True)
_auto_expand = (sel_rep != "All Reps")
with st.expander(
    f"Deals for: {sel_rep}" if _auto_expand else "Deal-Level Detail — select a Rep filter to focus",
    expanded=_auto_expand,
):
    if deals_filtered.empty:
        st.info("No credited deals match the current filters.")
    else:
        _dt = deals_filtered[["Rep","Opportunity Name","Close Date","Region",
                              "CW_ARR","LTC_Credit","Retention_Credit",
                              "Complete_Credit","Total_Credited"]].copy()
        _dt["Close Date"] = (pd.to_datetime(_dt["Close Date"], errors="coerce")
                             .dt.strftime("%b %d, %Y").fillna("—"))
        _dt = _dt.sort_values("Total_Credited", ascending=False)
        _dt.rename(columns={
            "CW_ARR":           "CW ARR",
            "LTC_Credit":       "LTC Uplift",
            "Retention_Credit": "Retention",
            "Complete_Credit":  "Complete",
            "Total_Credited":   "Total Credited",
        }, inplace=True)
        _mc = ["CW ARR","LTC Uplift","Retention","Complete","Total Credited"]
        st.dataframe(
            _dt.style.format({c: "${:,.0f}" for c in _mc}),
            use_container_width=True,
            height=min(450, 38 * len(_dt) + 42),
            hide_index=True,
        )
        # Summary totals beneath the table
        _c1, _c2, _c3, _c4, _c5 = st.columns(5)
        _c1.metric("CW ARR",    f"${deals_filtered['CW_ARR'].sum():,.0f}")
        _c2.metric("LTC",       f"${deals_filtered['LTC_Credit'].sum():,.0f}")
        _c3.metric("Retention", f"${deals_filtered['Retention_Credit'].sum():,.0f}")
        _c4.metric("Complete",  f"${deals_filtered['Complete_Credit'].sum():,.0f}")
        _c5.metric("Total",     f"${deals_filtered['Total_Credited'].sum():,.0f}")

# ── Row 6: Segment breakdown table ─────────────────────────────────────────
with st.expander("Segment Detail Table"):
    sd = seg_df[seg_df["Segment"].isin(["Key","Premier","Strategic"])][
        ["Segment","Total_Credited","PL_Quota","Pct_to_PL","Lin_Target","Pct_to_Lin","CW_Units"]].copy()
    sd["Pct_to_PL"]  = (sd["Pct_to_PL"]  * 100).round(1)
    sd["Pct_to_Lin"] = (sd["Pct_to_Lin"] * 100).round(1)
    sd.rename(columns={"Total_Credited":"Total Credited","PL_Quota":"P&L Quota",
                        "Pct_to_PL":"% to P&L","Lin_Target":"Linearity Target",
                        "Pct_to_Lin":"% to Lin","CW_Units":"CW Units"}, inplace=True)
    st.dataframe(sd.style.format(
        {"Total Credited":"${:,.0f}","P&L Quota":"${:,.0f}","Linearity Target":"${:,.0f}",
         "% to P&L":"{:.1f}%","% to Lin":"{:.1f}%"}),
        use_container_width=True)

# ── Row 7: Leader Excel Download ───────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("<div class='sh'>Leader Report Download</div>", unsafe_allow_html=True)

def _build_leader_excel(ldr_data, rep_data, seg_data, label):
    """Build a three-sheet Excel workbook: Leader Summary, Rep Detail, Segment Summary."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        # ── Sheet 1: Leader Summary ───────────────────────────────────────────
        ld = ldr_data[ldr_data["Segment"].isin(["Key","Premier","Strategic"])].copy()
        ld["Pct_to_Linearity"] = (ld["Pct_to_Linearity"].fillna(0) * 100).round(1)
        ld = ld[["Leader","Segment","Monthly_Quota","CW_ARR","LTC_Credit",
                  "Retention_Credit","Complete_Credit","Total_Credited",
                  "CW_Units","Pct_to_Linearity"]].copy()
        ld.rename(columns={
            "Monthly_Quota":    "Monthly Quota",
            "CW_ARR":           "CW ARR",
            "LTC_Credit":       "LTC Uplift",
            "Retention_Credit": "Retention",
            "Complete_Credit":  "Complete",
            "Total_Credited":   "Total Credited",
            "CW_Units":         "CW Units",
            "Pct_to_Linearity": "% to Linearity",
        }, inplace=True)
        ld.to_excel(writer, sheet_name="Leader Summary", index=False)

        # ── Sheet 2: Rep Detail ───────────────────────────────────────────────
        rd = rep_data.copy()
        rd["Pct_to_Linearity"] = (rd["Pct_to_Linearity"].fillna(0) * 100).round(1)
        rd = rd[["Rep","Leader","Segment","Region","Monthly_Quota","CW_ARR","LTC_Credit",
                  "Retention_Credit","Complete_Credit","Total_Credited",
                  "CW_Units","Pct_to_Linearity"]].copy()
        rd.rename(columns={
            "Monthly_Quota":    "Monthly Quota",
            "CW_ARR":           "CW ARR",
            "LTC_Credit":       "LTC Uplift",
            "Retention_Credit": "Retention",
            "Complete_Credit":  "Complete",
            "Total_Credited":   "Total Credited",
            "CW_Units":         "CW Units",
            "Pct_to_Linearity": "% to Linearity",
        }, inplace=True)
        rd.to_excel(writer, sheet_name="Rep Detail", index=False)

        # ── Sheet 3: Segment Summary ──────────────────────────────────────────
        sd2 = seg_data[seg_data["Segment"].isin(["Key","Premier","Strategic"])].copy()
        sd2["Pct_to_PL"]  = (sd2["Pct_to_PL"].fillna(0)  * 100).round(1)
        sd2["Pct_to_Lin"] = (sd2["Pct_to_Lin"].fillna(0) * 100).round(1)
        sd2 = sd2[["Segment","CW_ARR","LTC_Credit","Retention_Credit","Complete_Credit",
                    "Total_Credited","PL_Quota","Lin_Target",
                    "Pct_to_PL","Pct_to_Lin","CW_Units"]].copy()
        sd2.rename(columns={
            "CW_ARR":           "CW ARR",
            "LTC_Credit":       "LTC Uplift",
            "Retention_Credit": "Retention",
            "Complete_Credit":  "Complete",
            "Total_Credited":   "Total Credited",
            "PL_Quota":         "P&L Quota",
            "Lin_Target":       "Linearity Target",
            "Pct_to_PL":        "% to P&L",
            "Pct_to_Lin":       "% to Linearity",
            "CW_Units":         "CW Units",
        }, inplace=True)
        sd2.to_excel(writer, sheet_name="Segment Summary", index=False)

    buf.seek(0)
    return buf.getvalue()

_safe_label = re.sub(r"[^\w\-]", "_", period_label.replace("|","").replace(" ","_").strip())
_excel_file  = f"SMB_ClientSales_{_safe_label}.xlsx"
_excel_bytes = _build_leader_excel(ldr_filtered, rep_filtered, seg_df, period_label)

st.download_button(
    label="Download Excel",
    data=_excel_bytes,
    file_name=_excel_file,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Downloads Leader Summary, Rep Detail, and Segment Summary for the current filter selection.",
)

# ── CW Deal Detail Download ─────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("<div class='sh'>CW Deal Detail Download</div>", unsafe_allow_html=True)
st.caption("One row per CW opportunity — includes Account ID, Account Name, Opportunity ID, and all incentive amounts. Respects the current period selection.")

def _build_cw_detail_excel(master_df, period_lbl):
    """Build a deal-level Excel workbook from the master CW dataset."""
    buf = io.BytesIO()
    df = master_df.copy()

    # Detect Account ID column (SFDC exports various names)
    _acct_id_col = next(
        (c for c in df.columns if c.strip().lower() in (
            "account id", "account: id (18 char)", "account id (18 char)",
            "account id (18-char)", "accountid")),
        None)
    _acct_name_col = next(
        (c for c in df.columns if c.strip().lower() in (
            "account name", "account: name", "accountname")),
        None)

    # Detect opportunity ID column: prefer 18-char field
    _oid_col = ("ID (18 Char)" if "ID (18 Char)" in df.columns
                else "Opportunity ID" if "Opportunity ID" in df.columns
                else None)

    out = pd.DataFrame()
    out["Account ID"]        = df[_acct_id_col].fillna("")  if _acct_id_col else ""
    out["Account Name"]      = df[_acct_name_col].fillna("") if _acct_name_col else ""
    out["Opportunity ID"]    = df[_oid_col].fillna("")       if _oid_col else ""
    out["Opportunity Name"]  = df["Opportunity Name"].fillna("")
    out["Opportunity Owner"] = df["Opportunity Owner"].fillna("")
    out["Close Date"]        = pd.to_datetime(df["Close Date"], errors="coerce").dt.date
    out["Rep"]               = df.get("Rep",     df["Opportunity Owner"])
    out["Leader"]            = df.get("Leader",  "")
    out["Region"]            = df.get("Region",  "")
    out["Segment"]           = df.get("Segment", "")
    out["Forecast Amount"]   = df["Forecast_Amount_ARR"].round(2)
    out["CW ARR"]            = df["CW_ARR_Adjusted"].round(2)
    out["LTC Uplift"]        = df["LTC_Uplift"].round(2)
    out["Retention"]         = df["Retention_Credit"].round(2)
    out["Complete/Ref"]      = df["Complete_Credit"].round(2)
    out["Total Credited"]    = df["Total_Credited"].round(2)
    out["Endorsed App"]      = df.get("Endorsed_App_Flag", 0).astype(int)
    out["In Complete"]       = df.get("In_Complete", 0).astype(int)

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="CW Deal Detail", index=False)
        ws = writer.sheets["CW Deal Detail"]
        for col_cells in ws.columns:
            max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 50)

    buf.seek(0)
    return buf.getvalue()

_cw_detail_file  = f"CW_Deal_Detail_{_safe_label}.xlsx"
_cw_detail_bytes = _build_cw_detail_excel(data["master"], period_label)

st.download_button(
    label="Download CW Deal Detail",
    data=_cw_detail_bytes,
    file_name=_cw_detail_file,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="One row per CW opportunity with account, opportunity, and all incentive amounts.",
)

# ── Footer ─────────────────────────────────────────────────────────────────
row_counts = org.get("Row_Counts", {})
rc_str = " | ".join(f"{k}: {v} rows" for k,v in row_counts.items())
st.markdown(
    f"<div style='text-align:center;color:{C['dark_grey']};font-size:11px;margin-top:16px'>"
    f"SAP Concur US SMB Client Sales &nbsp;|&nbsp; {period_label} &nbsp;|&nbsp; "
    f"Source: Salesforce Reports &nbsp;|&nbsp; {rc_str}"
    f"</div>",
    unsafe_allow_html=True)
