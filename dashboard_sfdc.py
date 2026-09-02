"""
SAP Concur US SMB Client Sales — Live SFDC Dashboard
Connects to Salesforce, runs 4 source reports, applies the full calculation
engine, and renders an interactive monthly performance dashboard.

Run:  streamlit run dashboard_sfdc.py
"""

import io, re, warnings
from datetime import datetime
from pathlib import Path

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


def sf_run_report(sf, report_id):
    """GET Analytics API; return (DataFrame, row_count)."""
    result = sf.restful(
        path=f"analytics/reports/{report_id}",
        method="GET",
        params={"includeDetails": "true"},
    )
    rpt_type = result.get("reportMetadata", {}).get("reportFormat", "TABULAR")
    if rpt_type not in ("TABULAR",):
        # Summary / matrix: fall back to tabular columns
        pass
    meta   = result.get("reportMetadata", {})
    ext    = result.get("reportExtendedMetadata", {})
    fmap   = result.get("factMap", {})
    cols   = meta.get("detailColumns", [])
    cinfo  = ext.get("detailColumnInfo", {})
    labels = [cinfo.get(c, {}).get("label", c) for c in cols]
    rows   = fmap.get("T!T", {}).get("rows", [])
    records = []
    for row in rows:
        cells = row.get("dataCells", [])
        records.append({labels[i]: _cell_val(cells[i]) for i in range(min(len(labels),len(cells)))})
    df = pd.DataFrame(records)
    return df, len(records)


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
# CALCULATION ENGINE  (ported from build_recap_july.py)
# ─────────────────────────────────────────────────────────────────────────────
def run_calc(cw_raw, ltc_raw, ret_raw, comp_raw,
             ldr_quota_map, rep_quota_map, rep_leader_map, rep_team_map,
             seg_pl, month_nums, year, monthly_factor=0.066687):
    if isinstance(month_nums, int):
        month_nums = [month_nums]
    month_names = [MONTH_NAMES[m - 1] for m in month_nums]

    # ── Filter by period ─────────────────────────────────────────────────────
    cw = cw_raw.copy()
    cw["Close Date"] = pd.to_datetime(cw["Close Date"], errors="coerce")
    cw = cw[(cw["Close Date"].dt.month.isin(month_nums)) &
            (cw["Close Date"].dt.year  == year)].copy()
    cw["Opportunity Owner"] = cw["Opportunity Owner"].apply(normalize)
    mgr_col_cw = "Oppty Manager" if "Oppty Manager" in cw.columns else "Opportunity Owner: Manager"
    cw[mgr_col_cw] = cw[mgr_col_cw].apply(normalize)

    ltc = ltc_raw.copy()
    ltc["Close Date"] = pd.to_datetime(ltc["Close Date"], errors="coerce")
    ltc = ltc[(ltc["Close Date"].dt.month.isin(month_nums)) &
              (ltc["Close Date"].dt.year  == year)].copy()
    ltc["Opportunity Owner"] = ltc["Opportunity Owner"].apply(normalize)
    mgr_col_ltc = "Oppty Manager" if "Oppty Manager" in ltc.columns else "Opportunity Owner: Manager"
    ltc[mgr_col_ltc] = ltc[mgr_col_ltc].apply(normalize)

    ret = ret_raw.copy()
    ret = ret[ret["Final Month Closed"].isin(month_names)].copy()
    ret["Opportunity Owner"] = ret["Opportunity Owner"].apply(normalize)
    mgr_col_ret = "Oppty Manager" if "Oppty Manager" in ret.columns else "Opportunity Owner: Manager"
    ret[mgr_col_ret] = ret[mgr_col_ret].apply(normalize)
    ret["BMI_ARR"]     = pd.to_numeric(ret["BMI Sales ARR"], errors="coerce").fillna(0)
    ret["Roll_SC_Ret"] = pd.to_numeric(
        ret["Roll-up Sales Credit Calculation (converted)"], errors="coerce").fillna(0)
    split_mask = ret.get("ARR Disputes", pd.Series([""] * len(ret))).fillna("")
    split_mask = split_mask.str.contains("Split Opportunity", case=False, na=False)
    ret = ret[~split_mask].copy()

    comp = comp_raw.copy()
    comp["Close Month"] = pd.to_datetime(comp["Close Month"], errors="coerce")
    comp = comp[(comp["Close Month"].dt.month.isin(month_nums)) &
                (comp["Close Month"].dt.year  == year)].copy()
    comp["Opportunity Owner"] = comp["Opportunity Owner"].apply(normalize)
    comp["Complete_Credit_Val"] = pd.to_numeric(
        comp["Roll-up Sales Credit Calculation (converted)"], errors="coerce").fillna(0)

    # ── Lookup tables ────────────────────────────────────────────────────────
    complete_names  = set(comp["Opportunity Name"].str.strip())
    complete_lookup = dict(zip(comp["Opportunity Name"].str.strip(), comp["Complete_Credit_Val"]))

    ltc["LTC_Uplift_Calc"] = ltc.apply(
        lambda r: pd.to_numeric(r["Forecast Amount"], errors="coerce") * ltc_rate(r["Term (no. of months)"]), axis=1)
    ltc_lookup = dict(zip(ltc["Opportunity Name"].str.strip(), ltc["LTC_Uplift_Calc"]))

    ret_grp = ret.groupby("Opportunity ID").agg(
        BMI_SUM  = ("BMI_ARR",     "sum"),
        SC_SUM   = ("Roll_SC_Ret", "sum"),
        Rep      = ("Opportunity Owner", "first"),
        Leader   = (mgr_col_ret,        "first"),
        Team     = ("Oppty Team",        "first"),
        OppName  = ("Opportunity Name",  "first"),
    ).reset_index()
    ret_grp["Retention_Credit"] = (ret_grp["BMI_SUM"] - ret_grp["SC_SUM"]).clip(lower=0)
    ret_by_oppname = ret_grp.groupby("OppName")["Retention_Credit"].sum().to_dict()

    # ── Master dataset ───────────────────────────────────────────────────────
    master = cw.copy()
    master["Rep"]    = master["Opportunity Owner"]          # normalized above
    master["Leader"] = master[mgr_col_cw]                  # normalized above
    master["Region"] = master["Oppty Team"] if "Oppty Team" in master.columns else "Unknown"
    master["Segment"]= master["Oppty Team"].apply(seg_from_team)
    master["VP"]     = master["Segment"].map(VP_MAP).fillna("N/A")
    master["Forecast_Amount_ARR"] = pd.to_numeric(master["Forecast Amount"], errors="coerce").fillna(0)
    master["_OppName"] = master["Opportunity Name"].str.strip()
    master["In_Complete"] = master["_OppName"].isin(complete_names).astype(int)
    master["Complete_Credit"]  = master["_OppName"].map(complete_lookup).fillna(0)
    master["CW_ARR_Adjusted"]  = np.where(master["In_Complete"] == 1, 0, master["Forecast_Amount_ARR"])
    master["LTC_Uplift"]       = master["_OppName"].map(ltc_lookup).fillna(0)
    master["Retention_Credit"] = master["_OppName"].map(ret_by_oppname).fillna(0)
    master["Total_Credited"]   = (master["CW_ARR_Adjusted"] + master["LTC_Uplift"]
                                  + master["Retention_Credit"] + master["Complete_Credit"])

    # ── Standalone retention (not matched to CW) ─────────────────────────────
    cw_names    = set(master["_OppName"])
    standalone  = ret_grp[~ret_grp["OppName"].isin(cw_names)]
    total_ret   = ret_grp["Retention_Credit"].sum()

    # ── Rep aggregation ──────────────────────────────────────────────────────
    cw_by_rep = master.groupby("Rep").agg(
        Leader_cw       = ("Leader",          "first"),
        Region_cw       = ("Region",          "first"),
        Segment_cw      = ("Segment",         "first"),
        VP              = ("VP",              "first"),
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
        for src in [row.get("Leader_cw",""), row.get("Leader_ret","")]:
            if isinstance(src, str) and src.strip() not in ("","nan","Unknown","NaN"):
                return src.strip()
        return rep_leader_map.get(row["Rep"], "Unknown")

    def resolve_segment(row):
        for src in [row.get("Segment_cw",""), seg_from_team(row.get("Team_ret",""))]:
            if isinstance(src, str) and src not in ("","Unknown","nan"):
                return src
        return seg_from_team(rep_team_map.get(row["Rep"], ""))

    def resolve_region(row):
        for src in [row.get("Region_cw",""), row.get("Team_ret","")]:
            if isinstance(src, str) and src.strip() not in ("","nan","Unknown","NaN"):
                return src.strip()
        return rep_team_map.get(row["Rep"], "Unknown")

    rep_grp["Leader"]  = rep_grp.apply(resolve_leader, axis=1)
    rep_grp["Region"]  = rep_grp.apply(resolve_region, axis=1)
    rep_grp["Segment"] = rep_grp.apply(resolve_segment, axis=1)
    rep_grp["VP"]      = rep_grp["Segment"].map(VP_MAP).fillna("N/A")
    rep_grp["Total_Credited"] = (rep_grp["CW_ARR"] + rep_grp["LTC_Credit"]
                                 + rep_grp["Retention_Credit"] + rep_grp["Complete_Credit"])
    rep_grp["Monthly_Quota"] = rep_grp["Rep"].map(
        lambda r: rep_quota_map.get(r, 0) * monthly_factor)
    rep_grp["Pct_to_Linearity"] = np.where(
        rep_grp["Monthly_Quota"] > 0,
        rep_grp["Total_Credited"] / rep_grp["Monthly_Quota"], np.nan)
    rep_grp = rep_grp[["Rep","Leader","Region","Segment","VP",
                        "CW_ARR","LTC_Credit","Retention_Credit","Complete_Credit",
                        "Total_Credited","Monthly_Quota","Pct_to_Linearity","CW_Units"]]

    # ── Leader aggregation ───────────────────────────────────────────────────
    ldr_grp = rep_grp.groupby("Leader").agg(
        Segment         = ("Segment",          "first"),
        Region          = ("Region",           "first"),
        VP              = ("VP",               "first"),
        CW_ARR          = ("CW_ARR",           "sum"),
        LTC_Credit      = ("LTC_Credit",       "sum"),
        Retention_Credit= ("Retention_Credit", "sum"),
        Complete_Credit = ("Complete_Credit",  "sum"),
        Total_Credited  = ("Total_Credited",   "sum"),
        CW_Units        = ("CW_Units",         "sum"),
    ).reset_index()
    ldr_grp["Monthly_Quota"] = ldr_grp["Leader"].map(lambda l: ldr_quota_map.get(l, 0))
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
    seg_grp["PL_Quota"]   = seg_grp["Segment"].map(seg_pl)
    seg_grp["Lin_Target"] = seg_grp["Segment"].map(
        lambda s: sum(ldr_quota_map.get(l,0) for l,sg in LEADER_SEGMENT.items() if sg==s))
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

    return {"org": org, "segment": seg_grp, "leader": ldr_grp, "rep": rep_grp,
            "dist": dist, "master": master}


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────────────────────
for k, v in [("sf", None), ("data", None), ("last_run", None),
              ("raw_cw", None), ("raw_ltc", None), ("raw_ret", None), ("raw_comp", None),
              ("tgt_bytes", None), ("monthly_factor", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

# period_label / sel_region defaults (overwritten by sidebar widgets each run)
period_label = "—"
sel_region   = "All Regions"

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
            with st.spinner("Running Salesforce reports…"):
                try:
                    sf = st.session_state.sf
                    cw_df,  n1 = sf_run_report(sf, rpt_cw)
                    ltc_df, n2 = sf_run_report(sf, rpt_ltc)
                    ret_df, n3 = sf_run_report(sf, rpt_ret)
                    comp_df,n4 = sf_run_report(sf, rpt_comp)
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

                    # ── Column preview (helps catch naming mismatches) ────────
                    with st.expander("Report column preview (click to inspect)"):
                        for label, df in [("CW ARR", cw_df),("LTC", ltc_df),
                                          ("Retention", ret_df),("Complete", comp_df)]:
                            st.markdown(f"**{label}** — {len(df)} rows")
                            if len(df) > 0:
                                st.write(list(df.columns))
                                st.dataframe(df.head(2), use_container_width=True)
                            else:
                                st.caption("No rows returned.")

                except Exception as e:
                    st.error(f"Report error: {e}")
                    st.stop()

            with st.spinner("Calculating metrics…"):
                try:
                    lq, rq, rlm, rtm, sp, mf = load_quotas(
                        st.session_state.tgt_bytes, month_nums)
                    st.session_state.monthly_factor = mf
                    result = run_calc(
                        st.session_state.raw_cw,  st.session_state.raw_ltc,
                        st.session_state.raw_ret, st.session_state.raw_comp,
                        lq, rq, rlm, rtm, sp, month_nums, int(sel_year),
                        monthly_factor=mf)
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
        regions_avail = ["All Regions"] + sorted(
            _rep_seg["Region"].dropna().replace("Unknown","").str.strip()
            .loc[lambda s: s != ""].unique().tolist())
        sel_region = st.selectbox("Region", regions_avail)

        _rep_reg = _rep_seg[_rep_seg["Region"]==sel_region] if sel_region != "All Regions" else _rep_seg
        ldrs_avail = ["All Leaders"] + sorted(_rep_reg["Leader"].dropna().unique().tolist())
        sel_ldr = st.selectbox("Leader", ldrs_avail)
    else:
        sel_seg    = "All Segments"
        sel_region = "All Regions"
        sel_ldr    = "All Leaders"

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

# ── Row 5: Segment breakdown table ─────────────────────────────────────────
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

# ── Footer ─────────────────────────────────────────────────────────────────
row_counts = org.get("Row_Counts", {})
rc_str = " | ".join(f"{k}: {v} rows" for k,v in row_counts.items())
st.markdown(
    f"<div style='text-align:center;color:{C['dark_grey']};font-size:11px;margin-top:16px'>"
    f"SAP Concur US SMB Client Sales &nbsp;|&nbsp; {period_label} &nbsp;|&nbsp; "
    f"Source: Salesforce Reports &nbsp;|&nbsp; {rc_str}"
    f"</div>",
    unsafe_allow_html=True)
