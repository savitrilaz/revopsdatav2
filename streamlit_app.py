"""
RevOps Program Dashboard
Read-only live feed from OneDrive Excel.
Refreshes twice daily (cache TTL = 43200 seconds = 12 hours).
Edit data directly in the Excel file — dashboard picks it up automatically.
"""
import re
from collections import defaultdict
from io import BytesIO
from datetime import datetime
import hashlib

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

# ─────────────────────────────────────────────────────────────
ONEDRIVE_FILE_URL = (
    "https://emerson-my.sharepoint.com/:x:/p/savitri_lazarus/"
    "IQAQPOe1joHSTopYQHg4L61vAdgWzYvAdfVUHhZGNiI6TAM?e=YsNeJD"
)
REFRESH_TTL = 43200   # 12 hours = twice daily
EDIT_LINK   = ONEDRIVE_FILE_URL  # click-through to edit in Excel
# ─────────────────────────────────────────────────────────────

# ── PALETTE ───────────────────────────────────────────────────
C = dict(
    navy="#1B2552", blue="#004B8D", teal="#00AD7C", lblue="#1DB1DE",
    sgreen="#7CCF8B", lgreen="#75D3EB", green="#00573D", gray="#9FA1A4",
    red="#C0392B", amber="#D97706", white="#FFFFFF", bg="#F4F6FA",
)
PALETTE = [C["blue"],C["lblue"],C["teal"],C["sgreen"],C["navy"],C["lgreen"],C["green"],C["gray"]]
STATUS_COLOR = {
    "Delayed":     C["red"],
    "At Risk":     C["amber"],
    "On Track":    C["teal"],
    "Active":      C["blue"],
    "In Progress": C["lblue"],
    "Complete":    C["sgreen"],
    "Completed":   C["sgreen"],
    "Not Started": C["gray"],
    "Planning":    C["lgreen"],
}
BAND_COLOR = {
    "Top Priority":    C["teal"],
    "Middle Priority": C["blue"],
    "Lower Priority":  C["lblue"],
    "N/A":             C["gray"],
}

# ── PAGE CONFIG ───────────────────────────────────────────────
st.set_page_config(
    page_title="RevOps Program Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{{font-family:'DM Sans',sans-serif;background:{C['bg']};}}
.main{{background:{C['bg']};}}
.block-container{{padding:1.2rem 1.8rem 2rem;max-width:1520px;}}

/* ── KPI cards ── */
.kpi{{background:{C['white']};border-radius:10px;padding:16px 18px 12px;
  border:1px solid #E2E8F2;box-shadow:0 1px 6px rgba(0,0,0,0.05);}}
.kpi-val{{font-size:30px;font-weight:700;line-height:1;margin:6px 0 4px;}}
.kpi-lbl{{font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['gray']};}}
.kpi-sub{{font-size:10px;color:{C['gray']};margin-top:3px;}}
.kpi-bar{{height:3px;border-radius:2px;margin-bottom:8px;}}

/* ── Section title ── */
.sec{{font-size:11px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;
  color:{C['navy']};border-bottom:2px solid {C['blue']};padding-bottom:5px;
  display:inline-block;margin-bottom:12px;}}

/* ── Status badge ── */
.badge{{display:inline-block;padding:2px 8px;border-radius:12px;
  font-size:10px;font-weight:700;letter-spacing:.03em;}}

/* ── Project row card ── */
.prow{{background:{C['white']};border-radius:8px;padding:12px 16px;
  border:1px solid #E2E8F2;margin-bottom:7px;cursor:pointer;
  transition:box-shadow .15s;}}
.prow:hover{{box-shadow:0 3px 12px rgba(0,0,0,0.1);}}
.prow-id{{font-family:'DM Mono',monospace;font-size:10px;color:{C['gray']};}}
.prow-name{{font-size:13px;font-weight:600;color:{C['navy']};}}
.prow-desc{{font-size:11px;color:#555;margin-top:2px;line-height:1.4;}}

/* ── Detail panel ── */
.detail-hdr{{background:linear-gradient(135deg,{C['navy']} 0%,{C['blue']} 100%);
  border-radius:10px;padding:18px 22px;color:white;margin-bottom:14px;}}
.detail-field{{margin-bottom:10px;}}
.detail-lbl{{font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
  color:{C['gray']};margin-bottom:2px;}}
.detail-val{{font-size:13px;font-weight:500;color:{C['navy']};}}
.detail-val.ph{{color:{C['gray']};font-style:italic;font-weight:400;}}

/* ── Edit callout ── */
.edit-cta{{background:#EFF6FF;border:1px solid {C['lblue']};border-radius:8px;
  padding:12px 16px;font-size:12px;color:{C['navy']};margin:12px 0;}}

/* ── Divider ── */
hr.slim{{border:none;border-top:1px solid #E2E8F2;margin:20px 0;}}

/* ── Tab styling ── */
.stTabs [data-baseweb="tab-list"]{{gap:4px;background:#EAEEF4;border-radius:8px;padding:4px;}}
.stTabs [data-baseweb="tab"]{{border-radius:6px;font-size:12px;font-weight:500;padding:5px 14px;}}
.stTabs [aria-selected="true"]{{background:{C['white']};color:{C['navy']};}}

/* ── Expander ── */
.streamlit-expanderHeader{{font-size:12px;font-weight:600;}}

/* ── Refresh banner ── */
.refresh-note{{font-size:10px;color:{C['gray']};text-align:right;
  padding:4px 0 0;letter-spacing:.03em;}}

/* Hide default streamlit chrome */
#MainMenu{{visibility:hidden;}}footer{{visibility:hidden;}}header{{visibility:hidden;}}
</style>
""", unsafe_allow_html=True)

# ── HELPERS ───────────────────────────────────────────────────
def nc(df, *candidates):
    """Find first matching column name, including fuzzy."""
    for c in candidates:
        if c in df.columns: return c
    for c in candidates:
        for col in df.columns:
            if c.lower().replace(" ","").replace("?","") in col.lower().replace(" ","").replace("?",""):
                return col
    return None

def build_dl_url(url):
    if "/:x:/p/" in url or "/:x:/s/" in url:
        return url + ("&" if "?" in url else "?") + "download=1"
    if "_layouts/15/Doc.aspx" in url:
        m = re.search(r'sourcedoc=%7B([^%]+)%7D', url, re.I)
        if m:
            base = url.split("/_layouts/")[0]
            return f"{base}/_layouts/15/download.aspx?UniqueId={m.group(1)}"
    return url + ("&" if "?" in url else "?") + "download=1"

def normalize_cdm(v):
    if pd.isna(v) or str(v).strip() == "": return "Unknown"
    s = str(v).strip().upper()
    if s in ("Y","YES","TRUE","1"): return "Yes"
    if s in ("N","NO","FALSE","0"): return "No"
    return "Unknown"

def status_badge(s, size=10):
    cm = {
        "delayed":     ("#FEE2E2","#C0392B"),
        "at risk":     ("#FEF3C7","#D97706"),
        "on track":    ("#D1FAE5","#065F46"),
        "active":      ("#DBEAFE","#1E40AF"),
        "in progress": ("#E0F2FE","#0369A1"),
        "complete":    ("#D1FAE5","#065F46"),
        "completed":   ("#D1FAE5","#065F46"),
        "not started": ("#F3F4F6","#374151"),
        "planning":    ("#EDE9FE","#5B21B6"),
    }
    bg, fg = cm.get(str(s).lower(), ("#F3F4F6","#374151"))
    return f"<span class='badge' style='background:{bg};color:{fg};font-size:{size}px'>{s}</span>"

def cdm_badge(v):
    if v == "Yes":
        return "<span class='badge' style='background:#FEF3C7;color:#D97706'>⚠ CDM</span>"
    return "<span style='color:#9FA1A4;font-size:10px'>—</span>"

def fmt_val(v, prefix="$"):
    try:
        n = float(v)
        if n >= 1_000_000: return f"{prefix}{n/1_000_000:.1f}M"
        if n >= 1_000:     return f"{prefix}{n/1_000:.0f}K"
        return f"{prefix}{n:.0f}"
    except Exception:
        return str(v) if v else "—"

def chart_base(fig, height=260):
    fig.update_layout(
        height=height, margin=dict(t=12,b=12,l=8,r=8),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="DM Sans", size=11, color="#374151"),
        showlegend=False,
        xaxis=dict(gridcolor="#F0F4F8",linecolor="#E2E8F2",tickfont=dict(size=10)),
        yaxis=dict(gridcolor="#F0F4F8",linecolor="#E2E8F2",tickfont=dict(size=10)),
    )
    fig.update_traces(marker_line_width=0)
    return fig

# ── DATA LOAD (cached 12 hours) ───────────────────────────────
@st.cache_data(ttl=REFRESH_TTL, show_spinner=False)
def load_all(url):
    try:
        dl = build_dl_url(url)
        r = requests.get(dl, headers={"User-Agent":"Mozilla/5.0"}, timeout=30, allow_redirects=True)
        r.raise_for_status()
        if "html" in r.headers.get("Content-Type","").lower():
            fb = url + ("&download=1" if "?" in url else "?download=1")
            r = requests.get(fb, headers={"User-Agent":"Mozilla/5.0"}, timeout=30, allow_redirects=True)
            r.raise_for_status()
        buf = BytesIO(r.content)
        sheets = {}
        for s in ["Projects","Project_Resources","Dependencies","Project_Value_Map","Value_Category_Dictionary"]:
            try:
                df = pd.read_excel(buf, sheet_name=s, engine="openpyxl")
                df.columns = [c.strip() for c in df.columns]
                sheets[s] = df
            except Exception:
                sheets[s] = None
        return sheets, None, datetime.now()
    except Exception as e:
        return None, str(e), datetime.now()

# ── LOAD ──────────────────────────────────────────────────────
with st.spinner("Loading portfolio data…"):
    sheets, err, loaded_at = load_all(ONEDRIVE_FILE_URL)

if err:
    st.error(f"Could not load data: {err}")
    st.info("Check that the SharePoint link is set to 'Anyone with the link can view'.")
    st.stop()

proj_df = sheets["Projects"]
res_df  = sheets["Project_Resources"]
dep_df  = sheets["Dependencies"]
vm_df   = sheets["Project_Value_Map"]
vd_df   = sheets["Value_Category_Dictionary"]

if proj_df is None:
    st.error("Projects sheet missing."); st.stop()

# ── COLUMN MAP ────────────────────────────────────────────────
pid_c    = nc(proj_df,"Project ID","ProjectID","ID")
name_c   = nc(proj_df,"Project","Project Name","Name")
rank_c   = nc(proj_df,"Priority Rank","PriorityRank","Rank")
type_c   = nc(proj_df,"Priority Type","ProjectType","Type")
strat_c  = nc(proj_df,"Strategic Priority","FLMC Tag","Strategic")
owner_c  = nc(proj_df,"Owner","owner")
core_c   = nc(proj_df,"Core Team","CoreTeam","Requested By")
status_c = nc(proj_df,"Status","status")
cycle_c  = nc(proj_df,"Cycle","cycle")
effort_c = nc(proj_df,"Effort","effort")
impact_c = nc(proj_df,"Impact","impact")
invest_c = nc(proj_df,"Investment","investment")
delay_c  = nc(proj_df,"Delayed Flag","Delayed","delay_flag")
deli_c   = nc(proj_df,"If Delayed Impact","Delayed Impact")
band_c   = nc(proj_df,"Priority Band","Band")
cdm_c    = nc(proj_df,"CDM Dependency Flag","CDM Dependency","CDM")
bizprog_c= nc(proj_df,"Business Program","BizProg")
bv_c     = nc(proj_df,"Business Value ($)","Business Value","BizValue")
dar_c    = nc(proj_df,"Dollars at Risk ($)","Dollars at Risk","DAR")
rawval_c = nc(proj_df,"Raw Value Description","RawValue")
valgrp_c = nc(proj_df,"Value Groups","ValueGroups")
valcat_c = nc(proj_df,"Value Categories","ValueCategories")

# Resources
res_pid_c  = nc(res_df,"Project ID","ProjectID","ID") if res_df is not None else None
res_team_c = nc(res_df,"Team","team") if res_df is not None else None

# Dependencies
dep_pid_c = nc(dep_df,"Project ID","ProjectID","ID") if dep_df is not None else None
dep_on_c  = nc(dep_df,"Depends On Project ID","DependsOn","dependency") if dep_df is not None else None

# Normalise CDM
CDM = "__cdm__"
proj_df[CDM] = proj_df[cdm_c].apply(normalize_cdm) if cdm_c else "Unknown"

# Resource lookup
proj_teams = defaultdict(list)
if res_df is not None and res_pid_c and res_team_c:
    for _, row in res_df.iterrows():
        if pd.notna(row[res_pid_c]) and pd.notna(row[res_team_c]):
            proj_teams[str(row[res_pid_c])].append(str(row[res_team_c]))

# Dependency lookup
proj_deps = defaultdict(list)
if dep_df is not None and dep_pid_c and dep_on_c:
    for _, row in dep_df.iterrows():
        if pd.notna(row[dep_pid_c]) and pd.notna(row[dep_on_c]):
            raw = str(row[dep_on_c])
            deps = [d.strip() for d in raw.replace(";",",").split(",") if d.strip()]
            proj_deps[str(row[dep_pid_c])].extend(deps)

# ── FILTER TO REVOPS ──────────────────────────────────────────
revops_df = proj_df[proj_df[owner_c].str.strip() == "RevOps"].copy() if owner_c else proj_df.copy()

# ── SHARED METRICS ────────────────────────────────────────────
total      = len(revops_df)
delayed_m  = revops_df[delay_c].str.strip().str.upper() == "Y" if delay_c else pd.Series([False]*total, index=revops_df.index)
delayed_n  = int(delayed_m.sum())
strategic_n= int((revops_df[type_c].str.strip() == "Strategic").sum()) if type_c else 0
sustaining_n=int((revops_df[type_c].str.strip() == "Sustaining").sum()) if type_c else 0
cdm_yes_n  = int((revops_df[CDM] == "Yes").sum())
flmc_n     = int(revops_df[strat_c].str.contains("FLMC",na=False).sum()) if strat_c else 0

top_total  = int((revops_df[band_c] == "Top Priority").sum()) if band_c else 0
top_delayed= int(((revops_df[band_c] == "Top Priority") & delayed_m).sum()) if band_c else 0
top_ok     = top_total - top_delayed
strat_del  = int(((revops_df[type_c] == "Strategic") & delayed_m).sum()) if type_c else 0

bv_total   = None
dar_total  = None
if bv_c:
    v = pd.to_numeric(revops_df[bv_c], errors="coerce")
    if v.notna().any(): bv_total = v.sum()
if dar_c:
    v = pd.to_numeric(revops_df[dar_c], errors="coerce")
    if v.notna().any(): dar_total = v.sum()

# ── MANUAL REFRESH ────────────────────────────────────────────
col_hdr, col_refresh = st.columns([8,1])
with col_hdr:
    st.markdown(f"""
    <div style='display:flex;align-items:baseline;gap:12px;margin-bottom:4px;'>
      <span style='font-size:20px;font-weight:700;color:{C['navy']};'>RevOps Program Dashboard</span>
      <span style='font-size:12px;color:{C['gray']};'>FY26 · Owner = RevOps · Read-only live view</span>
    </div>
    """, unsafe_allow_html=True)
with col_refresh:
    if st.button("↺ Refresh", help="Force reload from OneDrive (auto-refreshes every 12 hours)"):
        st.cache_data.clear()
        st.rerun()

st.markdown(f"<div class='refresh-note'>Last loaded: {loaded_at.strftime('%B %d, %Y at %I:%M %p')} · "
            f"Next auto-refresh in ~{max(0,12 - int((datetime.now()-loaded_at).seconds/3600))} hr · "
            f"<a href='{EDIT_LINK}' target='_blank' style='color:{C['blue']};text-decoration:none;font-weight:600;'>"
            f"✏️ Edit in Excel →</a></div>", unsafe_allow_html=True)

st.markdown("<hr class='slim'>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# VIEW TOGGLE
# ═══════════════════════════════════════════════════════════════
view_col, _ = st.columns([3,5])
with view_col:
    view = st.radio("", ["📋 Executive Summary","📊 Portfolio Detail","🔍 Project Explorer"],
                    horizontal=True, label_visibility="collapsed",
                    key="main_view")

st.markdown("<hr class='slim'>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# ██  EXECUTIVE SUMMARY  ██
# ═══════════════════════════════════════════════════════════════
if view == "📋 Executive Summary":

    # ── KPI Row 1 ─────────────────────────────────────────────
    k = st.columns(6)
    kpi_data = [
        ("Total Projects",   str(total),          "Owner = RevOps",            C["navy"],    ""),
        ("Strategic",        str(strategic_n),    "Project type: Strategic",   C["teal"],    ""),
        ("Sustaining",       str(sustaining_n),   "Project type: Sustaining",  C["lblue"],   ""),
        ("Delayed",          str(delayed_n),       f"{round(delayed_n/total*100)}% of portfolio",
         C["red"] if delayed_n else C["gray"], "⚠ " if delayed_n else ""),
        ("CDM Dependent",    str(cdm_yes_n),      "Blocked pending P11 CDM",   C["amber"],   ""),
        ("FLMC SoaP Aligned",str(flmc_n),         "FLMC Strategy on a Page",   C["navy"],    ""),
    ]
    for col, (lbl, val, sub, color, prefix) in zip(k, kpi_data):
        col.markdown(f"""
        <div class='kpi'>
          <div class='kpi-bar' style='background:{color}'></div>
          <div class='kpi-lbl'>{lbl}</div>
          <div class='kpi-val' style='color:{color}'>{prefix}{val}</div>
          <div class='kpi-sub'>{sub}</div>
        </div>""", unsafe_allow_html=True)

    # ── KPI Row 2 — Cost of inaction ──────────────────────────
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    pct_delayed = round(delayed_n/total*100) if total else 0
    top_ok_str  = f"1 of {top_total}" if top_total else "—"

    st.markdown(f"""
    <div style='background:linear-gradient(135deg,{C['navy']} 0%,{C['blue']} 100%);
      border-radius:10px;padding:14px 20px;margin-bottom:4px;'>
      <div style='font-size:10px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;
        color:rgba(255,255,255,0.6);margin-bottom:10px;'>
        COST OF INACTION — IF NO DECISIONS ARE MADE THIS QUARTER
      </div>
      <div style='display:flex;gap:0;'>
    """, unsafe_allow_html=True)

    inaction = [
        (f"{pct_delayed}%",  "of RevOps portfolio delayed",   f"{delayed_n} of {total} programs", C["red"]),
        (top_ok_str,          "Top Priority programs on track", f"{top_delayed} of {top_total} already delayed", C["amber"]),
        (str(cdm_yes_n),      "programs blocked on CDM (P11)", "Cannot start or advance",           C["amber"]),
        (str(strat_del),      "Strategic programs slipping",   f"{strat_del} of {strategic_n} Strategic delayed", C["red"]),
    ]
    cols = st.columns(4)
    for col, (val, lbl, sub, color) in zip(cols, inaction):
        col.markdown(f"""
        <div style='background:rgba(255,255,255,0.08);border-radius:8px;padding:12px 14px;margin:0 4px;'>
          <div style='font-size:26px;font-weight:700;color:{color};line-height:1;'>{val}</div>
          <div style='font-size:11px;font-weight:600;color:rgba(255,255,255,0.85);margin-top:6px;'>{lbl}</div>
          <div style='font-size:10px;color:rgba(255,255,255,0.5);margin-top:3px;'>{sub}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<hr class='slim'>", unsafe_allow_html=True)

    # ── Charts row ────────────────────────────────────────────
    st.markdown("<div class='sec'>Executive Overview</div>", unsafe_allow_html=True)
    ch1, ch2, ch3 = st.columns(3)

    with ch1:
        st.caption("**Projects by Status**")
        if status_c:
            sc = revops_df[status_c].value_counts().reset_index()
            sc.columns = ["Status","Count"]
            colors = [STATUS_COLOR.get(s, C["gray"]) for s in sc["Status"]]
            fig = px.bar(sc, x="Status", y="Count", color="Status",
                         color_discrete_map=STATUS_COLOR, template="plotly_white")
            fig.update_layout(showlegend=False)
            st.plotly_chart(chart_base(fig,240), use_container_width=True)

    with ch2:
        st.caption("**Projects by Priority Band**")
        if band_c:
            bc = revops_df[band_c].value_counts().reset_index()
            bc.columns = ["Band","Count"]
            colors = [BAND_COLOR.get(b, C["gray"]) for b in bc["Band"]]
            fig2 = go.Figure(go.Bar(
                x=bc["Count"], y=bc["Band"], orientation="h",
                marker_color=colors))
            st.plotly_chart(chart_base(fig2,240), use_container_width=True)

    with ch3:
        st.caption("**CDM Dependency**")
        cdm_c2 = revops_df[CDM].value_counts().reset_index()
        cdm_c2.columns = ["CDM","Count"]
        cdm_colors = {"Yes":C["amber"],"No":C["teal"],"Unknown":C["gray"]}
        fig3 = px.pie(cdm_c2, names="CDM", values="Count",
                      color="CDM", color_discrete_map=cdm_colors,
                      hole=0.55, template="plotly_white")
        fig3.update_traces(textposition="outside", textfont_size=10,
                           marker=dict(line=dict(color="white",width=2)))
        fig3.update_layout(height=240, margin=dict(t=8,b=8,l=8,r=8),
                           showlegend=True, legend=dict(font=dict(size=10)),
                           paper_bgcolor="white")
        st.plotly_chart(fig3, use_container_width=True)

    # ── CDM Callout ───────────────────────────────────────────
    st.markdown(f"""
    <div style='background:#FFF8F0;border:1px solid {C['amber']};border-left:4px solid {C['amber']};
      border-radius:8px;padding:14px 18px;margin:8px 0;'>
      <div style='font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
        color:{C['amber']};margin-bottom:6px;'>CDM AS A CASE STUDY IN DELAY RISK</div>
      <div style='font-size:12px;color:#1a1a1a;line-height:1.6;'>
        <strong>P11 (CDM DAUT ID Replacement) is currently Delayed.</strong>
        It is the upstream dependency for 8 other projects: P27 (SFDC Daily Sync), P20 (CPQ/Revenue Cloud),
        P29 (SFDC DataCloud), P7 (Forecasting), P15 (MTM Expansion), P22 (Sales Planning),
        P38 (SYSS/MAS), and P60 (AFAG).<br>
        <span style='color:{C['red']};font-weight:600;'>
        Every month CDM slips, it delays quoting, forecasting, pricing, and customer data integrity
        across the commercial stack.</span>
      </div>
    </div>""", unsafe_allow_html=True)

    # ── Spotlight table ───────────────────────────────────────
    st.markdown("<hr class='slim'>", unsafe_allow_html=True)
    st.markdown("<div class='sec'>Project Spotlight — Top & Delayed</div>",
                unsafe_allow_html=True)
    st.caption("Ranked by delay status + impact score. Edit values directly in Excel.")

    spot = revops_df.copy()
    spot["__score"] = 0
    if delay_c:
        spot["__score"] += (spot[delay_c].str.strip().str.upper()=="Y").astype(int)*10
    if impact_c:
        spot["__score"] += pd.to_numeric(spot[impact_c], errors="coerce").fillna(0)
    if bv_c:
        spot["__score"] += pd.to_numeric(spot[bv_c], errors="coerce").fillna(0)/1000
    spot = spot.sort_values("__score", ascending=False).head(20)

    disp_cols = [c for c in [pid_c, name_c, band_c, status_c, type_c, CDM, bv_c, dar_c] if c]
    disp = spot[disp_cols].rename(columns={CDM:"CDM Dep"})
    if bv_c in disp.columns:  disp = disp.rename(columns={bv_c:"Biz Value ($)"})
    if dar_c in disp.columns: disp = disp.rename(columns={dar_c:"$ at Risk"})
    st.dataframe(disp.reset_index(drop=True), use_container_width=True, hide_index=True,
                 height=340)

    # ── Edit CTA ──────────────────────────────────────────────
    st.markdown(f"""
    <div class='edit-cta'>
      ✏️ <strong>Want to update status, scores, or add Business Value / $ at Risk?</strong>
      Open the source Excel file — the dashboard refreshes automatically twice daily.
      &nbsp;&nbsp;<a href='{EDIT_LINK}' target='_blank'
        style='color:{C['blue']};font-weight:700;'>Open Excel →</a>
    </div>""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# ██  PORTFOLIO DETAIL  ██
# ═══════════════════════════════════════════════════════════════
elif view == "📊 Portfolio Detail":

    # ── Sidebar-style filters in a horizontal strip ────────────
    st.markdown("<div class='sec'>Filters</div>", unsafe_allow_html=True)
    f1,f2,f3,f4,f5 = st.columns(5)

    band_opts = ["All"] + sorted([b for b in revops_df[band_c].dropna().unique() if b] ) if band_c else ["All"]
    status_opts = ["All"] + sorted([s for s in revops_df[status_c].dropna().unique() if s]) if status_c else ["All"]
    type_opts_f = ["All"] + sorted([t for t in revops_df[type_c].dropna().unique() if t]) if type_c else ["All"]
    cdm_opts_f  = ["All","Yes","No","Unknown"]
    core_opts   = ["All"] + sorted([c for c in revops_df[core_c].dropna().unique() if c]) if core_c else ["All"]

    sel_band   = f1.selectbox("Priority Band",  band_opts,    key="pb")
    sel_status = f2.selectbox("Status",         status_opts,  key="ps")
    sel_type   = f3.selectbox("Project Type",   type_opts_f,  key="pt")
    sel_cdm    = f4.selectbox("CDM Dependency", cdm_opts_f,   key="pc")
    sel_core   = f5.selectbox("Requested By",   core_opts,    key="pr")

    filt = revops_df.copy()
    if sel_band   != "All" and band_c:   filt = filt[filt[band_c]==sel_band]
    if sel_status != "All" and status_c: filt = filt[filt[status_c]==sel_status]
    if sel_type   != "All" and type_c:   filt = filt[filt[type_c]==sel_type]
    if sel_cdm    != "All":              filt = filt[filt[CDM]==sel_cdm]
    if sel_core   != "All" and core_c:   filt = filt[filt[core_c]==sel_core]

    st.caption(f"**{len(filt)}** of {total} RevOps projects shown")
    st.markdown("<hr class='slim'>", unsafe_allow_html=True)

    # ── Two charts ────────────────────────────────────────────
    ch1, ch2 = st.columns(2)
    with ch1:
        st.caption("**Status breakdown**")
        if status_c and not filt.empty:
            sc = filt[status_c].value_counts().reset_index()
            sc.columns = ["Status","Count"]
            fig = px.bar(sc.sort_values("Count",ascending=False),
                         x="Status", y="Count", color="Status",
                         color_discrete_map=STATUS_COLOR, template="plotly_white")
            fig.update_layout(showlegend=False)
            st.plotly_chart(chart_base(fig,220), use_container_width=True)

    with ch2:
        st.caption("**Effort vs Impact**")
        if effort_c and impact_c and not filt.empty:
            sdf = filt[[c for c in [pid_c,name_c,effort_c,impact_c,status_c,CDM] if c]].copy()
            sdf[effort_c] = pd.to_numeric(sdf[effort_c], errors="coerce")
            sdf[impact_c] = pd.to_numeric(sdf[impact_c], errors="coerce")
            sdf = sdf.dropna(subset=[effort_c,impact_c])
            if not sdf.empty:
                # deterministic jitter
                def jit(s, sc=0.12):
                    return s.apply(lambda v: ((int(hashlib.md5(str(v).encode()).hexdigest(),16)%1000)/1000-.5)*2*sc)
                sdf["__x"] = sdf[effort_c] + jit(sdf[effort_c])
                sdf["__y"] = sdf[impact_c] + jit(sdf[impact_c])
                hov = {c:True for c in [name_c,pid_c] if c}
                fig2 = px.scatter(sdf, x="__x", y="__y",
                                  color=status_c if status_c else None,
                                  color_discrete_map=STATUS_COLOR,
                                  hover_data=hov, template="plotly_white", opacity=0.8)
                fig2.update_traces(marker=dict(size=11,line=dict(width=1.5,color="white")))
                fig2.update_layout(xaxis_title="Effort",yaxis_title="Impact", showlegend=True,
                                   legend=dict(orientation="h",y=1.12,font=dict(size=10)))
                st.plotly_chart(chart_base(fig2,220), use_container_width=True)

    # ── Resource load chart ───────────────────────────────────
    if res_df is not None and res_pid_c and res_team_c and pid_c:
        st.caption("**Resource / Team Load (filtered projects)**")
        filt_pids = set(filt[pid_c].dropna().astype(str).unique())
        res_f = res_df[res_df[res_pid_c].astype(str).isin(filt_pids)]
        if not res_f.empty:
            tc = res_f.groupby(res_team_c)[res_pid_c].nunique().reset_index()
            tc.columns = ["Team","Projects"]
            tc = tc.sort_values("Projects", ascending=True)
            avg = tc["Projects"].mean()
            tc["Color"] = tc["Projects"].apply(
                lambda x: C["red"] if x>avg*1.4 else (C["amber"] if x>avg else C["teal"]))
            fig3 = go.Figure(go.Bar(
                x=tc["Projects"], y=tc["Team"], orientation="h",
                marker_color=tc["Color"].tolist()))
            fig3.update_layout(showlegend=False)
            st.plotly_chart(chart_base(fig3, max(220, len(tc)*28)), use_container_width=True)
            st.caption("🔴 Overloaded  🟠 Elevated  🟢 Normal  (relative to average)")

    st.markdown("<hr class='slim'>", unsafe_allow_html=True)

    # ── Value Groups ──────────────────────────────────────────
    st.markdown("<div class='sec'>Value Groups</div>", unsafe_allow_html=True)

    if vm_df is not None:
        vm_pid_c  = nc(vm_df,"Project ID","ProjectID","ID")
        vm_cat_c  = nc(vm_df,"Value Category","Category")
        vm_grp_c  = nc(vm_df,"Value Group","Group")
        if vm_pid_c and vm_grp_c:
            filt_pids = set(filt[pid_c].dropna().astype(str).unique()) if pid_c else set()
            vm_f = vm_df[vm_df[vm_pid_c].isin(filt_pids)] if filt_pids else vm_df
            grp_counts = vm_f.groupby(vm_grp_c)[vm_pid_c].nunique().reset_index()
            grp_counts.columns = ["Value Group","Projects"]
            vg_cols = st.columns(len(grp_counts))
            grp_colors = [C["teal"],C["blue"],C["lblue"],C["navy"],C["green"],C["amber"],C["gray"]]
            for i,(col,(_, row)) in enumerate(zip(vg_cols, grp_counts.iterrows())):
                clr = grp_colors[i % len(grp_colors)]
                col.markdown(f"""
                <div style='background:{C['white']};border-radius:8px;padding:12px 14px;
                  border:1px solid #E2E8F2;border-top:3px solid {clr};text-align:center;'>
                  <div style='font-size:22px;font-weight:700;color:{clr};'>{int(row['Projects'])}</div>
                  <div style='font-size:10px;font-weight:600;color:{C['navy']};margin-top:3px;'>
                    {row['Value Group']}</div>
                </div>""", unsafe_allow_html=True)

    st.markdown("<hr class='slim'>", unsafe_allow_html=True)

    # ── Full project table ────────────────────────────────────
    st.markdown("<div class='sec'>All Projects</div>", unsafe_allow_html=True)

    show_cols = [c for c in [pid_c,name_c,band_c,type_c,status_c,core_c,CDM,
                              effort_c,impact_c,bv_c,dar_c] if c]
    disp2 = filt[show_cols].rename(columns={CDM:"CDM Dep"})
    if bv_c  in disp2.columns: disp2 = disp2.rename(columns={bv_c:"Biz Value ($)"})
    if dar_c in disp2.columns: disp2 = disp2.rename(columns={dar_c:"$ at Risk"})
    st.dataframe(disp2.reset_index(drop=True), use_container_width=True,
                 hide_index=True, height=420)

    st.markdown(f"""
    <div class='edit-cta'>
      ✏️ <strong>To update any field</strong> — Status, Priority Band, CDM flag, Business Value,
      Dollars at Risk — open the Excel file. Dashboard refreshes every 12 hours automatically.
      &nbsp;&nbsp;<a href='{EDIT_LINK}' target='_blank'
        style='color:{C['blue']};font-weight:700;'>Open Excel →</a>
    </div>""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# ██  PROJECT EXPLORER  ██
# ═══════════════════════════════════════════════════════════════
else:
    st.markdown("<div class='sec'>Project Explorer</div>", unsafe_allow_html=True)

    # ── Search + filter ───────────────────────────────────────
    sc1, sc2, sc3 = st.columns([3,2,2])
    search = sc1.text_input("Search project name or ID", placeholder="Type to filter…", key="search")
    band_f = sc2.selectbox("Band", ["All"]+sorted([b for b in revops_df[band_c].dropna().unique() if b])) if band_c else "All"
    stat_f = sc3.selectbox("Status", ["All"]+sorted([s for s in revops_df[status_c].dropna().unique() if s])) if status_c else "All"

    exp_df = revops_df.copy()
    if search and pid_c and name_c:
        q = search.lower()
        exp_df = exp_df[
            exp_df[pid_c].astype(str).str.lower().str.contains(q, na=False) |
            exp_df[name_c].astype(str).str.lower().str.contains(q, na=False)
        ]
    if band_f != "All" and band_c:   exp_df = exp_df[exp_df[band_c]==band_f]
    if stat_f != "All" and status_c: exp_df = exp_df[exp_df[status_c]==stat_f]

    # Sort: Top first, then delayed
    if band_c:
        band_order = {"Top Priority":0,"Middle Priority":1,"Lower Priority":2,"N/A":3}
        exp_df = exp_df.copy()
        exp_df["__bs"] = exp_df[band_c].map(band_order).fillna(4)
        exp_df["__ds"] = (exp_df[delay_c].str.upper()=="Y").astype(int) * -1 if delay_c else 0
        exp_df = exp_df.sort_values(["__bs","__ds"])

    st.caption(f"**{len(exp_df)}** projects · click any row to see full details")

    left_col, right_col = st.columns([2,3])

    with left_col:
        if exp_df.empty:
            st.info("No projects match your filters.")
        else:
            sel_pid = None
            for _, row in exp_df.iterrows():
                pid  = str(row[pid_c]) if pid_c else "—"
                name = str(row[name_c]) if name_c else "—"
                band = str(row[band_c]) if band_c else ""
                stat = str(row[status_c]) if status_c else ""
                cdm  = row[CDM]
                is_delayed = str(row[delay_c]).strip().upper()=="Y" if delay_c else False
                card_class = "prow delayed" if is_delayed else "prow"
                bdg        = status_badge(stat)
                cdm_bdg    = cdm_badge(cdm)
                bc         = BAND_COLOR.get(band, C["gray"])
                desc       = str(row[rawval_c])[:80]+"…" if rawval_c and pd.notna(row[rawval_c]) and str(row[rawval_c]) != "None" else ""

                btn_key = f"proj_{pid}"
                clicked = st.button(
                    f"[{pid}]  {name[:42]}{'…' if len(name)>42 else ''}",
                    key=btn_key, use_container_width=True)
                if clicked:
                    st.session_state["selected_pid"] = pid
                # mini metadata under button
                st.markdown(
                    f"<div style='font-size:10px;color:{C['gray']};margin:-6px 0 6px 6px;'>"
                    f"<span style='color:{bc};font-weight:600;'>{band.replace(' Priority','')}</span>"
                    f"  ·  {bdg}  ·  {cdm_bdg}</div>",
                    unsafe_allow_html=True)

    # ── Right: detail panel ───────────────────────────────────
    with right_col:
        sel = st.session_state.get("selected_pid")
        if not sel and not exp_df.empty and pid_c:
            sel = str(exp_df.iloc[0][pid_c])
            st.session_state["selected_pid"] = sel

        if sel and pid_c:
            match = revops_df[revops_df[pid_c].astype(str)==sel]
            if match.empty:
                st.info("Select a project from the list.")
            else:
                row = match.iloc[0]
                pname   = str(row[name_c]) if name_c else sel
                pstatus = str(row[status_c]) if status_c else "—"
                pband   = str(row[band_c])   if band_c  else "—"
                ptype   = str(row[type_c])   if type_c  else "—"
                is_del  = str(row[delay_c]).strip().upper()=="Y" if delay_c else False
                pcdm    = row[CDM]
                accent  = C["red"] if is_del else (C["amber"] if pcdm=="Yes" else C["blue"])

                st.markdown(f"""
                <div class='detail-hdr' style='background:linear-gradient(135deg,{C['navy']} 0%,{accent} 100%);'>
                  <div style='font-size:11px;font-family:DM Mono;color:rgba(255,255,255,0.6);'>{sel}</div>
                  <div style='font-size:17px;font-weight:700;color:white;margin:4px 0;'>{pname}</div>
                  <div style='display:flex;gap:8px;flex-wrap:wrap;margin-top:6px;'>
                    {status_badge(pstatus,11)}
                    <span class='badge' style='background:rgba(255,255,255,0.15);color:white;font-size:10px;'>{pband}</span>
                    <span class='badge' style='background:rgba(255,255,255,0.15);color:white;font-size:10px;'>{ptype}</span>
                    {'<span class="badge" style="background:#FEF3C7;color:#D97706;font-size:10px;">⚠ CDM Dependent</span>' if pcdm=="Yes" else ''}
                    {'<span class="badge" style="background:#FEE2E2;color:#C0392B;font-size:10px;">⚠ Delayed</span>' if is_del else ''}
                  </div>
                </div>""", unsafe_allow_html=True)

                tab1, tab2, tab3, tab4 = st.tabs(["Overview","Resources","Dependencies","Risk & Value"])

                with tab1:
                    def dfield(label, col_key, fallback="Not yet captured"):
                        val = str(row[col_key]) if col_key and col_key in row.index and pd.notna(row[col_key]) and str(row[col_key]) not in ("None","nan","") else None
                        ph  = " ph" if not val else ""
                        return f"""
                        <div class='detail-field'>
                          <div class='detail-lbl'>{label}</div>
                          <div class='detail-val{ph}'>{val or fallback}</div>
                        </div>"""

                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown(
                            dfield("Business Program", bizprog_c) +
                            dfield("Core Team / Requested By", core_c) +
                            dfield("Strategic Priority", strat_c) +
                            dfield("Cycle", cycle_c) +
                            dfield("Investment", invest_c),
                            unsafe_allow_html=True)
                    with c2:
                        st.markdown(
                            dfield("Priority Rank", rank_c) +
                            dfield("Effort Score", effort_c) +
                            dfield("Impact Score", impact_c) +
                            dfield("If Delayed Impact", deli_c) +
                            dfield("Value Groups", valgrp_c),
                            unsafe_allow_html=True)

                    raw = str(row[rawval_c]) if rawval_c and pd.notna(row[rawval_c]) and str(row[rawval_c]) not in ("None","nan","") else None
                    if raw:
                        st.markdown(f"""
                        <div style='background:#F8FAFC;border-radius:6px;padding:10px 14px;
                          border-left:3px solid {C['blue']};margin-top:4px;'>
                          <div class='detail-lbl'>What This Project Delivers</div>
                          <div style='font-size:12px;color:#374151;margin-top:3px;line-height:1.5;'>{raw}</div>
                        </div>""", unsafe_allow_html=True)

                with tab2:
                    teams = proj_teams.get(sel, [])
                    if teams:
                        st.markdown(f"**{len(teams)} teams involved:**")
                        for t in sorted(teams):
                            st.markdown(f"- {t}")
                    else:
                        st.info("No resource data for this project.")

                with tab3:
                    deps = proj_deps.get(sel, [])
                    if deps:
                        st.markdown(f"**Depends on {len(deps)} project(s):**")
                        for d in deps:
                            dm = revops_df[revops_df[pid_c].astype(str)==d.strip()] if pid_c else pd.DataFrame()
                            if not dm.empty:
                                dn = str(dm.iloc[0][name_c]) if name_c else d
                                ds = str(dm.iloc[0][status_c]) if status_c else "—"
                                st.markdown(
                                    f"- **{d}** — {dn} &nbsp; {status_badge(ds)}",
                                    unsafe_allow_html=True)
                            else:
                                st.markdown(f"- {d}")
                    else:
                        st.info("No dependencies recorded for this project.")

                    # Projects that depend ON this one
                    blocking = [p for p,dlist in proj_deps.items() if sel in dlist]
                    if blocking:
                        st.markdown(f"**{len(blocking)} project(s) depend on this:**")
                        for b in blocking:
                            bm = revops_df[revops_df[pid_c].astype(str)==b] if pid_c else pd.DataFrame()
                            bn = str(bm.iloc[0][name_c]) if not bm.empty and name_c else b
                            st.markdown(f"- **{b}** — {bn}")

                with tab4:
                    r1, r2 = st.columns(2)
                    with r1:
                        def score_card(label, col_key, color):
                            v = pd.to_numeric(row[col_key], errors="coerce") if col_key else None
                            disp = str(int(v)) if v and not pd.isna(v) else "—"
                            clr  = color if v and not pd.isna(v) else C["gray"]
                            st.markdown(f"""
                            <div style='background:{C['white']};border-radius:8px;padding:14px;
                              border:1px solid #E2E8F2;margin-bottom:8px;'>
                              <div class='detail-lbl'>{label}</div>
                              <div style='font-size:28px;font-weight:700;color:{clr};'>{disp}</div>
                            </div>""", unsafe_allow_html=True)
                        score_card("Effort Score (1–5)", effort_c, C["blue"])
                        score_card("Impact Score (1–5)", impact_c, C["teal"])

                    with r2:
                        def money_card(label, col_key, color):
                            v = row[col_key] if col_key else None
                            disp = fmt_val(v) if v and pd.notna(v) and str(v) not in ("None","nan","") else "Not yet captured"
                            ph   = not (v and pd.notna(v) and str(v) not in ("None","nan",""))
                            clr  = color if not ph else C["gray"]
                            st.markdown(f"""
                            <div style='background:{C['white']};border-radius:8px;padding:14px;
                              border:1px solid #E2E8F2;margin-bottom:8px;'>
                              <div class='detail-lbl'>{label}</div>
                              <div style='font-size:20px;font-weight:700;color:{clr};
                                {'font-style:italic;font-size:14px;' if ph else ''}'>{disp}</div>
                            </div>""", unsafe_allow_html=True)
                        money_card("Business Value ($)",   bv_c,  C["teal"])
                        money_card("Dollars at Risk ($)",  dar_c, C["red"])

                    # Delay / CDM risk flags
                    if is_del:
                        st.markdown(f"""
                        <div style='background:#FEF3F2;border-left:3px solid {C['red']};
                          border-radius:0 6px 6px 0;padding:10px 14px;font-size:12px;'>
                          <strong style='color:{C['red']};'>⚠ Delay Flag Active</strong><br>
                          <span style='color:#374151;'>Review owner accountability and blockers.
                          {'This project is CDM-dependent — resolution requires P11 delivery.' if pcdm=='Yes' else ''}</span>
                        </div>""", unsafe_allow_html=True)
                    elif pcdm == "Yes":
                        st.markdown(f"""
                        <div style='background:#FFF8F0;border-left:3px solid {C['amber']};
                          border-radius:0 6px 6px 0;padding:10px 14px;font-size:12px;'>
                          <strong style='color:{C['amber']};'>CDM Dependency</strong><br>
                          <span style='color:#374151;'>On track now but cannot advance past current stage
                          until P11 (CDM DAUT ID Replacement) is delivered.</span>
                        </div>""", unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div style='background:#F0FFF8;border-left:3px solid {C['teal']};
                          border-radius:0 6px 6px 0;padding:10px 14px;font-size:12px;'>
                          <strong style='color:{C['teal']};'>✓ No Active Risk Flags</strong>
                        </div>""", unsafe_allow_html=True)

                st.markdown(f"""
                <div class='edit-cta' style='margin-top:12px;'>
                  To update <strong>{pname}</strong> — change status, scores, CDM flag, or add
                  business value — edit directly in Excel.
                  &nbsp;<a href='{EDIT_LINK}' target='_blank'
                    style='color:{C['blue']};font-weight:700;'>Open Excel →</a>
                </div>""", unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────────────────────
st.markdown(f"""
<hr class='slim'>
<div style='display:flex;justify-content:space-between;align-items:center;
  font-size:10px;color:{C['gray']};padding-bottom:8px;'>
  <span>RevOps Program Dashboard · FY26 · Emerson · Read-only live view</span>
  <span>Auto-refreshes every 12 hours ·
    <a href='{EDIT_LINK}' target='_blank'
      style='color:{C['blue']};text-decoration:none;font-weight:600;'>
      Edit source in Excel →</a></span>
</div>""", unsafe_allow_html=True)
