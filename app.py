"""
Control Tower – Warehouse Operations Dashboard (Streamlit)

Run:   streamlit run app.py
Data:  upload the WAREHOUSE_Visibility workbook in the sidebar, or place it next to
       this file as WAREHOUSE_Visibility_rename.xlsx.
"""
from __future__ import annotations

import io
import os
import urllib.parse
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

import engine as E

st.set_page_config(page_title="Control Tower – Warehouse Operations", page_icon="🏭", layout="wide")

# ───────────────────────────────────────────── design tokens
INK, PAPER, LINE, MUTED = "#1B2A3A", "#F4F6F8", "#D5DCE3", "#5B6B7B"
DOCK = "#F2B705"
GOOD, WATCH, RISK = "#2E7D4F", "#C98A00", "#C0392B"
SERIES = ["#1B2A3A", "#4F7CAC", "#9DB4C0", "#F2B705", "#7A6C5D", "#A23E48"]
STATUS_COLOR = {"Good": GOOD, "Watch": WATCH, "Risk": RISK, "High": RISK, "Medium": WATCH, "Low": GOOD}

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600&family=Barlow+Semi+Condensed:wght@500;600;700&display=swap');
html, body, [class*="css"], .stMarkdown, .stDataFrame, button, input, select, textarea {{
  font-family: 'Barlow', 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
}}
h1, h2, h3, h4 {{ font-family: 'Barlow Semi Condensed', 'Barlow', sans-serif !important; color: {INK}; letter-spacing: .2px; }}
.block-container {{ padding-top: 1.2rem; max-width: 1500px; }}
.ct-head {{ border-radius: 6px; overflow: hidden; margin-bottom: 1rem; background: {INK}; color: #fff; }}
.ct-stripe {{ height: 10px; background: repeating-linear-gradient(135deg, {DOCK} 0 14px, {INK} 14px 28px); }}
.ct-head-body {{ padding: 14px 20px 16px; display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 8px; }}
.ct-title {{ font-family: 'Barlow Semi Condensed', sans-serif; font-weight: 700; font-size: 1.9rem; line-height: 1.1; margin: 0; }}
.ct-sub {{ color: #C9D3DD; font-size: .95rem; margin-top: 4px; }}
.ct-through {{ font-family: 'Barlow Semi Condensed', sans-serif; font-size: 1.05rem; color: {DOCK}; text-align: right; }}
.kpi {{ background: #fff; border: 1px solid {LINE}; border-left: 5px solid var(--c); border-radius: 4px;
        padding: 10px 12px 9px; height: 100%; }}
.kpi-l {{ color: {MUTED}; font-size: .82rem; line-height: 1.2; min-height: 2em; }}
.kpi-v {{ font-family: 'Barlow Semi Condensed', sans-serif; font-weight: 700; font-size: 1.75rem; color: {INK}; line-height: 1.15; }}
.kpi-n {{ color: {MUTED}; font-size: .75rem; line-height: 1.25; }}
.pill {{ display:inline-block; padding: 1px 9px; border-radius: 10px; color:#fff; font-size:.78rem; font-weight:600; }}
.note {{ color: {MUTED}; font-size: .82rem; }}
.ins {{ background:#fff; border:1px solid {LINE}; border-radius:4px; padding:10px 14px; height:100%; }}
.ins h4 {{ margin: 0 0 6px 0; font-size: 1.05rem; }}
.ins ul {{ margin: 0; padding-left: 1.1rem; }} .ins li {{ margin-bottom: 4px; font-size: .92rem; }}
[data-testid="stSidebar"] {{ background: #E9EDF1; }}
.ct-tw {{ overflow: auto; border: 1px solid {LINE}; border-radius: 4px; background: #fff; margin-bottom: .6rem;
          -webkit-overflow-scrolling: touch; }}
.ct-table {{ border-collapse: collapse; width: 100%; font-size: .86rem; }}
.ct-table th {{ position: sticky; top: 0; background: #EEF2F5; color: {INK}; text-align: left; font-weight: 600;
               padding: 6px 9px; border-bottom: 1px solid {LINE}; white-space: nowrap; z-index: 1; }}
.ct-table td {{ padding: 5px 9px; border-bottom: 1px solid #EDF0F3; white-space: nowrap; }}
.ct-table tr:hover td {{ background: #F7F9FB; }}
.ct-tw, .js-plotly-plot {{ -webkit-user-drag: none; user-select: text; }}
.stTabs [data-baseweb="tab"] {{ font-family: 'Barlow Semi Condensed', sans-serif; font-size: 1.02rem; font-weight: 600; }}
</style>
""", unsafe_allow_html=True)


# ───────────────────────────────────────────── version-safe widgets
def plot(fig, key=None):
    fig.update_layout(template="plotly_white", font=dict(family="Barlow, sans-serif", color=INK, size=13),
                      margin=dict(l=10, r=10, t=46, b=10), colorway=SERIES,
                      legend=dict(orientation="h", y=-0.18), title_font=dict(family="Barlow Semi Condensed", size=17))
    if IS_ADMIN:
        cfg = {"displaylogo": False}
    else:  # viewers: hover/tap for details, no zoom / pan / drag, page scroll not captured
        fig.update_layout(dragmode=False)
        fig.update_xaxes(fixedrange=True)
        fig.update_yaxes(fixedrange=True)
        cfg = {"displayModeBar": False, "scrollZoom": False, "doubleClick": False, "showTips": False}
    try:
        st.plotly_chart(fig, width="stretch", key=key, config=cfg)
    except TypeError:
        st.plotly_chart(fig, use_container_width=True, key=key, config=cfg)


def table(df: pd.DataFrame, fmt: dict | None = None, status_cols=(), height=None, key=None):
    if df is None or df.empty:
        st.info("No rows for the selected filters.")
        return
    df = df.reset_index(drop=True)
    sty = df.style
    if fmt:
        sty = sty.format({k: v for k, v in fmt.items() if k in df.columns}, na_rep="–")
    for c in status_cols:
        if c in df.columns:
            sty = sty.map(lambda v: f"background-color:{STATUS_COLOR[v]};color:white;font-weight:600"
                          if v in STATUS_COLOR else "", subset=[c])
    show_styler(sty, height, key)


def show_styler(sty, height=None, key=None):
    """Admin: interactive grid. Viewers: fixed table – no column dragging, resizing or editing."""
    if IS_ADMIN:
        kw = dict(hide_index=True, key=key)
        if height:
            kw["height"] = height
        try:
            st.dataframe(sty, width="stretch", **kw)
        except TypeError:
            st.dataframe(sty, use_container_width=True, **kw)
        return
    html = sty.hide(axis="index").set_table_attributes('class="ct-table"').to_html()
    mh = f"max-height:{height}px;" if height else "max-height:560px;"
    st.markdown(f'<div class="ct-tw" style="{mh}">{html}</div>', unsafe_allow_html=True)


N0, N1, N2, P0, P1 = "{:,.0f}", "{:,.1f}", "{:,.2f}", "{:.0%}", "{:.1%}"


def kpi_tile(col, label, value, note, better="up", status=None):
    color = STATUS_COLOR.get(status, "#9DB4C0") if status else "#9DB4C0"
    arrow = {"up": "▲ higher is better", "down": "▼ lower for same output", "none": ""}[better]
    col.markdown(f"""<div class="kpi" style="--c:{color}"><div class="kpi-l">{label}</div>
      <div class="kpi-v">{value}</div><div class="kpi-n">{note}<br>{arrow}</div></div>""", unsafe_allow_html=True)


def pill(s):
    return f'<span class="pill" style="background:{STATUS_COLOR.get(s, MUTED)}">{s}</span>'


# ───────────────────────────────────────────── data loading (cached)
DEFAULT_FILE = "WAREHOUSE_Visibility_rename.xlsx"


@st.cache_data(show_spinner="Reading workbook – first load takes ~20–40 s…", max_entries=1)
def load_raw(content: bytes | None, path: str | None, _mtime: float = 0):
    return E.load_workbook(content if content is not None else path)


@st.cache_data(show_spinner=False, max_entries=3)
def build(_raw, raw_key, overrides_csv: str | None):
    ov = pd.read_csv(io.StringIO(overrides_csv), index_col=0) if overrides_csv else None
    m = E.build_model(_raw, ov)
    return m, E.enrich(m)


@st.cache_data(show_spinner="Matching pick → pack packages…", max_entries=4)
def build_pickpack(_raw, raw_key, basis):
    return E.pickpack(_raw, basis)


import tempfile
import time as _time
import urllib.request
from datetime import datetime

SHARED_PATH = Path(tempfile.gettempdir()) / "ct_shared_workbook.xlsx"
SHARED_META = Path(tempfile.gettempdir()) / "ct_shared_workbook.txt"


def _secret(name: str, default: str = "") -> str:
    """Read from Streamlit secrets (cloud) or environment variables (local)."""
    try:
        v = st.secrets.get(name, None)
        if v is not None:
            return str(v)
    except Exception:
        pass
    return os.environ.get(name, default)


ADMIN_PW = _secret("ADMIN_PASSWORD")
DATA_URL = _secret("DATA_URL")


def download_candidates(u: str) -> list[str]:
    """Possible direct-download URLs for Google Drive / OneDrive / SharePoint / Dropbox share links."""
    import base64
    import re as _re
    from urllib.parse import urlsplit
    u = u.strip()
    add = lambda url, q: url + ("&" if "?" in url else "?") + q
    m = _re.search(r"/file/d/([\w-]+)", u) or _re.search(r"[?&]id=([\w-]+)", u)
    if "drive.google.com" in u and m:
        return [f"https://drive.usercontent.google.com/download?id={m.group(1)}&export=download&confirm=t"]
    if "dropbox.com" in u:
        return [_re.sub(r"([?&])dl=0", r"\1dl=1", u) if "dl=0" in u else add(u, "dl=1")]
    if "sharepoint.com" in u or "1drv.ms" in u or "onedrive.live.com" in u:
        c = []
        # 1) share link + download=1
        c.append(u if "download=1" in u else add(u, "download=1"))
        # 2) SharePoint download.aspx?share=<token>
        sp = urlsplit(u)
        mm = _re.search(r"/:[a-z]:/[a-z]/personal/([^/]+)/([^/?]+)", sp.path)
        if mm:
            c.append(f"{sp.scheme}://{sp.netloc}/personal/{mm.group(1)}/_layouts/15/download.aspx?share={mm.group(2)}")
        # 3) OneDrive shares API (works for consumer OneDrive 'anyone' links)
        tok = base64.urlsafe_b64encode(u.encode()).decode().rstrip("=")
        c.append(f"https://api.onedrive.com/v1.0/shares/u!{tok}/root/content")
        return c
    return [u]


def direct_download_url(u: str) -> str:
    return download_candidates(u)[0]


@st.cache_data(ttl=3600, show_spinner="Downloading the latest workbook…", max_entries=1)
def fetch_url(url: str) -> bytes:
    import http.cookiejar
    notes = []
    for cand in download_candidates(url):
        jar = http.cookiejar.CookieJar()            # SharePoint sets a cookie, then redirects to the file
        opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
        opener.addheaders = [("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"), ("Accept", "*/*")]
        try:
            with opener.open(cand, timeout=180) as r:
                data = r.read()
                final = r.geturl()
        except Exception as ex:
            notes.append(f"{type(ex).__name__}: {ex}")
            continue
        if data.startswith(b"PK"):
            return data
        txt = data[:20000].decode("utf-8", "ignore").lower()
        if "login.microsoftonline" in final.lower() or "login.microsoftonline" in txt or "sign in" in txt:
            notes.append("asked for a Microsoft sign-in")
        else:
            notes.append("returned a web page")
    if any("sign-in" in n for n in notes):
        raise ValueError("The OneDrive/SharePoint link asks for a Microsoft sign-in, so it isn't public. "
                         "In OneDrive: Share → gear icon → 'Anyone with the link' → Can view → copy the new link. "
                         "If 'Anyone' isn't available, your IT has blocked public links – use 'Publish for everyone' instead.")
    raise ValueError("The link didn't return an Excel file (" + "; ".join(notes) + "). "
                     "Make sure it is shared as 'Anyone with the link can view'.")


def shared_info() -> str | None:
    if SHARED_PATH.exists():
        return SHARED_META.read_text() if SHARED_META.exists() else "uploaded workbook"
    return None


# ── admin login (one password unlocks every admin feature) ──
with st.sidebar:
    if ADMIN_PW:
        if st.session_state.get("is_admin"):
            c_a, c_b = st.columns([2, 1])
            c_a.markdown("**🔓 Admin mode**")
            if c_b.button("Log out"):
                st.session_state["is_admin"] = False
                st.rerun()
        else:
            with st.expander("🔒 Admin login"):
                pw_try = st.text_input("Admin password", type="password", key="pw_try")
                if pw_try:
                    if pw_try == ADMIN_PW:
                        st.session_state["is_admin"] = True
                        st.rerun()
                    st.error("Wrong password.")
    IS_ADMIN = bool(st.session_state.get("is_admin")) or not ADMIN_PW

# ── admin panel: publish data for everyone ──
with st.sidebar:
    st.markdown("### Data")
    info = shared_info()
    if info:
        st.caption(f"Showing: {info}")
    elif DATA_URL:
        st.caption("Showing: workbook from the shared link (refreshed hourly)")
    if not ADMIN_PW:
        st.caption("Tip: set ADMIN_PASSWORD in the app secrets – until then everyone has admin rights.")
    if IS_ADMIN:
      with st.expander("Update data (admin)", expanded=not (info or DATA_URL)):
        if True:
            up = st.file_uploader("Warehouse workbook (.xlsx)", type=["xlsx", "xlsm"])
            if up is not None and st.button("Publish for everyone", type="primary"):
                SHARED_PATH.write_bytes(up.getvalue())
                SHARED_META.write_text(f"{up.name} · published {datetime.now():%d-%b-%Y %H:%M}")
                st.success("Published – everyone opening the link now sees this data.")
                _time.sleep(1)
                st.rerun()
            if DATA_URL and st.button("Reload from shared link"):
                fetch_url.clear()
                SHARED_PATH.unlink(missing_ok=True)
                st.rerun()
            if info and st.button("Remove published workbook"):
                SHARED_PATH.unlink(missing_ok=True)
                SHARED_META.unlink(missing_ok=True)
                st.rerun()

# ── pick the data source: published upload → shared link → file next to app ──
local = Path(__file__).with_name(DEFAULT_FILE)
raw = None
try:
    if SHARED_PATH.exists():
        mt = SHARED_PATH.stat().st_mtime
        raw_key = f"shared:{mt}"
        raw = load_raw(None, str(SHARED_PATH), mt)
    elif DATA_URL:
        content = fetch_url(DATA_URL)
        raw_key = f"url:{len(content)}:{hash(content[:4096])}"
        raw = load_raw(content, None)
    elif local.exists():
        mt = local.stat().st_mtime
        raw_key = f"local:{mt}"
        raw = load_raw(None, str(local), mt)
except Exception as ex:
    st.error(f"Could not load the workbook: {ex}")
    st.stop()

if raw is None:
    st.markdown('<div class="ct-head"><div class="ct-stripe"></div><div class="ct-head-body">'
                '<div><div class="ct-title">Control Tower – Warehouse Operations</div>'
                '<div class="ct-sub">No data has been published yet.</div></div></div></div>',
                unsafe_allow_html=True)
    st.info("The dashboard owner needs to publish the warehouse workbook: open **Update data (admin)** "
            "in the sidebar, upload the file and click **Publish for everyone**. Once published, "
            "everyone opening this link sees the dashboard.")
    st.stop()

ov_csv = st.session_state.get(f"ov::{raw_key}")
try:
    model, L = build(raw, raw_key, ov_csv)
except Exception as ex:  # surface workbook structure problems clearly
    st.error(f"Could not build the model from this workbook: {ex}")
    st.stop()

# ───────────────────────────────────────────── sidebar filters & settings
with st.sidebar:
    st.markdown("### Filters")
    months = E.months_available(model)
    mlabels = {"All": "All months", **{mm: pd.Timestamp(mm + "-01").strftime("%b-%Y") for mm in months}}
    month = st.selectbox("Month", ["All"] + months, index=len(months), format_func=mlabels.get)
    week = st.selectbox("Week", [0, 1, 2, 3, 4], format_func=lambda w: "All weeks" if w == 0 else
                        f"Week {w} (days {7*w-6}–{'31' if w == 4 else 7*w})", disabled=(month == "All"))
    procs = [q for q in E.PROCESS_ORDER if q in set(model.long["Process"])]
    process = st.selectbox("Process", ["All"] + procs)
    acts = sorted(model.long.loc[(model.long["Process"] == process) | (process == "All"), "Activity"].unique())
    activity = st.selectbox("Activity", ["All"] + acts)
    f = E.Filters(month=month, week=0 if month == "All" else week, process=process, activity=activity)

    p = E.Params()
    with (st.expander("Thresholds & risk settings") if IS_ADMIN else st.container()):
      if IS_ADMIN:
        p.good_thr = st.number_input("Good if achievement ≥", 0.0, 2.0, 1.0, 0.05, format="%.2f")
        p.watch_thr = st.number_input("Watch if achievement ≥ (below = Risk)", 0.0, 2.0, 0.9, 0.05, format="%.2f")
        p.cv_thr = st.number_input("Weekly CV above = volatile", 0.0, 2.0, 0.25, 0.05, format="%.2f")
        st.caption("Risk score weights (re-weighted when a factor is missing)")
        c1, c2, c3 = st.columns(3)
        p.w_plan = c1.number_input("Plan", 0.0, 1.0, 0.5, 0.05)
        p.w_prod = c2.number_input("Prod.", 0.0, 1.0, 0.3, 0.05)
        p.w_vol = c3.number_input("Volat.", 0.0, 1.0, 0.2, 0.05)
        st.caption("Full risk reached at")
        c1, c2, c3 = st.columns(3)
        p.cap_plan = c1.number_input("Plan gap", 0.05, 1.0, 0.5, 0.05)
        p.cap_prod = c2.number_input("Prod gap", 0.05, 1.0, 0.5, 0.05)
        p.cap_vol = c3.number_input("CV", 0.05, 3.0, 1.0, 0.05)
        hi_o = st.text_input("High-risk cutoff override (blank = 80th percentile)", "")
        md_o = st.text_input("Medium-risk cutoff override (blank = median)", "")
        p.high_override = float(hi_o) if hi_o.strip() else None
        p.med_override = float(md_o) if md_o.strip() else None
        p.pick_basis = st.radio("Pick time basis", ["Created", "Updated"], horizontal=True,
                                help="Picklist Created or Updated time is used as the pick timestamp.")
        c1, c2 = st.columns(2)
        p.speed_good = c1.number_input("Speed Good ≥ (≤4 h share)", 0.0, 1.0, 0.9, 0.05)
        p.speed_watch = c2.number_input("Speed Watch ≥", 0.0, 1.0, 0.75, 0.05)
    if abs(p.w_plan + p.w_prod + p.w_vol - 1) > 1e-6:
        st.warning(f"Risk weights add up to {p.w_plan + p.w_prod + p.w_vol:.2f}; scores are normalised by the total.")
    if ov_csv:
        st.success("Edited plan / manpower is active for this session.")

# ───────────────────────────────────────────── calculations
bench = E.benchmarks(L, model)
K = E.kpis(L, model, f, p, bench)
PT = E.process_table(L, model, f, p, bench)
WK, wk_month = E.weekly_table(L, model, f)
MT = E.monthly_trend(L, model, f)
RK, HI, MD = E.risk_engine(L, model, p, bench)
DQ_CHK, DQ_FL = E.data_quality(L, model, f)
dq_total = int(DQ_CHK["Count"].sum())
INS = E.insights(K, PT, WK, MT, wk_month, dq_total)
showing = E.period_label(model, f)

# ───────────────────────────────────────────── header
st.markdown(f"""<div class="ct-head"><div class="ct-stripe"></div><div class="ct-head-body">
<div><div class="ct-title">Control Tower – Warehouse Operations</div><div class="ct-sub">Showing: {showing}</div></div>
<div class="ct-through">Data through {model.through:%d-%b-%Y}<br><span style="color:#C9D3DD;font-size:.85rem">
{K['wd']} working days · status {pill(K['status'])}</span></div></div></div>""", unsafe_allow_html=True)
for w in model.warnings:
    st.warning(w)

tabs = st.tabs(["Overview", "Productivity", "Risk & cutoffs", "Pick → Pack", "Data quality",
                "Plan & manpower", "Export & share", "How it's calculated"])

# ═════════════════════════════════════════════ OVERVIEW
with tabs[0]:
    r1 = st.columns(8)
    kpi_tile(r1[0], "Total actual", E.fmt_n(K["actual"]), "Σ units, all activities")
    kpi_tile(r1[1], "Total projected", E.fmt_n(K["proj"]), "Σ plan, days ≤ data-through")
    kpi_tile(r1[2], "Achievement %", E.fmt_p(K["ach"], 1), "Actual (planned) ÷ projected", status=K["status"])
    kpi_tile(r1[3], "Activity variance", E.fmt_n(K["var"]), "Actual (planned) − projected",
             status="Good" if K["var"] >= 0 else "Risk")
    kpi_tile(r1[4], "Manpower (person-days)", E.fmt_n(K["mp"]), "Σ manpower, process level", "down")
    kpi_tile(r1[5], "Avg daily manpower", E.fmt_n(K["avg_mp"], 1), f"Person-days ÷ {K['wd']} days with MP", "down")
    kpi_tile(r1[6], "Blended productivity", E.fmt_n(K["blended"], 1), "Σ actual ÷ Σ MP (MP days)")
    kpi_tile(r1[7], "Productivity index", E.fmt_n(K["prod_idx"], 2), f"Blended ÷ all-history {E.fmt_n(K['bench'])}",
             status=E.status_of(K["prod_idx"], p))
    st.write("")
    r2 = st.columns(8)
    kpi_tile(r2[0], "Projected productivity", E.fmt_n(K["proj_prod"], 1), "Σ plan ÷ Σ MP (MP days)")
    kpi_tile(r2[1], "Actual productivity (plan processes)", E.fmt_n(K["act_prod_plan"], 1), "Σ actual ÷ Σ MP")
    kpi_tile(r2[2], "Productivity variance", E.fmt_n(K["prod_var"], 1), "Actual − projected productivity",
             status="Good" if (K["prod_var"] or 0) >= 0 else "Risk")
    kpi_tile(r2[3], "Productivity achievement", E.fmt_p(K["prod_ach"], 1), "Actual ÷ projected productivity",
             status=E.status_of(K["prod_ach"], p))
    kpi_tile(r2[4], "Avg daily activity", E.fmt_n(K["avg_daily"]), f"Σ actual ÷ {K['active_days']} active days")
    kpi_tile(r2[5], "Avg weekly activity", E.fmt_n(K["avg_weekly"]), "Avg daily × 7")
    kpi_tile(r2[6], "Processes / activities", f"{K['n_proc']} / {K['n_act']}", "With activity in period", "none")
    kpi_tile(r2[7], "Days / weeks", f"{K['active_days']} / {K['n_weeks']}", "Active days / weeks", "none")

    st.write("")
    c1, c2 = st.columns([1.35, 1])
    with c1:
        st.markdown("#### Planned vs projected by process")
        if not PT.empty:
            d = PT.set_index("Process").reindex([q for q in E.PROCESS_ORDER if q in PT["Process"].values]).reset_index()
            fig = go.Figure()
            fig.add_bar(x=d["Process"], y=d["Projected"], name="Projected", marker_color="#C8D3DD")
            fig.add_bar(x=d["Process"], y=d["Actual (Planned)"], name="Actual (planned)",
                        marker_color=[STATUS_COLOR.get(s, INK) for s in d["Status"]],
                        text=[E.fmt_p(v) for v in d["Achievement %"]], textposition="outside")
            fig.update_layout(barmode="overlay", height=360, title=f"Bars coloured by status · {showing.split('|')[0]}")
            fig.data[1].width = 0.5
            plot(fig, "pvp")
    with c2:
        st.markdown("#### Blended productivity")
        bt = pd.DataFrame([
            ("Blended productivity (units/person-day)", K["blended"], "▲"),
            ("Projected productivity", K["proj_prod"], "▲"),
            ("Actual productivity (plan processes)", K["act_prod_plan"], "▲"),
            ("Productivity achievement %", K["prod_ach"], "▲"),
            ("Benchmark productivity (all history)", K["bench"], "▲"),
            ("Productivity index vs benchmark", K["prod_idx"], "▲"),
            ("Avg daily manpower", K["avg_mp"], "▼"),
        ], columns=["Metric", "Value", "Better"])
        bt["Value"] = [E.fmt_p(v, 1) if "%" in mt_ else E.fmt_n(v, 2 if "index" in mt_ else 1)
                       for mt_, v in zip(bt["Metric"], bt["Value"])]
        table(bt)
        st.markdown('<div class="note">Weighted: Σ actual ÷ Σ manpower, never an average of process ratios.</div>',
                    unsafe_allow_html=True)
        if not PT.empty:
            gp = PT[["Process", "Actual (Planned)", "Projected", "Variance", "Achievement %", "Status"]].copy()
            tot = pd.DataFrame([{"Process": "Total", "Actual (Planned)": gp["Actual (Planned)"].sum(),
                                 "Projected": gp["Projected"].sum(), "Variance": gp["Variance"].sum(),
                                 "Achievement %": E.safe_div(gp["Actual (Planned)"].sum(), gp["Projected"].sum()),
                                 "Status": K["status"]}])
            table(pd.concat([gp, tot]), {"Actual (Planned)": N0, "Projected": N0, "Variance": N0, "Achievement %": P1},
                  ["Status"])

    st.markdown("#### Process performance – ranked best to worst")
    cols = ["Rank", "Process", "Status", "Actual", "Projected", "Actual (Planned)", "Variance", "Achievement %",
            "Manpower", "Avg Daily MP", "Productivity", "Projected Prod.", "Benchmark", "Prod. Index", "Status Basis",
            "Share of MP", "Share of Actual"]
    table(PT[cols] if not PT.empty else PT,
          {"Actual": N0, "Projected": N0, "Actual (Planned)": N0, "Variance": N0, "Achievement %": P1, "Manpower": N0,
           "Avg Daily MP": N2, "Productivity": N1, "Projected Prod.": N1, "Benchmark": N1, "Prod. Index": N2,
           "Share of MP": P0, "Share of Actual": P0}, ["Status"])

    st.markdown("#### Month-wise planned vs projected by process")
    mp_tab = E.month_process_table(L, model, f)
    if not mp_tab.empty:
        c1, c2 = st.columns([1.2, 1])
        with c1:
            piv_a = mp_tab.pivot(index="Process", columns="Month", values="ActPl")
            piv_p = mp_tab.pivot(index="Process", columns="Month", values="Proj")
            rows = []
            for q in [x for x in E.PROCESS_ORDER if x in piv_a.index] + ["Total"]:
                a = piv_a.sum() if q == "Total" else piv_a.loc[q]
                pr = piv_p.sum() if q == "Total" else piv_p.loc[q]
                for metric, vals in [("Planned (actual)", a), ("Projected", pr), ("Gap", a - pr),
                                     ("Gap %", (a - pr) / pr.replace(0, np.nan))]:
                    r = {"Process": q, "Metric": metric}
                    r.update({c.strftime("%b-%y"): v for c, v in vals.items()})
                    r["Total"] = (a.sum() - pr.sum()) / pr.sum() if metric == "Gap %" and pr.sum() else (
                        vals.sum() if metric != "Gap %" else np.nan)
                    rows.append(r)
            mm = pd.DataFrame(rows)
            numc = [c for c in mm.columns if c not in ("Process", "Metric")]
            sty = mm.style.format(lambda v: "–" if pd.isna(v) else f"{v:,.0f}", subset=numc)
            sty = sty.format(lambda v: "–" if pd.isna(v) else f"{v:.1%}", subset=pd.IndexSlice[mm["Metric"] == "Gap %", numc])
            sty = sty.map(lambda v: f"color:{RISK}" if isinstance(v, (int, float)) and v < 0 else "",
                          subset=pd.IndexSlice[mm["Metric"].isin(["Gap", "Gap %"]), numc])
            show_styler(sty, 420)
            st.markdown('<div class="note">Follows Process / Activity filters, not the Month / Week filter.</div>',
                        unsafe_allow_html=True)
        with c2:
            mp_tab["Achievement %"] = mp_tab["ActPl"] / mp_tab["Proj"].replace(0, np.nan)
            mp_tab["Month label"] = mp_tab["Month"].dt.strftime("%b-%y")
            fig = px.bar(mp_tab, x="Process", y="Achievement %", color="Month label", barmode="group",
                         title="Achievement % by process and month", color_discrete_sequence=SERIES)
            fig.add_hline(y=p.good_thr, line_dash="dot", line_color=GOOD)
            fig.update_yaxes(tickformat=".0%")
            fig.update_layout(height=420, legend_title_text="")
            plot(fig, "mpa")

    c1, c2 = st.columns([1.5, 1])
    mlab = pd.Timestamp(wk_month + "-01").strftime("%b-%Y")
    with c1:
        st.markdown(f"#### Weekly performance – {mlab}")
        table(WK[["Week", "Actual", "Projected", "Variance", "Achievement %", "Manpower", "Productivity", "CV"]]
              if not WK.empty else WK,
              {"Actual": N0, "Projected": N0, "Variance": N0, "Achievement %": P1, "Manpower": N0,
               "Productivity": N1, "CV": P0})
        st.markdown(f'<div class="note">Week rule: Wk n = days 7n−6 to 7n; days 29–31 join Week 4. '
                    f'CV above {p.cv_thr:.0%} = volatile output.</div>', unsafe_allow_html=True)
    with c2:
        st.markdown("#### Week highlights")
        if not WK.empty:
            hl = []
            pick = lambda col, fn: WK.loc[getattr(WK[col], fn)()]
            hl.append(("Best week (achievement %)", pick("Achievement %", "idxmax")["Week"], E.fmt_p(WK["Achievement %"].max(), 1)))
            hl.append(("Worst week (achievement %)", pick("Achievement %", "idxmin")["Week"], E.fmt_p(WK["Achievement %"].min(), 1)))
            hl.append(("Peak week (actual)", pick("Actual", "idxmax")["Week"], E.fmt_n(WK["Actual"].max())))
            hl.append(("Lowest week (actual)", pick("Actual", "idxmin")["Week"], E.fmt_n(WK["Actual"].min())))
            hl.append(("Highest-productivity week", pick("Productivity", "idxmax")["Week"], E.fmt_n(WK["Productivity"].max(), 1)))
            hl.append(("Largest-gap week (variance)", pick("Variance", "idxmin")["Week"], E.fmt_n(WK["Variance"].min())))
            table(pd.DataFrame(hl, columns=["Highlight", "Week", "Value"]))

    c1, c2 = st.columns([1, 1.25])
    with c1:
        st.markdown("#### Monthly trend")
        table(MT[["Month", "Actual", "Projected", "Variance", "Achievement %", "Manpower", "Productivity", "MoM %"]]
              .assign(Month=MT["Month"].dt.strftime("%b-%y")) if not MT.empty else MT,
              {"Actual": N0, "Projected": N0, "Variance": N0, "Achievement %": P1, "Manpower": N0,
               "Productivity": N1, "MoM %": P1})
        st.markdown('<div class="note">MoM % compares actual per working day, so part months compare fairly.</div>',
                    unsafe_allow_html=True)
    with c2:
        st.markdown("#### Manpower analysis (vs benchmark productivity)")
        if not PT.empty:
            ma = PT[["Process", "Avg Daily MP", "Req. MP/day @ bench", "MP Utilisation %", "MP gap/day (+surplus)",
                     "MP/day for plan", "Plan MP gap/day", "Manpower Signal"]].copy()
            tot = {"Process": "Total", "Avg Daily MP": ma["Avg Daily MP"].sum(),
                   "Req. MP/day @ bench": ma["Req. MP/day @ bench"].sum(),
                   "MP/day for plan": ma["MP/day for plan"].sum()}
            tot["MP Utilisation %"] = E.safe_div(tot["Req. MP/day @ bench"], tot["Avg Daily MP"])
            tot["MP gap/day (+surplus)"] = tot["Avg Daily MP"] - tot["Req. MP/day @ bench"]
            tot["Plan MP gap/day"] = tot["Avg Daily MP"] - tot["MP/day for plan"]
            tot["Manpower Signal"] = ""
            ma = pd.concat([ma, pd.DataFrame([tot])])
            table(ma, {"Avg Daily MP": N2, "Req. MP/day @ bench": N2, "MP Utilisation %": P0,
                       "MP gap/day (+surplus)": N2, "MP/day for plan": N2, "Plan MP gap/day": N2})
            st.markdown('<div class="note">Negative plan MP gap = more manpower needed to hit plan at benchmark speed. '
                        'Utilisation below 100% = more people than the volume needed.</div>', unsafe_allow_html=True)

    c1, c2 = st.columns([1, 1])
    flow, roles = E.scorecard(L, model, f)
    with c1:
        st.markdown("#### Warehouse scorecard – flow")
        table(flow, {"Value": N0})
        st.markdown('<div class="note">Backlog = received − put away in the period (all processes, ignores Process filter).</div>',
                    unsafe_allow_html=True)
    with c2:
        st.markdown("#### By role")
        table(roles, {"Actual": N0, "Manpower": N0, "Productivity": N1})
        fig = go.Figure(go.Funnel(y=flow["Flow Metric"].iloc[[0, 1, 3, 4]].str.replace(r" \(.*\)", "", regex=True),
                                  x=flow["Value"].iloc[[0, 1, 3, 4]], marker_color=[SERIES[0], SERIES[1], SERIES[2], DOCK],
                                  textinfo="value"))
        fig.update_layout(height=250, title="Receive → store → pick → pack")
        plot(fig, "funnel")

    st.markdown("#### Management insights")
    c1, c2, c3 = st.columns(3)
    for col, title, items, mark in [(c1, "Good", INS["good"], "🟢"), (c2, "Performance gaps", INS["bad"], "🔴"),
                                    (c3, "Risk / attention (ranked by impact)", INS["risk"], "🟠")]:
        col.markdown(f'<div class="ins"><h4>{mark} {title}</h4><ul>' +
                     "".join(f"<li>{x}</li>" for x in items or ["Nothing flagged."]) + "</ul></div>",
                     unsafe_allow_html=True)
    st.markdown("#### Executive summary")
    st.markdown("\n".join(f"- {s}" for s in INS["summary"]))

# ═════════════════════════════════════════════ PRODUCTIVITY
with tabs[1]:
    st.markdown("#### Blended productivity & overall efficiency")
    st.markdown('<div class="note">Manpower present vs actual units handled across GRN, Putaway, Picking, Packing and Return.</div>',
                unsafe_allow_html=True)
    r = st.columns(7)
    kpi_tile(r[0], "Units handled", E.fmt_n(K["actual"]), "Σ actual")
    kpi_tile(r[1], "MP present", E.fmt_n(K["mp"]), "Person-days", "down")
    kpi_tile(r[2], "Blended productivity", E.fmt_n(K["blended"], 1), "Units ÷ MP (MP days)")
    kpi_tile(r[3], "Planned units", E.fmt_n(K["proj"]), "Σ plan to data-through")
    kpi_tile(r[4], "Required productivity", E.fmt_n(K["req_prod"], 1), "Plan ÷ MP")
    kpi_tile(r[5], "Overall efficiency", E.fmt_p(K["efficiency"], 1), "Units ÷ planned units",
             status=E.status_of(K["efficiency"], p))
    kpi_tile(r[6], "Avg MP / day", E.fmt_n(K["avg_mp"], 1), f"Over {K['wd']} working days", "down")

    st.write("")
    st.markdown("#### Process view")
    if not PT.empty:
        pv = PT[["Process", "Actual", "Manpower", "Productivity", "Projected"]].copy()
        pv["Required Productivity"] = pv["Projected"] / pv["Manpower"].replace(0, np.nan)
        pv["Overall Efficiency %"] = pv["Actual"] / pv["Projected"].replace(0, np.nan)
        pv["Benchmark"] = PT["Benchmark"]
        pv = pv.set_index("Process").reindex([q for q in E.PROCESS_ORDER if q in pv["Process"].values]).reset_index()
        tot = {"Process": "Total / blended", "Actual": K["actual"], "Manpower": K["mp"], "Productivity": K["blended"],
               "Projected": K["proj"], "Required Productivity": K["req_prod"], "Overall Efficiency %": K["efficiency"],
               "Benchmark": K["bench"]}
        pv = pd.concat([pv, pd.DataFrame([tot])]).rename(columns={
            "Actual": "Units handled", "Manpower": "MP present (person-days)",
            "Productivity": "Blended productivity", "Projected": "Planned units"})
        c1, c2 = st.columns([1.2, 1])
        with c1:
            table(pv, {"Units handled": N0, "MP present (person-days)": N0, "Blended productivity": N1,
                       "Planned units": N0, "Required Productivity": N1, "Overall Efficiency %": P1, "Benchmark": N1})
        with c2:
            d = pv[pv["Process"] != "Total / blended"]
            fig = go.Figure()
            fig.add_bar(x=d["Process"], y=d["Blended productivity"], name="Actual", marker_color=INK)
            fig.add_bar(x=d["Process"], y=d["Required Productivity"], name="Required (plan ÷ MP)", marker_color=DOCK)
            fig.add_scatter(x=d["Process"], y=d["Benchmark"], name="Benchmark", mode="markers",
                            marker=dict(symbol="line-ew-open", size=34, line=dict(width=3, color=RISK)))
            fig.update_layout(barmode="group", height=330, title="Productivity by process: actual vs required")
            plot(fig, "pbp")

    DT = E.daily_trend(L, model, f)
    c1, c2 = st.columns(2)
    with c1:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=DT["Date"], y=DT["Actual"], name="Units handled", marker_color="#9DB4C0")
        fig.add_scatter(x=DT["Date"], y=DT["Proj"], name="Planned", mode="lines", line=dict(color=INK, dash="dot"))
        fig.add_scatter(x=DT["Date"], y=DT["MP"], name="MP present", mode="lines+markers",
                        line=dict(color=DOCK, width=3), secondary_y=True)
        fig.update_yaxes(title_text="Units", secondary_y=False)
        fig.update_yaxes(title_text="Person-days", secondary_y=True, showgrid=False)
        fig.update_layout(height=360, title=f"Daily: units handled vs MP present – {process if process!='All' else 'all processes'}")
        plot(fig, "dum")
    with c2:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_scatter(x=DT["Date"], y=DT["Productivity"], name="Blended productivity", mode="lines+markers",
                        line=dict(color=INK, width=3))
        fig.add_scatter(x=DT["Date"], y=DT["Req. Productivity"], name="Required productivity", mode="lines",
                        line=dict(color=DOCK, dash="dash"))
        fig.add_bar(x=DT["Date"], y=DT["Efficiency %"], name="Overall efficiency %", marker_color="#C8D3DD",
                    secondary_y=True, opacity=0.8)
        fig.update_yaxes(title_text="Units / person-day", secondary_y=False)
        fig.update_yaxes(tickformat=".0%", secondary_y=True, showgrid=False)
        fig.update_layout(height=360, title="Daily: blended productivity & overall efficiency %")
        fig.data = (fig.data[2], fig.data[0], fig.data[1])
        plot(fig, "dpe")

    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.markdown("#### Month-wise trend")
        if not MT.empty:
            mtv = MT[["Month", "Actual", "Manpower", "Productivity", "Projected", "Req. Productivity",
                      "Efficiency %", "Efficiency MoM (pts)"]].copy()
            mtv["Month"] = mtv["Month"].dt.strftime("%b-%y")
            table(mtv.rename(columns={"Actual": "Units handled", "Manpower": "MP present", "Projected": "Planned units"}),
                  {"Units handled": N0, "MP present": N0, "Productivity": N1, "Planned units": N0,
                   "Req. Productivity": N1, "Efficiency %": P1, "Efficiency MoM (pts)": "{:+.1%}"})
    with c2:
        if not MT.empty:
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_bar(x=MT["Month"].dt.strftime("%b-%y"), y=MT["Efficiency %"], name="Overall efficiency %",
                        marker_color="#9DB4C0", text=[E.fmt_p(v) for v in MT["Efficiency %"]], textposition="outside")
            fig.add_scatter(x=MT["Month"].dt.strftime("%b-%y"), y=MT["Productivity"], name="Blended productivity",
                            mode="lines+markers", line=dict(color=INK, width=3), secondary_y=True)
            fig.update_yaxes(tickformat=".0%", secondary_y=False)
            fig.update_yaxes(showgrid=False, secondary_y=True)
            fig.update_layout(height=320, title="Month-wise overall efficiency & blended productivity")
            plot(fig, "mte")

    st.markdown("#### Daily detail")
    dd = DT[["Date", "Actual", "MP", "Productivity", "Proj", "Efficiency %"]].copy()
    dd["Date"] = dd["Date"].dt.strftime("%d-%b-%Y")
    table(dd.rename(columns={"Actual": "Units handled", "MP": "MP present", "Proj": "Planned units"}),
          {"Units handled": N0, "MP present": N0, "Productivity": N1, "Planned units": N0, "Efficiency %": P1}, height=320)

    st.markdown("#### Activity breakdown")
    AT = E.activity_table(L, model, f)
    table(AT.rename(columns={"ActPl": "Actual (Planned)"}),
          {"Actual": N0, "Projected": N0, "Actual (Planned)": N0, "Variance": N0, "Achievement %": P1})

# ═════════════════════════════════════════════ RISK
with tabs[2]:
    CT, cmonth = E.cutoff_table(L, model, f, RK, bench)
    clab = pd.Timestamp(cmonth + "-01").strftime("%b-%Y")
    r = st.columns(4)
    kpi_tile(r[0], "High-risk cutoff (score ≥)", E.fmt_n(HI), "80th percentile of all process-days"
             if p.high_override is None else "Manual override", "none", "High")
    kpi_tile(r[1], "Medium-risk cutoff (score ≥)", E.fmt_n(MD), "Median of all process-days"
             if p.med_override is None else "Manual override", "none", "Medium")
    last7 = RK[RK["Date"] > model.through - pd.Timedelta(days=7)]["Risk score"].mean()
    kpi_tile(r[2], "Warehouse risk, last 7 days", E.fmt_n(last7, 1), "Avg process-day score (0–100)", "down",
             E.level_of(last7, HI, MD))
    kpi_tile(r[3], "Weights plan / prod / volatility", f"{p.w_plan:.0%} / {p.w_prod:.0%} / {p.w_vol:.0%}",
             "Edit in sidebar settings", "none")

    st.markdown(f"#### Cutoff factors by process – {clab}")
    st.markdown('<div class="note">What each process must do per day from tomorrow to finish the month on plan, '
                'and what that means for productivity and headcount.</div>', unsafe_allow_html=True)
    table(CT, {"Benchmark prod": N1, "Month plan": N0, "Daily plan": N0, "Actual (Planned) MTD": N0,
               "Achievement MTD": P1, "Remaining plan": N0, "DAILY CUTOFF VOLUME": N0, "Catch-up uplift": "{:+.1%}",
               "Avg MP/day (MTD)": N2, "Actual prod (MTD)": N1, "Break-even prod (plan)": N1,
               "Break-even prod (cutoff)": N1, "Break-even MP (plan)": N2, "Break-even MP (cutoff)": N2,
               "Last-7-day risk": N1})
    st.markdown('<div class="note">Daily cutoff volume = (month plan − actual (planned) MTD) ÷ plan days remaining. '
                'Break-even productivity = units per person per day the current team needs. '
                'Break-even MP = people per day needed at benchmark productivity.</div>', unsafe_allow_html=True)

    st.markdown(f"#### Weekly risk roll-up – {clab}")
    WR = E.weekly_risk(RK, cmonth, HI, MD)
    table(WR, {**{f"Wk {w}": N1 for w in range(1, 5)}, "Month avg": N1, "Plan-gap risk": N2,
               "Prod-gap risk": N2, "Volatility risk": N2}, ["Level"])

    st.markdown("#### Daily risk heatmap (score 0–100)")
    hm = RK.pivot(index="Process", columns="Date", values="Risk score")
    hm = hm.reindex([q for q in E.PROCESS_ORDER if q in hm.index])
    fig = go.Figure(go.Heatmap(z=hm.values, x=hm.columns, y=hm.index, zmin=0, zmax=100,
                               colorscale=[[0, "#E3EFE7"], [MD / 100, "#F6E3A6"], [HI / 100, "#E7A08E"], [1, RISK]],
                               hovertemplate="%{y} · %{x|%d-%b}<br>Score %{z:.0f}<extra></extra>", xgap=1, ygap=2))
    fig.update_layout(height=300, title=f"Low < {MD:.0f} ≤ Medium < {HI:.0f} ≤ High")
    plot(fig, "heat")

    with st.expander("Risk engine – process × day detail"):
        rq = st.multiselect("Processes", procs, default=procs, key="rq")
        rv = RK[RK["Process"].isin(rq)][["Process", "Date", "Actual", "ActPl", "Proj", "MP", "Productivity",
                                             "Achievement %", "Benchmark", "Prod Index", "Plan-gap risk",
                                             "Prod-gap risk", "7-day CV", "Volatility risk", "Risk score", "Risk level"]]
        rv = rv.assign(Date=rv["Date"].dt.strftime("%d-%b-%Y")).rename(columns={"ActPl": "Actual (Planned)", "Proj": "Projected"})
        table(rv, {"Actual": N0, "Actual (Planned)": N0, "Projected": N0, "MP": N0, "Productivity": N1,
                   "Achievement %": P0, "Benchmark": N1, "Prod Index": N2, "Plan-gap risk": N2, "Prod-gap risk": N2,
                   "7-day CV": N2, "Volatility risk": N2, "Risk score": N1}, ["Risk level"], height=420)
        st.markdown('<div class="note">Plan-gap risk = min(1, (1 − achievement) ÷ cap). Prod-gap risk = min(1, (1 − index) ÷ cap). '
                    'Volatility = CV of the last 7 days with activity (needs ≥ 3 active days). '
                    'Missing factors are left out and the rest re-weighted.</div>', unsafe_allow_html=True)

# ═════════════════════════════════════════════ PICK → PACK
with tabs[3]:
    PP, no_pick = build_pickpack(raw, raw_key, p.pick_basis)
    if PP.empty:
        st.info("Pick → pack matching needs Shipping Package Code, Picklist Status and Created/Updated in Picking-Input, "
                "and Shipping Package Code + Invoice Created in Packed-Input.")
    else:
        s_, e_ = E.period_bounds(model, f)
        sel = PP[(PP["PickDate"] >= s_) & (PP["PickDate"] <= e_)]
        packed = int(sel["Packed"].sum())
        within4 = E.safe_div((sel["Bucket"] == "0-4 h").sum(), packed)
        r = st.columns(8)
        kpi_tile(r[0], "Packages picked", E.fmt_n(len(sel)), "Picklist status CLOSED")
        kpi_tile(r[1], "Packed (matched)", E.fmt_n(packed), "Same Shipping Package Code")
        kpi_tile(r[2], "Not packed yet", E.fmt_n(len(sel) - packed), "Picked, no invoice", "down")
        kpi_tile(r[3], "Match %", E.fmt_p(E.safe_div(packed, len(sel)), 1), "Packed ÷ picked")
        kpi_tile(r[4], "Avg pick → pack", f"{E.fmt_n(sel['Lead hrs'].mean(), 2)} h", "Mean lead time", "down")
        kpi_tile(r[5], "Median pick → pack", f"{E.fmt_n(sel['Lead hrs'].median(), 2)} h", "Half packed faster", "down")
        kpi_tile(r[6], "Packed within 4 h", E.fmt_p(within4, 1), f"Good ≥ {p.speed_good:.0%}, Watch ≥ {p.speed_watch:.0%}",
                 status=E.speed_status(within4, p))
        kpi_tile(r[7], "Packed pkgs with no pick record", E.fmt_n(no_pick), "All dates in Packed-Input", "down")

        st.write("")
        ppm = cmonth if f.month == "All" else f.month
        ms = pd.Timestamp(ppm + "-01")
        wk_pp = E.pickpack_summary(PP[PP["Month"] == ms], p, "Week")
        st.markdown(f"#### Week-wise timeline – {ms:%b-%Y}")
        if not wk_pp.empty:
            wb_ = PP[PP["Month"] == ms].groupby("Week")["PickDate"].agg(["min", "max"])
            wk_pp.insert(1, "Dates", [f"{wb_.loc[w,'min']:%d-%b} – {wb_.loc[w,'max']:%d-%b}" for w in wk_pp["Week"]])
            wk_pp["Week"] = "Wk " + wk_pp["Week"].astype(str)
            tot = E.pickpack_summary(PP[PP["Month"] == ms].assign(All="Total"), p, "All").rename(columns={"All": "Week"})
            tot.insert(1, "Dates", "")
            wk_pp = pd.concat([wk_pp, tot])
        ppcols = ["Pkgs picked", "Packed", "Not packed", "Match %", "Avg lead (hrs)", "Median lead (hrs)", "% ≤ 4 h",
                  "4-24 h", "24-48 h", "Over 48 h", "Same-day %", "Speed status"]
        ppfmt = {"Pkgs picked": N0, "Items picked": N0, "Packed": N0, "Not packed": N0, "Match %": P1,
                 "Avg lead (hrs)": N2, "Median lead (hrs)": N2, "% ≤ 4 h": P1, "0-4 h": N0, "4-24 h": N0,
                 "24-48 h": N0, "Over 48 h": N0, "Same-day %": P1}
        table(wk_pp[["Week", "Dates"] + ppcols] if not wk_pp.empty else wk_pp, ppfmt, ["Speed status"])

        day = E.pickpack_summary(sel, p, "PickDate")
        if not day.empty:
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            for b_, c_ in zip(["0-4 h", "4-24 h", "24-48 h", "Over 48 h"], [GOOD, "#9DB4C0", WATCH, RISK]):
                fig.add_bar(x=day["PickDate"], y=day[b_], name=b_, marker_color=c_)
            fig.add_bar(x=day["PickDate"], y=day["Not packed"], name="Not packed", marker_color="#D5DCE3")
            fig.add_scatter(x=day["PickDate"], y=day["% ≤ 4 h"], name="% within 4 h", mode="lines+markers",
                            line=dict(color=INK, width=3), secondary_y=True)
            fig.add_hline(y=p.speed_good, line_dash="dot", line_color=GOOD, secondary_y=True)
            fig.update_yaxes(title_text="Packages", secondary_y=False)
            fig.update_yaxes(tickformat=".0%", range=[0, 1.05], secondary_y=True, showgrid=False)
            fig.update_layout(barmode="stack", height=380, title="Day-wise packages by pick → pack lead-time bucket")
            plot(fig, "ppd")

            c1, c2 = st.columns([1.5, 1])
            with c1:
                st.markdown("#### Day-wise timeline (by pick date)")
                dv = day.assign(PickDate=day["PickDate"].dt.strftime("%d-%b-%Y")).rename(columns={"PickDate": "Pick date"})
                table(dv[["Pick date", "Pkgs picked", "Items picked", "Packed", "Not packed", "Match %", "Avg lead (hrs)",
                          "Median lead (hrs)", "0-4 h", "4-24 h", "24-48 h", "Over 48 h", "% ≤ 4 h", "Same-day %",
                          "Speed status"]], ppfmt, ["Speed status"], height=380)
            with c2:
                lh = sel["Lead hrs"].dropna()
                fig = px.histogram(lh[lh <= 48], nbins=48, title="Lead-time distribution (≤ 48 h shown)",
                                   color_discrete_sequence=[INK])
                fig.add_vline(x=4, line_dash="dot", line_color=GOOD, annotation_text="4 h")
                fig.update_layout(height=380, showlegend=False, xaxis_title="Hours from pick to invoice",
                                  yaxis_title="Packages")
                plot(fig, "pph")
        st.markdown('<div class="note">Lead time = pack time (Invoice Created) − pick time (picklist Created or Updated, '
                    'per settings). Key = Shipping Package Code. Buckets: 0-4 h, 4-24 h, 24-48 h, over 48 h.</div>',
                    unsafe_allow_html=True)

# ═════════════════════════════════════════════ DATA QUALITY
with tabs[4]:
    c1, c2 = st.columns([1, 1.6])
    with c1:
        st.markdown("#### Checks")
        table(pd.concat([DQ_CHK, pd.DataFrame([("Total exceptions", dq_total, "✔ OK" if dq_total == 0 else "⚠ Review")],
                                               columns=DQ_CHK.columns)]))
        st.markdown('<div class="note">Flags only – nothing is deleted or excluded. Missing manpower = activity with '
                    'MP = 0. Missing plan = actual on a line/day with no plan. Outlier = |z-score| of daily process '
                    'productivity above 3.</div>', unsafe_allow_html=True)
    with c2:
        st.markdown("#### Flagged rows (selected period)")
        table(DQ_FL.assign(Date=DQ_FL["Date"].dt.strftime("%d-%b-%Y")) if not DQ_FL.empty else DQ_FL,
              {"Actual": N0, "Manpower": N0, "z-score": N2}, height=360)
    st.markdown("#### Source coverage")
    cov = pd.DataFrame([(k_, v[0], v[1]) for k_, v in model.coverage.items()], columns=["Source", "First date", "Last date"])
    cov["Rule"] = "Live data used inside this range; stored History values outside it"
    table(cov.assign(**{"First date": cov["First date"].dt.strftime("%d-%b-%Y"),
                        "Last date": cov["Last date"].dt.strftime("%d-%b-%Y")}))
    um1, um2 = model.diag.get("grn_unmapped", []), model.diag.get("unmapped_channels", [])
    if um1:
        st.warning("GRN / Putaway categories not in the Category sheet (counted as Non-Apparels): " + ", ".join(um1))
    if um2:
        st.warning("Packed channels not in Settings channel table (excluded from packing): " + ", ".join(um2))

# ═════════════════════════════════════════════ PLAN & MANPOWER
with tabs[5]:
    if not IS_ADMIN:
        st.markdown("#### Plan & manpower")
        st.markdown('<div class="note">Planned units per activity and people per process per day (read-only). '
                    'Only the admin can edit these – log in from the sidebar.</div>', unsafe_allow_html=True)
        base_v = model.plan_mp_wide if not ov_csv else pd.read_csv(io.StringIO(ov_csv), index_col=0)
        vis_v = [c for c in base_v.columns if c.startswith(wk_month)] or list(base_v.columns)
        table(base_v[vis_v].reset_index().rename(columns={"index": "Line"}), {c: N0 for c in vis_v}, height=520)
    else:
        st.markdown("#### Plan & manpower grid")
        st.markdown('<div class="note">These are the manual rows of Main Table (planned units per activity and people per '
                    'process per day). Edit cells, then apply to recalculate the whole dashboard. Changes last for this '
                    'session – download the CSV to keep them, and upload it next time.</div>', unsafe_allow_html=True)
        base = model.plan_mp_wide if not ov_csv else pd.read_csv(io.StringIO(ov_csv), index_col=0)
        show_from = st.select_slider("Show dates from", options=list(base.columns),
                                     value=next((c for c in base.columns if c.startswith(wk_month)), base.columns[0]))
        cols_vis = list(base.columns[list(base.columns).index(show_from):])
        try:
            edited = st.data_editor(base[cols_vis], width="stretch", key=f"ed::{raw_key}::{show_from}")
        except TypeError:
            edited = st.data_editor(base[cols_vis], use_container_width=True, key=f"ed::{raw_key}::{show_from}")
        c1, c2, c3, c4 = st.columns(4)
        if c1.button("Apply edits", type="primary"):
            full = base.copy()
            full[cols_vis] = edited.values
            st.session_state[f"ov::{raw_key}"] = full.to_csv()
            st.rerun()
        if c2.button("Reset to workbook values", disabled=not ov_csv):
            st.session_state.pop(f"ov::{raw_key}", None)
            st.rerun()
        c3.download_button("Download grid (CSV)", base.to_csv().encode(), "plan_manpower.csv", "text/csv")
        upl = c4.file_uploader("Load a saved grid", type=["csv"], label_visibility="collapsed")
        if upl is not None and st.session_state.get("last_grid") != upl.name + str(upl.size):
            st.session_state["last_grid"] = upl.name + str(upl.size)
            st.session_state[f"ov::{raw_key}"] = upl.getvalue().decode()
            st.rerun()


# ═════════════════════════════════════════════ EXPORT & SHARE
with tabs[6]:
    if IS_ADMIN:
        st.markdown("#### Email the dashboard summary")
        recips = [r_["Email"] for r_ in model.recipients if r_["Send"].lower() == "yes"]
        to = st.text_input("To (comma-separated)", ", ".join(recips))
        subj = st.text_input("Subject", f"Warehouse Dashboard - data through {model.through:%d-%b-%Y}")
        body = (f"Hi, please review the Warehouse Dashboard ({showing}).\n\n" + "\n".join(f"• {s}" for s in INS["summary"]))
        body = st.text_area("Body", body, height=220)
        link = f"mailto:{to.replace(' ', '')}?subject={urllib.parse.quote(subj)}&body={urllib.parse.quote(body)}"
        st.link_button("Open in mail app", link, type="primary")
        if len(link) > 1800:
            st.caption("Long message – some mail apps cut mailto links around 2,000 characters. Copy the body if it's truncated.")

    st.markdown("#### Download tables")

    def export_xlsx() -> bytes:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as xw:
            pd.DataFrame([(k_, v) for k_, v in K.items()], columns=["KPI", "Value"]).to_excel(xw, sheet_name="KPIs", index=False)
            PT.to_excel(xw, sheet_name="Process performance", index=False)
            E.month_process_table(L, model, f).to_excel(xw, sheet_name="Month x process", index=False)
            WK.to_excel(xw, sheet_name="Weekly", index=False)
            MT.to_excel(xw, sheet_name="Monthly", index=False)
            E.daily_trend(L, model, f).to_excel(xw, sheet_name="Daily", index=False)
            CT.to_excel(xw, sheet_name="Cutoffs", index=False)
            E.weekly_risk(RK, cmonth, HI, MD).to_excel(xw, sheet_name="Weekly risk", index=False)
            RK.to_excel(xw, sheet_name="Risk engine", index=False)
            DQ_CHK.to_excel(xw, sheet_name="DQ checks", index=False)
            DQ_FL.to_excel(xw, sheet_name="DQ flags", index=False)
            L.to_excel(xw, sheet_name="Data model (long)", index=False)
            pd.DataFrame({"Summary": INS["summary"]}).to_excel(xw, sheet_name="Summary", index=False)
        return out.getvalue()

    c1, c2 = st.columns(2)
    if c1.button("Prepare Excel export"):
        st.session_state["xlsx"] = export_xlsx()
    if st.session_state.get("xlsx"):
        c1.download_button("Download dashboard tables (.xlsx)", st.session_state["xlsx"],
                           f"warehouse_dashboard_{model.through:%Y%m%d}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("Download data model (CSV)", L.to_csv(index=False).encode(),
                       f"data_model_{model.through:%Y%m%d}.csv", "text/csv")

# ═════════════════════════════════════════════ FORMULA GUIDE
with tabs[7]:
    st.markdown("#### How each number is calculated")
    guide = [
        ("Source → actuals", "GRN", "Σ Quantity Received per GRN Date, split Apparels / Non-Apparels via the Category sheet."),
        ("Source → actuals", "Putaway", "Σ Putaway Quantity per Updated date where Type = PUTAWAY_GRN_ITEM, split by category type."),
        ("Source → actuals", "Picking", "Count of picked items per Created date where Picklist Status = CLOSED."),
        ("Source → actuals", "Packing", "Count of Sale Order Item Codes per Invoice Created date, split B2C / B2B by the Settings channel table and by category type."),
        ("Source → actuals", "Return", "Σ quantity per Updated date where Type = PUTAWAY_COURIER_RETURNED_ITEMS or PUTAWAY_REVERSE_PICKUP_ITEM."),
        ("Source → actuals", "History rule", "Inside each input sheet's date range the live export is used; outside it, the stored History values."),
        ("Source → actuals", "Plan & manpower", "Read from the manual rows of Main Table (editable in the Plan & manpower tab)."),
        ("Data model", "Projected", "Plan × working day (date ≤ data-through), so future plan doesn't dilute achievement."),
        ("Data model", "Actual (planned)", "Actual only on lines/days that have a plan – like-for-like with projected."),
        ("Data model", "Weeks", "Wk n = days 7n−6 to 7n of the month; days 29–31 join Week 4."),
        ("KPIs", "Achievement %", "Actual (planned) ÷ projected. Good ≥ Good threshold, Watch ≥ Watch threshold, else Risk."),
        ("KPIs", "Avg daily manpower", "Manpower person-days ÷ days in period with any manpower."),
        ("Productivity", "Blended / process productivity", "Actual on days with manpower ÷ manpower person-days."),
        ("Productivity", "Benchmark", "Same ratio over all loaded dates (ignores Month/Week filter)."),
        ("Productivity", "Productivity index", "Period productivity ÷ benchmark (1.00 = normal)."),
        ("Productivity", "Projected productivity", "Plan on MP days ÷ manpower – what the plan silently assumes."),
        ("Manpower", "Req. MP/day @ bench", "(Actual ÷ working days) ÷ benchmark."),
        ("Manpower", "MP for plan / gap", "(Plan ÷ working days) ÷ benchmark; gap = avg MP − MP for plan."),
        ("Manpower", "Signal", "Activity index ≥ 1.1 & MP index ≤ 0.9 → possible shortage; ≤ 0.9 & ≥ 1.1 → under-utilised."),
        ("Risk", "Score (0–100)", "100 × (w_plan × plan risk + w_prod × prod risk + w_vol × volatility) ÷ weights used."),
        ("Risk", "Cutoffs", "High = 80th percentile, Medium = median of all process-day scores (unless overridden)."),
        ("Risk", "Daily cutoff volume", "(Month plan − actual (planned) MTD) ÷ plan days remaining."),
        ("Pick → pack", "Lead time", "(Invoice Created − picklist time) in hours; ≥ 90% within 4 h = Good, ≥ 75% = Watch."),
    ]
    table(pd.DataFrame(guide, columns=["Section", "Metric", "Logic"]))
