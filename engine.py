"""
Warehouse Control Tower – calculation engine.

Reproduces the WAREHOUSE_Visibility workbook logic (Main Table → Data Model → Calc →
Dashboard / Productivity View / Risk Factor / Pick-Pack Timeline) in pandas.

    raw input sheets ──► daily actuals (History FINAL rule) ──┐
    Main Table (manual plan + manpower rows) ─────────────────┴─► long model
    long model + filters ─► KPIs, tables, risk engine, insights
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import numpy as np
import pandas as pd
from python_calamine import CalamineWorkbook

PROCESS_ORDER = ["GRN", "Putaway", "Picking", "Packing", "Return"]
HIST_COLS = [
    "GRN Apparels", "GRN Non-Apparels", "Putaway Apparels", "Putaway Non-Apparels",
    "Picking", "Packing B2C Apparels", "Packing B2C Non-Apparels",
    "Packing B2B Apparels", "Packing B2B Non-Apparels", "Return",
]
RETURN_TYPES = {"PUTAWAY_COURIER_RETURNED_ITEMS", "PUTAWAY_REVERSE_PICKUP_ITEM"}
GRN_PUTAWAY_TYPE = "PUTAWAY_GRN_ITEM"


@dataclass
class Params:
    good_thr: float = 1.00
    watch_thr: float = 0.90
    cv_thr: float = 0.25
    w_plan: float = 0.5
    w_prod: float = 0.3
    w_vol: float = 0.2
    cap_plan: float = 0.5
    cap_prod: float = 0.5
    cap_vol: float = 1.0
    high_override: float | None = None
    med_override: float | None = None
    pick_basis: str = "Created"
    speed_good: float = 0.90
    speed_watch: float = 0.75


# ───────────────────────────────────────────── helpers
def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def _to_dt(s: pd.Series) -> pd.Series:
    """Handles real datetimes, Excel serial numbers and text dates mixed in one column."""
    s = s.replace("", np.nan)
    num = pd.to_numeric(s, errors="coerce")
    is_num = num.notna() & s.map(lambda v: isinstance(v, (int, float)) and not isinstance(v, bool))
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    if is_num.any():
        out[is_num] = pd.to_datetime(num[is_num], unit="D", origin="1899-12-30")
    rest = ~is_num & s.notna()
    if rest.any():
        out[rest] = pd.to_datetime(s[rest].astype(str) if s[rest].map(lambda v: isinstance(v, str)).all()
                                   else s[rest], errors="coerce", dayfirst=True)
    return out


def _cell_dt(v):
    if v == "" or v is None:
        return pd.NaT
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v)) if 20000 < v < 80000 else pd.NaT
    return pd.to_datetime(v, errors="coerce")


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.replace("", np.nan), errors="coerce").fillna(0.0)


def _sheet_df(wb, name: str) -> pd.DataFrame:
    rows = wb.get_sheet_by_name(name).to_python()
    if not rows:
        return pd.DataFrame()
    seen, cols = {}, []
    for h in rows[0]:
        k = str(h).strip() or "col"
        seen[k] = seen.get(k, 0) + 1
        cols.append(k if seen[k] == 1 else f"{k}.{seen[k]-1}")
    return pd.DataFrame(rows[1:], columns=cols)


def _find_col(df: pd.DataFrame, *names):
    lut = {_norm(c): c for c in df.columns}
    for n in names:
        if _norm(n) in lut:
            return lut[_norm(n)]
    return None


def safe_div(a, b):
    try:
        if b is None or pd.isna(b) or b == 0 or pd.isna(a):
            return np.nan
        return a / b
    except Exception:
        return np.nan


def week_of_day(day: int) -> int:
    """Wk n = days 7n−6 … 7n; days 29–31 join Week 4."""
    return min((day - 1) // 7 + 1, 4)


# ───────────────────────────────────────────── loading
@dataclass
class RawData:
    main_table: list
    history: pd.DataFrame
    inward_map: dict
    outward_map: dict
    channel_map: dict
    recipients: list
    grn: pd.DataFrame
    putaway: pd.DataFrame
    picking: pd.DataFrame
    packed: pd.DataFrame
    warnings: list = field(default_factory=list)


def load_workbook(src) -> RawData:
    """src = file path or raw bytes of the .xlsx."""
    wb = (CalamineWorkbook.from_filelike(io.BytesIO(src)) if isinstance(src, (bytes, bytearray))
          else CalamineWorkbook.from_path(str(src)))
    names = set(wb.sheet_names)
    for n in ["Main Table", "GRN-Input", "PutawayReturn-Input", "Picking-Input", "Packed-Input"]:
        if n not in names:
            raise ValueError(f"Sheet '{n}' not found in workbook.")
    warn = []

    inward_map, outward_map = {}, {}
    if "Category" in names:
        for r in wb.get_sheet_by_name("Category").to_python()[2:]:
            r = list(r) + [""] * 8
            if r[0] != "" and _norm(r[0]) not in inward_map:
                inward_map[_norm(r[0])] = str(r[1]).strip()
            for a, b in ((3, 4), (6, 7)):
                if r[a] != "" and _norm(r[a]) not in outward_map:
                    outward_map[_norm(r[a])] = str(r[b]).strip()

    channel_map, recipients = {}, []
    if "Settings" in names:
        srows = wb.get_sheet_by_name("Settings").to_python()
        hi = next((i for i, r in enumerate(srows) if len(r) > 8 and _norm(r[7]) == "channel name"), None)
        if hi is not None:
            for r in srows[hi + 1:]:
                if len(r) < 9 or r[7] == "" or str(r[7]).startswith("Channels in"):
                    break
                channel_map[_norm(r[7])] = str(r[8]).strip()
        ei = next((i for i, r in enumerate(srows) if len(r) > 11 and _norm(r[10]) == "email"), None)
        if ei is not None:
            for r in srows[ei + 1:]:
                if len(r) < 12 or "@" not in str(r[10]):
                    break
                recipients.append({"Email": str(r[10]).strip(), "Send": str(r[11]).strip() or "Yes",
                                   "Notes": str(r[12]).strip() if len(r) > 12 else ""})
    if not channel_map:
        warn.append("Channel → Registration Type table not found in Settings – all channels counted as Packing B2C.")

    history = pd.DataFrame(columns=["Date"] + HIST_COLS)
    if "History" in names:
        h = wb.get_sheet_by_name("History").to_python()
        hi = next((i for i, r in enumerate(h) if r and _norm(r[0]) == "date"), None)
        if hi is not None:
            body = [list(r[:11]) for r in h[hi + 1:] if r and r[0] != ""]
            history = pd.DataFrame(body, columns=["Date"] + HIST_COLS)
            history["Date"] = _to_dt(history["Date"]).dt.normalize()
            for c in HIST_COLS:
                history[c] = pd.to_numeric(history[c].replace("", np.nan), errors="coerce")
            history = history.dropna(subset=["Date"]).drop_duplicates("Date")

    return RawData(
        main_table=wb.get_sheet_by_name("Main Table").to_python(), history=history,
        inward_map=inward_map, outward_map=outward_map, channel_map=channel_map, recipients=recipients,
        grn=_sheet_df(wb, "GRN-Input"), putaway=_sheet_df(wb, "PutawayReturn-Input"),
        picking=_sheet_df(wb, "Picking-Input"), packed=_sheet_df(wb, "Packed-Input"), warnings=warn,
    )


# ───────────────────────────────────────────── actuals from raw exports
def _cat_type(series: pd.Series, mapping: dict) -> pd.Series:
    t = series.map(lambda x: mapping.get(_norm(x)))
    return t.fillna("Non-Apparels").map(lambda v: "Apparels" if _norm(v).startswith("apparel") else "Non-Apparels")


def build_actuals(raw: RawData):
    diag, frames, coverage = {}, [], {}
    blank = lambda df: pd.Series("", index=df.index)

    g = raw.grn
    cd, cq, cc = _find_col(g, "GRN Date"), _find_col(g, "Quantity Received"), _find_col(g, "Category")
    if cd and cq:
        d = pd.DataFrame({"Date": _to_dt(g[cd]).dt.normalize(), "Qty": _num(g[cq]),
                          "Type": _cat_type(g[cc] if cc else blank(g), raw.inward_map)}).dropna(subset=["Date"])
        coverage["GRN"] = (d["Date"].min(), d["Date"].max())
        d["Col"] = "GRN " + d["Type"]
        frames.append(d[["Date", "Col", "Qty"]])
        diag["grn_unmapped"] = sorted({str(x) for x in (g[cc] if cc else []) if x != "" and _norm(x) not in raw.inward_map})

    p = raw.putaway
    cu, cq, ct, cc = (_find_col(p, "Updated"), _find_col(p, "Putaway Quantity", "Quantity"),
                      _find_col(p, "Type"), _find_col(p, "Category"))
    if cu and cq and ct:
        d = pd.DataFrame({"Date": _to_dt(p[cu]).dt.normalize(), "Qty": _num(p[cq]),
                          "PType": p[ct].astype(str).str.strip(),
                          "Type": _cat_type(p[cc] if cc else blank(p), raw.inward_map)}).dropna(subset=["Date"])
        coverage["Putaway"] = coverage["Return"] = (d["Date"].min(), d["Date"].max())
        pa = d[d["PType"] == GRN_PUTAWAY_TYPE].assign(Col=lambda x: "Putaway " + x["Type"])
        rt = d[d["PType"].isin(RETURN_TYPES)].assign(Col="Return")
        frames += [pa[["Date", "Col", "Qty"]], rt[["Date", "Col", "Qty"]]]

    k = raw.picking
    ccr, cst = _find_col(k, "Created"), _find_col(k, "Picklist Status")
    if ccr and cst:
        d = pd.DataFrame({"Date": _to_dt(k[ccr]).dt.normalize(),
                          "St": k[cst].astype(str).str.strip().str.upper()}).dropna(subset=["Date"])
        coverage["Picking"] = (d["Date"].min(), d["Date"].max())
        d = d[d["St"] == "CLOSED"].assign(Qty=1.0, Col="Picking")
        frames.append(d[["Date", "Col", "Qty"]])

    pk = raw.packed
    ci, cs, cc, ch_ = (_find_col(pk, "Invoice Created"), _find_col(pk, "Sale Order Item Code"),
                       _find_col(pk, "Category"), _find_col(pk, "Channel Name"))
    if ci and cs:
        ch = pk[ch_].astype(str).str.strip() if ch_ else blank(pk)
        reg = ch.map(lambda x: raw.channel_map.get(_norm(x))) if raw.channel_map else pd.Series("Packing B2C", index=pk.index)
        d = pd.DataFrame({"Date": _to_dt(pk[ci]).dt.normalize(), "Has": pk[cs].astype(str).str.strip() != "",
                          "Reg": reg, "Ch": ch,
                          "Type": _cat_type(pk[cc] if cc else blank(pk), raw.outward_map)}).dropna(subset=["Date"])
        coverage["Packing"] = (d["Date"].min(), d["Date"].max())
        diag["unmapped_channels"] = sorted(set(d.loc[d["Reg"].isna(), "Ch"]) - {""})
        d = d[d["Has"] & d["Reg"].notna()].assign(Qty=1.0)
        d["Col"] = d["Reg"] + " " + d["Type"]
        frames.append(d[["Date", "Col", "Qty"]])

    live = (pd.concat(frames).groupby(["Date", "Col"])["Qty"].sum().unstack("Col")
            if frames else pd.DataFrame())
    return live.reindex(columns=HIST_COLS), coverage, diag


def final_actuals(raw: RawData, live, coverage, dates: pd.DatetimeIndex) -> pd.DataFrame:
    """History FINAL rule: live value when the date is inside that source's export range, else stored value."""
    stored = raw.history.set_index("Date").reindex(dates) if not raw.history.empty else pd.DataFrame(index=dates)
    lv = live.reindex(dates)
    out = pd.DataFrame(index=dates)
    for col in HIST_COLS:
        lo, hi = coverage.get(col.split()[0], (None, None))
        inr = ((dates >= lo) & (dates <= hi)) if lo is not None else np.zeros(len(dates), bool)
        lval = lv[col].fillna(0.0).to_numpy()
        sval = stored[col].to_numpy() if col in stored else np.full(len(dates), np.nan)
        out[col] = np.where(inr, lval, sval)
    return out


# ───────────────────────────────────────────── Main Table → long model
def _paren(label: str) -> str:
    m = re.search(r"\(([^)]*)\)", label)
    return m.group(1).strip() if m else ""


def parse_main_table(rows: list):
    # calamine trims leading blank rows, so locate the date header row by content
    hi = next((i for i, r in enumerate(rows[:10])
               if len(r) > 4 and sum(not pd.isna(_cell_dt(v)) for v in r[2:7]) >= 3), None)
    if hi is None:
        raise ValueError("Could not find the date header row in 'Main Table'.")
    hdr = rows[hi]
    date_idx = []
    for j, v in enumerate(hdr[2:], start=2):
        dt = _cell_dt(v)
        if pd.isna(dt):
            break
        date_idx.append((j, dt.normalize()))
    dates = pd.DatetimeIndex([d for _, d in date_idx])
    cols = [j for j, _ in date_idx]

    blocks, cur = {}, None
    for r in rows[hi + 1:]:
        if not r or len(r) < 2:
            continue
        if str(r[0]).strip():
            cur = str(r[0]).strip()
        lab = str(r[1]).strip()
        if cur is None or not lab:
            continue
        vals = np.array([pd.to_numeric(r[j], errors="coerce") if j < len(r) and r[j] != "" else np.nan
                         for j in cols], float)
        blocks.setdefault(cur, []).append((lab, vals))

    slots, mp_rows = [], {}
    for blk, lines in blocks.items():
        proc = blk.split()[0]
        acts = [(l, v) for l, v in lines if _norm(l).startswith("actual number")]
        plans = [(l, v) for l, v in lines if _norm(l).startswith("planned")]
        for l, v in lines:
            if _norm(l).startswith("manpower"):
                mp_rows[proc] = v
        used = set()
        for pl, pv in plans:
            pp = _paren(pl)
            keys = [f"{blk} {_paren(a)}".strip() for a, _ in acts if (_paren(a) == pp or pp == "")]
            used.update(keys)
            slots.append({"Process": proc, "Activity": f"{blk} – {pp}" if pp else blk,
                          "Hist": keys, "Plan": np.nan_to_num(pv)})
        for a, _ in acts:
            key = f"{blk} {_paren(a)}".strip()
            if key not in used:
                pp = _paren(a)
                slots.append({"Process": proc, "Activity": f"{blk} – {pp}" if pp else blk,
                              "Hist": [key], "Plan": np.zeros(len(dates))})
    return dates, slots, mp_rows


@dataclass
class Model:
    long: pd.DataFrame
    mp: pd.DataFrame
    through: pd.Timestamp
    first: pd.Timestamp
    coverage: dict
    diag: dict
    recipients: list
    warnings: list
    plan_mp_wide: pd.DataFrame


def build_model(raw: RawData, overrides: pd.DataFrame | None = None) -> Model:
    dates, slots, mp_rows = parse_main_table(raw.main_table)
    live, coverage, diag = build_actuals(raw)
    act = final_actuals(raw, live, coverage, dates)

    parts = []
    for s in slots:
        keys = [k for k in s["Hist"] if k in act.columns]
        a = act[keys].sum(axis=1, min_count=1) if keys else pd.Series(np.nan, index=dates)
        parts.append(pd.DataFrame({"Date": dates, "Process": s["Process"], "Activity": s["Activity"],
                                   "Actual": a.to_numpy(), "Plan": s["Plan"]}))
    long = pd.concat(parts, ignore_index=True)
    mp = pd.concat([pd.DataFrame({"Date": dates, "Process": p, "MP": v}) for p, v in mp_rows.items()],
                   ignore_index=True)

    wide = {f"Plan | {s['Activity']}": s["Plan"] for s in slots}
    wide.update({f"Manpower | {p}": np.nan_to_num(v) for p, v in mp_rows.items()})
    plan_mp_wide = pd.DataFrame(wide, index=[d.strftime("%Y-%m-%d") for d in dates]).T

    if overrides is not None and not overrides.empty:
        long, mp = apply_overrides(long, mp, overrides)

    nonzero = long.assign(a=long["Actual"].fillna(0)).groupby("Date")["a"].apply(lambda s: (s != 0).any())
    through = nonzero[nonzero].index.max() if nonzero.any() else dates.max()
    long["Actual"] = long["Actual"].where(long["Date"] <= through).fillna(0.0)
    mp["MP"] = mp["MP"].fillna(0.0)
    return Model(long=long, mp=mp, through=through, first=dates.min(), coverage=coverage, diag=diag,
                 recipients=raw.recipients, warnings=raw.warnings, plan_mp_wide=plan_mp_wide)


def apply_overrides(long, mp, wide: pd.DataFrame):
    long, mp = long.copy(), mp.copy()
    lk = long.set_index(["Activity", "Date"])
    mk = mp.set_index(["Process", "Date"])
    for lab, row in wide.iterrows():
        if "|" not in str(lab):
            continue
        kind, name = [x.strip() for x in str(lab).split("|", 1)]
        vals = pd.to_numeric(row, errors="coerce").fillna(0.0)
        idx = pd.to_datetime(vals.index)
        if kind == "Plan":
            for d, v in zip(idx, vals):
                if (name, d) in lk.index:
                    lk.loc[(name, d), "Plan"] = v
        else:
            for d, v in zip(idx, vals):
                if (name, d) in mk.index:
                    mk.loc[(name, d), "MP"] = v
    return lk.reset_index()[long.columns], mk.reset_index()[mp.columns]


# ───────────────────────────────────────────── period / filters
@dataclass
class Filters:
    month: str = "All"
    week: int = 0
    process: str = "All"
    activity: str = "All"


def months_available(m: Model) -> list[str]:
    return sorted({d.strftime("%Y-%m") for d in m.long["Date"] if d <= m.through})


def period_bounds(m: Model, f: Filters):
    if f.month == "All":
        return m.first, m.through
    ms = pd.Timestamp(f.month + "-01")
    me = ms + pd.offsets.MonthEnd(0)
    if f.week:
        s = ms + pd.Timedelta(days=7 * (f.week - 1))
        e = me if f.week == 4 else ms + pd.Timedelta(days=7 * f.week - 1)
    else:
        s, e = ms, me
    return s, min(e, m.through)


def period_label(m: Model, f: Filters) -> str:
    s, e = period_bounds(m, f)
    base = "All months" if f.month == "All" else (
        (f"Week {f.week} of " if f.week else "") + pd.Timestamp(f.month + "-01").strftime("%b-%Y"))
    t = f"{base} ({s:%d-%b-%Y} to {e:%d-%b-%Y})"
    t += f"  |  Process: {f.process}" + (f"  |  Activity: {f.activity}" if f.activity != "All" else "")
    return t


def enrich(m: Model) -> pd.DataFrame:
    L = m.long.merge(m.mp, on=["Date", "Process"], how="left")
    L["MP"] = L["MP"].fillna(0.0)
    L["WD"] = (L["Date"] <= m.through).astype(int)
    L["Proj"] = L["Plan"] * L["WD"]
    L["ActPl"] = L["Actual"] * (L["Plan"] > 0)
    L["Month"] = L["Date"].dt.to_period("M").dt.to_timestamp()
    L["Week"] = L["Date"].dt.day.map(week_of_day)
    return L


def filter_long(L, m, f, use_period=True, use_proc=True):
    x = L
    if use_period:
        s, e = period_bounds(m, f)
        x = x[(x["Date"] >= s) & (x["Date"] <= e)]
    if use_proc:
        if f.process != "All":
            x = x[x["Process"] == f.process]
        if f.activity != "All":
            x = x[x["Activity"] == f.activity]
    return x


def proc_daily(x: pd.DataFrame) -> pd.DataFrame:
    """Slots → process × day; MP counted once per process-day."""
    g = x.groupby(["Date", "Process"], as_index=False).agg(
        Actual=("Actual", "sum"), ActPl=("ActPl", "sum"), Proj=("Proj", "sum"),
        Plan=("Plan", "sum"), MP=("MP", "first"))
    g["ActOnMP"] = np.where(g["MP"] > 0, g["Actual"], 0.0)
    g["ProjOnMP"] = np.where(g["MP"] > 0, g["Proj"], 0.0)
    return g


def benchmarks(L, m) -> dict:
    g = proc_daily(L[L["Date"] <= m.through])
    out = {p: safe_div(d["ActOnMP"].sum(), d["MP"].sum()) for p, d in g.groupby("Process")}
    out["__ALL__"] = safe_div(g["ActOnMP"].sum(), g["MP"].sum())
    out["__avg_act__"] = {p: safe_div(d["Actual"].sum(), (d["Actual"] > 0).sum()) for p, d in g.groupby("Process")}
    out["__avg_mp__"] = {p: safe_div(d["MP"].sum(), (d["MP"] > 0).sum()) for p, d in g.groupby("Process")}
    return out


def status_of(v, p: Params):
    if v is None or pd.isna(v):
        return "–"
    return "Good" if v >= p.good_thr else ("Watch" if v >= p.watch_thr else "Risk")


def _wd(g):
    return int((g.groupby("Date")["MP"].sum() > 0).sum())


# ───────────────────────────────────────────── KPIs & tables
def kpis(L, m, f, p, bench) -> dict:
    x = filter_long(L, m, f)
    g = proc_daily(x)
    wd = _wd(g)
    active = int((g.groupby("Date")["Actual"].sum() > 0).sum())
    pp = g.groupby("Process")["Proj"].sum()
    gp = g[g["Process"].isin(pp[pp > 0].index)]
    k = dict(actual=g["Actual"].sum(), proj=g["Proj"].sum(), actpl=g["ActPl"].sum(), mp=g["MP"].sum(),
             wd=wd, active_days=active, act_on_mp=g["ActOnMP"].sum(), proj_on_mp=g["ProjOnMP"].sum())
    k["ach"] = safe_div(k["actpl"], k["proj"])
    k["var"] = k["actpl"] - k["proj"]
    k["avg_mp"] = safe_div(k["mp"], wd)
    k["blended"] = safe_div(k["act_on_mp"], k["mp"])
    k["bench"] = bench["__ALL__"] if f.process == "All" else bench.get(f.process, np.nan)
    k["prod_idx"] = safe_div(k["blended"], k["bench"])
    k["proj_prod"] = safe_div(k["proj_on_mp"], k["mp"])
    k["act_prod_plan"] = safe_div(gp["ActOnMP"].sum(), gp["MP"].sum())
    k["prod_var"] = k["act_prod_plan"] - k["proj_prod"]
    k["prod_ach"] = safe_div(k["act_prod_plan"], k["proj_prod"])
    k["avg_daily"] = safe_div(k["actual"], active)
    k["avg_weekly"] = k["avg_daily"] * 7
    k["n_proc"] = int((g.groupby("Process")["Actual"].sum() > 0).sum())
    k["n_act"] = int((x.groupby("Activity")["Actual"].sum() > 0).sum())
    k["n_weeks"] = int(x.loc[x["Actual"] > 0, ["Month", "Week"]].drop_duplicates().shape[0])
    k["req_prod"] = safe_div(k["proj"], k["mp"])
    k["efficiency"] = safe_div(k["actual"], k["proj"])
    k["status"] = status_of(k["ach"] if k["proj"] else k["prod_idx"], p)
    return k


def _mp_signal(r):
    a, ml, ix = r["Activity Level Idx"], r["MP Level Idx"], r["Prod. Index"]
    if pd.isna(a) or pd.isna(ml):
        return "–"
    if a >= 1.1 and ml <= 0.9:
        return "Possible MP shortage (volume ↑, MP ↓)"
    if a <= 0.9 and ml >= 1.1:
        return "Under-utilised (volume ↓, MP ↑)"
    if ml <= 0.9 and not pd.isna(ix) and ix < 0.9:
        return "Low MP + low productivity – review"
    if ml >= 1.1 and not pd.isna(ix) and ix < 0.9:
        return "High MP, low productivity – review deployment"
    return "Balanced"


def process_table(L, m, f, p, bench) -> pd.DataFrame:
    g = proc_daily(filter_long(L, m, f))
    wd = _wd(g)
    rows = []
    for proc, d in g.groupby("Process"):
        a, pr, ap, mp = d["Actual"].sum(), d["Proj"].sum(), d["ActPl"].sum(), d["MP"].sum()
        ach = safe_div(ap, pr)
        prod = safe_div(d["ActOnMP"].sum(), mp)
        b = bench.get(proc, np.nan)
        ix = safe_div(prod, b)
        amp = safe_div(mp, wd)
        req = safe_div(safe_div(a, wd), b)
        mpp = safe_div(safe_div(pr, wd), b)
        rows.append({
            "Process": proc, "Actual": a, "Projected": pr, "Actual (Planned)": ap, "Variance": ap - pr,
            "Achievement %": ach, "Manpower": mp, "Avg Daily MP": amp, "Productivity": prod,
            "Projected Prod.": safe_div(d["ProjOnMP"].sum(), mp), "Benchmark": b, "Prod. Index": ix,
            "Status Basis": "Achievement %" if pr > 0 else "Productivity index",
            "Status": status_of(ach if pr > 0 else ix, p), "Score": ach if pr > 0 else ix,
            "Activity Level Idx": safe_div(safe_div(a, (d["Actual"] > 0).sum()), bench["__avg_act__"].get(proc)),
            "MP Level Idx": safe_div(safe_div(mp, (d["MP"] > 0).sum()), bench["__avg_mp__"].get(proc)),
            "Req. MP/day @ bench": req, "MP Utilisation %": safe_div(req, amp),
            "MP gap/day (+surplus)": amp - req, "MP/day for plan": mpp, "Plan MP gap/day": amp - mpp,
        })
    t = pd.DataFrame(rows)
    if t.empty:
        return t
    t["Manpower Signal"] = t.apply(_mp_signal, axis=1)
    t["Share of MP"] = t["Manpower"] / (t["Manpower"].sum() or np.nan)
    t["Share of Actual"] = t["Actual"] / (t["Actual"].sum() or np.nan)
    t["Rank"] = t["Score"].rank(ascending=False, method="min").astype("Int64")
    t["_o"] = t["Process"].map({q: i for i, q in enumerate(PROCESS_ORDER)})
    return t.sort_values(["Rank", "_o"]).drop(columns="_o").reset_index(drop=True)


def month_process_table(L, m, f) -> pd.DataFrame:
    x = filter_long(L, m, f, use_period=False)
    return x[x["Date"] <= m.through].groupby(["Process", "Month"], as_index=False)[["ActPl", "Proj"]].sum()


def weekly_table(L, m, f):
    month = f.month if f.month != "All" else months_available(m)[-1]
    g = proc_daily(filter_long(L, m, Filters(month, 0, f.process, f.activity)))
    g["Week"] = g["Date"].dt.day.map(week_of_day)
    rows = []
    for w, d in g.groupby("Week"):
        daily = d.groupby("Date")["Actual"].sum()
        rows.append({"Week": f"Wk {w} ({d['Date'].min():%d-%b}–{d['Date'].max():%d-%b})", "WeekNo": w,
                     "Actual": d["Actual"].sum(), "Projected": d["Proj"].sum(), "ActPl": d["ActPl"].sum(),
                     "Manpower": d["MP"].sum(), "Productivity": safe_div(d["ActOnMP"].sum(), d["MP"].sum()),
                     "CV": safe_div(daily.std(ddof=0), daily.mean())})
    t = pd.DataFrame(rows)
    if not t.empty:
        t["Variance"] = t["ActPl"] - t["Projected"]
        t["Achievement %"] = t["ActPl"] / t["Projected"].replace(0, np.nan)
    return t, month


def monthly_trend(L, m, f) -> pd.DataFrame:
    x = filter_long(L, m, f, use_period=False)
    g = proc_daily(x[x["Date"] <= m.through])
    g["Month"] = g["Date"].dt.to_period("M").dt.to_timestamp()
    rows = [{"Month": mo, "Actual": d["Actual"].sum(), "Projected": d["Proj"].sum(), "ActPl": d["ActPl"].sum(),
             "Manpower": d["MP"].sum(), "WD": _wd(d), "Productivity": safe_div(d["ActOnMP"].sum(), d["MP"].sum())}
            for mo, d in g.groupby("Month")]
    t = pd.DataFrame(rows)
    if t.empty:
        return t
    t["Variance"] = t["ActPl"] - t["Projected"]
    t["Achievement %"] = t["ActPl"] / t["Projected"].replace(0, np.nan)
    t["Req. Productivity"] = t["Projected"] / t["Manpower"].replace(0, np.nan)
    t["Efficiency %"] = t["Actual"] / t["Projected"].replace(0, np.nan)
    apd = t["Actual"] / t["WD"].replace(0, np.nan)
    t["MoM %"] = apd / apd.shift(1) - 1
    t["Efficiency MoM (pts)"] = t["Efficiency %"].diff()
    return t


def daily_trend(L, m, f) -> pd.DataFrame:
    g = proc_daily(filter_long(L, m, f))
    d = g.groupby("Date", as_index=False).agg(Actual=("Actual", "sum"), MP=("MP", "sum"), Proj=("Proj", "sum"),
                                              ActOnMP=("ActOnMP", "sum"), ActPl=("ActPl", "sum"))
    d["Productivity"] = d["ActOnMP"] / d["MP"].replace(0, np.nan)
    d["Efficiency %"] = d["Actual"] / d["Proj"].replace(0, np.nan)
    d["Req. Productivity"] = d["Proj"] / d["MP"].replace(0, np.nan)
    return d


def activity_table(L, m, f) -> pd.DataFrame:
    g = filter_long(L, m, f).groupby(["Process", "Activity"], as_index=False).agg(
        Actual=("Actual", "sum"), Projected=("Proj", "sum"), ActPl=("ActPl", "sum"))
    g["Variance"] = g["ActPl"] - g["Projected"]
    g["Achievement %"] = g["ActPl"] / g["Projected"].replace(0, np.nan)
    return g


def scorecard(L, m, f):
    by = proc_daily(filter_long(L, m, Filters(f.month, f.week))).groupby("Process")[["Actual", "MP", "ActOnMP"]].sum()
    get = lambda q, c: float(by.loc[q, c]) if q in by.index else 0.0
    flow = pd.DataFrame([
        ("Units received (Receive)", get("GRN", "Actual")),
        ("Units put away (Store)", get("Putaway", "Actual")),
        ("Putaway backlog (received − put away)", get("GRN", "Actual") - get("Putaway", "Actual")),
        ("Units picked (Pick)", get("Picking", "Actual")),
        ("Units packed (Pack)", get("Packing", "Actual")),
        ("Pick − Pack gap", get("Picking", "Actual") - get("Packing", "Actual")),
        ("Returns processed (Return)", get("Return", "Actual")),
    ], columns=["Flow Metric", "Value"])
    roles = {"Receiving/Inbound": ["GRN", "Putaway"], "Outbound": ["Picking", "Packing"], "Returns": ["Return"]}
    rr = [(r, sum(get(q, "Actual") for q in ps), sum(get(q, "MP") for q in ps),
           safe_div(sum(get(q, "ActOnMP") for q in ps), sum(get(q, "MP") for q in ps))) for r, ps in roles.items()]
    return flow, pd.DataFrame(rr, columns=["Role", "Actual", "Manpower", "Productivity"])


# ───────────────────────────────────────────── data quality
def data_quality(L, m, f):
    ga = proc_daily(L[L["Date"] <= m.through])
    ga["Prod"] = ga["Actual"] / ga["MP"].replace(0, np.nan)
    ga["z"] = ga.groupby("Process")["Prod"].transform(
        lambda s: (s - s.mean()) / s.std(ddof=0) if s.std(ddof=0) else s * 0)
    s, e = period_bounds(m, f)
    g = ga[(ga["Date"] >= s) & (ga["Date"] <= e)]
    if f.process != "All":
        g = g[g["Process"] == f.process]
    fl = []
    for r in g.itertuples():
        if r.Actual > 0 and r.MP == 0:
            fl.append((r.Date, r.Process, "All (process)", "Missing manpower", r.Actual, r.MP, np.nan))
        if r.Actual == 0 and r.MP > 0:
            fl.append((r.Date, r.Process, "All (process)", "Zero activity with manpower", r.Actual, r.MP, r.z))
        if not pd.isna(r.z) and abs(r.z) > 3:
            fl.append((r.Date, r.Process, "All (process)", "Outlier productivity", r.Actual, r.MP, r.z))
        if r.Actual < 0 or r.MP < 0:
            fl.append((r.Date, r.Process, "All (process)", "Negative value", r.Actual, r.MP, np.nan))
    xs = filter_long(L, m, f)
    for r in xs[(xs["Date"] <= m.through) & (xs["Actual"] > 0) & (xs["Plan"] == 0)].itertuples():
        fl.append((r.Date, r.Process, r.Activity, "Missing plan", r.Actual, r.MP, np.nan))
    fl = pd.DataFrame(fl, columns=["Date", "Process", "Activity", "DQ Flag", "Actual", "Manpower", "z-score"])
    cnt = lambda n: int((fl["DQ Flag"] == n).sum())
    checks = {
        "Blank date": int(m.long["Date"].isna().sum()),
        "Duplicate date": int(m.long.groupby(["Activity", "Date"]).size().gt(1).sum()),
        "Missing manpower": cnt("Missing manpower"), "Missing plan": cnt("Missing plan"),
        "Zero activity with manpower": cnt("Zero activity with manpower"),
        "Negative value": cnt("Negative value"), "Outlier productivity": cnt("Outlier productivity"),
        "Unmapped GRN/Putaway categories": len(m.diag.get("grn_unmapped", [])),
        "Unmapped packing channels": len(m.diag.get("unmapped_channels", [])),
    }
    chk = pd.DataFrame([(k, v, "✔ OK" if v == 0 else "⚠ Review") for k, v in checks.items()],
                       columns=["Check", "Count", "Status"])
    return chk, fl.sort_values(["Date", "Process"]).reset_index(drop=True)


# ───────────────────────────────────────────── risk engine
def risk_engine(L, m, p: Params, bench):
    g = proc_daily(L[L["Date"] <= m.through]).sort_values(["Process", "Date"]).reset_index(drop=True)
    g["Productivity"] = g["Actual"] / g["MP"].replace(0, np.nan)
    g["Achievement %"] = g["ActPl"] / g["Proj"].replace(0, np.nan)
    g["Benchmark"] = g["Process"].map(bench)
    g["Prod Index"] = g["Productivity"] / g["Benchmark"]
    g["Plan-gap risk"] = ((1 - g["Achievement %"]) / p.cap_plan).clip(0, 1)
    g["Prod-gap risk"] = ((1 - g["Prod Index"]) / p.cap_prod).clip(0, 1)

    def cv7(s):
        v = s.to_numpy(float)
        out = np.full(len(v), np.nan)
        for i in range(len(v)):
            w = v[max(0, i - 6):i + 1]
            w = w[w > 0]                      # zero-activity days are ignored (matches workbook)
            if len(w) >= 3:
                out[i] = w.std() / w.mean()
        return pd.Series(out, index=s.index)

    g["7-day CV"] = g.groupby("Process", group_keys=False)["Actual"].apply(cv7)
    g["Volatility risk"] = (g["7-day CV"] / p.cap_vol).clip(0, 1)
    W = np.array([p.w_plan, p.w_prod, p.w_vol])
    R = g[["Plan-gap risk", "Prod-gap risk", "Volatility risk"]].to_numpy(float)
    ok = ~np.isnan(R)
    den = (ok * W).sum(axis=1)
    g["Risk score"] = np.where(den > 0, 100 * np.nansum(R * W, axis=1) / np.where(den > 0, den, 1), np.nan)
    sc = g["Risk score"].dropna()
    hi = p.high_override if p.high_override is not None else (float(np.percentile(sc, 80)) if len(sc) else 80.0)
    md = p.med_override if p.med_override is not None else (float(np.percentile(sc, 50)) if len(sc) else 50.0)
    g["Risk level"] = [level_of(v, hi, md) for v in g["Risk score"]]
    g["Month"] = g["Date"].dt.to_period("M").dt.to_timestamp()
    g["Week"] = g["Date"].dt.day.map(week_of_day)
    return g, hi, md


def level_of(score, hi, md):
    if pd.isna(score):
        return "–"
    return "High" if score >= hi else ("Medium" if score >= md else "Low")


def cutoff_table(L, m, f, risk, bench):
    month = f.month if f.month != "All" else months_available(m)[-1]
    ms = pd.Timestamp(month + "-01")
    x = L[(L["Date"] >= ms) & (L["Date"] <= ms + pd.offsets.MonthEnd(0))]
    g = proc_daily(x)
    wd = _wd(g[g["Date"] <= m.through])
    rows = []
    for proc, d in g.groupby("Process"):
        plan_m = d["Plan"].sum()
        pdays = int((d["Plan"] > 0).sum())
        rem = int(((d["Plan"] > 0) & (d["Date"] > m.through)).sum())
        dp = safe_div(plan_m, pdays)
        mtd = d[d["Date"] <= m.through]
        ap, pr = mtd["ActPl"].sum(), mtd["Proj"].sum()
        amp = safe_div(mtd["MP"].sum(), wd)
        remaining = plan_m - ap
        cut = safe_div(remaining, rem)
        b = bench.get(proc, np.nan)
        r7 = risk[(risk["Process"] == proc) & (risk["Date"] > m.through - pd.Timedelta(days=7))]["Risk score"].mean()
        rows.append({"Process": proc, "Benchmark prod": b, "Month plan": plan_m, "Plan days": pdays,
                     "Plan days left": rem, "Daily plan": dp, "Actual (Planned) MTD": ap,
                     "Achievement MTD": safe_div(ap, pr), "Remaining plan": remaining, "DAILY CUTOFF VOLUME": cut,
                     "Catch-up uplift": safe_div(cut, dp) - 1, "Avg MP/day (MTD)": amp,
                     "Actual prod (MTD)": safe_div(mtd["ActOnMP"].sum(), mtd["MP"].sum()),
                     "Break-even prod (plan)": safe_div(dp, amp), "Break-even prod (cutoff)": safe_div(cut, amp),
                     "Break-even MP (plan)": safe_div(dp, b), "Break-even MP (cutoff)": safe_div(cut, b),
                     "Last-7-day risk": r7})
    t = pd.DataFrame(rows)
    if not t.empty:
        t = t.sort_values("Process", key=lambda s: s.map({q: i for i, q in enumerate(PROCESS_ORDER)})).reset_index(drop=True)
    return t, month


def weekly_risk(risk, month, hi, md):
    r = risk[risk["Month"] == pd.Timestamp(month + "-01")]
    if r.empty:
        return pd.DataFrame()
    rows = []
    for proc, d in list(r.groupby("Process")) + [("All processes", r)]:
        wk = d.groupby("Week")["Risk score"].mean()
        row = {"Process": proc, **{f"Wk {w}": wk.get(w, np.nan) for w in range(1, 5)}}
        row["Month avg"] = d["Risk score"].mean()
        row["Level"] = level_of(row["Month avg"], hi, md)
        row["Plan-gap risk"] = d["Plan-gap risk"].mean()
        row["Prod-gap risk"] = d["Prod-gap risk"].mean()
        row["Volatility risk"] = d["Volatility risk"].mean()
        ws = wk.dropna()
        if len(ws) >= 2:
            df_ = ws.iloc[-1] - ws.iloc[-2]
            row["Trend vs prior week"] = (f"▲ Worsening (+{df_:.0f})" if df_ > 0.5 else
                                          f"▼ Improving ({df_:.0f})" if df_ < -0.5 else "► Flat")
        else:
            row["Trend vs prior week"] = "–"
        rows.append(row)
    order = {q: i for i, q in enumerate(PROCESS_ORDER + ["All processes"])}
    return pd.DataFrame(rows).sort_values("Process", key=lambda s: s.map(order)).reset_index(drop=True)


# ───────────────────────────────────────────── pick → pack
def pickpack(raw: RawData, basis: str = "Created") -> tuple[pd.DataFrame, int]:
    k, pk = raw.picking, raw.packed
    cs, cst = _find_col(k, "Shipping Package Code"), _find_col(k, "Picklist Status")
    ct = _find_col(k, basis) or _find_col(k, "Created")
    cs2, ci = _find_col(pk, "Shipping Package Code"), _find_col(pk, "Invoice Created")
    if not all([cs, cst, ct, cs2, ci]):
        return pd.DataFrame(), 0
    a = pd.DataFrame({"SPC": k[cs].astype(str).str.strip(), "T": _to_dt(k[ct]),
                      "St": k[cst].astype(str).str.strip().str.upper()})
    a = a[(a["St"] == "CLOSED") & (a["SPC"] != "") & a["T"].notna()]
    pick = a.groupby("SPC").agg(PickTime=("T", "min"), PickItems=("T", "size"))
    b = pd.DataFrame({"SPC": pk[cs2].astype(str).str.strip(), "T": _to_dt(pk[ci])})
    b = b[(b["SPC"] != "") & b["T"].notna()]
    pack = b.groupby("SPC").agg(PackTime=("T", "min"), PackItems=("T", "size"))
    d = pick.join(pack, how="left").reset_index()
    no_pick = int((~pack.index.isin(pick.index)).sum())
    d["Packed"] = d["PackTime"].notna()
    d["Lead hrs"] = (d["PackTime"] - d["PickTime"]).dt.total_seconds() / 3600
    d["Bucket"] = pd.cut(d["Lead hrs"], [-np.inf, 4, 24, 48, np.inf], labels=["0-4 h", "4-24 h", "24-48 h", "Over 48 h"])
    d["PickDate"] = d["PickTime"].dt.normalize()
    d["SameDay"] = d["Packed"] & (d["PackTime"].dt.normalize() == d["PickDate"])
    d["Month"] = d["PickDate"].dt.to_period("M").dt.to_timestamp()
    d["Week"] = d["PickDate"].dt.day.map(week_of_day)
    return d, no_pick


def speed_status(v, p: Params):
    if pd.isna(v):
        return "–"
    return "Good" if v >= p.speed_good else ("Watch" if v >= p.speed_watch else "Risk")


def pickpack_summary(d: pd.DataFrame, p: Params, key) -> pd.DataFrame:
    if d.empty:
        return pd.DataFrame()
    g = d.groupby(key).agg(**{"Pkgs picked": ("SPC", "size"), "Items picked": ("PickItems", "sum"),
                              "Packed": ("Packed", "sum"), "Avg lead (hrs)": ("Lead hrs", "mean"),
                              "Median lead (hrs)": ("Lead hrs", "median"), "Same-day": ("SameDay", "sum")})
    b = d.pivot_table(index=key, columns="Bucket", values="SPC", aggfunc="count", observed=False).fillna(0)
    g = g.join(b)
    g["Not packed"] = g["Pkgs picked"] - g["Packed"]
    g["Match %"] = g["Packed"] / g["Pkgs picked"]
    g["% ≤ 4 h"] = g["0-4 h"] / g["Packed"].replace(0, np.nan)
    g["Same-day %"] = g["Same-day"] / g["Packed"].replace(0, np.nan)
    g["Speed status"] = g["% ≤ 4 h"].map(lambda v: speed_status(v, p))
    return g.reset_index()


# ───────────────────────────────────────────── insights / summary text
def fmt_n(v, d=0):
    return "–" if v is None or pd.isna(v) else f"{v:,.{d}f}"


def fmt_p(v, d=0):
    return "–" if v is None or pd.isna(v) else f"{v*100:.{d}f}%"


def insights(k, pt, wk, mt, wk_month, dq_count) -> dict:
    good, bad, risk, summ = [], [], [], []
    mlab = pd.Timestamp(wk_month + "-01").strftime("%b-%y")
    planned = pt[pt["Projected"] > 0] if not pt.empty else pt
    if not pt.empty:
        if not planned.empty:
            best = planned.sort_values("Achievement %", ascending=False).iloc[0]
            worst = planned.sort_values("Achievement %").iloc[0]
            good.append(f"{best['Process']} is the best performer: {fmt_p(best['Achievement %'])} of plan.")
            bad.append(f"{worst['Process']} is lowest: {fmt_p(worst['Achievement %'])} of plan, "
                       f"short by {fmt_n(-worst['Variance'])} units.")
        hp = pt.sort_values("Productivity", ascending=False).iloc[0]
        good.append(f"Highest productivity: {hp['Process']} at {fmt_n(hp['Productivity'])} units/person-day.")
        for _, r in pt[(pt["Prod. Index"] >= 1) & (pt["Process"] != hp["Process"])].iterrows():
            good.append(f"{r['Process']} productivity is {fmt_p(r['Prod. Index'] - 1)} above its own benchmark.")
        for _, r in pt[pt["Status"].isin(["Risk", "Watch"])].sort_values("Variance").iterrows():
            need = (f"; plan needs ~{fmt_n(-r['Plan MP gap/day'], 1)} more MP/day at benchmark speed"
                    if not pd.isna(r["Plan MP gap/day"]) and r["Plan MP gap/day"] < 0 else "")
            risk.append(f"{r['Process']} ({r['Status']}): gap {fmt_n(-r['Variance'])} units{need}.")
        for _, r in pt[pt["Manpower Signal"].str.contains("review|shortage|Under", case=False, na=False)].iterrows():
            risk.append(f"{r['Process']}: {r['Manpower Signal']}.")
    if not pd.isna(k["ach"]):
        (good if k["ach"] >= 1 else bad).append(
            f"Overall achievement {fmt_p(k['ach'])} – {'surplus' if k['var'] >= 0 else 'gap'} of {fmt_n(abs(k['var']))} units.")
    if not wk.empty and wk["Achievement %"].notna().any():
        b, w = wk.loc[wk["Achievement %"].idxmax()], wk.loc[wk["Achievement %"].idxmin()]
        good.append(f"Best week of {mlab}: {b['Week']} at {fmt_p(b['Achievement %'])} of plan.")
        if len(wk) > 1:
            bad.append(f"Worst week of {mlab}: {w['Week']} at {fmt_p(w['Achievement %'])} of plan.")
    if not pd.isna(k["blended"]) and not pd.isna(k["proj_prod"]) and k["blended"] < k["proj_prod"]:
        bad.append(f"Blended productivity {fmt_n(k['blended'])} vs projected {fmt_n(k['proj_prod'])} units/person-day.")
    if len(mt) > 1 and not pd.isna(mt["MoM %"].iloc[-1]):
        v = mt["MoM %"].iloc[-1]
        (good if v >= 0 else bad).append(f"Actual per working day {'up' if v >= 0 else 'down'} {fmt_p(abs(v))} MoM.")
    if dq_count:
        risk.append(f"{dq_count} data-quality exception(s) to review (see Data Quality tab).")

    n_plan = len(planned)
    on_plan = int((planned["Achievement %"] >= 1).sum()) if n_plan else 0
    summ.append(f"Overall: {fmt_n(k['actual'])} units processed vs {fmt_n(k['proj'])} planned "
                f"({fmt_p(k['ach'])} achievement, variance {fmt_n(k['var'])}).")
    summ.append(f"Plan achievement: {on_plan} of {n_plan} processes on/above plan.")
    summ.append(f"Blended productivity {fmt_n(k['blended'])} units/person-day "
                f"(index {fmt_n(k['prod_idx'], 2)} vs all-history benchmark {fmt_n(k['bench'])}).")
    if n_plan:
        pl = planned.sort_values("Achievement %")
        summ.append(f"Best process: {pl.iloc[-1]['Process']} ({fmt_p(pl.iloc[-1]['Achievement %'])}). "
                    f"Worst: {pl.iloc[0]['Process']} ({fmt_p(pl.iloc[0]['Achievement %'])}).")
    summ.append(f"Manpower: {fmt_n(k['mp'])} person-days (avg {fmt_n(k['avg_mp'], 1)}/day over {k['wd']} working days).")
    if risk:
        summ.append("Top risk: " + risk[0])
    if n_plan:
        summ.append(f"Action: prioritise {planned.sort_values('Achievement %').iloc[0]['Process']} "
                    f"– close the plan gap through manpower and productivity.")
    return dict(good=good, bad=bad, risk=risk, summary=summ)
