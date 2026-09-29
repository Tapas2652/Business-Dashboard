from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="CEO | Business Performance Cockpit",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# CONFIGURATION
# =========================================================
# CEO reporting universe: only these BHs are included everywhere.
# TBA variants are consolidated into TBA.
ALLOWED_BHS = [
    "Sadhna Shukla", "Prathap Sagar", "Mehr Hashim", "Anuradha",
    "Ajay", "Deepak Desai", "Jawad Ulla Khan", "TBA", "ULITES"
]
BH_ORDER = {name: i for i, name in enumerate(ALLOWED_BHS)}

# Only these four domains are included in the CEO view.
# MS, International and unmapped/other domains are excluded everywhere.
ALLOWED_DOMAINS = ["Captive", "Services", "ITES", "ULITES"]
DOMAIN_ORDER = {name: i for i, name in enumerate(ALLOWED_DOMAINS)}

REPO_FILE = Path(__file__).with_name("CEO_MongoDB.xlsx")

SHEET_MAP = {
    "Demand": {"date": "Created_at", "bh": "BH", "kam": "KAM", "domain": "Domain", "client": "Company_name"},
    "Submission": {"date": "date", "bh": "BH", "kam": "KAM", "domain": "Domain", "client": "client"},
    "Interview": {"date": "Interview_date", "bh": "BH", "kam": "KAM", "domain": "Domain", "client": "company_name"},
    "Selection": {"date": "selection_date", "bh": "BH", "kam": "KAM", "domain": "Domain", "client": "company_name"},
    "Onboarding": {"date": "display_date", "bh": "BH", "kam": "KAM", "domain": "Domain", "client": "company_name"},
    "Onboarding Pipeline": {"date": "display_date", "bh": "BH", "kam": "KAM", "domain": "Domain", "client": "company_name"},
    "Exit": {"date": "last_work_day", "bh": "BH", "kam": "KAM", "domain": "Domain", "client": "company_name"},
    "Exit Pipeline": {"date": "tentative_exit_date", "bh": "BH", "kam": "KAM", "domain": "Domain", "client": "company_name"},
    "Active Headcount": {"date": "display_date", "bh": "BH", "kam": "KAM", "domain": "Domain", "client": "company_name"},
}

REQUIRED = {
    "Demand": ["no_of_opening", "Created_at", "BH", "KAM", "Domain", "Company_name"],
    "Submission": ["date", "BH", "KAM", "Domain", "client"],
    "Interview": ["Interview_date", "BH", "KAM", "Domain", "company_name"],
    "Selection": ["selection_date", "BH", "KAM", "Domain", "company_name", "po", "margin"],
    "Onboarding": ["display_date", "BH", "KAM", "Domain", "company_name", "p_o_value", "margin"],
    "Onboarding Pipeline": ["display_date", "BH", "KAM", "Domain", "company_name", "p_o_value", "margin"],
    "Exit": ["last_work_day", "BH", "KAM", "Domain", "company_name", "p_o_value", "margin"],
    "Exit Pipeline": ["tentative_exit_date", "BH", "KAM", "Domain", "company_name", "p_o_value", "margin"],
    "Active Headcount": ["display_date", "BH", "KAM", "Domain", "company_name", "p_o_value", "margin"],
}

# =========================================================
# STYLING
# =========================================================
st.markdown(
    """
    <style>
    .stApp { background: #071b3b; }
    .block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 1800px; }
    [data-testid="stHeader"] { background: rgba(0,0,0,0); }
    [data-testid="stSidebar"] { background: #061733; }
    [data-testid="stSidebar"] * { color: #eef5ff; }
    .hero {
        background: linear-gradient(100deg, #08295b 0%, #0b477d 48%, #0b294f 100%);
        border: 1px solid #2d6093; border-radius: 16px; padding: 18px 22px; margin-bottom: 12px;
        box-shadow: 0 8px 24px rgba(0,0,0,.22);
    }
    .hero h1 { color: white; margin: 0; font-size: 2rem; letter-spacing: -.03em; }
    .hero p { color: #c7d9ed; margin: 5px 0 0; font-size: .9rem; }
    .section-title { color: white; font-size: 1.15rem; font-weight: 700; margin: 12px 0 7px; }
    .metric {
        background: #f6f8fb; border: 2px solid #dbe5ef; border-radius: 12px;
        padding: 12px 14px; min-height: 104px; box-shadow: 0 4px 14px rgba(0,0,0,.12);
    }
    .metric .label { color: #405064; font-size: .78rem; font-weight: 700; text-transform: uppercase; }
    .metric .value { color: #092f59; font-size: 1.8rem; font-weight: 800; line-height: 1.1; margin-top: 6px; }
    .metric .sub { color: #66778a; font-size: .72rem; margin-top: 5px; }
    .metric.green { border-top: 5px solid #16823b; }
    .metric.red { border-top: 5px solid #cf3940; }
    .metric.blue { border-top: 5px solid #0b5a91; }
    .metric.gold { border-top: 5px solid #d3911e; }
    .metric.dark { background: #0c2e56; border-color: #29557e; }
    .metric.dark .label, .metric.dark .sub { color: #c8d9eb; }
    .metric.dark .value { color: white; }
    .chart-card { background: #f7f9fc; border-radius: 14px; padding: 4px 7px 2px; }
    .small-note { color: #c2d3e6; font-size: .78rem; }
    div[data-testid="stTabs"] button { color: #d7e6f7; font-weight: 700; }
    div[data-testid="stTabs"] button[aria-selected="true"] { color: #ffffff; }
    .stDataFrame { border-radius: 10px; overflow: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# DATA LOADING
# =========================================================
@st.cache_data(show_spinner=False)
def load_workbook(source_bytes=None, source_path=None):
    raw = BytesIO(source_bytes) if source_bytes is not None else source_path
    data = {}
    for sheet, cfg in SHEET_MAP.items():
        df = pd.read_excel(raw, sheet_name=sheet)
        missing = [c for c in REQUIRED[sheet] if c not in df.columns]
        if missing:
            raise ValueError(f"{sheet}: missing columns: {', '.join(missing)}")
        df = df.copy()
        df[cfg["date"]] = pd.to_datetime(df[cfg["date"]], errors="coerce").dt.normalize()
        for c in ["BH", "KAM", "Domain", cfg["client"]]:
            if c in df.columns:
                df[c] = df[c].fillna("Unmapped").astype(str).str.strip().replace({"": "Unmapped"})
        if "BH" in df.columns:
            df["BH"] = df["BH"].replace({"TBA - I": "TBA", "TBA-1": "TBA", "TBA - 1": "TBA", "TBA–1": "TBA"})
        for c in ["no_of_opening", "po", "margin", "p_o_value"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        data[sheet] = df
    return data


def get_repo_bytes():
    return REPO_FILE.read_bytes() if REPO_FILE.exists() else None


def month_label(value):
    """Return a consistent Month label for either a pandas Period or Timestamp/date."""
    if isinstance(value, pd.Period):
        ts = value.start_time
    else:
        ts = pd.Timestamp(value)
    return ts.strftime("%b'%y")


def period_from_label(label):
    return pd.Period(pd.to_datetime("01-" + label.replace("'", "-"), format="%d-%b-%y"), freq="M")


def date_filter(df, date_col, start, end):
    return df[(df[date_col] >= start) & (df[date_col] <= end)].copy()


def mtd(df, date_col, period, as_of):
    start = period.start_time.normalize()
    end = min(period.end_time.normalize(), pd.Timestamp(as_of).normalize())
    return date_filter(df, date_col, start, end)


def pipeline(df, date_col, period, as_of):
    # Pipeline means records scheduled after the selected DOD and inside the selected month.
    start = pd.Timestamp(as_of).normalize() + pd.Timedelta(days=1)
    end = period.end_time.normalize()
    return date_filter(df, date_col, start, end)


def money_lakh(value):
    return float(value or 0) / 100000


def fmt_int(value):
    return f"{int(round(value or 0)):,}"


def fmt_lakh(value):
    return f"{money_lakh(value):,.1f} L"


def pct(margin, po):
    return 0 if po == 0 else margin / po * 100


def metric_card(label, value, sub="", kind="blue"):
    st.markdown(
        f'''<div class="metric {kind}">
        <div class="label">{label}</div><div class="value">{value}</div><div class="sub">{sub}</div>
        </div>''',
        unsafe_allow_html=True,
    )


def filter_df(df, sheet, filters):
    """Apply the fixed CEO BH/domain universe, then user filters."""
    cfg=SHEET_MAP[sheet]
    x=df.copy()
    if "BH" in x.columns:
        x=x[x["BH"].isin(ALLOWED_BHS)]
    if "Domain" in x.columns:
        x=x[x["Domain"].isin(ALLOWED_DOMAINS)]
    for field,values in [("BH",filters["bh"]),("KAM",filters["kam"]),("Domain",filters["domain"]),(cfg["client"],filters["client"])]:
        if values and field in x.columns:
            x=x[x[field].isin(values)]
    return x


def universe_values(data,column):
    vals=set()
    for sheet in SHEET_MAP:
        if column in data[sheet].columns:
            vals.update(data[sheet][column].dropna().astype(str).unique())
    if column=="BH":
        return [x for x in ALLOWED_BHS if x in vals or x=="ULITES"]
    if column=="Domain":
        return [x for x in ALLOWED_DOMAINS if x in vals]
    return sorted(vals)


def ordered_sort(df,metric_col=None,group_cols=None):
    """Sort Domain -> BH -> KAM, with the requested metric descending within each group."""
    if df.empty:
        return df
    out=df.copy()
    if "Domain" in out.columns:
        out["__domain_order"]=out["Domain"].map(DOMAIN_ORDER).fillna(999)
    if "BH" in out.columns:
        out["__bh_order"]=out["BH"].map(BH_ORDER).fillna(999)
    sort_cols=[c for c in ["__domain_order","__bh_order"] if c in out.columns]
    if group_cols:
        sort_cols += [c for c in group_cols if c in out.columns and c not in sort_cols]
    if metric_col and metric_col in out.columns:
        sort_cols.append(metric_col)
        ascending=[True]*(len(sort_cols)-1)+[False]
    else:
        ascending=[True]*len(sort_cols)
    if sort_cols:
        out=out.sort_values(sort_cols,ascending=ascending,kind="stable")
    return out.drop(columns=[c for c in ["__domain_order","__bh_order"] if c in out.columns])

# =========================================================
# CORE AGGREGATIONS
# =========================================================
def overall_metrics(data, period, as_of, filters):
    def f(sheet):
        return filter_df(data[sheet], sheet, filters)

    demand = mtd(f("Demand"), "Created_at", period, as_of)
    sub = mtd(f("Submission"), "date", period, as_of)
    intr = mtd(f("Interview"), "Interview_date", period, as_of)
    sel = mtd(f("Selection"), "selection_date", period, as_of)
    ob = mtd(f("Onboarding"), "display_date", period, as_of)
    obp = pipeline(f("Onboarding Pipeline"), "display_date", period, as_of)
    ex = mtd(f("Exit"), "last_work_day", period, as_of)
    exp = pipeline(f("Exit Pipeline"), "tentative_exit_date", period, as_of)

    return {
        "demand": demand["no_of_opening"].sum(),
        "submission": len(sub),
        "interview": len(intr),
        "selection": len(sel),
        "ob_hc": len(ob),
        "ob_po": ob["p_o_value"].sum(),
        "ob_margin": ob["margin"].sum(),
        "ob_pipe_hc": len(obp),
        "ob_pipe_po": obp["p_o_value"].sum(),
        "ob_pipe_margin": obp["margin"].sum(),
        "exit_hc": len(ex),
        "exit_po": ex["p_o_value"].sum(),
        "exit_margin": ex["margin"].sum(),
        "exit_pipe_hc": len(exp),
        "exit_pipe_po": exp["p_o_value"].sum(),
        "exit_pipe_margin": exp["margin"].sum(),
    }


def active_snapshot(data, as_of, filters):
    ah = filter_df(data["Active Headcount"], "Active Headcount", filters)
    ah = ah[ah["display_date"] <= pd.Timestamp(as_of).normalize()]
    if ah.empty:
        return ah
    latest = ah["display_date"].max()
    return ah[ah["display_date"] == latest].copy()


def bh_scorecard(data, period, as_of, filters):
    bhs = filters["bh"] or ALLOWED_BHS
    rows = []
    for bh in bhs:
        local = dict(filters)
        local["bh"] = [bh]
        m = overall_metrics(data, period, as_of, local)
        ah = active_snapshot(data, as_of, local)
        ob_proj = m["ob_hc"] + m["ob_pipe_hc"]
        ex_proj = m["exit_hc"] + m["exit_pipe_hc"]
        rows.append({
            "BH": bh,
            "Demand": m["demand"], "Submission": m["submission"], "Interview": m["interview"], "Selection": m["selection"],
            "OB MTD": m["ob_hc"], "OB Pipeline": m["ob_pipe_hc"], "OB Projection": ob_proj,
            "Exit MTD": m["exit_hc"], "Exit Pipeline": m["exit_pipe_hc"], "Exit Projection": ex_proj,
            "MTD Net": m["ob_hc"] - m["exit_hc"], "Net Projection": ob_proj - ex_proj,
            "OB PO (L)": money_lakh(m["ob_po"]), "OB Margin (L)": money_lakh(m["ob_margin"]),
            "Exit PO (L)": money_lakh(m["exit_po"]), "Exit Margin (L)": money_lakh(m["exit_margin"]),
            "Active HC": len(ah), "Active PO (L)": money_lakh(ah["p_o_value"].sum()), "Active Margin (L)": money_lakh(ah["margin"].sum()),
        })
    return pd.DataFrame(rows)


def client_metrics(data, period, as_of, filters, event="Onboarding"):
    # Build client universe from all relevant source tabs after global filters.
    client_col_map = {s: SHEET_MAP[s]["client"] for s in SHEET_MAP}
    clients = set()
    for s, c in client_col_map.items():
        x = filter_df(data[s], s, filters)
        clients.update(x[c].dropna().astype(str).unique())

    rows = []
    for client in sorted(clients):
        local = dict(filters)
        local["client"] = [client]
        m = overall_metrics(data, period, as_of, local)
        ah = active_snapshot(data, as_of, local)
        # Resolve client attributes from the latest active snapshot first, then fall back to
        # current onboarding/exit records so the client table always has BH/KAM/Domain.
        attr_frames=[]
        for src in ["Active Headcount", "Onboarding", "Exit"]:
            xx=filter_df(data[src], src, local)
            cc=SHEET_MAP[src]["client"]
            if not xx.empty:
                attr_frames.append(xx[[c for c in ["BH","KAM","Domain"] if c in xx.columns]])
        attrs=pd.concat(attr_frames, ignore_index=True) if attr_frames else pd.DataFrame()
        def mode_attr(col):
            if attrs.empty or col not in attrs.columns:
                return ""
            vals=attrs[col].dropna().astype(str)
            return vals.mode().iat[0] if not vals.empty else ""

        rows.append({
            "Client": client,
            "BH": mode_attr("BH"),
            "KAM": mode_attr("KAM"),
            "Domain": mode_attr("Domain"),
            "Demand": m["demand"], "Submission": m["submission"], "Interview": m["interview"], "Selection": m["selection"],
            "OB MTD": m["ob_hc"], "OB Pipeline": m["ob_pipe_hc"], "OB Projection": m["ob_hc"] + m["ob_pipe_hc"],
            "OB PO (L)": money_lakh(m["ob_po"]), "OB Margin (L)": money_lakh(m["ob_margin"]),
            "OB Margin %": pct(m["ob_margin"], m["ob_po"]),
            "Exit MTD": m["exit_hc"], "Exit Pipeline": m["exit_pipe_hc"], "Exit Projection": m["exit_hc"] + m["exit_pipe_hc"],
            "Exit PO (L)": money_lakh(m["exit_po"]), "Exit Margin (L)": money_lakh(m["exit_margin"]),
            "Exit Margin %": pct(m["exit_margin"], m["exit_po"]),
            "MTD Net": m["ob_hc"] - m["exit_hc"], "Net Projection": m["ob_hc"] + m["ob_pipe_hc"] - m["exit_hc"] - m["exit_pipe_hc"],
            "OB Projection PO (L)": 0.0, "OB Projection Margin (L)": 0.0,
            "Exit Projection PO (L)": 0.0, "Exit Projection Margin (L)": 0.0,
            "Net PO (L)": money_lakh(m["ob_po"]-m["exit_po"]),
            "Net Margin (L)": money_lakh(m["ob_margin"]-m["exit_margin"]),
            "Net Projection PO (L)": 0.0, "Net Projection Margin (L)": 0.0,
        })
    return pd.DataFrame(rows)


def bh_metric_event(data, period, as_of, filters, sheet, date_col, value_col=None, pipeline_sheet=None, pipeline_date=None):
    bhs = filters["bh"] or ALLOWED_BHS
    rows = []
    for bh in bhs:
        local = dict(filters); local["bh"] = [bh]
        actual = filter_df(data[sheet], sheet, local)
        actual = mtd(actual, date_col, period, as_of)
        pipe = pd.DataFrame()
        if pipeline_sheet:
            pipe = filter_df(data[pipeline_sheet], pipeline_sheet, local)
            pipe = pipeline(pipe, pipeline_date, period, as_of)
        actual_value = actual[value_col].sum() if value_col else len(actual)
        pipe_value = pipe[value_col].sum() if value_col else len(pipe)
        rows.append({"BH": bh, "MTD": actual_value, "Pipeline": pipe_value, "Projection": actual_value + pipe_value})
    return pd.DataFrame(rows)

# =========================================================
# GENERIC CHART HELPERS
# =========================================================
def bar_chart(df, x, y, title, color=None, horizontal=False, height=380):
    fig = px.bar(df, x=x, y=y, color=color, orientation="h" if horizontal else "v", text_auto=True, height=height)
    fig.update_layout(
        title=dict(text=title, x=0.02, font=dict(size=15)),
        margin=dict(l=10, r=10, t=45, b=45), paper_bgcolor="#f7f9fc", plot_bgcolor="#f7f9fc",
        legend_title="", hovermode="x unified",
    )
    return fig


def line_chart(df, x, y, title, color=None, height=390):
    fig = px.line(df, x=x, y=y, color=color, markers=True, height=height)
    fig.update_layout(
        title=dict(text=title, x=0.02, font=dict(size=15)),
        margin=dict(l=10, r=10, t=45, b=45), paper_bgcolor="#f7f9fc", plot_bgcolor="#f7f9fc",
        legend_title="", hovermode="x unified",
    )
    return fig


def doughnut(values, labels, title):
    fig = go.Figure(go.Pie(labels=labels, values=values, hole=.58, textinfo="label+value"))
    fig.update_layout(title=dict(text=title, x=.02, font=dict(size=15)), height=360, margin=dict(l=10,r=10,t=45,b=10), paper_bgcolor="#f7f9fc")
    return fig


def daily_event(data, sheet, date_col, period, as_of, filters, value_col=None):
    x = filter_df(data[sheet], sheet, filters)
    x = mtd(x, date_col, period, as_of)
    if x.empty:
        return pd.DataFrame(columns=["Date", "Value"])
    if value_col:
        s = x.groupby(date_col)[value_col].sum()
    else:
        s = x.groupby(date_col).size()
    out = s.rename("Value").reset_index().rename(columns={date_col:"Date"})
    return out


def month_series(data, sheet, date_col, filters, start_period, end_period, value_col=None):
    x = filter_df(data[sheet], sheet, filters)
    x = x[(x[date_col] >= start_period.start_time) & (x[date_col] <= end_period.end_time)]
    if x.empty:
        return pd.DataFrame(columns=["Month", "Value"])
    if value_col:
        s = x.groupby(x[date_col].dt.to_period("M"))[value_col].sum()
    else:
        s = x.groupby(x[date_col].dt.to_period("M")).size()
    out = s.rename("Value").reset_index()
    out["Month"] = out[date_col].astype(str)
    return out[["Month","Value"]]

# =========================================================
# LOAD WORKBOOK + FILTERS
# =========================================================
repo_bytes = get_repo_bytes()
if repo_bytes is not None:
    data = load_workbook(source_path=str(REPO_FILE))
else:
    st.error("CEO_MongoDB.xlsx was not found beside app.py.")
    st.stop()

periods = set()
for sheet, cfg in SHEET_MAP.items():
    d = data[sheet][cfg["date"]].dropna()
    if not d.empty:
        periods.update(d.dt.to_period("M").tolist())
periods = sorted(periods)
period_labels = [month_label(p) for p in periods]
if not periods:
    st.error("No valid reporting dates were found in the workbook.")
    st.stop()

latest_dates = []
for sheet, cfg in SHEET_MAP.items():
    d = data[sheet][cfg["date"]].dropna()
    if not d.empty:
        latest_dates.append(d.max())
latest_available = max(latest_dates) if latest_dates else periods[-1].start_time
latest_label = month_label(latest_available)
default_idx = period_labels.index(latest_label) if latest_label in period_labels else len(period_labels) - 1

# =========================================================
# CLEAN EXECUTIVE THEME
# =========================================================
st.markdown("""
<style>
.stApp { background:#eef3f8; }
.block-container { max-width:1800px; padding:1.2rem 2rem 2.5rem; }
[data-testid="stHeader"] { background:#eef3f8; }
.hero { background:linear-gradient(135deg,#0b315b,#154f7d); border-radius:14px; padding:20px 24px; margin-bottom:14px; box-shadow:0 4px 16px rgba(21,53,84,.12); }
.hero h1 { color:#fff; margin:0; font-size:2rem; letter-spacing:-.025em; }
.hero p { color:#dceaf5; margin:5px 0 0; font-size:.9rem; }
.section-title { color:#19324d; font-size:1.1rem; font-weight:800; margin:15px 0 8px; }
.metric { background:#fff; border:1px solid #dbe4ec; border-radius:12px; padding:13px 15px; min-height:100px; box-shadow:0 2px 8px rgba(31,55,78,.06); }
.metric .label { color:#607286; font-size:.75rem; font-weight:800; text-transform:uppercase; }
.metric .value { color:#173b60; font-size:1.75rem; font-weight:850; line-height:1.15; margin-top:5px; }
.metric .sub { color:#77889a; font-size:.72rem; margin-top:5px; }
.metric.green { border-top:4px solid #2d8a55; }.metric.red{border-top:4px solid #d55a5a}.metric.blue{border-top:4px solid #2877b7}.metric.gold{border-top:4px solid #d79a35}.metric.dark{border-top:4px solid #496a86}
div[data-testid="stTabs"] button { color:#3c5268; font-weight:750; }
div[data-testid="stTabs"] button[aria-selected="true"] { color:#0c4c78; }
.stDataFrame { border-radius:10px; overflow:hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""<div class="hero"><h1>CEO BUSINESS PERFORMANCE COCKPIT</h1><p>Executive view &nbsp;•&nbsp; PO / Margin shown in ₹ Lakhs</p></div>""", unsafe_allow_html=True)

with st.expander("🔎 Filters", expanded=True):
    f1,f2,f3,f4,f5,f6 = st.columns([1.1,1.1,1.1,1.1,1.4,1.4])
    with f1:
        selected_label = st.selectbox("Reporting Month", period_labels, index=default_idx)
    selected_period = period_from_label(selected_label)
    today = pd.Timestamp.today().normalize()
    date_min = selected_period.start_time.date()
    date_max = max(date_min, min(selected_period.end_time.date(), today.date()))
    month_dates=[]
    for sheet,cfg in SHEET_MAP.items():
        d=data[sheet][cfg["date"]].dropna()
        d=d[d.dt.to_period("M")==selected_period]
        if not d.empty: month_dates.append(d.max())
    default_asof=max(selected_period.start_time.normalize(), min(max(month_dates) if month_dates else today, today, pd.Timestamp(date_max)))
    default_asof=max(selected_period.start_time.normalize(), min(default_asof,pd.Timestamp(date_max)))
    with f2:
        as_of=pd.Timestamp(st.date_input("DOD / As-of", value=default_asof.date(), min_value=date_min, max_value=date_max))
    bh_values=universe_values(data,"BH"); kam_values=universe_values(data,"KAM"); domain_values=universe_values(data,"Domain")
    client_values=sorted(set(universe_values(data,"company_name")+universe_values(data,"Company_name")+universe_values(data,"client")))
    with f3: selected_bh=st.multiselect("Business Head",bh_values,placeholder="All BHs")
    with f4: selected_kam=st.multiselect("KAM",kam_values,placeholder="All KAMs")
    with f5: selected_domain=st.multiselect("Domain",domain_values,placeholder="All Domains")
    with f6: selected_client=st.multiselect("Client",client_values,placeholder="All Clients")
    g1,g2=st.columns([1,3])
    with g1: analysis_view=st.selectbox("Breakdown By",["Business Head","Client","KAM","Domain"])
    with g2: st.caption("Every chart responds to the selected Month, DOD, BH, KAM, Domain and Client. Change Breakdown By for the comparison dimension.")

filters={"bh":selected_bh,"kam":selected_kam,"domain":selected_domain,"client":selected_client}

def selected_dimension(sheet):
    return {"Business Head":"BH","Client":SHEET_MAP[sheet]["client"],"KAM":"KAM","Domain":"Domain"}[analysis_view]

def clean_chart(fig, title):
    fig.update_layout(title=dict(text=title,x=.02,font=dict(size=15,color="#263b50")),paper_bgcolor="#ffffff",plot_bgcolor="#ffffff",font=dict(color="#53677b"),margin=dict(l=15,r=15,t=48,b=45),hovermode="x unified")
    fig.update_xaxes(showgrid=False); fig.update_yaxes(gridcolor="#e7edf3")
    return fig

def monthly_metric(sheet,date_col,value_col=None):
    x=filter_df(data[sheet],sheet,filters); start=selected_period-8
    x=x[(x[date_col]>=start.start_time)&(x[date_col]<=selected_period.end_time)]
    rows=[]
    for p in pd.period_range(start,selected_period,freq="M"):
        z=x[x[date_col].dt.to_period("M")==p]
        rows.append({"Month":month_label(p),"Value":z[value_col].sum() if value_col else len(z)})
    return pd.DataFrame(rows)

def breakdown_metric(sheet,date_col,value_col=None):
    x=mtd(filter_df(data[sheet],sheet,filters),date_col,selected_period,as_of)
    dim=selected_dimension(sheet)
    if x.empty: return pd.DataFrame(columns=[dim,"Value"])
    if analysis_view=="Client":
        dim=SHEET_MAP[sheet]["client"]
        group_cols=["Domain","BH","KAM",dim]
    elif analysis_view=="KAM":
        x=x.assign(**{"KAM View":x["BH"]+" | "+x["KAM"]})
        dim="KAM View"
        group_cols=["Domain","BH","KAM",dim]
    else:
        group_cols=[dim]
    if value_col:
        out=x.groupby(group_cols,dropna=False)[value_col].sum().reset_index(name="Value")
    else:
        out=x.groupby(group_cols,dropna=False).size().reset_index(name="Value")
    return ordered_sort(out,"Value",group_cols=[c for c in ["Domain","BH","KAM"] if c in out.columns])

def dod_metric(sheet,date_col,value_col=None):
    x=mtd(filter_df(data[sheet],sheet,filters),date_col,selected_period,as_of)
    if x.empty: return pd.DataFrame(columns=["Date","Value"])
    out=x.groupby(date_col)[value_col].sum().reset_index(name="Value") if value_col else x.groupby(date_col).size().reset_index(name="Value")
    return out.rename(columns={date_col:"Date"})

def render_operational_tab(sheet,date_col,label,value_col=None,color="blue",conversion=None):
    x=mtd(filter_df(data[sheet],sheet,filters),date_col,selected_period,as_of); total=x[value_col].sum() if value_col else len(x); days=max(1,(as_of-selected_period.start_time.normalize()).days+1); dim=selected_dimension(sheet)
    cards=st.columns(4)
    with cards[0]: metric_card(f"MTD {label}",fmt_lakh(total) if value_col else fmt_int(total),"₹ Lakhs" if value_col else "Count",color)
    with cards[1]: metric_card("Avg / Day",f"{total/days:.1f}","MTD run-rate","dark")
    with cards[2]: metric_card(f"{dim} Count",fmt_int(x[dim].nunique()) if not x.empty else "0",f"Distinct {dim}","gold")
    with cards[3]: metric_card("Conversion",f"{conversion():.1f}%" if conversion else "—","Current filtered view","green")
    c1,c2=st.columns(2)
    with c1:
        mm=monthly_metric(sheet,date_col,value_col); st.plotly_chart(clean_chart(px.bar(mm,x="Month",y="Value",text_auto=True,height=360),f"Month-on-Month {label}"),use_container_width=True)
    with c2:
        dd=dod_metric(sheet,date_col,value_col)
        fig_dod=px.line(dd,x="Date",y="Value",markers=True,text="Value",height=410)
        fig_dod.update_traces(textposition="top center", cliponaxis=False, marker=dict(size=7), line=dict(width=3))
        fig_dod.update_xaxes(tickformat="%d-%b", tickangle=-35, nticks=min(12, max(4, len(dd))))
        fig_dod.update_yaxes(automargin=True)
        st.plotly_chart(clean_chart(fig_dod,f"DOD {label} — {selected_label}"),use_container_width=True)
    bd=breakdown_metric(sheet,date_col,value_col); st.plotly_chart(clean_chart(px.bar(bd,x="Value",y=dim,orientation="h",text_auto=True,height=max(360,min(650,120+len(bd)*22))),f"{label} by {dim}"),use_container_width=True)
    return x

# =========================================================
# CEO OVERVIEW
# =========================================================
m=overall_metrics(data,selected_period,as_of,filters); ob_proj=m["ob_hc"]+m["ob_pipe_hc"]; ex_proj=m["exit_hc"]+m["exit_pipe_hc"]; net_mtd=m["ob_hc"]-m["exit_hc"]; net_proj=ob_proj-ex_proj
st.markdown('<div class="section-title">CEO Snapshot</div>',unsafe_allow_html=True)
r=st.columns(7)
items=[("Demand",fmt_int(m["demand"]),"MTD","blue"),("Submission",fmt_int(m["submission"]),"MTD","blue"),("Interview",fmt_int(m["interview"]),"MTD","blue"),("Selection",fmt_int(m["selection"]),"MTD","blue"),("Onboarding",fmt_int(m["ob_hc"]),f"PO {fmt_lakh(m['ob_po'])}","green"),("Exit",fmt_int(m["exit_hc"]),f"PO {fmt_lakh(m['exit_po'])}","red"),("Net",fmt_int(net_mtd),"OB − Exit","dark")]
for col,(lab,val,sub,kind) in zip(r,items):
    with col: metric_card(lab,val,sub,kind)
ob_proj_po=m["ob_po"]+m["ob_pipe_po"]; ob_proj_margin=m["ob_margin"]+m["ob_pipe_margin"]
ex_proj_po=m["exit_po"]+m["exit_pipe_po"]; ex_proj_margin=m["exit_margin"]+m["exit_pipe_margin"]
net_proj_po=ob_proj_po-ex_proj_po; net_proj_margin=ob_proj_margin-ex_proj_margin
r2=st.columns(3)
projection_cards=[
    ("OB PROJECTION",ob_proj,ob_proj_po,ob_proj_margin,"green"),
    ("EXIT PROJECTION",ex_proj,ex_proj_po,ex_proj_margin,"red"),
    ("NET PROJECTION",net_proj,net_proj_po,net_proj_margin,"dark"),
]
for col,(lab,hc,po_v,mar_v,kind) in zip(r2,projection_cards):
    with col:
        html=f'<div class="metric {kind}"><div class="label">{lab}</div><div class="value">{fmt_int(hc)} HC</div><div class="sub">PO {fmt_lakh(po_v)} &nbsp; | &nbsp; Margin {fmt_lakh(mar_v)}</div></div>'
        st.markdown(html,unsafe_allow_html=True)

st.markdown('<div class="section-title">CEO Trend Board — separate metric charts</div>',unsafe_allow_html=True)
chart_specs=[("Demand","Demand","Created_at",None),("Submission","Submission","date",None),("Interview","Interview","Interview_date",None),("Selection","Selection","selection_date",None),("Onboarding","Onboarding","display_date",None),("Exit","Exit","last_work_day",None)]
for idx in range(0,len(chart_specs),2):
    cols=st.columns(2)
    for col,spec in zip(cols,chart_specs[idx:idx+2]):
        label,sheet,dc,vc=spec; mm=monthly_metric(sheet,dc,vc)
        with col: st.plotly_chart(clean_chart(px.bar(mm,x="Month",y="Value",text_auto=True,height=330),f"{label} — Month-on-Month"),use_container_width=True)

tabs=st.tabs(["Demand","Submission","Interview","Selection","Onboarding","Exit","Net & Projection"])
with tabs[0]:
    render_operational_tab("Demand","Created_at","Demand",None,"blue",lambda:100*m["submission"]/m["demand"] if m["demand"] else 0)
    d=mtd(filter_df(data["Demand"],"Demand",filters),"Created_at",selected_period,as_of)
    detail=d.groupby(["Company_name","BH","KAM","Domain"],dropna=False)["no_of_opening"].sum().reset_index()
    detail=ordered_sort(detail,"no_of_opening")
    st.dataframe(detail,use_container_width=True,hide_index=True,height=360)
with tabs[1]:
    render_operational_tab("Submission","date","Submission",None,"blue",lambda:100*m["submission"]/m["demand"] if m["demand"] else 0)
with tabs[2]:
    render_operational_tab("Interview","Interview_date","Interview",None,"blue",lambda:100*m["interview"]/m["submission"] if m["submission"] else 0)
with tabs[3]:
    render_operational_tab("Selection","selection_date","Selection",None,"blue",lambda:100*m["selection"]/m["interview"] if m["interview"] else 0)
with tabs[4]:
    ob=mtd(filter_df(data["Onboarding"],"Onboarding",filters),"display_date",selected_period,as_of); obp=pipeline(filter_df(data["Onboarding Pipeline"],"Onboarding Pipeline",filters),"display_date",selected_period,as_of); po=ob["p_o_value"].sum(); mar=ob["margin"].sum()
    c=st.columns(7); vals=[("MTD OB HC",fmt_int(len(ob)),"Actual","green"),("MTD OB PO",fmt_lakh(po),"₹ Lakhs","green"),("MTD OB Margin",fmt_lakh(mar),"₹ Lakhs","green"),("Margin %",f"{pct(mar,po):.1f}%","Margin / PO","gold"),("OB Pipeline",fmt_int(len(obp)),"Remaining after DOD","dark"),("OB Projection",fmt_int(len(ob)+len(obp)),"MTD + Pipeline","green"),("Projected PO",fmt_lakh(ob["p_o_value"].sum()+obp["p_o_value"].sum()),"MTD + Pipeline","dark")]
    for col,(lab,val,sub,kind) in zip(c,vals):
        with col: metric_card(lab,val,sub,kind)
    mm=monthly_metric("Onboarding","display_date",None); st.plotly_chart(clean_chart(px.bar(mm,x="Month",y="Value",text_auto=True,height=380),"Onboarding — Month-on-Month HC"),use_container_width=True)
    c1,c2=st.columns(2)
    with c1:
        dd=dod_metric("Onboarding","display_date"); st.plotly_chart(clean_chart(px.line(dd,x="Date",y="Value",markers=True,text="Value",height=350),"Onboarding — DOD Trend"),use_container_width=True)
    with c2:
        bd=breakdown_metric("Onboarding","display_date"); dim=selected_dimension("Onboarding"); st.plotly_chart(clean_chart(px.bar(bd,x="Value",y=dim,orientation="h",text_auto=True,height=350),f"Onboarding by {dim}"),use_container_width=True)
    po_mm=monthly_metric("Onboarding","display_date","p_o_value"); po_mm["Value"]/=100000; mar_mm=monthly_metric("Onboarding","display_date","margin"); mar_mm["Value"]/=100000
    c1,c2=st.columns(2)
    with c1: st.plotly_chart(clean_chart(px.bar(po_mm,x="Month",y="Value",text_auto=".1f",height=330),"Onboarding PO — Month-on-Month (₹ Lakhs)"),use_container_width=True)
    with c2: st.plotly_chart(clean_chart(px.bar(mar_mm,x="Month",y="Value",text_auto=".1f",height=330),"Onboarding Margin — Month-on-Month (₹ Lakhs)"),use_container_width=True)
    detail=ob.groupby(["company_name","BH","KAM","Domain"]).agg(HC=("full_name","size"),PO=("p_o_value","sum"),Margin=("margin","sum")).reset_index(); detail["PO (L)"]=detail["PO"]/100000; detail["Margin (L)"]=detail["Margin"]/100000; detail["Margin %"]=np.where(detail["PO"]!=0,detail["Margin"]/detail["PO"]*100,0); detail=ordered_sort(detail,"HC")
    st.dataframe(detail[["company_name","BH","KAM","Domain","HC","PO (L)","Margin (L)","Margin %"]].sort_values("HC",ascending=False).style.format({"PO (L)":"{:,.1f}","Margin (L)":"{:,.1f}","Margin %":"{:,.1f}%"}),use_container_width=True,height=380,hide_index=True)
with tabs[5]:
    ex=mtd(filter_df(data["Exit"],"Exit",filters),"last_work_day",selected_period,as_of); exp=pipeline(filter_df(data["Exit Pipeline"],"Exit Pipeline",filters),"tentative_exit_date",selected_period,as_of); po=ex["p_o_value"].sum(); mar=ex["margin"].sum()
    c=st.columns(7); vals=[("MTD Exit HC",fmt_int(len(ex)),"Actual","red"),("MTD Exit PO",fmt_lakh(po),"₹ Lakhs","red"),("MTD Exit Margin",fmt_lakh(mar),"₹ Lakhs","red"),("Margin %",f"{pct(mar,po):.1f}%","Margin / PO","gold"),("Exit Pipeline",fmt_int(len(exp)),"Remaining after DOD","dark"),("Exit Projection",fmt_int(len(ex)+len(exp)),"MTD + Pipeline","red"),("Projected PO",fmt_lakh(ex["p_o_value"].sum()+exp["p_o_value"].sum()),"MTD + Pipeline","dark")]
    for col,(lab,val,sub,kind) in zip(c,vals):
        with col: metric_card(lab,val,sub,kind)
    mm=monthly_metric("Exit","last_work_day",None); st.plotly_chart(clean_chart(px.bar(mm,x="Month",y="Value",text_auto=True,height=380),"Exit — Month-on-Month HC"),use_container_width=True)
    c1,c2=st.columns(2)
    with c1:
        dd=dod_metric("Exit","last_work_day"); st.plotly_chart(clean_chart(px.line(dd,x="Date",y="Value",markers=True,text="Value",height=350),"Exit — DOD Trend"),use_container_width=True)
    with c2:
        bd=breakdown_metric("Exit","last_work_day"); dim=selected_dimension("Exit"); st.plotly_chart(clean_chart(px.bar(bd,x="Value",y=dim,orientation="h",text_auto=True,height=350),f"Exit by {dim}"),use_container_width=True)
    po_mm=monthly_metric("Exit","last_work_day","p_o_value"); po_mm["Value"]/=100000; mar_mm=monthly_metric("Exit","last_work_day","margin"); mar_mm["Value"]/=100000
    c1,c2=st.columns(2)
    with c1: st.plotly_chart(clean_chart(px.bar(po_mm,x="Month",y="Value",text_auto=".1f",height=330),"Exit PO — Month-on-Month (₹ Lakhs)"),use_container_width=True)
    with c2: st.plotly_chart(clean_chart(px.bar(mar_mm,x="Month",y="Value",text_auto=".1f",height=330),"Exit Margin — Month-on-Month (₹ Lakhs)"),use_container_width=True)
    detail=ex.groupby(["company_name","BH","KAM","Domain"]).agg(HC=("full_name","size"),PO=("p_o_value","sum"),Margin=("margin","sum")).reset_index(); detail["PO (L)"]=detail["PO"]/100000; detail["Margin (L)"]=detail["Margin"]/100000; detail["Margin %"]=np.where(detail["PO"]!=0,detail["Margin"]/detail["PO"]*100,0); detail=ordered_sort(detail,"HC")
    st.dataframe(detail[["company_name","BH","KAM","Domain","HC","PO (L)","Margin (L)","Margin %"]].sort_values("HC",ascending=False).style.format({"PO (L)":"{:,.1f}","Margin (L)":"{:,.1f}","Margin %":"{:,.1f}%"}),use_container_width=True,height=380,hide_index=True)
with tabs[6]:
    c=st.columns(6); vals=[("OB MTD",m["ob_hc"],"Actual","green"),("OB Pipeline",m["ob_pipe_hc"],"Remaining","green"),("OB Projection",ob_proj,"MTD + Pipeline","green"),("Exit MTD",m["exit_hc"],"Actual","red"),("Exit Pipeline",m["exit_pipe_hc"],"Remaining","red"),("Net Projection",net_proj,"OB Projection − Exit Projection","dark")]
    for col,(lab,val,sub,kind) in zip(c,vals):
        with col: metric_card(lab,fmt_int(val),sub,kind)
    net_mom=[]
    for p in pd.period_range(selected_period-8,selected_period,freq="M"):
        obx=filter_df(data["Onboarding"],"Onboarding",filters); obx=obx[obx["display_date"].dt.to_period("M")==p]; exx=filter_df(data["Exit"],"Exit",filters); exx=exx[exx["last_work_day"].dt.to_period("M")==p]
        net_mom.append({"Month":month_label(p),"Net":len(obx)-len(exx)})
    nm=pd.DataFrame(net_mom); c1,c2=st.columns(2)
    with c1: st.plotly_chart(clean_chart(px.bar(nm,x="Month",y="Net",text_auto=True,height=370),"Net HC — Month-on-Month"),use_container_width=True)
    with c2: st.plotly_chart(clean_chart(px.line(nm,x="Month",y="Net",markers=True,text="Net",height=370),"Net HC Trend"),use_container_width=True)
    st.markdown('<div class="section-title">BH Net / Projection</div>',unsafe_allow_html=True)
    score=bh_scorecard(data,selected_period,as_of,filters)
    # CEO scorecard order: Actual OB/Exit/Net, then Projection OB/Exit/Net.
    score["Net PO (L)"] = score["OB PO (L)"] - score["Exit PO (L)"]
    score["Net Margin (L)"] = score["OB Margin (L)"] - score["Exit Margin (L)"]
    # Initialize every projection financial column before row-wise enrichment.
    score["OB Projection PO (L)"] = 0.0
    score["OB Projection Margin (L)"] = 0.0
    score["Exit Projection PO (L)"] = 0.0
    score["Exit Projection Margin (L)"] = 0.0
    # Pull pipeline PO/Margin directly so projected financials are MTD + remaining pipeline.
    for i, row in score.iterrows():
        local = dict(filters); local["bh"] = [row["BH"]]
        mm = overall_metrics(data, selected_period, as_of, local)
        score.loc[i, "OB Projection PO (L)"] = money_lakh(mm["ob_po"] + mm["ob_pipe_po"])
        score.loc[i, "OB Projection Margin (L)"] = money_lakh(mm["ob_margin"] + mm["ob_pipe_margin"])
        score.loc[i, "Exit Projection PO (L)"] = money_lakh(mm["exit_po"] + mm["exit_pipe_po"])
        score.loc[i, "Exit Projection Margin (L)"] = money_lakh(mm["exit_margin"] + mm["exit_pipe_margin"])
    score["Net Projection PO (L)"] = score["OB Projection PO (L)"] - score["Exit Projection PO (L)"]
    score["Net Projection Margin (L)"] = score["OB Projection Margin (L)"] - score["Exit Projection Margin (L)"]
    bh_cols = [
        "BH",
        "OB MTD", "OB PO (L)", "OB Margin (L)",
        "Exit MTD", "Exit PO (L)", "Exit Margin (L)",
        "MTD Net", "Net PO (L)", "Net Margin (L)",
        "OB Projection", "OB Projection PO (L)", "OB Projection Margin (L)",
        "Exit Projection", "Exit Projection PO (L)", "Exit Projection Margin (L)",
        "Net Projection", "Net Projection PO (L)", "Net Projection Margin (L)"
    ]
    st.dataframe(
        ordered_sort(score[bh_cols],"Net Projection",group_cols=[]).style.format({c:"{:,.1f}" for c in bh_cols if c != "BH" and c not in ["OB MTD","Exit MTD","MTD Net","OB Projection","Exit Projection","Net Projection"]}),
        use_container_width=True, height=440, hide_index=True
    )

    st.markdown('<div class="section-title">Client Level Net / Projection</div>',unsafe_allow_html=True)
    client_tbl = client_metrics(data, selected_period, as_of, filters)
    if client_tbl.empty:
        st.info("No client records match the selected filters.")
    else:
        client_tbl["Net PO (L)"] = client_tbl["OB PO (L)"] - client_tbl["Exit PO (L)"]
        client_tbl["Net Margin (L)"] = client_tbl["OB Margin (L)"] - client_tbl["Exit Margin (L)"]
        # Financial projection = actual MTD + remaining pipeline.
        for i, row in client_tbl.iterrows():
            local = dict(filters); local["client"] = [row["Client"]]
            mm = overall_metrics(data, selected_period, as_of, local)
            client_tbl.loc[i, "OB Projection PO (L)"] = money_lakh(mm["ob_po"] + mm["ob_pipe_po"])
            client_tbl.loc[i, "OB Projection Margin (L)"] = money_lakh(mm["ob_margin"] + mm["ob_pipe_margin"])
            client_tbl.loc[i, "Exit Projection PO (L)"] = money_lakh(mm["exit_po"] + mm["exit_pipe_po"])
            client_tbl.loc[i, "Exit Projection Margin (L)"] = money_lakh(mm["exit_margin"] + mm["exit_pipe_margin"])
        client_tbl["Net Projection PO (L)"] = client_tbl["OB Projection PO (L)"] - client_tbl["Exit Projection PO (L)"]
        client_tbl["Net Projection Margin (L)"] = client_tbl["OB Projection Margin (L)"] - client_tbl["Exit Projection Margin (L)"]
        client_cols = [
            "Client", "BH", "KAM", "Domain",
            "OB MTD", "OB PO (L)", "OB Margin (L)",
            "Exit MTD", "Exit PO (L)", "Exit Margin (L)",
            "MTD Net", "Net PO (L)", "Net Margin (L)",
            "OB Projection", "OB Projection PO (L)", "OB Projection Margin (L)",
            "Exit Projection", "Exit Projection PO (L)", "Exit Projection Margin (L)",
            "Net Projection", "Net Projection PO (L)", "Net Projection Margin (L)"
        ]
        st.dataframe(
            ordered_sort(client_tbl[client_cols],"Net Projection").style.format({c:"{:,.1f}" for c in client_cols if c not in ["Client","BH","KAM","Domain"] and c not in ["OB MTD","Exit MTD","MTD Net","OB Projection","Exit Projection","Net Projection"]}),
            use_container_width=True, height=480, hide_index=True
        )

st.caption("CEO Business Performance Cockpit • Demand: Created_at • Submission: date • Interview: Interview_date • Selection: selection_date • Onboarding: display_date • Exit: last_work_day • Pipelines: after DOD through month-end • PO/Margin: ₹ Lakhs")
