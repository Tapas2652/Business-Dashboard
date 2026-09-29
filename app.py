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
    cfg = SHEET_MAP[sheet]
    x = df.copy()
    for field, values in [("BH", filters["bh"]), ("KAM", filters["kam"]), ("Domain", filters["domain"]), (cfg["client"], filters["client"])]:
        if values:
            x = x[x[field].isin(values)]
    return x


def universe_values(data, column):
    vals = set()
    for sheet in SHEET_MAP:
        if column in data[sheet].columns:
            vals.update(data[sheet][column].dropna().astype(str).unique())
    return sorted(vals)

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
    bhs = filters["bh"] or universe_values(data, "BH")
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
        rows.append({
            "Client": client,
            "BH": (ah["BH"].mode().iat[0] if not ah.empty else ""),
            "KAM": (ah["KAM"].mode().iat[0] if not ah.empty else ""),
            "Demand": m["demand"], "Submission": m["submission"], "Interview": m["interview"], "Selection": m["selection"],
            "OB MTD": m["ob_hc"], "OB Pipeline": m["ob_pipe_hc"], "OB Projection": m["ob_hc"] + m["ob_pipe_hc"],
            "OB PO (L)": money_lakh(m["ob_po"]), "OB Margin (L)": money_lakh(m["ob_margin"]),
            "OB Margin %": pct(m["ob_margin"], m["ob_po"]),
            "Exit MTD": m["exit_hc"], "Exit Pipeline": m["exit_pipe_hc"], "Exit Projection": m["exit_hc"] + m["exit_pipe_hc"],
            "Exit PO (L)": money_lakh(m["exit_po"]), "Exit Margin (L)": money_lakh(m["exit_margin"]),
            "Exit Margin %": pct(m["exit_margin"], m["exit_po"]),
            "MTD Net": m["ob_hc"] - m["exit_hc"], "Net Projection": m["ob_hc"] + m["ob_pipe_hc"] - m["exit_hc"] - m["exit_pipe_hc"],
        })
    return pd.DataFrame(rows)


def bh_metric_event(data, period, as_of, filters, sheet, date_col, value_col=None, pipeline_sheet=None, pipeline_date=None):
    bhs = filters["bh"] or universe_values(data, "BH")
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
# LOAD WORKBOOK
# =========================================================
repo_bytes = get_repo_bytes()

with st.sidebar:
    st.markdown("### ⚙️ Dashboard Controls")
    uploaded = st.file_uploader("Test updated Excel", type=["xlsx"], help="Production: replace CEO_MongoDB.xlsx in GitHub.")
    if st.button("↻ Refresh Data", use_container_width=True):
        st.cache_data.clear(); st.rerun()

if uploaded is not None:
    data = load_workbook(source_bytes=uploaded.getvalue())
    source_name = uploaded.name
elif repo_bytes is not None:
    data = load_workbook(source_path=str(REPO_FILE))
    source_name = "CEO_MongoDB.xlsx"
else:
    st.error("CEO_MongoDB.xlsx was not found beside app.py.")
    st.stop()

# =========================================================
# PERIOD + GLOBAL FILTERS
# =========================================================
periods = set()
for s, cfg in SHEET_MAP.items():
    d = data[s][cfg["date"]].dropna()
    periods.update(d.dt.to_period("M").tolist())
periods = sorted(periods)
period_labels = [month_label(p) for p in periods]
if not period_labels:
    st.error("No valid reporting dates were found in the workbook.")
    st.stop()

latest_dates = [data[s][SHEET_MAP[s]["date"]].dropna().max() for s in SHEET_MAP]
latest_dates = [d for d in latest_dates if pd.notna(d)]
latest_available = max(latest_dates) if latest_dates else periods[-1].start_time
latest_label = month_label(latest_available)
default_idx = period_labels.index(latest_label) if latest_label in period_labels else len(period_labels) - 1

with st.sidebar:
    selected_label = st.selectbox("Reporting Month", period_labels, index=default_idx)
    selected_period = period_from_label(selected_label)

    # As-of range is always valid, including future workbook months.
    today = pd.Timestamp.today().normalize()
    date_min = selected_period.start_time.date()
    date_max = max(date_min, min(selected_period.end_time.date(), today.date()))
    month_dates = []
    for s, cfg in SHEET_MAP.items():
        d = data[s][cfg["date"]].dropna()
        d = d[d.dt.to_period("M") == selected_period]
        if not d.empty: month_dates.append(d.max())
    default_asof = max(selected_period.start_time.normalize(), min(max(month_dates) if month_dates else today, today, selected_period.end_time.normalize()))
    default_asof = max(selected_period.start_time.normalize(), min(default_asof, pd.Timestamp(date_max)))
    as_of = pd.Timestamp(st.date_input("DOD / As-of Date", value=default_asof.date(), min_value=date_min, max_value=date_max))

    bh_values = universe_values(data, "BH")
    kam_values = universe_values(data, "KAM")
    domain_values = universe_values(data, "Domain")
    client_values = universe_values(data, "company_name")
    # Include Demand/Submission client variants.
    client_values = sorted(set(client_values + universe_values(data, "Company_name") + universe_values(data, "client")))

    selected_bh = st.multiselect("Business Head", bh_values)
    selected_kam = st.multiselect("KAM", kam_values)
    selected_domain = st.multiselect("Domain", domain_values)
    selected_client = st.multiselect("Client", client_values)

filters = {"bh": selected_bh, "kam": selected_kam, "domain": selected_domain, "client": selected_client}

# =========================================================
# HERO
# =========================================================
st.markdown(
    f'''<div class="hero"><h1>📊 CEO BUSINESS PERFORMANCE COCKPIT</h1>
    <p>{selected_label} &nbsp;•&nbsp; DOD {as_of.strftime('%d %b %Y')} &nbsp;•&nbsp; PO / Margin shown in ₹ Lakhs &nbsp;•&nbsp; Source: {source_name}</p></div>''',
    unsafe_allow_html=True,
)

# =========================================================
# CEO OVERVIEW
# =========================================================
m = overall_metrics(data, selected_period, as_of, filters)
ob_proj = m["ob_hc"] + m["ob_pipe_hc"]
ex_proj = m["exit_hc"] + m["exit_pipe_hc"]
net_mtd = m["ob_hc"] - m["exit_hc"]
net_proj = ob_proj - ex_proj

st.markdown('<div class="section-title">CEO Snapshot — MTD & Projection</div>', unsafe_allow_html=True)
row1 = st.columns(6)
for col, label, value, sub, kind in [
    (row1[0], "MTD Demand", fmt_int(m["demand"]), "Openings", "blue"),
    (row1[1], "MTD Submission", fmt_int(m["submission"]), "Candidates submitted", "blue"),
    (row1[2], "MTD Interview", fmt_int(m["interview"]), "Interviews", "blue"),
    (row1[3], "MTD Selection", fmt_int(m["selection"]), "Selections", "blue"),
    (row1[4], "MTD Onboarding", fmt_int(m["ob_hc"]), f"PO {fmt_lakh(m['ob_po'])} | Margin {fmt_lakh(m['ob_margin'])}", "green"),
    (row1[5], "MTD Exit", fmt_int(m["exit_hc"]), f"PO {fmt_lakh(m['exit_po'])} | Margin {fmt_lakh(m['exit_margin'])}", "red"),
]:
    with col: metric_card(label, value, sub, kind)

row2 = st.columns(6)
for col, label, value, sub, kind in [
    (row2[0], "MTD Net", fmt_int(net_mtd), "Onboarding − Exit", "green"),
    (row2[1], "OB Projection", fmt_int(ob_proj), f"{m['ob_pipe_hc']:,} pipeline", "green"),
    (row2[2], "Exit Projection", fmt_int(ex_proj), f"{m['exit_pipe_hc']:,} pipeline", "red"),
    (row2[3], "Net Projection", fmt_int(net_proj), "OB Projection − Exit Projection", "dark"),
    (row2[4], "OB Margin %", f"{pct(m['ob_margin'], m['ob_po']):.1f}%", f"PO {fmt_lakh(m['ob_po'])}", "gold"),
    (row2[5], "Exit Margin %", f"{pct(m['exit_margin'], m['exit_po']):.1f}%", f"PO {fmt_lakh(m['exit_po'])}", "gold"),
]:
    with col: metric_card(label, value, sub, kind)

# Main CEO charts
bh = bh_scorecard(data, selected_period, as_of, filters)
if not bh.empty:
    c1, c2 = st.columns(2)
    with c1:
        plot = bh[["BH","Demand","Submission","Interview","Selection"]].melt("BH", var_name="Metric", value_name="Count")
        st.plotly_chart(bar_chart(plot, "BH", "Count", "BH-wise Funnel — MTD", "Metric", height=410), use_container_width=True)
    with c2:
        plot = bh[["BH","OB MTD","OB Pipeline","Exit MTD","Exit Pipeline"]].melt("BH", var_name="Metric", value_name="HC")
        st.plotly_chart(bar_chart(plot, "BH", "HC", "BH-wise Onboarding vs Exit", "Metric", height=410), use_container_width=True)

    c3, c4 = st.columns([1.3, 1])
    with c3:
        st.markdown('<div class="section-title">CEO BH Scorecard</div>', unsafe_allow_html=True)
        cols = ["BH","Demand","Submission","Interview","Selection","OB MTD","OB Pipeline","OB Projection","Exit MTD","Exit Pipeline","Exit Projection","MTD Net","Net Projection","OB PO (L)","OB Margin (L)","Exit PO (L)","Exit Margin (L)"]
        st.dataframe(bh[cols].sort_values("Net Projection", ascending=False).style.format({c:"{:,.1f}" for c in cols if "(L)" in c}), use_container_width=True, height=430)
    with c4:
        st.plotly_chart(doughnut([m["ob_hc"], m["exit_hc"]], ["Onboarding","Exit"], "MTD HC Movement"), use_container_width=True)

# =========================================================
# ANALYTICAL TABS
# =========================================================
tabs = st.tabs(["🏠 CEO Overview", "📥 Demand", "📤 Submission", "🎯 Interview", "🟢 Onboarding", "🔴 Exit", "⚖️ Net & Projection"])

# ---- CEO Overview duplicate compact section ----
with tabs[0]:
    st.markdown('<div class="section-title">Executive Movement</div>', unsafe_allow_html=True)
    daily_parts = []
    for name, sheet, dc in [("Demand","Demand","Created_at"),("Submission","Submission","date"),("Interview","Interview","Interview_date"),("Selection","Selection","selection_date"),("Onboarding","Onboarding","display_date"),("Exit","Exit","last_work_day")]:
        dd = daily_event(data, sheet, dc, selected_period, as_of, filters)
        if not dd.empty:
            dd["Metric"] = name
            daily_parts.append(dd)
    if daily_parts:
        daily_all = pd.concat(daily_parts, ignore_index=True)
        st.plotly_chart(line_chart(daily_all, "Date", "Value", "Daily Operating Trend", "Metric", height=440), use_container_width=True)

    st.markdown('<div class="section-title">Conversion Indicators</div>', unsafe_allow_html=True)
    conv = pd.DataFrame({
        "Funnel": ["Demand → Submission", "Submission → Interview", "Interview → Selection", "Selection → Onboarding"],
        "Rate": [
            100*m["submission"]/m["demand"] if m["demand"] else 0,
            100*m["interview"]/m["submission"] if m["submission"] else 0,
            100*m["selection"]/m["interview"] if m["interview"] else 0,
            100*m["ob_hc"]/m["selection"] if m["selection"] else 0,
        ]
    })
    st.plotly_chart(bar_chart(conv, "Funnel", "Rate", "MTD Funnel Conversion %", height=350), use_container_width=True)

# ---- Demand ----
with tabs[1]:
    d = filter_df(data["Demand"], "Demand", filters)
    dm = mtd(d, "Created_at", selected_period, as_of)
    total = dm["no_of_opening"].sum()
    avg = total / max(1, (as_of - selected_period.start_time.normalize()).days + 1)
    c = st.columns(4)
    with c[0]: metric_card("MTD Demand", fmt_int(total), "Sum of no_of_opening", "blue")
    with c[1]: metric_card("Avg Demand / Day", f"{avg:.1f}", "MTD run-rate", "dark")
    with c[2]: metric_card("Clients with Demand", fmt_int(dm[SHEET_MAP['Demand']['client']].nunique()), "Distinct clients", "gold")
    with c[3]: metric_card("BHs with Demand", fmt_int(dm["BH"].nunique()), "Distinct BHs", "gold")
    c1,c2=st.columns(2)
    with c1:
        dd = dm.groupby("Created_at")["no_of_opening"].sum().reset_index(name="Demand")
        st.plotly_chart(line_chart(dd,"Created_at","Demand","DOD Demand Trend",height=390),use_container_width=True)
    with c2:
        cd = dm.groupby(SHEET_MAP["Demand"]["client"])["no_of_opening"].sum().reset_index().sort_values("no_of_opening",ascending=False).head(20)
        st.plotly_chart(bar_chart(cd,SHEET_MAP["Demand"]["client"],"no_of_opening","Top 20 Clients by Demand",height=390),use_container_width=True)
    bhd = dm.groupby("BH")["no_of_opening"].sum().reset_index().sort_values("no_of_opening",ascending=False)
    st.plotly_chart(bar_chart(bhd,"BH","no_of_opening","BH-wise MTD Demand",height=380),use_container_width=True)
    st.markdown('<div class="section-title">Client Demand Detail</div>',unsafe_allow_html=True)
    detail = dm.groupby([SHEET_MAP["Demand"]["client"],"BH","KAM","Domain"],dropna=False)["no_of_opening"].sum().reset_index().sort_values("no_of_opening",ascending=False)
    st.dataframe(detail,use_container_width=True,height=430,hide_index=True)

# ---- Submission ----
with tabs[2]:
    s = filter_df(data["Submission"], "Submission", filters)
    sm = mtd(s,"date",selected_period,as_of)
    c=st.columns(4)
    with c[0]: metric_card("MTD Submissions",fmt_int(len(sm)),"Candidate submissions","blue")
    with c[1]: metric_card("Avg / Day",f"{len(sm)/max(1,(as_of-selected_period.start_time.normalize()).days+1):.1f}","Submission run-rate","dark")
    with c[2]: metric_card("Active Clients",fmt_int(sm["client"].nunique()),"Distinct clients","gold")
    with c[3]: metric_card("Submission / Demand",f"{100*len(sm)/m['demand']:.1f}%" if m['demand'] else "0.0%","MTD conversion","green")
    c1,c2=st.columns(2)
    with c1:
        dd=sm.groupby("date").size().reset_index(name="Submissions")
        st.plotly_chart(line_chart(dd,"date","Submissions","DOD Submission Trend",height=390),use_container_width=True)
    with c2:
        cd=sm.groupby("client").size().reset_index(name="Submissions").sort_values("Submissions",ascending=False).head(20)
        st.plotly_chart(bar_chart(cd,"client","Submissions","Top 20 Clients by Submission",height=390),use_container_width=True)
    bhd=sm.groupby("BH").size().reset_index(name="Submissions").sort_values("Submissions",ascending=False)
    st.plotly_chart(bar_chart(bhd,"BH","Submissions","BH-wise MTD Submission",height=380),use_container_width=True)

# ---- Interview ----
with tabs[3]:
    i = filter_df(data["Interview"], "Interview", filters)
    im = mtd(i,"Interview_date",selected_period,as_of)
    c=st.columns(5)
    with c[0]: metric_card("MTD Interviews",fmt_int(len(im)),"Interview events","blue")
    with c[1]: metric_card("Avg / Day",f"{len(im)/max(1,(as_of-selected_period.start_time.normalize()).days+1):.1f}","Interview run-rate","dark")
    with c[2]: metric_card("Clients",fmt_int(im["company_name"].nunique()),"Distinct clients","gold")
    with c[3]: metric_card("Interview → Selection",f"{100*m['selection']/len(im):.1f}%" if len(im) else "0.0%","MTD conversion","green")
    with c[4]: metric_card("Selection",fmt_int(m["selection"]),"MTD","green")
    c1,c2=st.columns(2)
    with c1:
        dd=im.groupby("Interview_date").size().reset_index(name="Interviews")
        st.plotly_chart(line_chart(dd,"Interview_date","Interviews","DOD Interview Trend",height=390),use_container_width=True)
    with c2:
        cd=im.groupby("company_name").size().reset_index(name="Interviews").sort_values("Interviews",ascending=False).head(20)
        st.plotly_chart(bar_chart(cd,"company_name","Interviews","Top 20 Clients by Interview",height=390),use_container_width=True)
    bhd=im.groupby("BH").size().reset_index(name="Interviews").sort_values("Interviews",ascending=False)
    st.plotly_chart(bar_chart(bhd,"BH","Interviews","BH-wise MTD Interview",height=380),use_container_width=True)

# ---- Onboarding ----
with tabs[4]:
    ob = filter_df(data["Onboarding"],"Onboarding",filters)
    obm = mtd(ob,"display_date",selected_period,as_of)
    obp = pipeline(filter_df(data["Onboarding Pipeline"],"Onboarding Pipeline",filters),"display_date",selected_period,as_of)
    po=obm["p_o_value"].sum(); mar=obm["margin"].sum(); proj=len(obm)+len(obp)
    c=st.columns(6)
    with c[0]: metric_card("MTD OB HC",fmt_int(len(obm)),"Onboarded","green")
    with c[1]: metric_card("MTD OB PO",fmt_lakh(po),"₹ Lakhs","green")
    with c[2]: metric_card("MTD OB Margin",fmt_lakh(mar),"₹ Lakhs","green")
    with c[3]: metric_card("OB Margin %",f"{pct(mar,po):.1f}%","Margin / PO","gold")
    with c[4]: metric_card("OB Pipeline",fmt_int(len(obp)),"After DOD in month","dark")
    with c[5]: metric_card("OB Projection",fmt_int(proj),"MTD + Pipeline","green")

    # MOM onboarding — last 9 months including selected month
    start9 = selected_period - 8
    months = pd.period_range(start9, selected_period, freq="M")
    mom_rows=[]
    for p in months:
        x=filter_df(data["Onboarding"],"Onboarding",filters)
        x=x[x["display_date"].dt.to_period("M")==p]
        mom_rows.append({"Month":month_label(p),"HC":len(x),"PO (L)":money_lakh(x["p_o_value"].sum()),"Margin (L)":money_lakh(x["margin"].sum())})
    mom=pd.DataFrame(mom_rows)
    c1,c2=st.columns(2)
    with c1:
        st.plotly_chart(bar_chart(mom,"Month","HC","Month-on-Month Onboarding HC — Last 9 Months",height=390),use_container_width=True)
    with c2:
        fig=go.Figure()
        fig.add_trace(go.Bar(x=mom["Month"],y=mom["PO (L)"],name="PO (L)"))
        fig.add_trace(go.Scatter(x=mom["Month"],y=mom["Margin (L)"],name="Margin (L)",mode="lines+markers",yaxis="y2"))
        fig.update_layout(title="Onboarding PO vs Margin",height=390,margin=dict(l=10,r=10,t=45,b=45),paper_bgcolor="#f7f9fc",plot_bgcolor="#f7f9fc",yaxis=dict(title="PO (L)"),yaxis2=dict(title="Margin (L)",overlaying="y",side="right"))
        st.plotly_chart(fig,use_container_width=True)

    c1,c2=st.columns(2)
    with c1:
        dd=obm.groupby("display_date").size().reset_index(name="Onboarding")
        st.plotly_chart(line_chart(dd,"display_date","Onboarding","DOD Onboarding Trend",height=390),use_container_width=True)
    with c2:
        cd=obm.groupby("company_name").size().reset_index(name="Onboarding").sort_values("Onboarding",ascending=False).head(20)
        st.plotly_chart(bar_chart(cd,"company_name","Onboarding","Top 20 Clients by MTD Onboarding",height=390),use_container_width=True)

    c1,c2=st.columns(2)
    with c1:
        bhd=obm.groupby("BH").agg(HC=("full_name","size"),PO=("p_o_value","sum"),Margin=("margin","sum")).reset_index()
        bhd["PO (L)"]=bhd["PO"]/100000; bhd["Margin (L)"]=bhd["Margin"]/100000
        st.plotly_chart(bar_chart(bhd,"BH","HC","BH-wise MTD Onboarding",height=390),use_container_width=True)
    with c2:
        dom=obm.groupby("Domain").size().reset_index(name="HC")
        st.plotly_chart(doughnut(dom["HC"].tolist(),dom["Domain"].tolist(),"Onboarding by Domain"),use_container_width=True)

    st.markdown('<div class="section-title">Onboarding Client Detail — HC / PO / Margin</div>',unsafe_allow_html=True)
    cd=obm.groupby(["company_name","BH","KAM","Domain"]).agg(HC=("full_name","size"),PO=("p_o_value","sum"),Margin=("margin","sum")).reset_index()
    cd["PO (L)"]=cd["PO"]/100000; cd["Margin (L)"]=cd["Margin"]/100000; cd["Margin %"]=np.where(cd["PO"]!=0,cd["Margin"]/cd["PO"]*100,0)
    st.dataframe(cd[["company_name","BH","KAM","Domain","HC","PO (L)","Margin (L)","Margin %"]].sort_values("HC",ascending=False).style.format({"PO (L)":"{:,.1f}","Margin (L)":"{:,.1f}","Margin %":"{:,.1f}%"}),use_container_width=True,height=400,hide_index=True)

# ---- Exit ----
with tabs[5]:
    ex = filter_df(data["Exit"],"Exit",filters)
    exm = mtd(ex,"last_work_day",selected_period,as_of)
    exp = pipeline(filter_df(data["Exit Pipeline"],"Exit Pipeline",filters),"tentative_exit_date",selected_period,as_of)
    po=exm["p_o_value"].sum(); mar=exm["margin"].sum(); proj=len(exm)+len(exp)
    c=st.columns(6)
    with c[0]: metric_card("MTD Exit HC",fmt_int(len(exm)),"Exited","red")
    with c[1]: metric_card("MTD Exit PO",fmt_lakh(po),"₹ Lakhs","red")
    with c[2]: metric_card("MTD Exit Margin",fmt_lakh(mar),"₹ Lakhs","red")
    with c[3]: metric_card("Exit Margin %",f"{pct(mar,po):.1f}%","Margin / PO","gold")
    with c[4]: metric_card("Exit Pipeline",fmt_int(len(exp)),"After DOD in month","dark")
    with c[5]: metric_card("Exit Projection",fmt_int(proj),"MTD + Pipeline","red")

    start9=selected_period-8; months=pd.period_range(start9,selected_period,freq="M")
    rows=[]
    for p in months:
        x=filter_df(data["Exit"],"Exit",filters); x=x[x["last_work_day"].dt.to_period("M")==p]
        rows.append({"Month":month_label(p),"HC":len(x),"PO (L)":money_lakh(x["p_o_value"].sum()),"Margin (L)":money_lakh(x["margin"].sum())})
    mom=pd.DataFrame(rows)
    c1,c2=st.columns(2)
    with c1: st.plotly_chart(bar_chart(mom,"Month","HC","Month-on-Month Exit HC — Last 9 Months",height=390),use_container_width=True)
    with c2:
        fig=go.Figure(); fig.add_trace(go.Bar(x=mom["Month"],y=mom["PO (L)"],name="PO (L)")); fig.add_trace(go.Scatter(x=mom["Month"],y=mom["Margin (L)"],name="Margin (L)",mode="lines+markers",yaxis="y2"))
        fig.update_layout(title="Exit PO vs Margin",height=390,margin=dict(l=10,r=10,t=45,b=45),paper_bgcolor="#f7f9fc",plot_bgcolor="#f7f9fc",yaxis=dict(title="PO (L)"),yaxis2=dict(title="Margin (L)",overlaying="y",side="right"))
        st.plotly_chart(fig,use_container_width=True)
    c1,c2=st.columns(2)
    with c1:
        dd=exm.groupby("last_work_day").size().reset_index(name="Exit")
        st.plotly_chart(line_chart(dd,"last_work_day","Exit","DOD Exit Trend",height=390),use_container_width=True)
    with c2:
        cd=exm.groupby("company_name").size().reset_index(name="Exit").sort_values("Exit",ascending=False).head(20)
        st.plotly_chart(bar_chart(cd,"company_name","Exit","Top 20 Clients by MTD Exit",height=390),use_container_width=True)
    st.markdown('<div class="section-title">Exit Client Detail — HC / PO / Margin</div>',unsafe_allow_html=True)
    cd=exm.groupby(["company_name","BH","KAM","Domain"]).agg(HC=("full_name","size"),PO=("p_o_value","sum"),Margin=("margin","sum")).reset_index()
    cd["PO (L)"]=cd["PO"]/100000; cd["Margin (L)"]=cd["Margin"]/100000; cd["Margin %"]=np.where(cd["PO"]!=0,cd["Margin"]/cd["PO"]*100,0)
    st.dataframe(cd[["company_name","BH","KAM","Domain","HC","PO (L)","Margin (L)","Margin %"]].sort_values("HC",ascending=False).style.format({"PO (L)":"{:,.1f}","Margin (L)":"{:,.1f}","Margin %":"{:,.1f}%"}),use_container_width=True,height=400,hide_index=True)

# ---- Net & Projection ----
with tabs[6]:
    st.markdown('<div class="section-title">Net Headcount & Projection Bridge</div>',unsafe_allow_html=True)
    c=st.columns(6)
    for col,label,value,sub,kind in [
        (c[0],"OB MTD",fmt_int(m["ob_hc"]),"Actual","green"),(c[1],"OB Pipeline",fmt_int(m["ob_pipe_hc"]),"After DOD","green"),(c[2],"OB Projection",fmt_int(ob_proj),"MTD + Pipeline","green"),
        (c[3],"Exit MTD",fmt_int(m["exit_hc"]),"Actual","red"),(c[4],"Exit Pipeline",fmt_int(m["exit_pipe_hc"]),"After DOD","red"),(c[5],"Net Projection",fmt_int(net_proj),"Projection Net","dark")]:
        with col: metric_card(label,value,sub,kind)

    bridge=pd.DataFrame({"Stage":["Opening Baseline","+ Onboarding MTD","+ OB Pipeline","− Exit MTD","− Exit Pipeline","Projected Net"],"HC":[0,m["ob_hc"],m["ob_pipe_hc"],-m["exit_hc"],-m["exit_pipe_hc"],net_proj]})
    # waterfall
    fig=go.Figure(go.Waterfall(x=bridge["Stage"],y=bridge["HC"],measure=["absolute","relative","relative","relative","relative","total"],text=bridge["HC"].map(lambda x:f"{x:+,}"),textposition="outside"))
    fig.update_layout(title="Projection Bridge",height=420,margin=dict(l=10,r=10,t=45,b=45),paper_bgcolor="#f7f9fc",plot_bgcolor="#f7f9fc")
    st.plotly_chart(fig,use_container_width=True)

    score=bh.copy()
    score["OB vs Exit MTD"]=score["OB MTD"]-score["Exit MTD"]
    score["OB vs Exit Projection"]=score["OB Projection"]-score["Exit Projection"]
    st.markdown('<div class="section-title">BH Net & Projection Scorecard</div>',unsafe_allow_html=True)
    st.dataframe(score[["BH","OB MTD","OB Pipeline","OB Projection","Exit MTD","Exit Pipeline","Exit Projection","MTD Net","Net Projection","OB PO (L)","OB Margin (L)","Exit PO (L)","Exit Margin (L)"]].sort_values("Net Projection",ascending=False).style.format({c:"{:,.1f}" for c in ["OB PO (L)","OB Margin (L)","Exit PO (L)","Exit Margin (L)"]}),use_container_width=True,height=450,hide_index=True)

# =========================================================
# FOOTER / REFRESH RULE
# =========================================================
st.markdown("---")
st.caption("CEO Business Performance Cockpit • Demand uses Created_at • Submission uses date • Interview uses Interview_date • Selection uses selection_date • Onboarding uses display_date • Exit uses last_work_day • Pipelines use dates after DOD within the selected month • PO/Margin are ₹ Lakhs.")
