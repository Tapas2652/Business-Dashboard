
import os
from io import BytesIO
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="CEO | Business Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# Configuration
# -----------------------------
REPO_FILE = Path(__file__).with_name("CEO_MongoDB.xlsx")
SHEET_MAP = {
    "Demand": {"date": "Created_at", "bh": "BH", "client": "Company_name"},
    "Submission": {"date": "date", "bh": "BH", "client": "client"},
    "Interview": {"date": "Interview_date", "bh": "BH", "client": "company_name"},
    "Selection": {"date": "selection_date", "bh": "BH", "client": "company_name"},
    "Onboarding": {"date": "display_date", "bh": "BH", "client": "company_name"},
    "Onboarding Pipeline": {"date": "display_date", "bh": "BH", "client": "company_name"},
    "Exit": {"date": "last_work_day", "bh": "BH", "client": "company_name"},
    "Exit Pipeline": {"date": "tentative_exit_date", "bh": "BH", "client": "company_name"},
    "Active Headcount": {"date": "display_date", "bh": "BH", "client": "company_name"},
}

REQUIRED = {
    "Demand": ["no_of_opening", "Created_at", "BH", "KAM", "Company_name"],
    "Submission": ["date", "BH", "KAM", "client"],
    "Interview": ["Interview_date", "BH", "KAM", "company_name"],
    "Selection": ["selection_date", "BH", "KAM", "company_name", "po", "margin"],
    "Onboarding": ["display_date", "BH", "KAM", "company_name", "p_o_value", "margin"],
    "Onboarding Pipeline": ["display_date", "BH", "KAM", "company_name", "p_o_value", "margin"],
    "Exit": ["last_work_day", "BH", "KAM", "company_name", "p_o_value", "margin"],
    "Exit Pipeline": ["tentative_exit_date", "BH", "KAM", "company_name", "p_o_value", "margin"],
    "Active Headcount": ["display_date", "BH", "KAM", "company_name", "p_o_value", "margin"],
}

# -----------------------------
# Styling
# -----------------------------
st.markdown("""
<style>
.block-container {padding-top: 1.0rem; padding-bottom: 2rem;}
h1, h2, h3 {letter-spacing: -0.02em;}
.metric-card {
    background: linear-gradient(135deg, #111827 0%, #1f2937 100%);
    border: 1px solid #374151;
    border-radius: 14px;
    padding: 14px 16px;
    min-height: 112px;
}
.metric-label {font-size: 0.80rem; color: #9ca3af; margin-bottom: 5px;}
.metric-value {font-size: 1.65rem; font-weight: 750; color: white;}
.metric-sub {font-size: 0.74rem; color: #d1d5db; margin-top: 3px;}
.section-note {color:#6b7280; font-size:0.85rem;}
.small-muted {color:#6b7280; font-size:0.75rem;}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Data loading
# -----------------------------
@st.cache_data(show_spinner=False)
def load_workbook(source_bytes=None, source_path=None):
    if source_bytes is not None:
        raw = BytesIO(source_bytes)
    else:
        raw = source_path
    data = {}
    for sheet, cfg in SHEET_MAP.items():
        df = pd.read_excel(raw, sheet_name=sheet)
        missing = [c for c in REQUIRED[sheet] if c not in df.columns]
        if missing:
            raise ValueError(f"{sheet}: missing columns: {', '.join(missing)}")
        df = df.copy()
        df[cfg["date"]] = pd.to_datetime(df[cfg["date"]], errors="coerce")
        if "BH" in df.columns:
            df["BH"] = df["BH"].fillna("Unmapped").astype(str).str.strip().replace({"": "Unmapped"})
        if "KAM" in df.columns:
            df["KAM"] = df["KAM"].fillna("Unmapped").astype(str).str.strip().replace({"": "Unmapped"})
        client_col = cfg["client"]
        if client_col in df.columns:
            df[client_col] = df[client_col].fillna("Unmapped").astype(str).str.strip().replace({"": "Unmapped"})
        for c in ["no_of_opening", "po", "margin", "p_o_value"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        data[sheet] = df
    return data

def get_repo_bytes():
    if not REPO_FILE.exists():
        return None
    return REPO_FILE.read_bytes()

# -----------------------------
# Helpers
# -----------------------------
def month_label(ts):
    return pd.Timestamp(ts).strftime("%b'%y")

def month_start(month_period):
    return pd.Timestamp(month_period.start_time.date())

def month_end(month_period):
    return pd.Timestamp(month_period.end_time.date())

def filter_dates(df, date_col, start, end):
    d = df[date_col]
    return df[(d >= start) & (d <= end)].copy()

def in_month_to_date(df, date_col, period, as_of):
    start = month_start(period)
    end = min(month_end(period), pd.Timestamp(as_of).normalize())
    return filter_dates(df, date_col, start, end)

def in_month_pipeline(df, date_col, period):
    # Pipeline is a separate pipeline population for the selected month.
    # It is not restricted to dates after the DOD because the source pipeline
    # sheets contain pipeline records dated earlier in the month as well.
    return filter_dates(df, date_col, month_start(period), month_end(period))

def fmt_num(v):
    return f"{v:,.0f}"

def fmt_lakh(v):
    return f"{v:,.1f}"

def metric_card(label, value, sub=""):
    st.markdown(
        f"""<div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{sub}</div>
        </div>""",
        unsafe_allow_html=True,
    )

def agg_bh(data, period, as_of, bh_filter=None, kam_filter=None, client_filter=None):
    # BH universe comes from all relevant tabs.
    universe = set()
    for s, cfg in SHEET_MAP.items():
        df = data[s]
        universe.update(df["BH"].dropna().astype(str).unique())
    universe = sorted(universe)

    rows = []
    for bh in universe:
        if bh_filter and bh not in bh_filter:
            continue

        def apply_extra(df, cfg):
            x = df[df["BH"] == bh].copy()
            if kam_filter and "KAM" in x.columns:
                x = x[x["KAM"].isin(kam_filter)]
            if client_filter:
                x = x[x[cfg["client"]].isin(client_filter)]
            return x

        demand = apply_extra(data["Demand"], SHEET_MAP["Demand"])
        sub = apply_extra(data["Submission"], SHEET_MAP["Submission"])
        intr = apply_extra(data["Interview"], SHEET_MAP["Interview"])
        sel = apply_extra(data["Selection"], SHEET_MAP["Selection"])
        ob = apply_extra(data["Onboarding"], SHEET_MAP["Onboarding"])
        obp = apply_extra(data["Onboarding Pipeline"], SHEET_MAP["Onboarding Pipeline"])
        ex = apply_extra(data["Exit"], SHEET_MAP["Exit"])
        exp = apply_extra(data["Exit Pipeline"], SHEET_MAP["Exit Pipeline"])
        ah = apply_extra(data["Active Headcount"], SHEET_MAP["Active Headcount"])

        d_mtd = in_month_to_date(demand, "Created_at", period, as_of)
        s_mtd = in_month_to_date(sub, "date", period, as_of)
        i_mtd = in_month_to_date(intr, "Interview_date", period, as_of)
        se_mtd = in_month_to_date(sel, "selection_date", period, as_of)
        o_mtd = in_month_to_date(ob, "display_date", period, as_of)
        o_pipe = in_month_pipeline(obp, "display_date", period)
        e_mtd = in_month_to_date(ex, "last_work_day", period, as_of)
        e_pipe = in_month_pipeline(exp, "tentative_exit_date", period)

        active = ah[ah["display_date"] <= pd.Timestamp(as_of).normalize()]
        # Use the latest available Active Headcount snapshot up to as_of for the BH.
        if not active.empty:
            latest = active["display_date"].max()
            active = active[active["display_date"] == latest]

        ob_po = o_mtd["p_o_value"].sum()
        ob_margin = o_mtd["margin"].sum()
        obp_po = o_pipe["p_o_value"].sum()
        obp_margin = o_pipe["margin"].sum()
        ex_po = e_mtd["p_o_value"].sum()
        ex_margin = e_mtd["margin"].sum()
        exp_po = e_pipe["p_o_value"].sum()
        exp_margin = e_pipe["margin"].sum()

        rows.append({
            "BH": bh,
            "Demand": d_mtd["no_of_opening"].sum(),
            "Submissions": len(s_mtd),
            "Interview": len(i_mtd),
            "Selections": len(se_mtd),
            "Onboarding": len(o_mtd),
            "OB PO (L)": ob_po / 100000,
            "OB Margin (L)": ob_margin / 100000,
            "OB Pipeline": len(o_pipe),
            "OB Pipeline PO (L)": obp_po / 100000,
            "OB Pipeline Margin (L)": obp_margin / 100000,
            "OB Projection": len(o_mtd) + len(o_pipe),
            "OB Projection PO (L)": (ob_po + obp_po) / 100000,
            "OB Projection Margin (L)": (ob_margin + obp_margin) / 100000,
            "Exited": len(e_mtd),
            "Exit PO (L)": ex_po / 100000,
            "Exit Margin (L)": ex_margin / 100000,
            "Exit Pipeline": len(e_pipe),
            "Exit Pipeline PO (L)": exp_po / 100000,
            "Exit Pipeline Margin (L)": exp_margin / 100000,
            "Exit Projection": len(e_mtd) + len(e_pipe),
            "Exit Projection PO (L)": (ex_po + exp_po) / 100000,
            "Exit Projection Margin (L)": (ex_margin + exp_margin) / 100000,
            "MTD Net": len(o_mtd) - len(e_mtd),
            "Net Projection": len(o_mtd) + len(o_pipe) - len(e_mtd) - len(e_pipe),
            "Active HC": len(active),
            "Active PO (L)": active["p_o_value"].sum() / 100000,
            "Active Margin (L)": active["margin"].sum() / 100000,
        })

    return pd.DataFrame(rows)

def client_detail(data, period, as_of, selected_bhs=None, selected_kams=None):
    # Client list from the active population and MTD/pipeline tabs.
    source_cols = [
        ("Active Headcount", "company_name"),
        ("Onboarding", "company_name"),
        ("Onboarding Pipeline", "company_name"),
        ("Exit", "company_name"),
        ("Exit Pipeline", "company_name"),
        ("Demand", "Company_name"),
        ("Submission", "client"),
        ("Interview", "company_name"),
        ("Selection", "company_name"),
    ]
    clients = set()
    for s, c in source_cols:
        df = data[s]
        if selected_bhs:
            df = df[df["BH"].isin(selected_bhs)]
        if selected_kams:
            df = df[df["KAM"].isin(selected_kams)]
        clients.update(df[c].dropna().astype(str).unique())

    rows = []
    for client in sorted(clients):
        def x(s):
            cfg = SHEET_MAP[s]
            df = data[s]
            if selected_bhs:
                df = df[df["BH"].isin(selected_bhs)]
            if selected_kams:
                df = df[df["KAM"].isin(selected_kams)]
            return df[df[cfg["client"]] == client].copy()

        ah = x("Active Headcount")
        ob = x("Onboarding")
        obp = x("Onboarding Pipeline")
        ex = x("Exit")
        exp = x("Exit Pipeline")
        dem = x("Demand")
        sub = x("Submission")
        intr = x("Interview")
        sel = x("Selection")

        ah = ah[ah["display_date"] <= pd.Timestamp(as_of).normalize()]
        if not ah.empty:
            ah = ah[ah["display_date"] == ah["display_date"].max()]

        dem = in_month_to_date(dem, "Created_at", period, as_of)
        sub = in_month_to_date(sub, "date", period, as_of)
        intr = in_month_to_date(intr, "Interview_date", period, as_of)
        sel = in_month_to_date(sel, "selection_date", period, as_of)
        ob = in_month_to_date(ob, "display_date", period, as_of)
        obp = in_month_pipeline(obp, "display_date", period)
        ex = in_month_to_date(ex, "last_work_day", period, as_of)
        exp = in_month_pipeline(exp, "tentative_exit_date", period)

        rows.append({
            "Client": client,
            "BH": (ah["BH"].mode().iat[0] if not ah.empty else
                   (ob["BH"].mode().iat[0] if not ob.empty else "")),
            "KAM": (ah["KAM"].mode().iat[0] if not ah.empty else
                    (ob["KAM"].mode().iat[0] if not ob.empty else "")),
            "Active HC": len(ah),
            "Active PO (L)": ah["p_o_value"].sum()/100000,
            "Active Margin (L)": ah["margin"].sum()/100000,
            "Demand": dem["no_of_opening"].sum(),
            "Submissions": len(sub),
            "Interview": len(intr),
            "Selections": len(sel),
            "Onboarding": len(ob),
            "OB PO (L)": ob["p_o_value"].sum()/100000,
            "OB Margin (L)": ob["margin"].sum()/100000,
            "OB Pipeline": len(obp),
            "OB Pipeline PO (L)": obp["p_o_value"].sum()/100000,
            "OB Pipeline Margin (L)": obp["margin"].sum()/100000,
            "OB Projection": len(ob)+len(obp),
            "Exited": len(ex),
            "Exit PO (L)": ex["p_o_value"].sum()/100000,
            "Exit Margin (L)": ex["margin"].sum()/100000,
            "Exit Pipeline": len(exp),
            "Exit Pipeline PO (L)": exp["p_o_value"].sum()/100000,
            "Exit Pipeline Margin (L)": exp["margin"].sum()/100000,
            "Exit Projection": len(ex)+len(exp),
            "MTD Net": len(ob)-len(ex),
            "Net Projection": len(ob)+len(obp)-len(ex)-len(exp),
        })

    return pd.DataFrame(rows)

# -----------------------------
# Load source
# -----------------------------
repo_bytes = get_repo_bytes()

with st.sidebar:
    st.title("CEO Dashboard")
    st.caption("Business Head / Client performance")

    uploaded = st.file_uploader(
        "Optional: test with an updated Excel file",
        type=["xlsx"],
        help="For production, keep CEO_MongoDB.xlsx in the same GitHub repository as app.py.",
    )

    if uploaded is not None:
        source_bytes = uploaded.getvalue()
        source_name = uploaded.name
    elif repo_bytes is not None:
        source_bytes = None
        source_name = "CEO_MongoDB.xlsx (repository)"
    else:
        source_bytes = None
        source_name = None

    if st.button("↻ Refresh data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

if source_name is None:
    st.error("CEO_MongoDB.xlsx was not found beside app.py. Add the workbook to the GitHub repository.")
    st.stop()

try:
    if uploaded is not None:
        data = load_workbook(source_bytes=source_bytes)
    else:
        data = load_workbook(source_path=str(REPO_FILE))
except Exception as e:
    st.error(f"Could not load the workbook: {e}")
    st.stop()

# -----------------------------
# Determine periods / default as-of
# -----------------------------
all_dates = []
for s, cfg in SHEET_MAP.items():
    d = data[s][cfg["date"]].dropna()
    if not d.empty:
        all_dates.append(d.max())
latest_available = max(all_dates).normalize()

period_values = set()
for s, cfg in SHEET_MAP.items():
    d = data[s][cfg["date"]].dropna()
    period_values.update(d.dt.to_period("M").tolist())

periods = sorted(period_values)
period_labels = [month_label(p.start_time) for p in periods]
default_idx = period_labels.index(month_label(latest_available)) if month_label(latest_available) in period_labels else len(period_labels)-1

with st.sidebar:
    selected_label = st.selectbox("Reporting Month", period_labels, index=default_idx)
    selected_period = pd.Period(pd.to_datetime("01-"+selected_label.replace("'", "-"), format="%d-%b-%y"), freq="M")

    # Default to latest usable date for selected month, but never beyond today.
    dates_for_period = []
    for s, cfg in SHEET_MAP.items():
        d = data[s][cfg["date"]].dropna()
        d = d[d.dt.to_period("M") == selected_period]
        if not d.empty:
            dates_for_period.append(d.max())
    default_asof = min(max(dates_for_period), pd.Timestamp.today().normalize()) if dates_for_period else pd.Timestamp.today().normalize()
    as_of = st.date_input("DOD / As-of Date", value=default_asof.date(), min_value=selected_period.start_time.date(), max_value=min(selected_period.end_time.date(), pd.Timestamp.today().date()))
    as_of = pd.Timestamp(as_of)

    # BH filter
    bh_values = sorted(set().union(*[set(data[s]["BH"].dropna().astype(str).unique()) for s in SHEET_MAP]))
    selected_bhs = st.multiselect("Business Head", bh_values, default=[])

    # KAM filter
    kam_values = sorted(set().union(*[set(data[s]["KAM"].dropna().astype(str).unique()) for s in SHEET_MAP]))
    selected_kams = st.multiselect("KAM", kam_values, default=[])

# -----------------------------
# Main title
# -----------------------------
st.title("CEO Business Performance Dashboard")
st.markdown(
    f"**Reporting Month:** {selected_label} &nbsp; | &nbsp; "
    f"**As of:** {as_of.strftime('%d %b %Y')} &nbsp; | &nbsp; "
    f"**Data Source:** {source_name}"
)
st.caption("PO and Margin are displayed in ₹ Lakhs. Projection = MTD + future pipeline within the selected month. Net = Onboarding − Exit.")

bh_df = agg_bh(
    data,
    selected_period,
    as_of,
    bh_filter=selected_bhs or None,
    kam_filter=selected_kams or None,
)

# -----------------------------
# KPI cards
# -----------------------------
tot = bh_df.sum(numeric_only=True)
c = st.columns(6)
with c[0]: metric_card("Demand", fmt_num(tot["Demand"]), "MTD openings")
with c[1]: metric_card("Submissions", fmt_num(tot["Submissions"]), "MTD")
with c[2]: metric_card("Interview", fmt_num(tot["Interview"]), "MTD")
with c[3]: metric_card("Selection", fmt_num(tot["Selections"]), "MTD")
with c[4]: metric_card("Onboarding", fmt_num(tot["Onboarding"]), f"MTD | {fmt_lakh(tot['OB PO (L)'])} L PO")
with c[5]: metric_card("Exited", fmt_num(tot["Exited"]), f"MTD | {fmt_lakh(tot['Exit PO (L)'])} L PO")

c2 = st.columns(6)
with c2[0]: metric_card("MTD Net", fmt_num(tot["MTD Net"]), "OB − Exit")
with c2[1]: metric_card("OB Pipeline", fmt_num(tot["OB Pipeline"]), f"{fmt_lakh(tot['OB Pipeline PO (L)'])} L PO")
with c2[2]: metric_card("OB Projection", fmt_num(tot["OB Projection"]), "MTD + Pipeline")
with c2[3]: metric_card("Exit Pipeline", fmt_num(tot["Exit Pipeline"]), f"{fmt_lakh(tot['Exit Pipeline PO (L)'])} L PO")
with c2[4]: metric_card("Exit Projection", fmt_num(tot["Exit Projection"]), "MTD + Pipeline")
with c2[5]: metric_card("Net Projection", fmt_num(tot["Net Projection"]), "OB Projection − Exit Projection")

# -----------------------------
# Tabs
# -----------------------------
tab1, tab2, tab3, tab4 = st.tabs(["BH CEO View", "Client View", "DOD Trend", "Data Quality"])

with tab1:
    st.subheader("Business Head Scorecard")
    st.caption("All counts are MTD to the selected as-of date. Pipeline is future-dated within the selected reporting month.")

    display_cols = [
        "BH", "Active HC", "Active PO (L)", "Active Margin (L)",
        "Demand", "Submissions", "Interview", "Selections",
        "Onboarding", "OB PO (L)", "OB Margin (L)",
        "OB Pipeline", "OB Pipeline PO (L)", "OB Projection",
        "Exited", "Exit PO (L)", "Exit Pipeline", "Exit Projection",
        "MTD Net", "Net Projection",
    ]
    table = bh_df[display_cols].copy()
    table = table.sort_values(["Net Projection", "Onboarding"], ascending=[False, False])

    st.dataframe(
        table.style.format({
            c: "{:,.1f}" for c in table.columns if "(L)" in c
        }).format({
            c: "{:,.0f}" for c in table.columns if c not in ["BH"] and "(L)" not in c
        }),
        use_container_width=True,
        height=560,
    )

    st.subheader("BH Onboarding vs Exit")
    chart_df = bh_df[["BH", "Onboarding", "OB Pipeline", "Exited", "Exit Pipeline"]].melt(
        id_vars="BH", var_name="Metric", value_name="HC"
    )
    fig = px.bar(
        chart_df, x="BH", y="HC", color="Metric", barmode="group",
        text_auto=True, height=430,
    )
    fig.update_layout(margin=dict(l=10,r=10,t=20,b=80), legend_title="")
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("BH PO & Margin — ₹ Lakhs")
    pm = bh_df[["BH", "OB PO (L)", "OB Margin (L)", "OB Pipeline PO (L)", "OB Pipeline Margin (L)",
                "Exit PO (L)", "Exit Margin (L)", "Exit Pipeline PO (L)", "Exit Pipeline Margin (L)"]].copy()
    pm["OB Projection PO (L)"] = pm["OB PO (L)"] + pm["OB Pipeline PO (L)"]
    pm["Exit Projection PO (L)"] = pm["Exit PO (L)"] + pm["Exit Pipeline PO (L)"]
    pmm = pm[["BH", "OB Projection PO (L)", "Exit Projection PO (L)", "OB Margin (L)", "Exit Margin (L)"]].melt(
        id_vars="BH", var_name="Metric", value_name="₹ Lakhs"
    )
    fig2 = px.bar(pmm, x="BH", y="₹ Lakhs", color="Metric", barmode="group", height=430)
    fig2.update_layout(margin=dict(l=10,r=10,t=20,b=80), legend_title="")
    st.plotly_chart(fig2, use_container_width=True)

with tab2:
    st.subheader("Client-Level Drilldown")
    client_df = client_detail(data, selected_period, as_of, selected_bhs or None, selected_kams or None)
    search = st.text_input("Search client", "")
    if search:
        client_df = client_df[client_df["Client"].str.contains(search, case=False, na=False)]

    cols = [
        "Client","BH","KAM","Active HC","Active PO (L)","Active Margin (L)",
        "Demand","Submissions","Interview","Selections","Onboarding","OB PO (L)",
        "OB Pipeline","OB Pipeline PO (L)","OB Projection","Exited","Exit PO (L)",
        "Exit Pipeline","Exit Pipeline PO (L)","Exit Projection","MTD Net","Net Projection"
    ]
    st.dataframe(
        client_df[cols].sort_values("Net Projection", ascending=False).style.format(
            {c:"{:,.1f}" for c in cols if "(L)" in c}
        ).format(
            {c:"{:,.0f}" for c in cols if c not in ["Client","BH","KAM"] and "(L)" not in c}
        ),
        use_container_width=True,
        height=620,
    )

with tab3:
    st.subheader("Daily Operating Trend")
    start = month_start(selected_period)
    end = min(month_end(selected_period), as_of)
    dates = pd.date_range(start, end, freq="D")
    daily = pd.DataFrame({"Date": dates})

    def daily_count(sheet, date_col, value_col=None):
        df = data[sheet].copy()
        df = df[(df["BH"].isin(selected_bhs))] if selected_bhs else df
        df = df[(df["KAM"].isin(selected_kams))] if selected_kams else df
        df = df[(df[date_col] >= start) & (df[date_col] <= end)]
        if value_col:
            return df.groupby(date_col)[value_col].sum()
        return df.groupby(date_col).size()

    for name, sheet, date_col in [
        ("Demand","Demand","Created_at"),
        ("Submissions","Submission","date"),
        ("Interview","Interview","Interview_date"),
        ("Selection","Selection","selection_date"),
        ("Onboarding","Onboarding","display_date"),
        ("Exit","Exit","last_work_day"),
    ]:
        ser = daily_count(sheet, date_col).rename(name)
        daily = daily.merge(ser, left_on="Date", right_index=True, how="left")
    daily = daily.fillna(0)
    daily["MTD Net"] = daily["Onboarding"].cumsum() - daily["Exit"].cumsum()

    trend_metrics = st.multiselect(
        "Trend metrics",
        ["Demand","Submissions","Interview","Selection","Onboarding","Exit","MTD Net"],
        default=["Demand","Onboarding","Exit","MTD Net"]
    )
    if trend_metrics:
        plot = daily[["Date"] + trend_metrics].melt(id_vars="Date", var_name="Metric", value_name="HC")
        fig = px.line(plot, x="Date", y="HC", color="Metric", markers=True, height=470)
        fig.update_layout(margin=dict(l=10,r=10,t=20,b=30), legend_title="")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Daily Data Table")
    st.dataframe(daily, use_container_width=True, height=400)

with tab4:
    st.subheader("Data Quality & Refresh Diagnostics")
    quality_rows = []
    for s, cfg in SHEET_MAP.items():
        df = data[s]
        d = df[cfg["date"]]
        quality_rows.append({
            "Sheet": s,
            "Rows": len(df),
            "Date Column": cfg["date"],
            "Valid Dates": int(d.notna().sum()),
            "Blank Dates": int(d.isna().sum()),
            "Latest Date": d.max(),
            "Blank BH": int(df["BH"].isna().sum()),
            "Blank KAM": int(df["KAM"].isna().sum()),
        })
    q = pd.DataFrame(quality_rows)
    st.dataframe(q, use_container_width=True, hide_index=True)

    st.info(
        "Production refresh: replace/commit CEO_MongoDB.xlsx in the same GitHub repository. "
        "Streamlit Cloud will redeploy the repository; use the Refresh data button to clear the app cache."
    )

    st.subheader("Business Rules Applied")
    st.markdown("""
    - **Demand:** sum `no_of_opening`, using `Created_at`.
    - **Submission:** row count, using `date`.
    - **Interview:** row count, using `Interview_date`.
    - **Selection:** row count, using `selection_date`.
    - **Onboarding:** row count + PO/Margin, using `display_date`.
    - **Onboarding Pipeline:** future rows in the selected month after the as-of date, using `display_date`.
    - **Exit:** row count + PO/Margin, using `last_work_day`.
    - **Exit Pipeline:** future rows in the selected month after the as-of date, using `tentative_exit_date`.
    - **MTD Net:** Onboarding − Exit.
    - **Projection:** MTD + future pipeline.
    - **Net Projection:** Onboarding Projection − Exit Projection.
    - **Active Headcount:** latest available snapshot on/before the as-of date; PO/Margin shown in ₹ Lakhs.
    """)

# Footer
st.caption("CEO Performance Dashboard • Built from CEO_MongoDB.xlsx • Counts are headcount/opening events unless stated otherwise.")
