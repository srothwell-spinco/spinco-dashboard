import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import date
import os
import glob

# =============================
# CONFIG
# =============================
MODEL_PATH = "out/model.csv"
DATA_DIR = "data/incoming"
MIN_CLASSES = 4
PASSWORD = "spinco2025"

ACCENT = "#BBD7ED"
BLACK = "#000000"
WHITE = "#FFFFFF"
GREY = "#4D4D4D"
LIGHT = "#F4F4F4"

UTIL_SCALE = [
    [0.0, "#2471A3"],
    [0.3, "#A8CCE8"],
    [0.5, "#F0F0F0"],
    [0.7, "#E8A090"],
    [1.0, "#C0392B"],
]
DROPOUT_SCALE = [[0.0, "#FFFFFF"], [1.0, "#C0392B"]]
MIX_SCALE = [[0.0, "#7B2D8B"], [0.5, "#F0F0F0"], [1.0, "#27AE60"]]
CREDIT_SCALE = [[0.0, "#27AE60"], [0.5, "#F0F0F0"], [1.0, "#7B2D8B"]]
ORDERS_HEATMAP_SCALE = [
    [0.0,  "#2471A3"],
    [0.25, "#A8CCE8"],
    [0.5,  "#F0F0F0"],
    [0.75, "#E8A090"],
    [1.0,  "#C0392B"],
]

DOW_ORDER  = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
DOW_LABELS = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
SLOT_ORDER = ["06:00","07:00","08:00","09:00","09:30","10:15","11:30",
              "16:45","17:00","17:30","18:35","19:40"]

SCHOOL_YEAR_MONTHS = ["09","10","11","12","01","02","03","04"]
SUMMER_MONTHS      = ["05","06","07","08"]

GROUPS = ["Credits", "Memberships", "Intro Offers"]
GROUP_COLORS = {"Credits": ACCENT, "Memberships": GREY, "Intro Offers": BLACK}

PURCHASE_WINDOWS = [
    ("Early Morning", 5,  8),
    ("Mid Morning",   8,  11),
    ("Midday",        11, 14),
    ("Afternoon",     14, 17),
    ("Evening",       17, 22),
    ("Late Night",    22, 5),
]
WINDOW_ORDER = [w[0] for w in PURCHASE_WINDOWS]
WINDOW_COLORS = {
    "Early Morning": "#2471A3",
    "Mid Morning":   "#A8CCE8",
    "Midday":        "#BBD7ED",
    "Afternoon":     "#F0C080",
    "Evening":       "#E8A090",
    "Late Night":    "#4D4D4D",
}

def assign_window(hour):
    for name, start, end in PURCHASE_WINDOWS:
        if start < end:
            if start <= hour < end:
                return name
        else:
            if hour >= start or hour < end:
                return name
    return "Other"

st.set_page_config(
    page_title="SPINCO London — Analytics",
    page_icon="🚴",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main { background-color: #FFFFFF; }
    .block-container { padding-top:2rem; padding-bottom:1rem; padding-left:2rem; padding-right:2rem; }
    h1, h2, h3 { color: #000000; font-family: sans-serif; }
    .metric-card {
        background: #F4F4F4; border-radius: 6px; padding: 16px;
        text-align: center; border-left: 4px solid #BBD7ED; margin-bottom: 8px;
    }
    .metric-value { font-size: 26px; font-weight: bold; color: #000000; }
    .metric-label { font-size: 11px; color: #4D4D4D; margin-top: 4px; }
    .section-header { border-bottom: 2px solid #000000; padding-bottom: 4px; margin-bottom: 16px; margin-top: 8px; }
    .spinco-header {
        background: #000000; color: white; padding: 16px 24px;
        border-radius: 8px; margin-bottom: 20px; border-left: 6px solid #BBD7ED;
    }
    .description { font-size: 12px; color: #4D4D4D; margin-bottom: 12px; line-height: 1.5; }
    .stDataFrame td { text-align: center !important; }
    .stDataFrame th { text-align: center !important; }
</style>
""", unsafe_allow_html=True)

# =============================
# PASSWORD GATE
# =============================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div class="spinco-header">
        <span style="font-size:22px; font-weight:bold; color:white;">// SPINCO LONDON</span>
        <span style="font-size:13px; color:#BBD7ED; margin-left:12px;">Analytics Dashboard</span>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("### Access Required")
    pw = st.text_input("Password", type="password")
    if st.button("Login"):
        if pw == PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()

# =============================
# LOAD DATA
# =============================
@st.cache_data
def load_data():
    df = pd.read_csv(MODEL_PATH)
    df["month"] = df["month"].astype(str)
    df = df[df["report_instructor"] != "TEAM TEACH"].copy()
    df["instructor_first"] = df["report_instructor"].apply(lambda x: str(x).strip().split()[0])
    df["month_num"] = df["month"].str.slice(5, 7)
    return df

@st.cache_data
def load_revenue():
    files = glob.glob(os.path.join(DATA_DIR, "revenue_*.csv"))
    if not files:
        return pd.DataFrame()
    frames = []
    for f in files:
        month = os.path.basename(f).replace("revenue_","").replace(".csv","")
        try:
            df_r = pd.read_csv(f)
            df_r.columns = df_r.columns.str.strip()
            df_r["month"] = month
            frames.append(df_r)
        except:
            pass
    if not frames:
        return pd.DataFrame()
    rev = pd.concat(frames, ignore_index=True)
    rev["Realized Revenue"] = pd.to_numeric(
        rev["Realized Revenue"].astype(str).str.replace(",","").str.strip(), errors="coerce"
    ).fillna(0)
    return rev

@st.cache_data
def load_orders_data():
    purchases_path = "out/orders_purchases.csv"
    renewals_path  = "out/orders_renewals.csv"
    summary_path   = "out/orders_summary.csv"
    if not os.path.exists(purchases_path):
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    purchases = pd.read_csv(purchases_path, parse_dates=["order_dt"])
    renewals  = pd.read_csv(renewals_path,  parse_dates=["order_dt"]) if os.path.exists(renewals_path) else pd.DataFrame()
    summary   = pd.read_csv(summary_path)   if os.path.exists(summary_path)  else pd.DataFrame()
    if not purchases.empty and "hour" in purchases.columns:
        purchases["window"] = purchases["hour"].apply(assign_window)
    if not renewals.empty and "order_dt" in renewals.columns:
        renewals["month"] = pd.to_datetime(renewals["order_dt"]).dt.to_period("M").astype(str)
        renewals["hour"]  = pd.to_datetime(renewals["order_dt"]).dt.hour
        renewals["dow"]   = pd.to_datetime(renewals["order_dt"]).dt.day_name()
        if "Line Total" not in renewals.columns:
            renewals["Line Total"] = pd.to_numeric(renewals.get("Line Total", 0), errors="coerce").fillna(0)
    return purchases, renewals, summary

df        = load_data()
rev_df    = load_revenue()
ord_purchases, ord_renewals, ord_summary = load_orders_data()

available_months = sorted(df["month"].unique())
month_labels     = {m: pd.Period(m, freq="M").strftime("%B %Y") for m in available_months}
school_year      = [m for m in available_months if m.split("-")[1] in SCHOOL_YEAR_MONTHS]
summer           = [m for m in available_months if m.split("-")[1] in SUMMER_MONTHS]

def fmt_pct(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"+{round(v*100,1)}%" if v >= 0 else f"{round(v*100,1)}%"

def fmt_cad(v):
    return f"${v:,.0f}"

def desc(text):
    st.markdown(f'<div class="description">{text}</div>', unsafe_allow_html=True)

# =============================
# SIDEBAR
# =============================
with st.sidebar:
    st.markdown("### // SPINCO")
    st.markdown("**London Analytics**")
    st.markdown("---")

    period_options = []
    if school_year:
        period_options.append("School Year (Sep-Apr)")
    if summer:
        period_options.append("Summer (May-Aug)")
    period_options += available_months

    selected_raw = st.selectbox(
        "Report Period",
        options=period_options,
        index=len(period_options) - 1,
        format_func=lambda x: {
            "School Year (Sep-Apr)": "School Year — Sep to Apr",
            "Summer (May-Aug)": "Summer — May to Aug",
        }.get(x, month_labels.get(x, x))
    )

    if selected_raw == "School Year (Sep-Apr)":
        df_curr  = df[df["month"].isin(school_year)].copy()
        rev_curr = rev_df[rev_df["month"].isin(school_year)].copy() if not rev_df.empty else pd.DataFrame()
        ord_curr = ord_purchases[ord_purchases["month"].isin(school_year)].copy() if not ord_purchases.empty else pd.DataFrame()
        ren_curr = ord_renewals[ord_renewals["month"].isin(school_year)].copy() if not ord_renewals.empty else pd.DataFrame()
        sel_months     = school_year
        display_label  = "School Year (Sep-Apr)"
        selected_month = None
    elif selected_raw == "Summer (May-Aug)":
        df_curr  = df[df["month"].isin(summer)].copy()
        rev_curr = rev_df[rev_df["month"].isin(summer)].copy() if not rev_df.empty else pd.DataFrame()
        ord_curr = ord_purchases[ord_purchases["month"].isin(summer)].copy() if not ord_purchases.empty else pd.DataFrame()
        ren_curr = ord_renewals[ord_renewals["month"].isin(summer)].copy() if not ord_renewals.empty else pd.DataFrame()
        sel_months     = summer
        display_label  = "Summer (May-Aug)"
        selected_month = None
    else:
        selected_month = selected_raw
        df_curr  = df[df["month"] == selected_month].copy()
        rev_curr = rev_df[rev_df["month"] == selected_month].copy() if not rev_df.empty else pd.DataFrame()
        ord_curr = ord_purchases[ord_purchases["month"] == selected_month].copy() if not ord_purchases.empty else pd.DataFrame()
        ren_curr = ord_renewals[ord_renewals["month"] == selected_month].copy() if not ord_renewals.empty else pd.DataFrame()
        sel_months    = [selected_month]
        display_label = month_labels[selected_month]

    st.markdown("---")
    st.caption(f"Data: {month_labels[available_months[0]]} to {month_labels[available_months[-1]]}")
    st.markdown("---")
    if st.button("Logout"):
        st.session_state.authenticated = False
        st.rerun()
    st.caption("SPINCO London — Internal Use Only")

# =============================
# MOM / TRAILING
# =============================
idx        = available_months.index(selected_month) if selected_month else None
mom_change = None
trailing_3m = None
if idx is not None:
    curr_util = df_curr["util"].mean()
    if idx >= 1:
        mom_change  = curr_util - df[df["month"] == available_months[idx-1]]["util"].mean()
    if idx >= 3:
        trailing_3m = curr_util - df[df["month"] == available_months[idx-3]]["util"].mean()

# =============================
# REVENUE
# =============================
if not rev_curr.empty:
    rev_class      = rev_curr[rev_curr["Product Type"].isin(["Credit","Membership"])]["Realized Revenue"].sum()
    rev_penalty    = rev_curr[rev_curr["Product Type"] == "Penalty Fees"]["Realized Revenue"].sum()
    rev_credit     = rev_curr[rev_curr["Product Type"] == "Credit"]["Realized Revenue"].sum()
    rev_membership = rev_curr[rev_curr["Product Type"] == "Membership"]["Realized Revenue"].sum()
    total_riders_rev = int(df_curr["checked_in"].sum())
    rev_per_rider  = rev_class / total_riders_rev if total_riders_rev > 0 else 0
else:
    rev_class = rev_penalty = rev_credit = rev_membership = rev_per_rider = 0

# =============================
# HEADER
# =============================
st.markdown(f"""
<div class="spinco-header">
    <span style="font-size:22px; font-weight:bold; color:white;">// SPINCO LONDON</span>
    <span style="font-size:13px; color:#BBD7ED; margin-left:16px;">Performance Dashboard — {display_label}</span>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Overview", "🕐 Timeslots", "👤 Instructors", "📈 Trends", "🛒 Orders"
])

# =============================
# TAB 1 — OVERVIEW
# =============================
with tab1:
    st.markdown('<div class="section-header"><h3>Studio Snapshot</h3></div>', unsafe_allow_html=True)

    total_riders  = int(df_curr["checked_in"].sum())
    total_classes = len(df_curr)
    utilization   = df_curr["util"].mean()
    median_util   = df_curr["util"].median()
    pct_above_70  = (df_curr["util"] >= 0.70).mean()
    pct_below_40  = (df_curr["util"] < 0.40).mean()
    top_day       = df_curr.groupby("dow")["util"].mean().idxmax()
    slot_counts   = df_curr.groupby("slot_key")["util"].agg(["mean","count"])
    eligible      = slot_counts[slot_counts["count"] >= MIN_CLASSES]
    top_slot      = eligible["mean"].idxmax() if not eligible.empty else "N/A"

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{round(utilization*100,1)}%</div><div class="metric-label">Studio Utilization</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{total_riders:,}</div><div class="metric-label">Total Riders</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{total_classes:,}</div><div class="metric-label">Total Classes</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{round(median_util*100,1)}%</div><div class="metric-label">Median Utilization</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if rev_class > 0:
        r1, r2, r3, r4, r5 = st.columns(5)
        with r1:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_cad(rev_class)}</div><div class="metric-label">Class Revenue (CAD)</div></div>', unsafe_allow_html=True)
        with r2:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_cad(rev_membership)}</div><div class="metric-label">Membership Revenue</div></div>', unsafe_allow_html=True)
        with r3:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_cad(rev_credit)}</div><div class="metric-label">Credit Revenue</div></div>', unsafe_allow_html=True)
        with r4:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_cad(rev_per_rider)}</div><div class="metric-label">Revenue per Rider</div></div>', unsafe_allow_html=True)
        with r5:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_cad(rev_penalty)}</div><div class="metric-label">Penalty Fees</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    c5, c6, c7, c8, c9 = st.columns(5)
    with c5:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{round(pct_above_70*100,1)}%</div><div class="metric-label">Classes ≥70%</div></div>', unsafe_allow_html=True)
    with c6:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{round(pct_below_40*100,1)}%</div><div class="metric-label">Classes &lt;40%</div></div>', unsafe_allow_html=True)
    with c7:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_pct(mom_change)}</div><div class="metric-label">MoM Change</div></div>', unsafe_allow_html=True)
    with c8:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_pct(trailing_3m)}</div><div class="metric-label">Trailing 3M</div></div>', unsafe_allow_html=True)
    with c9:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{top_day}</div><div class="metric-label">Top Day</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    monthly_trend = (
        df.groupby("month")
        .agg(utilization=("util","mean"), riders=("checked_in","sum"))
        .reset_index().sort_values("month")
    )
    monthly_trend["label"]   = monthly_trend["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))
    monthly_trend["util_pct"] = (monthly_trend["utilization"] * 100).round(1)
    monthly_trend["current"] = monthly_trend["month"] == (selected_month or "")

    col_ov1, col_ov2 = st.columns(2)

    with col_ov1:
        fig_ov = go.Figure()
        fig_ov.add_trace(go.Scatter(
            x=monthly_trend["label"], y=monthly_trend["util_pct"],
            mode="lines+markers+text",
            line=dict(color=BLACK, width=2),
            marker=dict(
                color=[ACCENT if c else BLACK for c in monthly_trend["current"]],
                size=[14 if c else 7 for c in monthly_trend["current"]],
                line=dict(color=BLACK, width=1)
            ),
            text=[f"{v}%" for v in monthly_trend["util_pct"]],
            textposition="top center", textfont=dict(size=10),
        ))
        fig_ov.update_layout(
            title="Studio Utilization — Season to Date",
            height=300, margin=dict(l=20, r=20, t=40, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(ticksuffix="%", range=[0,100], gridcolor=LIGHT),
            xaxis=dict(gridcolor=WHITE), showlegend=False,
        )
        st.plotly_chart(fig_ov, use_container_width=True)

    with col_ov2:
        if not rev_df.empty:
            rev_monthly = (
                rev_df[rev_df["Product Type"].isin(["Credit","Membership"])]
                .groupby("month")["Realized Revenue"].sum()
                .reset_index().sort_values("month")
            )
            rev_monthly["label"]   = rev_monthly["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))
            rev_monthly["current"] = rev_monthly["month"] == (selected_month or "")
            fig_rev = go.Figure(go.Bar(
                x=rev_monthly["label"], y=rev_monthly["Realized Revenue"],
                marker_color=[BLACK if c else ACCENT for c in rev_monthly["current"]],
                text=[fmt_cad(v) for v in rev_monthly["Realized Revenue"]],
                textposition="outside",
            ))
            fig_rev.update_layout(
                title="Monthly Class Revenue (CAD)",
                height=300, margin=dict(l=20, r=20, t=40, b=20),
                paper_bgcolor=WHITE, plot_bgcolor=WHITE,
                yaxis=dict(gridcolor=LIGHT), xaxis_title="", showlegend=False,
            )
            st.plotly_chart(fig_rev, use_container_width=True)

    if not rev_df.empty:
        riders_by_month    = df.groupby("month")["checked_in"].sum().reset_index()
        riders_by_month.columns = ["month","total_riders"]
        rev_class_monthly  = (
            rev_df[rev_df["Product Type"].isin(["Credit","Membership"])]
            .groupby("month")["Realized Revenue"].sum().reset_index()
        )
        rpr = rev_class_monthly.merge(riders_by_month, on="month")
        rpr["rev_per_rider"] = (rpr["Realized Revenue"] / rpr["total_riders"]).round(2)
        rpr["label"]   = rpr["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))
        rpr["current"] = rpr["month"] == (selected_month or "")

        fig_rpr = go.Figure()
        fig_rpr.add_trace(go.Scatter(
            x=rpr["label"], y=rpr["rev_per_rider"],
            mode="lines+markers+text",
            line=dict(color=GREY, width=2),
            marker=dict(
                color=[ACCENT if c else GREY for c in rpr["current"]],
                size=[14 if c else 7 for c in rpr["current"]],
            ),
            text=[fmt_cad(v) for v in rpr["rev_per_rider"]],
            textposition="top center", textfont=dict(size=10),
        ))
        fig_rpr.update_layout(
            title="Revenue per Rider — Season to Date (CAD)",
            height=260, margin=dict(l=20, r=20, t=40, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(gridcolor=LIGHT, tickprefix="$"),
            xaxis=dict(gridcolor=WHITE), showlegend=False,
        )
        st.plotly_chart(fig_rpr, use_container_width=True)

# =============================
# TAB 2 — TIMESLOTS
# =============================
with tab2:
    st.markdown('<div class="section-header"><h3>Timeslot Performance</h3></div>', unsafe_allow_html=True)

    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        slot_filter = st.multiselect(
            "Filter by Timeslot", options=sorted(df_curr["slot_key"].unique()),
            default=[], placeholder="All timeslots"
        )
    with col_f2:
        hm_metric = st.radio(
            "Heatmap metric",
            ["Average Riders","Utilization %","Late Cancels + No Shows",
             "Late Cancels","No Shows","Membership %","Credit %"],
            horizontal=False
        )

    if hm_metric == "Average Riders":
        hm_raw = df_curr.groupby(["dow","slot_time"])["checked_in"].mean().reset_index()
        hm_raw["value"] = hm_raw["checked_in"].round(1)
        hm_raw["label"] = hm_raw["value"].round(0).astype(int).astype(str)
        colorbar_title = "Avg Riders"; scale = UTIL_SCALE; zmin, zmax = 0, 41
    elif hm_metric == "Utilization %":
        hm_raw = df_curr.groupby(["dow","slot_time"])["util"].mean().reset_index()
        hm_raw["value"] = hm_raw["util"].round(3)
        hm_raw["label"] = (hm_raw["value"]*100).round(0).astype(int).astype(str)+"%"
        colorbar_title = "Utilization"; scale = UTIL_SCALE; zmin, zmax = 0, 1.0
    elif hm_metric == "Late Cancels + No Shows":
        hm_raw = df_curr.groupby(["dow","slot_time"])["total_dropout"].mean().reset_index()
        hm_raw["value"] = hm_raw["total_dropout"].round(1)
        hm_raw["label"] = hm_raw["value"].astype(str)
        colorbar_title = "Avg Dropouts"; scale = DROPOUT_SCALE; zmin, zmax = 0, None
    elif hm_metric == "Late Cancels":
        hm_raw = df_curr.groupby(["dow","slot_time"])["late_cancel"].mean().reset_index()
        hm_raw["value"] = hm_raw["late_cancel"].round(1)
        hm_raw["label"] = hm_raw["value"].astype(str)
        colorbar_title = "Avg Late Cancels"; scale = DROPOUT_SCALE; zmin, zmax = 0, None
    elif hm_metric == "No Shows":
        hm_raw = df_curr.groupby(["dow","slot_time"])["no_show"].mean().reset_index()
        hm_raw["value"] = hm_raw["no_show"].round(1)
        hm_raw["label"] = hm_raw["value"].astype(str)
        colorbar_title = "Avg No Shows"; scale = DROPOUT_SCALE; zmin, zmax = 0, None
    elif hm_metric == "Membership %":
        hm_raw = df_curr.groupby(["dow","slot_time"])["membership_rate"].mean().reset_index()
        hm_raw["value"] = hm_raw["membership_rate"].round(3)
        hm_raw["label"] = (hm_raw["value"]*100).round(0).astype(int).astype(str)+"%"
        colorbar_title = "Member Rate"; scale = MIX_SCALE; zmin, zmax = 0, 1.0
    else:
        hm_raw = df_curr.groupby(["dow","slot_time"])["credit_rate"].mean().reset_index()
        hm_raw["value"] = hm_raw["credit_rate"].round(3)
        hm_raw["label"] = (hm_raw["value"]*100).round(0).astype(int).astype(str)+"%"
        colorbar_title = "Credit Rate"; scale = CREDIT_SCALE; zmin, zmax = 0, 1.0

    hm_matrix = pd.DataFrame(index=SLOT_ORDER, columns=DOW_ORDER, dtype=float)
    hm_text   = pd.DataFrame(index=SLOT_ORDER, columns=DOW_ORDER, data="")
    for _, row in hm_raw.iterrows():
        d, s = row["dow"], row["slot_time"]
        if d in DOW_ORDER and s in SLOT_ORDER:
            hm_matrix.loc[s, d] = row["value"]
            hm_text.loc[s, d]   = row["label"]

    fig_hm = go.Figure(data=go.Heatmap(
        z=hm_matrix.values.tolist(),
        x=DOW_LABELS, y=SLOT_ORDER,
        text=[[hm_text.loc[s, d] for d in DOW_ORDER] for s in SLOT_ORDER],
        texttemplate="%{text}",
        colorscale=scale, showscale=True,
        zmin=zmin, zmax=zmax if zmax else None,
        colorbar=dict(title=colorbar_title),
    ))
    fig_hm.update_layout(
        title=f"Schedule Heatmap — {hm_metric}",
        height=420, margin=dict(l=60, r=20, t=40, b=20),
        paper_bgcolor=WHITE, plot_bgcolor=WHITE,
        font=dict(size=11), yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig_hm, use_container_width=True)

    if hm_metric in ["Average Riders","Utilization %"]:
        st.caption("Red = busy/high. Blue = slow/low.")
    elif hm_metric == "Membership %":
        st.caption("Green = membership heavy. White = balanced. Purple = credit heavy.")
    elif hm_metric == "Credit %":
        st.caption("Purple = credit heavy. White = balanced. Green = membership heavy.")
    else:
        st.caption("Deeper red = higher dropout. White = minimal dropout.")

    slot_perf = (
        df_curr.groupby("slot_key")
        .agg(
            classes=("util","size"), riders=("checked_in","sum"),
            avg_riders=("checked_in","mean"), utilization=("util","mean"),
            baseline=("slot_baseline_util","mean"), delta=("delta_vs_slot","mean"),
            delta_trailing=("delta_vs_trailing_3m","mean"),
            avg_dropout=("total_dropout","mean"),
            membership_rate=("membership_rate","mean"), credit_rate=("credit_rate","mean"),
        ).reset_index()
    )
    slot_perf = slot_perf[slot_perf["classes"] >= MIN_CLASSES].copy()
    if slot_filter:
        slot_perf = slot_perf[slot_perf["slot_key"].isin(slot_filter)]

    for col in ["utilization","baseline","delta","delta_trailing","membership_rate","credit_rate"]:
        slot_perf[col] = slot_perf[col].round(3)
    for col in ["avg_riders","avg_dropout"]:
        slot_perf[col] = slot_perf[col].round(1)
    slot_perf = slot_perf.sort_values("delta", ascending=True)

    slot_display = slot_perf.copy()
    slot_display["utilization"]    = slot_display["utilization"].apply(lambda x: f"{round(x*100,1)}%")
    slot_display["baseline"]       = slot_display["baseline"].apply(lambda x: f"{round(x*100,1)}%")
    slot_display["delta"]          = slot_display["delta"].apply(lambda x: f"+{round(x*100,1)}%" if x >= 0 else f"{round(x*100,1)}%")
    slot_display["delta_trailing"] = slot_display["delta_trailing"].apply(
        lambda x: f"+{round(x*100,1)}%" if pd.notna(x) and x >= 0 else f"{round(x*100,1)}%" if pd.notna(x) else "N/A")
    slot_display["membership_rate"] = slot_display["membership_rate"].apply(lambda x: f"{round(x*100,1)}%")
    slot_display["credit_rate"]     = slot_display["credit_rate"].apply(lambda x: f"{round(x*100,1)}%")
    slot_display = slot_display.rename(columns={
        "slot_key":"Timeslot","classes":"Classes","riders":"Total Riders",
        "avg_riders":"Avg Riders","utilization":"Utilization","baseline":"Baseline",
        "delta":"Delta vs Baseline","delta_trailing":"Delta vs 3M",
        "avg_dropout":"Avg Dropout","membership_rate":"Member %","credit_rate":"Credit %"
    })
    st.dataframe(
        slot_display[["Timeslot","Classes","Total Riders","Avg Riders","Utilization",
                      "Baseline","Delta vs Baseline","Delta vs 3M","Avg Dropout","Member %","Credit %"]],
        use_container_width=True, hide_index=True
    )
    desc("Baseline = historical average utilization for this slot. Delta vs Baseline = how this period compared to that average. Delta vs 3M = performance vs the trailing 3-month average. Avg Dropout = average penalty cancellations and no-shows per class.")

    if slot_filter:
        st.markdown("**Slot Trend — Selected Timeslots**")
        slot_trend = (
            df[df["slot_key"].isin(slot_filter)]
            .groupby(["slot_key","month"])["util"].mean().reset_index()
        )
        slot_trend["label"]    = slot_trend["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))
        slot_trend["util_pct"] = (slot_trend["util"] * 100).round(1)
        fig_st = px.line(slot_trend, x="label", y="util_pct", color="slot_key",
                         markers=True, title="Timeslot Utilization Over Time", text="util_pct")
        fig_st.update_traces(textposition="top center", texttemplate="%{text}%")
        fig_st.update_layout(
            height=320, margin=dict(l=20, r=20, t=40, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(ticksuffix="%", range=[0,100], gridcolor=LIGHT),
            xaxis_title="", legend_title="Slot",
        )
        st.plotly_chart(fig_st, use_container_width=True)

# =============================
# TAB 3 — INSTRUCTORS
# =============================
with tab3:
    st.markdown('<div class="section-header"><h3>Instructor Performance</h3></div>', unsafe_allow_html=True)

    st.markdown("**Slot Difficulty by Instructor**")
    desc("This table shows the average historical difficulty of each instructor's assigned timeslots — based on the long-run average utilization of those slots before accounting for the instructor's impact. A lower baseline means the instructor is working harder slots. Read this alongside lift to understand the full picture.")

    slot_difficulty = (
        df_curr.groupby("instructor_first")
        .agg(classes=("util","size"), avg_slot_baseline=("slot_baseline_util","mean"),
             utilization=("util","mean"), lift=("delta_vs_slot","mean")).reset_index()
    )
    slot_difficulty = slot_difficulty[slot_difficulty["classes"] >= 1].copy()
    for col in ["avg_slot_baseline","utilization","lift"]:
        slot_difficulty[col] = slot_difficulty[col].round(3)
    slot_difficulty = slot_difficulty.sort_values("avg_slot_baseline", ascending=True)

    sd_display = slot_difficulty.copy()
    sd_display["avg_slot_baseline"] = sd_display["avg_slot_baseline"].apply(lambda x: f"{round(x*100,1)}%")
    sd_display["utilization"]       = sd_display["utilization"].apply(lambda x: f"{round(x*100,1)}%")
    sd_display["lift"]              = sd_display["lift"].apply(fmt_pct)
    sd_display = sd_display.rename(columns={
        "instructor_first":"Instructor","classes":"Classes Taught",
        "avg_slot_baseline":"Avg Slot Difficulty","utilization":"Actual Utilization","lift":"Lift vs Slot",
    })
    st.dataframe(sd_display[["Instructor","Classes Taught","Avg Slot Difficulty","Actual Utilization","Lift vs Slot"]],
                 use_container_width=True, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)

    all_instructors = sorted(df["instructor_first"].unique())
    instr_filter = st.multiselect("Filter by Instructor", options=all_instructors, default=[], placeholder="All instructors")

    instr_perf = (
        df_curr.groupby("instructor_first")
        .agg(classes=("util","size"), riders=("checked_in","sum"),
             avg_riders=("checked_in","mean"), utilization=("util","mean"),
             slot_baseline=("slot_baseline_util","mean"), lift=("delta_vs_slot","mean"),
             lift_trailing=("delta_vs_trailing_3m","mean"), avg_dropout=("total_dropout","mean")).reset_index()
    )
    for col in ["utilization","slot_baseline","lift","lift_trailing"]:
        instr_perf[col] = instr_perf[col].round(3)
    for col in ["avg_riders","avg_dropout"]:
        instr_perf[col] = instr_perf[col].round(1)
    instr_perf = instr_perf.sort_values("lift", ascending=False)

    instr_perf_display = instr_perf[instr_perf["instructor_first"].isin(instr_filter)] if instr_filter else instr_perf

    fig_lift = go.Figure(go.Bar(
        x=instr_perf_display["lift"], y=instr_perf_display["instructor_first"],
        orientation="h",
        marker_color=[ACCENT if v >= 0 else "#F4CCCC" for v in instr_perf_display["lift"]],
        text=[fmt_pct(v) for v in instr_perf_display["lift"]], textposition="outside",
    ))
    fig_lift.add_vline(x=0, line_dash="dash", line_color=BLACK, line_width=1)
    fig_lift.update_layout(
        title="Lift vs Slot Baseline",
        height=max(300, len(instr_perf_display)*38),
        margin=dict(l=20, r=80, t=40, b=20),
        paper_bgcolor=WHITE, plot_bgcolor=WHITE,
        xaxis=dict(tickformat=".0%"), yaxis=dict(autorange="reversed"), showlegend=False,
    )
    st.plotly_chart(fig_lift, use_container_width=True)
    desc("Lift vs Slot measures how an instructor performs relative to the historical average for their assigned timeslots. A positive lift means they fill their classes above what those slots typically produce.")

    instr_display = instr_perf_display.copy()
    instr_display["avg_riders"]     = instr_display["avg_riders"].apply(lambda x: round(x,1))
    instr_display["utilization"]    = instr_display["utilization"].apply(lambda x: f"{round(x*100,1)}%")
    instr_display["slot_baseline"]  = instr_display["slot_baseline"].apply(lambda x: f"{round(x*100,1)}%")
    instr_display["lift"]           = instr_display["lift"].apply(fmt_pct)
    instr_display["lift_trailing"]  = instr_display["lift_trailing"].apply(fmt_pct)
    instr_display = instr_display.rename(columns={
        "instructor_first":"Instructor","classes":"Classes","riders":"Total Riders",
        "avg_riders":"Avg Riders","utilization":"Utilization","slot_baseline":"Slot Baseline",
        "lift":"Lift vs Slot","lift_trailing":"Lift vs 3M","avg_dropout":"Avg Dropout",
    })
    st.dataframe(
        instr_display[["Instructor","Classes","Total Riders","Avg Riders","Utilization",
                       "Slot Baseline","Lift vs Slot","Lift vs 3M","Avg Dropout"]],
        use_container_width=True, hide_index=True
    )
    desc("Slot Baseline = historical average utilization for the slots this instructor teaches. Lift vs Slot = performance above or below that baseline. Lift vs 3M = trailing 3-month comparison. Avg Dropout = average penalty cancellations and no-shows per class.")

    if instr_filter:
        st.markdown("**Instructor Trend vs Studio Average**")
        studio_trend = df.groupby("month")["util"].mean().reset_index()
        studio_trend["label"]    = studio_trend["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))
        studio_trend["util_pct"] = (studio_trend["util"]*100).round(1)

        instr_trend = (
            df[df["instructor_first"].isin(instr_filter)]
            .groupby(["instructor_first","month"])["util"].mean().reset_index()
        )
        instr_trend["label"]    = instr_trend["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))
        instr_trend["util_pct"] = (instr_trend["util"]*100).round(1)

        fig_it = go.Figure()
        fig_it.add_trace(go.Scatter(
            x=studio_trend["label"], y=studio_trend["util_pct"],
            mode="lines+markers+text",
            line=dict(color=GREY, width=1.5, dash="dash"),
            marker=dict(color=GREY, size=5),
            text=[f"{v}%" for v in studio_trend["util_pct"]],
            textposition="bottom center", textfont=dict(size=8, color=GREY),
            name="Studio Average",
        ))
        colors = [BLACK, "#2471A3", "#C0392B", "#27AE60", "#F0C080"]
        for i, instr in enumerate(instr_filter):
            idata = instr_trend[instr_trend["instructor_first"] == instr]
            fig_it.add_trace(go.Scatter(
                x=idata["label"], y=idata["util_pct"],
                mode="lines+markers+text",
                line=dict(color=colors[i % len(colors)], width=2),
                marker=dict(size=7),
                text=[f"{v}%" for v in idata["util_pct"]],
                textposition="top center", textfont=dict(size=9),
                name=instr,
            ))
        fig_it.update_layout(
            height=340, margin=dict(l=20, r=20, t=20, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(ticksuffix="%", range=[0,100], gridcolor=LIGHT),
            xaxis_title="", yaxis_title="Utilization",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        )
        st.plotly_chart(fig_it, use_container_width=True)
        desc("Monthly utilization trend for selected instructors versus the studio average (dashed line).")

        st.markdown(f"**Slot Breakdown — {display_label}**")
        instr_slot = (
            df_curr[df_curr["instructor_first"].isin(instr_filter)]
            .groupby(["instructor_first","slot_key"])
            .agg(classes=("util","size"), avg_riders=("checked_in","mean"),
                 utilization=("util","mean"), lift=("delta_vs_slot","mean"),
                 avg_dropout=("total_dropout","mean")).reset_index()
        )
        instr_slot["utilization"] = instr_slot["utilization"].apply(lambda x: f"{round(x*100,1)}%")
        instr_slot["avg_riders"]  = instr_slot["avg_riders"].round(1)
        instr_slot["avg_dropout"] = instr_slot["avg_dropout"].round(1)
        instr_slot["lift"]        = instr_slot["lift"].apply(fmt_pct)
        instr_slot = instr_slot.rename(columns={
            "instructor_first":"Instructor","slot_key":"Slot","classes":"Classes",
            "avg_riders":"Avg Riders","utilization":"Utilization",
            "lift":"Lift vs Slot","avg_dropout":"Avg Dropout"
        })
        st.dataframe(instr_slot, use_container_width=True, hide_index=True)
        desc("Per-slot performance for the selected period.")

# =============================
# TAB 4 — TRENDS
# =============================
with tab4:
    st.markdown('<div class="section-header"><h3>Season Trends</h3></div>', unsafe_allow_html=True)

    monthly_trend = (
        df.groupby("month")
        .agg(classes=("util","size"), riders=("checked_in","sum"), utilization=("util","mean"))
        .reset_index().sort_values("month")
    )
    monthly_trend["label"]    = monthly_trend["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))
    monthly_trend["util_pct"] = (monthly_trend["utilization"]*100).round(1)
    monthly_trend["current"]  = monthly_trend["month"] == (selected_month or "")

    col_tl, col_tr = st.columns(2)
    with col_tl:
        fig_util = go.Figure()
        fig_util.add_trace(go.Scatter(
            x=monthly_trend["label"], y=monthly_trend["util_pct"],
            mode="lines+markers+text", line=dict(color=BLACK, width=2),
            marker=dict(
                color=[ACCENT if c else BLACK for c in monthly_trend["current"]],
                size=[14 if c else 7 for c in monthly_trend["current"]],
                line=dict(color=BLACK, width=1)
            ),
            text=[f"{v}%" for v in monthly_trend["util_pct"]],
            textposition="top center", textfont=dict(size=10),
        ))
        fig_util.update_layout(
            title="Monthly Studio Utilization", height=320,
            margin=dict(l=20, r=20, t=40, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(ticksuffix="%", range=[0,100], gridcolor=LIGHT),
            xaxis=dict(gridcolor=WHITE), showlegend=False,
        )
        st.plotly_chart(fig_util, use_container_width=True)

    with col_tr:
        fig_riders = go.Figure(go.Bar(
            x=monthly_trend["label"], y=monthly_trend["riders"],
            marker_color=[BLACK if c else ACCENT for c in monthly_trend["current"]],
            text=monthly_trend["riders"], textposition="outside",
        ))
        fig_riders.update_layout(
            title="Monthly Total Riders", height=320,
            margin=dict(l=20, r=20, t=40, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(gridcolor=LIGHT), xaxis_title="", showlegend=False,
        )
        st.plotly_chart(fig_riders, use_container_width=True)

    if "week" in df.columns:
        weekly_trend = df.groupby("week")["util"].mean().reset_index()
        weekly_trend["week_start"] = pd.to_datetime(weekly_trend["week"].str.slice(0,10))
        weekly_trend = weekly_trend.sort_values("week_start")
        weekly_trend["util_pct"] = (weekly_trend["util"]*100).round(1)
        weekly_trend["label"]    = weekly_trend["week_start"].dt.strftime("%b %d")
        fig_weekly = go.Figure()
        fig_weekly.add_trace(go.Scatter(
            x=weekly_trend["label"], y=weekly_trend["util_pct"],
            mode="lines+markers+text", line=dict(color=GREY, width=1.5),
            marker=dict(color=GREY, size=5),
            text=[f"{v}%" for v in weekly_trend["util_pct"]],
            textposition="top center", textfont=dict(size=8),
        ))
        fig_weekly.update_layout(
            title="Weekly Utilization — Full Season", height=300,
            margin=dict(l=20, r=20, t=40, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(ticksuffix="%", range=[0,100], gridcolor=LIGHT),
            xaxis=dict(gridcolor=WHITE), showlegend=False,
        )
        st.plotly_chart(fig_weekly, use_container_width=True)

    st.markdown(f"**Day of Week — {display_label}**")
    dow_summary = (
        df_curr.groupby("dow")
        .agg(classes=("util","size"), riders=("checked_in","sum"), utilization=("util","mean"))
        .reset_index()
    )
    dow_summary["dow"] = pd.Categorical(dow_summary["dow"], categories=DOW_ORDER, ordered=True)
    dow_summary = dow_summary.sort_values("dow")
    dow_summary["util_pct"] = (dow_summary["utilization"]*100).round(1)
    fig_dow = go.Figure(go.Bar(
        x=dow_summary["dow"], y=dow_summary["util_pct"], marker_color=ACCENT,
        text=[f"{v}%" for v in dow_summary["util_pct"]], textposition="outside",
    ))
    fig_dow.update_layout(
        title=f"Utilization by Day of Week — {display_label}", height=280,
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor=WHITE, plot_bgcolor=WHITE,
        yaxis=dict(ticksuffix="%", range=[0,100], gridcolor=LIGHT),
        xaxis_title="", showlegend=False,
    )
    st.plotly_chart(fig_dow, use_container_width=True)

    if not rev_df.empty:
        st.markdown("**Revenue vs Utilization — Season to Date**")
        desc("Class revenue (bars) and studio utilization (line) by month.")
        rev_monthly = (
            rev_df[rev_df["Product Type"].isin(["Credit","Membership"])]
            .groupby("month")["Realized Revenue"].sum().reset_index()
        )
        combined = monthly_trend.merge(rev_monthly, on="month", how="left")
        combined["label"] = combined["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))
        fig_combo = go.Figure()
        fig_combo.add_trace(go.Bar(
            x=combined["label"], y=combined["Realized Revenue"],
            name="Class Revenue (CAD)", marker_color=LIGHT,
            marker_line_color=GREY, marker_line_width=0.5, opacity=0.85, yaxis="y2",
        ))
        fig_combo.add_trace(go.Scatter(
            x=combined["label"], y=combined["util_pct"],
            name="Utilization %", mode="lines+markers+text",
            line=dict(color=BLACK, width=3),
            marker=dict(color=BLACK, size=9, line=dict(color=WHITE, width=2)),
            text=[f"{v}%" for v in combined["util_pct"]],
            textposition="top center", textfont=dict(size=10, color=BLACK), yaxis="y1",
        ))
        fig_combo.update_layout(
            height=360, margin=dict(l=20, r=80, t=20, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(ticksuffix="%", range=[0,100], gridcolor=LIGHT, title="Utilization %", layer="above traces"),
            yaxis2=dict(overlaying="y", side="right", tickprefix="$", title="Revenue (CAD)", showgrid=False, rangemode="tozero"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
            xaxis_title="",
        )
        st.plotly_chart(fig_combo, use_container_width=True)

    # ---- ORDERS TRENDS ----
    if not ord_purchases.empty:
        st.markdown("---")
        st.markdown('<div class="section-header"><h3>Orders Trends</h3></div>', unsafe_allow_html=True)

        core_all = ord_purchases[ord_purchases["is_core"]].copy()
        core_all["window"] = core_all["hour"].apply(assign_window)

        # CHART 1: Order value by group + utilization line
        # Groups: Credits, Intro Offers, Memberships (new purchases + renewals combined)
        st.markdown("**Order Value and Studio Utilization by Month**")
        desc(
            "Stacked bars show monthly order value by product group. Memberships includes both new membership "
            "purchases and auto-renewals. The line shows studio utilization. "
            "Use this to see whether sales activity and class fill rates are moving together or diverging."
        )

        credits_rev = (
            core_all[core_all["product_group"] == "Credits"]
            .groupby("month")["Line Total"].sum().reset_index()
            .rename(columns={"Line Total": "Credits"})
        )
        intro_rev = (
            core_all[core_all["product_group"] == "Intro Offers"]
            .groupby("month")["Line Total"].sum().reset_index()
            .rename(columns={"Line Total": "Intro Offers"})
        )
        # New membership purchases
        new_mem_rev = (
            core_all[core_all["product_group"] == "Memberships"]
            .groupby("month")["Line Total"].sum().reset_index()
            .rename(columns={"Line Total": "new_mem"})
        )
        # Renewals
        if not ord_renewals.empty and "Line Total" in ord_renewals.columns:
            renewal_rev = (
                ord_renewals.groupby("month")["Line Total"].sum().reset_index()
                .rename(columns={"Line Total": "renewal_mem"})
            )
        else:
            renewal_rev = pd.DataFrame(columns=["month","renewal_mem"])

        all_trend_months = sorted(core_all["month"].unique())
        trend_base = pd.DataFrame({"month": all_trend_months})
        trend_base["month_label"] = trend_base["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))

        trend_base = trend_base.merge(credits_rev, on="month", how="left")
        trend_base = trend_base.merge(intro_rev,   on="month", how="left")
        trend_base = trend_base.merge(new_mem_rev, on="month", how="left")
        trend_base = trend_base.merge(renewal_rev, on="month", how="left")
        trend_base = trend_base.fillna(0)
        trend_base["Memberships"] = trend_base["new_mem"] + trend_base["renewal_mem"]

        util_by_month = df.groupby("month")["util"].mean().reset_index()
        util_by_month["util_pct"]    = (util_by_month["util"]*100).round(1)
        util_by_month["month_label"] = util_by_month["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))

        fig_orders_util = go.Figure()
        bar_groups = [("Credits", ACCENT), ("Intro Offers", BLACK), ("Memberships", GREY)]
        for grp, color in bar_groups:
            if grp not in trend_base.columns:
                continue
            fig_orders_util.add_trace(go.Bar(
                x=trend_base["month_label"], y=trend_base[grp],
                name=grp, marker_color=color,
                text=[fmt_cad(v) if v > 0 else "" for v in trend_base[grp]],
                textposition="inside", textfont=dict(color=WHITE, size=10),
                yaxis="y2",
            ))

        fig_orders_util.add_trace(go.Scatter(
            x=util_by_month["month_label"], y=util_by_month["util_pct"],
            name="Utilization %", mode="lines+markers+text",
            line=dict(color=BLACK, width=3),
            marker=dict(color=BLACK, size=9, line=dict(color=WHITE, width=2)),
            text=[f"{v}%" for v in util_by_month["util_pct"]],
            textposition="top center", textfont=dict(size=10, color=BLACK),
            yaxis="y1",
        ))

        fig_orders_util.update_layout(
            barmode="stack", height=420,
            margin=dict(l=20, r=80, t=20, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(ticksuffix="%", range=[0,100], gridcolor=LIGHT,
                       title="Utilization %", layer="above traces"),
            yaxis2=dict(overlaying="y", side="right", tickprefix="$",
                        title="Order Value (CAD)", showgrid=False, rangemode="tozero"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
            xaxis_title="",
        )
        st.plotly_chart(fig_orders_util, use_container_width=True)

        # CHART 2: Order value by product group per month (standalone)
        st.markdown("**Order Value by Product Group — Monthly Trend**")
        desc(
            "Total order value (CAD) by product group per month. "
            "Note: this is order value at time of purchase, not recognized revenue. "
            "Memberships includes new purchases and renewals combined."
        )
        fig_trend_rev = go.Figure()
        for grp, color in bar_groups:
            if grp not in trend_base.columns:
                continue
            fig_trend_rev.add_trace(go.Bar(
                x=trend_base["month_label"], y=trend_base[grp],
                name=grp, marker_color=color,
                text=[fmt_cad(v) if v > 0 else "" for v in trend_base[grp]],
                textposition="inside", textfont=dict(color=WHITE, size=10),
            ))
        fig_trend_rev.update_layout(
            barmode="stack", height=340,
            margin=dict(l=20, r=20, t=20, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(gridcolor=LIGHT, tickprefix="$", title="Order Value (CAD)"),
            xaxis_title="",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        )
        st.plotly_chart(fig_trend_rev, use_container_width=True)

        # CHART 3: Orders by purchase window per month
        st.markdown("**When Orders Are Placed — Monthly Trend by Purchase Window**")
        desc(
            "Order volume by purchase window per month. "
            "Windows: Early Morning (5am-8am), Mid Morning (8am-11am), Midday (11am-2pm), "
            "Afternoon (2pm-5pm), Evening (5pm-10pm), Late Night (10pm-5am)."
        )
        trend_win = (
            core_all.groupby(["month","window"])["Line Quantity"].sum().reset_index()
        )
        trend_win["month_label"] = trend_win["month"].apply(lambda m: pd.Period(m, freq="M").strftime("%b %Y"))
        trend_win["window"] = pd.Categorical(trend_win["window"], categories=WINDOW_ORDER, ordered=True)
        trend_win = trend_win.sort_values(["month","window"])

        fig_trend_win = go.Figure()
        for window in WINDOW_ORDER:
            wdata = trend_win[trend_win["window"] == window].sort_values("month")
            if wdata.empty:
                continue
            fig_trend_win.add_trace(go.Bar(
                x=wdata["month_label"], y=wdata["Line Quantity"],
                name=window, marker_color=WINDOW_COLORS.get(window, GREY),
                text=wdata["Line Quantity"].astype(int),
                textposition="inside", textfont=dict(color=WHITE, size=9),
            ))
        fig_trend_win.update_layout(
            barmode="stack", height=340,
            margin=dict(l=20, r=20, t=20, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(gridcolor=LIGHT, title="Units Sold"),
            xaxis_title="",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        )
        st.plotly_chart(fig_trend_win, use_container_width=True)

# =============================
# TAB 5 — ORDERS
# =============================
with tab5:
    st.markdown('<div class="section-header"><h3>Sales and Orders</h3></div>', unsafe_allow_html=True)

    if ord_purchases.empty:
        st.warning("No orders data found. Run orders_pipeline.py first.")
        st.stop()

    core_curr = ord_curr[ord_curr["is_core"]].copy() if not ord_curr.empty else pd.DataFrame()

    # ---- PRODUCT GROUP TOGGLE ----
    st.markdown("**Product Group**")
    group_options  = ["All"] + GROUPS
    sel_group_btn  = st.radio("", group_options, horizontal=True, key="ord_group_btn", label_visibility="collapsed")
    show_rev       = st.radio("Metric", ["Orders", "Order Value (CAD)"], key="ord_metric", horizontal=True)
    use_rev        = (show_rev == "Order Value (CAD)")
    val_label      = "Order Value (CAD)" if use_rev else "Orders"

    filtered = pd.DataFrame()
    if not core_curr.empty:
        filtered = core_curr.copy() if sel_group_btn == "All" else core_curr[core_curr["product_group"] == sel_group_btn].copy()

    if filtered.empty:
        st.info("No data for the selected period.")
        st.stop()

    # ---- MEMBERSHIP KPI STRIP ----
    st.markdown("**Membership Activity**")
    desc("Renewals are auto-processed billing events excluded from purchase analysis. Net New = manual membership purchases only (excludes renewals and Intro Offers).")

    # Renewals for selected period
    total_ren  = len(ren_curr) if not ren_curr.empty else 0
    days_total = sum(pd.Period(m, freq="M").days_in_month for m in sel_months)
    avg_daily  = total_ren / days_total if days_total > 0 else 0

    # MoM on renewals -- compare last two months in full dataset, not just selection
    ren_mom = None
    if not ord_renewals.empty and "month" in ord_renewals.columns:
        all_ren_months = sorted(ord_renewals["month"].unique())
        # Find the last month in the current selection that also has a prior month
        sel_sorted = sorted(sel_months)
        last_sel   = sel_sorted[-1]
        if last_sel in all_ren_months:
            last_idx = all_ren_months.index(last_sel)
            if last_idx >= 1:
                prev_month = all_ren_months[last_idx - 1]
                last_count = len(ord_renewals[ord_renewals["month"] == last_sel])
                prev_count = len(ord_renewals[ord_renewals["month"] == prev_month])
                last_days  = pd.Period(last_sel, freq="M").days_in_month
                prev_days  = pd.Period(prev_month, freq="M").days_in_month
                last_daily = last_count / last_days if last_days > 0 else 0
                prev_daily = prev_count / prev_days if prev_days > 0 else 0
                if prev_daily > 0:
                    ren_mom = (last_daily - prev_daily) / prev_daily

    # Net new memberships (non-renewal membership purchases)
    net_new_mem = 0
    if not core_curr.empty:
        net_new_mem = int(core_curr[core_curr["product_group"] == "Memberships"]["Line Quantity"].sum())

    mk1, mk2, mk3, mk4 = st.columns(4)
    with mk1:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{total_ren:,}</div><div class="metric-label">Total Renewals</div></div>', unsafe_allow_html=True)
    with mk2:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{avg_daily:.1f}</div><div class="metric-label">Avg Daily Renewals</div></div>', unsafe_allow_html=True)
    with mk3:
        mom_str = fmt_pct(ren_mom) if ren_mom is not None else "N/A"
        st.markdown(f'<div class="metric-card"><div class="metric-value">{mom_str}</div><div class="metric-label">MoM Change (Renewals)</div></div>', unsafe_allow_html=True)
    with mk4:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{net_new_mem:,}</div><div class="metric-label">Net New Memberships</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- PURCHASE KPI STRIP ----
    st.markdown("**Purchase Activity**")

    total_orders = filtered["Order Number"].nunique()
    total_units  = int(filtered["Line Quantity"].sum())
    total_val    = filtered["Line Total"].sum()
    intro_units  = int(core_curr[core_curr["product_group"] == "Intro Offers"]["Line Quantity"].sum()) if not core_curr.empty else 0

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{total_orders:,}</div><div class="metric-label">Total Orders</div></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{total_units:,}</div><div class="metric-label">Units Sold</div></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_cad(total_val)}</div><div class="metric-label">Order Value (CAD)</div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{intro_units:,}</div><div class="metric-label">New Rider Purchases</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- HEATMAP ----
    st.markdown("**When Are Purchases Happening?**")
    st.info(
        "Monday 1pm is annotated separately -- this spike is driven by the weekly schedule release, "
        "not organic demand. The heat scale is calculated excluding this cell so the rest of the week reads clearly."
    )

    hm_mode = st.radio("View by", ["Hourly", "Purchase Window"], horizontal=True, key="hm_mode")

    DOW_ORDER_ORD = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    DOW_SHORT     = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]

    if hm_mode == "Hourly":
        HOUR_ORDER  = list(range(5, 24)) + list(range(0, 5))
        HOUR_LABELS = [f"{h:02d}:00" for h in HOUR_ORDER]

        if use_rev:
            hm_agg  = filtered.groupby(["hour","dow"])["Line Total"].sum().reset_index().rename(columns={"Line Total":"value"})
            text_fn = lambda v: f"${v:,.0f}" if v > 0 else ""
        else:
            hm_agg  = filtered.groupby(["hour","dow"]).size().reset_index(name="value")
            text_fn = lambda v: str(int(v)) if v > 0 else ""

        hm_pivot = (
            hm_agg.pivot(index="hour", columns="dow", values="value")
            .reindex(index=HOUR_ORDER, columns=DOW_ORDER_ORD).fillna(0)
        )

        # Extract and annotate Monday 13:00, then zero it out for scale purposes
        mon13_val = hm_pivot.loc[13, "Monday"] if 13 in hm_pivot.index and "Monday" in hm_pivot.columns else 0
        mon13_label = text_fn(mon13_val) + " *" if mon13_val > 0 else "*"
        hm_pivot_scaled = hm_pivot.copy()
        hm_pivot_scaled.loc[13, "Monday"] = 0  # exclude from colour scale

        text_vals = []
        for h in HOUR_ORDER:
            row_labels = []
            for d in DOW_ORDER_ORD:
                v = hm_pivot.loc[h, d]
                if h == 13 and d == "Monday":
                    row_labels.append(mon13_label)
                else:
                    row_labels.append(text_fn(v))
            text_vals.append(row_labels)

        fig_hm = go.Figure(data=go.Heatmap(
            z=hm_pivot_scaled.values.tolist(),
            x=DOW_SHORT, y=HOUR_LABELS,
            text=text_vals, texttemplate="%{text}",
            colorscale=ORDERS_HEATMAP_SCALE,
            showscale=True, colorbar=dict(title=val_label),
        ))
        fig_hm.update_layout(
            title=f"Purchase Heatmap — {val_label} by Hour and Day of Week",
            height=580, margin=dict(l=80, r=20, t=40, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            font=dict(size=10), yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig_hm, use_container_width=True)
        desc(
            "Red = high activity, Blue = low activity. "
            "* Monday 13:00 is annotated but excluded from the colour scale -- this spike reflects the weekly "
            "schedule release at 1pm, not organic purchase behaviour. "
            "Hours 00:00-04:00 appear at the bottom and represent genuine overnight purchases."
        )

    else:
        # Window heatmap -- 1pm Monday falls in Midday window, same treatment
        if use_rev:
            hm_agg_w  = filtered.groupby(["window","dow"])["Line Total"].sum().reset_index().rename(columns={"Line Total":"value"})
            text_fn_w = lambda v: f"${v:,.0f}" if v > 0 else ""
        else:
            hm_agg_w  = filtered.groupby(["window","dow"]).size().reset_index(name="value")
            text_fn_w = lambda v: str(int(v)) if v > 0 else ""

        hm_pivot_w = (
            hm_agg_w.pivot(index="window", columns="dow", values="value")
            .reindex(index=WINDOW_ORDER, columns=DOW_ORDER_ORD).fillna(0)
        )
        text_vals_w = [[text_fn_w(hm_pivot_w.loc[w, d]) for d in DOW_ORDER_ORD] for w in WINDOW_ORDER]

        fig_hm_w = go.Figure(data=go.Heatmap(
            z=hm_pivot_w.values.tolist(),
            x=DOW_SHORT, y=WINDOW_ORDER,
            text=text_vals_w, texttemplate="%{text}",
            colorscale=ORDERS_HEATMAP_SCALE,
            showscale=True, colorbar=dict(title=val_label),
        ))
        fig_hm_w.update_layout(
            title=f"Purchase Heatmap — {val_label} by Window and Day of Week",
            height=360, margin=dict(l=140, r=20, t=40, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            font=dict(size=11), yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig_hm_w, use_container_width=True)
        desc(
            "Purchase windows: Early Morning (5am-8am), Mid Morning (8am-11am), Midday (11am-2pm), "
            "Afternoon (2pm-5pm), Evening (5pm-10pm), Late Night (10pm-5am). "
            "Red = high activity, Blue = low activity. "
            "Note: Monday Midday includes the 1pm schedule release spike."
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- PRODUCT MIX BY WINDOW ----
    st.markdown("**Product Mix by Purchase Window**")
    desc(
        "Credits vs Memberships vs Intro Offers split by purchase window. "
        "Use this to understand what type of buyer is active at each time of day."
    )

    sel_days = st.multiselect(
        "Filter by Day of Week", options=DOW_ORDER_ORD, default=DOW_ORDER_ORD, key="mix_dow_filter"
    )
    if not sel_days:
        sel_days = DOW_ORDER_ORD

    mix_data = core_curr[core_curr["dow"].isin(sel_days)].copy() if not core_curr.empty else pd.DataFrame()

    if not mix_data.empty:
        mix_agg = (
            mix_data.groupby(["window","product_group"])
            .agg(orders=("Line Quantity","sum"), revenue=("Line Total","sum"))
            .reset_index()
        )
        mix_metric = "revenue" if use_rev else "orders"
        mix_ylabel  = "Order Value (CAD)" if use_rev else "Units"
        mix_agg["window"] = pd.Categorical(mix_agg["window"], categories=WINDOW_ORDER, ordered=True)
        mix_agg = mix_agg.sort_values("window")

        fig_mix = go.Figure()
        for group in GROUPS:
            gdata = mix_agg[mix_agg["product_group"] == group]
            if gdata.empty:
                continue
            fig_mix.add_trace(go.Bar(
                x=gdata["window"], y=gdata[mix_metric],
                name=group, marker_color=GROUP_COLORS.get(group, GREY),
                text=gdata[mix_metric].apply(lambda v: fmt_cad(v) if use_rev else str(int(v))),
                textposition="inside", textfont=dict(color=WHITE, size=10),
            ))

        fig_mix.update_layout(
            barmode="stack",
            title=f"Product Mix by Purchase Window — {mix_ylabel}",
            height=380,
            margin=dict(l=20, r=20, t=80, b=20),
            paper_bgcolor=WHITE, plot_bgcolor=WHITE,
            yaxis=dict(gridcolor=LIGHT, tickprefix="$" if use_rev else "", title=mix_ylabel),
            xaxis_title="Purchase Window",
            legend=dict(orientation="h", yanchor="bottom", y=1.12, xanchor="left", x=0),
        )
        st.plotly_chart(fig_mix, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- PIE + TABLE ----
    st.markdown("**Product Mix — Period Summary**")
    desc(f"Order value and unit split across product groups for the selected period: {display_label}.")

    if not core_curr.empty:
        pie_data = (
            core_curr.groupby("product_group")
            .agg(units=("Line Quantity","sum"), value=("Line Total","sum"))
            .reindex(GROUPS).dropna(how="all").reset_index()
        )

        col_pie, col_tbl = st.columns([1, 1])
        with col_pie:
            pie_metric = "value" if use_rev else "units"
            pie_labels = pie_data["product_group"].tolist()
            pie_values = pie_data[pie_metric].tolist()
            pie_colors = [GROUP_COLORS.get(g, GREY) for g in pie_labels]
            total_pie  = sum(pie_values)
            custom_text = [
                f"{fmt_cad(v)}<br>{v/total_pie*100:.1f}%" if use_rev
                else f"{int(v)} units<br>{v/total_pie*100:.1f}%"
                for v in pie_values
            ]
            fig_pie = go.Figure(data=go.Pie(
                labels=pie_labels, values=pie_values,
                marker=dict(colors=pie_colors),
                textinfo="label+percent",
                hovertemplate="%{label}<br>%{customdata}<extra></extra>",
                customdata=custom_text, hole=0.35,
            ))
            fig_pie.update_layout(
                height=320, margin=dict(l=20, r=20, t=40, b=20),
                paper_bgcolor=WHITE, showlegend=False,
                title=f"{'Order Value' if use_rev else 'Units'} by Product Group",
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_tbl:
            tbl = pie_data.copy()
            total_val_tbl = tbl["value"].sum()
            tbl["pct_value"] = (tbl["value"] / total_val_tbl * 100).round(1).astype(str) + "%"
            tbl["value"]     = tbl["value"].apply(fmt_cad)
            tbl = tbl.rename(columns={
                "product_group":"Product Group","units":"Units",
                "value":"Order Value (CAD)","pct_value":"% of Value",
            })
            st.dataframe(tbl, use_container_width=True, hide_index=True)
            desc(
                "Intro Offers = New Rider Special and 2-Week Unlimited only. "
                "These are acquisition signals -- read the unit count as new customers entering the studio."
            )

st.markdown("---")
st.caption(f"SPINCO London — Data through {month_labels[available_months[-1]]} — Internal Use Only")