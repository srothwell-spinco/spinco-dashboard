# =============================
# ORDERS TAB -- INTEGRATION NOTES
# =============================
# 1. Add orders_pipeline.py to your spinco_dashboard/ folder (same level as dashboard.py)
#
# 2. Add this import at the top of dashboard.py, alongside existing imports:
#       from orders_pipeline import load_orders, build_orders_summary
#
# 3. Add this cache function alongside load_data() and load_revenue():
#
#   @st.cache_data
#   def load_orders_data():
#       purchases = pd.read_csv("out/orders_purchases.csv", parse_dates=["order_dt"])
#       renewals  = pd.read_csv("out/orders_renewals.csv",  parse_dates=["order_dt"])
#       summary   = pd.read_csv("out/orders_summary.csv")
#       return purchases, renewals, summary
#
# 4. Change your tab definition line to add tab5:
#       tab1, tab2, tab3, tab4, tab5 = st.tabs([
#           "📊 Overview", "🕐 Timeslots", "👤 Instructors", "📈 Trends", "🛒 Orders"
#       ])
#
# 5. Paste the block below as: with tab5:
# =============================


with tab5:

    # ---- LOAD ----
    try:
        ord_purchases, ord_renewals, ord_summary = load_orders_data()
    except Exception as e:
        st.warning(f"Orders data not found. Run orders_pipeline.py first. ({e})")
        st.stop()

    st.markdown('<div class="section-header"><h3>Sales and Orders</h3></div>', unsafe_allow_html=True)

    GROUPS    = ["Credits", "Memberships", "Intro Offers"]
    DOW_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    DOW_SHORT = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    GROUP_COLORS = {
        "Credits":     ACCENT,   # light blue
        "Memberships": GREY,     # mid grey
        "Intro Offers": BLACK,   # black -- acquisition stands out
    }

    # ---- FILTERS ----
    available_order_months = sorted(ord_purchases["month"].unique())
    order_month_labels     = {m: pd.Period(m, freq="M").strftime("%B %Y") for m in available_order_months}

    col_f1, col_f2, col_f3 = st.columns([2, 2, 1])
    with col_f1:
        sel_months = st.multiselect(
            "Month(s)",
            options=available_order_months,
            default=available_order_months,
            format_func=lambda m: order_month_labels.get(m, m),
            key="ord_months"
        )
    with col_f2:
        sel_groups = st.multiselect(
            "Product Group(s)",
            options=GROUPS,
            default=GROUPS,
            key="ord_groups"
        )
    with col_f3:
        show_rev = st.radio("Metric", ["Orders", "Revenue (CAD)"], key="ord_metric")

    if not sel_months:
        sel_months = available_order_months
    if not sel_groups:
        sel_groups = GROUPS

    use_rev     = (show_rev == "Revenue (CAD)")
    val_label   = "Revenue (CAD)" if use_rev else "Orders"

    # Core filter applied throughout
    filtered = ord_purchases[
        ord_purchases["is_core"] &
        ord_purchases["month"].isin(sel_months) &
        ord_purchases["product_group"].isin(sel_groups)
    ].copy()

    if filtered.empty:
        st.info("No data for the selected filters.")
        st.stop()

    # ---- RENEWAL KPI STRIP ----
    st.markdown("**Membership Renewals**")
    desc("Auto-renewals are excluded from all purchase analysis below. Shown here as a retention signal only -- rising avg daily renewals means your membership base is holding.")

    if not ord_renewals.empty:
        ren_sel     = ord_renewals[ord_renewals["month"].isin(sel_months)]
        total_ren   = len(ren_sel)
        days_total  = sum(pd.Period(m, freq="M").days_in_month for m in sel_months)
        avg_daily   = total_ren / days_total if days_total > 0 else 0

        # MoM change in avg daily renewals across last two selected months
        ren_mom = None
        if len(sel_months) >= 2 and not ord_summary.empty and "avg_daily_renewals" in ord_summary.columns:
            s = sorted(sel_months)
            last_row = ord_summary[ord_summary["month"] == s[-1]]["avg_daily_renewals"].values
            prev_row = ord_summary[ord_summary["month"] == s[-2]]["avg_daily_renewals"].values
            if len(last_row) and len(prev_row) and prev_row[0] > 0:
                ren_mom = (last_row[0] - prev_row[0]) / prev_row[0]

        rk1, rk2, rk3 = st.columns(3)
        with rk1:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{total_ren:,}</div><div class="metric-label">Total Renewals</div></div>', unsafe_allow_html=True)
        with rk2:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{avg_daily:.1f}</div><div class="metric-label">Avg Daily Renewals</div></div>', unsafe_allow_html=True)
        with rk3:
            mom_str = fmt_pct(ren_mom) if ren_mom is not None else "N/A"
            st.markdown(f'<div class="metric-card"><div class="metric-value">{mom_str}</div><div class="metric-label">MoM Change (Avg Daily)</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- PURCHASE KPI STRIP ----
    st.markdown("**Purchase Activity**")

    total_orders = filtered["Order Number"].nunique()
    total_units  = int(filtered["Line Quantity"].sum())
    total_rev    = filtered["Line Total"].sum()
    intro_units  = int(filtered[filtered["product_group"] == "Intro Offers"]["Line Quantity"].sum())

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{total_orders:,}</div><div class="metric-label">Total Orders</div></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{total_units:,}</div><div class="metric-label">Units Sold</div></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_cad(total_rev)}</div><div class="metric-label">Revenue (CAD)</div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{intro_units:,}</div><div class="metric-label">Intro Offer Units (New Riders)</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- HEATMAP ----
    st.markdown("**When Are Purchases Happening?**")

    st.info(
        "Monday is the highest-volume purchase day but this is driven by the 1pm schedule "
        "release, not organic demand. It is a conversion moment for people already engaged, "
        "not an advertising window."
    )

    # Build pivot: hours 5-23 first, then 0-4 (overnight at bottom)
    HOUR_ORDER  = list(range(5, 24)) + list(range(0, 5))
    HOUR_LABELS = [f"{h:02d}:00" for h in HOUR_ORDER]

    if use_rev:
        hm_agg = (
            filtered.groupby(["hour", "dow"])["Line Total"]
            .sum().reset_index().rename(columns={"Line Total": "value"})
        )
        text_fmt = lambda v: f"${v:,.0f}" if v > 0 else ""
    else:
        hm_agg = (
            filtered.groupby(["hour", "dow"]).size()
            .reset_index(name="value")
        )
        text_fmt = lambda v: str(int(v)) if v > 0 else ""

    hm_pivot = (
        hm_agg.pivot(index="hour", columns="dow", values="value")
        .reindex(index=HOUR_ORDER, columns=DOW_ORDER)
        .fillna(0)
    )

    text_vals = [
        [text_fmt(hm_pivot.loc[h, d]) for d in DOW_ORDER]
        for h in HOUR_ORDER
    ]

    fig_hm = go.Figure(data=go.Heatmap(
        z=hm_pivot.values.tolist(),
        x=DOW_SHORT,
        y=HOUR_LABELS,
        text=text_vals,
        texttemplate="%{text}",
        colorscale=[
            [0.0,  "#F4F4F4"],
            [0.25, "#BBD7ED"],
            [0.6,  "#2471A3"],
            [1.0,  "#000000"],
        ],
        showscale=True,
        colorbar=dict(title=val_label),
    ))
    fig_hm.update_layout(
        title=f"Purchase Heatmap -- {val_label} by Hour and Day of Week",
        height=580,
        margin=dict(l=80, r=20, t=40, b=20),
        paper_bgcolor=WHITE,
        plot_bgcolor=WHITE,
        font=dict(size=10),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig_hm, use_container_width=True)
    desc(
        "Each cell shows total purchases (or revenue) at that hour on that day of week, "
        "across all selected months. Darker = more activity. Rows with no activity show blank. "
        "Hours 00:00-04:00 appear at the bottom -- these are genuine overnight purchases, not system renewals (those are excluded)."
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- PRODUCT MIX BY HOUR ----
    st.markdown("**Product Mix by Hour of Day**")
    desc(
        "Credits (existing rider top-ups) vs Memberships (commitment purchases) vs Intro Offers "
        "(first-time acquisition). The split by hour tells you what kind of buyer is active at each time -- "
        "useful for deciding what to promote and when."
    )

    mix_agg = (
        filtered.groupby(["hour", "product_group"])
        .agg(orders=("Line Quantity", "sum"), revenue=("Line Total", "sum"))
        .reset_index()
    )
    mix_metric = "revenue" if use_rev else "orders"
    mix_ylabel  = "Revenue (CAD)" if use_rev else "Units"

    # Only show hours with at least 3 total units to suppress noise
    active_hours = (
        mix_agg.groupby("hour")[mix_metric].sum()
        .loc[lambda x: x >= 3].index.tolist()
    )
    mix_agg = mix_agg[mix_agg["hour"].isin(active_hours)].copy()
    mix_agg["hour_label"] = mix_agg["hour"].apply(lambda h: f"{h:02d}:00")
    # Sort by the HOUR_ORDER defined above (5am first)
    hour_pos = {h: i for i, h in enumerate(HOUR_ORDER)}
    mix_agg["hour_sort"] = mix_agg["hour"].map(hour_pos)
    mix_agg = mix_agg.sort_values("hour_sort")

    fig_mix = go.Figure()
    for group in GROUPS:
        gdata = mix_agg[mix_agg["product_group"] == group]
        if gdata.empty:
            continue
        fig_mix.add_trace(go.Bar(
            x=gdata["hour_label"],
            y=gdata[mix_metric],
            name=group,
            marker_color=GROUP_COLORS.get(group, GREY),
        ))

    fig_mix.update_layout(
        barmode="stack",
        title=f"Product Mix by Hour -- {mix_ylabel}",
        height=320,
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor=WHITE,
        plot_bgcolor=WHITE,
        yaxis=dict(
            gridcolor=LIGHT,
            tickprefix="$" if use_rev else "",
            title=mix_ylabel,
        ),
        xaxis_title="Hour of Day",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    st.plotly_chart(fig_mix, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- MONTHLY TREND ----
    if len(available_order_months) > 1:
        st.markdown("**Monthly Sales Trend by Product Group**")
        desc("Total units and revenue per month split by product group. Use this to see whether mix is shifting -- for example, if Intro Offers are growing month-over-month, your acquisition funnel is working.")

        trend_agg = (
            ord_purchases[ord_purchases["is_core"] & ord_purchases["month"].isin(sel_months)]
            .groupby(["month", "product_group"])
            .agg(orders=("Line Quantity", "sum"), revenue=("Line Total", "sum"))
            .reset_index()
        )
        trend_agg["month_label"] = trend_agg["month"].apply(
            lambda m: pd.Period(m, freq="M").strftime("%b %Y")
        )
        trend_metric = "revenue" if use_rev else "orders"
        trend_ylabel  = "Revenue (CAD)" if use_rev else "Units"

        fig_trend = go.Figure()
        for group in GROUPS:
            gdata = trend_agg[trend_agg["product_group"] == group].sort_values("month")
            if gdata.empty:
                continue
            fig_trend.add_trace(go.Bar(
                x=gdata["month_label"],
                y=gdata[trend_metric],
                name=group,
                marker_color=GROUP_COLORS.get(group, GREY),
            ))

        fig_trend.update_layout(
            barmode="stack",
            title=f"Monthly Sales by Product Group -- {trend_ylabel}",
            height=320,
            margin=dict(l=20, r=20, t=40, b=20),
            paper_bgcolor=WHITE,
            plot_bgcolor=WHITE,
            yaxis=dict(
                gridcolor=LIGHT,
                tickprefix="$" if use_rev else "",
                title=trend_ylabel,
            ),
            xaxis_title="",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        )
        st.plotly_chart(fig_trend, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

    # ---- PRODUCT GROUP SUMMARY TABLE ----
    st.markdown("**Product Group Summary**")

    tbl = (
        filtered.groupby("product_group")
        .agg(
            orders  = ("Order Number", "nunique"),
            units   = ("Line Quantity", "sum"),
            revenue = ("Line Total", "sum"),
        )
        .reindex(GROUPS)
        .dropna(how="all")
        .reset_index()
    )
    total_rev_tbl = tbl["revenue"].sum()
    tbl["pct_revenue"] = (tbl["revenue"] / total_rev_tbl * 100).round(1).astype(str) + "%"
    tbl["revenue"]     = tbl["revenue"].apply(fmt_cad)
    tbl = tbl.rename(columns={
        "product_group": "Product Group",
        "orders":        "Orders",
        "units":         "Units",
        "revenue":       "Revenue (CAD)",
        "pct_revenue":   "% of Revenue",
    })
    st.dataframe(tbl, use_container_width=True, hide_index=True)
    desc(
        "Intro Offers = New Rider Special and 2-Week Unlimited only. "
        "These represent first-time purchase decisions and should be read as an acquisition metric, "
        "not a revenue metric."
    )
