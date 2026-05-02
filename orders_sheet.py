import pandas as pd
import numpy as np

# =============================
# ORDERS SHEET -- INTEGRATION NOTES
# =============================
# 1. Add this file (orders_sheet.py) to your spinco_dashboard/ folder.
#
# 2. In step3_outputs.py, add at the top:
#       from orders_pipeline import load_orders, build_orders_summary
#       from orders_sheet import write_orders_sheet
#
# 3. After your existing model load, add:
#       orders_purchases, orders_renewals = load_orders()
#       orders_summary = build_orders_summary(orders_purchases, orders_renewals)
#
# 4. Inside your ExcelWriter block, call:
#       write_orders_sheet(writer, orders_purchases, orders_renewals, orders_summary, month)
#    where month is the YYYY-MM string for the report month.
# =============================

GROUPS    = ["Credits", "Memberships", "Intro Offers"]
DOW_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
HOUR_ORDER = list(range(5, 24)) + list(range(0, 5))


def write_orders_sheet(writer, orders_purchases, orders_renewals, orders_summary, month):
    """
    Writes the Orders sheet to monthly_pack.xlsx for the given month (YYYY-MM).

    Sections:
      1. Renewal KPIs
      2. Purchase KPIs
      3. Product group breakdown table
      4. Hour-of-day distribution table (Credits / Memberships / Intro Offers)
      5. Day-of-week distribution table
    """
    wb         = writer.book
    sheet_name = "Orders"

    if orders_purchases.empty:
        pd.DataFrame([{"Note": "No orders data available -- run orders_pipeline.py"}]).to_excel(
            writer, sheet_name=sheet_name, index=False
        )
        return

    # Filter to month, core products only
    month_purch = orders_purchases[
        (orders_purchases["month"] == month) &
        (orders_purchases["is_core"])
    ].copy()

    month_ren = (
        orders_renewals[orders_renewals["month"] == month].copy()
        if not orders_renewals.empty else pd.DataFrame()
    )

    ws = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws

    # ---- FORMATS ----
    fmt_title = wb.add_format({
        "bold": True, "font_size": 13, "font_color": "#FFFFFF",
        "bg_color": "#000000", "align": "left", "valign": "vcenter",
        "bottom": 2,
    })
    fmt_section = wb.add_format({
        "bold": True, "font_size": 11, "font_color": "#000000",
        "bg_color": "#BBD7ED", "align": "left", "valign": "vcenter",
        "top": 1, "bottom": 1,
    })
    fmt_header = wb.add_format({
        "bold": True, "bg_color": "#000000", "font_color": "#FFFFFF",
        "align": "center", "valign": "vcenter", "border": 1,
    })
    fmt_label = wb.add_format({
        "bold": True, "align": "left", "valign": "vcenter",
        "bg_color": "#F4F4F4", "border": 1,
    })
    fmt_num = wb.add_format({"num_format": "#,##0",   "align": "center", "border": 1})
    fmt_cad = wb.add_format({"num_format": "$#,##0",  "align": "center", "border": 1})
    fmt_pct = wb.add_format({"num_format": "0.0%",    "align": "center", "border": 1})
    fmt_dec = wb.add_format({"num_format": "0.0",     "align": "center", "border": 1})
    fmt_mom_pos = wb.add_format({
        "num_format": "+0.0%;-0.0%", "align": "center", "border": 1,
        "font_color": "#27AE60",
    })
    fmt_mom_neg = wb.add_format({
        "num_format": "+0.0%;-0.0%", "align": "center", "border": 1,
        "font_color": "#C0392B",
    })
    fmt_blank = wb.add_format({"bg_color": "#FFFFFF"})

    # Column widths
    ws.set_column(0, 0, 26)
    ws.set_column(1, 8, 16)

    row = 0

    # ---- TITLE ----
    month_label = pd.Period(month, freq="M").strftime("%B %Y")
    ws.merge_range(row, 0, row, 4, f"SPINCO London -- Orders & Sales -- {month_label}", fmt_title)
    row += 2

    # ---- SECTION 1: RENEWAL KPIs ----
    ws.merge_range(row, 0, row, 4, "Membership Renewals (Auto-Processed)", fmt_section)
    row += 1

    total_ren   = len(month_ren)
    days        = pd.Period(month, freq="M").days_in_month
    avg_daily   = round(total_ren / days, 1) if days > 0 else 0

    # MoM change
    ren_mom = None
    if not orders_summary.empty and "avg_daily_renewals" in orders_summary.columns:
        months_sorted = sorted(orders_summary["month"].unique())
        if month in months_sorted:
            idx = months_sorted.index(month)
            if idx >= 1:
                last = orders_summary[orders_summary["month"] == month]["avg_daily_renewals"].values
                prev_m = months_sorted[idx - 1]
                prev = orders_summary[orders_summary["month"] == prev_m]["avg_daily_renewals"].values
                if len(last) and len(prev) and prev[0] > 0:
                    ren_mom = (last[0] - prev[0]) / prev[0]

    ren_kpis = [
        ("Total Renewals",         total_ren,  fmt_num),
        ("Avg Daily Renewals",     avg_daily,  fmt_dec),
        ("MoM Change (Avg Daily)", ren_mom,    fmt_mom_pos if (ren_mom or 0) >= 0 else fmt_mom_neg),
    ]
    for label, val, fmt in ren_kpis:
        ws.write(row, 0, label, fmt_label)
        if val is None:
            ws.write(row, 1, "N/A", fmt_num)
        else:
            ws.write(row, 1, val, fmt)
        row += 1

    row += 1

    # ---- SECTION 2: PURCHASE KPIs ----
    ws.merge_range(row, 0, row, 4, "Purchase Activity (Volitional Purchases Only)", fmt_section)
    row += 1

    if month_purch.empty:
        ws.write(row, 0, "No purchase data for this month", fmt_label)
        row += 2
    else:
        total_orders = month_purch["Order Number"].nunique()
        total_units  = int(month_purch["Line Quantity"].sum())
        total_rev    = month_purch["Line Total"].sum()
        intro_units  = int(
            month_purch[month_purch["product_group"] == "Intro Offers"]["Line Quantity"].sum()
        )

        pur_kpis = [
            ("Total Orders",               total_orders,  fmt_num),
            ("Units Sold",                 total_units,   fmt_num),
            ("Total Revenue (CAD)",        total_rev,     fmt_cad),
            ("Intro Offer Units (New Riders)", intro_units, fmt_num),
        ]
        for label, val, fmt in pur_kpis:
            ws.write(row, 0, label, fmt_label)
            ws.write(row, 1, val, fmt)
            row += 1

    row += 1

    # ---- SECTION 3: PRODUCT GROUP BREAKDOWN ----
    ws.merge_range(row, 0, row, 4, "Product Group Breakdown", fmt_section)
    row += 1

    headers = ["Product Group", "Orders", "Units", "Revenue (CAD)", "% of Revenue"]
    for c, h in enumerate(headers):
        ws.write(row, c, h, fmt_header)
    row += 1

    if not month_purch.empty:
        grp_tbl = (
            month_purch.groupby("product_group")
            .agg(orders=("Order Number", "nunique"),
                 units=("Line Quantity", "sum"),
                 revenue=("Line Total", "sum"))
            .reindex(GROUPS)
            .dropna(how="all")
            .reset_index()
        )
        total_rev_grp = grp_tbl["revenue"].sum()
        for _, r in grp_tbl.iterrows():
            pct = (r["revenue"] / total_rev_grp) if total_rev_grp > 0 else 0
            ws.write(row, 0, r["product_group"],   fmt_label)
            ws.write(row, 1, int(r["orders"]),     fmt_num)
            ws.write(row, 2, int(r["units"]),      fmt_num)
            ws.write(row, 3, r["revenue"],         fmt_cad)
            ws.write(row, 4, pct,                  fmt_pct)
            row += 1

        # Total row
        ws.write(row, 0, "Total", fmt_label)
        ws.write(row, 1, int(grp_tbl["orders"].sum()),  fmt_num)
        ws.write(row, 2, int(grp_tbl["units"].sum()),   fmt_num)
        ws.write(row, 3, grp_tbl["revenue"].sum(),      fmt_cad)
        ws.write(row, 4, 1.0,                           fmt_pct)
        row += 1

    row += 1

    # ---- SECTION 4: PURCHASES BY HOUR ----
    ws.merge_range(row, 0, row, len(GROUPS), "Purchases by Hour of Day (Units)", fmt_section)
    row += 1

    hour_headers = ["Hour"] + GROUPS + ["Total"]
    for c, h in enumerate(hour_headers):
        ws.write(row, c, h, fmt_header)
    row += 1

    if not month_purch.empty:
        hour_agg = (
            month_purch.groupby(["hour", "product_group"])["Line Quantity"]
            .sum().unstack(fill_value=0)
            .reindex(columns=GROUPS, fill_value=0)
        )
        for h in HOUR_ORDER:
            hour_str = f"{h:02d}:00"
            ws.write(row, 0, hour_str, fmt_label)
            row_total = 0
            for c, group in enumerate(GROUPS, start=1):
                val = int(hour_agg.loc[h, group]) if h in hour_agg.index else 0
                ws.write(row, c, val, fmt_num)
                row_total += val
            ws.write(row, len(GROUPS) + 1, row_total, fmt_num)
            row += 1

        # Column totals
        ws.write(row, 0, "Total", fmt_label)
        for c, group in enumerate(GROUPS, start=1):
            ws.write(row, c, int(month_purch[month_purch["product_group"] == group]["Line Quantity"].sum()), fmt_num)
        ws.write(row, len(GROUPS) + 1, int(month_purch["Line Quantity"].sum()), fmt_num)
        row += 1

    row += 1

    # ---- SECTION 5: PURCHASES BY DAY OF WEEK ----
    ws.merge_range(row, 0, row, len(GROUPS), "Purchases by Day of Week (Units)", fmt_section)
    row += 1

    for c, h in enumerate(hour_headers):
        ws.write(row, c, h, fmt_header)
    row += 1

    if not month_purch.empty:
        dow_agg = (
            month_purch.groupby(["dow", "product_group"])["Line Quantity"]
            .sum().unstack(fill_value=0)
            .reindex(index=DOW_ORDER, columns=GROUPS, fill_value=0)
        )
        for day in DOW_ORDER:
            ws.write(row, 0, day, fmt_label)
            row_total = 0
            for c, group in enumerate(GROUPS, start=1):
                val = int(dow_agg.loc[day, group]) if day in dow_agg.index else 0
                ws.write(row, c, val, fmt_num)
                row_total += val
            ws.write(row, len(GROUPS) + 1, row_total, fmt_num)
            row += 1

        # Column totals
        ws.write(row, 0, "Total", fmt_label)
        for c, group in enumerate(GROUPS, start=1):
            ws.write(row, c, int(month_purch[month_purch["product_group"] == group]["Line Quantity"].sum()), fmt_num)
        ws.write(row, len(GROUPS) + 1, int(month_purch["Line Quantity"].sum()), fmt_num)
        row += 1

    ws.write(row + 1, 0, "Intro Offers = New Rider Special and 2-Week Unlimited only. These are acquisition signals, not recurring revenue.", fmt_blank)
