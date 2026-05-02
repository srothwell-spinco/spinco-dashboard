import os
import pandas as pd
import numpy as np
from datetime import date
from orders_pipeline import load_orders, build_orders_summary
from orders_sheet import write_orders_sheet

# =============================
# CONFIG
# =============================
OUT_DIR = "out"
os.makedirs(OUT_DIR, exist_ok=True)
MODEL_PATH = os.path.join(OUT_DIR, "model.csv")

MIN_CLASSES = 4
WEAK_DELTA = -0.07
STRONG_DELTA = 0.07

# SPINCO Brand Colours
BLACK = "#000000"
WHITE = "#FFFFFF"
LIGHT_BLUE = "#BBD7ED"
LIGHT_GREY = "#F4F4F4"
DARK_GREY = "#4D4D4D"
RED = "#F4CCCC"
DELTA_GREEN = "#C6EFCE"
DELTA_RED = "#F4CCCC"

# =============================
# LOAD MODEL
# =============================
df = pd.read_csv(MODEL_PATH)
df["month"] = df["month"].astype(str)

this_month = pd.Period(date.today(), freq="M")
months = pd.PeriodIndex(df["month"], freq="M").unique()
full_months = [m for m in months if m < this_month]
report_month = max(full_months) if full_months else max(months)
current_month = report_month.strftime("%Y-%m")

df_curr = df[df["month"] == current_month].copy()
print(f"Report month: {current_month} ({len(df_curr)} classes)")

# Exclude TEAM TEACH globally
df = df[df["report_instructor"] != "TEAM TEACH"].copy()
df_curr = df_curr[df_curr["report_instructor"] != "TEAM TEACH"].copy()

# =============================
# LOAD ORDERS
# =============================
print("Loading orders data...")
orders_purchases, orders_renewals = load_orders()
orders_summary = build_orders_summary(orders_purchases, orders_renewals)

# =============================
# DATE FORMATTING HELPERS
# =============================
def fmt_month(m):
    try:
        return pd.Period(m, freq="M").strftime("%b %Y")
    except:
        return m

# =============================
# PAGE 1 — EXECUTIVE SUMMARY
# =============================
total_riders = int(df_curr["checked_in"].sum())
total_classes = len(df_curr)
utilization = round(df_curr["util"].mean(), 3)
median_util = round(df_curr["util"].median(), 3)

pct_above_70 = f"{round((df_curr['util'] >= 0.70).mean() * 100, 1)}%"
pct_below_40 = f"{round((df_curr['util'] < 0.40).mean() * 100, 1)}%"

top_day = df_curr.groupby("dow")["util"].mean().idxmax()

slot_counts = df_curr.groupby("slot_key")["util"].agg(["mean", "count"])
eligible_slots = slot_counts[slot_counts["count"] >= MIN_CLASSES]
top_slot = eligible_slots["mean"].idxmax() if not eligible_slots.empty else "N/A"

sorted_months = sorted([m.strftime("%Y-%m") for m in full_months])
mom_change = round(utilization - df[df["month"] == sorted_months[-2]]["util"].mean(), 3) if len(sorted_months) >= 2 else "N/A"
trailing_3m_change = round(utilization - df[df["month"] == sorted_months[-4]]["util"].mean(), 3) if len(sorted_months) >= 4 else "N/A"

exec_rows = [
    ("Report Month", fmt_month(current_month)),
    ("Total Riders", f"{total_riders:,}"),
    ("Total Classes", f"{total_classes:,}"),
    ("Studio Utilization", f"{round(utilization * 100, 1)}%"),
    ("Median Utilization", f"{round(median_util * 100, 1)}%"),
    ("% Classes >= 70%", pct_above_70),
    ("% Classes < 40%", pct_below_40),
    ("Top Day", top_day),
    ("Top Timeslot (min 4 classes)", top_slot),
    ("MoM Change", f"{round(mom_change * 100, 1)}%" if mom_change != "N/A" else "N/A"),
    ("Trailing 3-Month Change", f"{round(trailing_3m_change * 100, 1)}%" if trailing_3m_change != "N/A" else "N/A"),
]
exec_summary = pd.DataFrame(exec_rows, columns=["Metric", "Value"])

# =============================
# PAGE 2 — UTILIZATION BUCKETS
# =============================
KNOWN_SLOTS = ["06:00","07:00","08:00","09:00","09:30","10:15","11:30",
               "16:45","17:00","17:30","18:35","19:40"]

def bucket_row(r):
    if str(r["dow"]).lower() in ["sat","saturday","sun","sunday"]:
        return "Weekend"
    if r["slot_time"] in ["06:00","07:00"]:
        return "Morning"
    if r["slot_time"] in ["08:00","09:00","09:30"]:
        return "Midday"
    if r["slot_time"] in ["10:15","11:30"]:
        return "Late Morning"
    return "Evening"

df_curr["bucket"] = df_curr.apply(bucket_row, axis=1)

unknown_slots = df_curr[~df_curr["slot_time"].isin(KNOWN_SLOTS)]["slot_time"].unique()
if len(unknown_slots) > 0:
    print(f"WARNING: Unknown slot times found: {unknown_slots}")

weekday_summary = (
    df_curr[df_curr["bucket"] != "Weekend"]
    .groupby("bucket")
    .agg(classes=("util","size"), riders=("checked_in","sum"), utilization=("util","mean"))
    .reset_index()
)
weekday_summary["utilization"] = weekday_summary["utilization"].round(3)

weekend_summary = (
    df_curr[df_curr["bucket"] == "Weekend"]
    .groupby("slot_time")
    .agg(classes=("util","size"), riders=("checked_in","sum"), utilization=("util","mean"))
    .reset_index()
)
weekend_summary["utilization"] = weekend_summary["utilization"].round(3)

# =============================
# PAGE 2B — DAY OF WEEK
# =============================
dow_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
dow_summary = (
    df_curr.groupby("dow")
    .agg(classes=("util","size"), riders=("checked_in","sum"),
         utilization=("util","mean"), utilization_median=("util","median"))
    .reset_index()
)
dow_summary["utilization"] = dow_summary["utilization"].round(3)
dow_summary["utilization_median"] = dow_summary["utilization_median"].round(3)
dow_summary["dow"] = pd.Categorical(dow_summary["dow"], categories=dow_order, ordered=True)
dow_summary = dow_summary.sort_values("dow")

# =============================
# PAGE 3 — TIMESLOT PERFORMANCE
# =============================
slot_perf = (
    df_curr.groupby("slot_key")
    .agg(
        dow_group=("dow_group","first"),
        slot_time=("slot_time","first"),
        classes=("util","size"),
        riders=("checked_in","sum"),
        utilization=("util","mean"),
        utilization_stddev=("util","std"),
        baseline_util=("slot_baseline_util","mean"),
        delta_vs_baseline=("delta_vs_slot","mean"),
        delta_vs_trailing_3m=("delta_vs_trailing_3m","mean"),
        delta_vs_prev_month=("delta_vs_prev_month","mean"),
    )
    .reset_index()
)
for col in ["utilization","utilization_stddev","baseline_util",
            "delta_vs_baseline","delta_vs_trailing_3m","delta_vs_prev_month"]:
    slot_perf[col] = slot_perf[col].round(3)

slot_perf["flag_low_sample"] = slot_perf["classes"].apply(lambda x: "Yes" if x < MIN_CLASSES else "")
slot_perf["flag_low_sample_negative"] = slot_perf.apply(
    lambda r: "Yes" if r["classes"] < MIN_CLASSES and r["delta_vs_baseline"] != "" and r["delta_vs_baseline"] < 0 else "", axis=1)
slot_perf["flag_structurally_weak"] = slot_perf.apply(
    lambda r: "Yes" if r["classes"] >= MIN_CLASSES and r["delta_vs_baseline"] != "" and r["delta_vs_baseline"] <= WEAK_DELTA else "", axis=1)
slot_perf["flag_opportunity"] = slot_perf.apply(
    lambda r: "Yes" if r["classes"] >= MIN_CLASSES and r["delta_vs_baseline"] != "" and r["delta_vs_baseline"] >= STRONG_DELTA else "", axis=1)
slot_perf = slot_perf.sort_values(["delta_vs_baseline","classes"], ascending=[True,False])

# =============================
# PAGE 3B — SLOT TRAJECTORY
# =============================
all_months_sorted = sorted(df["month"].unique())
recent_3 = all_months_sorted[-3:]
prior_3 = all_months_sorted[-6:-3] if len(all_months_sorted) >= 6 else all_months_sorted[:-3]

recent_slot = df[df["month"].isin(recent_3)].groupby("slot_key")["util"].agg(classes_recent="count", util_recent="mean").reset_index()
prior_slot = df[df["month"].isin(prior_3)].groupby("slot_key")["util"].agg(classes_prior="count", util_prior="mean").reset_index()

slot_trajectory = recent_slot.merge(prior_slot, on="slot_key", how="left")
slot_trajectory["util_recent"] = slot_trajectory["util_recent"].round(3)
slot_trajectory["util_prior"] = slot_trajectory["util_prior"].round(3)
slot_trajectory["trajectory_delta"] = (slot_trajectory["util_recent"] - slot_trajectory["util_prior"]).round(3)
slot_trajectory["trend"] = slot_trajectory["trajectory_delta"].apply(
    lambda x: "Growing" if x >= 0.05 else ("Declining" if x <= -0.05 else "Stable") if pd.notna(x) else "No prior data"
)
slot_trajectory = slot_trajectory.sort_values("trajectory_delta", ascending=True)

# =============================
# PAGE 3C — SLOT LONGITUDINAL
# =============================
slot_monthly = (
    df.groupby(["slot_key","month"])
    .agg(utilization=("util","mean"), baseline=("slot_baseline_util","mean"))
    .reset_index()
)
slot_monthly["delta"] = (slot_monthly["utilization"] - slot_monthly["baseline"]).round(3)
slot_monthly["utilization"] = slot_monthly["utilization"].round(3)
slot_monthly["baseline"] = slot_monthly["baseline"].round(3)

months_available = sorted(df["month"].unique())
slot_long_rows = []
for slot in sorted(slot_monthly["slot_key"].unique()):
    row = {"Slot": slot}
    s = slot_monthly[slot_monthly["slot_key"] == slot]
    for m in months_available:
        mf = fmt_month(m)
        md = s[s["month"] == m]
        row[f"{mf} Util"] = md["utilization"].values[0] if len(md) else ""
        row[f"{mf} Baseline"] = md["baseline"].values[0] if len(md) else ""
        row[f"{mf} Delta"] = md["delta"].values[0] if len(md) else ""
    slot_long_rows.append(row)
slot_longitudinal = pd.DataFrame(slot_long_rows)

# =============================
# PAGE 4 — INSTRUCTOR PERFORMANCE
# =============================
df_curr = df_curr[df_curr["report_instructor"] != "TEAM TEACH"].copy()
instr_perf = (
    df_curr.groupby("report_instructor")
    .agg(
        classes=("util","size"),
        riders=("checked_in","sum"),
        utilization=("util","mean"),
        utilization_stddev=("util","std"),
        slot_baseline=("slot_baseline_util","mean"),
        lift_vs_slot=("delta_vs_slot","mean"),
        lift_vs_trailing_3m=("delta_vs_trailing_3m","mean"),
        lift_vs_prev_month=("delta_vs_prev_month","mean"),
    )
    .reset_index()
)
for col in ["utilization","utilization_stddev","slot_baseline",
            "lift_vs_slot","lift_vs_trailing_3m","lift_vs_prev_month"]:
    instr_perf[col] = instr_perf[col].round(3)
instr_perf = instr_perf.sort_values(["lift_vs_slot","classes"], ascending=[False,False])

# =============================
# PAGE 4B — INSTRUCTOR LONGITUDINAL
# =============================
instr_overall = (
    df.groupby(["report_instructor","month"])
    .agg(utilization=("util","mean"), baseline=("slot_baseline_util","mean"))
    .reset_index()
)
instr_overall["slot_key"] = "OVERALL"

instr_slot = (
    df.groupby(["report_instructor","slot_key","month"])
    .agg(utilization=("util","mean"), baseline=("slot_baseline_util","mean"))
    .reset_index()
)

instr_overall = instr_overall[instr_overall["report_instructor"] != "TEAM TEACH"]
instr_slot = instr_slot[instr_slot["report_instructor"] != "TEAM TEACH"]

instr_long = pd.concat([
    instr_overall[["report_instructor","slot_key","month","utilization","baseline"]],
    instr_slot[["report_instructor","slot_key","month","utilization","baseline"]]
], ignore_index=True)

instr_slot_counts = instr_long[instr_long["slot_key"] != "OVERALL"].groupby(
    ["report_instructor","slot_key"])["utilization"].count().reset_index()
instr_slot_counts.columns = ["report_instructor","slot_key","total_classes"]
valid_slots = instr_slot_counts[instr_slot_counts["total_classes"] >= 3][["report_instructor","slot_key"]]
valid_slots["_keep"] = True
instr_long = instr_long.merge(valid_slots, on=["report_instructor","slot_key"], how="left")
instr_long = instr_long[(instr_long["slot_key"] == "OVERALL") | (instr_long["_keep"] == True)]
instr_long = instr_long.drop(columns=["_keep"])
instr_long["delta"] = (instr_long["utilization"] - instr_long["baseline"]).round(3)
instr_long["utilization"] = instr_long["utilization"].round(3)
instr_long["baseline"] = instr_long["baseline"].round(3)

instr_long_rows = []
instructors = sorted(instr_long["report_instructor"].unique())
for instr in instructors:
    instr_data = instr_long[instr_long["report_instructor"] == instr]
    slots = ["OVERALL"] + sorted([s for s in instr_data["slot_key"].unique() if s != "OVERALL"])
    for slot in slots:
        row = {"Instructor": instr, "Slot": slot}
        s = instr_data[instr_data["slot_key"] == slot]
        for m in months_available:
            mf = fmt_month(m)
            md = s[s["month"] == m]
            row[f"{mf} Util"] = md["utilization"].values[0] if len(md) else ""
            row[f"{mf} Baseline"] = md["baseline"].values[0] if len(md) else ""
            row[f"{mf} Delta"] = md["delta"].values[0] if len(md) else ""
        instr_long_rows.append(row)
instr_longitudinal = pd.DataFrame(instr_long_rows)

# =============================
# PAGE 5 — TRENDS
# =============================
monthly_trend = (
    df.groupby("month")
    .agg(classes=("util","size"), riders=("checked_in","sum"), utilization=("util","mean"))
    .reset_index()
    .sort_values("month")
)
monthly_trend["month"] = monthly_trend["month"].apply(fmt_month)
monthly_trend["utilization"] = monthly_trend["utilization"].round(3)

if "week" in df.columns:
    weekly_trend = (
        df.groupby("week")
        .agg(classes=("util","size"), riders=("checked_in","sum"), utilization=("util","mean"))
        .reset_index()
    )
    weekly_trend["week_start"] = pd.to_datetime(weekly_trend["week"].str.slice(0,10))
    weekly_trend = weekly_trend.sort_values("week_start")
    weekly_trend["week"] = weekly_trend["week_start"].dt.strftime("%b %d, %Y")
    weekly_trend = weekly_trend.drop(columns=["week_start"])
    weekly_trend["utilization"] = weekly_trend["utilization"].round(3)
else:
    weekly_trend = pd.DataFrame({"note": ["No week column found in model.csv"]})

# =============================
# CLEAN NaN VALUES
# =============================
slot_longitudinal = slot_longitudinal.fillna("")
slot_trajectory = slot_trajectory.fillna("")
instr_longitudinal = instr_longitudinal.fillna("")
slot_perf = slot_perf.fillna("")

# =============================
# FORMATTING HELPERS
# =============================
COLOUR_KEY = "Colour coding: Blue rows = utilization at or above 70% (strong). Red rows = utilization below 40% (weak). Green delta cell = performing above baseline. Red delta cell = performing below baseline. Alternating grey rows are for readability only."

def write_explanation(worksheet, workbook, start_row, text):
    fmt = workbook.add_format({
        "font_name": "Arial", "font_size": 11, "font_color": DARK_GREY,
        "bold": True, "text_wrap": True, "valign": "top",
    })
    explanation_row = start_row + 2
    worksheet.merge_range(explanation_row, 0, explanation_row, 8, text, fmt)
    worksheet.set_row(explanation_row, 60)
    colour_fmt = workbook.add_format({
        "font_name": "Arial", "font_size": 9, "font_color": DARK_GREY,
        "italic": True, "text_wrap": True, "valign": "top",
    })
    worksheet.merge_range(explanation_row + 1, 0, explanation_row + 1, 8, COLOUR_KEY, colour_fmt)
    worksheet.set_row(explanation_row + 1, 30)

def format_sheet(writer, sheet_name, df, col_widths=None, header_bg=BLACK, header_fg=WHITE,
                 util_col=None, pct_cols=None, int_cols=None, delta_cols=None,
                 overall_col=None, alt_rows=True, explanation=None, freeze_cols=0):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    nrows, ncols = df.shape

    def make_fmt(bg, num_format=None, bold=False, font_color=BLACK):
        props = {
            "font_name": "Arial", "font_size": 10,
            "bg_color": bg, "border": 0,
            "align": "left", "valign": "vcenter",
            "bold": bold, "font_color": font_color,
        }
        if num_format:
            props["num_format"] = num_format
        return workbook.add_format(props)

    header_fmt = make_fmt(header_bg, bold=True, font_color=header_fg)
    fmt_base = make_fmt(WHITE)
    fmt_alt = make_fmt(LIGHT_GREY)
    fmt_high = make_fmt(LIGHT_BLUE)
    fmt_low = make_fmt(RED)
    fmt_overall = make_fmt(DARK_GREY, bold=True, font_color=WHITE)
    fmt_pct_base = make_fmt(WHITE, "0.0%")
    fmt_pct_alt = make_fmt(LIGHT_GREY, "0.0%")
    fmt_pct_high = make_fmt(LIGHT_BLUE, "0.0%")
    fmt_pct_low = make_fmt(RED, "0.0%")
    fmt_pct_overall = make_fmt(DARK_GREY, "0.0%", bold=True, font_color=WHITE)
    fmt_int_base = make_fmt(WHITE, "#,##0")
    fmt_int_alt = make_fmt(LIGHT_GREY, "#,##0")
    fmt_int_high = make_fmt(LIGHT_BLUE, "#,##0")
    fmt_int_low = make_fmt(RED, "#,##0")
    fmt_delta_green = make_fmt(DELTA_GREEN, "0.0%")
    fmt_delta_red = make_fmt(DELTA_RED, "0.0%")
    fmt_delta_green_overall = make_fmt(DELTA_GREEN, "0.0%", bold=True)
    fmt_delta_red_overall = make_fmt(DELTA_RED, "0.0%", bold=True)

    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_fmt)

    cols = list(df.columns)
    util_idx = cols.index(util_col) if util_col and util_col in cols else None
    pct_idxs = [cols.index(c) for c in (pct_cols or []) if c in cols]
    int_idxs = [cols.index(c) for c in (int_cols or []) if c in cols]
    delta_idxs = [cols.index(c) for c in (delta_cols or []) if c in cols]
    overall_idx = cols.index(overall_col) if overall_col and overall_col in cols else None

    for row_num in range(nrows):
        is_overall = overall_idx is not None and str(df.iloc[row_num, overall_idx]) == "OVERALL"
        is_alt = (row_num % 2 == 1) and alt_rows and not is_overall
        util_val = df.iloc[row_num][util_col] if util_idx is not None else None

        try:
            uv = float(util_val) if util_val not in [None, "", "N/A"] else None
        except:
            uv = None

        if is_overall:
            row_bg = "overall"
        elif uv is not None:
            if uv >= 0.70:
                row_bg = "high"
            elif uv < 0.40:
                row_bg = "low"
            elif is_alt:
                row_bg = "alt"
            else:
                row_bg = "base"
        elif is_alt:
            row_bg = "alt"
        else:
            row_bg = "base"

        for col_num in range(ncols):
            val = df.iloc[row_num, col_num]
            is_pct = col_num in pct_idxs
            is_int = col_num in int_idxs
            is_delta = col_num in delta_idxs

            if is_delta and val != "":
                try:
                    dv = float(val)
                    fmt = (fmt_delta_green_overall if is_overall else fmt_delta_green) if dv >= 0 else (fmt_delta_red_overall if is_overall else fmt_delta_red)
                except:
                    fmt = fmt_overall if is_overall else fmt_base
            elif row_bg == "overall":
                fmt = fmt_pct_overall if is_pct else fmt_overall
            elif row_bg == "high":
                fmt = fmt_pct_high if is_pct else (fmt_int_high if is_int else fmt_high)
            elif row_bg == "low":
                fmt = fmt_pct_low if is_pct else (fmt_int_low if is_int else fmt_low)
            elif row_bg == "alt":
                fmt = fmt_pct_alt if is_pct else (fmt_int_alt if is_int else fmt_alt)
            else:
                fmt = fmt_pct_base if is_pct else (fmt_int_base if is_int else fmt_base)

            worksheet.write(row_num + 1, col_num, val, fmt)

    if col_widths:
        for col_num, width in enumerate(col_widths):
            worksheet.set_column(col_num, col_num, width)
    else:
        for col_num in range(ncols):
            worksheet.set_column(col_num, col_num, 18)

    worksheet.freeze_panes(1, freeze_cols)

    if explanation:
        write_explanation(worksheet, workbook, nrows + 1, explanation)

# =============================
# WRITE TO EXCEL
# =============================
out_xlsx = os.path.join(OUT_DIR, "monthly_pack.xlsx")
with pd.ExcelWriter(out_xlsx, engine="xlsxwriter", mode="w",
                    engine_kwargs={"options": {"nan_inf_to_errors": True}}) as writer:

    exec_summary.to_excel(writer, sheet_name="01_Executive_Summary", index=False)
    weekday_summary.to_excel(writer, sheet_name="02_Weekday_Buckets", index=False)
    weekend_summary.to_excel(writer, sheet_name="02_Weekend_Slots", index=False)
    dow_summary.to_excel(writer, sheet_name="02_Day_of_Week", index=False)
    slot_perf.to_excel(writer, sheet_name="03_Timeslot_Performance", index=False)
    slot_trajectory.to_excel(writer, sheet_name="03_Slot_Trajectory", index=False)
    slot_longitudinal.to_excel(writer, sheet_name="03_Slot_Longitudinal", index=False)
    instr_perf.to_excel(writer, sheet_name="04_Instructor_Performance", index=False)
    instr_longitudinal.to_excel(writer, sheet_name="04_Instructor_Longitudinal", index=False)
    monthly_trend.to_excel(writer, sheet_name="05_Monthly_Trend", index=False)
    weekly_trend.to_excel(writer, sheet_name="05_Weekly_Trend", index=False)

    format_sheet(writer, "01_Executive_Summary", exec_summary,
                 col_widths=[35, 20],
                 explanation="This page shows a snapshot of studio performance for the report month. Utilization is the average percentage of bikes filled per class, treating each class equally regardless of size. MoM Change compares this month to last month. Trailing 3-Month Change compares this month to the same month three months ago.")

    format_sheet(writer, "02_Weekday_Buckets", weekday_summary,
                 col_widths=[20, 12, 14, 18],
                 util_col="utilization", pct_cols=["utilization"],
                 int_cols=["classes","riders"],
                 explanation="Utilization broken down by weekday time bucket. Morning = 6am/7am. Midday = 8am/9am/9:30am. Late Morning = 10:15am/11:30am. Evening = all remaining weekday slots.")

    format_sheet(writer, "02_Weekend_Slots", weekend_summary,
                 col_widths=[20, 12, 14, 18],
                 util_col="utilization", pct_cols=["utilization"],
                 int_cols=["classes","riders"],
                 explanation="Weekend utilization broken down by timeslot. Use this to identify which weekend slots are performing above or below expectations.")

    format_sheet(writer, "02_Day_of_Week", dow_summary,
                 col_widths=[18, 12, 14, 18, 22],
                 util_col="utilization", pct_cols=["utilization","utilization_median"],
                 int_cols=["classes","riders"],
                 explanation="Average and median utilization by day of week for the report month. Use this to identify which days are structurally strong or weak across all timeslots.")

    format_sheet(writer, "03_Timeslot_Performance", slot_perf,
                 col_widths=[22, 14, 12, 10, 12, 14, 16, 16, 16, 16, 18, 22, 18],
                 util_col="utilization",
                 pct_cols=["utilization","baseline_util","delta_vs_baseline",
                           "utilization_stddev","delta_vs_trailing_3m","delta_vs_prev_month"],
                 int_cols=["classes","riders"],
                 delta_cols=["delta_vs_baseline","delta_vs_trailing_3m","delta_vs_prev_month"],
                 explanation="Each row is a recurring timeslot. Baseline = full season average. Delta vs Baseline = performance vs season average. Delta vs Trailing 3M = performance vs last 3 months. Delta vs Prev Month = performance vs last month. Green = above baseline, Red = below.")

    format_sheet(writer, "03_Slot_Trajectory", slot_trajectory,
                 col_widths=[22, 14, 14, 14, 16, 18],
                 pct_cols=["util_recent","util_prior","trajectory_delta"],
                 delta_cols=["trajectory_delta"],
                 explanation="Compares each slot's recent 3-month average against the prior period. Growing = improved by 5 or more percentage points. Declining = dropped by 5 or more points. Stable = within that range.")

    slot_long_pct_cols = [c for c in slot_longitudinal.columns if "Util" in c or "Baseline" in c]
    slot_long_delta_cols = [c for c in slot_longitudinal.columns if "Delta" in c]
    format_sheet(writer, "03_Slot_Longitudinal", slot_longitudinal,
                 col_widths=[22] + [13, 13, 13] * len(months_available),
                 pct_cols=slot_long_pct_cols,
                 delta_cols=slot_long_delta_cols,
                 freeze_cols=1,
                 explanation="Month by month utilization for each timeslot. Each month shows Util (actual), Baseline (season average to date), and Delta. Green delta = beating baseline. Red delta = below baseline.")

    format_sheet(writer, "04_Instructor_Performance", instr_perf,
                 col_widths=[20, 10, 12, 14, 16, 14, 16, 16],
                 util_col="utilization",
                 pct_cols=["utilization","utilization_stddev","slot_baseline",
                           "lift_vs_slot","lift_vs_trailing_3m","lift_vs_prev_month"],
                 int_cols=["classes","riders"],
                 delta_cols=["lift_vs_slot","lift_vs_trailing_3m","lift_vs_prev_month"],
                 explanation="Each row is an instructor. Lift vs Slot = performance vs season baseline. Lift vs Trailing 3M = performance vs last 3 months baseline. Lift vs Prev Month = performance vs last month baseline. Green = positive lift, Red = negative lift.")

    instr_long_pct_cols = [c for c in instr_longitudinal.columns if "Util" in c or "Baseline" in c]
    instr_long_delta_cols = [c for c in instr_longitudinal.columns if "Delta" in c]
    format_sheet(writer, "04_Instructor_Longitudinal", instr_longitudinal,
                 col_widths=[18, 20] + [13, 13, 13] * len(months_available),
                 pct_cols=instr_long_pct_cols,
                 delta_cols=instr_long_delta_cols,
                 overall_col="Slot",
                 freeze_cols=2,
                 explanation="Month by month performance for each instructor by slot. OVERALL rows show blended performance. Each month shows Util, Baseline, and Delta. Green delta = beating baseline. Red delta = below baseline.")

    format_sheet(writer, "05_Monthly_Trend", monthly_trend,
                 col_widths=[18, 12, 14, 18],
                 util_col="utilization", pct_cols=["utilization"],
                 int_cols=["classes","riders"],
                 explanation="Monthly utilization trend across the full season.")

    format_sheet(writer, "05_Weekly_Trend", weekly_trend,
                 col_widths=[28, 12, 14, 18],
                 util_col="utilization", pct_cols=["utilization"],
                 int_cols=["classes","riders"],
                 explanation="Weekly utilization trend across the full season.")

    # Orders sheet
    write_orders_sheet(writer, orders_purchases, orders_renewals, orders_summary, current_month)

    for tab in ["01_Executive_Summary","02_Weekday_Buckets","02_Weekend_Slots","02_Day_of_Week"]:
        writer.sheets[tab].set_tab_color(LIGHT_BLUE)
    for tab in ["03_Timeslot_Performance","03_Slot_Trajectory","03_Slot_Longitudinal"]:
        writer.sheets[tab].set_tab_color("#C6EFCE")
    for tab in ["04_Instructor_Performance","04_Instructor_Longitudinal"]:
        writer.sheets[tab].set_tab_color("#D9B3FF")
    for tab in ["05_Monthly_Trend","05_Weekly_Trend"]:
        writer.sheets[tab].set_tab_color(DARK_GREY)
    if "Orders" in writer.sheets:
        writer.sheets["Orders"].set_tab_color(LIGHT_BLUE)

print(f"Monthly pack written: {out_xlsx}")