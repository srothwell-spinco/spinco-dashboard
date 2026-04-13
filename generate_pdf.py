import os
import sys
import pandas as pd
import numpy as np
from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
import anthropic

# =============================
# CONFIG
# =============================
OUT_DIR = "out"
MODEL_PATH = os.path.join(OUT_DIR, "model.csv")
ENV_PATH = "/Users/stephenrothwell/Desktop/spinco_dashboard/.env"

with open(ENV_PATH) as f:
    for line in f:
        if 'ANTHROPIC_API_KEY' in line:
            os.environ['ANTHROPIC_API_KEY'] = line.strip().split('=', 1)[1]

BLACK = colors.HexColor("#000000")
WHITE = colors.HexColor("#FFFFFF")
LIGHT_BLUE = colors.HexColor("#BBD7ED")
LIGHT_GREY = colors.HexColor("#F4F4F4")
DARK_GREY = colors.HexColor("#4D4D4D")
MID_GREY = colors.HexColor("#888888")
RED = colors.HexColor("#F4CCCC")
DELTA_GREEN = colors.HexColor("#C6EFCE")
DELTA_RED = colors.HexColor("#F4CCCC")

HM_DEEP_RED = colors.HexColor("#C0392B")
HM_LIGHT_RED = colors.HexColor("#E8A090")
HM_ORANGE = colors.HexColor("#F0C080")
HM_NEUTRAL = colors.HexColor("#F0F0F0")
HM_LIGHT_BLUE = colors.HexColor("#A8CCE8")
HM_DEEP_BLUE = colors.HexColor("#2471A3")

PAGE_W, PAGE_H = A4
MARGIN = 18 * mm
MIN_CLASSES = 4

DOW_ORDER = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
DOW_LABELS = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
SLOT_ORDER = ["06:00","07:00","08:00","09:00","09:30","10:15","11:30",
              "16:45","17:00","17:30","18:35","19:40"]

# =============================
# LOAD MODEL
# =============================
df = pd.read_csv(MODEL_PATH)
df["month"] = df["month"].astype(str)

if len(sys.argv) > 1:
    current_month = sys.argv[1]
    print(f"Running report for specified month: {current_month}")
else:
    this_month = pd.Period(date.today(), freq="M")
    months = pd.PeriodIndex(df["month"], freq="M").unique()
    full_months = [m for m in months if m < this_month]
    report_month = max(full_months) if full_months else max(months)
    current_month = report_month.strftime("%Y-%m")

available_months = sorted(df["month"].unique())
if current_month not in available_months:
    print(f"ERROR: Month {current_month} not found in model.csv")
    print(f"Available months: {', '.join(available_months)}")
    sys.exit(1)

report_month_label = pd.Period(current_month, freq="M").strftime("%B %Y")
OUT_PDF = os.path.join(OUT_DIR, f"monthly_pack_{current_month}.pdf")

df = df[df["report_instructor"] != "TEAM TEACH"].copy()
df_curr = df[df["month"] == current_month].copy()
df_curr = df_curr[df_curr["report_instructor"] != "TEAM TEACH"].copy()

print(f"Report month: {current_month} ({len(df_curr)} classes)")

def first_name(full):
    return str(full).strip().split()[0]

# =============================
# METRICS
# =============================
total_riders = int(df_curr["checked_in"].sum())
total_classes = len(df_curr)
utilization = df_curr["util"].mean()
median_util = df_curr["util"].median()
pct_above_70 = (df_curr["util"] >= 0.70).mean()
pct_below_40 = (df_curr["util"] < 0.40).mean()
top_day = df_curr.groupby("dow")["util"].mean().idxmax()

slot_counts = df_curr.groupby("slot_key")["util"].agg(["mean","count"])
eligible_slots = slot_counts[slot_counts["count"] >= MIN_CLASSES]
top_slot = eligible_slots["mean"].idxmax() if not eligible_slots.empty else "N/A"

idx = available_months.index(current_month)
mom_change = utilization - df[df["month"] == available_months[idx-1]]["util"].mean() if idx >= 1 else None
trailing_3m = utilization - df[df["month"] == available_months[idx-3]]["util"].mean() if idx >= 3 else None

slot_perf = (
    df_curr.groupby("slot_key")
    .agg(
        dow_group=("dow_group","first"),
        slot_time=("slot_time","first"),
        classes=("util","size"),
        riders=("checked_in","sum"),
        utilization=("util","mean"),
        baseline=("slot_baseline_util","mean"),
        delta=("delta_vs_slot","mean"),
        delta_trailing=("delta_vs_trailing_3m","mean"),
    )
    .reset_index()
)
slot_perf = slot_perf[slot_perf["classes"] >= MIN_CLASSES].copy()
for col in ["utilization","baseline","delta","delta_trailing"]:
    slot_perf[col] = slot_perf[col].round(3)
slot_perf = slot_perf.sort_values("delta", ascending=True)

instr_perf = (
    df_curr.groupby("report_instructor")
    .agg(
        classes=("util","size"),
        riders=("checked_in","sum"),
        utilization=("util","mean"),
        slot_baseline=("slot_baseline_util","mean"),
        lift=("delta_vs_slot","mean"),
        lift_trailing=("delta_vs_trailing_3m","mean"),
    )
    .reset_index()
)
for col in ["utilization","slot_baseline","lift","lift_trailing"]:
    instr_perf[col] = instr_perf[col].round(3)
instr_perf["report_instructor"] = instr_perf["report_instructor"].apply(first_name)
instr_perf = instr_perf.sort_values("lift", ascending=False)

trend_data_full = (
    df.groupby("month")
    .agg(classes=("util","size"), riders=("checked_in","sum"), utilization=("util","mean"))
    .reset_index()
    .sort_values("month")
)
trend_data_full["month_label"] = trend_data_full["month"].apply(
    lambda m: pd.Period(m, freq="M").strftime("%B %Y")
)

instr_perf_raw = df_curr.groupby("report_instructor").agg(lift=("delta_vs_slot","mean")).reset_index()
weak_slots = slot_perf[slot_perf["delta"] <= -0.07]["slot_key"].tolist()
strong_slots = slot_perf[slot_perf["delta"] >= 0.07]["slot_key"].tolist()
top_instructors = [first_name(i) for i in instr_perf_raw[instr_perf_raw["lift"] > 0]["report_instructor"].tolist()]
bottom_instructors = [first_name(i) for i in instr_perf_raw[instr_perf_raw["lift"] < 0]["report_instructor"].tolist()]

# =============================
# HEATMAP DATA
# =============================
heatmap_data = df_curr.groupby(["dow","slot_time"])["util"].mean().reset_index()

def get_heatmap_color(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return colors.HexColor("#E0E0E0")
    if val < 0.30:
        return HM_DEEP_RED
    elif val < 0.45:
        return HM_LIGHT_RED
    elif val < 0.55:
        return HM_ORANGE
    elif val < 0.65:
        return HM_NEUTRAL
    elif val < 0.75:
        return HM_LIGHT_BLUE
    else:
        return HM_DEEP_BLUE

# =============================
# AI NARRATIVE
# =============================
print("Generating AI narrative...")

prompt = f"""You are summarizing monthly performance for SPINCO London, a boutique indoor cycling studio. Write in first person plural as the studio owner. Be observational only. No judgment. No corporate language. No em dashes. Only note things that are interesting or not immediately obvious from the data. Use first names only for instructors.

Metrics for {report_month_label}:
- Studio utilization: {round(utilization * 100, 1)}%
- Median utilization: {round(median_util * 100, 1)}%
- Total riders: {total_riders:,}
- Total classes: {total_classes}
- Classes at or above 70%: {round(pct_above_70 * 100, 1)}%
- Classes below 40%: {round(pct_below_40 * 100, 1)}%
- Month on month change: {f"+{round(mom_change * 100, 1)}%" if mom_change and mom_change >= 0 else f"{round(mom_change * 100, 1)}%" if mom_change else "N/A"}
- Trailing 3-month change: {f"+{round(trailing_3m * 100, 1)}%" if trailing_3m and trailing_3m >= 0 else f"{round(trailing_3m * 100, 1)}%" if trailing_3m else "N/A"}
- Top performing day: {top_day}
- Top performing timeslot: {top_slot}
- Slots underperforming vs season baseline: {', '.join(weak_slots) if weak_slots else 'None'}
- Slots outperforming vs season baseline: {', '.join(strong_slots) if strong_slots else 'None'}
- Instructors with positive lift: {', '.join(top_instructors) if top_instructors else 'None'}
- Instructors with negative lift: {', '.join(bottom_instructors) if bottom_instructors else 'None'}

Write 4-5 bullet points. Each bullet is one short observation. Start each with a dash. Focus on what is interesting or non-obvious. Do not repeat what is already visible in the tables. No headers. No preamble."""

client = anthropic.Anthropic()
response = client.messages.create(
    model="claude-sonnet-4-20250514",
    max_tokens=350,
    messages=[{"role": "user", "content": prompt}]
)
narrative = response.content[0].text
print("Narrative generated.")

# =============================
# STYLES
# =============================
def style(name, **kwargs):
    defaults = dict(fontName="Helvetica", fontSize=10, textColor=BLACK, leading=14, spaceAfter=4)
    defaults.update(kwargs)
    return ParagraphStyle(name, **defaults)

S_SECTION_TITLE = style("section_title", fontName="Helvetica-Bold", fontSize=16, textColor=BLACK, leading=20, spaceAfter=4)
S_SECTION_SUB = style("section_sub", fontName="Helvetica", fontSize=9, textColor=DARK_GREY, leading=13, spaceAfter=8)
S_NARRATIVE = style("narrative", fontName="Helvetica", fontSize=10, textColor=BLACK, leading=17, spaceAfter=5)
S_METRIC_LABEL = style("metric_label", fontName="Helvetica", fontSize=8, textColor=DARK_GREY, leading=11, alignment=TA_CENTER)
S_METRIC_VALUE = style("metric_value", fontName="Helvetica-Bold", fontSize=20, textColor=BLACK, leading=24, alignment=TA_CENTER)
S_METRIC_VALUE_SM = style("metric_value_sm", fontName="Helvetica-Bold", fontSize=13, textColor=BLACK, leading=17, alignment=TA_CENTER)
S_TABLE_HEADER = style("table_header", fontName="Helvetica-Bold", fontSize=8, textColor=WHITE, leading=11, alignment=TA_CENTER)
S_TABLE_CELL = style("table_cell", fontName="Helvetica", fontSize=8, textColor=BLACK, leading=11, alignment=TA_LEFT)
S_TABLE_CELL_C = style("table_cell_c", fontName="Helvetica", fontSize=8, textColor=BLACK, leading=11, alignment=TA_CENTER)
S_CAPTION = style("caption", fontName="Helvetica-Oblique", fontSize=7, textColor=DARK_GREY, leading=10, spaceAfter=4)
S_HM_HEADER = style("hm_header", fontName="Helvetica-Bold", fontSize=7, textColor=WHITE, leading=9, alignment=TA_CENTER)
S_HM_LABEL = style("hm_label", fontName="Helvetica", fontSize=7, textColor=BLACK, leading=9, alignment=TA_RIGHT)

def pct(v):
    return f"{round(v * 100, 1)}%" if v is not None else "N/A"

def signed_pct(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    sign = "+" if v >= 0 else ""
    return f"{sign}{round(v * 100, 1)}%"

def hr(color=LIGHT_GREY, thickness=0.5):
    return HRFlowable(width="100%", thickness=thickness, color=color, spaceAfter=6, spaceBefore=2)

# =============================
# PAGE CANVASES
# =============================
def first_page(canvas, doc):
    canvas.saveState()

    # Full black background
    canvas.setFillColor(BLACK)
    canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    # Light blue left spine
    canvas.setFillColor(LIGHT_BLUE)
    canvas.rect(0, 0, 5*mm, PAGE_H, fill=1, stroke=0)

    # // SPINCO — large white text
    canvas.setFont("Helvetica-Bold", 48)
    canvas.setFillColor(WHITE)
    canvas.drawString(MARGIN, PAGE_H * 0.62, "// SPINCO")

    # Light blue accent rule below title
    canvas.setFillColor(LIGHT_BLUE)
    canvas.rect(MARGIN, PAGE_H * 0.58, PAGE_W - 2*MARGIN, 1.5, fill=1, stroke=0)

    # Monthly Performance Report
    canvas.setFont("Helvetica", 13)
    canvas.setFillColor(LIGHT_BLUE)
    canvas.drawString(MARGIN, PAGE_H * 0.53, "MONTHLY PERFORMANCE REPORT")

    # Report month
    canvas.setFont("Helvetica-Bold", 20)
    canvas.setFillColor(WHITE)
    canvas.drawString(MARGIN, PAGE_H * 0.47, report_month_label.upper())

    # Divider
    canvas.setFillColor(DARK_GREY)
    canvas.rect(MARGIN, PAGE_H * 0.43, PAGE_W - 2*MARGIN, 0.5, fill=1, stroke=0)

    # Key stats on cover
    canvas.setFont("Helvetica", 9)
    canvas.setFillColor(MID_GREY)
    canvas.drawString(MARGIN, PAGE_H * 0.38, f"Studio Utilization")
    canvas.setFont("Helvetica-Bold", 9)
    canvas.setFillColor(WHITE)
    canvas.drawString(MARGIN, PAGE_H * 0.34, pct(utilization))

    canvas.setFont("Helvetica", 9)
    canvas.setFillColor(MID_GREY)
    canvas.drawString(MARGIN + 45*mm, PAGE_H * 0.38, f"Total Riders")
    canvas.setFont("Helvetica-Bold", 9)
    canvas.setFillColor(WHITE)
    canvas.drawString(MARGIN + 45*mm, PAGE_H * 0.34, f"{total_riders:,}")

    canvas.setFont("Helvetica", 9)
    canvas.setFillColor(MID_GREY)
    canvas.drawString(MARGIN + 90*mm, PAGE_H * 0.38, f"MoM Change")
    canvas.setFont("Helvetica-Bold", 9)
    canvas.setFillColor(WHITE)
    canvas.drawString(MARGIN + 90*mm, PAGE_H * 0.34, signed_pct(mom_change))

    # Bottom confidential note
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(DARK_GREY)
    canvas.drawString(MARGIN, 12*mm, "Confidential — Internal Use Only")
    canvas.drawRightString(PAGE_W - MARGIN, 12*mm, f"SPINCO London — {report_month_label}")

    # Bottom blue rule
    canvas.setFillColor(LIGHT_BLUE)
    canvas.rect(5*mm, 9*mm, PAGE_W - 5*mm, 1, fill=1, stroke=0)

    canvas.restoreState()

def later_pages(canvas, doc):
    canvas.saveState()

    # Left blue spine
    canvas.setFillColor(LIGHT_BLUE)
    canvas.rect(0, 0, 5*mm, PAGE_H, fill=1, stroke=0)

    # Black top header
    canvas.setFillColor(BLACK)
    canvas.rect(5*mm, PAGE_H - 10*mm, PAGE_W - 5*mm, 10*mm, fill=1, stroke=0)
    canvas.setFont("Helvetica-Bold", 8)
    canvas.setFillColor(WHITE)
    canvas.drawString(MARGIN, PAGE_H - 6.5*mm, "// SPINCO")
    canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 6.5*mm,
                           f"{report_month_label} — Performance Report")

    # Blue accent under header
    canvas.setFillColor(LIGHT_BLUE)
    canvas.rect(5*mm, PAGE_H - 10*mm - 1.5, PAGE_W - 5*mm, 1.5, fill=1, stroke=0)

    # Footer
    canvas.setFillColor(LIGHT_GREY)
    canvas.rect(5*mm, 0, PAGE_W - 5*mm, 9*mm, fill=1, stroke=0)
    canvas.setFillColor(LIGHT_BLUE)
    canvas.rect(5*mm, 9*mm, PAGE_W - 5*mm, 1, fill=1, stroke=0)
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(DARK_GREY)
    canvas.drawString(MARGIN, 3*mm, "Confidential — Internal Use Only")
    canvas.drawRightString(PAGE_W - MARGIN, 3*mm, f"Page {doc.page - 1}")

    canvas.restoreState()

# =============================
# BUILD STORY
# =============================
content_w = PAGE_W - 2 * MARGIN
story = []
story.append(PageBreak())

# =============================
# PAGE 2 — STUDIO SNAPSHOT
# =============================
story.append(Paragraph("Studio Snapshot", S_SECTION_TITLE))
story.append(Paragraph(f"Performance overview for {report_month_label}.", S_SECTION_SUB))
story.append(hr(BLACK, 1))
story.append(Spacer(1, 4*mm))

for line in narrative.strip().split("\n"):
    line = line.strip()
    if line:
        story.append(Paragraph(line, S_NARRATIVE))
story.append(Spacer(1, 5*mm))
story.append(hr())
story.append(Spacer(1, 4*mm))

col_w = content_w / 4
kpi_data = [
    [Paragraph(pct(utilization), S_METRIC_VALUE), Paragraph(f"{total_riders:,}", S_METRIC_VALUE),
     Paragraph(str(total_classes), S_METRIC_VALUE), Paragraph(pct(median_util), S_METRIC_VALUE)],
    [Paragraph("Studio Utilization", S_METRIC_LABEL), Paragraph("Total Riders", S_METRIC_LABEL),
     Paragraph("Total Classes", S_METRIC_LABEL), Paragraph("Median Utilization", S_METRIC_LABEL)],
]
kpi_table = Table(kpi_data, colWidths=[col_w]*4)
kpi_table.setStyle(TableStyle([
    ("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("BACKGROUND", (0,0), (-1,-1), LIGHT_GREY),
    ("TOPPADDING", (0,0), (-1,-1), 10), ("BOTTOMPADDING", (0,0), (-1,-1), 10),
    ("LINEAFTER", (0,0), (2,1), 0.5, WHITE),
]))
story.append(kpi_table)
story.append(Spacer(1, 3*mm))

col_w2 = content_w / 3
sec_data = [
    [Paragraph(pct(pct_above_70), S_METRIC_VALUE_SM), Paragraph(pct(pct_below_40), S_METRIC_VALUE_SM),
     Paragraph(signed_pct(mom_change), S_METRIC_VALUE_SM)],
    [Paragraph("Classes at or above 70%", S_METRIC_LABEL), Paragraph("Classes below 40%", S_METRIC_LABEL),
     Paragraph("Month on Month Change", S_METRIC_LABEL)],
]
sec_table = Table(sec_data, colWidths=[col_w2]*3)
sec_table.setStyle(TableStyle([
    ("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("TOPPADDING", (0,0), (-1,-1), 8), ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ("LINEAFTER", (0,0), (1,1), 0.5, LIGHT_GREY),
    ("LINEBELOW", (0,0), (-1,0), 0.5, LIGHT_GREY),
]))
story.append(sec_table)
story.append(Spacer(1, 3*mm))

hl_data = [
    [Paragraph("TOP DAY", S_TABLE_HEADER), Paragraph("TOP TIMESLOT", S_TABLE_HEADER),
     Paragraph("TRAILING 3-MONTH CHANGE", S_TABLE_HEADER)],
    [Paragraph(top_day, style("td", fontName="Helvetica-Bold", fontSize=12, alignment=TA_CENTER)),
     Paragraph(top_slot, style("ts", fontName="Helvetica-Bold", fontSize=10, alignment=TA_CENTER)),
     Paragraph(signed_pct(trailing_3m), style("t3", fontName="Helvetica-Bold", fontSize=12, alignment=TA_CENTER))],
]
hl_table = Table(hl_data, colWidths=[content_w/3]*3)
hl_table.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,0), BLACK), ("BACKGROUND", (0,1), (-1,1), LIGHT_BLUE),
    ("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("TOPPADDING", (0,0), (-1,-1), 9), ("BOTTOMPADDING", (0,0), (-1,-1), 9),
    ("LINEAFTER", (0,0), (1,-1), 0.5, WHITE),
]))
story.append(hl_table)
story.append(PageBreak())

# =============================
# PAGE 3 — TIMESLOT PERFORMANCE
# =============================
story.append(Paragraph("Timeslot Performance", S_SECTION_TITLE))
story.append(Paragraph(
    "Timeslots with a minimum of 4 classes this month. Baseline = full season average to date. Delta vs 3M = performance vs trailing 3-month slot average.",
    S_SECTION_SUB))
story.append(hr(BLACK, 1))
story.append(Spacer(1, 3*mm))

slot_headers = ["Timeslot", "Classes", "Riders", "Utilization", "Baseline", "Delta vs Baseline", "Delta vs 3M"]
slot_col_w = [content_w*0.24, content_w*0.08, content_w*0.08,
              content_w*0.12, content_w*0.12, content_w*0.18, content_w*0.18]

slot_table_data = [[Paragraph(h, S_TABLE_HEADER) for h in slot_headers]]
for _, row in slot_perf.iterrows():
    slot_table_data.append([
        Paragraph(str(row["slot_key"]), S_TABLE_CELL),
        Paragraph(str(int(row["classes"])), S_TABLE_CELL_C),
        Paragraph(f"{int(row['riders']):,}", S_TABLE_CELL_C),
        Paragraph(pct(row["utilization"]), S_TABLE_CELL_C),
        Paragraph(pct(row["baseline"]), S_TABLE_CELL_C),
        Paragraph(signed_pct(row["delta"]), S_TABLE_CELL_C),
        Paragraph(signed_pct(row["delta_trailing"]), S_TABLE_CELL_C),
    ])

slot_table = Table(slot_table_data, colWidths=slot_col_w, repeatRows=1)
slot_style = [
    ("BACKGROUND", (0,0), (-1,0), BLACK),
    ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, LIGHT_GREY]),
    ("ALIGN", (0,0), (-1,-1), "CENTER"), ("ALIGN", (0,1), (0,-1), "LEFT"),
    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
]
for i, (_, row) in enumerate(slot_perf.iterrows()):
    delta_bg = DELTA_GREEN if row["delta"] >= 0 else DELTA_RED
    try:
        delta_t_bg = DELTA_GREEN if float(row["delta_trailing"]) >= 0 else DELTA_RED
    except:
        delta_t_bg = LIGHT_GREY
    util_bg = LIGHT_BLUE if row["utilization"] >= 0.70 else (RED if row["utilization"] < 0.40 else None)
    slot_style.append(("BACKGROUND", (5, i+1), (5, i+1), delta_bg))
    slot_style.append(("BACKGROUND", (6, i+1), (6, i+1), delta_t_bg))
    if util_bg:
        slot_style.append(("BACKGROUND", (3, i+1), (3, i+1), util_bg))

slot_table.setStyle(TableStyle(slot_style))
story.append(slot_table)
story.append(Spacer(1, 5*mm))
story.append(hr())
story.append(Spacer(1, 3*mm))

# Heatmap
story.append(Paragraph("Schedule Heatmap", style("hm_title",
    fontName="Helvetica-Bold", fontSize=11, textColor=BLACK, leading=14, spaceAfter=3)))
story.append(Paragraph(
    "Average utilization by day and timeslot. Deep blue = 75%+. Light blue = 65-75%. White/orange = average. Red = below 45%. Grey = no classes.",
    S_CAPTION))
story.append(Spacer(1, 2*mm))

hm_label_w = content_w * 0.13
hm_cell_w = (content_w - hm_label_w) / len(DOW_ORDER)

hm_table_data = [[Paragraph("", S_HM_HEADER)] + [Paragraph(d, S_HM_HEADER) for d in DOW_LABELS]]

for slot in SLOT_ORDER:
    row_data = [Paragraph(slot, S_HM_LABEL)]
    for dow in DOW_ORDER:
        match = heatmap_data[(heatmap_data["dow"] == dow) & (heatmap_data["slot_time"] == slot)]
        if len(match) > 0:
            val = match["util"].values[0]
            label = f"{round(val * 100)}%"
            use_white = val < 0.45 or val >= 0.65
            cell_s = style(f"hm_{dow}_{slot}",
                fontName="Helvetica-Bold", fontSize=7,
                textColor=WHITE if use_white else BLACK,
                leading=9, alignment=TA_CENTER)
            row_data.append(Paragraph(label, cell_s))
        else:
            row_data.append(Paragraph("", S_HM_LABEL))
    hm_table_data.append(row_data)

hm_table = Table(hm_table_data, colWidths=[hm_label_w] + [hm_cell_w] * len(DOW_ORDER))
hm_style = [
    ("BACKGROUND", (0,0), (-1,0), BLACK),
    ("BACKGROUND", (0,1), (0,-1), LIGHT_GREY),
    ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ("ALIGN", (0,1), (0,-1), "RIGHT"),
    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ("RIGHTPADDING", (0,1), (0,-1), 4),
    ("GRID", (0,0), (-1,-1), 0.5, WHITE),
]
for r_idx, slot in enumerate(SLOT_ORDER):
    for c_idx, dow in enumerate(DOW_ORDER):
        match = heatmap_data[(heatmap_data["dow"] == dow) & (heatmap_data["slot_time"] == slot)]
        val = match["util"].values[0] if len(match) > 0 else None
        hm_style.append(("BACKGROUND", (c_idx+1, r_idx+1), (c_idx+1, r_idx+1), get_heatmap_color(val)))

hm_table.setStyle(TableStyle(hm_style))
story.append(hm_table)
story.append(Spacer(1, 2*mm))
story.append(Paragraph("Days run Monday through Sunday left to right.", S_CAPTION))
story.append(PageBreak())

# =============================
# PAGE 4 — INSTRUCTOR PERFORMANCE
# =============================
story.append(Paragraph("Instructor Performance", S_SECTION_TITLE))
story.append(Paragraph(
    "Ranked by Lift vs Slot. Positive lift = performing above the historical average for assigned timeslots. Lift vs 3M shows recent momentum vs trailing 3-month slot average.",
    S_SECTION_SUB))
story.append(hr(BLACK, 1))
story.append(Spacer(1, 3*mm))

instr_headers = ["Instructor", "Classes", "Riders", "Utilization", "Slot Baseline", "Lift vs Slot", "Lift vs 3M"]
instr_col_w = [content_w*0.20, content_w*0.08, content_w*0.09,
               content_w*0.12, content_w*0.13, content_w*0.19, content_w*0.19]

instr_table_data = [[Paragraph(h, S_TABLE_HEADER) for h in instr_headers]]
for _, row in instr_perf.iterrows():
    instr_table_data.append([
        Paragraph(str(row["report_instructor"]), S_TABLE_CELL),
        Paragraph(str(int(row["classes"])), S_TABLE_CELL_C),
        Paragraph(f"{int(row['riders']):,}", S_TABLE_CELL_C),
        Paragraph(pct(row["utilization"]), S_TABLE_CELL_C),
        Paragraph(pct(row["slot_baseline"]), S_TABLE_CELL_C),
        Paragraph(signed_pct(row["lift"]), S_TABLE_CELL_C),
        Paragraph(signed_pct(row["lift_trailing"]), S_TABLE_CELL_C),
    ])

instr_table = Table(instr_table_data, colWidths=instr_col_w, repeatRows=1)
instr_style = [
    ("BACKGROUND", (0,0), (-1,0), BLACK),
    ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, LIGHT_GREY]),
    ("ALIGN", (0,0), (-1,-1), "CENTER"), ("ALIGN", (0,1), (0,-1), "LEFT"),
    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
]
for i, (_, row) in enumerate(instr_perf.iterrows()):
    lift_bg = DELTA_GREEN if row["lift"] >= 0 else DELTA_RED
    try:
        lift_t_bg = DELTA_GREEN if float(row["lift_trailing"]) >= 0 else DELTA_RED
    except:
        lift_t_bg = LIGHT_GREY
    util_bg = LIGHT_BLUE if row["utilization"] >= 0.70 else (RED if row["utilization"] < 0.40 else None)
    instr_style.append(("BACKGROUND", (5, i+1), (5, i+1), lift_bg))
    instr_style.append(("BACKGROUND", (6, i+1), (6, i+1), lift_t_bg))
    if util_bg:
        instr_style.append(("BACKGROUND", (3, i+1), (3, i+1), util_bg))

instr_table.setStyle(TableStyle(instr_style))
story.append(instr_table)
story.append(Spacer(1, 4*mm))
story.append(Paragraph(
    "Blue utilization = at or above 70%. Red = below 40%. Green lift = above baseline. Red lift = below baseline.",
    S_CAPTION))
story.append(PageBreak())

# =============================
# PAGE 5 — TRENDS
# =============================
story.append(Paragraph("Monthly Trend", S_SECTION_TITLE))
story.append(Paragraph(
    "Studio utilization month by month since September 2025. Current report month highlighted in blue.",
    S_SECTION_SUB))
story.append(hr(BLACK, 1))
story.append(Spacer(1, 4*mm))

trend_headers = ["Month", "Classes", "Riders", "Utilization"]
trend_col_w = [content_w*0.35, content_w*0.20, content_w*0.20, content_w*0.25]

trend_table_data = [[Paragraph(h, S_TABLE_HEADER) for h in trend_headers]]
for _, row in trend_data_full.iterrows():
    is_current = row["month"] == current_month
    cs = style("tc", fontName="Helvetica-Bold", fontSize=8) if is_current else S_TABLE_CELL
    csc = style("tcc", fontName="Helvetica-Bold", fontSize=8, alignment=TA_CENTER) if is_current else S_TABLE_CELL_C
    trend_table_data.append([
        Paragraph(row["month_label"], cs),
        Paragraph(str(int(row["classes"])), csc),
        Paragraph(f"{int(row['riders']):,}", csc),
        Paragraph(pct(row["utilization"]), csc),
    ])

trend_table = Table(trend_table_data, colWidths=trend_col_w, repeatRows=1)
trend_style_list = [
    ("BACKGROUND", (0,0), (-1,0), BLACK),
    ("ALIGN", (0,0), (-1,-1), "CENTER"), ("ALIGN", (0,1), (0,-1), "LEFT"),
    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("TOPPADDING", (0,0), (-1,-1), 5), ("BOTTOMPADDING", (0,0), (-1,-1), 5),
]
for i, (_, row) in enumerate(trend_data_full.iterrows()):
    is_current = row["month"] == current_month
    bg = LIGHT_BLUE if is_current else (WHITE if i % 2 == 0 else LIGHT_GREY)
    util_bg = LIGHT_BLUE if row["utilization"] >= 0.70 else (RED if row["utilization"] < 0.40 else bg)
    trend_style_list.append(("BACKGROUND", (0, i+1), (-1, i+1), bg))
    trend_style_list.append(("BACKGROUND", (3, i+1), (3, i+1), util_bg))

trend_table.setStyle(TableStyle(trend_style_list))
story.append(trend_table)

# =============================
# BUILD
# =============================
doc = SimpleDocTemplate(
    OUT_PDF, pagesize=A4,
    leftMargin=MARGIN, rightMargin=MARGIN,
    topMargin=16*mm, bottomMargin=14*mm,
)
doc.build(story, onFirstPage=first_page, onLaterPages=later_pages)
print(f"PDF written: {OUT_PDF}")