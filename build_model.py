import os
import pandas as pd
import numpy as np

# =============================
# CONFIG
# =============================
DATA_DIR = "data/incoming"
OUT_DIR = "out"
os.makedirs(OUT_DIR, exist_ok=True)

CAPACITY = 41
SEASON_START = "2025-09"

SLOT_BAND_MAP = {
    "10:00": "10:15",
    "11:15": "11:30",
    "17:15": "17:30",
    "18:00": "18:35",
    "18:30": "18:35",
    "18:33": "18:35",
    "18:34": "18:35",
    "18:35": "18:35",
    "19:40": "19:40",
    "19:43": "19:40",
    "19:44": "19:40",
    "19:45": "19:40",
}

CANONICAL_SLOTS = [
    "06:00","07:00","08:00","09:00","09:30","10:15","11:30",
    "16:45","17:00","17:30","18:35","19:40"
]

# =============================
# SLOT BANDING
# =============================
def band_slot(time_str):
    try:
        if time_str in SLOT_BAND_MAP:
            return SLOT_BAND_MAP[time_str]
        h, m = map(int, time_str.split(":"))
        total = h * 60 + m
        rounded = round(total / 15) * 15
        bh, bm = divmod(rounded, 60)
        result = f"{bh:02d}:{bm:02d}"
        return SLOT_BAND_MAP.get(result, result)
    except:
        return time_str

# =============================
# LOAD CLASS FILES
# =============================
class_files = sorted([
    f for f in os.listdir(DATA_DIR)
    if f.startswith("classes_") and f.endswith(".csv")
])

if not class_files:
    raise SystemExit("No class files found in data/incoming/. Expected files named classes_YYYY-MM.csv")

frames = []
for f in class_files:
    month = f.replace("classes_", "").replace(".csv", "")
    path = os.path.join(DATA_DIR, f)
    print(f"Reading: {path}")
    df_f = pd.read_csv(path)
    df_f["_file_month"] = month
    frames.append(df_f)

df = pd.concat(frames, ignore_index=True)
print(f"Total rows loaded: {len(df):,}")

# =============================
# VALIDATE COLUMNS
# =============================
required = [
    "Class Date", "Class Time", "Instructors",
    "Membership Checked In Reservations",
    "Credit Checked In Reservations",
    "Membership Penalty Cancelled Reservations",
    "Credit Penalty Cancelled Reservations",
    "Membership Penalty No Showed Reservations",
    "Credit Penalty No Showed Reservations",
]
missing = [c for c in required if c not in df.columns]
if missing:
    raise SystemExit(f"Missing required columns: {missing}")

# =============================
# CLEAN + PARSE
# =============================
dt_str = df["Class Date"].astype(str).str.strip() + " " + df["Class Time"].astype(str).str.strip()
df["start_dt"] = pd.to_datetime(dt_str, errors="coerce")
failed_mask = df["start_dt"].isna()
if failed_mask.sum() > 0:
    df.loc[failed_mask, "start_dt"] = pd.to_datetime(
        dt_str[failed_mask], errors="coerce", format="%m/%d/%Y %I:%M %p"
    )
still_failed = df["start_dt"].isna().sum()
if still_failed > 0:
    print(f"WARNING: {still_failed} rows failed date parsing and will be dropped")
else:
    print(f"Date parsing clean — all rows parsed successfully")

# Numeric conversions
def to_num(col):
    return pd.to_numeric(df[col].astype(str).str.replace(",",""), errors="coerce").fillna(0)

df["membership_checked_in"] = to_num("Membership Checked In Reservations")
df["credit_checked_in"] = to_num("Credit Checked In Reservations")
df["membership_late_cancel"] = to_num("Membership Penalty Cancelled Reservations")
df["credit_late_cancel"] = to_num("Credit Penalty Cancelled Reservations")
df["membership_no_show"] = to_num("Membership Penalty No Showed Reservations")
df["credit_no_show"] = to_num("Credit Penalty No Showed Reservations")

df["checked_in"] = df["membership_checked_in"] + df["credit_checked_in"]
df["late_cancel"] = df["membership_late_cancel"] + df["credit_late_cancel"]
df["no_show"] = df["membership_no_show"] + df["credit_no_show"]
df["total_dropout"] = df["late_cancel"] + df["no_show"]
df["capacity"] = CAPACITY

# Drop rows with no check-ins and no dropouts — truly empty classes
df = df.dropna(subset=["start_dt"])
df = df[df["checked_in"] > 0].copy()

# Membership mix
df["membership_rate"] = (df["membership_checked_in"] / df["checked_in"].replace(0, np.nan)).fillna(0).clip(0, 1)
df["credit_rate"] = (df["credit_checked_in"] / df["checked_in"].replace(0, np.nan)).fillna(0).clip(0, 1)

# Utilization
df["util"] = (df["checked_in"] / CAPACITY).clip(0, 1)

# =============================
# TIME FIELDS
# =============================
df["date"] = df["start_dt"].dt.date.astype(str)
df["month"] = df["start_dt"].dt.to_period("M").astype(str)
df["week"] = df["start_dt"].dt.to_period("W").astype(str)
df["dow"] = df["start_dt"].dt.day_name()
df["time"] = df["start_dt"].dt.strftime("%H:%M")
df["slot_time"] = df["time"].apply(band_slot)

def dow_group(d):
    if d in ["Monday", "Wednesday", "Friday"]:
        return "MWF"
    if d in ["Tuesday", "Thursday"]:
        return "TueThu"
    return "Weekend"

df["dow_group"] = df["dow"].apply(dow_group)
df["slot_key"] = df["dow_group"] + " | " + df["slot_time"]

# =============================
# INSTRUCTOR
# =============================
def clean_instructor(val):
    val = str(val).strip()
    for sep in ["/", ",", "&", "+", " and "]:
        if sep in val:
            return "TEAM TEACH"
    return val

df["report_instructor"] = df["Instructors"].apply(clean_instructor)

# Class type
df["class_kind"] = df["Class Type"].astype(str).str.lower().str.strip() if "Class Type" in df.columns else "standard"

# =============================
# SEASON FILTER
# =============================
df = df[df["month"] >= SEASON_START].copy()
print(f"Rows after season filter (>= {SEASON_START}): {len(df):,}")

# =============================
# CANONICAL SLOT FILTER
# =============================
dropped = df[~df["slot_time"].isin(CANONICAL_SLOTS)]["slot_time"].value_counts()
if len(dropped) > 0:
    print(f"WARNING: Dropping {dropped.sum()} rows with unrecognised slot times:")
    print(dropped.to_string())

df = df[df["slot_time"].isin(CANONICAL_SLOTS)].copy()
print(f"Rows after canonical slot filter: {len(df):,}")

# =============================
# DEDUPLICATE
# =============================
before = len(df)
df = df.drop_duplicates(subset=["start_dt", "report_instructor"], keep="first")
dropped_dupes = before - len(df)
if dropped_dupes > 0:
    print(f"WARNING: {dropped_dupes} duplicate rows removed")
else:
    print(f"Deduplication clean — no duplicates found")
print(f"Rows after deduplication: {len(df):,}")

# =============================
# DYNAMIC BASELINES PER MONTH
# =============================
months_sorted = sorted(df["month"].unique())
all_rows = []

for i, month in enumerate(months_sorted):
    available = months_sorted[:i+1]
    df_avail = df[df["month"].isin(available)]
    df_month = df[df["month"] == month].copy()

    # Slot baselines
    slot_baseline = df_avail.groupby("slot_key")["util"].mean().rename("slot_baseline_util")
    day_baseline = df_avail.groupby("dow")["util"].mean().rename("day_baseline_util")
    studio_baseline = df_avail["util"].mean()

    # Trailing 3-month baseline
    trailing = months_sorted[max(0, i-2):i+1]
    df_trail = df[df["month"].isin(trailing)]
    slot_trailing = df_trail.groupby("slot_key")["util"].mean().rename("slot_trailing_3m_util")

    # Previous month baseline
    if i >= 1:
        prev = months_sorted[i-1]
        slot_prev = df[df["month"] == prev].groupby("slot_key")["util"].mean().rename("slot_prev_month_util")
    else:
        slot_prev = pd.Series(dtype=float, name="slot_prev_month_util")

    # Dropout baselines
    slot_dropout_baseline = df_avail.groupby("slot_key")["total_dropout"].mean().rename("slot_dropout_baseline")

    # Join
    df_month = df_month.join(slot_baseline, on="slot_key")
    df_month = df_month.join(day_baseline, on="dow")
    df_month["studio_baseline_util"] = studio_baseline
    df_month = df_month.join(slot_trailing, on="slot_key")
    df_month = df_month.join(slot_prev, on="slot_key")
    df_month = df_month.join(slot_dropout_baseline, on="slot_key")

    # Deltas
    df_month["delta_vs_slot"] = df_month["util"] - df_month["slot_baseline_util"]
    df_month["delta_vs_day"] = df_month["util"] - df_month["day_baseline_util"]
    df_month["delta_vs_studio"] = df_month["util"] - df_month["studio_baseline_util"]
    df_month["delta_vs_trailing_3m"] = df_month["util"] - df_month["slot_trailing_3m_util"]
    df_month["delta_vs_prev_month"] = df_month["util"] - df_month["slot_prev_month_util"]

    # Rider metrics
    df_month["slot_expected_riders"] = df_month["slot_baseline_util"] * CAPACITY
    df_month["slot_lift_riders"] = df_month["checked_in"] - df_month["slot_expected_riders"]
    df_month["slot_lift_util"] = df_month["delta_vs_slot"]

    all_rows.append(df_month)

model = pd.concat(all_rows, ignore_index=True)

# =============================
# OUTPUT COLUMNS
# =============================
model_cols = [
    "start_dt", "date", "week", "month", "dow", "dow_group",
    "time", "slot_time", "slot_key", "report_instructor", "class_kind",
    "checked_in", "membership_checked_in", "credit_checked_in",
    "capacity", "util", "membership_rate", "credit_rate",
    "late_cancel", "no_show", "total_dropout", "slot_dropout_baseline",
    "slot_baseline_util", "slot_trailing_3m_util", "slot_prev_month_util",
    "day_baseline_util", "studio_baseline_util",
    "delta_vs_slot", "delta_vs_day", "delta_vs_studio",
    "delta_vs_trailing_3m", "delta_vs_prev_month",
    "slot_expected_riders", "slot_lift_riders", "slot_lift_util"
]

model = model[model_cols].sort_values("start_dt").reset_index(drop=True)
out_path = os.path.join(OUT_DIR, "model.csv")
model.to_csv(out_path, index=False)
print(f"model.csv written: {len(model):,} rows → {out_path}")