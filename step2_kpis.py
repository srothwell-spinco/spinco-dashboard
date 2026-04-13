import os
import pandas as pd
import numpy as np

DATA_DIR = "data/incoming"

# Load first file (we'll do multi-file later)
files = [f for f in os.listdir(DATA_DIR) if f.lower().endswith((".csv", ".xlsx", ".xls"))]
if not files:
    raise SystemExit("No data files found in data/incoming")

path = os.path.join(DATA_DIR, sorted(files)[0])
print("Reading:", path)

df = pd.read_excel(path) if path.lower().endswith((".xlsx", ".xls")) else pd.read_csv(path)

# ---- Clean + types ----
# Standardize column names (keep originals too, but easier to reference)
# (Your schema is stable; we just reference directly)
required = [
    "Class Date", "Class Time", "Instructors",
    "Checked In Reservations", "Actual Capacity",
    "Class Type"
]
missing = [c for c in required if c not in df.columns]
if missing:
    raise SystemExit(f"Missing required columns: {missing}")

# Parse datetime from separate date + time columns
dt_str = df["Class Date"].astype(str).str.strip() + " " + df["Class Time"].astype(str).str.strip()
df["start_dt"] = pd.to_datetime(dt_str, errors="coerce")

# Numeric conversions
df["checked_in"] = pd.to_numeric(df["Checked In Reservations"], errors="coerce")
df["capacity"] = pd.to_numeric(df["Actual Capacity"], errors="coerce")

# Drop bad rows
df = df.dropna(subset=["start_dt", "checked_in", "capacity"])
df = df[df["capacity"] > 0]

# Derived fields
df["date"] = df["start_dt"].dt.date
df["month"] = df["start_dt"].dt.to_period("M").astype(str)
df["week"] = df["start_dt"].dt.to_period("W").astype(str)
df["dow"] = df["start_dt"].dt.day_name()
df["time"] = df["start_dt"].dt.strftime("%H:%M")
df["is_weekend"] = df["start_dt"].dt.dayofweek >= 5

# Utilization (our own)
df["util"] = (df["checked_in"] / df["capacity"]).clip(lower=0, upper=1)

# Optional: detect team teach (two names in instructors field)
df["is_team_teach"] = df["Instructors"].astype(str).str.contains(r"[\/,&+]| and ", regex=True)

# ---- KPIs ----
total_capacity = df["capacity"].sum()
total_checked_in = df["checked_in"].sum()
weighted_util = total_checked_in / total_capacity if total_capacity else np.nan

print("\n=== OVERALL KPIs ===")
print(f"Classes: {len(df):,}")
print(f"Total riders (checked in): {int(total_checked_in):,}")
print(f"Total capacity: {int(total_capacity):,}")
print(f"Weighted utilization (checked_in / capacity): {weighted_util:.1%}")
print(f"Team-teach classes (detected): {int(df['is_team_teach'].sum()):,}")

# ---- Breakdown 1: Instructor table ----
inst = (
    df.groupby("Instructors", dropna=False)
      .agg(classes=("util", "count"),
           riders=("checked_in", "sum"),
           capacity=("capacity", "sum"))
      .reset_index()
)
inst["weighted_util"] = inst["riders"] / inst["capacity"]
inst = inst.sort_values("weighted_util", ascending=False)

print("\n=== TOP 15 Instructors by Weighted Utilization ===")
print(inst.head(15).to_string(index=False))

# ---- Breakdown 2: Daypart buckets (Weekday/Weekend + morning/afternoon/evening) ----
def tod_bucket(t):
    # t is a datetime.time
    if t.hour < 9:
        return "Morning (pre-9)"
    if t.hour < 12:
        return "Late morning (9-12)"
    if t.hour < 17:
        return "Afternoon (12-5)"
    if t.hour < 20:
        return "Evening (5-8)"
    return "Late evening (8+)"

df["tod_bucket"] = df["start_dt"].dt.time.map(tod_bucket)
df["daypart"] = np.where(df["is_weekend"], "Weekend", "Weekday") + " " + df["tod_bucket"]

daypart = (
    df.groupby("daypart")
      .agg(classes=("util", "count"),
           riders=("checked_in", "sum"),
           capacity=("capacity", "sum"))
      .reset_index()
)
daypart["weighted_util"] = daypart["riders"] / daypart["capacity"]
daypart = daypart.sort_values("weighted_util", ascending=False)

print("\n=== Daypart (Weekday/Weekend buckets) ===")
print(daypart.to_string(index=False))

