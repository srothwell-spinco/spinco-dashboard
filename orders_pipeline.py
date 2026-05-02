import os
import glob
import pandas as pd
import numpy as np

# =============================
# CONFIG
# =============================
DATA_DIR = "data/incoming"
OUT_DIR  = "out"

# =============================
# PRODUCT CLASSIFICATION
# =============================
# Three analytical groups:
#   Credits      -- all credit purchases (existing rider behaviour)
#   Memberships  -- all recurring membership types including student
#   Intro Offers -- New Rider Special + 2-Week Unlimited (acquisition signal)
#   Other        -- Staff, Water, Accessories, Clothing (excluded from core)

def classify_product_group(row):
    pt   = str(row.get("Product Type", ""))
    prod = str(row.get("Product", ""))

    if pt == "Credits":
        return "Credits"

    if pt == "Memberships":
        if any(k in prod for k in ["New Rider", "2-Week"]):
            return "Intro Offers"
        if "Staff" in prod:
            return "Other"
        return "Memberships"

    # Water, Accessories, Clothing, Email Credit Gift Cards etc.
    return "Other"


# =============================
# LOAD ALL HOURLY ORDER FILES
# =============================
def load_orders(data_dir=DATA_DIR):
    """
    Loads all HourlyOrders_YYYY-MM.csv files from data/incoming/.

    Returns:
        orders_purchases : DataFrame  true volitional purchases (core analysis)
        orders_renewals  : DataFrame  system-processed auto-renewals (KPI only)

    Penalties are excluded entirely.
    System-processed membership renewals are separated -- they are billing
    events not purchase decisions, and distort demand-timing analysis.
    The midnight cluster (hour=0) in purchases is real but small (~12 orders
    in Sept); these are retained and visible in the heatmap.
    """
    files = sorted(glob.glob(os.path.join(data_dir, "HourlyOrders_*.csv")))
    if not files:
        print("WARNING: No HourlyOrders files found in data/incoming/")
        return pd.DataFrame(), pd.DataFrame()

    frames = []
    for f in files:
        print(f"  Orders: reading {os.path.basename(f)}")
        try:
            df_f = pd.read_csv(f)
            df_f.columns = df_f.columns.str.strip()
            frames.append(df_f)
        except Exception as e:
            print(f"  WARNING: Failed to read {f}: {e}")

    if not frames:
        return pd.DataFrame(), pd.DataFrame()

    raw = pd.concat(frames, ignore_index=True)
    print(f"  Orders raw rows loaded: {len(raw):,}")

    # Parse datetime
    raw["order_dt"] = pd.to_datetime(
        raw["Order Date (Local)"].astype(str).str.strip() + " " +
        raw["Order Time (Local)"].astype(str).str.strip(),
        format="%m/%d/%Y %I:%M %p",
        errors="coerce"
    )
    failed = raw["order_dt"].isna().sum()
    if failed > 0:
        print(f"  WARNING: {failed} order rows failed datetime parsing -- dropped")
    raw = raw[raw["order_dt"].notna()].copy()

    # Time fields
    raw["month"] = raw["order_dt"].dt.to_period("M").astype(str)
    raw["date"]  = raw["order_dt"].dt.date.astype(str)
    raw["hour"]  = raw["order_dt"].dt.hour
    raw["dow"]   = raw["order_dt"].dt.day_name()
    raw["week"]  = raw["order_dt"].dt.to_period("W").astype(str)

    # Numeric fields
    raw["Line Total"]    = pd.to_numeric(raw["Line Total"],    errors="coerce").fillna(0)
    raw["Line Subtotal"] = pd.to_numeric(raw["Line Subtotal"], errors="coerce").fillna(0)
    raw["Line Quantity"] = pd.to_numeric(raw["Line Quantity"], errors="coerce").fillna(1)

    # Completed only
    completed = raw[raw["Order Status"] == "Completed"].copy()

    # Define populations
    is_penalty = completed["Product Type"] == "Penalty Fees"
    is_renewal  = (
        (completed["Processed by System?"] == True) &
        (completed["Product Type"] == "Memberships")
    )

    orders_renewals  = completed[is_renewal].copy()
    orders_purchases = completed[~is_renewal & ~is_penalty].copy()

    # Classify
    orders_purchases["product_group"] = orders_purchases.apply(classify_product_group, axis=1)
    orders_purchases["is_core"]        = orders_purchases["product_group"] != "Other"

    core = orders_purchases[orders_purchases["is_core"]]
    breakdown = " | ".join(
        f"{g}: {n}" for g, n in core.groupby("product_group").size().items()
    )
    print(
        f"  True purchases: {len(orders_purchases):,} | "
        f"Renewals: {len(orders_renewals):,} | "
        f"Penalties: {is_penalty.sum():,} (excluded)"
    )
    print(f"  Core breakdown -- {breakdown}")

    return orders_purchases, orders_renewals


# =============================
# MONTHLY ORDERS SUMMARY
# =============================
def build_orders_summary(orders_purchases, orders_renewals):
    """
    Month-level summary for the monthly report. One row per month.
    Columns: totals + units/revenue split by Credits, Memberships, Intro Offers
    + renewal KPIs.
    """
    if orders_purchases.empty:
        return pd.DataFrame()

    GROUPS = ["Credits", "Memberships", "Intro Offers"]
    core   = orders_purchases[orders_purchases["is_core"]].copy()

    agg = (
        core.groupby(["month", "product_group"])
        .agg(units=("Line Quantity", "sum"), revenue=("Line Total", "sum"))
        .reset_index()
    )
    units_pivot = (
        agg.pivot(index="month", columns="product_group", values="units")
        .reindex(columns=GROUPS).fillna(0)
    )
    rev_pivot = (
        agg.pivot(index="month", columns="product_group", values="revenue")
        .reindex(columns=GROUPS).fillna(0)
    )
    units_pivot.columns = [f"units_{g.lower().replace(' ', '_')}" for g in GROUPS]
    rev_pivot.columns   = [f"rev_{g.lower().replace(' ', '_')}"   for g in GROUPS]

    totals = (
        core.groupby("month")
        .agg(
            total_orders  = ("Order Number", "nunique"),
            total_units   = ("Line Quantity", "sum"),
            total_revenue = ("Line Total", "sum"),
        )
        .reset_index()
    )

    if not orders_renewals.empty:
        ren = (
            orders_renewals.groupby("month")
            .agg(total_renewals=("Order Number", "nunique"))
            .reset_index()
        )
        ren["days_in_month"]      = ren["month"].apply(
            lambda m: pd.Period(m, freq="M").days_in_month
        )
        ren["avg_daily_renewals"] = (
            ren["total_renewals"] / ren["days_in_month"]
        ).round(1)
    else:
        ren = pd.DataFrame(columns=["month", "total_renewals", "avg_daily_renewals"])

    summary = (
        totals
        .merge(units_pivot.reset_index(), on="month", how="left")
        .merge(rev_pivot.reset_index(),   on="month", how="left")
        .merge(
            ren[["month", "total_renewals", "avg_daily_renewals"]],
            on="month", how="left"
        )
        .fillna(0)
        .sort_values("month")
        .reset_index(drop=True)
    )

    return summary


# =============================
# STANDALONE ENTRY POINT
# =============================
if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)
    print("Loading orders...")
    orders_purchases, orders_renewals = load_orders(DATA_DIR)

    if orders_purchases.empty:
        print("No orders data found -- skipping")
    else:
        orders_purchases.to_csv(os.path.join(OUT_DIR, "orders_purchases.csv"), index=False)
        print(f"orders_purchases.csv written: {len(orders_purchases):,} rows")

        summary = build_orders_summary(orders_purchases, orders_renewals)
        summary.to_csv(os.path.join(OUT_DIR, "orders_summary.csv"), index=False)
        print(f"orders_summary.csv written: {len(summary):,} rows")

        orders_renewals.to_csv(os.path.join(OUT_DIR, "orders_renewals.csv"), index=False)
        print(f"orders_renewals.csv written: {len(orders_renewals):,} rows")

    print("Orders pipeline complete.")
