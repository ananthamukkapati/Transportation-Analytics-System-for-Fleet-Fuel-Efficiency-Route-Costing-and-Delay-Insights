"""
Transportation Analytics System
Step 2: ETL + Build Unified Master Analytics Table
"""
import pandas as pd
import numpy as np
import json

# ── Load all sources ──────────────────────────────────────────────────────────
vehicles  = pd.read_csv("data/vehicles.csv")
drivers   = pd.read_csv("data/drivers.csv")
routes    = pd.read_csv("data/route_logs.csv", parse_dates=["trip_date"])
fuel      = pd.read_excel("data/fuel_logs.xlsx")
with open("data/delivery_timelines.json") as f:
    delivery = pd.DataFrame(json.load(f))
maint     = pd.read_csv("data/maintenance_history.csv", parse_dates=["maint_date"])

print("=== DATA QUALITY REPORT (Pre-clean) ===")
for name, df in [("Vehicles",vehicles),("Drivers",drivers),("Routes",routes),("Fuel",fuel),("Delivery",delivery),("Maint",maint)]:
    print(f"  {name}: {df.shape[0]} rows, {df.isnull().sum().sum()} nulls")

# ── Data Cleaning ─────────────────────────────────────────────────────────────
# Fill missing categoricals with mode
routes["traffic_level"] = routes["traffic_level"].fillna(routes["traffic_level"].mode()[0])
routes["weather"] = routes["weather"].fillna(routes["weather"].mode()[0])

# Cap fuel outliers at 99th percentile
p99 = fuel["fuel_consumed_l"].quantile(0.99)
fuel["fuel_consumed_l"] = fuel["fuel_consumed_l"].clip(upper=p99)
fuel["fuel_efficiency_kml"] = routes.set_index("trip_id").loc[fuel["trip_id"], "distance_km"].values / fuel["fuel_consumed_l"]

# Remove duplicate trip IDs
routes = routes.drop_duplicates(subset="trip_id")
fuel   = fuel.drop_duplicates(subset="trip_id")

# ── Aggregate maintenance cost per vehicle ────────────────────────────────────
maint_agg = maint.groupby("vehicle_id").agg(
    total_maint_cost_inr=("maint_cost_inr","sum"),
    maint_count=("maint_type","count"),
    avg_downtime_h=("downtime_hours","mean"),
).reset_index()

# ── Build MASTER TABLE ────────────────────────────────────────────────────────
master = routes.copy()
master = master.merge(fuel,     on="trip_id",    how="left")
master = master.merge(delivery, on="trip_id",    how="left")
master = master.merge(vehicles, on="vehicle_id", how="left")
# Drop driver avg_speed to avoid collision; use route avg_speed_kmph_x
drivers2 = drivers.drop(columns=["avg_speed_kmph"])
master = master.merge(drivers2,  on="driver_id",  how="left")
master = master.merge(maint_agg,on="vehicle_id", how="left")
# Rename suffixed columns
master.rename(columns={"avg_speed_kmph_x":"avg_speed_kmph"}, inplace=True)

# ── Derived Features ──────────────────────────────────────────────────────────
# Toll + labour cost estimation
master["toll_cost_inr"]   = np.where(master["route_category"].isin(["Highway","Mixed"]),
                                      master["distance_km"] * np.random.uniform(1.5, 4, len(master)), 0)
master["labour_cost_inr"] = master["actual_duration_h"] * np.random.uniform(150, 300, len(master))
master["total_trip_cost_inr"] = (master["fuel_cost_inr"] +
                                  master["toll_cost_inr"] +
                                  master["labour_cost_inr"] +
                                  master["total_maint_cost_inr"].fillna(0) / master["maint_count"].fillna(1))
master["cost_per_km"]     = (master["total_trip_cost_inr"] / master["distance_km"]).round(2)

# Driver performance score (higher = better)
master["driver_perf_score"] = (
    (master["fuel_efficiency_kml"] / master["fuel_efficiency_kml"].max()) * 40 +
    (1 - master["delay_minutes"] / master["delay_minutes"].max()) * 40 +
    (master["safety_rating"] / 5) * 20
).round(2)

# Route difficulty tier
master["difficulty_tier"] = pd.cut(master["road_difficulty"],
                                    bins=[0,3,6,10],
                                    labels=["Easy","Moderate","Hard"])

# Month / Quarter
master["trip_month"] = master["trip_date"].dt.month
master["trip_quarter"] = master["trip_date"].dt.quarter

# Final column selection for master CSV
MASTER_COLS = [
    "trip_id","vehicle_id","driver_id","driver_name","trip_date","trip_month","trip_quarter",
    "route_name","route_category","distance_km",
    "fuel_consumed_l","fuel_efficiency_kml","fuel_cost_inr",
    "road_difficulty","difficulty_tier","traffic_level","weather",
    "delay_minutes","delivery_status","expected_duration_h","actual_duration_h",
    "total_maint_cost_inr","maint_count","avg_downtime_h",
    "toll_cost_inr","labour_cost_inr","total_trip_cost_inr","cost_per_km",
    "driver_perf_score","safety_rating","experience_years",
    "vehicle_type","fuel_type","year_mfg","base_km_per_l","capacity_kg",
    "customer_location","avg_speed_kmph",
]
master = master[MASTER_COLS].copy()
master.fillna({"total_maint_cost_inr":0,"maint_count":0,"avg_downtime_h":0}, inplace=True)

master.to_csv("data/master_analytics_table.csv", index=False)
print(f"\n✅ Master table built: {master.shape[0]} rows × {master.shape[1]} cols")
print(f"   Nulls remaining: {master.isnull().sum().sum()}")

# ── Quick Stats ───────────────────────────────────────────────────────────────
print("\n=== SUMMARY STATISTICS ===")
print(f"  Avg fuel efficiency : {master['fuel_efficiency_kml'].mean():.2f} km/l")
print(f"  Avg delay           : {master['delay_minutes'].mean():.1f} min")
print(f"  Avg cost per km     : ₹{master['cost_per_km'].mean():.2f}")
print(f"  On-time deliveries  : {(master['delivery_status']=='On Time').mean()*100:.1f}%")
