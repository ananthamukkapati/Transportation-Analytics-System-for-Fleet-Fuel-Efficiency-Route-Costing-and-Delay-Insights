"""
Transportation Analytics System
Step 1: Generate synthetic multi-source datasets
"""
import pandas as pd
import numpy as np
import json
import os

np.random.seed(42)
os.makedirs("data", exist_ok=True)
os.makedirs("outputs", exist_ok=True)
os.makedirs("charts", exist_ok=True)

N_VEHICLES = 20
N_DRIVERS = 25
N_TRIPS = 500

# ── 1. Vehicles ──────────────────────────────────────────────────────────────
vehicle_types = ["Truck", "Van", "Sedan", "SUV", "Bus"]
fuel_types    = ["Diesel", "Petrol", "CNG", "Electric"]
vehicles = pd.DataFrame({
    "vehicle_id":   [f"V{i:03d}" for i in range(1, N_VEHICLES+1)],
    "vehicle_type": np.random.choice(vehicle_types, N_VEHICLES),
    "fuel_type":    np.random.choice(fuel_types, N_VEHICLES, p=[0.45,0.35,0.12,0.08]),
    "capacity_kg":  np.random.randint(500, 10000, N_VEHICLES),
    "year_mfg":     np.random.randint(2010, 2024, N_VEHICLES),
    "base_km_per_l":np.round(np.random.uniform(6, 18, N_VEHICLES), 2),
})
vehicles.to_csv("data/vehicles.csv", index=False)

# ── 2. Drivers ────────────────────────────────────────────────────────────────
drivers = pd.DataFrame({
    "driver_id":         [f"D{i:03d}" for i in range(1, N_DRIVERS+1)],
    "driver_name":       [f"Driver_{i}" for i in range(1, N_DRIVERS+1)],
    "experience_years":  np.random.randint(1, 20, N_DRIVERS),
    "license_class":     np.random.choice(["A","B","C"], N_DRIVERS),
    "safety_rating":     np.round(np.random.uniform(3.0, 5.0, N_DRIVERS), 1),
    "avg_speed_kmph":    np.random.randint(45, 90, N_DRIVERS),
})
drivers.to_csv("data/drivers.csv", index=False)

# ── 3. GPS / Route logs ───────────────────────────────────────────────────────
routes = ["Route_City_A", "Route_City_B", "Route_Highway_1",
          "Route_Highway_2", "Route_Rural_X", "Route_Rural_Y", "Route_Mixed"]
categories = {"Route_City_A":"City","Route_City_B":"City",
               "Route_Highway_1":"Highway","Route_Highway_2":"Highway",
               "Route_Rural_X":"Rural","Route_Rural_Y":"Rural","Route_Mixed":"Mixed"}
traffic_levels = ["Low","Medium","High","Very High"]
weather_conds  = ["Clear","Rain","Fog","Storm","Hot"]

route_logs = pd.DataFrame({
    "trip_id":        [f"T{i:04d}" for i in range(1, N_TRIPS+1)],
    "vehicle_id":     np.random.choice(vehicles["vehicle_id"], N_TRIPS),
    "driver_id":      np.random.choice(drivers["driver_id"],  N_TRIPS),
    "route_name":     np.random.choice(routes, N_TRIPS),
    "trip_date":      pd.date_range("2023-01-01", periods=N_TRIPS, freq="16h"),
    "distance_km":    np.round(np.random.uniform(20, 400, N_TRIPS), 1),
    "traffic_level":  np.random.choice(traffic_levels, N_TRIPS, p=[0.2,0.4,0.3,0.1]),
    "weather":        np.random.choice(weather_conds,  N_TRIPS, p=[0.5,0.2,0.1,0.05,0.15]),
    "road_difficulty":np.round(np.random.uniform(1, 10, N_TRIPS), 1),
    "avg_speed_kmph": np.round(np.random.uniform(30, 100, N_TRIPS), 1),
})
route_logs["route_category"] = route_logs["route_name"].map(categories)
# Inject ~5% missing values
for col in ["traffic_level","weather"]:
    mask = np.random.random(N_TRIPS) < 0.05
    route_logs.loc[mask, col] = np.nan
route_logs.to_csv("data/route_logs.csv", index=False)

# ── 4. Fuel logs ──────────────────────────────────────────────────────────────
base_fuel_map = vehicles.set_index("vehicle_id")["base_km_per_l"].to_dict()
fuel_logs_rows = []
for _, row in route_logs.iterrows():
    base_eff = base_fuel_map.get(row["vehicle_id"], 10)
    traffic_pen = {"Low":0,"Medium":-0.5,"High":-1.5,"Very High":-3}.get(row["traffic_level"] or "Medium", -0.5)
    weather_pen = {"Clear":0,"Rain":-0.8,"Fog":-0.5,"Storm":-2,"Hot":-0.3}.get(row["weather"] or "Clear", 0)
    efficiency  = max(3, base_eff + traffic_pen + weather_pen + np.random.normal(0, 0.5))
    fuel_consumed = round(row["distance_km"] / efficiency, 2)
    fuel_cost     = round(fuel_consumed * np.random.uniform(85, 110), 2)
    fuel_logs_rows.append({
        "trip_id":       row["trip_id"],
        "fuel_consumed_l": fuel_consumed,
        "fuel_efficiency_kml": round(row["distance_km"] / fuel_consumed, 3),
        "fuel_cost_inr": fuel_cost,
        "refuel_count":  np.random.randint(0, 3),
    })
fuel_logs = pd.DataFrame(fuel_logs_rows)
# Inject outliers
outlier_idx = np.random.choice(len(fuel_logs), 10, replace=False)
fuel_logs.loc[outlier_idx, "fuel_consumed_l"] *= np.random.uniform(1.8, 2.5, 10)
fuel_logs.to_excel("data/fuel_logs.xlsx", index=False)

# ── 5. Delivery timelines ─────────────────────────────────────────────────────
delivery_rows = []
for _, row in route_logs.iterrows():
    expected_hrs   = row["distance_km"] / 60
    traffic_delay  = {"Low":0,"Medium":10,"High":30,"Very High":60}.get(row["traffic_level"] or "Medium", 10)
    weather_delay  = {"Clear":0,"Rain":20,"Fog":15,"Storm":60,"Hot":5}.get(row["weather"] or "Clear", 0)
    random_delay   = int(np.random.exponential(15))
    total_delay    = traffic_delay + weather_delay + random_delay
    # Occasionally no delay
    if np.random.random() < 0.25: total_delay = 0
    delivery_rows.append({
        "trip_id":            row["trip_id"],
        "expected_duration_h":round(expected_hrs, 2),
        "actual_duration_h":  round(expected_hrs + total_delay/60, 2),
        "delay_minutes":      total_delay,
        "delivery_status":    "On Time" if total_delay == 0 else ("Minor Delay" if total_delay < 30 else "Major Delay"),
        "customer_location":  np.random.choice(["North Zone","South Zone","East Zone","West Zone","Central"]),
    })
delivery_df = pd.DataFrame(delivery_rows)
delivery_df.to_json("data/delivery_timelines.json", orient="records", indent=2)

# ── 6. Maintenance history ────────────────────────────────────────────────────
maint_types = ["Oil Change","Tire Rotation","Brake Service","Engine Check","Full Service"]
maint_rows = []
for vid in vehicles["vehicle_id"]:
    n_maint = np.random.randint(1, 6)
    for _ in range(n_maint):
        maint_rows.append({
            "vehicle_id":    vid,
            "maint_date":    pd.Timestamp("2023-01-01") + pd.Timedelta(days=int(np.random.randint(0,365))),
            "maint_type":    np.random.choice(maint_types),
            "maint_cost_inr":round(np.random.uniform(500, 25000), 2),
            "downtime_hours":round(np.random.uniform(1, 48), 1),
        })
maint_df = pd.DataFrame(maint_rows)
maint_df.to_csv("data/maintenance_history.csv", index=False)

print("✅ All source datasets generated successfully.")
print(f"   Vehicles: {len(vehicles)} | Drivers: {len(drivers)} | Trips: {N_TRIPS}")
print(f"   Maintenance records: {len(maint_df)}")
