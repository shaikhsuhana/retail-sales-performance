import numpy as np
import pandas as pd
from datetime import date, timedelta

np.random.seed(7)

regions = ["North", "South", "East", "West"]
region_base = {"North": 1.15, "South": 0.75, "East": 1.05, "West": 1.0}  # South underperforms
categories = ["Electronics", "Apparel", "Home & Kitchen", "Beauty", "Sports"]
cat_price = {"Electronics": 180, "Apparel": 45, "Home & Kitchen": 65, "Beauty": 28, "Sports": 55}
channels = ["Online", "In-Store", "Mobile App"]

start = date(2024, 10, 1)
days = 365
rows = []
order_id = 5000

for d in range(days):
    day = start + timedelta(days=d)
    month = day.month
    # seasonal multiplier: Nov-Dec holiday spike, summer dip
    season = 1.0
    if month in (11, 12):
        season = 1.55
    elif month in (6, 7):
        season = 0.85
    elif month in (1, 2):
        season = 0.8

    n_orders_today = np.random.poisson(14 * season)
    for _ in range(n_orders_today):
        region = np.random.choice(regions, p=[0.28, 0.20, 0.27, 0.25])
        category = np.random.choice(categories, p=[0.25, 0.22, 0.2, 0.18, 0.15])
        channel = np.random.choice(channels, p=[0.52, 0.30, 0.18])
        units = np.random.randint(1, 5)
        unit_price = cat_price[category] * np.random.normal(1, 0.12)
        revenue = round(unit_price * units * region_base[region], 2)
        cost = round(revenue * np.random.uniform(0.55, 0.68), 2)
        order_id += 1
        rows.append([order_id, day.isoformat(), region, category, channel, units, round(unit_price,2), revenue, cost])

df = pd.DataFrame(rows, columns=["order_id","date","region","category","channel","units","unit_price","revenue","cost"])
df.to_csv("/home/claude/projects/retail-sales/raw_sales_data.csv", index=False)
print(df.shape)
print(df.groupby("region")["revenue"].sum().sort_values(ascending=False))
print(df["date"].min(), df["date"].max())
