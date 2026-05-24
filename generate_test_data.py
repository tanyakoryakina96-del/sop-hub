"""Generate realistic CPG test data for the S&OP Dashboard.

Produces 7 CSVs in ./test_data/ matching the canonical field names from
SPEC.md §4.1, so the Upload page's column mapper auto-detects them.

Shape:
  - 10 SKUs across 3 brands (Bluestar beverages / Crux snacks / Nordica dairy)
  - 5 channels, 4 regions, 3 plants
  - 16 months total: Apr 2025 -> Jul 2026 (cycle = May 2026 per SPEC.md)
      • Apr 2025 -> Apr 2026 (13 months) = HISTORY with actuals
      • May 2026 -> Jul 2026 (3 months)  = FORWARD PLANNING HORIZON
        demand has consensus + statistical only (actuals blank);
        supply and financial stay history-only (13 months) since the
        Scenario page reads price proxy from history and only needs
        demand consensus on the forward horizon.
  - Realistic seasonality, trend, forecast bias, and a mix of RAG performance
    so every dashboard (Demand / Supply / Financial / Scorecard) shows variety.

Run with:  py generate_test_data.py
"""
import csv
import random
from datetime import date
from pathlib import Path

random.seed(42)

OUT_DIR = Path(__file__).parent / "test_data"
OUT_DIR.mkdir(exist_ok=True)


# ---------------------------------------------------------------------------
# Master data
# ---------------------------------------------------------------------------
# SKU fields:
#   (code, name, brand, category, subcategory, uom, uom_to_cases,
#    base_monthly_vol_cases, price_eur, gm_pct, mape_target, bias_dir, plant)
# bias_dir: +1 = consensus tends to under-forecast (actuals > consensus)
#           -1 = consensus tends to over-forecast (actuals < consensus)
SKUS = [
    ("BST-001", "Bluestar Energy Original 250ml", "Bluestar", "Beverages", "Energy Drinks", "cases", 1.0, 10000, 18.0, 0.38, 0.07,  0, "PL-AMS"),
    ("BST-002", "Bluestar Energy Zero 250ml",     "Bluestar", "Beverages", "Energy Drinks", "cases", 1.0,  4000, 18.0, 0.40, 0.18, +1, "PL-AMS"),
    ("BST-003", "Bluestar Energy Original 500ml", "Bluestar", "Beverages", "Energy Drinks", "cases", 1.0,  2000, 30.0, 0.42, 0.25,  0, "PL-AMS"),
    ("CRX-001", "Crux Classic Salted 150g",       "Crux",     "Snacks",    "Chips",         "cases", 1.0,  6000, 12.0, 0.28, 0.08,  0, "PL-WAR"),
    ("CRX-002", "Crux Sour Cream 150g",           "Crux",     "Snacks",    "Chips",         "cases", 1.0,  3000, 13.0, 0.27, 0.12,  0, "PL-WAR"),
    ("CRX-003", "Crux BBQ 150g",                  "Crux",     "Snacks",    "Chips",         "cases", 1.0,  2500, 13.0, 0.27, 0.10,  0, "PL-WAR"),
    ("CRX-004", "Crux Multipack 6x30g",           "Crux",     "Snacks",    "Chips",         "cases", 1.0,  2500, 18.0, 0.30, 0.30,  0, "PL-WAR"),
    ("NOR-001", "Nordica Plain Yogurt 500g",      "Nordica",  "Dairy",     "Yogurt",        "cases", 1.0,  4500, 14.0, 0.22, 0.09,  0, "PL-MIL"),
    ("NOR-002", "Nordica Strawberry Yogurt 500g", "Nordica",  "Dairy",     "Yogurt",        "cases", 1.0,  3500, 15.0, 0.23, 0.11,  0, "PL-MIL"),
    ("NOR-003", "Nordica Greek Yogurt 200g",      "Nordica",  "Dairy",     "Yogurt",        "cases", 1.0,  7500, 25.0, 0.25, 0.15, -1, "PL-MIL"),
]

CHANNELS = [
    ("RET-MT", "Modern Trade",      "retail"),
    ("RET-TT", "Traditional Trade", "retail"),
    ("ECOM",   "E-commerce",        "ecomm"),
    ("FOOD",   "Foodservice",       "foodservice"),
    ("EXP",    "Export",            "export"),
]

REGIONS = [
    ("WEU", "Western Europe",         "NL", "EU-West"),
    ("NEU", "Northern Europe",        "SE", "EU-North"),
    ("SEU", "Southern Europe",        "IT", "EU-South"),
    ("CEE", "Central Eastern Europe", "PL", "EU-CEE"),
]

PLANTS = [
    ("PL-AMS", "Amsterdam Beverages Plant", "NL"),
    ("PL-WAR", "Warsaw Snacks Plant",       "PL"),
    ("PL-MIL", "Milan Dairy Plant",         "IT"),
]

# Where each SKU is sold: list of (channel_code, region_code, volume_share)
DEMAND_ALLOC = {
    "BST-001": [("RET-MT", "WEU", 0.50), ("ECOM",   "WEU", 0.20), ("FOOD", "WEU", 0.30)],
    "BST-002": [("RET-MT", "WEU", 0.70), ("ECOM",   "WEU", 0.30)],
    "BST-003": [("RET-MT", "WEU", 0.75), ("EXP",    "CEE", 0.25)],
    "CRX-001": [("RET-MT", "WEU", 0.50), ("RET-TT", "CEE", 0.30), ("ECOM", "WEU", 0.20)],
    "CRX-002": [("RET-MT", "WEU", 0.60), ("RET-TT", "CEE", 0.40)],
    "CRX-003": [("RET-MT", "WEU", 0.70), ("ECOM",   "WEU", 0.30)],
    "CRX-004": [("RET-MT", "WEU", 0.65), ("ECOM",   "WEU", 0.35)],
    "NOR-001": [("RET-MT", "WEU", 0.50), ("RET-MT", "SEU", 0.30), ("FOOD", "WEU", 0.20)],
    "NOR-002": [("RET-MT", "WEU", 0.70), ("RET-MT", "SEU", 0.30)],
    "NOR-003": [("RET-MT", "WEU", 0.50), ("ECOM",   "WEU", 0.20), ("FOOD", "WEU", 0.30)],
}

# YoY trend (cumulative over 12 months) applied to volume
TREND_YOY = {
    "BST-001":  0.00, "BST-002":  0.25, "BST-003": -0.05,
    "CRX-001":  0.02, "CRX-002":  0.00, "CRX-003":  0.05, "CRX-004":  0.10,
    "NOR-001": -0.03, "NOR-002":  0.00, "NOR-003":  0.15,
}

# Seasonal multipliers (1.0 = average) by month-of-year, keyed by category
SEASONALITY = {
    "Beverages": {1:0.80, 2:0.85, 3:0.95, 4:1.00, 5:1.15, 6:1.30, 7:1.40, 8:1.35, 9:1.15, 10:1.00, 11:0.85, 12:0.95},
    "Snacks":    {1:0.85, 2:0.95, 3:1.00, 4:1.00, 5:1.00, 6:1.00, 7:1.05, 8:1.05, 9:1.00, 10:1.00, 11:1.10, 12:1.30},
    "Dairy":     {1:0.95, 2:0.95, 3:1.05, 4:1.10, 5:1.15, 6:1.15, 7:1.20, 8:1.10, 9:1.00, 10:0.95, 11:0.90, 12:0.95},
}

# Revenue actuals vs budget delta (annual avg) — drives the Financial Review RAG mix
BUDGET_PERF = {
    "BST-001":  0.02,   # slight over-perform
    "BST-002":  0.20,   # well over (growing zero)
    "BST-003": -0.10,   # under (declining premium)
    "CRX-001":  0.00,
    "CRX-002": -0.04,
    "CRX-003":  0.01,
    "CRX-004": -0.12,   # promo miss
    "NOR-001":  0.05,
    "NOR-002":  0.00,
    "NOR-003": -0.08,
}

# Supply targets per SKU: (DOS days, fill_rate %, production_adherence %)
SUPPLY_PERF = {
    "BST-001": (30, 98, 96),
    "BST-002": (22, 92, 88),   # under-stocked & late on growing demand
    "BST-003": (50, 99, 98),   # excess inventory on declining SKU
    "CRX-001": (28, 97, 95),
    "CRX-002": (32, 96, 94),
    "CRX-003": (30, 98, 96),
    "CRX-004": (25, 94, 90),   # promo planning gaps
    "NOR-001": (15, 99, 97),   # dairy: short shelf-life => low DOS by design
    "NOR-002": (14, 98, 96),
    "NOR-003": (12, 97, 95),
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def months_range(n: int = 16):
    """Sequence of n first-of-month dates starting Apr 2025.

    Default 16 covers Apr 2025 -> Jul 2026: 13 months of history (through Apr
    2026) plus a 3-month forward planning horizon (May / Jun / Jul 2026) that
    only the Demand fact carries.
    """
    out = []
    y, m = 2025, 4
    for _ in range(n):
        out.append(date(y, m, 1))
        m += 1
        if m == 13:
            m, y = 1, y + 1
    return out


# Cycle close = last month with actuals. Anything strictly after is forward.
ACTUALS_THROUGH = date(2026, 4, 1)


def write_csv(path, header, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


# ---------------------------------------------------------------------------
# Dimension files
# ---------------------------------------------------------------------------
write_csv(
    OUT_DIR / "dim_sku.csv",
    ["sku_code", "sku_name", "brand", "category", "subcategory", "uom", "uom_to_cases"],
    [(s[0], s[1], s[2], s[3], s[4], s[5], s[6]) for s in SKUS],
)
write_csv(OUT_DIR / "dim_channel.csv", ["channel_code", "channel_name", "channel_type"], CHANNELS)
write_csv(OUT_DIR / "dim_region.csv",  ["region_code", "region_name", "country", "cluster"], REGIONS)
write_csv(OUT_DIR / "dim_plant.csv",   ["plant_code", "plant_name", "country"], PLANTS)


# ---------------------------------------------------------------------------
# Fact: Demand
# ---------------------------------------------------------------------------
months = months_range()
history_months = [p for p in months if p <= ACTUALS_THROUGH]
demand_rows = []

# caches needed for downstream supply + financial. Only history populates these;
# forward months contribute consensus to fact_demand but never to supply/financial.
sku_channel_month_actual = {}   # (sku, channel, period) -> volume actual (region-summed)
sku_channel_month_cons   = {}   # (sku, channel, period) -> volume consensus (region-summed)

for sku in SKUS:
    code, name, brand, cat, _sub, _uom, _conv, base_vol, _price, _gm, mape_t, bias_dir, _plant = sku
    seasonal = SEASONALITY[cat]
    yoy = TREND_YOY[code]
    m_trend = (1 + yoy) ** (1 / 12) - 1

    for i, period in enumerate(months):
        is_forward = period > ACTUALS_THROUGH
        true_demand = base_vol * ((1 + m_trend) ** i) * seasonal[period.month]

        stat = true_demand * random.uniform(0.95, 1.05)

        if code == "BST-002":
            cons = stat * random.uniform(1.05, 1.12)  # commercial sees growth
        elif code == "NOR-003":
            cons = stat * random.uniform(1.03, 1.08)
        elif code == "CRX-004":
            cons = stat * random.uniform(0.90, 1.15)  # noisy promo overlay
        else:
            cons = stat * random.uniform(0.97, 1.03)

        if is_forward:
            actual = None     # planning horizon: no actuals yet
        else:
            bias = bias_dir * mape_t * 0.5
            actual = cons * (1 + random.gauss(bias, mape_t))
            actual = max(actual, cons * 0.5)

        for ch, rg, share in DEMAND_ALLOC[code]:
            r_stat   = round(stat   * share, 1)
            r_cons   = round(cons   * share, 1)
            r_actual = "" if actual is None else round(actual * share, 1)
            demand_rows.append([period.isoformat(), code, ch, rg, r_stat, r_cons, r_actual])

            if actual is not None:
                sku_channel_month_actual[(code, ch, period)] = sku_channel_month_actual.get((code, ch, period), 0) + r_actual
            sku_channel_month_cons[(code, ch, period)]   = sku_channel_month_cons.get((code, ch, period), 0) + r_cons

write_csv(
    OUT_DIR / "fact_demand.csv",
    ["period_date", "sku_code", "channel_code", "region_code",
     "statistical_fcst", "consensus_fcst", "actuals"],
    demand_rows,
)


# ---------------------------------------------------------------------------
# Fact: Supply
# ---------------------------------------------------------------------------
supply_rows = []
for sku in SKUS:
    code, *_rest, plant = sku
    dos_target, fr_target, adh_target = SUPPLY_PERF[code]

    for period in history_months:
        total_cons = sum(v for (s, _c, p), v in sku_channel_month_cons.items() if s == code and p == period)
        if total_cons == 0:
            continue

        inventory   = round(total_cons * (dos_target / 30.44) * random.uniform(0.92, 1.08), 1)
        prod_plan   = round(total_cons * random.uniform(0.97, 1.03), 1)
        prod_actual = round(prod_plan * (adh_target / 100) * random.uniform(0.97, 1.03), 1)
        capacity    = round(prod_plan * random.uniform(1.10, 1.30), 1)
        orders_req  = round(total_cons * random.uniform(0.95, 1.05), 1)
        orders_del  = min(orders_req, round(orders_req * (fr_target / 100) * random.uniform(0.98, 1.02), 1))

        supply_rows.append([
            period.isoformat(), code, plant,
            inventory, prod_plan, prod_actual, capacity, orders_req, orders_del,
        ])

write_csv(
    OUT_DIR / "fact_supply.csv",
    ["period_date", "sku_code", "plant_code",
     "inventory_qty", "production_plan", "production_actual",
     "capacity_plan", "orders_requested", "orders_delivered"],
    supply_rows,
)


# ---------------------------------------------------------------------------
# Fact: Financial
# ---------------------------------------------------------------------------
fin_rows = []

# collapse demand (sku, channel) across regions
agg_actual = {}
for (s, c, p), v in sku_channel_month_actual.items():
    agg_actual[(s, c, p)] = agg_actual.get((s, c, p), 0) + v

sku_price = {s[0]: s[8] for s in SKUS}
sku_gm    = {s[0]: s[9] for s in SKUS}

for (sku, ch, period), actual_vol in sorted(agg_actual.items()):
    price  = sku_price[sku]
    gm_pct = sku_gm[sku]
    perf   = BUDGET_PERF[sku]

    revenue_actual = round(actual_vol * price, 2)
    revenue_budget = round(revenue_actual / (1 + perf), 2)
    revenue_le     = round((revenue_actual + revenue_budget) / 2, 2)

    gm_actual = round(revenue_actual * gm_pct * random.uniform(0.95, 1.05), 2)
    gm_budget = round(revenue_budget * gm_pct, 2)

    promo_actual = round(revenue_actual * random.uniform(0.02, 0.05), 2)
    promo_budget = round(revenue_budget * 0.03, 2)

    fin_rows.append([
        period.isoformat(), sku, ch,
        revenue_actual, revenue_budget, revenue_le,
        gm_actual, gm_budget, promo_actual, promo_budget, "EUR",
    ])

write_csv(
    OUT_DIR / "fact_financial.csv",
    ["period_date", "sku_code", "channel_code",
     "revenue_actual", "revenue_budget", "revenue_le",
     "gm_actual", "gm_budget", "promo_spend_actual", "promo_spend_budget", "currency_code"],
    fin_rows,
)


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
print(f"Generated CSVs in {OUT_DIR}:")
for f in sorted(OUT_DIR.iterdir()):
    with open(f, encoding="utf-8") as fh:
        n = sum(1 for _ in fh) - 1
    print(f"  {f.name:24s}  {n:5d} data rows")
