"""
======================================================================
  CHOCOLATE SHIPMENTS ANALYSIS — STEP 1: DATA CLEANING
======================================================================
Project  : Global Chocolate Shipments Data Analysis
Author   : [Your Name]
Purpose  : Clean raw shipment data before analysis.
           Handles duplicates, nulls, invalid values, and type issues.

Resume context:
  Analyzed global shipment data using Python (data processing &
  trend analysis) to identify top-selling products and key import
  markets, enabling insights for pricing strategy, demand forecasting,
  inventory optimization, and market expansion.
======================================================================
"""

import pandas as pd
import numpy as np
import os

# ── Paths ──────────────────────────────────────────────────────────
RAW_DIR    = "data"
OUTPUT_DIR = "data"
os.makedirs(OUTPUT_DIR, exist_ok=True)

VALID_STATUSES = {"Delivered", "Cancelled", "Pending"}

# ======================================================================
# 1. LOAD RAW DATA
# ======================================================================
print("=" * 60)
print("STEP 1 — Loading raw data")
print("=" * 60)

raw = pd.read_excel(os.path.join(RAW_DIR, "raw_shipments_data.xlsx"))
print(f"  Rows loaded        : {len(raw):,}")
print(f"  Columns            : {list(raw.columns)}")
print(f"  Null values (total): {raw.isnull().sum().sum():,}\n")

# Quick snapshot before cleaning
issues_log = {}

# ======================================================================
# 2. REMOVE DUPLICATE ROWS
# ======================================================================
print("STEP 2 — Removing duplicates")
before = len(raw)
raw.drop_duplicates(inplace=True)
raw.reset_index(drop=True, inplace=True)
removed_dups = before - len(raw)
issues_log["Duplicates removed"] = removed_dups
print(f"  Duplicates removed : {removed_dups}")
print(f"  Remaining rows     : {len(raw):,}\n")

# ======================================================================
# 3. FIX DATE COLUMN
# ======================================================================
print("STEP 3 — Standardising Shipdate column")
raw["Shipdate"] = pd.to_datetime(raw["Shipdate"], errors="coerce", dayfirst=False)
bad_dates = raw["Shipdate"].isna().sum()
issues_log["Unparseable dates"] = bad_dates
if bad_dates:
    print(f"  WARNING: {bad_dates} unparseable dates → dropped")
    raw.dropna(subset=["Shipdate"], inplace=True)
print(f"  Date range : {raw['Shipdate'].min().date()}  →  {raw['Shipdate'].max().date()}\n")

# ======================================================================
# 4. HANDLE MISSING NUMERIC VALUES
# ======================================================================
print("STEP 4 — Handling missing/invalid numeric values")

# Missing Amount — fill with median of same product
null_amount = raw["Amount"].isna().sum()
issues_log["Null Amount rows"] = null_amount
print(f"  Null Amount rows   : {null_amount}")
raw["Amount"] = raw.groupby("PID")["Amount"].transform(
    lambda x: x.fillna(x.median())
)

# Missing Boxes — fill with median of same product
null_boxes = raw["Boxes"].isna().sum()
issues_log["Null Boxes rows"] = null_boxes
print(f"  Null Boxes rows    : {null_boxes}")
raw["Boxes"] = raw.groupby("PID")["Boxes"].transform(
    lambda x: x.fillna(x.median())
)

# Negative amounts (data entry errors) — take absolute value
neg_amounts = (raw["Amount"] < 0).sum()
issues_log["Negative Amount rows"] = neg_amounts
print(f"  Negative Amount rows fixed: {neg_amounts}")
raw["Amount"] = raw["Amount"].abs()

# Cast types
raw["Amount"] = raw["Amount"].round(2)
raw["Boxes"]  = raw["Boxes"].fillna(0).astype(int)
print()

# ======================================================================
# 5. STANDARDISE ORDER STATUS
# ======================================================================
print("STEP 5 — Standardising Order_Status")

# Map common typos / variants
status_map = {
    "DELIVRD"  : "Delivered",
    "delivrd"  : "Delivered",
    "cancled"  : "Cancelled",
    "cancelled": "Cancelled",
    "CANCELLED": "Cancelled",
    "delivered": "Delivered",
    "DELIVERED": "Delivered",
    "pending"  : "Pending",
    "PENDING"  : "Pending",
    "Unknown"  : "Delivered",   # assume delivered if unknown
    ""         : "Delivered",
}
raw["Order_Status"] = raw["Order_Status"].str.strip().replace(status_map)

invalid_status = ~raw["Order_Status"].isin(VALID_STATUSES)
issues_log["Remaining invalid statuses"] = invalid_status.sum()
if invalid_status.sum():
    print(f"  {invalid_status.sum()} remaining unknown statuses → defaulted to 'Delivered'")
    raw.loc[invalid_status, "Order_Status"] = "Delivered"

status_dist = raw["Order_Status"].value_counts()
for s, cnt in status_dist.items():
    pct = cnt / len(raw) * 100
    print(f"    {s:<12}: {cnt:>6,}  ({pct:.1f}%)")
print()

# ======================================================================
# 6. ADD DERIVED COLUMNS
# ======================================================================
print("STEP 6 — Adding derived date columns")
raw["Year"]       = raw["Shipdate"].dt.year
raw["Month"]      = raw["Shipdate"].dt.month
raw["Month_Name"] = raw["Shipdate"].dt.strftime("%B")
raw["Quarter"]    = raw["Shipdate"].dt.quarter.apply(lambda q: f"Q{q}")
raw["WeekDay"]    = raw["Shipdate"].dt.day_name()
print("  Added: Year, Month, Month_Name, Quarter, WeekDay\n")

# ======================================================================
# 7. MERGE DIMENSION TABLES
# ======================================================================
print("STEP 7 — Merging dimension tables (Product, Geo, Salesperson)")

products = pd.read_excel(os.path.join(RAW_DIR, "products.xlsx"))
geo      = pd.read_excel(os.path.join(RAW_DIR, "geo.xlsx"))
sales    = pd.read_excel(os.path.join(RAW_DIR, "salespersons.xlsx"))

raw = raw.merge(products, on="PID",  how="left")
raw = raw.merge(geo,      on="GID",  how="left")
raw = raw.merge(sales,    on="SPID", how="left")

print(f"  Final columns: {list(raw.columns)}\n")

# ======================================================================
# 8. FINAL VALIDATION
# ======================================================================
print("STEP 8 — Final validation")
print(f"  Total rows (clean)  : {len(raw):,}")
print(f"  Remaining nulls     : {raw.isnull().sum().sum()}")
print(f"  Duplicate rows left : {raw.duplicated().sum()}")
print(f"  Amount range        : ₹{raw['Amount'].min():,.2f} — ₹{raw['Amount'].max():,.2f}")
print()

# ── Issues Summary ──────────────────────────────────────────────────
print("─" * 40)
print("  CLEANING SUMMARY")
print("─" * 40)
for issue, count in issues_log.items():
    print(f"  {issue:<35}: {count}")

# ======================================================================
# 9. SAVE CLEAN DATA
# ======================================================================
clean_path = os.path.join(OUTPUT_DIR, "cleaned_shipments.xlsx")
raw.to_excel(clean_path, index=False)
print(f"\n  ✓ Clean data saved → {clean_path}")
print("=" * 60)
print("  Data cleaning complete. Run 02_analysis.py next.")
print("=" * 60)
