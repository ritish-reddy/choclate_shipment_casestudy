"""
======================================================================
  CHOCOLATE SHIPMENTS ANALYSIS — STEP 2: TREND ANALYSIS & INSIGHTS
======================================================================
Project  : Global Chocolate Shipments Data Analysis
Author   : [Your Name]
Purpose  : Perform comprehensive trend analysis on cleaned shipment
           data to extract actionable business insights.

Insights generated:
  • Top-selling products (by revenue & volume)
  • Key import markets (geography & region)
  • Monthly/Quarterly revenue trends
  • Category-level performance
  • Pricing strategy recommendations
  • Demand forecasting (12-month rolling forecast)
  • Inventory optimisation signals
  • Market expansion opportunities

Resume context:
  Analyzed global shipment data using Python (data processing &
  trend analysis) to identify top-selling products and key import
  markets, enabling insights for pricing strategy, demand forecasting,
  inventory optimization, and market expansion.
======================================================================
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import warnings
import os

warnings.filterwarnings("ignore")

# ── Paths ──────────────────────────────────────────────────────────
DATA_PATH = "data/cleaned_shipments.xlsx"
OUT_DIR   = "outputs"
os.makedirs(OUT_DIR, exist_ok=True)

# ── Style ──────────────────────────────────────────────────────────
BROWN   = "#5C3317"
GOLD    = "#C8860A"
CREAM   = "#FFF8E7"
DARK    = "#2D1B00"
PALETTE = ["#5C3317","#C8860A","#A0522D","#D2691E","#8B4513",
           "#CD853F","#DEB887","#F4A460","#DAA520","#B8860B"]

plt.rcParams.update({
    "figure.facecolor" : CREAM,
    "axes.facecolor"   : CREAM,
    "axes.edgecolor"   : BROWN,
    "axes.labelcolor"  : DARK,
    "axes.titlecolor"  : BROWN,
    "text.color"       : DARK,
    "xtick.color"      : DARK,
    "ytick.color"      : DARK,
    "font.family"      : "DejaVu Sans",
    "axes.titlesize"   : 13,
    "axes.labelsize"   : 10,
})

def save_fig(name):
    path = os.path.join(OUT_DIR, f"{name}.png")
    plt.savefig(path, dpi=150, bbox_inches="tight", facecolor=CREAM)
    plt.close()
    print(f"  Chart saved → {path}")

def fmt_currency(val):
    if val >= 1_000_000:
        return f"${val/1_000_000:.2f}M"
    return f"${val/1_000:.1f}K"

# ======================================================================
# LOAD CLEAN DATA
# ======================================================================
print("=" * 60)
print("  CHOCOLATE SHIPMENTS — DATA ANALYSIS")
print("=" * 60)

df = pd.read_excel(DATA_PATH)
df["Shipdate"] = pd.to_datetime(df["Shipdate"])

# Only analyse Delivered shipments for revenue metrics
delivered = df[df["Order_Status"] == "Delivered"].copy()

print(f"  Total records        : {len(df):,}")
print(f"  Delivered shipments  : {len(delivered):,}")
print(f"  Date range           : {df['Shipdate'].min().date()} → {df['Shipdate'].max().date()}")
print(f"  Countries covered    : {df['Geo'].nunique()}")
print(f"  Unique products      : {df['Product'].nunique()}")
print(f"  Sales persons        : {df['Sales_person'].nunique()}")
print()

total_rev   = delivered["Amount"].sum()
total_boxes = delivered["Boxes"].sum()
avg_order   = delivered["Amount"].mean()

print(f"  Total Revenue        : {fmt_currency(total_rev)}")
print(f"  Total Boxes Shipped  : {total_boxes:,}")
print(f"  Avg Order Value      : {fmt_currency(avg_order)}")
print()

# ======================================================================
# ANALYSIS 1 — TOP SELLING PRODUCTS (Revenue & Volume)
# ======================================================================
print("─" * 60)
print("  [1/7]  TOP SELLING PRODUCTS")
print("─" * 60)

prod_rev = (delivered.groupby(["Product", "Category"])
            .agg(Total_Revenue=("Amount","sum"),
                 Total_Boxes=("Boxes","sum"),
                 Shipment_Count=("ShipmentID","count"))
            .reset_index()
            .sort_values("Total_Revenue", ascending=False))

prod_rev["Revenue_Share_%"] = (prod_rev["Total_Revenue"] / total_rev * 100).round(2)
prod_rev["Avg_Revenue_per_Shipment"] = (prod_rev["Total_Revenue"] / prod_rev["Shipment_Count"]).round(2)

print(prod_rev[["Product","Category","Total_Revenue","Total_Boxes","Revenue_Share_%"]].to_string(index=False))
print()

# Bar chart — Top 10 by Revenue
top10 = prod_rev.head(10)
fig, axes = plt.subplots(1, 2, figsize=(16, 6))
fig.suptitle("Top Selling Products — Revenue & Volume", fontsize=15, color=BROWN, fontweight="bold")

# Revenue
bars = axes[0].barh(top10["Product"][::-1], top10["Total_Revenue"][::-1]/1e6, color=PALETTE)
axes[0].set_xlabel("Total Revenue (Millions $)")
axes[0].set_title("By Revenue")
for bar, val in zip(bars, top10["Total_Revenue"][::-1]/1e6):
    axes[0].text(val + 0.01, bar.get_y() + bar.get_height()/2,
                 f"${val:.2f}M", va="center", fontsize=8)

# Volume (boxes)
bars2 = axes[1].barh(top10["Product"][::-1], top10["Total_Boxes"][::-1]/1000, color=PALETTE)
axes[1].set_xlabel("Total Boxes Shipped (Thousands)")
axes[1].set_title("By Volume (Boxes)")
for bar, val in zip(bars2, top10["Total_Boxes"][::-1]/1000):
    axes[1].text(val + 0.5, bar.get_y() + bar.get_height()/2,
                 f"{val:.1f}K", va="center", fontsize=8)

plt.tight_layout()
save_fig("01_top_products")

# ======================================================================
# ANALYSIS 2 — KEY IMPORT MARKETS
# ======================================================================
print("─" * 60)
print("  [2/7]  KEY IMPORT MARKETS")
print("─" * 60)

geo_rev = (delivered.groupby(["Geo", "Region"])
           .agg(Total_Revenue=("Amount","sum"),
                Total_Boxes=("Boxes","sum"),
                Shipment_Count=("ShipmentID","count"))
           .reset_index()
           .sort_values("Total_Revenue", ascending=False))

geo_rev["Market_Share_%"] = (geo_rev["Total_Revenue"] / total_rev * 100).round(2)

print(geo_rev.to_string(index=False))
print()

fig, axes = plt.subplots(1, 2, figsize=(14, 6))
fig.suptitle("Key Import Markets — Revenue Distribution", fontsize=15, color=BROWN, fontweight="bold")

# Pie chart
wedge_props = dict(width=0.5, edgecolor="white")
axes[0].pie(geo_rev["Total_Revenue"], labels=geo_rev["Geo"],
            autopct="%1.1f%%", colors=PALETTE[:len(geo_rev)],
            wedgeprops=wedge_props, startangle=90)
axes[0].set_title("Market Revenue Share (Donut)")

# Bar chart by region
reg_rev = delivered.groupby("Region")["Amount"].sum().sort_values(ascending=False)
axes[1].bar(reg_rev.index, reg_rev.values / 1e6, color=[BROWN, GOLD, "#A0522D"])
axes[1].set_xlabel("Region")
axes[1].set_ylabel("Revenue (Millions $)")
axes[1].set_title("Revenue by Region")
for i, (region, val) in enumerate(reg_rev.items()):
    axes[1].text(i, val/1e6 + 0.02, fmt_currency(val), ha="center", fontsize=9)

plt.tight_layout()
save_fig("02_import_markets")

# ======================================================================
# ANALYSIS 3 — MONTHLY & QUARTERLY REVENUE TREND
# ======================================================================
print("─" * 60)
print("  [3/7]  MONTHLY & QUARTERLY REVENUE TREND")
print("─" * 60)

monthly = (delivered.groupby(["Year","Month","Month_Name"])
           .agg(Revenue=("Amount","sum"), Boxes=("Boxes","sum"))
           .reset_index()
           .sort_values(["Year","Month"]))

monthly["Period"] = monthly["Month_Name"].str[:3] + " " + monthly["Year"].astype(str)

quarterly = (delivered.groupby(["Year","Quarter"])
             .agg(Revenue=("Amount","sum"), Boxes=("Boxes","sum"))
             .reset_index()
             .sort_values(["Year","Quarter"]))
quarterly["Label"] = quarterly["Quarter"] + " " + quarterly["Year"].astype(str)

print("Monthly Revenue (last 12 months):")
print(monthly.tail(12)[["Period","Revenue","Boxes"]].to_string(index=False))
print()

fig, axes = plt.subplots(2, 1, figsize=(16, 10))
fig.suptitle("Revenue Trend Analysis", fontsize=15, color=BROWN, fontweight="bold")

# Monthly trend
axes[0].plot(range(len(monthly)), monthly["Revenue"]/1e6,
             color=BROWN, linewidth=2, marker="o", markersize=4)
axes[0].fill_between(range(len(monthly)), monthly["Revenue"]/1e6, alpha=0.2, color=GOLD)
axes[0].set_xticks(range(len(monthly)))
axes[0].set_xticklabels(monthly["Period"], rotation=45, ha="right", fontsize=7)
axes[0].set_ylabel("Revenue (Millions $)")
axes[0].set_title("Monthly Revenue Trend")
axes[0].yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.1fM"))

# Quarterly trend
colors_q = [BROWN if "2023" in l else GOLD for l in quarterly["Label"]]
axes[1].bar(quarterly["Label"], quarterly["Revenue"]/1e6, color=colors_q, edgecolor="white")
axes[1].set_xlabel("Quarter")
axes[1].set_ylabel("Revenue (Millions $)")
axes[1].set_title("Quarterly Revenue Trend")
axes[1].tick_params(axis="x", rotation=30)
for i, val in enumerate(quarterly["Revenue"]):
    axes[1].text(i, val/1e6 + 0.05, fmt_currency(val),
                 ha="center", fontsize=7.5, color=DARK)

plt.tight_layout()
save_fig("03_revenue_trend")

# ======================================================================
# ANALYSIS 4 — CATEGORY PERFORMANCE
# ======================================================================
print("─" * 60)
print("  [4/7]  CATEGORY PERFORMANCE")
print("─" * 60)

cat_perf = (delivered.groupby("Category")
            .agg(Revenue=("Amount","sum"),
                 Boxes=("Boxes","sum"),
                 Orders=("ShipmentID","count"),
                 Avg_Order_Value=("Amount","mean"))
            .reset_index()
            .sort_values("Revenue", ascending=False))

cat_perf["Revenue_Share_%"] = (cat_perf["Revenue"] / total_rev * 100).round(2)
cat_perf["Avg_Order_Value"]  = cat_perf["Avg_Order_Value"].round(2)
print(cat_perf.to_string(index=False))
print()

fig, axes = plt.subplots(1, 3, figsize=(16, 6))
fig.suptitle("Category Performance Breakdown", fontsize=15, color=BROWN, fontweight="bold")

metrics = [("Revenue", "Revenue (Millions $)", 1e6), ("Boxes", "Boxes (Thousands)", 1e3), ("Orders", "Shipment Count", 1)]
for ax, (col, label, div) in zip(axes, metrics):
    vals = cat_perf[col] / div
    bars = ax.bar(cat_perf["Category"], vals, color=[BROWN, GOLD, "#A0522D"])
    ax.set_ylabel(label)
    ax.set_title(f"Category — {col}")
    for bar, v in zip(bars, vals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + vals.max()*0.01,
                f"{v:.1f}", ha="center", fontsize=9)

plt.tight_layout()
save_fig("04_category_performance")

# ======================================================================
# ANALYSIS 5 — PRICING STRATEGY INSIGHTS
# ======================================================================
print("─" * 60)
print("  [5/7]  PRICING STRATEGY INSIGHTS")
print("─" * 60)

pricing = delivered.copy()

pricing["Revenue_per_box"]  = pricing["Amount"] / pricing["Boxes"].replace(0, np.nan)
pricing["Profit_Margin_pct"] = ((pricing["Revenue_per_box"] - pricing["Cost_per_box"])
                                 / pricing["Revenue_per_box"] * 100).round(2)

price_summary = (pricing.groupby(["Product","Category"])
                 .agg(Cost_per_box=("Cost_per_box","first"),
                      Avg_Revenue_per_box=("Revenue_per_box","mean"),
                      Avg_Margin_pct=("Profit_Margin_pct","mean"),
                      Total_Revenue=("Amount","sum"))
                 .reset_index()
                 .sort_values("Avg_Margin_pct", ascending=False))

price_summary["Cost_per_box"]       = price_summary["Cost_per_box"].round(2)
price_summary["Avg_Revenue_per_box"] = price_summary["Avg_Revenue_per_box"].round(2)
price_summary["Avg_Margin_pct"]      = price_summary["Avg_Margin_pct"].round(1)

print("Pricing & Margin Analysis:")
print(price_summary[["Product","Category","Cost_per_box","Avg_Revenue_per_box","Avg_Margin_pct"]].to_string(index=False))
print()

# Scatter — cost vs revenue per box (bubble = total revenue)
fig, ax = plt.subplots(figsize=(12, 7))
for _, row in price_summary.iterrows():
    size = row["Total_Revenue"] / price_summary["Total_Revenue"].max() * 600
    color = BROWN if row["Category"] == "Bars" else GOLD if row["Category"] == "Bites" else "#A0522D"
    ax.scatter(row["Cost_per_box"], row["Avg_Revenue_per_box"], s=size, color=color, alpha=0.7, edgecolors="white")
    ax.annotate(row["Product"], (row["Cost_per_box"], row["Avg_Revenue_per_box"]),
                fontsize=7, ha="center", va="bottom")

ax.set_xlabel("Cost per Box ($)")
ax.set_ylabel("Avg Revenue per Box ($)")
ax.set_title("Pricing Strategy — Cost vs Revenue per Box\n(Bubble size = Total Revenue)", color=BROWN)
from matplotlib.lines import Line2D
legend_elements = [Line2D([0],[0], marker="o", color="w", markerfacecolor=BROWN, markersize=10, label="Bars"),
                   Line2D([0],[0], marker="o", color="w", markerfacecolor=GOLD,  markersize=10, label="Bites"),
                   Line2D([0],[0], marker="o", color="w", markerfacecolor="#A0522D", markersize=10, label="Other")]
ax.legend(handles=legend_elements, loc="upper left")
save_fig("05_pricing_strategy")

# ======================================================================
# ANALYSIS 6 — DEMAND FORECASTING (12-month Rolling Average)
# ======================================================================
print("─" * 60)
print("  [6/7]  DEMAND FORECASTING")
print("─" * 60)

# Build monthly series
monthly_demand = (delivered.groupby(["Year","Month"])
                  .agg(Revenue=("Amount","sum"), Boxes=("Boxes","sum"))
                  .reset_index()
                  .sort_values(["Year","Month"]))

monthly_demand["Date"] = pd.to_datetime(
    monthly_demand[["Year","Month"]].assign(day=1).rename(columns={"Year":"year","Month":"month"}))
monthly_demand.set_index("Date", inplace=True)

# 3-month moving average
monthly_demand["MA3_Revenue"] = monthly_demand["Revenue"].rolling(3).mean()
monthly_demand["MA6_Revenue"] = monthly_demand["Revenue"].rolling(6).mean()

# Simple trend extrapolation — linear regression on last 12 months
from numpy.polynomial import polynomial as P

last12 = monthly_demand.tail(12).copy()
x      = np.arange(len(last12))
coef   = np.polyfit(x, last12["Revenue"], 1)   # slope, intercept

# Forecast next 6 months
forecast_x      = np.arange(len(last12), len(last12) + 6)
forecast_vals   = np.polyval(coef, forecast_x)
forecast_dates  = pd.date_range(start=monthly_demand.index[-1] + pd.DateOffset(months=1), periods=6, freq="MS")

print("  Trend (linear fit on last 12 months):")
print(f"    Monthly revenue slope : ${coef[0]:,.0f}/month")
print(f"    6-month forecast total: {fmt_currency(forecast_vals.sum())}\n")
print("  6-Month Revenue Forecast:")
for d, v in zip(forecast_dates, forecast_vals):
    print(f"    {d.strftime('%b %Y')} → {fmt_currency(v)}")
print()

fig, ax = plt.subplots(figsize=(15, 6))
ax.plot(monthly_demand.index, monthly_demand["Revenue"]/1e6,
        label="Actual Revenue", color=BROWN, linewidth=2)
ax.plot(monthly_demand.index, monthly_demand["MA3_Revenue"]/1e6,
        label="3-Month MA", color=GOLD, linewidth=1.5, linestyle="--")
ax.plot(monthly_demand.index, monthly_demand["MA6_Revenue"]/1e6,
        label="6-Month MA", color="#A0522D", linewidth=1.5, linestyle=":")
ax.plot(forecast_dates, forecast_vals/1e6,
        label="6-Month Forecast (Linear)", color="red", linewidth=2, linestyle="--", marker="D", markersize=6)
ax.fill_between(forecast_dates, (forecast_vals*0.9)/1e6, (forecast_vals*1.1)/1e6,
                alpha=0.15, color="red", label="±10% Confidence Band")
ax.axvline(x=monthly_demand.index[-1], color="grey", linestyle=":", linewidth=1.5, label="Forecast start")
ax.set_title("Demand Forecasting — Monthly Revenue with 6-Month Projection", color=BROWN)
ax.set_ylabel("Revenue (Millions $)")
ax.set_xlabel("Date")
ax.legend(loc="upper left", fontsize=8)
ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.1fM"))
save_fig("06_demand_forecast")

# ======================================================================
# ANALYSIS 7 — INVENTORY OPTIMISATION
# ======================================================================
print("─" * 60)
print("  [7/7]  INVENTORY OPTIMISATION SIGNALS")
print("─" * 60)

# Monthly demand by product
monthly_prod = (delivered.groupby(["Product","Year","Month"])["Boxes"]
                .sum().reset_index())

inv_stats = (monthly_prod.groupby("Product")["Boxes"]
             .agg(Mean_Monthly_Demand="mean",
                  Std_Monthly_Demand="std",
                  Min_Monthly_Demand="min",
                  Max_Monthly_Demand="max")
             .reset_index())

# Safety stock = 1.65 * std (95% service level)
inv_stats["Safety_Stock_Boxes"]  = (inv_stats["Std_Monthly_Demand"] * 1.65).round(0).astype(int)
inv_stats["Reorder_Point_Boxes"] = (inv_stats["Mean_Monthly_Demand"] + inv_stats["Safety_Stock_Boxes"]).round(0).astype(int)
inv_stats["Max_Stock_Boxes"]     = (inv_stats["Mean_Monthly_Demand"] * 1.5 + inv_stats["Safety_Stock_Boxes"]).round(0).astype(int)
inv_stats = inv_stats.sort_values("Mean_Monthly_Demand", ascending=False)

print("Inventory Optimisation (95% Service Level):")
print(inv_stats.to_string(index=False))
print()

fig, ax = plt.subplots(figsize=(14, 8))
y_pos = range(len(inv_stats))
ax.barh(y_pos, inv_stats["Mean_Monthly_Demand"], color=GOLD, label="Avg Monthly Demand", alpha=0.8)
ax.barh(y_pos, inv_stats["Safety_Stock_Boxes"],
        left=inv_stats["Mean_Monthly_Demand"], color=BROWN, label="Safety Stock Buffer", alpha=0.8)
ax.set_yticks(y_pos)
ax.set_yticklabels(inv_stats["Product"], fontsize=8)
ax.set_xlabel("Boxes")
ax.set_title("Inventory Optimisation — Avg Monthly Demand + Safety Stock\n(Reorder when stock hits Reorder Point)", color=BROWN)
ax.legend()
save_fig("07_inventory_optimisation")

# ======================================================================
# SAVE SUMMARY REPORT (Excel)
# ======================================================================
print("─" * 60)
print("  SAVING SUMMARY REPORTS")
print("─" * 60)

summary_path = os.path.join(OUT_DIR, "analysis_summary.xlsx")
with pd.ExcelWriter(summary_path, engine="openpyxl") as writer:
    prod_rev.to_excel(writer, sheet_name="Top_Products",        index=False)
    geo_rev.to_excel(writer,  sheet_name="Import_Markets",      index=False)
    cat_perf.to_excel(writer, sheet_name="Category_Performance",index=False)
    price_summary.to_excel(writer, sheet_name="Pricing_Strategy", index=False)
    inv_stats.to_excel(writer, sheet_name="Inventory_Signals",  index=False)
    monthly_demand.reset_index().to_excel(writer, sheet_name="Monthly_Trend", index=False)

print(f"  Summary Excel → {summary_path}")

# ======================================================================
# FINAL INSIGHT SUMMARY
# ======================================================================
top_product   = prod_rev.iloc[0]["Product"]
top_country   = geo_rev.iloc[0]["Geo"]
top_category  = cat_perf.iloc[0]["Category"]
top_region    = reg_rev.idxmax()
best_margin   = price_summary.iloc[0]["Product"]
yoy_chg       = ((monthly_demand.loc[monthly_demand.index.year==2024,"Revenue"].sum() /
                   monthly_demand.loc[monthly_demand.index.year==2023,"Revenue"].sum()) - 1) * 100

print()
print("=" * 60)
print("  ★  KEY BUSINESS INSIGHTS")
print("=" * 60)
print(f"  1. Top revenue product    : {top_product}")
print(f"  2. Largest import market  : {top_country} ({geo_rev.iloc[0]['Market_Share_%']:.1f}% share)")
print(f"  3. Highest revenue region : {top_region}")
print(f"  4. Leading category       : {top_category} ({cat_perf.iloc[0]['Revenue_Share_%']:.1f}% of revenue)")
print(f"  5. Best-margin product    : {best_margin} ({price_summary.iloc[0]['Avg_Margin_pct']:.1f}% margin)")
print(f"  6. YoY Revenue Growth     : {yoy_chg:+.1f}% (2023 → 2024)")
print(f"  7. 6-Month Forecast Total : {fmt_currency(forecast_vals.sum())}")
print()
print("  Recommendations:")
print(f"  → Pricing  : Leverage {best_margin}'s high margin; review low-margin SKUs.")
print(f"  → Demand   : {'Upward' if coef[0] > 0 else 'Downward'} trend — {'scale up' if coef[0] > 0 else 'reassess'} inventory.")
print(f"  → Markets  : {top_country} drives the most revenue; expand presence in under-penetrated markets.")
print(f"  → Inventory: Safety stock model built — reorder points available in summary Excel.")
print("=" * 60)
print("  Analysis complete. All charts saved to /outputs folder.")
print("=" * 60)
