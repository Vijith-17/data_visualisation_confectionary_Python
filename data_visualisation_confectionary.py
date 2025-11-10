# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go

# 2. Load & Inspect Dataset
# -------------------------------
df = pd.read_excel("/content/Confectionary [4564].xlsx")

print("Initial Data Preview:")
print(df.head())
print(df.info())

# 3. Clean & Prepare Data
# -------------------------------

# Clean column names
df.columns = (
    df.columns.str.strip()
    .str.replace(" ", "_")
    .str.replace("(£)", "GBP")
    .str.replace(r"[\(\)]", "", regex=True)
)

# Convert Date column
df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")

# Remove duplicates / missing rows
df = df.drop_duplicates().dropna(subset=["Date"])

# Sort by date
df = df.sort_values("Date")

# Create derived columns
df["Year"] = df["Date"].dt.year
df["Month"] = df["Date"].dt.month_name()
df["Profit_Margin_%"] = (df["ProfitGBP"] / df["RevenueGBP"]) * 100

print("\nCleaned Data Sample:")
print(df.head())

# 4. Descriptive Statistics
# -------------------------------
summary = df.describe()
print("\nSummary Statistics:")
print(summary)

# 5. Analysis - Profit & Revenue Insights
# -------------------------------

# Mean profit margin by confectionary
avg_margin = (
    df.groupby("Confectionary")["Profit_Margin_%"]
    .mean()
    .sort_values(ascending=False)
)
print("\nAverage Profit Margin by Confectionary:")
print(avg_margin)

# Total revenue by region
region_revenue = (
    df.groupby("CountryUK")["RevenueGBP"]
    .sum()
    .sort_values(ascending=False)
)
print("\nTotal Revenue by Region:")
print(region_revenue)

# Identify peak revenue month
peak_month = (
    df.groupby(["Year", "Month"])["RevenueGBP"]
    .sum()
    .reset_index()
    .sort_values("RevenueGBP", ascending=False)
    .head(1)
)
print("\nPeak Sales Period:")
print(peak_month)

# 6. Static Visualisations (Matplotlib / Seaborn)
# -------------------------------

sns.set_theme(style="whitegrid", palette="muted")

# (a) Average profit margin by confectionary
plt.figure(figsize=(8,5))
sns.barplot(x=avg_margin.index, y=avg_margin.values, palette="Blues_d")
plt.title("Average Profit Margin by Confectionary")
plt.ylabel("Profit Margin (%)")
plt.xlabel("Confectionary Type")
plt.tight_layout()
plt.savefig("1_avg_profit_margin.png")
plt.show()

# (b) Revenue trend over time by confectionary
plt.figure(figsize=(10,6))
sns.lineplot(data=df, x="Date", y="RevenueGBP", hue="Confectionary", marker="o")
plt.title("Revenue Over Time by Confectionary")
plt.xlabel("Date")
plt.ylabel("Revenue (£)")
plt.legend(title="Confectionary")
plt.tight_layout()
plt.savefig("2_revenue_trend.png")
plt.show()

# (c) Heatmap: monthly revenue by region
pivot = df.pivot_table(
    values="RevenueGBP", index="Month", columns="CountryUK", aggfunc="sum"
)
plt.figure(figsize=(8,6))
sns.heatmap(pivot, cmap="YlOrBr", annot=False)
plt.title("Monthly Revenue Heatmap by Region")
plt.tight_layout()
plt.savefig("3_monthly_heatmap.png")
plt.show()

# (d) Scatter: Units sold vs Profit
plt.figure(figsize=(8,5))
sns.scatterplot(data=df, x="Units_Sold", y="ProfitGBP", hue="Confectionary", size="RevenueGBP", sizes=(20,200))
plt.title("Units Sold vs Profit by Confectionary")
plt.tight_layout()
plt.savefig("4_units_vs_profit.png")
plt.show()

# 7. Interactive Visualisations (Plotly)
# -------------------------------

# (a) Profit by Confectionary and Region
fig1 = px.bar(
    df,
    x="Confectionary",
    y="ProfitGBP",
    color="CountryUK",
    title="Profit by Confectionary and Region",
    barmode="group"
)
fig1.update_layout(xaxis_title="Confectionary", yaxis_title="Profit (£)")
fig1.write_html("interactive_profit_region.html")
fig1.show()

# (b) Revenue trend over time
fig2 = px.line(
    df,
    x="Date",
    y="RevenueGBP",
    color="Confectionary",
    title="Revenue Trend Over Time"
)
fig2.update_layout(xaxis_title="Date", yaxis_title="Revenue (£)")
fig2.write_html("interactive_revenue_trend.html")
fig2.show()

# (c) Profit Margin Distribution
fig3 = px.box(
    df,
    x="Confectionary",
    y="Profit_Margin_%",
    color="Confectionary",
    title="Profit Margin Distribution by Confectionary"
)
fig3.write_html("interactive_margin_distribution.html")
fig3.show()

# 8. Dashboard-Style Combined Visual (Optional)
# -------------------------------

# Combine key KPIs into one figure (optional)
kpi_fig = go.Figure()
for confection in df["Confectionary"].unique():
    subset = df[df["Confectionary"] == confection]
    kpi_fig.add_trace(
        go.Scatter(
            x=subset["Date"],
            y=subset["RevenueGBP"],
            mode="lines+markers",
            name=confection
        )
    )
kpi_fig.update_layout(
    title="Combined Revenue Trend Dashboard View",
    xaxis_title="Date",
    yaxis_title="Revenue (£)",
)
kpi_fig.write_html("interactive_dashboard.html")
kpi_fig.show()

# 9. Export Cleaned Data & Insights
# -------------------------------
df.to_csv("Cleaned_Confectionary.csv", index=False)
avg_margin.to_csv("Avg_Profit_Margin.csv")
region_revenue.to_csv("Region_Revenue.csv")

print("\nAll analysis completed. Charts and CSVs exported successfully.")
