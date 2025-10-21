import pandas as pd, numpy as np, matplotlib.pyplot as plt, os

# --- 1) Load Excel ---
xls = pd.ExcelFile("Marketing_Campaign_Performance_2025.xlsx")
df = xls.parse(xls.sheet_names[0])
df.columns = [c.replace("¬£", "£") for c in df.columns]  # fix weird pound sign if present

# --- 2) Clean & calculate ---
for c in ["Impressions","Clicks","Leads","Conversions"]:
    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
for c in ["Spend (£)","Revenue (£)"]:
    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

# Derived metrics
df["ROAS"] = np.where(df["Spend (£)"] > 0, df["Revenue (£)"] / df["Spend (£)"], np.nan)
df["CPA (£)"] = np.where(df["Conversions"] > 0, df["Spend (£)"] / df["Conversions"], np.nan)

# Month ordering
month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
df["Month"] = pd.Categorical(df["Month"], categories=month_order, ordered=True)

# --- 3) Summaries ---
monthly = df.groupby("Month", observed=True).agg({
    "Spend (£)":"sum",
    "Revenue (£)":"sum",
    "Conversions":"sum"
}).reset_index()

channel = df.groupby("Channel", observed=True).agg({
    "Spend (£)":"sum",
    "Revenue (£)":"sum",
    "Leads":"sum",
    "Conversions":"sum"
}).reset_index()

channel["ROAS"] = np.where(channel["Spend (£)"] > 0, channel["Revenue (£)"] / channel["Spend (£)"], np.nan)
channel["CPA (£)"] = np.where(channel["Conversions"] > 0, channel["Spend (£)"] / channel["Conversions"], np.nan)

# --- 4) Create folders ---
os.makedirs("images", exist_ok=True)
os.makedirs("data", exist_ok=True)

# --- 5) Charts (matplotlib; one chart per figure; no custom colors) ---
plt.figure()
plt.plot(monthly["Month"].astype(str), monthly["Spend (£)"], marker="o", label="Spend (£)")
plt.plot(monthly["Month"].astype(str), monthly["Revenue (£)"], marker="o", label="Revenue (£)")
plt.title("Monthly Spend vs Revenue (2025)")
plt.xlabel("Month")
plt.ylabel("£")
plt.legend()
plt.tight_layout()
plt.savefig("images/monthly_spend_vs_revenue.png", dpi=150)
plt.close()

plt.figure()
plt.bar(channel["Channel"], channel["ROAS"])
plt.title("ROAS by Channel")
plt.xlabel("Channel")
plt.ylabel("ROAS (Revenue/Spend)")
plt.xticks(rotation=20, ha="right")
plt.tight_layout()
plt.savefig("images/roas_by_channel.png", dpi=150)
plt.close()

# --- 6) Save clean data ---
monthly.to_csv("data/monthly_summary.csv", index=False)
channel.to_csv("data/channel_summary.csv", index=False)

# --- 7) README content (f-string) ---
readme = f"""# Marketing Campaign Dashboard 2025

A static, GitHub-friendly dashboard showing campaign efficiency metrics.

## Preview
![Monthly Spend vs Revenue](images/monthly_spend_vs_revenue.png)
![ROAS by Channel](images/roas_by_channel.png)

## Key Metrics
- **ROAS** = Revenue / Spend
- **CPA (£)** = Spend / Conversions

**Total Revenue:** £{monthly['Revenue (£)'].sum():,.0f}  
**Total Spend:** £{monthly['Spend (£)'].sum():,.0f}
"""

with open("README.md", "w", encoding="utf-8") as f:
    f.write(readme)

print("✅ Dashboard generated successfully!")
