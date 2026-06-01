#!/usr/bin/env python
# coding: utf-8

# # ABC Retail Solutions – Retail Sales Data Processing & Business Insights
# 
# This notebook contains the complete data engineering pipeline developed for the retail sales use case.
# 
# The objective of this work is to read the retail transaction data from Excel, clean and standardize the records, protect sensitive customer information, generate business KPIs, and prepare the final dataset for Power BI reporting.
# 
# The notebook is written step-by-step to clearly show the entire workflow from raw data to final analytics output.

# In[1]:


import pandas as pd
import numpy as np
import hashlib
import logging
import warnings
import os
from datetime import datetime

warnings.filterwarnings("ignore")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("retail_pipeline")


# ## 1. Configuration
# 
# In this section, I am setting up the source Excel file and the output folder.
# 
# The source file contains all three required sheets for this use case.  
# The output folder will store the cleaned dataset and KPI files generated during processing.
# 
# I also defined the category mapping here so that inconsistent values from the raw dataset can be standardized later in the pipeline.

# In[2]:


SOURCE_FILE = "USECASE - Data Engineering.xlsx"
OUTPUT_DIR    = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

VALID_CATEGORIES = {
    "electronics", "clothing", "furniture", "home appliances"
}

CATEGORY_ALIAS = {
    "elec":            "Electronics",
    "electronics":     "Electronics",
    "cloth":           "Clothing",
    "clothing":        "Clothing",
    "furn":            "Furniture",
    "furniture":       "Furniture",
    "home":            "Home Appliances",
    "home appliances": "Home Appliances",
}


# ## 2. Data Ingestion
# 
# The retail data is provided in one Excel workbook with three sheets:
# 
# - product_details
# - retail_data1
# - retail_data2
# 
# Here I am loading all three sheets into separate DataFrames.
# 
# I am also validating the expected columns to make sure both retail datasets have the required schema before moving to the next step.

# In[3]:


log.info("=== STEP 1: DATA INGESTION ===")

try:
    product_dim = pd.read_excel(SOURCE_FILE, sheet_name="product_details")
    retail1     = pd.read_excel(SOURCE_FILE, sheet_name="retail_data1")
    retail2     = pd.read_excel(SOURCE_FILE, sheet_name="retail_data2")
    log.info("Files loaded — product_dim: %d rows | retail_data1: %d rows | retail_data2: %d rows",
             len(product_dim), len(retail1), len(retail2))
except FileNotFoundError:
    log.error("Source file '%s' not found.  Update SOURCE_FILE variable.", SOURCE_FILE)
    raise

EXPECTED_COLS = {
    "transaction_id", "customer_id", "customer_name", "product_id",
    "price", "product_name", "category", "purchase_location", "city",
    "transaction_date", "quantity", "payment_method", "discount",
    "email", "phone", "payment_status",
}
for name, df in [("retail_data1", retail1), ("retail_data2", retail2)]:
    missing = EXPECTED_COLS - set(df.columns.str.lower())
    if missing:
        log.warning("Schema mismatch in %s — missing columns: %s", name, missing)
    else:
        log.info("%s schema OK", name)


# ## 3. Combining Retail Datasets
# 
# The retail transactions are collected from two different source systems.
# 
# To process everything consistently, both retail datasets are combined into one DataFrame.
# 
# This makes the cleaning and transformation steps easier because the same logic can be applied to the full dataset in one place.

# In[4]:


log.info("=== STEP 2: COMBINING RAW DATASETS ===")
raw = pd.concat([retail1, retail2], ignore_index=True)
log.info("Combined rows (before cleaning): %d", len(raw))


# ## 4. Data Cleaning and Transformation
# 
# This is the main processing stage of the pipeline.
# 
# The raw retail data contains several issues such as duplicate transactions, missing prices, inconsistent category values, mixed product naming, different date formats, and invalid quantity values.
# 
# Each issue is handled step-by-step so that the final dataset becomes clean, consistent, and ready for analysis.

# In[5]:


log.info("=== STEP 3: CLEANING & TRANSFORMATION ===")

df = raw.copy()

df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

before = len(df)
df = df.drop_duplicates()
log.info("Exact duplicates removed: %d", before - len(df))

before = len(df)
df = (
    df.sort_values("payment_status", ascending=False)
      .drop_duplicates(subset=["transaction_id", "customer_id", "product_id",
                                "quantity", "transaction_date"], keep="first")
)
log.info("Duplicate transactions (multi-status) removed: %d", before - len(df))

before = len(df)
df = df[df["payment_status"].str.lower() == "successful"].copy()
log.info("Failed-payment rows removed: %d", before - len(df))


# ## 4.1 Date and Category Standardization
# 
# In this part, transaction dates are converted into a proper date format.
# 
# The category column is also cleaned and standardized so that similar values like `elec` and `electronics` are treated as the same category.
# 
# Rows with invalid dates are removed because they cannot be used correctly in time-based analysis.

# In[6]:


def parse_date(val):
    """Handle Excel serial numbers, datetime objects, and string formats."""
    if pd.isna(val):
        return pd.NaT
    if isinstance(val, datetime):
        return val
    if isinstance(val, (int, float)):
        try:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(val))
        except Exception:
            return pd.NaT
    try:
        return pd.to_datetime(str(val), dayfirst=False)
    except Exception:
        return pd.NaT

df["transaction_date"] = df["transaction_date"].apply(parse_date)
invalid_dates = df["transaction_date"].isna().sum()
if invalid_dates:
    log.warning("Rows with unparseable dates (will be dropped): %d", invalid_dates)
    df = df[df["transaction_date"].notna()]

df["category"] = (
    df["category"]
    .astype(str).str.strip().str.lower()
    .map(CATEGORY_ALIAS)
    .fillna("Unknown")
)
unknown_cat = (df["category"] == "Unknown").sum()
if unknown_cat:
    log.warning("Rows with unrecognised category (mapped to 'Unknown'): %d", unknown_cat)


# ## 4.2 Product, Location, City, Price, and Quantity Cleaning
# 
# Here, product names are corrected using the product details sheet.
# 
# Location and city values are standardized to keep the format consistent.
# 
# Missing prices are filled using the product master data. If price is still missing after that, the row is removed because revenue cannot be calculated without price.
# 
# Invalid quantity values are also removed.

# In[7]:


name_map = product_dim.set_index("product_id")["product_name"].to_dict()
df["product_name"] = df["product_id"].map(name_map).fillna(df["product_name"].str.title())

df["purchase_location"] = df["purchase_location"].str.strip().str.lower().str.capitalize()

df["city"] = df["city"].str.strip().str.title()

price_map = product_dim.set_index("product_id")["price"].to_dict()
missing_price_before = df["price"].isna().sum()
df["price"] = df.apply(
    lambda r: price_map.get(r["product_id"], np.nan) if pd.isna(r["price"]) else r["price"],
    axis=1,
)
log.info("Missing prices imputed from product_dim: %d", missing_price_before - df["price"].isna().sum())

before = len(df)
df = df[df["price"].notna()]
log.info("Rows dropped due to irrecoverable missing price: %d", before - len(df))

before = len(df)
df = df[df["quantity"].notna() & (df["quantity"] > 0)]
log.info("Rows with invalid quantity (≤0 or null) removed: %d", before - len(df))


# ## 4.3 Discount, Payment Method, and Category Enrichment
# 
# Discount values are converted into numeric format and limited between 0 and 1.
# 
# Payment method values are formatted properly.
# 
# If any category is still unknown, the product details sheet is used again to fill the correct category wherever possible.

# In[8]:


df["discount"] = pd.to_numeric(df["discount"], errors="coerce").fillna(0)
df["discount"] = df["discount"].clip(0, 1)

df["payment_method"] = df["payment_method"].str.strip().str.title()

df["payment_method"] = df["payment_method"].replace({
    "Upi": "UPI",
    "Netbanking": "NetBanking"
})

cat_map = product_dim.set_index("product_id")["category"].to_dict()
df["category"] = df.apply(
    lambda r: cat_map.get(r["product_id"], r["category"]) if r["category"] == "Unknown" else r["category"],
    axis=1,
)

log.info("Rows after all cleaning: %d", len(df))


# ## 5. PII Masking
# 
# The dataset contains customer email addresses and phone numbers.
# 
# Since these are sensitive personal details, they need to be protected before exporting the final data.
# 
# In this step:
# - the email local part is masked using hashing
# - phone numbers are masked while keeping the last four digits visible
# 
# This protects customer privacy while keeping the dataset usable for reporting.

# In[9]:


log.info("=== STEP 4: PII MASKING ===")

def mask_email(email: str) -> str:
    if pd.isna(email) or "@" not in str(email):
        return "***"
    local, domain = str(email).split("@", 1)
    hashed = hashlib.sha256(local.encode()).hexdigest()[:8]
    return f"{hashed}@{domain}"

def mask_phone(phone) -> str:
    p = str(phone).strip()
    if len(p) >= 4:
        return "XXXXXX" + p[-4:]
    return "XXXXXXXXXX"

df["email"] = df["email"].apply(mask_email)
df["phone"] = df["phone"].apply(mask_phone)

log.info("PII masking applied to email and phone columns.")


# ## 6. Feature Engineering
# 
# After cleaning the records, additional columns are created for analysis.
# 
# Revenue is calculated using price, quantity, and discount.
# 
# Date-based columns like year, month, month name, and quarter are also extracted.
# 
# These fields help in business reporting and dashboard visualization.

# In[10]:


log.info("=== STEP 5: FEATURE ENGINEERING ===")

df["revenue"] = df["price"] * df["quantity"] * (1 - df["discount"])

df["year"] = df["transaction_date"].dt.year
df["month"] = df["transaction_date"].dt.month
df["month_name"] = df["transaction_date"].dt.strftime("%B")
df["quarter"] = df["transaction_date"].dt.quarter

log.info("Revenue and date features computed.")


# ## 7. KPI Aggregation
# 
# After preparing the dataset, the main business KPIs are calculated.
# 
# These include:
# - total revenue
# - total transactions
# - unique customers
# - average order value
# - average discount
# 
# These values help measure business performance and are also useful for Power BI KPI cards.

# In[11]:


log.info("=== STEP 6: KPI AGGREGATION ===")

total_revenue = df["revenue"].sum()
total_transactions = len(df)
total_customers = df["customer_id"].nunique()
avg_order_value = df["revenue"].mean()
avg_discount = df["discount"].mean() * 100

log.info("─ Total Revenue        : ₹{:,.2f}".format(total_revenue))
log.info("─ Total Transactions   : {:,}".format(total_transactions))
log.info("─ Unique Customers     : {:,}".format(total_customers))
log.info("─ Avg Order Value      : ₹{:,.2f}".format(avg_order_value))
log.info("─ Avg Discount Applied : {:.1f}%".format(avg_discount))


# ## 7.1 Revenue Analysis by Business Dimensions
# 
# Along with summary KPIs, grouped revenue analysis is created.
# 
# Revenue is calculated based on:
# 
# - category
# - city
# - product
# - month
# - payment method
# - purchase location
# 
# These tables help identify trends and compare business performance across different areas.

# In[12]:


rev_category = (
    df.groupby("category")["revenue"]
    .sum().reset_index()
    .rename(columns={"revenue": "total_revenue"})
    .sort_values("total_revenue", ascending=False)
)

rev_city = (
    df.groupby("city")["revenue"]
    .sum().reset_index()
    .rename(columns={"revenue": "total_revenue"})
    .sort_values("total_revenue", ascending=False)
)

rev_product = (
    df.groupby(["product_id", "product_name", "category"])["revenue"]
    .sum().reset_index()
    .rename(columns={"revenue": "total_revenue"})
    .sort_values("total_revenue", ascending=False)
)

rev_monthly = (
    df.groupby(["year", "month", "month_name"])["revenue"]
    .sum().reset_index()
    .rename(columns={"revenue": "total_revenue"})
    .sort_values(["year", "month"])
)

rev_payment = (
    df.groupby("payment_method")["revenue"]
    .sum().reset_index()
    .rename(columns={"revenue": "total_revenue"})
    .sort_values("total_revenue", ascending=False)
)

rev_channel = (
    df.groupby("purchase_location")["revenue"]
    .sum().reset_index()
    .rename(columns={"revenue": "total_revenue"})
)


# ## 7.2 Customer and Regional Analysis
# 
# Additional business analysis is also created to understand customer contribution and regional performance.
# 
# This includes:
# 
# - city-wise and category-wise transactions
# - top customers based on revenue
# - quarterly performance summary
# 
# These outputs support deeper analysis in reporting dashboards.

# In[13]:


city_category = (
    df.groupby(["city", "category"])
    .agg(
        transactions=("transaction_id", "count"),
        revenue=("revenue", "sum")
    )
    .reset_index()
)

top_customers = (
    df.groupby(["customer_id", "customer_name"])["revenue"]
    .sum().reset_index()
    .rename(columns={"revenue": "total_revenue"})
    .sort_values("total_revenue", ascending=False)
    .head(20)
)

quarterly = (
    df.groupby(["year", "quarter"])
    .agg(
        transactions=("transaction_id", "count"),
        revenue=("revenue", "sum"),
        unique_customers=("customer_id", "nunique"),
        avg_discount=("discount", "mean"),
    )
    .reset_index()
)

log.info("All KPI aggregations complete.")


# ## 8. Exporting Final Outputs
# 
# After processing is complete, all cleaned data and KPI tables are exported into CSV format.
# 
# These files are saved inside the output folder and used directly in Power BI.
# 
# This makes the final dataset reporting-ready and easy to visualize.

# In[14]:


log.info("=== STEP 7: EXPORTING OUTPUTS ===")

cleaned_path = os.path.join(OUTPUT_DIR, "retail_cleaned.csv")
df.to_csv(cleaned_path, index=False)

kpi_tables = {
    "kpi_revenue_by_category": rev_category,
    "kpi_revenue_by_city": rev_city,
    "kpi_revenue_by_product": rev_product,
    "kpi_revenue_monthly": rev_monthly,
    "kpi_revenue_by_payment": rev_payment,
    "kpi_revenue_by_channel": rev_channel,
    "kpi_city_category": city_category,
    "kpi_quarterly": quarterly,
    "kpi_top_customers": top_customers,
}

for name, table in kpi_tables.items():
    path = os.path.join(OUTPUT_DIR, f"{name}.csv")
    table.to_csv(path, index=False)


# ## 8.1 KPI Summary Export
# 
# A separate summary table is created for important KPI card values.
# 
# This includes:
# 
# - total revenue
# - total transactions
# - unique customers
# - average order value
# - average discount percentage
# 
# This file is useful for Power BI KPI cards and quick business overview.

# In[15]:


summary = pd.DataFrame([{
    "total_revenue": round(total_revenue, 2),
    "total_transactions": total_transactions,
    "unique_customers": total_customers,
    "avg_order_value": round(avg_order_value, 2),
    "avg_discount_pct": round(avg_discount, 2),
}])

summary.to_csv(
    os.path.join(OUTPUT_DIR, "kpi_summary.csv"),
    index=False
)

log.info("Summary KPI card exported.")


# ## Conclusion
# 
# The retail transaction data was successfully processed through an end-to-end data engineering pipeline.
# 
# The workflow included:
# 
# - data ingestion
# - schema validation
# - cleaning and standardization
# - PII masking
# - feature engineering
# - KPI generation
# - exporting final datasets
# 
# The final output is a clean and structured dataset that can be directly connected to Power BI for reporting, dashboard creation, and business insights.

# In[16]:


log.info("=== PIPELINE COMPLETE ===")

log.info(
    "Output files written to '%s/'\n"
    "  • retail_cleaned.csv\n"
    "  • kpi_summary.csv\n"
    "  • kpi_revenue_by_*.csv\n"
    "  • kpi_city_category.csv\n"
    "  • kpi_quarterly.csv\n"
    "  • kpi_top_customers.csv",
    OUTPUT_DIR,
)

