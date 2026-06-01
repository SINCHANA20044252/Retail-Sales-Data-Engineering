# Retail Sales Data Engineering — ABC Retail Solutions

An end-to-end data engineering pipeline that ingests, cleans, transforms, and prepares retail transaction data for business intelligence reporting via Power BI.

---

## Project Overview

ABC Retail Solutions operates across multiple cities through online and offline sales channels. This pipeline addresses data quality challenges across transactional datasets — including duplicates, missing values, inconsistent naming, invalid quantities, and varying date formats — to produce a clean, analytics-ready dataset.

---

## Repository Structure

```
Retail-Sales-Data-Engineering/
│
├── Code/
│   └── retail_pipeline.py          # Full data engineering pipeline (Python)
│
├── Documentation/
│   └── documentation.docx          # Architecture diagram, data flow, assumptions,
│                                   # cleaning strategy, transformation logic
│
└── Power BI/
    └── Retail_sales_dashboard.pbix # Interactive Power BI dashboard
```

---

## Pipeline Stages

### 1. Data Ingestion
- Loads three sheets from the source Excel file: `retail_data1`, `retail_data2`, and `product_details`
- Validates schema against expected columns for both transaction datasets
- Logs any missing columns or schema mismatches

### 2. Data Cleaning & Transformation
- Removes exact duplicates and duplicate transactions across source systems
- Filters out failed/unsuccessful payment records
- Standardizes date formats (handles Excel serial numbers, datetime objects, and string formats)
- Maps inconsistent category values (`elec` → `Electronics`, `home` → `Home Appliances`, etc.)
- Corrects product names using the `product_details` dimension table
- Standardizes city names, purchase location, and payment method values
- Imputes missing prices from the product master; drops rows where price is irrecoverable
- Removes rows with null or non-positive quantity values
- Normalizes discount values to the 0–1 range

### 3. PII Masking
- **Email**: local part hashed using SHA-256; domain retained (`abc12345@gmail.com`)
- **Phone**: masked with `XXXXXX` prefix, last 4 digits visible (`XXXXXX4321`)

### 4. Feature Engineering
- `revenue` = `price × quantity × (1 − discount)`
- Extracted date features: `year`, `month`, `month_name`, `quarter`

### 5. KPI Aggregation
| KPI | Description |
|-----|-------------|
| Total Revenue | Sum of all transaction revenue |
| Total Transactions | Count of valid records |
| Unique Customers | Count of distinct customer IDs |
| Average Order Value | Mean revenue per transaction |
| Average Discount | Mean discount percentage applied |
| Revenue by Category | Grouped by product category |
| Revenue by City | Grouped by purchase city |
| Revenue by Product | Grouped by product ID and name |
| Revenue by Month | Monthly revenue trend |
| Revenue by Quarter | Quarterly performance summary |
| Revenue by Payment Method | Breakdown by UPI, Card, NetBanking, etc. |
| Revenue by Channel | Online vs Offline |
| Top 20 Customers | Ranked by total revenue contribution |
| City × Category Matrix | Transactions and revenue by city and category |

### 6. Output Export
All outputs are saved to an `output/` folder as CSV files, ready for Power BI connection:
- `retail_cleaned.csv` — final cleaned and enriched dataset
- `kpi_summary.csv` — summary KPI card values
- `kpi_revenue_by_category.csv`
- `kpi_revenue_by_city.csv`
- `kpi_revenue_by_product.csv`
- `kpi_revenue_monthly.csv`
- `kpi_revenue_by_payment.csv`
- `kpi_revenue_by_channel.csv`
- `kpi_city_category.csv`
- `kpi_quarterly.csv`
- `kpi_top_customers.csv`

---

## Tech Stack

| Tool | Purpose |
|------|---------|
| Python 3.x | Pipeline development |
| pandas | Data manipulation and transformation |
| numpy | Numerical operations |
| hashlib | PII masking (SHA-256) |
| logging | Pipeline logging and monitoring |
| Power BI Desktop | Dashboard and visualization |

---

## How to Run

1. Clone the repository:
   ```bash
   git clone https://github.com/SINCHANA20044252/Retail-Sales-Data-Engineering.git
   cd Retail-Sales-Data-Engineering/Code
   ```

2. Install dependencies:
   ```bash
   pip install pandas numpy openpyxl
   ```

3. Place the source Excel file (`USECASE - Data Engineering.xlsx`) in the same directory as `retail_pipeline.py`

4. Run the pipeline:
   ```bash
   python retail_pipeline.py
   ```

5. Cleaned datasets and KPI tables will be saved in the `output/` folder.

6. Open `Power BI/Retail_sales_dashboard.pbix` and refresh the data source to point to your local `output/` folder.

---

## Power BI Dashboard

The dashboard includes the following pages:

- **Revenue Analysis** — Total revenue, monthly trends, quarterly performance
- **Product Performance** — Top products, revenue by product and category
- **Category Trends** — Category-wise revenue comparison and growth
- **Regional Insights** — City-wise revenue breakdown and channel analysis

Interactive features include slicers for date, city, category, and payment method; KPI cards; and drill-through visuals.

---

## Assumptions

- Only transactions with `payment_status = Successful` are included in analysis
- Missing prices are filled from `product_details`; rows without a recoverable price are excluded
- Rows with quantity ≤ 0 are treated as invalid and removed
- Unrecognized category values are mapped using the product dimension table; remaining unknowns are labeled `Unknown`
- Discount values outside the 0–1 range are clipped to valid bounds

---

## Author

**Sinchana** — Data Engineering Assessment Submission  
CMRIT 
