# Project Four – Excel Retail BI Dashboard (2023–2025)

An **Excel-based Business Intelligence dashboard** built on a multi-year retail sales dataset (2023–2025).  
The project demonstrates time-intelligence KPIs, year-over-year growth analysis, product/category mix, and regional performance — all driven by PivotTables, slicers, and structured Excel formulas.

---

## 📁 Folder Structure

```
Project-Four/
│
├── data/
│   ├── raw/
│   │   └── retail_sales_2023_2025.csv      ← source dataset (17,000+ transactions)
│   └── cleaned/
│       └── retail_sales_2023_2025_clean.xlsx  ← after import & light cleaning in Excel
│
├── excel_dashboard/
│   └── Retail_Sales_KPI_Dashboard.xlsx     ← main dashboard workbook
│
├── docs/
│   └── EXCEL_DASHBOARD_SETUP.md            ← step-by-step build & refresh guide
│
└── README.md
```

---

## 📊 Dataset Overview

| Field | Detail |
|-------|--------|
| **Rows** | ~17,000 transactions |
| **Period** | 1 Jan 2023 – 31 Dec 2025 (3 full years) |
| **Channels** | Online, In-Store |
| **Regions** | North, South, East, West, Midwest |
| **Categories** | Electronics · Footwear · Apparel · Home & Kitchen · Sports & Fitness |
| **Products** | 25 SKUs |
| **Customers** | 500 unique customers |

### Key Columns

| Column | Description |
|--------|-------------|
| `Transaction_ID` | Unique ID per transaction |
| `Date` / `Year` / `Month` / `Quarter` | Time dimensions |
| `Customer_ID`, `Customer_Name` | Customer attributes |
| `Product_ID`, `Product_Name`, `Category` | Product hierarchy |
| `Channel`, `Region` | Sales channel & geography |
| `Quantity`, `Unit_Price`, `Discount_Pct` | Transaction details |
| `Revenue`, `COGS`, `Gross_Profit` | Financial metrics |

---

## 🎯 Dashboard KPIs

| KPI | Formula |
|-----|---------|
| **YTD Revenue** | `=GETPIVOTDATA("Revenue", pvt_YearlyRevenue, "Year", 2025)` |
| **YoY Growth % (2025 vs 2024)** | `=(Revenue_2025 − Revenue_2024) / Revenue_2024` |
| **YoY Growth % (2024 vs 2023)** | `=(Revenue_2024 − Revenue_2023) / Revenue_2023` |
| **Gross Margin %** | `=SUM(Gross_Profit) / SUM(Revenue)` |
| **Top Region** | `=INDEX(...)` on region pivot |
| **Top Category** | `=INDEX(...)` on category pivot |

---

## 📈 Dashboard Layout

```
┌─────────────────────────────────────────────────────────────┐
│  [KPI Card: YTD Revenue]  [YoY Growth ▲]  [Top Region]  [Top Category]  │
├─────────────────────────────────────────────────────────────┤
│  Line Chart: Monthly Revenue by Year          [Slicer: Year / Channel]   │
├──────────────────────────────┬──────────────────────────────┤
│  Stacked Bar: Revenue by     │  Doughnut: Online vs In-Store│
│  Product Category            │  (2025 split)                │
├─────────────────────────────────────────────────────────────┤
│  Clustered Column: Revenue by Region – YoY  [Slicer: Region]            │
└─────────────────────────────────────────────────────────────┘
```

---

## 🚀 Quick Start

### 1. Open & Clean the Raw Data
1. Open `data/raw/retail_sales_2023_2025.csv` in Excel.
2. **Data → Get Data → From Text/CSV** → import into a table.
3. Name the table **`tblSales_2023_2025`**.
4. Save as `data/cleaned/retail_sales_2023_2025_clean.xlsx`.

### 2. Build the Dashboard
Follow the detailed walkthrough in **[docs/EXCEL_DASHBOARD_SETUP.md](docs/EXCEL_DASHBOARD_SETUP.md)**.

Steps covered:
- Creating and naming PivotTables
- KPI formula sheet with GETPIVOTDATA
- Chart types and slicer connections
- Conditional formatting for growth arrows
- Moving-average trendline
- Sheet protection & final polish

### 3. Refresh
**Data → Refresh All** (Ctrl+Alt+F5) after updating the source data.

---

## 💡 Key Insights (Sample)

- **Online channel** accounts for ~55% of all transactions across all years.
- Revenue shows a clear **Q4 seasonal spike** (Nov–Dec) every year.
- **YoY growth ~12% (2023→2024) and ~7% (2024→2025)** reflecting healthy but moderating growth.
- **Electronics** is the highest-revenue category; **Sports & Fitness** shows the strongest per-unit margin.

---

## 🛠 Tools Used

- **Microsoft Excel** – Power Query, PivotTables, PivotCharts, Slicers, GETPIVOTDATA
- **Python** – synthetic dataset generation (`data/raw/retail_sales_2023_2025.csv`)

---

## 📄 License

This project uses a synthetic dataset generated for portfolio and educational purposes.
