# Project Four – Retail Sales KPI Dashboard (Excel BI)

This project builds an interactive Excel BI dashboard on top of a multi-year retail sales dataset (2023–2025). It showcases how Excel can be used as a lightweight BI tool to monitor revenue, profitability, and channel performance for a retail business.

📥 **Download the dashboard:** [excel_dashboard/Retail_Sales_KPI_Dashboard.xlsx](excel_dashboard/Retail_Sales_KPI_Dashboard.xlsx)

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

## 📊 Dataset

**File:** `data/cleaned/retail_sales_2023_2025_clean.xlsx`  
**Rows:** Transaction-level records across multiple years  
**Source:** <a href="https://www.kaggle.com/datasets/rafliprasetyo/retail-sales-dataset-from-2023-to-2025">Retail Sales Dataset from 2023 to 2025</a> by `rafliprasetyo` on Kaggle.  

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
| `Date`, `Year`, `Month`, `Quarter` | Time dimensions |
| `Customer_ID`, `Customer_Name` | Customer attributes |
| `Product_ID`, `Product_Name`, `Category` | Product hierarchy |
| `Channel`, `Region` | Sales channel & geography |
| `Quantity`, `Unit_Price`, `Discount_Pct` | Transaction details |
| `Revenue`, `COGS`, `Gross_Profit` | Financial metrics |

All data is stored in an Excel Table named `tblSales_2023_2025` for easy formulas, PivotTables, and refresh.

---

## 🔍 Business Questions

The dashboard is designed to help a retail manager answer:

- How are **Revenue** and **Gross Profit** trending over time (by month and year)?
- Which **channels** (online vs in-store) and **regions** drive the most revenue and margin?
- Which **product categories** and products contribute the most to total sales?
- How do **discounts** affect revenue and profitability?

---

## 🎯 Dashboard KPIs

| KPI | Formula |
|-----|---------|
| **Total Revenue** | `=SUM(tblSales_2023_2025[Revenue])` |
| **Gross Margin %** | `=SUM(tblSales_2023_2025[Gross_Profit]) / SUM(tblSales_2023_2025[Revenue])` |
| **Total Transactions** | `=COUNTA(tblSales_2023_2025[Transaction_ID])` |
| **Avg Order Value** | `=SUM(tblSales_2023_2025[Revenue]) / COUNTA(tblSales_2023_2025[Transaction_ID])` |
| **YoY Revenue Growth %** | `=(KPI_Calc!B2 − KPI_Calc!B3) / KPI_Calc!B3` where B2 = Revenue 2025, B3 = Revenue 2024 — with ▲▼ conditional icon |

---

## 📈 Dashboard Design

**File:** `excel_dashboard/Retail_Sales_KPI_Dashboard.xlsx`

The `Dashboard` sheet contains:

### KPI Cards (top row)
- **Total Revenue**
- **Gross Margin %**
- **Total Transactions**
- **Average Order Value**
- *(Optional)* YoY Revenue Growth % with up/down arrows

### Visuals
- **Line chart:** Revenue by Month (multi-year trend)
- **Column chart:** Revenue by Channel
- **Bar chart:** Revenue & Gross Profit by Region
- **Bar chart:** Top Categories by Revenue

### Filters (Slicers)
- Year
- Region
- Channel
- Category

Slicers are connected to all relevant PivotTables, so selecting a year or channel instantly updates all KPIs and charts.

### Layout

```
┌──────────────────────────────────────────────────────────────────────┐
│  Dark header: "Retail Sales Performance Dashboard (2023–2025)"       │
├──────────────────────────────────────────────────────────────────────┤
│ [Total Revenue] [Gross Margin %] [Total Transactions] [Avg Order Val]│
├───────────────────────────────────┬──────────────────────────────────┤
│  Line: Revenue by Month (trend)   │  Column: Revenue by Channel      │
├───────────────────────────────────┼──────────────────────────────────┤
│  Bar: Revenue & GP by Region      │  Bar: Top Categories by Revenue  │
└───────────────────────────────────┴──────────────────────────────────┘
  Slicers: Year / Channel (left)    Region / Category (top-right)
```

---

## 🚀 How to Use / Refresh

1. Open `data/raw/retail_sales_2023_2025.csv` in Excel.
2. **Data → Get Data → From Text/CSV** → import into a table named **`tblSales_2023_2025`**.
3. Save as `data/cleaned/retail_sales_2023_2025_clean.xlsx`.
4. Open `excel_dashboard/Retail_Sales_KPI_Dashboard.xlsx`.
5. **Data → Refresh All** (Ctrl+Alt+F5) to update all PivotTables and KPIs.
6. Interact with slicers (Year, Channel, Region, Category) to explore performance.

For a detailed step-by-step walkthrough, see **[docs/EXCEL_DASHBOARD_SETUP.md](docs/EXCEL_DASHBOARD_SETUP.md)**.

---

## ⚙️ Implementation Details

- Data model built on a single fact table (`tblSales_2023_2025`) with derived columns for `Year`, `Month`, `Month_Name`, and `Quarter`.
- KPIs calculated with standard Excel functions: `SUM`, `COUNTA`, `SUMIFS`, and percentage formulas for Gross Margin and YoY growth.
- All visuals are PivotTable-based, making the dashboard easy to refresh when new data is appended to the table.
- Workbook sheets: `Dashboard` (visuals + KPI cards), `PivotData` (all PivotTables), `KPI_Calc` (formula calculations).

---

## Key Insights

- **Revenue concentration:** A small subset of product categories accounts for a disproportionately large share of total revenue (an 80/20 pattern), indicating that focusing merchandising and inventory decisions on these top categories can drive the biggest financial impact.

- **Seasonality:** Monthly sales exhibit clear seasonal peaks across the 2023–2025 period, with higher average order values during peak months. This pattern can inform staffing plans, promotional timing, and inventory replenishment cycles.

- **Pricing vs. volume tradeoff:** Higher-priced categories generate more revenue per transaction but lower unit volume, while low-priced essentials dominate total quantity sold. This suggests using margin optimization strategies for premium categories and efficiency/throughput strategies for high-volume staples.

---

## 🛠 Tools Used

- **Microsoft Excel** – Power Query, PivotTables, PivotCharts, Slicers, GETPIVOTDATA
- **Python** – used for light preprocessing / exploration (optional)

---

## 📄 License

This project uses the 'Retail Sales Dataset from 2023 to 2025' from Kaggle for educational and portfolio purposes.
