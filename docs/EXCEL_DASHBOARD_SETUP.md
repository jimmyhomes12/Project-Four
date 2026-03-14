# Excel Retail BI Dashboard – Setup Guide

This guide walks you through converting the raw CSV into a fully functional Excel BI dashboard with KPI cards, charts, and PivotTables.

---

## Prerequisites

| Item | Detail |
|------|--------|
| Excel version | Microsoft Excel 2016 or later (Power Query included) |
| Raw data | `data/raw/retail_sales_2023_2025.csv` |

---

## Step 1 – Import & Clean the Data

1. Open a **new blank workbook** in Excel.
2. Go to **Data → Get Data → From File → From Text/CSV**.
3. Browse to `data/raw/retail_sales_2023_2025.csv` and click **Import**.
4. In the Power Query preview, confirm:
   - `Date` column is detected as **Date**
   - `Revenue`, `COGS`, `Gross_Profit`, `Unit_Price` are **Decimal Number**
   - `Year`, `Month`, `Quantity` are **Whole Number**
   - `Discount_Pct` is **Decimal Number**
5. Click **Close & Load To…** → choose **Table** in the **existing worksheet** (cell A1).
6. The table will be auto-named `Table1`. Rename it to **`tblSales_2023_2025`**:  
   - Click anywhere inside the table → **Table Design** tab → change the name in the *Table Name* box.
7. Save this workbook to `data/cleaned/retail_sales_2023_2025_clean.xlsx`.

> **Tip:** All 17,000+ rows will load automatically. The helper columns (`Year`, `Month`, `Quarter`, `Month_Name`) are already present in the CSV so no further calculated columns are needed for basic analysis.

---

## Step 2 – Create the Dashboard Workbook

1. Open a **new workbook** and save it immediately as `excel_dashboard/Retail_Sales_KPI_Dashboard.xlsx`.
2. Create the following sheets (right-click sheet tab → Insert):

| Sheet name | Purpose |
|------------|---------|
| `Dashboard` | Visual dashboard (KPIs + charts) |
| `PivotData` | All PivotTables that feed the dashboard |
| `KPI_Calc` | Named-range formula calculations |

---

## Step 3 – Connect PivotTables to the Source Data

> You can link PivotTables directly to the clean workbook so a single data refresh keeps everything current.

1. In `Retail_Sales_KPI_Dashboard.xlsx`, go to the **`PivotData`** sheet.
2. **Insert → PivotTable → From External Data Source → Use an existing connection**.
3. Click **Browse for More…** and point to `retail_sales_2023_2025_clean.xlsx`, then select the `tblSales_2023_2025` table.

Alternatively (simpler): copy-paste `tblSales_2023_2025` into a `RawData` sheet inside the dashboard workbook, then build all PivotTables from that table.

### PivotTables to create (on the `PivotData` sheet)

| PivotTable name | Rows | Columns | Values |
|-----------------|------|---------|--------|
| `pvt_YearlyRevenue` | Year | — | Sum of Revenue |
| `pvt_MonthlyRevenue` | Year, Month_Name | — | Sum of Revenue |
| `pvt_ChannelRevenue` | Year | Channel | Sum of Revenue |
| `pvt_RegionRevenue` | Year | Region | Sum of Revenue |
| `pvt_CategoryRevenue` | Year | Category | Sum of Revenue |
| `pvt_TopProducts` | Product_Name | — | Sum of Revenue |
| `pvt_GrossMargin` | Year | Category | Sum of Gross_Profit, Sum of Revenue |

---

## Step 4 – KPI Formulas (on the `KPI_Calc` sheet)

Place these formulas starting at cell **B2** and label column A accordingly.

```
A2: YTD Revenue (2025)
B2: =GETPIVOTDATA("Revenue",pvt_YearlyRevenue,"Year",2025)

A3: Revenue 2024
B3: =GETPIVOTDATA("Revenue",pvt_YearlyRevenue,"Year",2024)

A4: Revenue 2023
B4: =GETPIVOTDATA("Revenue",pvt_YearlyRevenue,"Year",2023)

A5: YoY Growth % (2025 vs 2024)
B5: =(B2-B3)/B3

A6: YoY Growth % (2024 vs 2023)
B6: =(B3-B4)/B4

A7: Total Revenue (all years)
B7: =B2+B3+B4

A8: Overall Gross Margin %
B8: =SUMIF(tblSales_2023_2025[Year],2025,tblSales_2023_2025[Gross_Profit])
    / SUMIF(tblSales_2023_2025[Year],2025,tblSales_2023_2025[Revenue])

A9: Top Region (2025)
B9: =INDEX(pvt_RegionRevenue[Region],
        MATCH(MAX(pvt_RegionRevenue[2025]),pvt_RegionRevenue[2025],0))

A10: Top Category (2025)
B10: =INDEX(pvt_CategoryRevenue[Category],
        MATCH(MAX(pvt_CategoryRevenue[2025]),pvt_CategoryRevenue[2025],0))
```

> Format B5 and B6 as **Percentage (1 decimal place)**.
> Format B2, B3, B4, B7 as **Currency ($, 0 decimal places)**.

---

## Step 5 – Dashboard Layout (on the `Dashboard` sheet)

### Row 1 – KPI Cards (columns A–P, rows 1–6)

Create 4 KPI card blocks using merged cells and borders:

| Card | Metric | Formula reference |
|------|--------|-------------------|
| Card 1 | **YTD Revenue 2025** | `=KPI_Calc!B2` |
| Card 2 | **YoY Growth % (2025 vs 2024)** | `=KPI_Calc!B5` with ▲▼ conditional icon |
| Card 3 | **Top Region** | `=KPI_Calc!B9` |
| Card 4 | **Top Category** | `=KPI_Calc!B10` |

**Conditional formatting for growth arrow:**
- Select the YoY Growth card → Home → Conditional Formatting → Icon Sets → 3 Arrows.
- Rule: ≥ 0 = green up arrow, < 0 = red down arrow.

### Row 2 – Time-Series Chart (columns A–P, rows 8–22)

1. Select `pvt_MonthlyRevenue` → Insert → **Line Chart with Markers**.
2. Add a **Slicer** for `Year` (PivotTable Analyze → Insert Slicer → Year).
3. Place the slicer to the right of the chart.
4. Title: *Monthly Revenue by Year*.

### Row 3 – Category & Channel Mix (columns A–P, rows 24–38)

**Left half – Stacked Bar by Category:**
1. Use `pvt_CategoryRevenue` → Insert → **Stacked Bar Chart**.
2. Add data labels showing revenue amounts.
3. Title: *Revenue by Product Category*.

**Right half – Donut by Channel:**
1. Use `pvt_ChannelRevenue` (2025 only) → Insert → **Doughnut Chart**.
2. Add a data label showing channel names + percentages.
3. Title: *2025 Revenue Split: Online vs In-Store*.

### Row 4 – Regional Performance (columns A–P, rows 40–54)

1. Use `pvt_RegionRevenue` → Insert → **Clustered Column Chart**.
2. Group by Region; series = Year (2023, 2024, 2025).
3. Add a Slicer for `Region`.
4. Title: *Revenue by Region – Year-over-Year*.

---

## Step 6 – Slicers & Interactivity

1. Go to **PivotTable Analyze → Insert Slicer** for each of:
   - `Year`
   - `Channel`
   - `Category`
   - `Region`
2. Connect each slicer to **all PivotTables** (right-click slicer → Report Connections → select all).
3. Style slicers with a consistent colour theme (Slicer tab → Slicer Styles).

---

## Step 7 – Moving-Average Trendline

1. Click the Monthly Revenue line chart → **Chart Design → Add Chart Element → Trendline → Moving Average**.
2. Set Period = **3** (3-month moving average).
3. Format the trendline as a dashed line in a contrasting colour.

---

## Step 8 – Final Polish

- Apply a consistent colour theme: **Page Layout → Themes** (e.g., *Office* or *Facet*).
- Hide gridlines on the Dashboard sheet: **View → uncheck Gridlines**.
- Lock the Dashboard sheet: **Review → Protect Sheet** (allow only selecting cells).
- Add a header row with the project title using WordArt or a large merged cell.

---

## Refreshing Data

When source data changes:
1. Open `retail_sales_2023_2025_clean.xlsx` and update/replace the table.
2. In the dashboard workbook: **Data → Refresh All** (or Ctrl+Alt+F5).
3. All PivotTables, charts, and KPI formulas update automatically.

---

## Column Reference

| Column | Type | Description |
|--------|------|-------------|
| `Transaction_ID` | Text | Unique transaction identifier (T000001…) |
| `Date` | Date | Transaction date (YYYY-MM-DD) |
| `Year` | Integer | 2023, 2024, or 2025 |
| `Month` | Integer | 1–12 |
| `Month_Name` | Text | January … December |
| `Quarter` | Text | Q1, Q2, Q3, Q4 |
| `Customer_ID` | Text | Unique customer identifier (C0001…) |
| `Customer_Name` | Text | Full customer name |
| `Product_ID` | Text | Product code (P001–P025) |
| `Product_Name` | Text | Product display name |
| `Category` | Text | Electronics, Footwear, Apparel, Home & Kitchen, Sports & Fitness |
| `Channel` | Text | Online or In-Store |
| `Region` | Text | North, South, East, West, Midwest |
| `Quantity` | Integer | Units sold per transaction |
| `Unit_Price` | Decimal | List price per unit ($) |
| `Discount_Pct` | Decimal | Discount applied (0, 0.05, 0.10, 0.15, 0.20) |
| `Revenue` | Decimal | `Unit_Price × Quantity × (1 − Discount_Pct)` |
| `COGS` | Decimal | Cost of goods sold (`Unit_Cost × Quantity`) |
| `Gross_Profit` | Decimal | `Revenue − COGS` |
