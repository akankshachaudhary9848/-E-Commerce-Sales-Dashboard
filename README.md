# E-Commerce Sales Dashboard (2023–2025)

**Group 4 – Final Project** · Excel workbook: `group_4_project_work_4_Final_Project.xlsx`

An interactive Excel dashboard that analyses 5,000 e-commerce orders placed between **4 Oct 2023 and 3 Oct 2025** across four Indian regions. It shows sales, profit and quantity by category, product, region, payment mode and time, and can be filtered with **Category** and **Year** slicers.

---

## Headline numbers (full dataset)

| Metric | Value |
|---|---|
| Orders | 5,000 |
| Total Sales | ₹53,36,66,024.35 (≈ ₹53.37 crore) |
| Total Profit | ₹7,97,08,734.91 (≈ ₹7.97 crore) |
| Total Quantity | 14,963 units |
| Overall profit margin | ≈ 14.9 % |
| Best year | 2024 – ₹27.15 cr sales / ₹4.04 cr profit |
| Top category by sales | Home Decor (₹5.72 cr) |
| Top category by profit | Furniture (₹86.9 lakh) |
| Top region by sales and units | North (₹14.36 cr, 3,949 units) |
| Best month by profit | May (₹76.7 lakh) |

---

## Workbook structure

| Sheet | Purpose |
|---|---|
| **Dashboard** | The presentation layer: title banner, 3 KPI tiles, 10 charts and 2 slicers. Gridlines are hidden. |
| **pivot table** | Eleven PivotTables plus three summary cells (`SUM SALES`, `TOTAL PROFIT`, `TOTAL QUANTITY`). Every dashboard chart reads from here. |
| **Clean data** | The cleaned dataset, stored as an Excel Table named `Table1` (5,000 rows × 21 columns). This is the PivotTable source. |
| **Group_4 - (Ecommerce Sales Data…** | The raw data as received (5,000 rows × 14 columns). It is kept unchanged for reference. |

### Data flow

```
Raw sheet  →  Clean data (Table1, + derived columns)  →  PivotTables  →  Charts + slicers on Dashboard
```

---

## Dashboard contents

**KPI tiles:** Total Sales · Total Profit · Total Quantity.

**Charts (10)**

1. Category-wise Sales vs Profit analysis (combo: columns for sales, line for profit)
2. Top 10 products by sales (bar)
3. Bottom 10 products by sales (bar)
4. Region-wise max profit (single largest order profit per region)
5. Region-wise total profit (pie)
6. Articles sold region-wise (bar)
7. Year-wise payment mode and sales (clustered bar)
8. Month-wise profit (bar)
9. Year-wise total profit (line)
10. Average discount by category (bar)

**Slicers:** `Category` and `Year`. They are connected to the PivotTables, so selecting a value re-filters the charts.

---

## Data dictionary

### Raw / source columns (14)

| Column | Description |
|---|---|
| Order ID | Unique order number (10001–15000) |
| Order Date | Date of order |
| Customer Name | Customer who placed the order |
| Region | North, South, East or West |
| City | One of 20 cities |
| Category | One of 10 product categories |
| Sub-Category | One of 50 sub-categories |
| Product Name | Product (sub-category + a suffix) |
| Quantity | Units ordered |
| Unit Price | Price per unit (₹) |
| Discount | Discount in percent: 0, 5, 10, 15 or 20 |
| Sales | Order value after discount (₹) = Quantity × Unit Price × (1 − Discount/100) |
| Profit | Profit on the order (₹) |
| Payment Mode | COD, Credit Card, Debit Card, Net Banking or UPI |

### Derived columns added in `Clean data` (all formulas)

| Column | Logic |
|---|---|
| Month Name | `TEXT(Order Date, "mmmm")` |
| Year | `TEXT(Order Date, "yyyy")` (stored as text) |
| Discount Value | Quantity × Unit Price × Discount / 100 (₹ discount given) |
| COGS | Sales − Profit |
| Profit Segment | Profit ≥ ₹40,000 → **High**; ≥ ₹10,000 → **Medium**; otherwise **Low** |
| Customer Segment | Sales ≥ ₹3,00,000 → **VIP Platinum**; ≥ ₹1,00,000 → **Gold**; otherwise **Regular** |
| discount percentage | Discount / 100 (used for the average-discount pivot) |

---

## How to use

1. Open the file in **desktop Excel 2013 or later**. Slicers need Excel 2010+, and the dashboard is best viewed at 100 % zoom.
2. Go to the **Dashboard** sheet.
3. Click a **Category** and/or **Year** button to filter. Hold `Ctrl` to select several, and use the clear-filter icon (top right of the slicer) to reset.
4. To refresh after editing data in `Clean data`, use **Data → Refresh All**.
5. To add new rows, type directly below the table. `Table1` expands, and the derived columns fill themselves.

---

## Known issues and caveats

These were found while documenting the file. Worth a look before presenting it.

- **"Region-wise total profit" pie is mislabelled and probably stale.** The underlying PivotTable (`pivot table`!A67) sums **Sales**, not Profit. Its saved values (grand total ₹6.40 cr) match **2023 sales only**, so a Year filter was probably active when the workbook was saved. Fix: change the value field to *Sum of Profit* and clear the Year filter. Correct all-years profit by region is East ₹2.05 cr, North ₹2.13 cr, South ₹1.83 cr, West ₹1.96 cr.
- **Partial years.** Data runs Oct 2023 to Oct 2025, so 2023 (3 months) and 2025 (9 months) are incomplete. The 2024 peak and the year-on-year trend therefore reflect different period lengths and are not like-for-like.
- **KPI tile values look static.** The three tiles contain typed text, not cell links. They match the data now but will not change with the slicers or new data. Link them to `pivot table`!F1:F3 (or GETPIVOTDATA) to make them dynamic.
- **Pivot chart for "Max profit" shows an order-level maximum**, not a total. Do not read it as regional performance.
- **Average discount** is stored as a fraction (e.g. 0.10), so format the axis as a percentage.
- Top/Bottom-10 pivots include a *Grand Total* row in the cell range. Charts use ranges that extend to it (A32:A42, A46:A56), so check that the total is not plotted as an 11th bar.
- `Year` is text, not a number, so it sorts and filters alphabetically.
- Product and customer names look synthetic (placeholder-style product suffixes). The dataset appears to be training/demo data.
- Currency is shown in rupees using the Indian digit-grouping format (lakh/crore).

---

## Suggested extensions

- Add Region and Payment Mode slicers.
- Add a profit-margin % KPI and a Profit Segment / Customer Segment chart. Those columns are already prepared but unused on the dashboard.
- Add a city-level view (20 cities).
- Normalise year comparisons (e.g. monthly averages) to account for partial years.

---

## Tech notes

- Built in Microsoft Excel (PivotTables, PivotCharts-style charts, slicers, structured table references).
- No macros, VBA or external data connections.
- File size ≈ 2 MB.
