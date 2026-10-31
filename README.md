# King's House Excel Projects Portfolio

**Author:** Clinton Onezi Aiyede
**Program:** BoyCode Africa, King's House team
**Tools:** Microsoft Excel and Google Sheets

This repository holds the project assignments I completed during the program. It covers a full e-commerce sales and profitability analysis (done two ways), a set of advanced spreadsheet-function exercises, and my earlier core exercises on logical functions, lookups, formulas, named ranges and formatting.

---

## Projects at a glance

| # | Project | What it is | Folder |
|---|---------|-----------|--------|
| 1 | **Excel Capstone: Sales & Profitability (Group Project, Personal Copy)** | Cleaned 1,800 e-commerce orders and analysed realised (delivered) revenue, profit, fulfilment risk, regions, segments and discounts | [`01-excel-capstone-group-project`](./01-excel-capstone-group-project) |
| 2 | **Executive Sales Dashboard** | Executive-level dashboard and management recommendations built on the same cleaned data, covering all orders | [`02-executive-sales-dashboard`](./02-executive-sales-dashboard) |
| 3 | **Week 3: Advanced Functions (Exercises 02 to 05)** | UNIQUE, FILTER, SORT, XLOOKUP, CHOOSE, SWITCH, MEDIAN, MODE, LARGE, SMALL, RANK, ROUND, MROUND, AGGREGATE | [`03-week-3-advanced-functions`](./03-week-3-advanced-functions) |
| 4 | **Core Excel Exercises 01 to 05 (screenshots)** | Logical functions, lookups, formulas, named ranges, number formats and formatting, done in Google Sheets | [`04-excel-exercises-01-05`](./04-excel-exercises-01-05) |

### Repository structure

```
program-projects-portfolio/
├── README.md
├── 01-excel-capstone-group-project/
├── 02-executive-sales-dashboard/
├── 03-week-3-advanced-functions/
└── 04-excel-exercises-01-05/
```

---

## Project 1: Excel Capstone, Sales & Profitability (Group Project, Personal Copy)

**Folder:** `01-excel-capstone-group-project`
**Type:** Group capstone, my personal copy of the workbook
**Tools:** Excel (SUMIFS, SUMIF, COUNTIFS, COUNTIF, AVERAGEIFS, SUMPRODUCT, IFERROR, MONTH, TEXT), charts

### The brief
Every House in the program received the same e-commerce dataset (orders for FY 2025) and was asked to study it from a **Sales & Profitability** angle. The raw file had quality problems, so cleaning came before any analysis.

### Data cleaning decisions

| Issue found | Decision |
|-------------|----------|
| 12 exact duplicate order rows | Removed. The dataset went from 1,812 to **1,800** rows |
| 10 orders with no Region | Labelled `Unknown` and kept, since revenue and profit on those rows were valid |
| 7 orders with no Acquisition Channel | Labelled `Unknown` and kept, for the same reason |
| About 1 in 6 orders with no customer rating | Left blank and treated as "not rated", never as zero |
| Cancelled and Returned orders (about a third of all orders) | **Not counted as real revenue**, because the money never came in. Kept visible in a dedicated sheet because it turned out to be a main finding |

### How the workbook is organised
Everything is formula-driven from one cleaned data sheet, so if a value in the data changes, the rest of the workbook updates.

| Sheet | What it shows |
|-------|---------------|
| README | What we did, cleaning notes and a guide to the sheets |
| Dashboard | One-page summary of the main KPIs and 4 charts. Start here |
| Sales KPIs | Headline numbers: revenue, profit, margin, average order value, ROAS, revenue at risk |
| Category Analysis | Orders, revenue, profit, margin, average order value and units by product category |
| Region Analysis | The same breakdown across 10 regions plus `Unknown` |
| Monthly Trend | Revenue and profit month by month across 2025 |
| Segment Analysis | New vs Returning vs Loyal customers, and the effect of discount level on margin |
| Order Fulfilment | Delivered vs Cancelled vs Returned, on-time vs late delivery, and the revenue-at-risk calculation |
| Clean Data | The full cleaned dataset |

### Headline results (delivered orders only unless stated)

| Metric | Value |
|--------|-------|
| Total orders (all statuses) | 1,800 |
| Delivered orders | 1,175 (65.3%) |
| Realised revenue | ₦61,218,235 |
| Realised gross profit | ₦19,461,585 |
| Gross profit margin | 31.8% |
| Average order value | ₦52,101 |
| Marketing spend (delivered orders) | ₦1,342,144 |
| Revenue per marketing naira (ROAS) | 45.61x |
| Revenue lost to cancellations | ₦16,105,715 |
| Revenue lost to returns | ₦17,642,192 |
| **Total revenue at risk** | **₦33,747,907** (about 35.5% of the ₦94,966,143 booked) |
| Average customer rating | 3.60 |

### Key findings
- **Fulfilment is the biggest issue.** 16.5% of orders were cancelled and 18.2% were returned, so 34.7% of orders never became revenue.
- **Electronics and Home drive the business.** Electronics brings in about 36% of realised revenue (₦22.0M) and Home about 27% (₦16.3M). Together they produce roughly 62% of gross profit.
- **Margins are tight and consistent.** Every category sits between 31.4% and 32.3%, so profit follows revenue rather than mix.
- **Regional revenue is evenly spread** (about ₦5.5M to ₦6.9M per region). Abuja, Enugu and Edo lead on revenue. Kano (33.5%), Kaduna (33.1%) and Delta (32.9%) have the best margins, while Oyo (30.5%) and Lagos (30.7%) have the lowest. Edo has the highest average order value (₦65,064).
- **New customers contribute the most revenue** (₦26.9M from 533 delivered orders), but Returning customers have a higher average order value (₦53,952 vs ₦50,460).
- **Discounts did not visibly erode margin.** Margin stayed between 31.3% and 32.2% across the 0%, 1 to 15% and 16 to 25% discount bands, and no orders carried a discount above 25%.
- **31.6% of delivered orders arrived late** (371 of 1,175). Late orders averaged a slightly higher rating (3.70) than on-time ones (3.56), which is counter-intuitive and worth investigating.

---

## Project 2: Executive Sales Dashboard

**Folder:** `02-executive-sales-dashboard`
**Type:** Executive reporting workbook
**Tools:** Excel (SUMIFS, SUMIF, COUNTIF, COUNTIFS, IFERROR, structured table references), 4 charts, KPI cards

### The brief
Turn the same cleaned order data into a single executive view and give management clear recommendations. Where Project 1 looks at **realised (delivered) revenue**, this dashboard reports on **all 1,800 booked orders**, so its headline figures are higher.

### Dashboard headline KPIs (all orders)

| KPI | Value |
|-----|-------|
| Total revenue | ₦94,966,142.60 |
| Total cost | ₦64,775,648.38 |
| Gross profit | ₦30,190,494.22 |
| Gross margin | 31.79% |
| Orders | 1,800 |
| Units sold | 2,945 |
| Average order value | ₦52,758.97 |

### What the dashboard contains
- **KPI cards** for revenue, cost, gross profit, margin, orders and units sold
- **Top 7 products by revenue** (led by Power Bank, Smart Watch and Wireless Earbuds)
- **Top 7 regions by revenue** (led by Abuja, Enugu and Edo)
- **Category revenue, cost and profit** chart
- **Monthly revenue and profit trend** chart
- **Product gross profit and revenue** chart
- **Regional revenue** chart
- A written **Recommendation** panel for management

![Executive dashboard preview](./02-executive-sales-dashboard/dashboard-preview.jpg)

### Recommendations given to management
1. **What to sell more of:** push Electronics-led winners and the strongest individual profit contributors, and use bundles and cross-selling rather than deeper discounts.
2. **What to improve:** protect margin on high-revenue, low-margin items. Review supplier cost, promotional depth, bundle structure and price floors, because the 15% to 25% discount bands are not clearly buying better margin.
3. **Where to focus:** scale Abuja, Enugu and Edo for volume. Use Kano, Kaduna and Delta as margin benchmarks. Lagos and Oyo need margin improvement before aggressive expansion. Keep retaining and upselling Returning customers while continuing to acquire New ones.
4. **What to investigate operationally:** the 34.7% non-delivery rate (returns plus cancellations) is the biggest structural problem, especially for high-volume product and region combinations. September's revenue dip should also be broken down by product, channel, region and order status.
5. **What to track weekly:** delivered revenue, delivered gross profit, gross margin, cancellation rate, return rate, discount levels, marketing cost per delivered order, and top and bottom product contribution. The goal is **profitable delivered growth, not raw order volume.**

### Workbook sheets

| Sheet | Purpose |
|-------|---------|
| DASHBOARD | The executive dashboard |
| Cleaned Data | 1,800 cleaned orders, including Cost and Gross Profit columns |
| Monthly Trend | Revenue, cost, gross profit, margin and units by month |
| Category & Region | Revenue, cost, profit, margin and units by category and by region |
| Product Profitability | The same measures per product |
| Segment & Discount | Performance by customer segment and by discount band |
| Data Quality | Log of every data issue, how it was treated and why |
| KPI | The headline KPI table that feeds the dashboard |

### Data quality log

| Issue | Count | Treatment | Why |
|-------|-------|-----------|-----|
| Duplicate Order_ID rows | 12 | Removed exact duplicates, kept the first occurrence | Prevents double-counting orders, revenue, cost and profit |
| Missing Region | 10 | Replaced with `Unknown` | Keeps the transaction without inventing a region |
| Missing Acquisition Channel | 7 | Replaced with `Unknown` | Keeps totals complete |
| Missing Customer Rating | 314 | Left blank, treated as missing and not zero | Avoids distorting satisfaction metrics |

---

## Project 3: Week 3, Advanced Functions (Exercises 02 to 05)

**Folder:** `03-week-3-advanced-functions`
**Files:** `Exercise_02-1.xlsx`, `Exercise_03.xlsx`, `Exercise_04.xlsx`, `Exercise_05.xlsx`
**Focus:** Modern lookup, filtering, statistical and rounding functions

| Exercise | Sheet | Functions practised | What I built |
|----------|-------|--------------------|--------------|
| **02** | UNIQUE | `UNIQUE`, `SORT`, `COUNTA` | Extracted the unique months from a list, sorted them and counted the distinct values |
| | FILTER & SORT | `FILTER`, `SUMIF` | Filtered sales rows for a selected agent and totalled that agent's revenue |
| | XLOOKUP | `XLOOKUP` | Looked up a selected agent's team, revenue and profit |
| | RAND | `RANDBETWEEN` | Generated random revenue and profit figures for practice data |
| **03** | 2-WAY XLOOKUP | nested `XLOOKUP` | Returned the value at the intersection of a chosen month and team |
| | CHOOSE | `CHOOSE`, `RANDBETWEEN` | Assigned each team member a random colour from a list |
| | SWITCH | `SWITCH` | Converted job ratings 1 to 4 into grades (Excellent, Very Good, OK, Poor) |
| **04** | MEDIAN AND MODE | `AVERAGE`, `MEDIAN`, `MODE`, `IFERROR` | Summarised monthly quarterly figures and returned "No Mode" when no value repeated |
| | LARGE and SMALL | `LARGE`, `SMALL`, `INDEX`, `MATCH` | Found the top 3 and bottom 3 sales figures and the product behind each |
| | RANK | `RANK` | Ranked students' scores from highest to lowest |
| **05** | ROUNDING | `ROUND` | Calculated invoice totals including 15% sales tax, rounded to 2 decimals |
| | MROUND | `MROUND` | Calculated a 4% salary bonus rounded to the nearest 50 |
| | AGGREGATE | `AGGREGATE` | Totalled and averaged a price list that contains error values by telling the function to ignore errors |

> **Compatibility note:** `UNIQUE`, `FILTER`, `SORT` and `XLOOKUP` need Excel 365/2021 or later, or Google Sheets. In older versions of Excel those cells can show `#NAME?`. If that happens, open the file in Google Sheets. The `AGGREGATE` sheet also contains a few deliberate errors (for example a division by zero) to demonstrate how the function skips them.

---

## Project 4: Core Excel Exercises 01 to 05 (Screenshots)

**Folder:** `04-excel-exercises-01-05`
**Date captured:** 7 August 2026
**Tool:** Google Sheets (mobile)
**Format:** 9 screenshots of the completed exercise sheets

| Exercise | Topic | What the screenshots show | Files |
|----------|-------|---------------------------|-------|
| **01** | Logical Functions | Olympic-style athlete table (country, gender, medal, prize money) with a medal-to-prize table, then criteria questions such as athletes from the USA, Australian women, UK male silver medallists, and prize money by country group | `Screenshot_20260807-151131.png`, `Screenshot_20260807-151154.png` |
| **02** | Lookup Functions | Marathon athlete table (bib number, event, route, finishing position) with a lookup panel that returns the bib number, route and position for a chosen athlete | `Screenshot_20260807-151218.png` |
| **03** | Working with formulas and functions in calculations | Customer order sheet for 50 customers with line total, 4.5% sales tax and total including tax, plus a summary of total sales, average quantity, minimum and maximum price and number of customers | `Screenshot_20260807-151238.png`, `Screenshot_20260807-151255.png` |
| **04** | Naming ranges and using named ranges in formulas | The same customer data revisited for the named-ranges exercise, with the sales tax changed to 6% | `Screenshot_20260807-151305.png`, `Screenshot_20260807-151317.png` |
| **05** | Applying number formats and formatting | A client invoice sheet with date formatting, currency formatting and a 15% sales tax column | `Screenshot_20260807-151324.png`, `Screenshot_20260807-151342.png` |

---

## Skills demonstrated

- **Data cleaning:** duplicate removal, missing-value handling and a documented data quality log
- **Analysis formulas:** SUMIFS, SUMIF, COUNTIFS, COUNTIF, AVERAGEIFS, SUMPRODUCT, IFERROR
- **Lookup and dynamic functions:** XLOOKUP (including 2-way), INDEX/MATCH, UNIQUE, FILTER, SORT, CHOOSE, SWITCH
- **Statistics and rounding:** AVERAGE, MEDIAN, MODE, LARGE, SMALL, RANK, ROUND, MROUND, AGGREGATE
- **Dashboard design:** KPI cards, charts, top-N panels and written management recommendations
- **Business thinking:** separating realised revenue from booked revenue, measuring revenue at risk, and turning findings into actions
- **Spreadsheet practice:** named ranges, number and date formatting, formula-driven workbooks

## How to explore this repository

1. Start with **Project 1** and open its `Dashboard` sheet, then read the `README` sheet inside the workbook.
2. Open **Project 2** to see the executive dashboard and its recommendations.
3. Browse **Projects 3 and 4** to see the function-by-function exercise work.

## Notes
- All monetary values in Projects 1 and 2 are in Nigerian naira (₦). Project 4 exercises use the dollar format from the exercise templates.
- Projects 1 and 2 use the same cleaned dataset. Their figures differ because Project 1 counts delivered orders only, while Project 2 reports on all booked orders.

---

*Prepared by Clinton Onezi Aiyede, King's House, BoyCode Africa.*
