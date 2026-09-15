# Retail Sales Performance Analysis

**Business question:** Which regions, categories, and channels are driving revenue — and
where is the business underperforming?

A full year (Oct 2024 – Sep 2025) of ~5,200 retail transactions, modeled as a real Excel
workbook: live formulas (`SUMIF`, `COUNTIF`, `IFERROR`), a region/category/channel summary,
and 3 charts — built to be recreated directly as a Power BI dashboard.

## Key findings

- **South region generates $165K in revenue — less than half of North's $347K** — despite
  running the same product mix and channel split. This is the single biggest revenue gap in
  the business.
- Profit margin is nearly identical across all four regions (~38.4-38.6%), which rules out
  pricing or cost discipline as the explanation — **the South gap is a volume problem, not a
  margin problem.**
- Revenue shows a clear seasonal pattern: **Nov-Dec holiday months run ~55% above baseline**,
  while Jun-Jul and Jan-Feb dip 15-20% below it — relevant for inventory and staffing
  planning.
- **Online is the leading channel**, followed by In-Store, then Mobile App — useful for
  where to prioritise UX or marketing investment.

## Recommendation

Investigate why South underperforms on volume specifically (store count, local marketing
spend, staffing) rather than pricing — since margin is healthy there, the fix is demand
generation, not discounting.

## What's inside

- `retail_sales_performance.xlsx` — the deliverable: Raw Data sheet (5,212 transactions) +
  Summary sheet (region/category/channel tables with live formulas + 2 bar charts) +
  Monthly Trend sheet (line chart)
- `raw_sales_data.csv` — the source data
- `generate_data.py` / `build_workbook.py` — how it was built

## How it was built

```
pip install pandas openpyxl
python generate_data.py     # synthesises a realistic year of transactions
python build_workbook.py    # builds the Excel workbook with formulas + charts
```

**Note on the dataset:** synthesised for this project (real retail data is confidential),
with a deliberately underperforming region and realistic seasonality built in, so the
findings reflect genuine pattern-detection rather than random noise.

## Recreating this as a Power BI dashboard

The same Summary tables map directly onto Power BI visuals:
1. Import `raw_sales_data.csv` as the data source
2. Region/Category/Channel tables → clustered bar charts
3. Monthly Trend → a line chart with a date-hierarchy slicer
4. Add a region slicer to let a reviewer filter into the South gap live

## Tools

Excel (formulas, pivot-style summary tables) · Python (openpyxl, pandas) for generation ·
translatable directly to Power BI / DAX

## Author

Suhana Shaikh · [Portfolio](https://shaikhsuhana.github.io/suhana-portfolio) ·
[GitHub](https://github.com/shaikhsuhana)
