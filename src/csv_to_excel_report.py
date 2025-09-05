"""
Path: src/csv_to_excel_report.py

Purpose: Read multiple CSVs, combine, compute revenue, produce a styled Excel report with:

Sheet Transactions (cleaned rows + computed revenue = units*unit_price)

Sheet Summary (pivot by region and product: sum units and revenue)

A bar chart on Summary showing revenue by region

Basic formatting (bold headers, thousands separators, currency for revenue)

CLI spec (must work exactly):

python src/csv_to_excel_report.py \
  --inputs data/sales_q1.csv data/sales_q2.csv data/sales_q3.csv data/sales_q4.csv \
  --out reports/sales_report.xlsx


Acceptance criteria:

Creates reports/sales_report.xlsx

Transactions has all rows from inputs + revenue column (numeric)

Summary pivot exact columns: region, product, units_sum, revenue_sum

A bar chart titled “Revenue by Region” using revenue_sum aggregated by region

Number formats: units_sum as integer; revenue/revenue_sum as currency

No exceptions on missing optional columns (but fail if required headers missing)

Implementation notes (pseudocode-level, exact ops to perform):

Read CSVs with pd.read_csv, pd.concat (ignore_index=True)

Validate required cols: {"date","region","product","units","unit_price"}

Cast units to Int64, unit_price to float, compute revenue = units*unit_price

Build pivot with groupby(["region","product"]).agg({"units":"sum","revenue":"sum"})

Write with pd.ExcelWriter(engine="openpyxl")

Use openpyxl to:

bold header rows,

apply number formats: #,##0 for units; "$"#,##0.00 for revenue,

insert BarChart on Summary (X = regions, Y = revenue_sum)
"""
import csv

