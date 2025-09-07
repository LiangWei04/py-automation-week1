python src/csv_to_excel_report.py \
  --inputs data/sales_q1.csv data/sales_q2.csv data/sales_q3.csv data/sales_q4.csv \
  --out reports/sales_report.xlsx

Go to the folder name reports, there'll be an Excel file. Open the excel file and check the following: 
Sheet Transactions (cleaned rows + computed revenue = units*unit_price),

Sheet Summary (pivot by region and product: sum units and revenue),

A bar chart on Summary showing revenue by region,

Basic formatting (bold headers, thousands separators, currency for revenue).