# py-automation-week1

Core focus: **CSV/Excel automation with Python** + **Git basics**.  
This repo contains two scripts:

1. `csv_to_excel_report.py` → Combines multiple sales CSVs into a single **Excel report** with:
   - Transactions sheet (all rows + revenue column)
   - Summary sheet (KPIs, pivot by region/product, totals)
   - Bar chart of Revenue by Region
2. `merge_excels.py` → Merges multiple Excel files (with identical headers) into one consolidated sheet.

---

## 📦 Installation

```bash
git clone https://github.com/<your-username>/py-automation-week1.git
cd py-automation-week1
python -m venv .venv
# On Linux/Mac
source .venv/bin/activate
# On Windows
.venv\Scripts\activate
pip install -r requirements.txt
