'''
Path: src/merge_excels.py

Purpose: Merge multiple Excel files (identical headers) into a single Excel, one sheet.

CLI spec (must work exactly):

python src/merge_excels.py \
  --inputs data/regions.xlsx data/regions.xlsx \
  --sheet Sheet1 \
  --out data/merged_output/regions_merged.xlsx


Acceptance criteria:

Creates data/merged_output/regions_merged.xlsx

One sheet Merged with all rows appended (headers once)

Validates that requested --sheet exists in every input

Fails with clear error if headers mismatch

Implementation notes:

For each input: pd.read_excel(file, sheet_name=sheet)

Confirm equal ordered headers across inputs

pd.concat → write to Excel (Merged)
'''
import argparse
import pandas as pd
from pathlib import Path
import sys

def arg_parse():
    p = argparse.ArgumentParser(description="Merge multiple Excel files into one")
    p.add_argument("--inputs", nargs="+", required=True,
                   help="One or more input CSVs (e.g., --inputs data/q1.csv data/q2.csv)")
    p.add_argument("--out", required=True,
                   help="Output Excel file (e.g., reports/sales_report.xlsx)")
    p.add_argument("--sheet", type=str, required=True, help="Select which sheet to merge")
    return p.parse_args()

def main():
    args = arg_parse()
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    all_part = []
    expected_headers = None
    for f in args.inputs:
        try:
            df_part = pd.read_excel(f, sheet_name=args.sheet)
        except ValueError:
            sys.exit(f"Error: Sheet '{args.sheet}' not found in {f}")

        df_part.columns = [c.strip().lower() for c in df_part.columns]

        # Header validation
        if expected_headers is None:
            expected_headers = list(df_part.columns)
        else:
            if list(df_part.columns) != expected_headers:
                sys.exit(f"Error: Header mismatch in file {f}. "
                         f"Expected {expected_headers}, got {list(df_part.columns)}")

        all_parts.append(df_part)
    df_merged = pd.concat(all_part, ignore_index=True)
    df_merged.to_excel(out_path, index=False)
    print("Merging...")

if __name__ == "__main__":
    main()