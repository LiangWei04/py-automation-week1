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