python src/merge_excels.py \
  --inputs data/regions.xlsx data/regions.xlsx \
  --sheet Sheet1 \
  --out data/merged_output/regions_merged.xlsx

Go to data folder, then merged_output folder. There'll be an Excel file, open it and check the following:
Creates data/merged_output/regions_merged.xlsx,

One sheet Merged with all rows appended (headers once),

Validates that requested --sheet exists in every input,

Fails with clear error if headers mismatch.