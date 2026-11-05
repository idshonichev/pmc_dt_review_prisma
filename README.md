# The Digital Twin of Polymer Matrix Composites—From Non-Destructive Testing to Property Prediction: A Scoping Review

This repository contains the supplementary materials for the PRISMA-ScR scoping review paper.

Content:

1. Data files in `./data/`:

   - `./data/raw_export` - directory with the raw database export files
   - `./data/raw_export/scopus_export.csv` - Scopus export csv file
   - `./data/raw_export/scopus_export_fixed.xlsx` - xlsx file created with `./scripts/fix_scopus.py` from the Scopus export csv file with correct data split and column titles
   - `./data/raw_export/wos_export.xlsx` - WoS export file
   - `./data/screening.xlsm` - screening workbook with the inclusion/exclusion decision for every screened record: year, name, DOI, decision, reason, class and classification rationale
   - `./data/scopus_wos_merged.xlsx` - merged workbook created with `./scripts/merge_scopus_wos.py` from the fixed Scopus and the WoS exports

2. Helper scripts:

   - `./scripts/fix_scopus.py` - repairs the malformed Scopus CSV export (broken quoting / field boundaries) and writes the cleaned table to `.xlsx` or `.csv` based on the output extension. This is how `scopus_export_fixed.xlsx` was produced from `scopus_export.csv`:

     ```
     python scripts/fix_scopus.py data/raw_export/scopus_export.csv data/raw_export/scopus_export_fixed.xlsx
     ```

   - `./scripts/merge_scopus_wos.py` - combines the Scopus and WoS exports into a single Excel workbook with deduplication: papers are matched across the databases by DOI first, then by title similarity (with year / first-author boosts). The result contains four sheets: `Combined_Data` (unique papers), `Scopus_Original` (all Scopus records), `WoS_Original` (all WoS records) and `Statistics`:

     ```
     python scripts/merge_scopus_wos.py data/raw_export/scopus_export_fixed.xlsx data/raw_export/wos_export.xlsx -o data/scopus_wos_merged.xlsx
     ```

Both scripts can also be run with no arguments: missing file paths are prompted for interactively, and for Excel inputs the sheets are listed so one can be picked by number or name (Enter keeps the default). In an interactive terminal the path prompts support Tab completion of file and directory names; piped runs skip all prompts and fall back to defaults (first sheet, `<input_stem>_fixed.xlsx` / `<scopus_stem>_combined.xlsx`).

Run `python scripts/<script>.py --help` for the full list of options (encoding, sheet overrides, title similarity threshold, etc.).

