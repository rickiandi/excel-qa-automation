# Excel QA Automation — Data-Entry Freelance

Portfolio project focused on Excel quality assurance for freelance data-entry workflows.

## Workflow

Browse Excel → choose QA Profile → Analyze → review findings → Safe Fix → Export Report.

## Built-in profiles

- General Data Entry
- Customer Database
- Product Catalog
- Sales Data
- Inventory

## Main checks

Duplicate rows, duplicate IDs, missing values, empty/duplicate headers, mixed types, invalid email,
phone pattern, dates, numbers stored as text, negative values, outliers (IQR), whitespace,
formatting inconsistency, formula errors and broken #REF! references, hidden sheets and merged cells.

## Workflow features

- Reusable QA profiles
- QA score
- Audit history stored locally
- Reviewable findings with severity
- Safe Fix output copy
- Excel report
- Keyword command filter
- Original file is never modified

## Rules / limitations

- An outlier is only a potential anomaly, not proof of an error.
- Business correctness requires business rules.
- openpyxl does not recalculate Excel formulas.
- XLSM macros are not executed.
- Phone validation is a pattern check, not verification that a number is active.
- Command input is keyword-based; it is not a general LLM.
- Safe Fix only targets selected low-risk transformations.

## Build Windows EXE

Install Python 64-bit, then run `build_exe.bat`. The executable will be in `dist/ExcelQA_DataEntry.exe`.

## Portfolio note

This project is designed to demonstrate Python automation, Excel data quality, QA workflow design,
report generation, reusable rules and safe data-cleaning practices in a freelance data-entry context.
