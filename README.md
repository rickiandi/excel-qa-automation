# Excel QA Automation — Data-Entry Freelance

A Python-based Excel Quality Assurance (QA) automation tool designed for freelance data-entry workflows.

The application helps identify common Excel data-quality issues, review findings by severity, apply selected low-risk fixes, maintain audit history, and generate QA reports.

## Demo Result

The application was tested using a product catalog dataset containing 120 product records.

Example QA workflow:

1. Load an Excel workbook
2. Select a QA profile
3. Run Analyze
4. Review detected findings
5. Apply Safe Fix where appropriate
6. Run Analyze again
7. Export the QA report

Example final result:

**QA Score: 100/100**  
**Findings: 0**  
**Critical: 0**  
**High: 0**  
**Medium: 0**  
**Low: 0**

This demonstrates the complete QA → fix → re-check workflow.

## Workflow

Browse Excel → choose QA Profile → Analyze → review findings → Safe Fix → Export Report.

## Built-in Profiles

- General Data Entry
- Customer Database
- Product Catalog
- Sales Data
- Inventory

## Main QA Checks

The application can check for:

- Duplicate rows
- Duplicate IDs
- Missing values
- Empty headers
- Duplicate headers
- Mixed data types
- Invalid email patterns
- Phone number patterns
- Invalid dates
- Numbers stored as text
- Negative values
- Potential outliers using IQR
- Leading/trailing whitespace
- Formatting inconsistencies
- Formula errors
- Broken `#REF!` references
- Hidden sheets
- Merged cells

## Main Features

- Reusable QA profiles
- Automated QA scoring
- Findings grouped by severity
- Reviewable cell-level findings
- Audit history
- Safe Fix output copy
- Excel QA report generation
- Keyword-based command filter
- Original workbook is never modified

## Example QA Workflow

A typical data-entry quality-control process:

**Before Fix**

Detected issues may include:

- Duplicate records
- Missing stock values
- Mixed data types
- Formatting inconsistencies

**After Fix**

The workbook can be analyzed again to verify whether the selected issues have been resolved.

The goal is not simply to modify the spreadsheet, but to provide a repeatable QA process with traceable results.

## Rules and Limitations

- An outlier is only a potential anomaly, not proof of an error.
- Business correctness requires business rules.
- `openpyxl` does not recalculate Excel formulas.
- XLSM macros are not executed.
- Phone validation is a pattern check, not verification that a number is active.
- Command input is keyword-based; it is not a general LLM.
- Safe Fix only targets selected low-risk transformations.

## Build Windows EXE

Install Python 64-bit and run:

`build_exe.bat`

The executable will be generated at:

`dist/ExcelQA_DataEntry.exe`

## Technology

- Python
- Tkinter
- openpyxl
- PyInstaller
- Excel / XLSX

## Project Purpose

This project demonstrates practical automation skills for freelance data-entry and data-quality workflows.

It focuses on:

- Python automation
- Excel processing
- Data validation
- QA workflow design
- Safe data cleaning
- Auditability
- Automated reporting
- Reusable validation rules

## Portfolio Note

This project was built as a practical automation portfolio project to demonstrate how repetitive Excel quality-control tasks can be standardized and automated.
