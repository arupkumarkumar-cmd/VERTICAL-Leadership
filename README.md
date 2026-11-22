# DS Weekly Operations Dashboard – Automated Python Application

An automated, reusable Python application for generating the Data Services (DS) Team Weekly Operations Reporting Dashboard from DART export reports.

---

## 1. Overview & Architecture

The application reconstructs full historical ticket movements from raw DART CSV exports, calculates point-in-time snapshot balances (opening and closing), analyzes weekly throughput and movements (new, completed, and returned tickets), segments open ticket aging across standard buckets, tracks SLA compliance, and populates an executive-ready Excel workbook formatted according to leadership templates.

### Directory Structure

```text
Vertical_leadership_Latest/
│
├── input/                  <-- Place weekly DART CSV exports here (e.g., DART_REPORT_new.csv)
│   └── DART_REPORT_new.csv
│
├── template/               <-- Master Excel template (never overwritten)
│   └── Weekly_Operations_Report_11052026 - New Template.xlsx
│
├── output/                 <-- Generated weekly Excel reports are saved here
│   └── DS_Weekly_Operations_Dashboard_YYYYMMDD_YYYYMMDD.xlsx
│
├── logs/                   <-- Execution logs and audit trail
│   └── processing_log.txt
│
├── config/                 <-- User configuration settings
│   └── config.json
│
├── src/                    <-- Application modular source code
│   ├── __init__.py
│   ├── main.py             <-- Execution orchestrator and CLI interface
│   ├── data_loader.py      <-- CSV file discovery, loading, and encoding handling
│   ├── ticket_processor.py <-- DS ticket filtering, sorting, and master mapping
│   ├── movement_analyzer.py<-- Historical snapshot and weekly movement reconstruction
│   ├── aging_analyzer.py   <-- Open ticket aging calculation and category segmentation
│   ├── sla_analyzer.py     <-- SLA compliance evaluation and group breakdown
│   ├── dashboard_calculator.py <-- Consolidated KPI payload assembler
│   ├── excel_generator.py  <-- OpenPyXL multi-sheet workbook generation & chart rendering
│   └── validators.py       <-- Date validation, column checks, and data quality audits
│
├── Run_DS_Dashboard.bat    <-- One-click Windows batch launcher
├── requirements.txt        <-- Required Python packages (pandas, openpyxl)
├── instruction.txt         <-- Business logic specifications
└── README.md               <-- User manual and operational guide
```

---

## 2. Weekly Execution Workflow

Running the report every week requires only three simple steps:

### Step 1: Place the Latest DART CSV
Place your weekly DART export file into the `input/` folder:
- Path: `D:\ARUP\WorkSpace\Python_Codes\Vertical_leadership_Latest\input\`
- If exactly one CSV is present, it is selected automatically.
- If multiple CSV files are present, the application presents a numbered selection list.

### Step 2: Double-Click the Batch File
Double-click `Run_DS_Dashboard.bat` in the project root directory.
- The script automatically detects your Python environment (virtualenv or Anaconda), verifies prerequisites, and launches the application.

### Step 3: Enter the Reporting Date Range
When prompted in the console window, enter the start and end dates:

```text
==========================================
       DS WEEKLY OPERATIONS DASHBOARD     
==========================================

Enter Reporting Period (or 'Q' to quit):

Start Date (DD-MM-YYYY): 05-11-2026
End Date (DD-MM-YYYY)  : 11-11-2026

Selected Period:
  05-Nov-2026 to 11-Nov-2026

Proceed with processing? (Y/N): Y
```

Type `Y` and press `Enter`. The application runs through all 7 stages and opens the console completion banner when done.

---

## 3. Output Worksheets Reference

The generated Excel file is saved in the `output/` directory:
- Filename format: `DS_Weekly_Operations_Dashboard_YYYYMMDD_YYYYMMDD.xlsx`

The generated workbook contains 7 dedicated worksheets:

| Sheet Name | Contents |
| :--- | :--- |
| **`DS_Dashboard`** | Executive summary populated from the master template. Contains the weekly ticket movement table, open ticket aging matrix (Maintenance vs New Format), aging bar chart, New Format project tracking, SLA compliance by group (FP, EP, COM), and SLA bar chart. |
| **`DS_Ticket_Tracker`** | Master inventory with one row per distinct `TRACKER ID` that touched DS. Includes first DS assignment date/user, latest group/user, current ownership status, aging days, category, and SLA breach status. |
| **`DS_Aging_Analysis`** | Detailed audit of all open tickets grouped by their aging bucket (`<2 days`, `2-4 days`, `4-7 days`, `2 weeks`, `3 weeks`, `4 weeks`, `More than 1 month`) and category. |
| **`DS_Weekly_Movement`** | Reconciled breakdown of each KPI with complete lists of ticket numbers: Opening Balance tickets, New Tickets, Completed Tickets, Returned Tickets, and Closing Balance tickets. |
| **`DS_Movement_Audit`** | Chronological audit trail of all historical group and user handoffs for every ticket, tracking movement types (e.g. `First DS Assignment`, `Transferred to COM`, `Returned to DS`). |
| **`DART_Raw_Data`** | Preserved raw data rows from the source CSV for all tickets that touched DS, maintaining the original column structure for ad-hoc queries. |
| **`DS_Data_Quality`** | Comprehensive integrity check report documenting any missing fields, duplicate history rows, invalid dates, unparseable values, or reconciliation variances. |

---

## 4. Advanced Execution (CLI Options)

For automated scheduling, command-line arguments can be provided:

```bash
# Explicit dates in DD-MM-YYYY format
python -m src.main --start-date 05-11-2026 --end-date 11-11-2026

# Non-interactive mode (ideal for scheduled tasks)
python -m src.main -s 05-11-2026 -e 11-11-2026 --non-interactive

# Specify custom input CSV and custom config
python -m src.main -s 05-11-2026 -e 11-11-2026 -i "input\custom_dart.csv" -c "config\config.json"
```

---

## 5. Configuration Settings

Settings can be customized without editing Python code in `config/config.json`:

```json
{
  "project_name": "DS Weekly Operations Dashboard",
  "paths": {
    "input_dir": "input",
    "output_dir": "output",
    "template_path": "template/Weekly_Operations_Report_11052026 - New Template.xlsx",
    "log_dir": "logs"
  },
  "team_identifier": "DS",
  "keywords": {
    "new_format": "New Format",
    "sla": "SLA"
  },
  "date_settings": {
    "input_date_format": "%d-%m-%Y",
    "display_date_format": "%d-%b-%Y",
    "excel_date_format": "DD-MMM-YYYY",
    "allow_future_dates": false
  },
  "aging_buckets": [
    { "name": "<2 days", "min_days": 0, "max_days": 1 },
    { "name": "2-4 days", "min_days": 2, "max_days": 3 },
    { "name": "4-7 days", "min_days": 4, "max_days": 6 },
    { "name": "2 weeks", "min_days": 7, "max_days": 13 },
    { "name": "3 weeks", "min_days": 14, "max_days": 20 },
    { "name": "4 weeks", "min_days": 21, "max_days": 27 },
    { "name": "More than 1 month", "min_days": 28, "max_days": null }
  ]
}
```

---

## 6. Troubleshooting & FAQs

### Problem: "No input CSV file found"
- **Solution**: Ensure your DART export CSV file is placed in `Vertical_leadership_Latest\input\`.
- Make sure the file extension is `.csv`.

### Problem: "Missing required columns"
- **Solution**: The input CSV must contain standard DART columns:
  `TRACKER ID`, `TRACKER HIST ID`, `SHORT_DESCRIPTION`, `ASSIGNED_TO_GROUP`, `ASSIGNED TO USER`, `CREATE_DT`, `FOLLOW_UP_DT`, `UPDATE_DT`, `TRACKER STATUS`, `TRACKER STATUS HISTORY`.
  Check that the export was not truncated.

### Problem: "Invalid date format"
- **Solution**: Enter dates in `DD-MM-YYYY` format (e.g., `05-11-2026`). Slashes (e.g., `05/11/2026`) and ISO format (`2026-11-05`) are also supported.

### Problem: "Future reporting dates are not permitted"
- **Solution**: The application by default restricts reporting dates beyond today to avoid reporting future data. If you are running simulation tests on future data, set `"allow_future_dates": true` in `config/config.json`.

### Problem: "How is the single Closing Balance calculated?"
- **Solution**: The application uses the authoritative snapshot of open DS tickets as of the report end date (excluding any tickets with `TRACKER STATUS == 'Completed'`). The movement equation strictly reconciles: `Opening Balance + New Tickets + Returned Tickets - Tickets Closed this Week = Closing Balance`. There is only one authoritative closing balance presented to leadership.

### Checking Logs
All executions append detailed diagnostic and audit messages to:
`Vertical_leadership_Latest\logs\processing_log.txt`
