#  Employee Attendance Tracking & Reporting System

> A reusable, automated monthly attendance tracking solution built with **Microsoft Excel VBA** and **SAP Crystal Reports** — designed for HR teams to clean, process, and report employee attendance data with minimal manual effort every month.

---

##  Table of Contents

- [Overview](#overview)
- [Project Structure](#project-structure)
- [Tools & Requirements](#tools--requirements)
- [Workflow](#workflow)
- [VBA Cleaning Macro](#vba-cleaning-macro)
- [Crystal Reports Template](#crystal-reports-template)
- [Report Output](#report-output)
- [Known Issues & Notes](#known-issues--notes)
- [Author](#author)

---

## Overview

This project automates the monthly process of cleaning raw employee attendance data exported from an attendance system and generating a professional printed report showing:

- Total attendance records per employee
- Employee with the **highest** attendance
- Employee with the **lowest** attendance
- A **bar chart** of attendance per employee

The solution is built to be **reusable every month** — only the data file changes. The VBA macro and Crystal Reports template stay the same.

---

## Project Structure

```
Employee_Attendance/
│
├── Data/
│   ├── Raw/                        # Raw attendance exports (one per month)
│   │   ├── Original Records Report26.csv
│   └── Cleaned/                    # Cleaned Excel files ready for Crystal Reports
│       ├── Employee Attendance Records Workbook.xlsx
│
├── Macro/
│   └── AttendanceAutomation.bas    # VBA cleaning macro (import into Excel)
│
├── Report/
│   └── EMPLOYEE ATTENDANCE REPORT FOR NOVEMBER 2025.rpt      # Crystal Reports template (reusable)
│
└── README.md
```

---

## Tools & Requirements

| Tool | Version | Purpose |
|---|---|---|
| Microsoft Excel | 2016 or later | Data cleaning via VBA macro |
| SAP Crystal Reports | 2020 or later | Report design and generation |
| Microsoft ACE OLEDB 12.0 | Installed with MS Office | Excel-to-Crystal data connection |

---

## Workflow

The end-to-end process follows these stages every month:

```
Raw CSV Export
      ↓
Open in Excel → Run VBA Macro (clean data)
      ↓
Save as .xlsx
      ↓
Open Attendance_Report.rpt in Crystal Reports
      ↓
Update data source path → Enter Month & Year when prompted
      ↓
Preview → Export as PDF
```

---

## VBA Cleaning Macro

The macro (`AttendanceAutomation.bas`) performs the following cleaning steps on the raw data:

### What It Does

| Step | Action |
|---|---|
| 1 | Inserts column headers — `Employee Code`, `Employee Name`, `Date Present` |
| 2 | Strips the time component from DateTime values, keeping date only |
| 3 | Trims leading, trailing and middle spaces from Employee Code and Name columns |
| 4 | Formats the date column as `DD/MM/YYYY` |

---

## Crystal Reports Template

The `Attendance_Report.rpt` file is a reusable template that:

- Prompts for **Month** (1–12) and **Year** at runtime — no hardcoding
- Filters records to the selected month and year automatically
- Groups records by **Employee Name**
- Counts total attendance records per employee
- Displays highest and lowest attendance employee using Crystal variable formulas
- Renders a **bar chart** of attendance per employee
- Includes company logo and formatted header

### Crystal Reports Formula Reference

| Formula | Location | Purpose |
|---|---|---|
| `ResetVars` | Report Header | Initialises max/min tracking variables |
| `CaptureMinMax` | Group Footer | Updates highest/lowest as each group prints |
| `ShowHighest` | Report Footer | Displays employee with highest attendance |
| `ShowLowest` | Report Footer | Displays employee with lowest attendance |
| `ShowHighestCount` | Report Footer | Displays highest attendance count |
| `ShowLowestCount` | Report Footer | Displays lowest attendance count |

### Record Selection Formula

```
Month({Sheet1$.Date Present}) = {?Month}
And
Year({Sheet1$.Date Present}) = {?Year}
```

---

## Report Output

The generated report includes:

- **Header** — Company logo, company name, report title with dynamic month and year
- **Data Table** — Employee names and their total attendance count for the month
- **Bar Chart** — Visual comparison of attendance across all employees
- **Summary Section** — Total attendance records, highest and lowest attendance employee with counts
- **Footer** — Report creation date and author name

---

## Known Issues & Notes

- **Mixed date formats in CSV exports** — Always run Text to Columns before the macro. The source system exports dates inconsistently which causes type mismatch errors if skipped.
- **Excel file must be saved as `.xlsx`** — The ACE OLEDB driver used by Crystal Reports does not read legacy `.xls` or `.csv` files reliably.
- **Data source path must be updated monthly** — The `.rpt` file stores an absolute path to the Excel file. Update it via Database → Set Datasource Location each month.
- **Multiple time entries per day** — The source data records multiple entries per employee per day (e.g. different sessions). Each entry is counted individually as an attendance record, not deduplicated by date.

---

## Author

**Ursula Aseye Amoaku**  
