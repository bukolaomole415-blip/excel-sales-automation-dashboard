# excel-sales-automation-dashboard
Data cleaning, conditional formatting error handling, and VBA macro automation in Excel.

# 📊 Excel Sales Data Cleaning, Automation & Dashboard Simulation

## 🚀 Project Overview
This repository contains an end-to-end data processing, conditional formatting, error handling, and macro-automation project completed in Microsoft Excel. The project transforms raw sales account data across multiple years into a clean, automated summary report.

## 🛠️ Key Technical Features & Implementation
* **Advanced Conditional Formatting:** 
  * Implemented custom logical formulas (`=AND(ISNUMBER(...))`, `=ISERROR(...)`) to dynamically format numeric data while preventing text strings and formula errors from breaking format rules.
  * Applied targeted highlights for accounts missing Product 2 cross-sells, missing data values, top 10% unit sales, high CAGR growth ($\ge 100\%$), and negative CAGR.
* **Complex Formulas & Error Handling:** 
  * Utilized `IFERROR` alongside multi-year Compound Annual Growth Rate (CAGR) calculations (`(L2/H2)^(1/5)-1`) with clean error replacements (`"Formula Error"`, `"Check data"`).
  * Executed category aggregations using `COUNTIF` and `SUMIF`.
* **VBA Automation:** 
  * Developed custom Visual Basic for Applications (VBA) macro routines (such as the `UnitSales` macro) to automate multi-column sorting and workflow reporting.
* **Documentation & Visualization:** 
  * Built dynamic summary tables and packaged complete macro logs and visual screenshots into an executive-ready PDF report.

## 📸 Project Screenshots

### 1. Formatted Workbook Overview
![Workbook Overview](screenshots/dashboard_overview.png)

### 2. VBA Code Automation Module
![VBA Code Module](screenshots/vba_code_module.png)

## 📂 Repository Contents
* `Account_Sales_Data_for_Automation.xlsm`: Main macro-enabled Excel workbook containing all data sheets, conditional formatting rules, PivotTables, and VBA code modules.
* `excel_macro_automation_report.pdf`: Executive documentation report showcasing workbook screenshots and technical VBA code logs.
