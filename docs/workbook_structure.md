📁 Workbook Structure and Macro Overview

## Sheet-by-Sheet Breakdown

### 1. DATA VISUALIZATION
- Main dashboard interface.
- Compares sales metrics across MONTH 1 to MONTH 4.
- Powered by background formulas and macros.

### 2. ACCESS SHEETS
- Checkboxes to allow/disallow access to individual sheets.
- Tied to VBA event handlers that toggle worksheet visibility.

### 3. SALES RAW DATA
- Detailed sales transaction log.
- Fields include: JE Code, Salesperson, Item, Region, Price, Quantity, Discounts, Sales/Total Value.

### 4. BACKGROUND CALCULATIONS
- Core logic and pricing table for multiple months.
- Used in calculations for charts and visual summaries.
- Notes on VBA functionality:
  - `Hide Unhide Sheets`
  - `Check Uncheck Box`

## Macros and Automation

- **Hide/Unhide Logic**: Controlled by checkboxes on `ACCESS SHEETS`.
- **Macros likely used**:
  - `Worksheet_Change` to detect checkbox value changes.
  - Procedures to `Sheet.Visible = xlSheetVeryHidden` or `xlSheetVisible`.
- Macro file is stored in `vbaProject.bin` for modular use.

---

## Suggestions for Enhancements

- Add slicers or dropdowns for region/store-specific filters.
- Add conditional formatting for discount thresholds.
- Implement user form for easier manual entry.
