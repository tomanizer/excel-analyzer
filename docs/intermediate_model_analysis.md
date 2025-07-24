# Excel Workbook Analysis: intermediate_model.xlsx

*Generated on: 2025-07-24T01:17:51.903758*

## 📊 Executive Summary

- **File Size**: 6.92 KB
- **Sheets**: 3
- **Cells with Data**: 66
- **Formulas**: 17
- **Tables**: 1
- **Charts**: 0
- **Named Ranges**: 0
- **Cross-sheet References**: 0
- **Complexity Score**: 135

## 📋 File Metadata

- **Filename**: intermediate_model.xlsx
- **File Size**: 6.92 KB
- **Last Modified**: 2025-07-24T01:05:53.602401
- **File Type**: .xlsx
- **VBA Enabled**: False
- **Sheet Count**: 3

## 🌐 Global Features

### Document Properties

- **Author**: openpyxl

## 📄 Sheet Analysis

### Sheet: Data

- **Dimensions**: 6 rows × 4 columns
- **Cells with Data**: 24
- **Formulas**: 0
- **Tables**: 1
- **Charts**: 0
- **Data Validations**: 0
- **Merged Cells**: 0

#### Formal Tables

- **ProductTable** (Range: A1:D6)
  - Style: TableStyleMedium9

#### Sample Data

| Cell | Value | Type | Formula |
|------|-------|------|---------|
| A1 | Product ID | str | No |
| B1 | Product Name | str | No |
| C1 | Price | str | No |
| D1 | Category | str | No |
| A2 | 101 | int | No |
| B2 | Widget A | str | No |
| C2 | 25.5 | float | No |
| D2 | Electronics | str | No |
| A3 | 102 | int | No |
| B3 | Widget B | str | No |
| ... | ... | ... | ... | *(showing 10 of 24 cells)* |

### Sheet: Analysis

- **Dimensions**: 8 rows × 6 columns
- **Cells with Data**: 37
- **Formulas**: 15
- **Tables**: 0
- **Charts**: 0
- **Data Validations**: 0
- **Merged Cells**: 0

#### Sample Data

| Cell | Value | Type | Formula |
|------|-------|------|---------|
| A1 | Sales Analysis | str | No |
| A3 | Order ID | str | No |
| B3 | Product ID | str | No |
| C3 | Quantity | str | No |
| D3 | Product Name | str | No |
| E3 | Unit Price | str | No |
| F3 | Total | str | No |
| A4 | 1 | int | No |
| B4 | 101 | int | No |
| C4 | 5 | int | No |
| ... | ... | ... | ... | *(showing 10 of 37 cells)* |

#### Formula Analysis

**Most Common Functions:**

- VLOOKUP: 10 occurrences

**Sample Formulas:**

- **D4**: `=VLOOKUP(B4,Data!A1:D6,2,FALSE)`
- **E4**: `=VLOOKUP(B4,Data!A1:D6,3,FALSE)`
- **F4**: `=C4*E4`
- **D5**: `=VLOOKUP(B5,Data!A1:D6,2,FALSE)`
- **E5**: `=VLOOKUP(B5,Data!A1:D6,3,FALSE)`
- ... *(showing 5 of 15 formulas)*

### Sheet: Summary

- **Dimensions**: 4 rows × 2 columns
- **Cells with Data**: 5
- **Formulas**: 2
- **Tables**: 0
- **Charts**: 0
- **Data Validations**: 0
- **Merged Cells**: 0

#### Sample Data

| Cell | Value | Type | Formula |
|------|-------|------|---------|
| A1 | Sales Summary | str | No |
| A3 | Total Revenue | str | No |
| B3 | =SUM(Analysis!F4:F8) | str | Yes |
| A4 | Average Order Value | str | No |
| B4 | =B3/COUNT(Analysis!A4:A8) | str | Yes |

#### Formula Analysis

**Most Common Functions:**

- SUM: 1 occurrences
- COUNT: 1 occurrences

**Sample Formulas:**

- **B3**: `=SUM(Analysis!F4:F8)`
- **B4**: `=B3/COUNT(Analysis!A4:A8)`

## 🔗 Relationships

## 📈 Data Type Analysis

### Data Types Distribution

- **str**: 41 cells (62.1%)
- **int**: 22 cells (33.3%)
- **float**: 3 cells (4.5%)

### Formula Functions Distribution

- **VLOOKUP**: 10 uses (58.8%)
- **SUM**: 1 uses (5.9%)
- **COUNT**: 1 uses (5.9%)

## 💡 Analysis & Recommendations

**Complexity Level**: Moderate
- This file has moderate complexity with some advanced features

### Key Observations

- Contains 17 formulas indicating active calculations
- Uses 1 formal Excel tables for structured data
