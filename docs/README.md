# Test Files for Excel Parser

This directory contains a comprehensive set of Excel test files with varying complexity to thoroughly test the excel parser's capabilities.

## Test Files Overview

### 1. simple_model.xlsx (4.9 KB)
**Complexity**: Basic
**Description**: Basic financial model with simple calculations
**Features**:
- Single sheet with basic assumptions
- Simple formulas (revenue - costs = profit)
- Informal data islands
- No advanced features

**Expected Parser Results**:
- 0 formal tables
- 2 data islands
- No charts, named ranges, or data validation

### 2. intermediate_model.xlsx (7.1 KB)
**Complexity**: Intermediate
**Description**: Model with VLOOKUPs and multiple sheets
**Features**:
- Multiple sheets (Data, Analysis, Summary)
- Formal Excel table (ProductTable)
- VLOOKUP formulas for product lookups
- Cross-sheet references
- Basic calculations

**Expected Parser Results**:
- 1 formal table (ProductTable)
- 1 data island (Analysis sheet)
- No charts, named ranges, or data validation

### 3. advanced_model.xlsx (7.6 KB)
**Complexity**: Advanced
**Description**: Advanced model with external links and charts
**Features**:
- Multiple sheets (External_Data, Main_Model)
- Data validation rules
- Line chart with revenue projections
- External data references (simulated)
- Complex formulas with growth calculations

**Expected Parser Results**:
- 0 formal tables
- 4 data islands
- 1 chart (LineChart)
- Data validation rules present
- No named ranges

### 4. complex_model.xlsx (9.9 KB)
**Complexity**: Complex
**Description**: Complex model with named ranges and projections
**Features**:
- Multiple sheets (Inputs, Historical_Data, Calculations, Summary)
- Named ranges (Discount_Rate, Growth_Rate, Tax_Rate)
- Formal Excel table (HistoricalTable)
- Data validation with dropdown lists
- Bar chart with projections
- Complex financial calculations
- NPV calculations

**Expected Parser Results**:
- 1 formal table (HistoricalTable)
- 7 data islands
- 1 chart (BarChart)
- Named ranges present
- Data validation rules present

### 5. enterprise_model.xlsx (10.7 KB)
**Complexity**: Enterprise
**Description**: Enterprise model with multiple business units
**Features**:
- 5 sheets (Manufacturing, Sales, Consolidation, Scenarios, Dashboard)
- Multiple business unit data
- Cross-sheet consolidation formulas
- Scenario analysis
- Executive dashboard
- Multiple charts (LineChart, BarChart)
- Complex financial modeling

**Expected Parser Results**:
- 0 formal tables
- 11 data islands
- 2 charts (LineChart, BarChart)
- No named ranges or data validation

## Test Results Summary

Based on the comprehensive test run:

| File | Tables | Islands | Charts | Named Ranges | Data Validation |
|------|--------|---------|--------|--------------|-----------------|
| simple_model.xlsx | 0 | 2 | No | No | No |
| intermediate_model.xlsx | 1 | 1 | No | No | No |
| advanced_model.xlsx | 0 | 4 | Yes | No | Yes |
| complex_model.xlsx | 1 | 7 | Yes | Yes | Yes |
| enterprise_model.xlsx | 0 | 11 | Yes | No | No |

## Feature Coverage

The test files collectively test all major features of the excel parser:

✅ **Formal Tables**: 2 files contain Excel ListObjects
✅ **Data Islands**: All 5 files contain informal data blocks
✅ **Charts**: 3 files contain various chart types
✅ **Named Ranges**: 1 file contains named ranges
✅ **Data Validation**: 2 files contain validation rules
✅ **Multiple Sheets**: 4 files have multiple worksheets
✅ **Complex Formulas**: All files contain various formula types
✅ **Cross-sheet References**: 3 files have cross-sheet dependencies

## Usage

### Running Individual Tests
```bash
# Test a specific file
python -c "from excel_parser import analyze_workbook_final; from pathlib import Path; analyze_workbook_final(Path('test_files/simple_model.xlsx'))"
```

### Running All Tests
```bash
# Run comprehensive test suite
python test_parser.py
```

### Expected Parser Behavior

The parser should successfully:
1. **Detect all structural elements** in each file
2. **Identify formal tables** vs informal data islands
3. **Map relationships** between sheets and cells
4. **Recognize advanced features** like charts and named ranges
5. **Handle varying complexity** from simple to enterprise-level models

## File Sizes and Performance

| File | Size | Sheets | Expected Parse Time |
|------|------|--------|-------------------|
| simple_model.xlsx | 4.9 KB | 1 | < 1 second |
| intermediate_model.xlsx | 7.1 KB | 3 | < 1 second |
| advanced_model.xlsx | 7.6 KB | 2 | < 1 second |
| complex_model.xlsx | 9.9 KB | 4 | < 2 seconds |
| enterprise_model.xlsx | 10.7 KB | 5 | < 2 seconds |

## Notes

- All files are created using openpyxl and are compatible with Excel 2010+
- The files contain realistic financial modeling scenarios
- Chart detection may show verbose output due to openpyxl's chart object structure
- Some warnings about deprecated functions may appear (these don't affect functionality)
- The test files are designed to be representative of real-world financial models

## Maintenance

To regenerate all test files:
```bash
python test_file_generator.py
```

This will recreate all test files with fresh data and ensure consistency across test runs. 