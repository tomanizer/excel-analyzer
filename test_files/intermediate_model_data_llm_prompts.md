# LLM Analysis Prompts for intermediate_model.xlsx

This file contains various prompts you can use to analyze the extracted Excel data with an LLM.

## 1. Basic Summary

**Prompt:**

```
Analyze this Excel workbook data and provide a comprehensive summary:

**File Information:**
- Filename: intermediate_model.xlsx
- Size: 6.92 KB
- Sheets: 3
- Sheet Names: Data, Analysis, Summary

**Key Statistics:**
- Total cells with data: 66
- Total formulas: 17
- Total tables: 1
- Total charts: 0
- Named ranges: 0
- Cross-sheet references: 0

**Complexity Score:** 135

Please provide:
1. A brief overview of what this Excel file appears to be for
2. The main types of data and calculations it contains
3. Key insights about its structure and complexity
4. Recommendations for understanding or working with this file
```

---

## 2. Structure Analysis

**Prompt:**

```
Analyze the structure of this Excel workbook and explain how it's organized:

**Sheet Structure:**
- **Data**: 24 cells, 0 formulas, 1 tables, 0 charts
- **Analysis**: 37 cells, 15 formulas, 0 tables, 0 charts
- **Summary**: 5 cells, 2 formulas, 0 tables, 0 charts

**Data Types Found:**
- str: 41 cells (62.1%)
- int: 22 cells (33.3%)
- float: 3 cells (4.5%)

**Formula Analysis:**
- VLOOKUP: 10 uses (58.8%)
- SUM: 1 uses (5.9%)
- COUNT: 1 uses (5.9%)

Please provide:
1. How the workbook is structured across different sheets
2. What each sheet appears to be responsible for
3. How data flows between sheets (if any)
4. The overall architecture and design patterns used
```

---

## 3. Business Logic Analysis

**Prompt:**

```
Analyze the business logic and calculations in this Excel workbook:

**Key Calculations:**

**Analysis:**
- D4: =VLOOKUP(B4,Data!A1:D6,2,FALSE)
- E4: =VLOOKUP(B4,Data!A1:D6,3,FALSE)
- F4: =C4*E4

**Summary:**
- B3: =SUM(Analysis!F4:F8)
- B4: =B3/COUNT(Analysis!A4:A8)

**Data Relationships:**
No cross-sheet references found

**Tables and Structured Data:**

**Data:**
- ProductTable: A1:D6

Please provide:
1. What business processes or models this Excel file represents
2. The main calculations and their business purpose
3. Key assumptions and variables used
4. Potential business insights that could be derived
```

---

## 4. Data Quality Assessment

**Prompt:**

```
Assess the data quality and structure of this Excel workbook:

**Data Distribution:**
- **Data**: 24 cells (36.4%)
- **Analysis**: 37 cells (56.1%)
- **Summary**: 5 cells (7.6%)

**Formula Complexity:**
Total formulas: 17
Most complex functions:
- VLOOKUP: 10 uses
- SUM: 1 uses
- COUNT: 1 uses

**Validation and Controls:**
No data validation rules found

Please provide:
1. Assessment of data quality and consistency
2. Potential data integrity issues
3. Recommendations for data validation
4. Suggestions for improving data structure
```

---

## 5. Migration Analysis

**Prompt:**

```
Analyze this Excel workbook for potential migration to other systems:

**Current Structure:**
- Complexity Score: 135
- Total Formulas: 17
- Cross-sheet References: 0
- External Dependencies: 0

**Technical Features:**
No advanced technical features found

Please provide:
1. Assessment of migration complexity
2. Recommended migration approach
3. Potential challenges and risks
4. Suggested target systems or platforms
5. Estimated effort and timeline
```

---

