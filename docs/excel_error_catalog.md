# Excel Error Catalog: Complex Errors That Break Financial Models

This document catalogs complex Excel errors that are difficult to detect but can cause serious issues with financial model results. These errors often go unnoticed during development but can lead to catastrophic failures in production.

## Table of Contents

1. [Complex Calculation Errors](#complex-calculation-errors)
2. [Range Incomplete Formula Errors](#range-incomplete-formula-errors)
3. [Formula Reference Anchoring Errors](#formula-reference-anchoring-errors)
4. [Data Integrity Errors](#data-integrity-errors)
5. [Performance and Structural Errors](#performance-and-structural-errors)

---

## Complex Calculation Errors

### 1. Hidden Rows/Columns in Data Ranges

**Issue**: Data ranges include hidden rows/columns that contain incorrect or outdated values.

**Impact**: Calculations use wrong data without users realizing, leading to incorrect results.

**Example**:
```
SUM(A1:A100) includes hidden rows with old data
```

**Detection Challenge**: Need to check if ranges include hidden cells and verify data integrity.

**Severity**: 🔴 **HIGH**

---

### 2. Circular References in Named Ranges

**Issue**: Named ranges that reference each other in a circular manner.

**Impact**: Can cause infinite calculation loops or incorrect results.

**Example**:
```
Named Range "Revenue" = SUM(Expenses)
Named Range "Expenses" = Revenue * 0.8
```

**Detection Challenge**: Complex dependency mapping across multiple named ranges.

**Severity**: 🔴 **HIGH**

---

### 3. Inconsistent Date Formats in Date Calculations

**Issue**: Mixed date formats (text vs. actual dates) in date-based calculations.

**Impact**: Date arithmetic returns errors or wrong results.

**Example**:
```
"01/01/2024" (text) vs 01/01/2024 (date)
=DATEDIF("01/01/2024", TODAY(), "D") # Returns error
```

**Detection Challenge**: Need to identify cells that should be dates but are stored as text.

**Severity**: 🟡 **MEDIUM**

---

### 4. Array Formula Spill Errors

**Issue**: Array formulas that can't spill properly due to insufficient space or conflicts.

**Impact**: Formulas return #SPILL! errors or partial results.

**Example**:
```
=UNIQUE(A1:A100) # Spills into occupied cells
```

**Detection Challenge**: Need to check spill ranges and potential conflicts.

**Severity**: 🟡 **MEDIUM**

---

### 5. Volatile Functions in Large Models

**Issue**: NOW(), TODAY(), RAND(), OFFSET(), INDIRECT() causing excessive recalculations.

**Impact**: Model becomes slow and may produce inconsistent results.

**Example**:
```
=INDIRECT("A" & ROW()) # Recalculates on every change
```

**Detection Challenge**: Identifying performance impact of volatile functions.

**Severity**: 🟡 **MEDIUM**

---

### 6. Cross-Sheet Reference Errors

**Issue**: References to cells in other sheets that have been moved/deleted.

**Impact**: #REF! errors or references to wrong cells.

**Example**:
```
=Sheet1!A1 # Sheet1 was renamed or A1 was deleted
```

**Detection Challenge**: Tracking dependencies across multiple sheets.

**Severity**: 🔴 **HIGH**

---

### 7. Data Type Inconsistencies in Lookup Tables

**Issue**: Mixed data types in lookup tables (numbers stored as text, etc.).

**Impact**: VLOOKUP/HLOOKUP/XLOOKUP return wrong results or errors.

**Example**:
```
VLOOKUP table: "123" (text) vs 123 (number)
```

**Detection Challenge**: Identifying data type mismatches in lookup ranges.

**Severity**: 🔴 **HIGH**

---

### 8. Conditional Formatting Overlap Conflicts

**Issue**: Multiple conditional formatting rules that conflict or overlap.

**Impact**: Visual confusion and potential calculation issues.

**Example**:
```
Rule 1: A1:A100 > 0 (green)
Rule 2: A1:A100 < 0 (red)
Rule 3: A1:A100 = 0 (yellow) # Conflicts with Rule 1
```

**Detection Challenge**: Analyzing rule precedence and conflicts.

**Severity**: 🟢 **LOW**

---

### 9. External Data Connection Failures

**Issue**: Links to external databases or files that are broken or outdated.

**Impact**: Models use stale or missing data.

**Example**:
```
=SQL.REQUEST("connection_string", "SELECT * FROM table")
# Connection string is invalid or table doesn't exist
```

**Detection Challenge**: Testing external connections and data freshness.

**Severity**: 🔴 **HIGH**

---

### 10. Precision Errors in Financial Calculations

**Issue**: Floating-point precision errors in financial calculations.

**Impact**: Rounding errors accumulate and cause significant discrepancies.

**Example**:
```
=0.1 + 0.2 # Returns 0.30000000000000004 instead of 0.3
```

**Detection Challenge**: Identifying calculations that need special precision handling.

**Severity**: 🟡 **MEDIUM**

---

## Range Incomplete Formula Errors

### 1. Incomplete Formula Drag (Formula Cutoff)

**Issue**: Developer drags formula down but stops short, leaving cells without formulas.

**Impact**: Missing calculations for remaining rows, leading to incomplete totals.

**Example**:
```
Revenue calculation formula in A1:A100, but only dragged to A50
Rows A51:A100 have no formulas
```

**Detection**: Find formula patterns that suddenly stop mid-range.

**Severity**: 🔴 **HIGH**

---

### 2. False Range End Detection (Empty Cell Trap)

**Issue**: Empty cell in middle of data range makes developer think it's the end.

**Impact**: Missing calculations for data beyond the empty cell.

**Example**:
```
Data in A1:A100, but A50 is empty
Formula only covers A1:A49, missing A51:A100
```

**Detection**: Identify gaps in data ranges and check if formulas extend beyond them.

**Severity**: 🔴 **HIGH**

---

### 3. Partial Formula Propagation (Missing Edge Cases)

**Issue**: Developer copies formula to most cells but misses one or more.

**Impact**: Inconsistent calculations, especially dangerous in totals/sums.

**Example**:
```
Formula copied to A1:A99 but A100 is missing the formula
```

**Detection**: Find cells that should have formulas based on surrounding patterns.

**Severity**: 🔴 **HIGH**

---

### 4. Inconsistent Formula Application (Mixed Range)

**Issue**: Some cells in a range have formulas, others don't, creating mixed calculation methods.

**Impact**: Inconsistent calculation logic, hard to audit.

**Example**:
```
A1:A50 has formulas, A51:A100 has hardcoded values
```

**Detection**: Identify ranges with mixed formula/hardcoded patterns.

**Severity**: 🟡 **MEDIUM**

---

### 5. Formula Boundary Mismatch (Range Misalignment)

**Issue**: Formula range doesn't match the actual data range.

**Impact**: Incomplete aggregations, wrong totals.

**Example**:
```
SUM formula covers A1:A50 but data extends to A100
```

**Detection**: Compare formula ranges with actual data boundaries.

**Severity**: 🔴 **HIGH**

---

### 6. Copy-Paste Formula Gaps

**Issue**: When copying formulas, some cells get skipped due to selection errors.

**Impact**: Missing calculations in unexpected places.

**Example**:
```
Copying A1:A10 but accidentally missing A5
```

**Detection**: Find gaps in formula sequences.

**Severity**: 🟡 **MEDIUM**

---

### 7. Formula Range vs Data Range Discrepancy

**Issue**: Formulas reference ranges that don't match the actual data extent.

**Impact**: Lookups fail or return wrong results for data beyond the range.

**Example**:
```
VLOOKUP table range is A1:B50 but data goes to B100
```

**Detection**: Compare referenced ranges with actual data boundaries.

**Severity**: 🔴 **HIGH**

---

## Formula Reference Anchoring Errors

### 1. Missing Dollar Sign Anchors (Relative Reference Errors)

**Issue**: Formula should have absolute references but uses relative references.

**Impact**: References shift incorrectly when copied, causing wrong calculations.

**Example**:
```
=A1*B1 should be =$A$1*B1 when copying down
```

**Detection**: Identify formulas that should be anchored based on usage patterns.

**Severity**: 🔴 **HIGH**

---

### 2. Wrong Row/Column Anchoring (Partial Lock Errors)

**Issue**: Dollar sign on wrong part of reference.

**Impact**: Formulas copy incorrectly, creating calculation errors.

**Examples**:
```
=$A1 when it should be =A$1 (locked column, should lock row)
=A$1 when it should be =$A1 (locked row, should lock column)
=$A$1 when it should be =A1 (over-locked, should be relative)
```

**Detection**: Analyze anchoring patterns and identify inconsistencies.

**Severity**: 🔴 **HIGH**

---

### 3. Over-Anchored References (Unnecessary Absolute)

**Issue**: Dollar signs on references that should be relative.

**Impact**: Formula doesn't adapt when copied, missing the intended pattern.

**Example**:
```
=$A$1+$B$1 in every cell when it should be =A1+B1
```

**Detection**: Identify over-anchored references that should be relative.

**Severity**: 🟡 **MEDIUM**

---

### 4. Inconsistent Anchoring in Ranges

**Issue**: Mixed anchoring within the same formula.

**Impact**: Range expands incorrectly when copied.

**Example**:
```
=SUM($A$1:A10) - first reference anchored, second not
```

**Detection**: Check for mixed anchoring patterns in formulas.

**Severity**: 🟡 **MEDIUM**

---

### 5. Anchoring Errors in Lookup Functions

**Issue**: Wrong anchoring in VLOOKUP, HLOOKUP, INDEX/MATCH.

**Impact**: Lookups fail or return wrong results when copied.

**Examples**:
```
=VLOOKUP(A1,$B$1:$C$100,2) - lookup value should be $A1 when copying across
=INDEX($A$1:$A$100,MATCH(B1,$B$1:$B$100,0)) - mixed anchoring issues
```

**Detection**: Analyze anchoring patterns in lookup functions.

**Severity**: 🔴 **HIGH**

---

### 6. Anchoring in Array Formulas

**Issue**: Incorrect anchoring in array formulas.

**Impact**: Array formulas don't work correctly when applied to different ranges.

**Example**:
```
=SUM(IF($A$1:$A$100>0,$B$1:$B$100,0)) when ranges should be relative
```

**Detection**: Check anchoring consistency in array formulas.

**Severity**: 🟡 **MEDIUM**

---

### 7. Cross-Sheet Anchoring Errors

**Issue**: Wrong anchoring when referencing other sheets.

**Impact**: Cross-sheet references don't adapt properly.

**Example**:
```
=Sheet1!$A$1 when it should be =Sheet1!A1 for copying
```

**Detection**: Analyze cross-sheet reference anchoring patterns.

**Severity**: 🟡 **MEDIUM**

---

### 8. Named Range Anchoring Issues

**Issue**: Named ranges used without proper anchoring.

**Impact**: Named ranges don't work as expected when copied.

**Example**:
```
=MyRange*B1 when it should be =MyRange*$B$1
```

**Detection**: Check anchoring patterns with named ranges.

**Severity**: 🟡 **MEDIUM**

---

### 9. Conditional Formatting Anchoring Errors

**Issue**: Wrong anchoring in conditional formatting rules.

**Impact**: Conditional formatting doesn't work correctly.

**Example**:
```
=$A$1>0 applied to entire column when it should be =A1>0
```

**Detection**: Analyze anchoring in conditional formatting rules.

**Severity**: 🟢 **LOW**

---

### 10. Data Validation Anchoring Errors

**Issue**: Incorrect anchoring in data validation rules.

**Impact**: Data validation doesn't apply correctly to all cells.

**Example**:
```
List validation using =$A$1:$A$10 when it should be relative
```

**Detection**: Check anchoring patterns in data validation rules.

**Severity**: 🟢 **LOW**

---

## Data Integrity Errors

### 1. Mixed Data Types in Calculations

**Issue**: Combining text and numbers in calculations.

**Impact**: Calculations fail or return unexpected results.

**Example**:
```
="123" + 456 # Returns "123456" instead of 579
```

**Detection**: Identify cells with mixed data types in calculation ranges.

**Severity**: 🔴 **HIGH**

---

### 2. Hidden Data in Ranges

**Issue**: Hidden rows/columns contain data that affects calculations.

**Impact**: Users unaware of hidden data affecting results.

**Example**:
```
SUM(A1:A100) includes hidden row with old data
```

**Detection**: Check for hidden data in calculation ranges.

**Severity**: 🔴 **HIGH**

---

### 3. Stale Data in Named Ranges

**Issue**: Named ranges reference outdated or incorrect data.

**Impact**: Calculations use wrong data sources.

**Example**:
```
Named Range "CurrentYear" still points to 2023 data
```

**Detection**: Verify named range references are current.

**Severity**: 🔴 **HIGH**

---

## Performance and Structural Errors

### 1. Excessive Volatile Functions

**Issue**: Too many volatile functions causing performance issues.

**Impact**: Model becomes slow and unresponsive.

**Example**:
```
Multiple NOW(), RAND(), OFFSET() functions throughout model
```

**Detection**: Count and analyze volatile function usage.

**Severity**: 🟡 **MEDIUM**

---

### 2. Large Unused Ranges

**Issue**: Formulas reference unnecessarily large ranges.

**Impact**: Performance degradation and potential errors.

**Example**:
```
=SUM(A:A) instead of =SUM(A1:A1000)
```

**Detection**: Identify overly broad range references.

**Severity**: 🟡 **MEDIUM**

---

### 3. Broken External Links

**Issue**: External file references that no longer exist.

**Impact**: #REF! errors and missing data.

**Example**:
```
=[OldFile.xlsx]Sheet1!A1 # File no longer exists
```

**Detection**: Check external link validity.

**Severity**: 🔴 **HIGH**

---

## Error Detection Priorities

### 🔴 **Critical (High Priority)**
- Hidden data in ranges
- Circular references in named ranges
- Cross-sheet reference errors
- Data type inconsistencies
- External data connection failures
- Incomplete formula propagation
- Formula boundary mismatches
- Missing dollar sign anchors
- Wrong row/column anchoring

### 🟡 **Important (Medium Priority)**
- Inconsistent date formats
- Array formula spill errors
- Volatile functions in large models
- Conditional formatting conflicts
- Precision errors in financial calculations
- Inconsistent formula application
- Copy-paste formula gaps
- Over-anchored references
- Inconsistent anchoring in ranges
- Anchoring errors in lookup functions
- Mixed data types in calculations
- Excessive volatile functions
- Large unused ranges

### 🟢 **Minor (Low Priority)**
- Conditional formatting anchoring errors
- Data validation anchoring errors

---

## Detection Strategies

### 1. **Pattern Analysis**
- Identify formula patterns that break unexpectedly
- Check for consistency in formula application
- Analyze anchoring patterns across ranges

### 2. **Range Validation**
- Compare formula ranges with actual data boundaries
- Check for hidden data in calculation ranges
- Verify named range references

### 3. **Dependency Mapping**
- Track cross-sheet references
- Map named range dependencies
- Identify circular reference chains

### 4. **Data Type Checking**
- Verify consistent data types in lookup tables
- Check for mixed text/number formats
- Validate date format consistency

### 5. **Performance Analysis**
- Count volatile function usage
- Identify overly broad range references
- Check for unnecessary calculations

---

## Prevention Best Practices

### 1. **Formula Design**
- Use consistent anchoring patterns
- Avoid volatile functions when possible
- Keep ranges as small as necessary

### 2. **Data Management**
- Regularly audit named ranges
- Clean up unused external links
- Validate data types consistently

### 3. **Testing Procedures**
- Test formulas across different scenarios
- Verify calculations with known values
- Check for hidden data regularly

### 4. **Documentation**
- Document complex formulas
- Maintain change logs
- Keep external link inventories

---

*This catalog serves as a reference for identifying and preventing complex Excel errors that can compromise financial model integrity. Regular audits using these categories can help maintain model reliability and accuracy.* 