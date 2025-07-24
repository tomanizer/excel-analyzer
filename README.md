# CFO Models - Excel to Python Conversion Tool

A powerful tool for converting complex Excel financial models into standardized Python code. This project aims to automate the conversion of 300 financial and accounting Excel models while preserving their logic and relationships.

## 🚀 Quick Start

### Prerequisites
- Python 3.12+
- Virtual environment (recommended)

### Installation
```bash
# Clone the repository
git clone <repository-url>
cd cfo_models

# Create and activate virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Basic Usage
```bash
# Analyze a single Excel file (console output)
python excel_parser.py excel_files/mycoolsample.xlsx

# Generate structured JSON data
python excel_parser.py excel_files/mycoolsample.xlsx --json

# Generate markdown report
python excel_parser.py excel_files/mycoolsample.xlsx --markdown

# Extract data to pandas DataFrames
python excel_parser.py excel_files/mycoolsample.xlsx --dataframes

# All outputs at once
python excel_parser.py excel_files/mycoolsample.xlsx --json --markdown --dataframes
```

### Programmatic Usage
```python
from pathlib import Path
from excel_parser import analyze_workbook_final, generate_markdown_report, extract_data_to_dataframes

# Get structured analysis data
file_path = Path("excel_files/mycoolsample.xlsx")
analysis_data = analyze_workbook_final(file_path, return_data=True)

# Extract to pandas DataFrames
dataframes = extract_data_to_dataframes(analysis_data, file_path)

# Generate markdown report
report = generate_markdown_report(analysis_data)
```

## 🎯 What It Does

This tool performs comprehensive structural analysis of Excel workbooks:

### ✅ Detects and Analyzes
- **Formal Excel Tables** (ListObjects)
- **Informal Data Islands** (contiguous data blocks)
- **Pivot Tables** and their locations
- **Named Ranges** and variables
- **Data Validation Rules** (input cells)
- **Charts** and their data sources
- **External Workbook Links**
- **VBA Macros** presence
- **VLOOKUP/HLOOKUP** dependencies
- **Cross-sheet formulas** and relationships

### 📊 Output Formats

#### 1. Console Analysis
```
--- Comprehensive Analysis for: mycoolsample.xlsx ---

VBA Project Detected: False

Named Ranges:
  - mycoolrange: [('Sheet2', '$C$2:$C$7')]

--- Sheet-Level Analysis ---

Processing Sheet: Sheet1

Processing Sheet: Sheet2
  Pivot Tables Found:
    - 'PivotTable1' at range H7:N9

--- Discovered Data Tables & Islands ---
  - Table1 (Formal Table) on sheet 'Sheet3' at range A1:E3
  - Island_C2:D7 (Informal Data Island) on sheet 'Sheet2' at range C2:D7
```

#### 2. Structured JSON Data
```json
{
  "metadata": {
    "filename": "mycoolsample.xlsx",
    "file_size_kb": 15.9,
    "analysis_timestamp": "2025-07-24T22:54:24.730648"
  },
  "global_features": {
    "vba_detected": false,
    "external_links": [],
    "named_ranges": {
      "mycoolrange": [["Sheet2", "$C$2:$C$7"]]
    }
  },
  "sheets": {
    "Sheet2": {
      "pivot_tables": [
        {
          "name": "PivotTable1",
          "range": "H7:N9"
        }
      ]
    }
  },
  "summary": {
    "total_sheets": 3,
    "total_formal_tables": 1,
    "total_pivot_tables": 1,
    "total_data_islands": 6
  }
}
```

#### 3. Markdown Reports
```markdown
# Excel Analysis Report: mycoolsample.xlsx

**Analysis Date:** 2025-07-24T22:54:24.730648
**File Size:** 15.9 KB

## 📊 Executive Summary
- **Total Sheets:** 3
- **Formal Tables:** 1
- **Pivot Tables:** 1
- **Data Islands:** 6

## 📋 Sheet-by-Sheet Analysis
### Sheet: Sheet2
**Pivot Tables:**
- `PivotTable1` at range `H7:N9`
```

#### 4. Pandas DataFrames
```python
# Extract all tables and data islands as DataFrames
dataframes = extract_data_to_dataframes(analysis_data, file_path)

# Access specific DataFrames
table1_df = dataframes['Table1']  # Formal table
island_df = dataframes['Island_C2:D7']  # Data island
pivot_source_df = dataframes['Island_C2:D7']  # Pivot table source data
```

## 🏗️ Architecture

### Phase 1: Structural Analysis (Algorithmic)
- Multi-layered table discovery
- Relationship mapping
- Advanced structure detection
- Comprehensive profiling

### Phase 2: Semantic Analysis (AI-Enhanced) - Coming Soon
- Table classification
- Data flow synthesis
- Code translation

## 📁 Project Structure
```
cfo_models/
├── excel_parser.py          # Core analysis engine
├── excel_extractor.py       # Advanced extraction tool
├── requirements.txt         # Python dependencies
├── venv/                   # Virtual environment
├── excel_files/            # Excel files for analysis
├── examples/               # Example scripts and demos
├── docs/                   # Documentation files
├── reports/                # Generated analysis reports
└── README.md              # This file
```

## 🔧 Development

### Code Standards
- **Modular Design**: Single responsibility functions
- **Logging**: No print statements, proper logging
- **Type Hints**: Full type annotation
- **Docstrings**: Restructured text format
- **Testing**: Unit and integration tests

### Adding Features
1. Follow the modular architecture
2. Add comprehensive logging
3. Include type hints
4. Write tests for new functionality
5. Update documentation

## 🎯 Current Status

### ✅ Completed
- Core Excel parsing engine
- Table discovery (formal + informal)
- Pivot table detection
- Relationship mapping
- Advanced structure detection
- Comprehensive profiling
- Structured data output (JSON)
- Markdown report generation
- Pandas DataFrame extraction

### 🔄 In Progress
- AI integration for semantic analysis
- Python code generation
- CLI interface development
- Comprehensive testing

### 📋 Planned
- Batch processing for 300+ models
- Web-based GUI interface
- Performance optimization
- Enterprise features

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Follow coding standards
4. Add tests for new features
5. Submit a pull request

## 📄 License

[Add your license information here]

## 📞 Support

For questions or support, please [create an issue](link-to-issues) or contact the development team.

--- 