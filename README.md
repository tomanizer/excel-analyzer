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
# Analyze a single Excel file
python excel_parser.py

# The script will create sample files and demonstrate the analysis
```

## 🎯 What It Does

This tool performs comprehensive structural analysis of Excel workbooks:

### ✅ Detects and Analyzes
- **Formal Excel Tables** (ListObjects)
- **Informal Data Islands** (contiguous data blocks)
- **Named Ranges** and variables
- **Data Validation Rules** (input cells)
- **Charts** and their data sources
- **Pivot Tables** and relationships
- **External Workbook Links**
- **VBA Macros** presence
- **VLOOKUP/HLOOKUP** dependencies
- **Cross-sheet formulas** and relationships

### 📊 Sample Output
```
--- Comprehensive Analysis for: financial_model.xlsx ---

VBA Project Detected: False

Named Ranges:
  - Discount_Rate: B5
  - Growth_Rate: B6

External Dependencies:
  - C:\Models\SourceData.xlsx

--- Sheet-Level Analysis ---

Processing Sheet: Assumptions
  Data Validation Rules Found:
    - B5: "0.01,0.02,0.03,0.04,0.05"

--- Discovered Data Tables & Islands ---
  - AssumptionsTable (Formal Table) on sheet 'Assumptions' at range A1:B10
  - RevenueCalc (Informal Data Island) on sheet 'Calculations' at range A1:D12
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
├── requirements.txt         # Python dependencies
├── venv/                   # Virtual environment
├── sample_model.xlsx       # Test files
├── final_model.xlsm        # Comprehensive test file
├── PROJECT_SUMMARY.md      # Detailed project documentation
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
- Relationship mapping
- Advanced structure detection
- Comprehensive profiling

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

**Note**: This is a work in progress. The tool is currently in development and may have limitations with complex Excel models. For production use, please ensure thorough testing with your specific models. 