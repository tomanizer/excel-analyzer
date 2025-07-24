# Project Structure

This document provides an overview of the organized project structure for the Excel Analyzer tool.

## 📁 Directory Structure

```
cfo_models/
├── 📁 src/                    # Source code directory
│   └── 📁 excel_analyzer/     # Main package
│       ├── 📄 __init__.py     # Package initialization
│       ├── 📄 excel_parser.py # Core analysis engine
│       └── 📄 excel_extractor.py # Advanced extraction tool
├── 📄 requirements.txt         # Python dependencies
├── 📄 README.md               # Main project documentation
├── 📄 PROJECT_STRUCTURE.md    # This file
├── 📄 setup.py                # Package installation
├── 📄 pyproject.toml          # Modern Python packaging
├── 📄 LICENSE                 # MIT License
├── 📄 CHANGELOG.md            # Version history
├── 📄 CONTRIBUTING.md         # Contributor guidelines
├── 📄 MANIFEST.in             # Package distribution
├── 📄 Makefile                # Development tasks
├── 📁 venv/                   # Virtual environment
├── 📁 excel_files/            # Excel files for analysis
│   ├── 📄 simple_model.xlsx
│   ├── 📄 intermediate_model.xlsx
│   ├── 📄 advanced_model.xlsx
│   ├── 📄 complex_model.xlsx
│   ├── 📄 enterprise_model.xlsx
│   ├── 📄 mycoolsample.xlsx
│   ├── 📄 Book 3.xlsx
│   ├── 📄 sample_model.xlsx
│   ├── 📄 final_model.xlsm
│   └── 📄 external_source.xlsx
├── 📁 examples/               # Example scripts and demos
│   ├── 📄 example_usage.py           # Programmatic usage example
│   ├── 📄 demo_parser.py             # Parser testing demo
│   ├── 📄 demo_file_generator.py     # Test file generator
│   ├── 📄 demo_llm_analysis.py       # LLM analysis demo
│   └── 📄 demo_llm_usage.py          # LLM usage demo
├── 📁 docs/                   # Documentation files
│   ├── 📄 PROJECT_SUMMARY.md         # Detailed project overview
│   ├── 📄 LLM_TESTING_GUIDE.md       # LLM testing guide
│   ├── 📄 intermediate_model_analysis.md
│   ├── 📄 intermediate_model_data_demo_prompts.md
│   ├── 📄 intermediate_model_data_llm_prompts.md
│   └── 📄 README.md                  # Excel files documentation
├── 📁 reports/                # Generated analysis reports
    ├── 📄 mycoolsample.json
    ├── 📄 mycoolsample.analysis.json
    └── 📄 intermediate_model_data.json
└── 📁 tests/                  # Test suite
    ├── 📄 __init__.py
    └── 📄 test_parser.py
```

## 🎯 Purpose of Each Directory

### 📁 `excel_files/`
- **Purpose**: Contains all Excel files used for analysis and testing
- **Content**: Only `.xlsx` and `.xlsm` files
- **Usage**: Source files for the analyzer

### 📁 `examples/`
- **Purpose**: Contains demonstration and example scripts
- **Content**: Python files showing how to use the tools
- **Naming**: All files prefixed with `demo_` or `example_`

### 📁 `docs/`
- **Purpose**: Project documentation and guides
- **Content**: Markdown files with detailed explanations
- **Usage**: Reference material for users and developers

### 📁 `reports/`
- **Purpose**: Generated analysis outputs
- **Content**: JSON files with structured analysis data
- **Usage**: Results from running the analyzer

## 🔧 Core Files

### 📁 `src/excel_analyzer/`
- **Main package directory**
- Contains all core functionality
- Proper Python package structure

### 📄 `src/excel_analyzer/excel_parser.py`
- Main analysis engine
- CLI interface
- Structured data output
- Markdown report generation
- Pandas DataFrame extraction

### 📄 `src/excel_analyzer/excel_extractor.py`
- Advanced extraction tool
- Comprehensive data extraction
- LLM-ready output formats

### 📄 `src/excel_analyzer/__init__.py`
- Package initialization
- Public API exports
- Version information

### 📄 `requirements.txt`
- Python dependencies
- Version specifications
- Easy installation

## 🚀 Usage Examples

### Command Line
```bash
# Basic analysis
excel-analyzer excel_files/mycoolsample.xlsx

# Generate reports
excel-analyzer excel_files/mycoolsample.xlsx --json --markdown --dataframes

# Using Python module
python -m excel_analyzer.excel_parser excel_files/mycoolsample.xlsx
```

### Programmatic
```python
from pathlib import Path
from excel_analyzer.excel_parser import analyze_workbook_final

# Analyze file
file_path = Path("excel_files/mycoolsample.xlsx")
analysis_data = analyze_workbook_final(file_path, return_data=True)
```

## 📋 File Naming Conventions

### Excel Files
- Descriptive names: `simple_model.xlsx`, `complex_model.xlsx`
- User files: `mycoolsample.xlsx`, `Book 3.xlsx`
- Generated files: `final_model.xlsm`, `external_source.xlsx`

### Python Files
- Core tools: `src/excel_analyzer/excel_parser.py`, `src/excel_analyzer/excel_extractor.py`
- Examples: `example_usage.py`
- Demos: `demo_*.py`
- Tests: `tests/test_*.py`

### Documentation
- Main docs: `README.md`, `PROJECT_SUMMARY.md`
- Guides: `LLM_TESTING_GUIDE.md`
- Structure: `PROJECT_STRUCTURE.md`

## 🔄 Workflow

1. **Input**: Place Excel files in `excel_files/`
2. **Analysis**: Run `excel_parser.py` on target files
3. **Output**: Generated reports go to `reports/`
4. **Documentation**: Reference `docs/` for guidance
5. **Examples**: Use `examples/` for learning

## 🎯 Benefits of This Structure

- **Clean Separation**: Each directory has a clear purpose
- **Easy Navigation**: Logical organization makes files easy to find
- **Scalable**: Structure supports growth and new features
- **Professional**: Follows industry best practices
- **Maintainable**: Clear organization reduces confusion 