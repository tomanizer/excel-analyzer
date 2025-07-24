"""
Excel Analyzer - A powerful tool for converting complex Excel financial models into standardized Python code.

This package provides comprehensive Excel file analysis capabilities including:
- Table discovery (formal and informal)
- Pivot table detection
- Chart analysis
- Data validation detection
- Named range analysis
- External link detection
- VBA macro detection
- Structured data output
- Markdown report generation
- Pandas DataFrame extraction
"""

__version__ = "0.1.0"
__author__ = "Thomas"
__email__ = "thomas@example.com"

from .excel_parser import analyze_workbook_final, generate_markdown_report, extract_data_to_dataframes
from .cli import main as cli_main

__all__ = [
    "analyze_workbook_final",
    "generate_markdown_report", 
    "extract_data_to_dataframes",
    "cli_main"
] 