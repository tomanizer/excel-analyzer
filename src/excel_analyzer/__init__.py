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
from .excel_extractor import ExcelExtractor, extract_excel_to_markdown
from .extractor_cli import main as extractor_cli_main
from .excel_error_sniffer import ExcelErrorSniffer, sniff_excel_errors
from .error_sniffer_cli import main as error_sniffer_cli_main
from .probabilistic_error_detector import (
    ProbabilisticErrorSniffer, 
    ErrorDetector, 
    ErrorDetectionResult,
    detect_excel_errors_probabilistic
)

__all__ = [
    "analyze_workbook_final",
    "generate_markdown_report", 
    "extract_data_to_dataframes",
    "cli_main",
    "ExcelExtractor",
    "extract_excel_to_markdown",
    "extractor_cli_main",
    "ExcelErrorSniffer",
    "sniff_excel_errors",
    "error_sniffer_cli_main",
    "ProbabilisticErrorSniffer",
    "ErrorDetector",
    "ErrorDetectionResult",
    "detect_excel_errors_probabilistic"
] 