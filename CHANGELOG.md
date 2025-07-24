# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Professional project structure with organized directories
- Comprehensive documentation and guides
- Structured data output (JSON format)
- Markdown report generation
- Pandas DataFrame extraction
- Pivot table detection
- CLI with multiple output options

### Changed
- Reorganized project structure for better maintainability
- Renamed test files to demo files for clarity
- Updated all file references to use new structure

### Fixed
- CLI now properly analyzes specified files instead of dummy data
- Improved resource handling to reduce zipfile warnings
- Fixed DataFrame extraction for proper range parsing

## [0.1.0] - 2025-01-24

### Added
- Initial release of Excel Analyzer
- Core Excel parsing engine with openpyxl
- Table discovery (formal and informal)
- Chart detection
- Data validation detection
- Named range detection
- External link detection
- VBA macro detection
- Data island detection algorithm
- Basic CLI interface
- Comprehensive Excel structure analysis

### Features
- Multi-layered table discovery
- Relationship mapping
- Advanced structure detection
- Comprehensive profiling
- Modular architecture with single responsibility functions
- Proper logging and error handling
- Type hints and documentation

### Technical
- Python 3.12+ compatibility
- Openpyxl integration for Excel file parsing
- Networkx for dependency graph construction
- Pandas for data manipulation
- Comprehensive test suite with various Excel file types 