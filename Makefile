.PHONY: help install install-dev test test-cov lint format type-check quality clean build dist publish docs

# Default target
help:
	@echo "Excel Analyzer - Development Commands"
	@echo "====================================="
	@echo ""
	@echo "Installation:"
	@echo "  install      Install package in development mode"
	@echo "  install-dev  Install package with development dependencies"
	@echo ""
	@echo "Testing:"
	@echo "  test         Run tests"
	@echo "  test-cov     Run tests with coverage report"
	@echo ""
	@echo "Code Quality:"
	@echo "  lint         Run linting checks"
	@echo "  format       Format code with Black"
	@echo "  type-check   Run type checking with MyPy"
	@echo "  quality      Run all quality checks"
	@echo ""
	@echo "Building:"
	@echo "  build        Build package"
	@echo "  dist         Create distribution files"
	@echo "  clean        Clean build artifacts"
	@echo ""
	@echo "Documentation:"
	@echo "  docs         Build documentation"
	@echo ""
	@echo "Examples:"
	@echo "  demo         Run example analysis"

# Installation
install:
	pip install -e .

install-dev:
	pip install -e .[dev]
	pre-commit install

# Testing
test:
	python -m pytest tests/ -v

test-parser:
	python -m pytest tests/test_parser.py -v

test-extractor:
	python -m pytest tests/test_extractor.py -v

test-cov:
	python -m pytest tests/ --cov=excel_analyzer --cov-report=term-missing --cov-report=html

# Code Quality
lint:
	flake8 src/excel_analyzer/ examples/

format:
	black src/excel_analyzer/ examples/

type-check:
	mypy src/excel_analyzer/

quality: format lint type-check test

# Building
build:
	python -m build

dist: clean build

clean:
	rm -rf build/
	rm -rf dist/
	rm -rf *.egg-info/
	rm -rf htmlcov/
	rm -rf .coverage
	rm -rf .pytest_cache/
	rm -rf .mypy_cache/

# Documentation
docs:
	cd docs && make html

# Examples
demo:
	excel-analyzer excel_files/mycoolsample.xlsx --json --markdown --dataframes --save-dfs

demo-all:
	python examples/demo_parser.py

# Development helpers
setup: install-dev
	@echo "Development environment setup complete!"

check-all: quality test-cov
	@echo "All checks passed!"

# Git helpers
commit:
	git add .
	git commit -m "$${MESSAGE}"

push:
	git push origin main

# Package management
update-deps:
	pip install --upgrade pip
	pip install --upgrade -r requirements.txt

# Analysis examples
analyze-simple:
	excel-analyzer excel_files/simple_model.xlsx --json --markdown --summary

analyze-complex:
	excel-analyzer excel_files/complex_model.xlsx --json --markdown --summary

analyze-enterprise:
	excel-analyzer excel_files/enterprise_model.xlsx --json --markdown --summary

# Quick development cycle
dev-cycle: format lint test
	@echo "Development cycle complete!"

# CLI examples
cli-help:
	excel-analyzer --help

cli-batch:
	excel-analyzer "excel_files/*.xlsx" --output-dir ./batch_results --json --summary --batch

cli-dataframes:
	excel-analyzer excel_files/mycoolsample.xlsx --dataframes --save-dfs --dfs-format excel --output-dir ./dataframe_results

cli-verbose:
	excel-analyzer excel_files/complex_model.xlsx --verbose --json --markdown --dataframes --save-dfs

# Excel Extractor CLI examples
extractor-help:
	excel-extractor --help

extractor-basic:
	excel-extractor excel_files/mycoolsample.xlsx --markdown --json

extractor-llm:
	excel-extractor excel_files/enterprise_model.xlsx --llm-optimized --output-dir ./llm_reports

extractor-batch:
	excel-extractor "excel_files/*.xlsx" --output-dir ./extractor_batch --json --summary --batch --timing

extractor-verbose:
	excel-extractor excel_files/complex_model.xlsx --verbose --markdown --json --timing

# Click CLI examples
click-help:
	excel-analyzer-click --help

click-parser-help:
	excel-analyzer-click parser --help

click-extractor-help:
	excel-analyzer-click extractor --help

click-analyze-help:
	excel-analyzer-click analyze --help

click-parser-basic:
	excel-analyzer-click parser excel_files/simple_model.xlsx --summary

click-parser-full:
	excel-analyzer-click parser excel_files/simple_model.xlsx --json --markdown --dataframes --save-dfs

click-extractor-basic:
	excel-analyzer-click extractor excel_files/simple_model.xlsx --summary

click-extractor-full:
	excel-analyzer-click extractor excel_files/simple_model.xlsx --json --markdown --llm-optimized --timing

click-analyze-comprehensive:
	excel-analyzer-click analyze excel_files/simple_model.xlsx --json --markdown --dataframes --save-dfs --llm-optimized --timing

click-batch:
	excel-analyzer-click parser "excel_files/*.xlsx" --output-dir ./click_batch --json --summary --verbose

# Error Sniffer examples
error-sniffer-help:
	excel-error-sniffer --help

error-sniffer-basic:
	excel-error-sniffer excel_files/simple_model.xlsx --summary

error-sniffer-full:
	excel-error-sniffer excel_files/simple_model.xlsx --json --markdown --verbose --timing

error-sniffer-high-severity:
	excel-error-sniffer excel_files/simple_model.xlsx --severity high --summary

error-sniffer-batch:
	excel-error-sniffer "excel_files/*.xlsx" --output-dir ./error_reports --json --summary --verbose

error-sniffer-comprehensive:
	excel-error-sniffer excel_files/complex_model.xlsx --json --markdown --verbose --timing --output-dir ./error_reports

# Click CLI Error Detection examples
click-error-sniff-help:
	excel-analyzer-click error-sniff --help

click-error-sniff-basic:
	excel-analyzer-click error-sniff excel_files/simple_model.xlsx --summary

click-error-sniff-full:
	excel-analyzer-click error-sniff excel_files/simple_model.xlsx --json --markdown --verbose --timing

click-detect-errors-help:
	excel-analyzer-click detect-errors --help

click-detect-errors-basic:
	excel-analyzer-click detect-errors excel_files/simple_model.xlsx --summary

click-detect-errors-full:
	excel-analyzer-click detect-errors excel_files/simple_model.xlsx --json --markdown --verbose --timing

click-detect-errors-low-threshold:
	excel-analyzer-click detect-errors excel_files/simple_model.xlsx --error-threshold 0.3 --summary --verbose

click-detect-errors-batch:
	excel-analyzer-click detect-errors "excel_files/*.xlsx" --output-dir ./click_error_reports --json --summary --verbose
