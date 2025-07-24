#!/usr/bin/env python3
"""
Setup script for Excel Analyzer - CFO Models Tool

A powerful tool for converting complex Excel financial models into standardized Python code.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding="utf-8")

# Read requirements
requirements = []
with open("requirements.txt", "r", encoding="utf-8") as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith("#")]

setup(
    name="excel-analyzer",
    version="0.1.0",
    author="Thomas",
    author_email="thomas@example.com",
    description="A powerful tool for converting complex Excel financial models into standardized Python code",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/tomanizer/excel-analyzer",
    project_urls={
        "Bug Reports": "https://github.com/tomanizer/excel-analyzer/issues",
        "Source": "https://github.com/tomanizer/excel-analyzer",
        "Documentation": "https://github.com/tomanizer/excel-analyzer#readme",
    },
    packages=find_packages(),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Intended Audience :: Financial and Insurance Industry",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Topic :: Office/Business :: Financial",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Text Processing :: Markup",
    ],
    python_requires=">=3.12",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
            "mypy>=1.0.0",
            "pre-commit>=3.0.0",
        ],
        "docs": [
            "sphinx>=6.0.0",
            "sphinx-rtd-theme>=1.0.0",
            "myst-parser>=1.0.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "excel-analyzer=excel_parser:main",
        ],
    },
    include_package_data=True,
    zip_safe=False,
    keywords="excel, financial, models, analysis, pandas, openpyxl",
) 