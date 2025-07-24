#!/usr/bin/env python3
"""
Tests for the Excel Parser module.
"""

import pytest
from pathlib import Path

from excel_analyzer.excel_parser import analyze_workbook_final, generate_markdown_report, extract_data_to_dataframes


class TestExcelParser:
    """Test cases for Excel parser functionality."""
    
    def test_analyze_workbook_final_with_data(self):
        """Test that analyze_workbook_final returns structured data."""
        # Use a simple test file
        test_file = Path("excel_files/simple_model.xlsx")
        
        if test_file.exists():
            result = analyze_workbook_final(test_file, return_data=True)
            
            # Check that we get a dictionary with expected keys
            assert isinstance(result, dict)
            assert "metadata" in result
            assert "global_features" in result
            assert "sheets" in result
            assert "summary" in result
            assert "all_tables" in result
            
            # Check metadata
            assert result["metadata"]["filename"] == "simple_model.xlsx"
            assert result["metadata"]["file_size_kb"] > 0
            
            # Check summary
            assert result["summary"]["total_sheets"] >= 0
            assert result["summary"]["total_data_islands"] >= 0
    
    def test_analyze_workbook_final_nonexistent_file(self):
        """Test that analyze_workbook_final handles nonexistent files."""
        test_file = Path("nonexistent_file.xlsx")
        result = analyze_workbook_final(test_file, return_data=True)
        assert result == {}
    
    def test_generate_markdown_report(self):
        """Test markdown report generation."""
        # Create minimal test data
        test_data = {
            "metadata": {
                "filename": "test.xlsx",
                "file_size_kb": 10.0,
                "analysis_timestamp": "2025-01-24T00:00:00"
            },
            "global_features": {
                "vba_detected": False,
                "external_links": [],
                "named_ranges": {}
            },
            "sheets": {},
            "summary": {
                "total_sheets": 1,
                "total_formal_tables": 0,
                "total_pivot_tables": 0,
                "total_charts": 0,
                "total_data_islands": 0,
                "total_data_validation_rules": 0
            },
            "all_tables": []
        }
        
        report = generate_markdown_report(test_data)
        
        # Check that report contains expected content
        assert "# Excel Analysis Report: test.xlsx" in report
        assert "## 📊 Executive Summary" in report
        assert "**Total Sheets:** 1" in report
    
    def test_generate_markdown_report_empty_data(self):
        """Test markdown report generation with empty data."""
        report = generate_markdown_report({})
        assert "No analysis data provided." in report
    
    def test_extract_data_to_dataframes(self):
        """Test DataFrame extraction."""
        # Use a simple test file
        test_file = Path("excel_files/simple_model.xlsx")
        
        if test_file.exists():
            # First get analysis data
            analysis_data = analyze_workbook_final(test_file, return_data=True)
            
            # Then extract DataFrames
            dataframes = extract_data_to_dataframes(analysis_data, test_file)
            
            # Check that we get a dictionary
            assert isinstance(dataframes, dict)
            
            # Check that all DataFrames are either None or pandas DataFrames
            import pandas as pd
            for name, df in dataframes.items():
                assert df is None or isinstance(df, pd.DataFrame)


class TestExcelParserIntegration:
    """Integration tests for Excel parser."""
    
    def test_full_analysis_workflow(self):
        """Test the complete analysis workflow."""
        test_file = Path("excel_files/mycoolsample.xlsx")
        
        if test_file.exists():
            # 1. Analyze file
            analysis_data = analyze_workbook_final(test_file, return_data=True)
            assert analysis_data is not None
            
            # 2. Generate markdown report
            report = generate_markdown_report(analysis_data)
            assert len(report) > 0
            
            # 3. Extract DataFrames
            dataframes = extract_data_to_dataframes(analysis_data, test_file)
            assert len(dataframes) > 0
            
            # 4. Check that we have the expected structure
            assert "metadata" in analysis_data
            assert "sheets" in analysis_data
            assert "summary" in analysis_data


if __name__ == "__main__":
    pytest.main([__file__]) 