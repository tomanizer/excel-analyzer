#!/usr/bin/env python3
"""
Test Script for Excel Parser

Runs the excel parser on all test files and provides a summary of results.
"""

from pathlib import Path
from excel_analyzer.excel_parser import analyze_workbook_final
import sys
import io
from contextlib import redirect_stdout

def test_all_files():
    """Test the excel parser on all test files."""
    test_dir = Path("excel_files")
    
    if not test_dir.exists():
        print("Excel files directory not found. Please run demo_file_generator.py first.")
        return
    
    test_files = list(test_dir.glob("*.xlsx"))
    
    if not test_files:
        print("No Excel files found in excel_files directory.")
        return
    
    print("Testing Excel Parser on All Test Files")
    print("=" * 60)
    print()
    
    results = []
    
    for test_file in sorted(test_files):
        print(f"Testing: {test_file.name}")
        print("-" * 40)
        
        # Capture the output
        output = io.StringIO()
        with redirect_stdout(output):
            try:
                analyze_workbook_final(test_file)
                success = True
                error = None
            except Exception as e:
                success = False
                error = str(e)
        
        # Parse the output to extract key information
        output_text = output.getvalue()
        
        # Extract key metrics
        vba_detected = "VBA Project Detected: True" in output_text
        charts_found = "Charts Found:" in output_text
        tables_found = "Formal Table" in output_text
        islands_found = "Informal Data Island" in output_text
        external_links = "External Dependencies:" in output_text
        named_ranges = "Named Ranges:" in output_text
        data_validation = "Data Validation Rules Found:" in output_text
        
        # Count tables and islands
        table_count = output_text.count("Formal Table")
        island_count = output_text.count("Informal Data Island")
        
        results.append({
            'file': test_file.name,
            'success': success,
            'error': error,
            'vba_detected': vba_detected,
            'charts_found': charts_found,
            'tables_found': tables_found,
            'islands_found': islands_found,
            'external_links': external_links,
            'named_ranges': named_ranges,
            'data_validation': data_validation,
            'table_count': table_count,
            'island_count': island_count,
            'output': output_text
        })
        
        if success:
            print(f"✓ Successfully analyzed {test_file.name}")
            print(f"  - Tables found: {table_count}")
            print(f"  - Data islands found: {island_count}")
            print(f"  - Charts detected: {'Yes' if charts_found else 'No'}")
            print(f"  - Named ranges: {'Yes' if named_ranges else 'No'}")
            print(f"  - Data validation: {'Yes' if data_validation else 'No'}")
        else:
            print(f"✗ Failed to analyze {test_file.name}: {error}")
        
        print()
    
    # Print summary
    print("=" * 60)
    print("SUMMARY")
    print("=" * 60)
    
    successful_tests = [r for r in results if r['success']]
    failed_tests = [r for r in results if not r['success']]
    
    print(f"Total files tested: {len(results)}")
    print(f"Successful analyses: {len(successful_tests)}")
    print(f"Failed analyses: {len(failed_tests)}")
    print()
    
    if successful_tests:
        print("Successful Tests:")
        for result in successful_tests:
            print(f"  ✓ {result['file']}")
            print(f"    - Tables: {result['table_count']}, Islands: {result['island_count']}")
            if result['charts_found']:
                print(f"    - Charts: Yes")
            if result['named_ranges']:
                print(f"    - Named Ranges: Yes")
            if result['data_validation']:
                print(f"    - Data Validation: Yes")
        print()
    
    if failed_tests:
        print("Failed Tests:")
        for result in failed_tests:
            print(f"  ✗ {result['file']}: {result['error']}")
        print()
    
    # Feature detection summary
    print("Feature Detection Summary:")
    print(f"  - Files with formal tables: {sum(1 for r in successful_tests if r['tables_found'])}")
    print(f"  - Files with data islands: {sum(1 for r in successful_tests if r['islands_found'])}")
    print(f"  - Files with charts: {sum(1 for r in successful_tests if r['charts_found'])}")
    print(f"  - Files with named ranges: {sum(1 for r in successful_tests if r['named_ranges'])}")
    print(f"  - Files with data validation: {sum(1 for r in successful_tests if r['data_validation'])}")
    print(f"  - Files with external links: {sum(1 for r in successful_tests if r['external_links'])}")
    print(f"  - Files with VBA: {sum(1 for r in successful_tests if r['vba_detected'])}")
    
    return results

if __name__ == "__main__":
    test_all_files() 