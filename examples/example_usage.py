#!/usr/bin/env python3
"""
Example Usage of Excel Analyzer

This script demonstrates how to use the Excel analyzer programmatically
to get structured data and work with pandas DataFrames.
"""

from pathlib import Path
from excel_analyzer.excel_parser import analyze_workbook_final, generate_markdown_report, extract_data_to_dataframes
import json

def main():
    """Demonstrate the new Excel analyzer functionality."""
    
    # File to analyze
    file_path = Path("excel_files/mycoolsample.xlsx")
    
    print("🔍 Excel Analyzer - Programmatic Usage Example")
    print("=" * 50)
    print()
    
    # 1. Get structured analysis data
    print("1. 📊 Getting structured analysis data...")
    analysis_data = analyze_workbook_final(file_path, return_data=True)
    
    print(f"   ✅ Analyzed {analysis_data['metadata']['filename']}")
    print(f"   📁 File size: {analysis_data['metadata']['file_size_kb']:.1f} KB")
    print(f"   📋 Total sheets: {analysis_data['summary']['total_sheets']}")
    print(f"   📊 Total tables/islands: {analysis_data['summary']['total_data_islands'] + analysis_data['summary']['total_formal_tables']}")
    print()
    
    # 2. Extract data to pandas DataFrames
    print("2. 🐼 Extracting data to pandas DataFrames...")
    dataframes = extract_data_to_dataframes(analysis_data, file_path)
    
    print(f"   ✅ Extracted {len(dataframes)} DataFrames:")
    for name, df in dataframes.items():
        if df is not None:
            print(f"      - {name}: {df.shape[0]} rows × {df.shape[1]} columns")
            # Show first few rows for larger DataFrames
            if df.shape[0] > 1 and df.shape[1] > 1:
                print(f"        Preview:")
                print(f"        {df.head(2).to_string()}")
                print()
        else:
            print(f"      - {name}: Error extracting data")
    print()
    
    # 3. Generate markdown report
    print("3. 📝 Generating markdown report...")
    report = generate_markdown_report(analysis_data)
    print(f"   ✅ Generated {len(report)} character report")
    print()
    
    # 4. Save structured data as JSON
    print("4. 💾 Saving structured data as JSON...")
    json_file = Path("reports") / f"{file_path.stem}.analysis.json"
    json_file.parent.mkdir(exist_ok=True)
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(analysis_data, f, indent=2, default=str)
    print(f"   ✅ Saved to {json_file}")
    print()
    
    # 5. Example: Working with specific DataFrames
    print("5. 🔧 Example: Working with specific DataFrames...")
    
    # Find the formal table
    formal_tables = [t for t in analysis_data['all_tables'] if t['type'] == 'Formal Table']
    if formal_tables:
        table_name = formal_tables[0]['name']
        if table_name in dataframes and dataframes[table_name] is not None:
            df = dataframes[table_name]
            print(f"   📊 Formal table '{table_name}':")
            print(f"      Shape: {df.shape}")
            print(f"      Columns: {list(df.columns)}")
            print(f"      Data types: {df.dtypes.to_dict()}")
            print()
    
    # Find pivot tables
    pivot_tables = [t for t in analysis_data['all_tables'] if 'Pivot' in str(t.get('name', ''))]
    if pivot_tables:
        print(f"   🔄 Found {len(pivot_tables)} pivot table(s):")
        for pivot in pivot_tables:
            print(f"      - {pivot['name']} on {pivot['sheet']} at {pivot['range']}")
        print()
    
    print("🎉 Analysis complete! You can now:")
    print("   - Use the DataFrames for data analysis")
    print("   - Recreate pivot tables using pandas")
    print("   - Process the structured JSON data")
    print("   - Generate custom reports from the markdown")

if __name__ == "__main__":
    main() 