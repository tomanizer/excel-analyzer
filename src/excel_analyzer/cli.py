#!/usr/bin/env python3
"""
Command Line Interface for Excel Analyzer

Provides a professional CLI with configurable output options and flexible file handling.
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Optional, Dict, Any
import pandas as pd

from .excel_parser import analyze_workbook_final, generate_markdown_report, extract_data_to_dataframes


def create_parser() -> argparse.ArgumentParser:
    """Create and configure the argument parser."""
    parser = argparse.ArgumentParser(
        prog="excel-analyzer",
        description="Comprehensive Excel file analysis tool for financial models",
        epilog="""
Examples:
  excel-analyzer file.xlsx                           # Basic analysis
  excel-analyzer file.xlsx --output-dir ./results    # Custom output directory
  excel-analyzer file.xlsx --json --markdown         # Generate reports
  excel-analyzer file.xlsx --dataframes --save-dfs   # Extract and save DataFrames
  excel-analyzer *.xlsx --batch                      # Process multiple files
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    # Input file argument
    parser.add_argument(
        "file",
        nargs="+",
        help="Excel file(s) to analyze (.xlsx, .xlsm)"
    )
    
    # Output options
    parser.add_argument(
        "--output-dir", "-o",
        type=Path,
        default=Path("reports"),
        help="Directory to save output files (default: ./reports)"
    )
    
    parser.add_argument(
        "--json", "-j",
        action="store_true",
        help="Generate JSON report with structured analysis data"
    )
    
    parser.add_argument(
        "--markdown", "-m",
        action="store_true",
        help="Generate markdown report with formatted analysis"
    )
    
    parser.add_argument(
        "--dataframes", "-d",
        action="store_true",
        help="Extract data to pandas DataFrames"
    )
    
    parser.add_argument(
        "--save-dfs",
        action="store_true",
        help="Save extracted DataFrames to CSV files"
    )
    
    parser.add_argument(
        "--dfs-format",
        choices=["csv", "excel", "parquet"],
        default="csv",
        help="Format for saving DataFrames (default: csv)"
    )
    
    # Processing options
    parser.add_argument(
        "--batch", "-b",
        action="store_true",
        help="Process multiple files in batch mode"
    )
    
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose output"
    )
    
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress non-essential output"
    )
    
    parser.add_argument(
        "--summary",
        action="store_true",
        help="Show summary statistics only"
    )
    
    return parser


def validate_file(file_path: Path) -> bool:
    """Validate that the file exists and is an Excel file."""
    if not file_path.exists():
        print(f"❌ Error: File not found: {file_path}")
        return False
    
    if not file_path.suffix.lower() in ['.xlsx', '.xlsm']:
        print(f"❌ Error: Not an Excel file: {file_path}")
        return False
    
    return True


def save_dataframes(dataframes: Dict[str, pd.DataFrame], 
                   output_dir: Path, 
                   file_stem: str, 
                   format_type: str = "csv") -> None:
    """Save DataFrames to files in the specified format."""
    dfs_dir = output_dir / "dataframes" / file_stem
    dfs_dir.mkdir(parents=True, exist_ok=True)
    
    for name, df in dataframes.items():
        if df is None:
            continue
            
        # Clean filename
        safe_name = name.replace(":", "_").replace("/", "_").replace("\\", "_")
        
        if format_type == "csv":
            output_file = dfs_dir / f"{safe_name}.csv"
            df.to_csv(output_file, index=False)
        elif format_type == "excel":
            output_file = dfs_dir / f"{safe_name}.xlsx"
            df.to_excel(output_file, index=False)
        elif format_type == "parquet":
            output_file = dfs_dir / f"{safe_name}.parquet"
            df.to_parquet(output_file, index=False)
        
        print(f"  📊 Saved DataFrame '{name}' to: {output_file}")


def process_single_file(file_path: Path, 
                       args: argparse.Namespace) -> Dict[str, Any]:
    """Process a single Excel file and return results."""
    results = {
        "file": file_path.name,
        "success": False,
        "error": None,
        "outputs": []
    }
    
    try:
        if args.verbose:
            print(f"🔍 Analyzing: {file_path}")
        
        # Create output directory
        args.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Get analysis data
        analysis_data = analyze_workbook_final(file_path, return_data=True)
        results["success"] = True
        
        # Generate JSON report
        if args.json:
            json_file = args.output_dir / f"{file_path.stem}.json"
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(analysis_data, f, indent=2, default=str)
            results["outputs"].append(f"JSON: {json_file}")
            if not args.quiet:
                print(f"📄 JSON report saved to: {json_file}")
        
        # Generate markdown report
        if args.markdown:
            markdown_file = args.output_dir / f"{file_path.stem}.md"
            generate_markdown_report(analysis_data, markdown_file)
            results["outputs"].append(f"Markdown: {markdown_file}")
            if not args.quiet:
                print(f"📝 Markdown report saved to: {markdown_file}")
        
        # Extract DataFrames
        if args.dataframes:
            dataframes = extract_data_to_dataframes(analysis_data, file_path)
            results["dataframes"] = len(dataframes)
            
            if not args.quiet:
                print(f"🐼 Extracted {len(dataframes)} DataFrames:")
                for name, df in dataframes.items():
                    if df is not None:
                        print(f"  - {name}: {df.shape[0]} rows × {df.shape[1]} columns")
                    else:
                        print(f"  - {name}: Error extracting data")
            
            # Save DataFrames if requested
            if args.save_dfs:
                save_dataframes(dataframes, args.output_dir, file_path.stem, args.dfs_format)
        
        # Show summary if requested
        if args.summary:
            summary = analysis_data["summary"]
            print(f"\n📊 Summary for {file_path.name}:")
            print(f"  Sheets: {summary['total_sheets']}")
            print(f"  Formal Tables: {summary['total_formal_tables']}")
            print(f"  Pivot Tables: {summary['total_pivot_tables']}")
            print(f"  Charts: {summary['total_charts']}")
            print(f"  Data Islands: {summary['total_data_islands']}")
            print(f"  Data Validation Rules: {summary['total_data_validation_rules']}")
        
        # Standard console output if no specific outputs requested
        if not any([args.json, args.markdown, args.dataframes, args.summary]):
            if not args.quiet:
                print(f"\n--- Analysis for: {file_path.name} ---")
            analyze_workbook_final(file_path)
        
    except Exception as e:
        results["success"] = False
        results["error"] = str(e)
        print(f"❌ Error processing {file_path}: {e}")
    
    return results


def main():
    """Main CLI entry point."""
    parser = create_parser()
    args = parser.parse_args()
    
    # Handle help
    if len(args.file) == 1 and args.file[0] in ['-h', '--help']:
        parser.print_help()
        return
    
    # Validate files
    valid_files = []
    for file_pattern in args.file:
        if '*' in file_pattern or '?' in file_pattern:
            # Handle glob patterns
            pattern_path = Path(file_pattern)
            matching_files = list(pattern_path.parent.glob(pattern_path.name))
            valid_files.extend([f for f in matching_files if validate_file(f)])
        else:
            # Single file
            file_path = Path(file_pattern)
            if validate_file(file_path):
                valid_files.append(file_path)
    
    if not valid_files:
        print("❌ No valid Excel files found to process.")
        sys.exit(1)
    
    # Process files
    results = []
    total_files = len(valid_files)
    
    if args.batch or total_files > 1:
        print(f"🚀 Processing {total_files} file(s)...")
    
    for i, file_path in enumerate(valid_files, 1):
        if args.batch or total_files > 1:
            print(f"\n[{i}/{total_files}] Processing: {file_path.name}")
        
        result = process_single_file(file_path, args)
        results.append(result)
    
    # Summary
    successful = sum(1 for r in results if r["success"])
    failed = len(results) - successful
    
    if args.batch or total_files > 1:
        print(f"\n✅ Processing complete!")
        print(f"   Successfully processed: {successful}/{total_files}")
        if failed > 0:
            print(f"   Failed: {failed}")
        
        if args.output_dir:
            print(f"   Output directory: {args.output_dir.absolute()}")


if __name__ == "__main__":
    main() 