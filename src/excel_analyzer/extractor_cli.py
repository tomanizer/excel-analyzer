#!/usr/bin/env python3
"""
Command Line Interface for Excel Extractor

Provides a professional CLI with configurable output options for comprehensive Excel data extraction.
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Optional, Dict, Any
import time
from datetime import datetime

from .excel_extractor import ExcelExtractor, extract_excel_to_markdown


def create_parser() -> argparse.ArgumentParser:
    """Create and configure the argument parser."""
    parser = argparse.ArgumentParser(
        prog="excel-extractor",
        description="Comprehensive Excel data extraction tool for AI/LLM analysis",
        epilog="""
Examples:
  excel-extractor file.xlsx                           # Basic extraction
  excel-extractor file.xlsx --output-dir ./results    # Custom output directory
  excel-extractor file.xlsx --markdown --json         # Generate both formats
  excel-extractor file.xlsx --llm-optimized           # LLM-optimized output
  excel-extractor *.xlsx --batch                      # Process multiple files
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    # Input file argument
    parser.add_argument(
        "file",
        nargs="+",
        help="Excel file(s) to extract (.xlsx, .xlsm)"
    )
    
    # Output options
    parser.add_argument(
        "--output-dir", "-o",
        type=Path,
        default=Path("extractor_reports"),
        help="Directory to save output files (default: ./extractor_reports)"
    )
    
    parser.add_argument(
        "--markdown", "-m",
        action="store_true",
        help="Generate markdown report (LLM-optimized)"
    )
    
    parser.add_argument(
        "--json", "-j",
        action="store_true",
        help="Generate JSON data export"
    )
    
    parser.add_argument(
        "--llm-optimized",
        action="store_true",
        help="Optimize output for LLM/AI analysis (implies --markdown)"
    )
    
    parser.add_argument(
        "--include-formulas",
        action="store_true",
        help="Include detailed formula analysis in output"
    )
    
    parser.add_argument(
        "--include-styles",
        action="store_true",
        help="Include cell style and formatting information"
    )
    
    parser.add_argument(
        "--include-relationships",
        action="store_true",
        help="Include cross-sheet relationships and dependencies"
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
    
    parser.add_argument(
        "--timing",
        action="store_true",
        help="Show detailed timing information"
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


def process_single_file(file_path: Path, 
                       args: argparse.Namespace) -> Dict[str, Any]:
    """Process a single Excel file and return results."""
    results = {
        "file": file_path.name,
        "success": False,
        "error": None,
        "outputs": [],
        "timing": {}
    }
    
    start_time = time.time()
    
    try:
        if args.verbose:
            print(f"🔍 Extracting data from: {file_path}")
        
        # Create output directory
        args.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Create extractor instance
        extractor = ExcelExtractor(file_path)
        
        # Extract all data
        extraction_start = time.time()
        extracted_data = extractor.extract_all()
        extraction_time = time.time() - extraction_start
        
        results["success"] = True
        results["timing"]["extraction"] = extraction_time
        
        # Generate markdown report
        if args.markdown or args.llm_optimized:
            markdown_start = time.time()
            markdown_content = extractor.to_markdown()
            markdown_time = time.time() - markdown_start
            
            # Save markdown report
            markdown_file = args.output_dir / f"{file_path.stem}_extractor_report.md"
            with open(markdown_file, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            
            results["outputs"].append(f"Markdown: {markdown_file}")
            results["timing"]["markdown"] = markdown_time
            
            if not args.quiet:
                print(f"📝 Markdown report saved to: {markdown_file}")
        
        # Save JSON data
        if args.json:
            json_start = time.time()
            json_file = args.output_dir / f"{file_path.stem}_extracted_data.json"
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(extracted_data, f, indent=2, default=str)
            json_time = time.time() - json_start
            
            results["outputs"].append(f"JSON: {json_file}")
            results["timing"]["json"] = json_time
            
            if not args.quiet:
                print(f"📄 JSON data saved to: {json_file}")
        
        # Show summary if requested
        if args.summary:
            summary = extracted_data.get('summary', {})
            metadata = extracted_data.get('metadata', {})
            
            print(f"\n📊 Summary for {file_path.name}:")
            print(f"  📁 File size: {metadata.get('file_size_kb', 0):.2f} KB")
            print(f"  📊 Sheets: {metadata.get('sheet_count', 0)}")
            print(f"  📝 Cells with data: {summary.get('total_cells_with_data', 0):,}")
            print(f"  🧮 Formulas: {summary.get('total_formulas', 0):,}")
            print(f"  📋 Tables: {summary.get('total_tables', 0)}")
            print(f"  📈 Charts: {summary.get('total_charts', 0)}")
            print(f"  🔗 Cross-sheet references: {len(extracted_data.get('relationships', {}).get('cross_sheet_references', []))}")
        
        # Show timing information
        if args.timing:
            total_time = time.time() - start_time
            print(f"\n⏱️  Timing for {file_path.name}:")
            print(f"  🔍 Data extraction: {extraction_time:.3f}s")
            if args.markdown or args.llm_optimized:
                print(f"  📝 Markdown generation: {markdown_time:.3f}s")
            if args.json:
                print(f"  📄 JSON export: {json_time:.3f}s")
            print(f"  ⏱️  Total time: {total_time:.3f}s")
        
        # Standard console output if no specific outputs requested
        if not any([args.markdown, args.json, args.summary]):
            if not args.quiet:
                print(f"\n--- Extraction for: {file_path.name} ---")
            
            # Use the extractor's built-in console output
            extractor.extract_all()  # This will print progress messages
        
    except Exception as e:
        total_time = time.time() - start_time
        results["success"] = False
        results["error"] = str(e)
        results["timing"]["total"] = total_time
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
    overall_start = time.time()
    
    if args.batch or total_files > 1:
        print(f"🚀 Processing {total_files} file(s) with Excel Extractor...")
    
    for i, file_path in enumerate(valid_files, 1):
        if args.batch or total_files > 1:
            print(f"\n[{i}/{total_files}] Processing: {file_path.name}")
        
        result = process_single_file(file_path, args)
        results.append(result)
    
    # Summary
    overall_time = time.time() - overall_start
    successful = sum(1 for r in results if r["success"])
    failed = len(results) - successful
    
    if args.batch or total_files > 1:
        print(f"\n✅ Extraction complete!")
        print(f"   Successfully processed: {successful}/{total_files}")
        if failed > 0:
            print(f"   Failed: {failed}")
        
        if args.output_dir:
            print(f"   Output directory: {args.output_dir.absolute()}")
        
        if args.timing:
            print(f"   ⏱️  Total processing time: {overall_time:.2f}s")
            print(f"   ⏱️  Average time per file: {overall_time/len(results):.2f}s")


if __name__ == "__main__":
    main() 