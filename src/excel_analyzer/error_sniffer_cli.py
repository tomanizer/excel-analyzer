#!/usr/bin/env python3
"""
Command Line Interface for Excel Error Sniffer

Provides a CLI for detecting and analyzing Excel errors and issues.
"""

import argparse
import sys
from pathlib import Path
import time
from typing import List

from .excel_error_sniffer import ExcelErrorSniffer, sniff_excel_errors


def create_parser() -> argparse.ArgumentParser:
    """Create the argument parser for the error sniffer CLI."""
    parser = argparse.ArgumentParser(
        description="Excel Error Sniffer - Detect and analyze Excel errors and issues",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s file.xlsx                    # Basic error detection
  %(prog)s file.xlsx --json --markdown  # Generate reports
  %(prog)s "*.xlsx" --batch --summary   # Batch processing
  %(prog)s file.xlsx --verbose --timing # Detailed output with timing
        """
    )
    
    parser.add_argument(
        'files',
        nargs='+',
        help='Excel file(s) to analyze (supports glob patterns)'
    )
    
    parser.add_argument(
        '--output-dir', '-o',
        type=Path,
        default=Path("error_reports"),
        help='Directory to save output files (default: error_reports)'
    )
    
    parser.add_argument(
        '--json', '-j',
        action='store_true',
        help='Generate JSON error report'
    )
    
    parser.add_argument(
        '--markdown', '-m',
        action='store_true',
        help='Generate markdown error report'
    )
    
    parser.add_argument(
        '--summary', '-s',
        action='store_true',
        help='Show summary statistics only'
    )
    
    parser.add_argument(
        '--batch', '-b',
        action='store_true',
        help='Process multiple files in batch mode'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose output'
    )
    
    parser.add_argument(
        '--quiet', '-q',
        action='store_true',
        help='Suppress non-essential output'
    )
    
    parser.add_argument(
        '--timing', '-t',
        action='store_true',
        help='Show detailed timing information'
    )
    
    parser.add_argument(
        '--severity',
        choices=['high', 'medium', 'low', 'all'],
        default='all',
        help='Filter issues by severity (default: all)'
    )
    
    return parser


def validate_file(file_path: Path) -> bool:
    """Validate that the file exists and is an Excel file."""
    if not file_path.exists():
        print(f"❌ File not found: {file_path}")
        return False
    
    if not file_path.suffix.lower() in ['.xlsx', '.xlsm', '.xls']:
        print(f"❌ Not an Excel file: {file_path}")
        return False
    
    return True


def process_single_file(file_path: Path, args: argparse.Namespace) -> dict:
    """Process a single Excel file for errors."""
    if not validate_file(file_path):
        return {'success': False, 'error': 'File validation failed'}
    
    start_time = time.time()
    
    try:
        if args.verbose:
            print(f"🔍 Analyzing errors in: {file_path.name}")
        
        # Create output directory if needed
        if args.json or args.markdown:
            args.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Perform error detection
        sniffer = ExcelErrorSniffer(file_path)
        errors = sniffer.sniff_errors()
        
        # Filter by severity if requested
        if args.severity != 'all':
            filtered_errors = {}
            for error_type, error_list in errors.items():
                if isinstance(error_list, list):
                    filtered_errors[error_type] = [
                        error for error in error_list 
                        if error.get('severity', 'low') == args.severity
                    ]
                else:
                    filtered_errors[error_type] = error_list
            errors = filtered_errors
        
        # Generate outputs
        if args.json:
            json_path = args.output_dir / f"{file_path.stem}_error_analysis.json"
            sniffer.save_json(json_path)
            if not args.quiet:
                print(f"📄 JSON report saved to: {json_path}")
        
        if args.markdown:
            markdown_path = args.output_dir / f"{file_path.stem}_error_analysis.md"
            sniffer.save_markdown(markdown_path)
            if not args.quiet:
                print(f"📝 Markdown report saved to: {markdown_path}")
        
        # Show summary
        if args.summary:
            summary = errors['summary']
            print(f"\n📊 Error Summary for {file_path.name}:")
            print(f"  🔴 High Severity: {summary['severity_breakdown']['high']}")
            print(f"  🟡 Medium Severity: {summary['severity_breakdown']['medium']}")
            print(f"  🟢 Low Severity: {summary['severity_breakdown']['low']}")
            print(f"  📋 Total Issues: {summary['total_issues']}")
        
        # Show timing
        if args.timing:
            processing_time = time.time() - start_time
            print(f"\n⏱️  Processing time: {processing_time:.3f}s")
        
        # Show detailed errors if not quiet and not summary-only
        if not args.quiet and not args.summary:
            summary = errors['summary']
            if summary['total_issues'] > 0:
                print(f"\n🔍 Found {summary['total_issues']} issues in {file_path.name}:")
                
                for error_type, error_list in errors.items():
                    if isinstance(error_list, list) and error_list:
                        print(f"\n  📋 {error_type.replace('_', ' ').title()}: {len(error_list)}")
                        for error in error_list[:5]:  # Show first 5 errors
                            severity = error.get('severity', 'low')
                            severity_emoji = {'high': '🔴', 'medium': '🟡', 'low': '🟢'}[severity]
                            print(f"    {severity_emoji} {error.get('description', 'Unknown error')}")
                        
                        if len(error_list) > 5:
                            print(f"    ... and {len(error_list) - 5} more")
            else:
                print(f"\n✅ No issues found in {file_path.name}")
        
        return {
            'success': True,
            'file': file_path.name,
            'total_issues': errors['summary']['total_issues'],
            'processing_time': time.time() - start_time
        }
        
    except Exception as e:
        print(f"❌ Error processing {file_path}: {e}")
        return {
            'success': False,
            'file': file_path.name,
            'error': str(e),
            'processing_time': time.time() - start_time
        }


def main():
    """Main CLI entry point."""
    parser = create_parser()
    args = parser.parse_args()
    
    # Handle glob patterns
    import glob
    all_files = []
    for pattern in args.files:
        if '*' in pattern or '?' in pattern:
            # Handle glob pattern
            matched_files = glob.glob(pattern)
            if not matched_files:
                print(f"⚠️  No files matched pattern: {pattern}")
            all_files.extend([Path(f) for f in matched_files])
        else:
            # Single file
            all_files.append(Path(pattern))
    
    if not all_files:
        print("❌ No valid files to process.")
        sys.exit(1)
    
    # Remove duplicates and sort
    all_files = sorted(list(set(all_files)))
    
    if args.verbose:
        print(f"🚀 Processing {len(all_files)} file(s) for errors...")
    
    results = []
    for i, file_path in enumerate(all_files, 1):
        if args.verbose and len(all_files) > 1:
            print(f"\n[{i}/{len(all_files)}] Processing: {file_path.name}")
        
        result = process_single_file(file_path, args)
        results.append(result)
    
    # Summary for batch processing
    if len(all_files) > 1:
        successful = sum(1 for r in results if r['success'])
        failed = len(results) - successful
        total_issues = sum(r.get('total_issues', 0) for r in results if r['success'])
        
        print(f"\n✅ Error detection complete!")
        print(f"   Successfully processed: {successful}/{len(results)}")
        if failed > 0:
            print(f"   Failed: {failed}")
        print(f"   Total issues found: {total_issues}")
        
        if args.output_dir:
            print(f"   Output directory: {args.output_dir.absolute()}")


if __name__ == '__main__':
    main() 