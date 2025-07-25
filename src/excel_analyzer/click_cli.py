#!/usr/bin/env python3
"""
Click-based Command Line Interface for Excel Analyzer

Provides an elegant CLI using Click decorators with the same functionality
as the argparse version but with a more modern, decorator-based approach.
"""

import click
import json
import sys
from pathlib import Path
from typing import Optional, Dict, Any
import time
from datetime import datetime

from .excel_parser import analyze_workbook_final, generate_markdown_report, extract_data_to_dataframes
from .excel_extractor import ExcelExtractor


def validate_excel_file(ctx, param, value):
    """Validate that the file exists and is an Excel file."""
    if not value:
        return value
    
    # Handle both single files and tuples of files
    if isinstance(value, tuple):
        validated_files = []
        for file_path_str in value:
            file_path = Path(file_path_str)
            if not file_path.exists():
                raise click.BadParameter(f"File not found: {file_path}")
            
            if not file_path.suffix.lower() in ['.xlsx', '.xlsm']:
                raise click.BadParameter(f"Not an Excel file: {file_path}")
            
            validated_files.append(file_path)
        return validated_files
    else:
        file_path = Path(value)
        if not file_path.exists():
            raise click.BadParameter(f"File not found: {file_path}")
        
        if not file_path.suffix.lower() in ['.xlsx', '.xlsm']:
            raise click.BadParameter(f"Not an Excel file: {file_path}")
        
        return [file_path]


def process_files_with_parser(files, output_dir, json_output, markdown_output, dataframes, save_dfs, dfs_format, summary, verbose, quiet):
    """Process files using the Excel Parser."""
    results = []
    total_files = len(files)
    
    if verbose and total_files > 1:
        click.echo(f"🚀 Processing {total_files} file(s) with Excel Parser...")
    
    for i, file_path in enumerate(files, 1):
        if verbose and total_files > 1:
            click.echo(f"\n[{i}/{total_files}] Processing: {file_path.name}")
        
        start_time = time.time()
        
        try:
            # Create output directory
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Get analysis data
            analysis_data = analyze_workbook_final(file_path, return_data=True)
            
            # Generate JSON report
            if json_output:
                json_file = output_dir / f"{file_path.stem}.json"
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(analysis_data, f, indent=2, default=str)
                if not quiet:
                    click.echo(f"📄 JSON report saved to: {json_file}")
            
            # Generate markdown report
            if markdown_output:
                markdown_file = output_dir / f"{file_path.stem}.md"
                generate_markdown_report(analysis_data, markdown_file)
                if not quiet:
                    click.echo(f"📝 Markdown report saved to: {markdown_file}")
            
            # Extract DataFrames
            if dataframes:
                dataframes_dict = extract_data_to_dataframes(analysis_data, file_path)
                
                if not quiet:
                    click.echo(f"🐼 Extracted {len(dataframes_dict)} DataFrames:")
                    for name, df in dataframes_dict.items():
                        if df is not None:
                            click.echo(f"  - {name}: {df.shape[0]} rows × {df.shape[1]} columns")
                        else:
                            click.echo(f"  - {name}: Error extracting data")
                
                # Save DataFrames if requested
                if save_dfs:
                    dfs_dir = output_dir / "dataframes" / file_path.stem
                    dfs_dir.mkdir(parents=True, exist_ok=True)
                    
                    for name, df in dataframes_dict.items():
                        if df is None:
                            continue
                        
                        # Clean filename
                        safe_name = name.replace(":", "_").replace("/", "_").replace("\\", "_")
                        
                        if dfs_format == "csv":
                            output_file = dfs_dir / f"{safe_name}.csv"
                            df.to_csv(output_file, index=False)
                        elif dfs_format == "excel":
                            output_file = dfs_dir / f"{safe_name}.xlsx"
                            df.to_excel(output_file, index=False)
                        elif dfs_format == "parquet":
                            output_file = dfs_dir / f"{safe_name}.parquet"
                            df.to_parquet(output_file, index=False)
                        
                        click.echo(f"  📊 Saved DataFrame '{name}' to: {output_file}")
            
            # Show summary if requested
            if summary:
                summary_data = analysis_data["summary"]
                click.echo(f"\n📊 Summary for {file_path.name}:")
                click.echo(f"  Sheets: {summary_data['total_sheets']}")
                click.echo(f"  Formal Tables: {summary_data['total_formal_tables']}")
                click.echo(f"  Pivot Tables: {summary_data['total_pivot_tables']}")
                click.echo(f"  Charts: {summary_data['total_charts']}")
                click.echo(f"  Data Islands: {summary_data['total_data_islands']}")
                click.echo(f"  Data Validation Rules: {summary_data['total_data_validation_rules']}")
            
            # Standard console output if no specific outputs requested
            if not any([json_output, markdown_output, dataframes, summary]):
                if not quiet:
                    click.echo(f"\n--- Analysis for: {file_path.name} ---")
                analyze_workbook_final(file_path)
            
            results.append({
                "file": file_path.name,
                "success": True,
                "processing_time": time.time() - start_time
            })
            
        except Exception as e:
            results.append({
                "file": file_path.name,
                "success": False,
                "error": str(e),
                "processing_time": time.time() - start_time
            })
            click.echo(f"❌ Error processing {file_path}: {e}")
    
    return results


def process_files_with_extractor(files, output_dir, json_output, markdown_output, llm_optimized, summary, verbose, quiet, timing):
    """Process files using the Excel Extractor."""
    results = []
    total_files = len(files)
    
    if verbose and total_files > 1:
        click.echo(f"🚀 Processing {total_files} file(s) with Excel Extractor...")
    
    for i, file_path in enumerate(files, 1):
        if verbose and total_files > 1:
            click.echo(f"\n[{i}/{total_files}] Processing: {file_path.name}")
        
        start_time = time.time()
        
        try:
            # Create output directory
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Create extractor instance
            extractor = ExcelExtractor(file_path)
            
            # Extract all data
            extraction_start = time.time()
            extracted_data = extractor.extract_all()
            extraction_time = time.time() - extraction_start
            
            # Generate markdown report
            if markdown_output or llm_optimized:
                markdown_start = time.time()
                markdown_content = extractor.to_markdown()
                markdown_time = time.time() - markdown_start
                
                # Save markdown report
                markdown_file = output_dir / f"{file_path.stem}_extractor_report.md"
                with open(markdown_file, 'w', encoding='utf-8') as f:
                    f.write(markdown_content)
                
                if not quiet:
                    click.echo(f"📝 Markdown report saved to: {markdown_file}")
            
            # Save JSON data
            if json_output:
                json_start = time.time()
                json_file = output_dir / f"{file_path.stem}_extracted_data.json"
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(extracted_data, f, indent=2, default=str)
                json_time = time.time() - json_start
                
                if not quiet:
                    click.echo(f"📄 JSON data saved to: {json_file}")
            
            # Show summary if requested
            if summary:
                summary_data = extracted_data.get('summary', {})
                metadata = extracted_data.get('metadata', {})
                
                click.echo(f"\n📊 Summary for {file_path.name}:")
                click.echo(f"  📁 File size: {metadata.get('file_size_kb', 0):.2f} KB")
                click.echo(f"  📊 Sheets: {metadata.get('sheet_count', 0)}")
                click.echo(f"  📝 Cells with data: {summary_data.get('total_cells_with_data', 0):,}")
                click.echo(f"  🧮 Formulas: {summary_data.get('total_formulas', 0):,}")
                click.echo(f"  📋 Tables: {summary_data.get('total_tables', 0)}")
                click.echo(f"  📈 Charts: {summary_data.get('total_charts', 0)}")
                click.echo(f"  🔗 Cross-sheet references: {len(extracted_data.get('relationships', {}).get('cross_sheet_references', []))}")
            
            # Show timing information
            if timing:
                total_time = time.time() - start_time
                click.echo(f"\n⏱️  Timing for {file_path.name}:")
                click.echo(f"  🔍 Data extraction: {extraction_time:.3f}s")
                if markdown_output or llm_optimized:
                    click.echo(f"  📝 Markdown generation: {markdown_time:.3f}s")
                if json_output:
                    click.echo(f"  📄 JSON export: {json_time:.3f}s")
                click.echo(f"  ⏱️  Total time: {total_time:.3f}s")
            
            results.append({
                "file": file_path.name,
                "success": True,
                "processing_time": time.time() - start_time
            })
            
        except Exception as e:
            results.append({
                "file": file_path.name,
                "success": False,
                "error": str(e),
                "processing_time": time.time() - start_time
            })
            click.echo(f"❌ Error processing {file_path}: {e}")
    
    return results


@click.group()
@click.version_option(version="0.1.0", prog_name="excel-analyzer")
def cli():
    """
    Excel Analyzer - Comprehensive Excel file analysis tool.
    
    This tool provides two main analysis modes:
    
    \b
    PARSER: Fast analysis and DataFrame extraction
    EXTRACTOR: Comprehensive data extraction for AI/LLM analysis
    
    Choose the appropriate command based on your needs.
    """
    pass


@cli.command()
@click.argument('files', nargs=-1, required=True, callback=validate_excel_file)
@click.option('--output-dir', '-o', type=click.Path(file_okay=False, dir_okay=True, path_type=Path), 
              default=Path("reports"), help="Directory to save output files")
@click.option('--json', '-j', is_flag=True, help="Generate JSON report")
@click.option('--markdown', '-m', is_flag=True, help="Generate markdown report")
@click.option('--dataframes', '-d', is_flag=True, help="Extract data to pandas DataFrames")
@click.option('--save-dfs', is_flag=True, help="Save extracted DataFrames to files")
@click.option('--dfs-format', type=click.Choice(['csv', 'excel', 'parquet']), 
              default='csv', help="Format for saving DataFrames")
@click.option('--summary', is_flag=True, help="Show summary statistics only")
@click.option('--verbose', '-v', is_flag=True, help="Enable verbose output")
@click.option('--quiet', '-q', is_flag=True, help="Suppress non-essential output")
def parser(files, output_dir, json, markdown, dataframes, save_dfs, dfs_format, summary, verbose, quiet):
    """
    Analyze Excel files using the fast parser.
    
    This command provides quick analysis and DataFrame extraction.
    Ideal for:
    - Quick analysis and discovery
    - DataFrame extraction for analysis
    - Summary reports for decision making
    - Command-line automation
    """
    if not files:
        click.echo("❌ No valid Excel files found to process.")
        sys.exit(1)
    
    results = process_files_with_parser(
        files, output_dir, json, markdown, dataframes, save_dfs, dfs_format, summary, verbose, quiet
    )
    
    # Summary
    if len(files) > 1:
        successful = sum(1 for r in results if r["success"])
        failed = len(results) - successful
        
        click.echo(f"\n✅ Processing complete!")
        click.echo(f"   Successfully processed: {successful}/{len(results)}")
        if failed > 0:
            click.echo(f"   Failed: {failed}")
        
        if output_dir:
            click.echo(f"   Output directory: {output_dir.absolute()}")


@cli.command()
@click.argument('files', nargs=-1, required=True, callback=validate_excel_file)
@click.option('--output-dir', '-o', type=click.Path(file_okay=False, dir_okay=True, path_type=Path), 
              default=Path("extractor_reports"), help="Directory to save output files")
@click.option('--json', '-j', is_flag=True, help="Generate JSON data export")
@click.option('--markdown', '-m', is_flag=True, help="Generate markdown report (LLM-optimized)")
@click.option('--llm-optimized', is_flag=True, help="Optimize output for LLM/AI analysis")
@click.option('--summary', is_flag=True, help="Show summary statistics only")
@click.option('--verbose', '-v', is_flag=True, help="Enable verbose output")
@click.option('--quiet', '-q', is_flag=True, help="Suppress non-essential output")
@click.option('--timing', is_flag=True, help="Show detailed timing information")
def extractor(files, output_dir, json, markdown, llm_optimized, summary, verbose, quiet, timing):
    """
    Extract comprehensive data from Excel files.
    
    This command provides complete data extraction optimized for AI/LLM analysis.
    Ideal for:
    - AI/LLM analysis of Excel files
    - Complete documentation of financial models
    - Forensic analysis of Excel structure
    - Data auditing and compliance
    - Dependency mapping in complex models
    """
    if not files:
        click.echo("❌ No valid Excel files found to process.")
        sys.exit(1)
    
    results = process_files_with_extractor(
        files, output_dir, json, markdown, llm_optimized, summary, verbose, quiet, timing
    )
    
    # Summary
    if len(files) > 1:
        successful = sum(1 for r in results if r["success"])
        failed = len(results) - successful
        
        click.echo(f"\n✅ Extraction complete!")
        click.echo(f"   Successfully processed: {successful}/{len(results)}")
        if failed > 0:
            click.echo(f"   Failed: {failed}")
        
        if output_dir:
            click.echo(f"   Output directory: {output_dir.absolute()}")


@cli.command()
@click.argument('files', nargs=-1, required=True, callback=validate_excel_file)
@click.option('--output-dir', '-o', type=click.Path(file_okay=False, dir_okay=True, path_type=Path), 
              default=Path("analysis_reports"), help="Directory to save output files")
@click.option('--json', '-j', is_flag=True, help="Generate JSON reports")
@click.option('--markdown', '-m', is_flag=True, help="Generate markdown reports")
@click.option('--dataframes', '-d', is_flag=True, help="Extract DataFrames (parser only)")
@click.option('--save-dfs', is_flag=True, help="Save DataFrames to files")
@click.option('--dfs-format', type=click.Choice(['csv', 'excel', 'parquet']), 
              default='csv', help="Format for saving DataFrames")
@click.option('--llm-optimized', is_flag=True, help="LLM-optimized output (extractor only)")
@click.option('--summary', is_flag=True, help="Show summary statistics")
@click.option('--verbose', '-v', is_flag=True, help="Enable verbose output")
@click.option('--quiet', '-q', is_flag=True, help="Suppress non-essential output")
@click.option('--timing', is_flag=True, help="Show timing information")
def analyze(files, output_dir, json, markdown, dataframes, save_dfs, dfs_format, llm_optimized, summary, verbose, quiet, timing):
    """
    Analyze Excel files using both parser and extractor.
    
    This command runs both analysis tools and provides comprehensive results.
    Combines the speed of the parser with the depth of the extractor.
    """
    if not files:
        click.echo("❌ No valid Excel files found to process.")
        sys.exit(1)
    
    click.echo("🔍 Running comprehensive analysis with both parser and extractor...")
    
    # Create separate output directories
    parser_dir = output_dir / "parser"
    extractor_dir = output_dir / "extractor"
    
    # Run parser analysis
    click.echo("\n📊 Running Parser Analysis...")
    parser_results = process_files_with_parser(
        files, parser_dir, json, markdown, dataframes, save_dfs, dfs_format, summary, verbose, quiet
    )
    
    # Run extractor analysis
    click.echo("\n📄 Running Extractor Analysis...")
    extractor_results = process_files_with_extractor(
        files, extractor_dir, json, markdown, llm_optimized, summary, verbose, quiet, timing
    )
    
    # Combined summary
    if len(files) > 1:
        parser_successful = sum(1 for r in parser_results if r["success"])
        extractor_successful = sum(1 for r in extractor_results if r["success"])
        
        click.echo(f"\n✅ Comprehensive analysis complete!")
        click.echo(f"   Parser: {parser_successful}/{len(files)} successful")
        click.echo(f"   Extractor: {extractor_successful}/{len(files)} successful")
        click.echo(f"   Output directory: {output_dir.absolute()}")


def process_files_with_error_sniffer(files, output_dir, json_output, markdown_output, summary, verbose, quiet, timing):
    """Process files using the Excel Error Sniffer."""
    results = []
    total_files = len(files)
    
    if verbose and total_files > 1:
        click.echo(f"🔍 Processing {total_files} file(s) with Excel Error Sniffer...")
    
    for i, file_path in enumerate(files, 1):
        if verbose and total_files > 1:
            click.echo(f"\n[{i}/{total_files}] Processing: {file_path.name}")
        
        start_time = time.time()
        
        try:
            # Create output directory
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Import here to avoid circular imports
            from .excel_error_sniffer import ExcelErrorSniffer
            
            # Initialize error sniffer
            sniffer = ExcelErrorSniffer(file_path)
            
            # Sniff for errors
            error_data = sniffer.sniff_errors()
            
            # Generate JSON report
            if json_output:
                json_file = output_dir / f"{file_path.stem}_errors.json"
                sniffer.save_json(json_file)
                if not quiet:
                    click.echo(f"📄 Error JSON report saved to: {json_file}")
            
            # Generate markdown report
            if markdown_output:
                markdown_file = output_dir / f"{file_path.stem}_errors.md"
                sniffer.save_markdown(markdown_file)
                if not quiet:
                    click.echo(f"📝 Error markdown report saved to: {markdown_file}")
            
            # Show summary
            if summary:
                total_errors = sum(len(errors) for errors in error_data.values() if isinstance(errors, list))
                click.echo(f"🔍 Found {total_errors} potential issues in {file_path.name}")
                
                if verbose:
                    for error_type, errors in error_data.items():
                        if isinstance(errors, list) and errors:
                            click.echo(f"   {error_type}: {len(errors)} issues")
            
            # Timing information
            if timing:
                elapsed_time = time.time() - start_time
                click.echo(f"⏱️  Error detection completed in {elapsed_time:.2f} seconds")
            
            results.append({
                "file": file_path,
                "success": True,
                "errors": error_data,
                "elapsed_time": time.time() - start_time
            })
            
        except Exception as e:
            if not quiet:
                click.echo(f"❌ Error processing {file_path.name}: {e}")
            results.append({
                "file": file_path,
                "success": False,
                "error": str(e),
                "elapsed_time": time.time() - start_time
            })
    
    return results


def process_files_with_probabilistic_detector(files, output_dir, json_output, markdown_output, error_threshold, detectors, summary, verbose, quiet, timing):
    """Process files using the Probabilistic Error Detector."""
    results = []
    total_files = len(files)
    
    if verbose and total_files > 1:
        click.echo(f"🎯 Processing {total_files} file(s) with Probabilistic Error Detector...")
    
    for i, file_path in enumerate(files, 1):
        if verbose and total_files > 1:
            click.echo(f"\n[{i}/{total_files}] Processing: {file_path.name}")
        
        start_time = time.time()
        
        try:
            # Create output directory
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Import here to avoid circular imports
            from .probabilistic_error_detector import detect_excel_errors_probabilistic
            
            # Detect errors probabilistically
            error_data = detect_excel_errors_probabilistic(
                file_path, 
                error_threshold=error_threshold,
                output_dir=output_dir if json_output or markdown_output else None
            )
            
            # Show summary
            if summary:
                total_errors = len(error_data.get('errors', []))
                click.echo(f"🎯 Found {total_errors} potential errors (threshold: {error_threshold}) in {file_path.name}")
                
                if verbose and 'errors' in error_data:
                    error_counts = {}
                    for error in error_data['errors']:
                        error_type = error.get('error_type', 'unknown')
                        error_counts[error_type] = error_counts.get(error_type, 0) + 1
                    
                    for error_type, count in error_counts.items():
                        click.echo(f"   {error_type}: {count} errors")
            
            # Timing information
            if timing:
                elapsed_time = time.time() - start_time
                click.echo(f"⏱️  Probabilistic detection completed in {elapsed_time:.2f} seconds")
            
            results.append({
                "file": file_path,
                "success": True,
                "errors": error_data,
                "elapsed_time": time.time() - start_time
            })
            
        except Exception as e:
            if not quiet:
                click.echo(f"❌ Error processing {file_path.name}: {e}")
            results.append({
                "file": file_path,
                "success": False,
                "error": str(e),
                "elapsed_time": time.time() - start_time
            })
    
    return results


@cli.command()
@click.argument('files', nargs=-1, required=True, callback=validate_excel_file)
@click.option('--output-dir', '-o', type=click.Path(file_okay=False, dir_okay=True, path_type=Path), 
              default=Path("error_reports"), help="Directory to save output files")
@click.option('--json', '-j', is_flag=True, help="Generate JSON error reports")
@click.option('--markdown', '-m', is_flag=True, help="Generate markdown error reports")
@click.option('--summary', is_flag=True, help="Show error summary statistics")
@click.option('--verbose', '-v', is_flag=True, help="Enable verbose output")
@click.option('--quiet', '-q', is_flag=True, help="Suppress non-essential output")
@click.option('--timing', is_flag=True, help="Show detailed timing information")
def error_sniff(files, output_dir, json, markdown, summary, verbose, quiet, timing):
    """
    Detect common Excel errors and issues.
    
    This command uses the Excel Error Sniffer to detect various types of errors:
    - Formula errors (#N/A, #VALUE!, #REF!, etc.)
    - Circular references
    - Broken links and references
    - Data validation issues
    - Performance problems
    - Structural issues
    - Compatibility warnings
    """
    if not files:
        click.echo("❌ No valid Excel files found to process.")
        sys.exit(1)
    
    results = process_files_with_error_sniffer(
        files, output_dir, json, markdown, summary, verbose, quiet, timing
    )
    
    # Summary
    if len(files) > 1:
        successful = sum(1 for r in results if r["success"])
        failed = len(results) - successful
        
        click.echo(f"\n✅ Error detection complete!")
        click.echo(f"   Successfully processed: {successful}/{len(results)}")
        if failed > 0:
            click.echo(f"   Failed: {failed}")
        
        if output_dir:
            click.echo(f"   Output directory: {output_dir.absolute()}")


@cli.command()
@click.argument('files', nargs=-1, required=True, callback=validate_excel_file)
@click.option('--output-dir', '-o', type=click.Path(file_okay=False, dir_okay=True, path_type=Path), 
              default=Path("probabilistic_error_reports"), help="Directory to save output files")
@click.option('--json', '-j', is_flag=True, help="Generate JSON error reports")
@click.option('--markdown', '-m', is_flag=True, help="Generate markdown error reports")
@click.option('--error-threshold', '-t', type=float, default=0.7, 
              help="Minimum probability threshold for reporting errors (0.0 to 1.0)")
@click.option('--summary', is_flag=True, help="Show error summary statistics")
@click.option('--verbose', '-v', is_flag=True, help="Enable verbose output")
@click.option('--quiet', '-q', is_flag=True, help="Suppress non-essential output")
@click.option('--timing', is_flag=True, help="Show detailed timing information")
def detect_errors(files, output_dir, json, markdown, error_threshold, summary, verbose, quiet, timing):
    """
    Detect Excel errors using advanced probabilistic models.
    
    This command uses the Probabilistic Error Detector to find complex errors:
    - Circular named ranges
    - Inconsistent date formats
    - Array formula spill errors
    - Volatile function usage
    - Cross-sheet reference errors
    - Data type inconsistencies
    - Conditional formatting conflicts
    - External data connection failures
    - Precision errors in financial calculations
    - Incomplete drag formulas
    - False range end detection
    - Partial formula propagation
    - Formula boundary mismatches
    - Copy-paste formula gaps
    - Formula range vs data range discrepancies
    - Inconsistent formula application
    - Missing dollar sign anchors
    - Wrong row/column anchoring
    - Over-anchored references
    - Inconsistent anchoring in ranges
    - Lookup function anchoring issues
    - Array formula anchoring issues
    - Cross-sheet anchoring issues
    """
    if not files:
        click.echo("❌ No valid Excel files found to process.")
        sys.exit(1)
    
    # Validate error threshold
    if not 0.0 <= error_threshold <= 1.0:
        click.echo("❌ Error threshold must be between 0.0 and 1.0")
        sys.exit(1)
    
    if verbose:
        click.echo(f"🎯 Using error threshold: {error_threshold}")
    
    results = process_files_with_probabilistic_detector(
        files, output_dir, json, markdown, error_threshold, None, summary, verbose, quiet, timing
    )
    
    # Summary
    if len(files) > 1:
        successful = sum(1 for r in results if r["success"])
        failed = len(results) - successful
        
        click.echo(f"\n✅ Probabilistic error detection complete!")
        click.echo(f"   Successfully processed: {successful}/{len(results)}")
        if failed > 0:
            click.echo(f"   Failed: {failed}")
        
        if output_dir:
            click.echo(f"   Output directory: {output_dir.absolute()}")


if __name__ == '__main__':
    cli() 