#!/usr/bin/env python3
"""
Probabilistic Excel Error Detector - Advanced error detection using pattern analysis and probability models.

This module provides a framework for detecting complex Excel errors using probabilistic models
and pattern analysis. Each error type has its own detection algorithm that returns a probability
score, allowing users to set thresholds to control false positives.
"""

import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple, Callable
import warnings
from datetime import datetime
import json
import re
from dataclasses import dataclass
from enum import Enum
from collections import Counter

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.external_reference import ExternalReference

logger = logging.getLogger(__name__)


class ErrorSeverity(Enum):
    """Error severity levels."""
    HIGH = "high"
    MEDIUM = "medium"
    LOW = "low"


@dataclass
class ErrorDetectionResult:
    """Result of an error detection algorithm."""
    error_type: str
    description: str
    probability: float  # 0.0 to 1.0
    severity: ErrorSeverity
    location: Optional[str] = None  # e.g., "Sheet1!A1", "NamedRange:Revenue"
    details: Optional[Dict[str, Any]] = None
    suggested_fix: Optional[str] = None


class ErrorDetector:
    """Base class for error detection algorithms."""
    
    def __init__(self, name: str, description: str, severity: ErrorSeverity):
        self.name = name
        self.description = description
        self.severity = severity
    
    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        """
        Detect errors of this type in the workbook.
        
        Args:
            workbook: The Excel workbook to analyze
            **kwargs: Additional parameters for detection
            
        Returns:
            List of detected errors with probability scores
        """
        raise NotImplementedError("Subclasses must implement detect()")


class ProbabilisticErrorSniffer:
    """
    Advanced Excel error detection using probabilistic models and pattern analysis.
    
    This framework allows for:
    - Probabilistic error detection with confidence scores
    - Configurable thresholds to control false positives
    - Extensible architecture for adding new error types
    - Pattern-based analysis for complex error detection
    """
    
    def __init__(self, file_path: Path, error_threshold: float = 0.7):
        """
        Initialize the Probabilistic Error Sniffer.
        
        Args:
            file_path: Path to the Excel file to analyze
            error_threshold: Minimum probability threshold for reporting errors (0.0 to 1.0)
        """
        self.file_path = Path(file_path)
        self.workbook = None
        self.error_threshold = error_threshold
        self.detectors: List[ErrorDetector] = []
        self.detection_results: Dict[str, List[ErrorDetectionResult]] = {}
        
        # Register built-in error detectors
        self._register_builtin_detectors()
    
    def _load_workbook(self) -> None:
        """Load the Excel workbook safely."""
        try:
            self.workbook = openpyxl.load_workbook(
                self.file_path, 
                data_only=False,  # Keep formulas for error detection
                keep_vba=True
            )
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            raise
    
    def register_detector(self, detector: ErrorDetector) -> None:
        """Register a new error detector."""
        self.detectors.append(detector)
        logger.info(f"Registered error detector: {detector.name}")
    
    def _register_builtin_detectors(self) -> None:
        """Register all built-in error detectors."""
        # We'll add these one by one as we implement them
        pass
    
    def detect_all_errors(self) -> Dict[str, Any]:
        """
        Run all registered error detectors and return results above threshold.
        
        Returns:
            Dictionary containing all detected errors organized by type
        """
        logger.info(f"Starting probabilistic error detection for: {self.file_path}")
        
        try:
            self._load_workbook()
            
            # Run all detectors
            for detector in self.detectors:
                try:
                    results = detector.detect(self.workbook)
                    # Filter results by threshold
                    filtered_results = [
                        result for result in results 
                        if result.probability >= self.error_threshold
                    ]
                    self.detection_results[detector.name] = filtered_results
                    
                    logger.info(f"Detector '{detector.name}' found {len(filtered_results)} errors above threshold")
                    
                except Exception as e:
                    logger.error(f"Error in detector '{detector.name}': {e}")
                    self.detection_results[detector.name] = []
            
            # Generate summary
            summary = self._generate_summary()
            self.detection_results['summary'] = summary
            
            logger.info(f"Error detection completed. Found {summary['total_errors']} errors above threshold.")
            
        except Exception as e:
            logger.error(f"Error during detection: {e}")
            raise
        finally:
            if self.workbook:
                self.workbook.close()
        
        return self.detection_results
    
    def _generate_summary(self) -> Dict[str, Any]:
        """Generate summary of detection results."""
        total_errors = sum(len(results) for key, results in self.detection_results.items() if key != 'summary')
        
        severity_counts = {severity.value: 0 for severity in ErrorSeverity}
        for results in self.detection_results.values():
            if isinstance(results, list):
                for result in results:
                    severity_counts[result.severity.value] += 1
        
        return {
            'total_errors': total_errors,
            'severity_breakdown': severity_counts,
            'error_types': {
                detector.name: len(self.detection_results.get(detector.name, []))
                for detector in self.detectors
            },
            'threshold_used': self.error_threshold,
            'timestamp': datetime.now().isoformat(),
            'file_path': str(self.file_path),
            'file_size_mb': round(self.file_path.stat().st_size / (1024 * 1024), 2)
        }


# ============================================================================
# ERROR DETECTION ALGORITHMS
# ============================================================================

class HiddenDataInRangesDetector(ErrorDetector):
    """
    Detector for hidden rows/columns in data ranges.
    
    Algorithm:
    1. Identify data ranges (continuous cells with data)
    2. Check if ranges include hidden rows/columns
    3. Analyze data consistency within ranges
    4. Calculate probability based on hidden data patterns
    """
    
    def __init__(self):
        super().__init__(
            name="hidden_data_in_ranges",
            description="Hidden rows/columns in data ranges that may contain incorrect data",
            severity=ErrorSeverity.HIGH
        )
    
    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        """Detect hidden data in ranges."""
        results = []
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            
            # Find data ranges
            data_ranges = self._find_data_ranges(sheet)
            
            for range_info in data_ranges:
                # Check for hidden rows/columns in this range
                hidden_analysis = self._analyze_hidden_data(sheet, range_info)
                
                if hidden_analysis['has_hidden_data']:
                    probability = self._calculate_hidden_data_probability(hidden_analysis)
                    
                    if probability > 0:
                        results.append(ErrorDetectionResult(
                            error_type=self.name,
                            description=f"Data range {range_info['range']} contains hidden {hidden_analysis['hidden_type']} with potentially inconsistent data",
                            probability=probability,
                            severity=self.severity,
                            location=f"{sheet_name}!{range_info['range']}",
                            details=hidden_analysis,
                            suggested_fix="Review hidden data in range and ensure consistency with visible data"
                        ))
        
        return results
    
    def _find_data_ranges(self, sheet) -> List[Dict[str, Any]]:
        """Find continuous data ranges in the sheet."""
        ranges = []
        
        # Get the used range
        min_row = sheet.min_row
        max_row = sheet.max_row
        min_col = sheet.min_column
        max_col = sheet.max_column
        
        if min_row is None or max_row is None:
            return ranges
        
        # Find continuous data blocks
        current_range_start = None
        current_range_end = None
        
        for row in range(min_row, max_row + 1):
            row_has_data = any(
                sheet.cell(row=row, column=col).value is not None
                for col in range(min_col, max_col + 1)
            )
            
            if row_has_data and current_range_start is None:
                current_range_start = row
            elif not row_has_data and current_range_start is not None:
                current_range_end = row - 1
                ranges.append({
                    'start_row': current_range_start,
                    'end_row': current_range_end,
                    'range': f"{get_column_letter(min_col)}{current_range_start}:{get_column_letter(max_col)}{current_range_end}"
                })
                current_range_start = None
        
        # Handle range that extends to the end
        if current_range_start is not None:
            ranges.append({
                'start_row': current_range_start,
                'end_row': max_row,
                'range': f"{get_column_letter(min_col)}{current_range_start}:{get_column_letter(max_col)}{max_row}"
            })
        
        return ranges
    
    def _analyze_hidden_data(self, sheet, range_info: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze hidden data within a range."""
        hidden_rows = []
        hidden_cols = []
        
        # Check for hidden rows
        for row in range(range_info['start_row'], range_info['end_row'] + 1):
            if sheet.row_dimensions[row].hidden:
                hidden_rows.append(row)
        
        # Check for hidden columns
        for col in range(sheet.min_column, sheet.max_column + 1):
            if sheet.column_dimensions[get_column_letter(col)].hidden:
                hidden_cols.append(col)
        
        # Analyze data consistency
        visible_data = []
        hidden_data = []
        
        for row in range(range_info['start_row'], range_info['end_row'] + 1):
            for col in range(sheet.min_column, sheet.max_column + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value is not None:
                    if row in hidden_rows or col in hidden_cols:
                        hidden_data.append(cell.value)
                    else:
                        visible_data.append(cell.value)
        
        return {
            'has_hidden_data': len(hidden_rows) > 0 or len(hidden_cols) > 0,
            'hidden_rows': hidden_rows,
            'hidden_cols': hidden_cols,
            'hidden_type': 'rows' if len(hidden_rows) > 0 else 'columns' if len(hidden_cols) > 0 else 'none',
            'visible_data_count': len(visible_data),
            'hidden_data_count': len(hidden_data),
            'data_consistency_score': self._calculate_data_consistency(visible_data, hidden_data)
        }
    
    def _calculate_data_consistency(self, visible_data: List, hidden_data: List) -> float:
        """Calculate consistency score between visible and hidden data."""
        if not visible_data or not hidden_data:
            return 1.0  # No inconsistency if one set is empty
        
        # Simple consistency check - can be enhanced with more sophisticated analysis
        visible_types = set(type(val) for val in visible_data)
        hidden_types = set(type(val) for val in hidden_data)
        
        type_overlap = len(visible_types.intersection(hidden_types))
        total_types = len(visible_types.union(hidden_types))
        
        return type_overlap / total_types if total_types > 0 else 1.0
    
    def _calculate_hidden_data_probability(self, analysis: Dict[str, Any]) -> float:
        """Calculate probability of hidden data causing issues."""
        if not analysis['has_hidden_data']:
            return 0.0
        
        # Base probability from having hidden data
        base_prob = 0.3
        
        # Adjust based on data consistency
        consistency_factor = 1.0 - analysis['data_consistency_score']
        base_prob += consistency_factor * 0.4
        
        # Adjust based on amount of hidden data
        total_data = analysis['visible_data_count'] + analysis['hidden_data_count']
        if total_data > 0:
            hidden_ratio = analysis['hidden_data_count'] / total_data
            base_prob += hidden_ratio * 0.3
        
        return min(base_prob, 1.0)


class CircularNamedRangesDetector(ErrorDetector):
    """
    Detector for circular references in named ranges.
    
    Algorithm:
    1. Extract all named ranges from the workbook
    2. Parse each named range's formula to identify dependencies
    3. Build a dependency graph
    4. Detect cycles using graph algorithms
    5. Calculate probability based on cycle characteristics
    """
    
    def __init__(self):
        super().__init__(
            name="circular_named_ranges",
            description="Circular references in named ranges that can cause infinite calculation loops",
            severity=ErrorSeverity.HIGH
        )
    
    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        """Detect circular references in named ranges."""
        results = []
        
        # Extract all named ranges
        named_ranges = self._extract_named_ranges(workbook)
        
        if not named_ranges:
            return results
        
        # Build dependency graph
        dependency_graph = self._build_dependency_graph(named_ranges)
        
        # Detect cycles
        cycles = self._detect_cycles(dependency_graph)
        
        # Generate results for each cycle
        for cycle in cycles:
            probability = self._calculate_circular_probability(cycle, named_ranges, dependency_graph)
            
            if probability > 0:
                # Get formulas for the cycle
                cycle_formulas = {name: named_ranges[name]['formula'] for name in cycle}
                
                results.append(ErrorDetectionResult(
                    error_type=self.name,
                    description=f"Circular reference detected: {' → '.join(cycle)} → {cycle[0]}",
                    probability=probability,
                    severity=self.severity,
                    location=f"NamedRanges: {', '.join(cycle[:-1])}",  # Exclude the duplicate at the end
                    details={
                        'cycle': cycle,
                        'cycle_length': len(cycle),
                        'formulas': cycle_formulas,
                        'dependency_graph': dependency_graph
                    },
                    suggested_fix="Break the circular dependency by introducing intermediate calculations or using iterative calculation settings"
                ))
        
        return results
    
    def _extract_named_ranges(self, workbook: openpyxl.Workbook) -> Dict[str, Dict[str, Any]]:
        """Extract all named ranges and their formulas."""
        named_ranges = {}
        
        for name in workbook.defined_names:
            try:
                # Get the DefinedName object
                defined_name = workbook.defined_names[name]
                
                # Get the formula/reference
                formula = defined_name.attr_text if hasattr(defined_name, 'attr_text') else str(defined_name)
                
                named_ranges[name] = {
                    'formula': formula,
                    'scope': defined_name.localSheetId if hasattr(defined_name, 'localSheetId') else None,
                    'comment': defined_name.comment if hasattr(defined_name, 'comment') else None
                }
            except Exception as e:
                logger.warning(f"Could not extract named range {name}: {e}")
                continue
        
        return named_ranges
    
    def _parse_named_range_formula(self, formula: str) -> List[str]:
        """Parse a named range formula to extract dependencies."""
        dependencies = []
        
        if not formula or not isinstance(formula, str):
            return dependencies
        
        # Remove leading '=' if present
        if formula.startswith('='):
            formula = formula[1:]
        
        # Extract named range references using regex
        import re
        
        # Pattern for named range references
        # Matches: standalone names, names in functions, names in operations
        patterns = [
            r'\b([A-Za-z_][A-Za-z0-9_]*)\b',  # Basic named range pattern
            r'([A-Za-z_][A-Za-z0-9_]*)',      # More permissive pattern
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, formula)
            for match in matches:
                # Filter out common Excel functions and keywords (case-sensitive)
                excel_functions = {
                    'SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'IF', 'AND', 'OR',
                    'TRUE', 'FALSE', 'PI', 'TODAY', 'NOW', 'ROW', 'COLUMN',
                    'ABS', 'ROUND', 'INT', 'MOD', 'POWER', 'SQRT', 'LOG', 'LN',
                    'SIN', 'COS', 'TAN', 'ASIN', 'ACOS', 'ATAN', 'RAND', 'RANDBETWEEN'
                }
                
                # Filter out cell references (like A1, B2, etc.)
                is_cell_reference = (
                    len(match) >= 2 and 
                    match[0].isalpha() and 
                    match[1:].isdigit()
                )
                
                if (match not in excel_functions and 
                    len(match) >= 1 and 
                    not is_cell_reference):
                    dependencies.append(match)
        
        return list(set(dependencies))  # Remove duplicates
    
    def _build_dependency_graph(self, named_ranges: Dict[str, Dict[str, Any]]) -> Dict[str, List[str]]:
        """Build a dependency graph from named ranges."""
        graph = {}
        
        for name, info in named_ranges.items():
            dependencies = self._parse_named_range_formula(info['formula'])
            # Only include dependencies that are actual named ranges
            valid_dependencies = [dep for dep in dependencies if dep in named_ranges]
            graph[name] = valid_dependencies
        
        return graph
    
    def _detect_cycles(self, graph: Dict[str, List[str]]) -> List[List[str]]:
        """Detect cycles in the dependency graph using DFS."""
        cycles = []
        
        def dfs(node: str, path: List[str], visited: set, rec_stack: set):
            """Depth-first search to detect cycles."""
            if node in rec_stack:
                # Found a cycle
                cycle_start = path.index(node)
                cycle = path[cycle_start:] + [node]
                # Only add if it's a real cycle (length > 2, meaning at least 2 different nodes)
                if len(set(cycle[:-1])) > 1:  # More than one unique node (excluding the duplicate at the end)
                    cycles.append(cycle)
                return
            
            if node in visited:
                return
            
            visited.add(node)
            rec_stack.add(node)
            path.append(node)
            
            for neighbor in graph.get(node, []):
                dfs(neighbor, path.copy(), visited.copy(), rec_stack.copy())
            
            rec_stack.remove(node)
        
        # Run DFS from each node
        for node in graph:
            visited = set()
            rec_stack = set()
            dfs(node, [], visited, rec_stack)
        
        # Remove duplicate cycles (same cycle starting from different nodes)
        unique_cycles = []
        for cycle in cycles:
            # Normalize cycle by starting from the lexicographically smallest node
            cycle_nodes = cycle[:-1]  # Exclude the last node (duplicate of first)
            min_node = min(cycle_nodes)
            start_idx = cycle_nodes.index(min_node)
            normalized_cycle = cycle_nodes[start_idx:] + cycle_nodes[:start_idx] + [min_node]
            
            # Check if this normalized cycle is already in unique_cycles
            is_duplicate = False
            for existing_cycle in unique_cycles:
                existing_nodes = existing_cycle[:-1]
                if set(existing_nodes) == set(cycle_nodes):
                    is_duplicate = True
                    break
            
            if not is_duplicate:
                unique_cycles.append(normalized_cycle)
        
        return unique_cycles
    
    def _calculate_circular_probability(self, cycle: List[str], named_ranges: Dict[str, Dict[str, Any]], graph: Dict[str, List[str]]) -> float:
        """Calculate probability of circular reference being problematic."""
        if not cycle:
            return 0.0
        
        # Base probability based on cycle length
        cycle_length = len(cycle)
        if cycle_length == 2:
            base_prob = 0.9  # Very high probability for 2-range cycles
        elif cycle_length == 3:
            base_prob = 0.8  # High probability for 3-range cycles
        elif cycle_length == 4:
            base_prob = 0.7  # Medium-high probability for 4-range cycles
        else:
            base_prob = 0.6  # Medium probability for longer cycles
        
        # Adjust based on formula complexity
        complexity_factor = 0.0
        for name in cycle:
            formula = named_ranges[name]['formula']
            # Count functions, operators, and references
            function_count = formula.count('(') + formula.count(')')
            operator_count = formula.count('+') + formula.count('-') + formula.count('*') + formula.count('/')
            reference_count = len(self._parse_named_range_formula(formula))
            
            complexity = (function_count + operator_count + reference_count) / 10.0
            complexity_factor = max(complexity_factor, complexity)
        
        base_prob += complexity_factor * 0.2
        
        # Adjust based on usage frequency (how many other named ranges depend on these)
        usage_factor = 0.0
        for name in cycle:
            # Count how many other named ranges depend on this one
            dependents = sum(1 for deps in graph.values() if name in deps)
            usage_factor = max(usage_factor, dependents / 10.0)  # Normalize
        
        base_prob += usage_factor * 0.1
        
        # Check for aggregation functions in cycle (more dangerous)
        aggregation_functions = ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'SUMPRODUCT']
        for name in cycle:
            formula = named_ranges[name]['formula'].upper()
            if any(func in formula for func in aggregation_functions):
                base_prob += 0.1
                break
        
        return min(base_prob, 1.0)


class InconsistentDateFormatsDetector(ErrorDetector):
    """
    Detector for inconsistent date formats in date calculations.
    
    Algorithm:
    1. Scan for formulas that perform date arithmetic or use date functions
    2. Identify ranges/columns involved in date calculations
    3. Analyze data types in those ranges (Excel date, text that looks like date, other)
    4. Flag if a range contains a mix of true dates and text-formatted dates
    5. Calculate probability based on proportion of inconsistencies
    """
    def __init__(self):
        super().__init__(
            name="inconsistent_date_formats",
            description="Mixed date formats (text vs. actual dates) in date-based calculations",
            severity=ErrorSeverity.HIGH
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            checked_ranges = set()
            # 1. Scan all columns for mixed date types or all text dates
            for col in range(1, sheet.max_column + 1):
                col_letter = openpyxl.utils.get_column_letter(col)
                if col_letter in checked_ranges:
                    continue
                checked_ranges.add(col_letter)
                analysis = self._analyze_range(sheet, col_letter)
                # Flag if mixed types OR all text dates (and at least 1 text date)
                is_all_text_dates = analysis['text_date_count'] > 0 and analysis['date_count'] == 0
                if analysis['mixed_types'] or is_all_text_dates:
                    probability = self._calculate_probability(analysis, is_all_text_dates=is_all_text_dates)
                    if probability > 0:
                        results.append(ErrorDetectionResult(
                            error_type=self.name,
                            description=f"{'All text dates' if is_all_text_dates else 'Mixed date formats'} detected in column {col_letter} on sheet {sheet_name}",
                            probability=probability,
                            severity=self.severity,
                            location=f"{sheet_name}!{col_letter}",
                            details=analysis,
                            suggested_fix="Convert all dates in the column to Excel date format for consistency"
                        ))
            # 2. Also scan columns referenced by date formulas (legacy logic)
            date_formula_cells = self._find_date_formula_cells(sheet)
            for cell in date_formula_cells:
                referenced_ranges = self._extract_referenced_ranges(cell)
                for rng in referenced_ranges:
                    if rng in checked_ranges:
                        continue
                    checked_ranges.add(rng)
                    analysis = self._analyze_range(sheet, rng)
                    is_all_text_dates = analysis['text_date_count'] > 0 and analysis['date_count'] == 0
                    if analysis['mixed_types'] or is_all_text_dates:
                        probability = self._calculate_probability(analysis, is_all_text_dates=is_all_text_dates)
                        if probability > 0:
                            results.append(ErrorDetectionResult(
                                error_type=self.name,
                                description=f"{'All text dates' if is_all_text_dates else 'Mixed date formats'} detected in range {rng} on sheet {sheet_name}",
                                probability=probability,
                                severity=self.severity,
                                location=f"{sheet_name}!{rng}",
                                details=analysis,
                                suggested_fix="Convert all dates in the range to Excel date format for consistency"
                            ))
        return results

    def _find_date_formula_cells(self, sheet) -> List[openpyxl.cell.cell.Cell]:
        """Find cells with formulas that perform date arithmetic or use date functions."""
        date_functions = {'DATEDIF', 'DATE', 'YEAR', 'MONTH', 'DAY', 'EDATE', 'EOMONTH', 'TODAY', 'NOW', 'NETWORKDAYS', 'WORKDAY'}
        date_formula_cells = []
        for row in sheet.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and cell.value:
                    formula = str(cell.value).upper()
                    # Check for date functions or date arithmetic (e.g., +, -, with cell refs)
                    if any(func in formula for func in date_functions) or self._is_date_arithmetic(formula):
                        date_formula_cells.append(cell)
        return date_formula_cells

    def _is_date_arithmetic(self, formula: str) -> bool:
        # Simple heuristic: look for + or - between cell references
        import re
        # e.g., =A1-B1 or =B2+30
        return bool(re.search(r"[A-Z]+[0-9]+\s*[-+]\s*[A-Z0-9]+", formula))

    def _extract_referenced_ranges(self, cell) -> List[str]:
        # For simplicity, just extract all cell references in the formula
        import re
        formula = str(cell.value)
        # Matches A1, B2, C10, etc.
        refs = re.findall(r"[A-Z]+[0-9]+", formula)
        # Group by column (e.g., all A1, A2, A3 -> A)
        columns = set(ref[0] for ref in refs if len(ref) >= 2)
        # For now, treat each column as a range (e.g., A)
        return [f"{col}" for col in columns]

    def _analyze_range(self, sheet, col: str) -> dict:
        # Analyze all cells in the given column
        from openpyxl.utils import column_index_from_string
        col_idx = column_index_from_string(col)
        date_count = 0
        text_date_count = 0
        other_count = 0
        total = 0
        text_date_examples = []
        for row in sheet.iter_rows(min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value is None:
                    continue
                total += 1
                if self._is_excel_date(cell):
                    date_count += 1
                elif self._looks_like_date(cell.value):
                    text_date_count += 1
                    if len(text_date_examples) < 3:
                        text_date_examples.append(cell.value)
                else:
                    other_count += 1
        mixed_types = date_count > 0 and text_date_count > 0
        return {
            'date_count': date_count,
            'text_date_count': text_date_count,
            'other_count': other_count,
            'total': total,
            'mixed_types': mixed_types,
            'text_date_examples': text_date_examples
        }

    def _is_excel_date(self, cell) -> bool:
        # openpyxl stores Excel dates as numbers with a date format
        if cell.is_date:
            return True
        # Sometimes dates are stored as numbers with a date format
        if cell.data_type == 'n' and cell.number_format and 'yy' in cell.number_format.lower():
            return True
        return False

    def _looks_like_date(self, value) -> bool:
        import re
        # Match common date patterns: YYYY-MM-DD, DD/MM/YYYY, MM/DD/YYYY, etc.
        if not isinstance(value, str):
            return False
        patterns = [
            r'\b\d{4}-\d{2}-\d{2}\b',  # 2023-01-01
            r'\b\d{2}/\d{2}/\d{4}\b',  # 01/01/2023
            r'\b\d{1,2} [A-Za-z]{3,9} \d{4}\b',  # 1 Jan 2023
            r'\b\d{2}\.\d{2}\.\d{4}\b',  # 01.01.2023
        ]
        return any(re.search(pat, value) for pat in patterns)

    def _calculate_probability(self, analysis: dict, is_all_text_dates: bool = False) -> float:
        if analysis['total'] == 0:
            return 0.0
        if is_all_text_dates:
            # All text dates, no Excel dates
            ratio = analysis['text_date_count'] / analysis['total']
            if ratio > 0.9:
                return 0.5  # Lower probability, but still a warning
            elif ratio > 0.5:
                return 0.2
            else:
                return 0.1
        ratio = analysis['text_date_count'] / analysis['total']
        if ratio > 0.1:
            return 0.9
        elif ratio > 0.01:
            return 0.5
        elif ratio > 0:
            return 0.2
        return 0.0


class ArrayFormulaSpillErrorsDetector(ErrorDetector):
    """
    Detector for array formula spill errors.
    
    Algorithm:
    1. Scan for array formulas (curly braces or dynamic array functions)
    2. Look for #SPILL! errors in cells
    3. For each array formula, check the intended spill range for conflicts
    4. Calculate probability based on evidence of spill errors or conflicts
    """
    def __init__(self):
        super().__init__(
            name="array_formula_spill_errors",
            description="Array formulas that can't spill properly due to insufficient space or conflicts",
            severity=ErrorSeverity.HIGH
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    # 1. Check for #SPILL! error
                    if isinstance(cell.value, str) and cell.value.strip().upper() == '#SPILL!':
                        results.append(ErrorDetectionResult(
                            error_type=self.name,
                            description=f"#SPILL! error detected in cell {cell.coordinate} on sheet {sheet_name}",
                            probability=0.95,
                            severity=self.severity,
                            location=f"{sheet_name}!{cell.coordinate}",
                            details={'cell': cell.coordinate, 'error': '#SPILL!'},
                            suggested_fix="Check for blocked cells, merged cells, or other conflicts in the intended spill range."
                        ))
                    # 2. Check for array formulas (legacy or dynamic)
                    if self._is_array_formula(cell):
                        spill_range, conflict_cells = self._analyze_spill_range(sheet, cell)
                        if conflict_cells:
                            # High probability if most/all spill cells are blocked
                            spill_ratio = len(conflict_cells) / (len(spill_range) - 1)  # -1 to exclude original cell
                            probability = 0.9 if spill_ratio >= 0.8 else 0.7
                            results.append(ErrorDetectionResult(
                                error_type=self.name,
                                description=f"Array formula in {cell.coordinate} cannot spill due to conflicts in range {spill_range[0]}:{spill_range[-1]}",
                                probability=probability,
                                severity=self.severity,
                                location=f"{sheet_name}!{cell.coordinate}",
                                details={
                                    'cell': cell.coordinate,
                                    'spill_range': [c.coordinate for c in spill_range],
                                    'conflict_cells': [c.coordinate for c in conflict_cells],
                                },
                                suggested_fix="Clear or move conflicting cells in the spill range."
                            ))
        return results

    def _is_array_formula(self, cell) -> bool:
        # Legacy array formulas: openpyxl marks with cell.data_type == 'f' and formula starts with {= or ends with }
        if cell.data_type == 'f' and cell.value:
            formula = str(cell.value)
            if formula.startswith('{=') or formula.endswith('}'):  # legacy CSE
                return True
            # Dynamic array functions
            dynamic_funcs = ['SEQUENCE', 'UNIQUE', 'SORT', 'FILTER', 'RANDARRAY', 'TRANSPOSE', 'XMATCH', 'XLOOKUP']
            if any(func in formula.upper() for func in dynamic_funcs):
                return True
        return False

    def _analyze_spill_range(self, sheet, cell):
        # Heuristic: For legacy arrays, assume 1xN or Nx1 block; for dynamic, try to estimate
        # For this implementation, just check the right and down cells for conflicts (up to 10x10 block)
        spill_range = [cell]
        conflict_cells = []
        max_rows, max_cols = 10, 10
        # Check rightwards
        for dc in range(1, max_cols):
            col = cell.column + dc
            if col > sheet.max_column:
                break
            c = sheet.cell(row=cell.row, column=col)
            if c.value is not None:
                conflict_cells.append(c)
            spill_range.append(c)
        # Check downwards
        for dr in range(1, max_rows):
            row = cell.row + dr
            if row > sheet.max_row:
                break
            c = sheet.cell(row=row, column=cell.column)
            if c.value is not None:
                conflict_cells.append(c)
            spill_range.append(c)
        return spill_range, conflict_cells


class VolatileFunctionsDetector(ErrorDetector):
    """
    Detector for volatile functions that can cause performance issues in large models.
    
    Algorithm:
    1. Scan for volatile functions (NOW, TODAY, RAND, OFFSET, INDIRECT, etc.)
    2. Analyze usage patterns and dependency impact
    3. Calculate performance impact based on frequency and context
    4. Flag if volatile functions are used inappropriately or excessively
    """
    def __init__(self):
        super().__init__(
            name="volatile_functions",
            description="Volatile functions causing excessive recalculations and performance issues",
            severity=ErrorSeverity.MEDIUM
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        
        # Collect all volatile functions across the workbook
        volatile_cells = []
        total_formulas = 0
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == 'f' and cell.value:
                        total_formulas += 1
                        volatile_funcs = self._find_volatile_functions(str(cell.value))
                        if volatile_funcs:
                            volatile_cells.append({
                                'cell': cell,
                                'sheet': sheet_name,
                                'functions': volatile_funcs,
                                'dependencies': self._count_dependencies(sheet, cell)
                            })
        
        if not volatile_cells:
            return results
        
        # Analyze overall impact
        total_volatile_funcs = sum(len(cell['functions']) for cell in volatile_cells)
        model_size_factor = min(total_formulas / 100, 1.0)  # Normalize by model size
        
        # Calculate probability based on usage patterns
        probability = self._calculate_volatile_probability(
            total_volatile_funcs, 
            len(volatile_cells), 
            total_formulas,
            volatile_cells
        )
        
        if probability > 0:
            # Group by severity level
            high_impact_cells = [c for c in volatile_cells if c['dependencies'] > 5]
            medium_impact_cells = [c for c in volatile_cells if 2 <= c['dependencies'] <= 5]
            
            results.append(ErrorDetectionResult(
                error_type=self.name,
                description=f"Found {total_volatile_funcs} volatile functions across {len(volatile_cells)} cells",
                probability=probability,
                severity=self.severity,
                location=f"Workbook-wide: {len(volatile_cells)} cells affected",
                details={
                    'total_volatile_functions': total_volatile_funcs,
                    'affected_cells': len(volatile_cells),
                    'total_formulas': total_formulas,
                    'high_impact_cells': len(high_impact_cells),
                    'medium_impact_cells': len(medium_impact_cells),
                    'volatile_cells': [
                        {
                            'cell': f"{cell['sheet']}!{cell['cell'].coordinate}",
                            'functions': cell['functions'],
                            'dependencies': cell['dependencies']
                        }
                        for cell in volatile_cells
                    ]
                },
                suggested_fix="Replace volatile functions with static alternatives where possible, or use iterative calculation settings."
            ))
        
        return results

    def _find_volatile_functions(self, formula: str) -> List[str]:
        """Find volatile functions in a formula."""
        volatile_functions = {
            # Time-based
            'NOW', 'TODAY', 'RAND', 'RANDBETWEEN',
            # Reference-based
            'OFFSET', 'INDIRECT', 'ADDRESS', 'COLUMN', 'ROW',
            # Information functions
            'CELL', 'INFO',
            # Database functions
            'DSUM', 'DCOUNT', 'DAVERAGE', 'DMAX', 'DMIN', 'DSTDEV', 'DVAR',
            # Other volatile functions
            'AREAS', 'COLUMNS', 'ROWS', 'HYPERLINK'
        }
        
        found_funcs = []
        formula_upper = formula.upper()
        
        for func in volatile_functions:
            if func in formula_upper:
                # Check if it's actually a function call (not part of another word)
                pattern = r'\b' + func + r'\s*\('
                import re
                if re.search(pattern, formula_upper):
                    found_funcs.append(func)
        
        return found_funcs

    def _count_dependencies(self, sheet, cell) -> int:
        """Count how many other cells reference this cell."""
        dependencies = 0
        cell_ref = cell.coordinate
        
        for row in sheet.iter_rows():
            for other_cell in row:
                if other_cell.data_type == 'f' and other_cell.value:
                    formula = str(other_cell.value)
                    if cell_ref in formula:
                        dependencies += 1
        
        return dependencies

    def _calculate_volatile_probability(self, total_volatile_funcs: int, affected_cells: int, total_formulas: int, volatile_cells: List[dict]) -> float:
        """Calculate probability of performance issues from volatile functions."""
        if total_formulas == 0:
            return 0.0
        
        # For small models, use absolute counts
        if total_formulas < 10:
            if total_volatile_funcs == 1:
                base_prob = 0.2
            elif total_volatile_funcs <= 3:
                base_prob = 0.4
            else:
                base_prob = 0.6
        else:
            # Base probability from frequency (more nuanced)
            volatile_ratio = total_volatile_funcs / max(total_formulas, 1)
            if volatile_ratio > 0.3:  # More than 30% of formulas are volatile
                base_prob = 0.8
            elif volatile_ratio > 0.15:  # 15-30% volatile
                base_prob = 0.6
            elif volatile_ratio > 0.05:  # 5-15% volatile
                base_prob = 0.4
            else:  # Less than 5% volatile
                base_prob = 0.2
        
        # Adjust based on impact (high dependency cells)
        high_impact_count = sum(1 for cell in volatile_cells if cell['dependencies'] > 5)
        if high_impact_count > 0:
            impact_factor = min(high_impact_count / max(affected_cells, 1), 1.0)
            base_prob += impact_factor * 0.7  # High impact can add up to 70%
        
        # Adjust based on model size (larger models are more sensitive)
        if total_formulas > 100:  # Large model
            base_prob += 0.4
        elif total_formulas > 50:  # Medium model
            base_prob += 0.2
        
        # Ensure high probability for high-impact or large model
        if high_impact_count > 0 or total_formulas > 100:
            base_prob = max(base_prob, 0.8)
        
        return min(base_prob, 1.0)


class CrossSheetReferenceErrorsDetector(ErrorDetector):
    """
    Detector for cross-sheet reference errors (broken, missing, or invalid references).
    
    Algorithm:
    1. Scan all formulas for cross-sheet references
    2. Check if referenced sheet exists
    3. Check if referenced cell/range exists in the target sheet
    4. Look for #REF! errors in formulas or cell values
    5. Calculate probability based on severity
    """
    def __init__(self):
        super().__init__(
            name="cross_sheet_reference_errors",
            description="References to cells in other sheets that have been moved or deleted",
            severity=ErrorSeverity.HIGH
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        sheet_names = set(workbook.sheetnames)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == 'f' and cell.value:
                        formula = str(cell.value)
                        cross_refs = self._extract_cross_sheet_references(formula)
                        for ref in cross_refs:
                            ref_sheet, ref_cell = ref
                            if ref_sheet not in sheet_names:
                                # Missing sheet
                                results.append(ErrorDetectionResult(
                                    error_type=self.name,
                                    description=f"Reference to missing sheet '{ref_sheet}' in formula {cell.coordinate} on sheet {sheet_name}",
                                    probability=0.95,
                                    severity=self.severity,
                                    location=f"{sheet_name}!{cell.coordinate}",
                                    details={'formula': formula, 'missing_sheet': ref_sheet},
                                    suggested_fix="Restore the missing sheet or update the formula to reference an existing sheet."
                                ))
                            else:
                                # Check if cell exists in target sheet
                                target_sheet = workbook[ref_sheet]
                                if not self._cell_exists(target_sheet, ref_cell):
                                    # Check for #REF! in formula
                                    if '#REF!' in formula.upper():
                                        prob = 0.95
                                    else:
                                        prob = 0.7
                                    results.append(ErrorDetectionResult(
                                        error_type=self.name,
                                        description=f"Reference to missing cell/range '{ref_cell}' in sheet '{ref_sheet}' from formula {cell.coordinate} on sheet {sheet_name}",
                                        probability=prob,
                                        severity=self.severity,
                                        location=f"{sheet_name}!{cell.coordinate}",
                                        details={'formula': formula, 'target_sheet': ref_sheet, 'missing_cell': ref_cell},
                                        suggested_fix="Update the formula to reference a valid cell/range in the target sheet."
                                    ))
                                else:
                                    # Check if target cell is empty
                                    tgt_cell = target_sheet[ref_cell]  # openpyxl always returns a cell object
                                    if tgt_cell.value is None or tgt_cell.value == '':
                                        results.append(ErrorDetectionResult(
                                            error_type=self.name,
                                            description=f"Reference to empty cell '{ref_cell}' in sheet '{ref_sheet}' from formula {cell.coordinate} on sheet {sheet_name}",
                                            probability=0.3,
                                            severity=ErrorSeverity.LOW,
                                            location=f"{sheet_name}!{cell.coordinate}",
                                            details={'formula': formula, 'target_sheet': ref_sheet, 'empty_cell': ref_cell},
                                            suggested_fix="Check if the referenced cell should contain data."
                                        ))
                    # Also check for #REF! in cell value (not just formula)
                    if isinstance(cell.value, str) and '#REF!' in cell.value.upper():
                        results.append(ErrorDetectionResult(
                            error_type=self.name,
                            description=f"#REF! error in cell {cell.coordinate} on sheet {sheet_name}",
                            probability=0.95,
                            severity=self.severity,
                            location=f"{sheet_name}!{cell.coordinate}",
                            details={'cell': cell.coordinate, 'value': cell.value},
                            suggested_fix="Update the formula to reference a valid cell or sheet."
                        ))
        return results

    def _extract_cross_sheet_references(self, formula: str) -> List[tuple]:
        # Extract references like 'Sheet2'!A1 or Sheet3!B2
        import re
        pattern = r"(?:'([^']+)'|([A-Za-z0-9_]+))!([A-Za-z]+[0-9]+)"
        matches = re.findall(pattern, formula)
        refs = []
        for match in matches:
            sheet = match[0] if match[0] else match[1]
            cell = match[2]
            refs.append((sheet, cell))
        return refs

    def _cell_exists(self, sheet, cell_ref: str) -> bool:
        from openpyxl.utils import coordinate_to_tuple
        try:
            row, col = coordinate_to_tuple(cell_ref)
            max_row = sheet.max_row
            max_col = sheet.max_column
            return 1 <= row <= max_row and 1 <= col <= max_col
        except Exception:
            return False


class DataTypeInconsistenciesInLookupTablesDetector(ErrorDetector):
    """
    Detector for data type inconsistencies in lookup tables (e.g., numbers stored as text).
    
    Algorithm:
    1. Scan for lookup functions (VLOOKUP, HLOOKUP, XLOOKUP, MATCH, INDEX)
    2. Extract lookup ranges
    3. Analyze data types in lookup key columns/rows
    4. Flag if mixed types (number/text/date)
    5. Calculate probability based on proportion of inconsistencies
    """
    def __init__(self):
        super().__init__(
            name="data_type_inconsistencies_in_lookup_tables",
            description="Mixed data types in lookup tables (numbers stored as text, etc.)",
            severity=ErrorSeverity.HIGH
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == 'f' and cell.value:
                        formula = str(cell.value)
                        lookup_ranges = self._extract_lookup_ranges(formula)
                        for rng in lookup_ranges:
                            analysis = self._analyze_lookup_range(sheet, rng)
                            if analysis['mixed_types']:
                                probability = self._calculate_probability(analysis)
                                if probability > 0:
                                    results.append(ErrorDetectionResult(
                                        error_type=self.name,
                                        description=f"Mixed data types detected in lookup range {rng} on sheet {sheet_name}",
                                        probability=probability,
                                        severity=self.severity,
                                        location=f"{sheet_name}!{rng}",
                                        details=analysis,
                                        suggested_fix="Convert all lookup keys to a consistent data type (all numbers or all text)."
                                    ))
        return results

    def _extract_lookup_ranges(self, formula: str) -> List[str]:
        # Extract ranges from lookup functions (VLOOKUP, HLOOKUP, XLOOKUP, MATCH, INDEX)
        import re
        # Simple pattern for ranges like A1:B10, Sheet2!A1:B10, etc.
        pattern = r"([A-Za-z0-9_']+!|)([A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+)"
        matches = re.findall(pattern, formula)
        ranges = []
        for match in matches:
            prefix = match[0]
            rng = match[1]
            ranges.append(prefix + rng)
        return ranges

    def _analyze_lookup_range(self, sheet, rng: str) -> dict:
        from openpyxl.utils import range_boundaries
        # Only analyze ranges on the current sheet for now
        if '!' in rng:
            sheet_name, rng = rng.split('!')
            if sheet.title != sheet_name.replace("'", ""):
                return {'mixed_types': False}
        try:
            min_col, min_row, max_col, max_row = range_boundaries(rng)
        except Exception:
            return {'mixed_types': False}
        type_counts = {'number': 0, 'text': 0, 'date': 0, 'other': 0, 'numeric_text': 0}
        total = 0
        all_text = True
        all_numeric_text = True
        for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for cell in row:
                if cell.value is None:
                    continue
                total += 1
                if isinstance(cell.value, (int, float)):
                    type_counts['number'] += 1
                    all_text = False
                elif self._is_date(cell):
                    type_counts['date'] += 1
                    all_text = False
                elif isinstance(cell.value, str):
                    if self._is_numeric_string(cell.value):
                        type_counts['numeric_text'] += 1
                    else:
                        type_counts['text'] += 1
                        all_numeric_text = False
                else:
                    type_counts['other'] += 1
                    all_text = False
                    all_numeric_text = False
        # If all non-empty cells are strings and all are numeric strings, treat as numbers
        if total > 0 and all_text and all_numeric_text and type_counts['numeric_text'] > 0:
            type_counts['number'] = type_counts['numeric_text']
            type_counts['numeric_text'] = 0
        else:
            # Otherwise, count numeric_text as text
            type_counts['text'] += type_counts['numeric_text']
            type_counts['numeric_text'] = 0
        # Mixed types if more than one type is present (excluding 'other')
        nonzero_types = [k for k, v in type_counts.items() if v > 0 and k not in ('other', 'numeric_text')]
        mixed_types = len(nonzero_types) > 1
        return {
            'type_counts': type_counts,
            'total': total,
            'mixed_types': mixed_types
        }

    def _is_number(self, value) -> bool:
        return isinstance(value, (int, float))

    def _is_numeric_string(self, value) -> bool:
        if not isinstance(value, str):
            return False
        try:
            float(value.replace(',', ''))
            return True
        except Exception:
            return False

    def _is_date(self, cell) -> bool:
        return hasattr(cell, 'is_date') and cell.is_date

    def _calculate_probability(self, analysis: dict) -> float:
        if analysis['total'] == 0:
            return 0.0
        type_counts = analysis['type_counts']
        majority_type = max((k for k in type_counts if k != 'other'), key=lambda k: type_counts[k], default=None)
        inconsistent = sum(v for k, v in type_counts.items() if k != majority_type and k != 'other')
        ratio = inconsistent / analysis['total'] if analysis['total'] else 0
        if ratio > 0.1:
            return 0.9
        elif ratio > 0.01:
            return 0.5
        elif ratio > 0:
            return 0.2
        return 0.0


class ConditionalFormattingOverlapConflictsDetector(ErrorDetector):
    """
    Detector for overlapping/conflicting conditional formatting rules.
    
    Algorithm:
    1. Extract all conditional formatting rules and their ranges
    2. Detect overlapping ranges
    3. Analyze rule types and formats for conflicts
    4. Calculate probability based on severity of overlap/conflict
    """
    def __init__(self):
        super().__init__(
            name="conditional_formatting_overlap_conflicts",
            description="Multiple conditional formatting rules that conflict or overlap",
            severity=ErrorSeverity.MEDIUM
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            # 1. Extract all conditional formatting rules and their ranges
            cf_rules = self._extract_conditional_formatting_rules(sheet)
            # Compare all pairs, including rules within the same range
            n = len(cf_rules)
            for i in range(n):
                for j in range(i + 1, n):
                    rule1 = cf_rules[i]
                    rule2 = cf_rules[j]
                    overlap_cells = rule1['cells'] & rule2['cells']
                    if not overlap_cells:
                        continue
                    # 3. Analyze for conflicts
                    conflict_type, probability = self._analyze_conflict(rule1, rule2)
                    if probability > 0:
                        results.append(ErrorDetectionResult(
                            error_type=self.name,
                            description=f"Conditional formatting overlap/conflict between rules on {sheet_name}: {rule1['range']} and {rule2['range']}",
                            probability=probability,
                            severity=self.severity,
                            location=f"{sheet_name}!{rule1['range']} & {rule2['range']}",
                            details={
                                'rule1': rule1,
                                'rule2': rule2,
                                'overlap_cells': list(overlap_cells),
                                'conflict_type': conflict_type
                            },
                            suggested_fix="Review overlapping conditional formatting rules and resolve conflicts."
                        ))
        return results

    def _extract_conditional_formatting_rules(self, sheet) -> List[dict]:
        # openpyxl stores conditional formatting in sheet.conditional_formatting
        rules = []
        cf = getattr(sheet, 'conditional_formatting', None)
        if cf is None:
            return rules
        # Iterate over cf._cf_rules.items(), use cf_range.sqref for the range string
        for cf_range, rule_list in cf._cf_rules.items():
            range_str = str(getattr(cf_range, 'sqref', '')) if hasattr(cf_range, 'sqref') else None
            if not range_str:
                continue
            from openpyxl.utils import range_boundaries, get_column_letter
            try:
                min_col, min_row, max_col, max_row = range_boundaries(range_str)
            except Exception:
                continue
            cells = set()
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    cell = f"{get_column_letter(col)}{row}"
                    cells.add(cell)
            for rule in rule_list:
                rules.append({
                    'range': range_str,
                    'cells': cells,
                    'type': getattr(rule, 'type', None),
                    'formula': getattr(rule, 'formula', None),
                    'dxf': getattr(rule, 'dxf', None),
                    'priority': getattr(rule, 'priority', None),
                    'rule_obj': rule
                })
        return rules

    def _analyze_conflict(self, rule1: dict, rule2: dict) -> tuple:
        # Check for type and format conflicts
        type1 = rule1['type']
        type2 = rule2['type']
        dxf1 = rule1['dxf']
        dxf2 = rule2['dxf']
        # High probability: both set fill/font and are different
        if dxf1 and dxf2:
            fill1 = getattr(dxf1, 'fill', None)
            fill2 = getattr(dxf2, 'fill', None)
            font1 = getattr(dxf1, 'font', None)
            font2 = getattr(dxf2, 'font', None)
            if (fill1 and fill2 and fill1 != fill2) or (font1 and font2 and font1 != font2):
                return ('conflicting_format', 0.9)
        # Medium probability: different types (e.g., color scale vs formula)
        if type1 != type2:
            return ('different_types', 0.6)
        # Low probability: overlap but compatible
        return ('overlap', 0.3)


class ExternalDataConnectionFailuresDetector(ErrorDetector):
    """
    Detector for external data connection failures (broken or outdated links).
    
    Algorithm:
    1. Scan for external links and data connections
    2. Check if target file/database exists (if accessible)
    3. Check for broken/unavailable links in workbook metadata
    4. Check for error values in cells that depend on external data
    5. Check last refresh date/time if available
    6. Calculate probability based on severity
    """
    def __init__(self):
        super().__init__(
            name="external_data_connection_failures",
            description="Links to external databases or files that are broken or outdated",
            severity=ErrorSeverity.HIGH
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        import os
        results = []
        # 1. Scan for external links (to other workbooks/files)
        external_links = getattr(workbook, 'external_links', [])
        for link in external_links:
            target = getattr(link, 'target', None)
            if not target:
                continue
            # 2. Check if the target file exists (if it's a file path)
            if os.path.isfile(target):
                probability = 0.2  # Low probability if file exists
            else:
                probability = 0.95  # High probability if file is missing
            results.append(ErrorDetectionResult(
                error_type=self.name,
                description=f"External link to file '{target}' is {'missing' if probability > 0.5 else 'present'}",
                probability=probability,
                severity=self.severity if probability > 0.5 else ErrorSeverity.LOW,
                location=f"external_link:{target}",
                details={'target': target, 'link': str(link)},
                suggested_fix="Update or remove the broken external link."
            ))
        # 2b. Scan for data connections (databases, web queries, etc.)
        connections = getattr(workbook, 'connections', None)
        if connections:
            for conn in connections:
                # Check last refresh date/time if available
                last_refresh = getattr(conn, 'last_refresh', None)
                if last_refresh:
                    import datetime
                    if isinstance(last_refresh, str):
                        try:
                            last_refresh = datetime.datetime.fromisoformat(last_refresh)
                        except Exception:
                            last_refresh = None
                    if last_refresh:
                        days_since = (datetime.datetime.now() - last_refresh).days
                        if days_since > 30:
                            probability = 0.6
                        else:
                            probability = 0.2
                    else:
                        probability = 0.3
                else:
                    probability = 0.3
                results.append(ErrorDetectionResult(
                    error_type=self.name,
                    description=f"External data connection '{getattr(conn, 'name', 'unknown')}' may be outdated or unverifiable",
                    probability=probability,
                    severity=ErrorSeverity.MEDIUM if probability > 0.3 else ErrorSeverity.LOW,
                    location=f"connection:{getattr(conn, 'name', 'unknown')}",
                    details={'connection': str(conn)},
                    suggested_fix="Check the data connection and refresh if needed."
                ))
        # 3. Check for error values in cells that depend on external data
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.upper() in {'#REF!', '#VALUE!', '#N/A'}:
                        results.append(ErrorDetectionResult(
                            error_type=self.name,
                            description=f"Error value '{cell.value}' in cell {cell.coordinate} on sheet {sheet_name} may be due to external data connection failure",
                            probability=0.8,
                            severity=self.severity,
                            location=f"{sheet_name}!{cell.coordinate}",
                            details={'cell': cell.coordinate, 'value': cell.value},
                            suggested_fix="Check external data sources and update or fix broken links."
                        ))
        return results


class PrecisionErrorsInFinancialCalculationsDetector(ErrorDetector):
    """
    Detector for floating-point precision errors in financial calculations.
    
    Algorithm:
    1. Scan for formulas with decimal arithmetic or financial functions
    2. Detect absence of rounding in such formulas
    3. Analyze for known precision issues (chained arithmetic, subtraction of nearly equal numbers)
    4. Calculate probability based on severity
    """
    def __init__(self):
        super().__init__(
            name="precision_errors_in_financial_calculations",
            description="Floating-point precision errors in financial calculations",
            severity=ErrorSeverity.MEDIUM
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        import re
        results = []
        financial_funcs = {'PMT', 'NPV', 'IRR', 'FV', 'PV', 'RATE', 'XNPV', 'XIRR', 'MIRR', 'DURATION', 'YIELD', 'COUPON', 'PRICE', 'DISC', 'TBILL', 'SLN', 'SYD', 'DB', 'DDB', 'VDB', 'AMORDEGRC', 'AMORLINC'}
        rounding_funcs = {'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'MROUND', 'TRUNC', 'INT', 'CEILING', 'FLOOR'}
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == 'f' and cell.value:
                        formula = str(cell.value).upper()
                        # 1. Check for financial functions
                        uses_financial_func = any(func in formula for func in financial_funcs)
                        uses_rounding = any(func in formula for func in rounding_funcs)
                        # 2. Check for decimal arithmetic (division, multiplication, subtraction, decimal point)
                        has_decimal_point = '.' in formula
                        cell_refs = re.findall(r"[A-Z][0-9]+", formula)
                        any_float = False
                        for ref in cell_refs:
                            try:
                                ref_cell = sheet[ref]
                                if isinstance(ref_cell.value, float):
                                    any_float = True
                                    break
                            except Exception:
                                continue
                        has_decimal_arithmetic = has_decimal_point or any_float
                        # 3. Check for subtraction of nearly equal numbers (e.g., A1-A2)
                        subtraction_matches = re.findall(r"([A-Z][0-9]+)\s*-\s*([A-Z][0-9]+)", formula)
                        # 4. Check for chained arithmetic (multiple operators)
                        operator_count = formula.count('+') + formula.count('-') + formula.count('*') + formula.count('/')
                        # 5. Ignore integer-only calculations
                        cell_refs = re.findall(r"[A-Z][0-9]+", formula)
                        all_integer = True
                        for ref in cell_refs:
                            try:
                                ref_cell = sheet[ref]
                                if not isinstance(ref_cell.value, int):
                                    all_integer = False
                                    break
                            except Exception:
                                all_integer = False
                                break
                        is_integer_only = all_integer and not has_decimal_arithmetic and not uses_financial_func
                        if is_integer_only:
                            continue
                        # 6. Probability calculation
                        probability = 0.0
                        if uses_financial_func and not uses_rounding:
                            probability = 0.9
                        elif operator_count >= 3 and not uses_rounding:
                            probability = 0.8
                        elif subtraction_matches and not uses_rounding:
                            probability = 0.8
                        elif has_decimal_arithmetic and not uses_rounding:
                            probability = 0.6
                        elif uses_rounding and operator_count >= 2:
                            probability = 0.3
                        if probability > 0:
                            results.append(ErrorDetectionResult(
                                error_type=self.name,
                                description=f"Potential precision error in formula {cell.coordinate} on sheet {sheet_name}",
                                probability=probability,
                                severity=self.severity if probability >= 0.6 else ErrorSeverity.LOW,
                                location=f"{sheet_name}!{cell.coordinate}",
                                details={
                                    'formula': formula,
                                    'uses_financial_func': uses_financial_func,
                                    'uses_rounding': uses_rounding,
                                    'has_decimal_arithmetic': has_decimal_arithmetic,
                                    'subtraction_matches': subtraction_matches,
                                    'operator_count': operator_count
                                },
                                suggested_fix="Use explicit rounding (e.g., ROUND) in all financial and decimal calculations."
                            ))
        return results


class IncompleteDragFormulaDetector(ErrorDetector):
    """
    Detector for incomplete drag formulas (formula cutoff / range incomplete formula errors).
    
    Algorithm:
    1. For each column, scan for contiguous blocks of formulas
    2. Identify expected data range (based on adjacent columns or max data row)
    3. Detect cutoffs (formulas stop before end of data) or gaps (missing formulas in middle)
    4. Calculate probability based on severity
    """
    def __init__(self):
        super().__init__(
            name="incomplete_drag_formula",
            description="Formula drag/copy stopped short, leaving cells without formulas (formula cutoff)",
            severity=ErrorSeverity.HIGH
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            max_row = sheet.max_row
            max_col = sheet.max_column
            # For each column, scan for formula blocks
            for col in range(1, max_col + 1):
                formula_rows = []
                data_rows = []
                for row in range(1, max_row + 1):
                    cell = sheet.cell(row=row, column=col)
                    if cell.value is not None:
                        data_rows.append(row)
                    if cell.data_type == 'f' and cell.value:
                        formula_rows.append(row)
                if not formula_rows or not data_rows:
                    continue
                # Expected range: from min(data_rows) to max(data_rows)
                expected_rows = set(range(min(data_rows), max(data_rows) + 1))
                formula_rows_set = set(formula_rows)
                missing_formula_rows = expected_rows - formula_rows_set
                if not missing_formula_rows:
                    continue  # No cutoff/gap
                # Probability calculation
                cutoff_at_end = max(formula_rows) < max(data_rows)
                gap_in_middle = any(row > min(formula_rows) and row < max(formula_rows) for row in missing_formula_rows)
                if gap_in_middle:
                    probability = 0.9
                elif cutoff_at_end and len(missing_formula_rows) > 1:
                    probability = 0.8
                elif cutoff_at_end:
                    probability = 0.6
                else:
                    probability = 0.3
                if probability > 0:
                    from openpyxl.utils import get_column_letter
                    col_letter = get_column_letter(col)
                    results.append(ErrorDetectionResult(
                        error_type=self.name,
                        description=f"Incomplete drag/copy of formula in column {col_letter} on sheet {sheet_name}; missing formulas at rows {sorted(missing_formula_rows)}",
                        probability=probability,
                        severity=self.severity if probability >= 0.6 else ErrorSeverity.LOW,
                        location=f"{sheet_name}!{col_letter}{min(data_rows)}:{col_letter}{max(data_rows)}",
                        details={
                            'column': col_letter,
                            'missing_formula_rows': sorted(missing_formula_rows),
                            'formula_rows': formula_rows,
                            'data_rows': data_rows
                        },
                        suggested_fix="Drag/copy the formula to all data rows; check for gaps or cutoffs."
                    ))
        return results


class FalseRangeEndDetectionDetector(ErrorDetector):
    """
    Detector for false range end detection (empty cell trap).
    
    Algorithm:
    1. For each column, find contiguous blocks of non-empty cells
    2. Identify empty cells in the middle of a data range
    3. Check if formulas only cover up to the first empty cell
    4. Flag if data exists after the gap but formulas stop at the gap
    5. Calculate probability based on severity
    """
    def __init__(self):
        super().__init__(
            name="false_range_end_detection",
            description="Empty cell in the middle of a data range causes formulas to stop early (false range end)",
            severity=ErrorSeverity.HIGH
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            max_row = sheet.max_row
            max_col = sheet.max_column
            for col in range(1, max_col + 1):
                data_rows = []
                formula_rows = []
                for row in range(1, max_row + 1):
                    cell = sheet.cell(row=row, column=col)
                    if cell.value is not None:
                        data_rows.append(row)
                    if cell.data_type == 'f' and cell.value:
                        formula_rows.append(row)
                if not data_rows or not formula_rows:
                    continue
                # Find the first empty cell in the data range
                min_data_row = min(data_rows)
                max_data_row = max(data_rows)
                empty_in_middle = None
                for row in range(min_data_row, max_data_row + 1):
                    cell = sheet.cell(row=row, column=col)
                    if cell.value is None:
                        empty_in_middle = row
                        break
                if empty_in_middle is None:
                    continue  # No gap in the middle
                # Check if there is data after the gap
                data_after_gap = [row for row in range(empty_in_middle + 1, max_data_row + 1) if sheet.cell(row=row, column=col).value is not None]
                if not data_after_gap:
                    continue  # No data after gap
                # Find missing formulas after the gap
                formula_rows_set = set(formula_rows)
                missing_formula_rows = set(row for row in range(empty_in_middle + 1, max_data_row + 1) if row not in formula_rows_set and sheet.cell(row=row, column=col).value is not None)
                if not missing_formula_rows:
                    continue
                # Probability calculation
                gap_size = len(missing_formula_rows)
                if gap_size > 2:
                    probability = 0.9
                elif gap_size > 0:
                    probability = 0.6
                else:
                    probability = 0.3
                from openpyxl.utils import get_column_letter
                col_letter = get_column_letter(col)
                results.append(ErrorDetectionResult(
                    error_type=self.name,
                    description=f"False range end detected in column {col_letter} on sheet {sheet_name}; missing formulas after empty cell at row {empty_in_middle}: {sorted(missing_formula_rows)}",
                    probability=probability,
                    severity=self.severity if probability >= 0.6 else ErrorSeverity.LOW,
                    location=f"{sheet_name}!{col_letter}{min_data_row}:{col_letter}{max_data_row}",
                    details={
                        'column': col_letter,
                        'empty_in_middle': empty_in_middle,
                        'data_after_gap': data_after_gap,
                        'missing_formula_rows': sorted(missing_formula_rows)
                    },
                    suggested_fix="Check for empty cells in the middle of data ranges and ensure formulas cover all data rows."
                ))
        return results


class PartialFormulaPropagationDetector(ErrorDetector):
    def __init__(self):
        super().__init__(
            name="partial_formula_propagation",
            description="Cells in a data range are missing formulas that are present in most other cells (partial formula propagation)",
            severity=ErrorSeverity.HIGH
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            max_row = sheet.max_row
            max_col = sheet.max_column
            for col in range(1, max_col + 1):
                formula_rows = []
                non_formula_rows = []
                formulas = {}
                for row in range(1, max_row + 1):
                    cell = sheet.cell(row=row, column=col)
                    if cell.data_type == 'f' and cell.value:
                        formula_rows.append(row)
                        formulas[row] = cell.value
                    elif cell.value is not None:
                        non_formula_rows.append(row)
                total_data_rows = len(formula_rows) + len(non_formula_rows)
                if total_data_rows < 5:
                    continue  # skip small ranges
                if len(formula_rows) / total_data_rows < 0.7:
                    continue  # skip if not mostly formulas
                # Find non-formula cells surrounded by formulas or at the edge
                missing_candidates = []
                for row in non_formula_rows:
                    prev_formula = any(r < row for r in formula_rows)
                    next_formula = any(r > row for r in formula_rows)
                    if (prev_formula and next_formula) or row == 1 or row == max_row:
                        missing_candidates.append(row)
                if not missing_candidates:
                    continue
                # Use the most common formula as the expected one
                from collections import Counter
                formula_counter = Counter(formulas.values())
                if not formula_counter:
                    continue
                expected_formula, _ = formula_counter.most_common(1)[0]
                # Probability: more missing = lower, more surrounded = higher, edge = 0.5
                for row in missing_candidates:
                    if row == 1 or row == max_row:
                        probability = 0.5
                        severity = ErrorSeverity.MEDIUM
                    else:
                        probability = 0.8 if len(missing_candidates) <= 2 else 0.6
                        severity = self.severity if probability >= 0.7 else ErrorSeverity.MEDIUM
                    from openpyxl.utils import get_column_letter
                    col_letter = get_column_letter(col)
                    results.append(ErrorDetectionResult(
                        error_type=self.name,
                        description=f"Cell {col_letter}{row} on sheet {sheet_name} is missing a formula, but most cells in this column have '{expected_formula}'",
                        probability=probability,
                        severity=severity,
                        location=f"{sheet_name}!{col_letter}{row}",
                        details={
                            'column': col_letter,
                            'row': row,
                            'expected_formula': expected_formula,
                            'missing_candidates': missing_candidates
                        },
                        suggested_fix=f"Copy the formula '{expected_formula}' to cell {col_letter}{row} if appropriate."
                    ))
        return results


class FormulaBoundaryMismatchDetector(ErrorDetector):
    def __init__(self):
        super().__init__(
            name="formula_boundary_mismatch",
            description="Aggregation formula references a range that does not cover all data (range misalignment)",
            severity=ErrorSeverity.HIGH
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        agg_funcs = ["SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN"]
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            max_row = sheet.max_row
            max_col = sheet.max_column
            for row in range(1, max_row + 1):
                for col in range(1, max_col + 1):
                    cell = sheet.cell(row=row, column=col)
                    if cell.data_type == 'f' and cell.value:
                        formula = str(cell.value).upper()
                        for func in agg_funcs:
                            if formula.startswith(f"={func}"):
                                # Extract range, e.g., =SUM(A1:A50)
                                import re
                                match = re.search(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", formula)
                                if not match:
                                    continue
                                start_col, start_row, end_col, end_row = match.groups()
                                if start_col != end_col:
                                    continue  # Only handle single-column ranges for now
                                col_idx = openpyxl.utils.column_index_from_string(start_col)
                                start_row = int(start_row)
                                end_row = int(end_row)
                                # Find actual data extent in this column
                                data_rows = [r for r in range(1, max_row + 1) if sheet.cell(row=r, column=col_idx).value is not None]
                                if not data_rows:
                                    continue
                                max_data_row = max(data_rows)
                                if max_data_row > end_row:
                                    # Data exists beyond the referenced range
                                    extra_data_count = max_data_row - end_row
                                    total_data_count = max_data_row - start_row + 1
                                    probability = min(0.9, 0.5 + 0.4 * (extra_data_count / total_data_count))
                                    from openpyxl.utils import get_column_letter
                                    col_letter = get_column_letter(col_idx)
                                    results.append(ErrorDetectionResult(
                                        error_type=self.name,
                                        description=f"Formula in {sheet_name}!{get_column_letter(col)}{row} references {col_letter}{start_row}:{col_letter}{end_row}, but data extends to row {max_data_row}.",
                                        probability=probability,
                                        severity=self.severity if probability >= 0.7 else ErrorSeverity.MEDIUM,
                                        location=f"{sheet_name}!{get_column_letter(col)}{row}",
                                        details={
                                            'formula': formula,
                                            'referenced_range': f"{col_letter}{start_row}:{col_letter}{end_row}",
                                            'max_data_row': max_data_row,
                                            'extra_data_count': extra_data_count
                                        },
                                        suggested_fix=f"Check if the aggregation formula should include data up to row {max_data_row}."
                                    ))
        return results


class CopyPasteFormulaGapsDetector(ErrorDetector):
    def __init__(self):
        super().__init__(
            name="copy_paste_formula_gaps",
            description="Gaps in formula sequences due to copy-paste errors (missing formulas in expected locations)",
            severity=ErrorSeverity.HIGH
        )

    def detect(self, workbook: openpyxl.Workbook, **kwargs) -> List[ErrorDetectionResult]:
        results = []
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            max_row = sheet.max_row
            max_col = sheet.max_column
            for col in range(1, max_col + 1):
                formula_rows = []
                non_formula_rows = []
                formulas = {}
                for row in range(1, max_row + 1):
                    cell = sheet.cell(row=row, column=col)
                    if cell.data_type == 'f' and cell.value:
                        formula_rows.append(row)
                        formulas[row] = cell.value
                    elif cell.value is not None:
                        non_formula_rows.append(row)
                if len(formula_rows) < 3:
                    continue  # Need at least 3 formulas to detect gaps
                # Find gaps in formula sequences
                formula_rows.sort()
                gaps = []
                for i in range(len(formula_rows) - 1):
                    current_row = formula_rows[i]
                    next_row = formula_rows[i + 1]
                    if next_row - current_row > 1:
                        # Check if there are non-formula cells in the gap
                        gap_cells = [r for r in range(current_row + 1, next_row) if r in non_formula_rows]
                        if gap_cells:
                            gaps.append((current_row, next_row, gap_cells))
                for start_row, end_row, gap_cells in gaps:
                    # Check if surrounding formulas are similar
                    start_formula = formulas[start_row]
                    end_formula = formulas[end_row]
                    if self._are_formulas_similar(start_formula, end_formula):
                        probability = 0.8 if len(gap_cells) <= 2 else 0.6
                        from openpyxl.utils import get_column_letter
                        col_letter = get_column_letter(col)
                        results.append(ErrorDetectionResult(
                            error_type=self.name,
                            description=f"Formula gap detected in column {col_letter} on sheet {sheet_name} between rows {start_row} and {end_row}; missing formulas in rows {gap_cells}.",
                            probability=probability,
                            severity=self.severity if probability >= 0.7 else ErrorSeverity.MEDIUM,
                            location=f"{sheet_name}!{col_letter}{start_row}:{col_letter}{end_row}",
                            details={
                                'column': col_letter,
                                'start_row': start_row,
                                'end_row': end_row,
                                'gap_cells': gap_cells,
                                'start_formula': start_formula,
                                'end_formula': end_formula
                            },
                            suggested_fix=f"Check for missing formulas in rows {gap_cells}; consider copying the formula pattern from adjacent cells."
                        ))
        return results

    def _are_formulas_similar(self, formula1: str, formula2: str) -> bool:
        """Check if two formulas follow a similar pattern (e.g., incrementing row references)."""
        # Simple heuristic: check if formulas have similar structure
        # This could be enhanced with more sophisticated pattern matching
        if not formula1 or not formula2:
            return False
        # Remove cell references and compare structure
        import re
        clean1 = re.sub(r'[A-Z]+\d+', 'CELL', formula1)
        clean2 = re.sub(r'[A-Z]+\d+', 'CELL', formula2)
        return clean1 == clean2


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def detect_excel_errors_probabilistic(
    file_path: Path, 
    error_threshold: float = 0.7,
    output_dir: Optional[Path] = None
) -> Dict[str, Any]:
    """
    Detect Excel errors using probabilistic models.
    
    Args:
        file_path: Path to the Excel file
        error_threshold: Minimum probability threshold (0.0 to 1.0)
        output_dir: Optional directory to save reports
        
    Returns:
        Dictionary containing detection results
    """
    sniffer = ProbabilisticErrorSniffer(file_path, error_threshold)
    
    # Register detectors
    sniffer.register_detector(HiddenDataInRangesDetector())
    sniffer.register_detector(CircularNamedRangesDetector())
    sniffer.register_detector(InconsistentDateFormatsDetector())
    sniffer.register_detector(ArrayFormulaSpillErrorsDetector())
    sniffer.register_detector(VolatileFunctionsDetector())
    sniffer.register_detector(CrossSheetReferenceErrorsDetector())
    sniffer.register_detector(DataTypeInconsistenciesInLookupTablesDetector())
    sniffer.register_detector(ConditionalFormattingOverlapConflictsDetector())
    sniffer.register_detector(ExternalDataConnectionFailuresDetector())
    sniffer.register_detector(PrecisionErrorsInFinancialCalculationsDetector())
    sniffer.register_detector(IncompleteDragFormulaDetector())
    sniffer.register_detector(FalseRangeEndDetectionDetector())
    sniffer.register_detector(PartialFormulaPropagationDetector())
    sniffer.register_detector(FormulaBoundaryMismatchDetector())
    sniffer.register_detector(CopyPasteFormulaGapsDetector())
    
    # Run detection
    results = sniffer.detect_all_errors()
    
    # Save reports if output directory specified
    if output_dir:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Ensure file_path is a Path object
        file_path = Path(file_path)
        
        # Save JSON report
        json_path = output_dir / f"{file_path.stem}_probabilistic_errors.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            # Convert results to serializable format
            serializable_results = {}
            for key, value in results.items():
                if key == 'summary':
                    serializable_results[key] = value
                else:
                    serializable_results[key] = [
                        {
                            'error_type': r.error_type,
                            'description': r.description,
                            'probability': r.probability,
                            'severity': r.severity.value,
                            'location': r.location,
                            'details': r.details,
                            'suggested_fix': r.suggested_fix
                        }
                        for r in value
                    ]
            json.dump(serializable_results, f, indent=2, default=str)
        
        logger.info(f"Probabilistic error analysis saved to: {json_path}")
    
    return results


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python probabilistic_error_detector.py <excel_file> [threshold] [output_dir]")
        sys.exit(1)
    
    file_path = Path(sys.argv[1])
    threshold = float(sys.argv[2]) if len(sys.argv) > 2 else 0.7
    output_dir = Path(sys.argv[3]) if len(sys.argv) > 3 else None
    
    try:
        results = detect_excel_errors_probabilistic(file_path, threshold, output_dir)
        print(f"Probabilistic error detection completed. Found {results['summary']['total_errors']} errors above threshold {threshold}.")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1) 