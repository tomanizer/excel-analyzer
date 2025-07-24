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
    
    # Register the first detector
    sniffer.register_detector(HiddenDataInRangesDetector())
    
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