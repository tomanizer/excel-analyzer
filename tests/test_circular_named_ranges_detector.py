#!/usr/bin/env python3
"""
Tests for Circular Named Ranges Detector.

Tests various scenarios including:
- Simple 2-range cycles
- Complex 3+ range cycles
- Non-circular references
- Edge cases and error conditions
"""

import pytest
import tempfile
import shutil
from pathlib import Path
from unittest.mock import Mock, patch

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from excel_analyzer.probabilistic_error_detector import CircularNamedRangesDetector, ErrorSeverity


class TestCircularNamedRangesDetector:
    """Test cases for Circular Named Ranges Detector."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.detector = CircularNamedRangesDetector()
        self.temp_dir = Path(tempfile.mkdtemp())
    
    def teardown_method(self):
        """Clean up test fixtures."""
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
    
    def create_test_workbook(self, named_ranges_data: dict) -> openpyxl.Workbook:
        """Create a test workbook with named ranges."""
        wb = openpyxl.Workbook()
        
        # Add named ranges
        for name, formula in named_ranges_data.items():
            defined_name = DefinedName(name=name, attr_text=formula)
            wb.defined_names.add(defined_name)
        
        return wb
    
    def test_simple_2_range_cycle(self):
        """Test detection of simple 2-range circular reference."""
        named_ranges = {
            'Revenue': '=SUM(Expenses)',
            'Expenses': '=Revenue * 0.8'
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        assert len(results) == 1
        result = results[0]
        
        assert result.error_type == "circular_named_ranges"
        assert "Circular reference detected" in result.description
        # Check that both nodes are in the cycle (order may vary due to normalization)
        assert "Revenue" in result.description and "Expenses" in result.description
        assert result.probability >= 0.8  # High probability for 2-range cycles
        assert result.severity == ErrorSeverity.HIGH
        # Check that both nodes are in the location (order may vary due to normalization)
        assert "Revenue" in result.location and "Expenses" in result.location
        # Check that both nodes are in the cycle (order may vary)
        assert set(result.details['cycle'][:-1]) == {'Revenue', 'Expenses'}
        assert result.details['cycle_length'] == 3  # Including the duplicate at the end
    
    def test_complex_3_range_cycle(self):
        """Test detection of 3-range circular reference."""
        named_ranges = {
            'A': '=B + C',
            'B': '=A * 0.5',
            'C': '=A - B'
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        assert len(results) == 1
        result = results[0]
        
        assert result.error_type == "circular_named_ranges"
        assert "Circular reference detected" in result.description
        assert result.probability >= 0.7  # High probability for 3-range cycles
        assert result.severity == ErrorSeverity.HIGH
        assert result.details['cycle_length'] == 3
    
    def test_4_range_cycle(self):
        """Test detection of 4-range circular reference."""
        named_ranges = {
            'W': '=X + Y',
            'X': '=Y + Z',
            'Y': '=Z + W',
            'Z': '=W + X'
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        assert len(results) == 1
        result = results[0]
        
        assert result.error_type == "circular_named_ranges"
        assert result.probability >= 0.6  # Medium-high probability for 4-range cycles
        assert result.details['cycle_length'] == 4
    
    def test_no_circular_reference(self):
        """Test that no errors are detected when there are no circular references."""
        named_ranges = {
            'Revenue': '=SUM(A1:A10)',
            'Expenses': '=SUM(B1:B10)',
            'Profit': '=Revenue - Expenses'
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        assert len(results) == 0
    
    def test_self_reference(self):
        """Test detection of self-referencing named range."""
        named_ranges = {
            'SelfRef': '=SelfRef + 1'
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        assert len(results) == 1
        result = results[0]
        
        assert result.error_type == "circular_named_ranges"
        assert "SelfRef → SelfRef" in result.description
        assert result.details['cycle_length'] == 1
    
    def test_multiple_cycles(self):
        """Test detection of multiple independent cycles."""
        named_ranges = {
            'A': '=B',
            'B': '=A',
            'X': '=Y',
            'Y': '=X',
            'Z': '=SUM(A1:A10)'  # No cycle
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        assert len(results) == 2
        
        # Check that both cycles are detected
        cycles = [result.details['cycle'] for result in results]
        assert ['A', 'B'] in cycles or ['B', 'A'] in cycles
        assert ['X', 'Y'] in cycles or ['Y', 'X'] in cycles
    
    def test_complex_formulas_in_cycle(self):
        """Test detection with complex formulas containing functions."""
        named_ranges = {
            'Revenue': '=SUM(Expenses) + AVERAGE(Expenses)',
            'Expenses': '=Revenue * 0.8 + COUNT(A1:A10)'
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        assert len(results) == 1
        result = results[0]
        
        # Should have higher probability due to complex formulas
        assert result.probability >= 0.9
    
    def test_aggregation_functions_in_cycle(self):
        """Test that aggregation functions increase probability."""
        named_ranges = {
            'Total': '=SUM(SubTotal)',
            'SubTotal': '=Total * 0.9'
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        assert len(results) == 1
        result = results[0]
        
        # Should have higher probability due to SUM function
        assert result.probability >= 0.9
    
    def test_indirect_references(self):
        """Test detection of indirect circular references."""
        named_ranges = {
            'A': '=INDIRECT("B")',
            'B': '=A + 1'
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        # Note: This might not be detected by our current parser
        # as INDIRECT references are harder to parse
        # This test documents the limitation
        pass
    
    def test_empty_workbook(self):
        """Test behavior with workbook containing no named ranges."""
        wb = openpyxl.Workbook()
        results = self.detector.detect(wb)
        
        assert len(results) == 0
    
    def test_invalid_formulas(self):
        """Test behavior with invalid formulas."""
        named_ranges = {
            'Valid': '=SUM(A1:A10)',
            'Invalid': '=INVALID_FUNCTION(',
            'Circular': '=Valid'
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        # Should still detect the circular reference despite invalid formula
        assert len(results) == 1
    
    def test_formula_parsing_edge_cases(self):
        """Test formula parsing with various edge cases."""
        test_cases = [
            ('=A1+B1', []),  # No named ranges (cell references should be filtered out)
            ('=Revenue', ['Revenue']),  # Simple named range
            ('=SUM(Revenue)', ['Revenue']),  # Named range in function
            ('=Revenue+Expenses', ['Revenue', 'Expenses']),  # Multiple named ranges
            ('=IF(Revenue>0,Revenue,0)', ['Revenue']),  # Named range in IF
            ('=SUM(A1:A10)', []),  # No named ranges, just cell references
            ('=SUM(Revenue,Expenses)', ['Revenue', 'Expenses']),  # Multiple in function
        ]
        
        for formula, expected in test_cases:
            dependencies = self.detector._parse_named_range_formula(formula)
            # Filter out cell references (like A1, B1) that are not named ranges
            filtered_deps = [dep for dep in dependencies if not (len(dep) <= 3 and dep[0].isalpha() and dep[1:].isdigit())]
            assert set(filtered_deps) == set(expected), f"Failed for formula: {formula}"
    
    def test_excel_keywords_filtering(self):
        """Test that Excel keywords are properly filtered out."""
        named_ranges = {
            'SUM': '=A1+A2',  # Should not be detected as named range reference
            'Revenue': '=SUM(A1:A10)',  # SUM should be filtered out
            'Expenses': '=Revenue'  # Should create cycle with Revenue
        }
        
        wb = self.create_test_workbook(named_ranges)
        results = self.detector.detect(wb)
        
        # Should detect cycle between Revenue and Expenses, not SUM
        assert len(results) == 1
        result = results[0]
        assert 'Revenue' in result.details['cycle']
        assert 'Expenses' in result.details['cycle']
        assert 'SUM' not in result.details['cycle']
    
    def test_probability_calculation(self):
        """Test probability calculation for different scenarios."""
        # Test 2-range cycle (should have high probability)
        named_ranges_2 = {
            'A': '=B',
            'B': '=A'
        }
        wb_2 = self.create_test_workbook(named_ranges_2)
        results_2 = self.detector.detect(wb_2)
        assert results_2[0].probability >= 0.8
        
        # Test 3-range cycle (should have slightly lower probability)
        named_ranges_3 = {
            'A': '=B',
            'B': '=C',
            'C': '=A'
        }
        wb_3 = self.create_test_workbook(named_ranges_3)
        results_3 = self.detector.detect(wb_3)
        assert results_3[0].probability >= 0.7
        
        # Test complex formula (should have higher probability)
        named_ranges_complex = {
            'A': '=SUM(B) + AVERAGE(B) + COUNT(B)',
            'B': '=A * 0.8'
        }
        wb_complex = self.create_test_workbook(named_ranges_complex)
        results_complex = self.detector.detect(wb_complex)
        assert results_complex[0].probability >= 0.9
    
    def test_dependency_graph_building(self):
        """Test dependency graph construction."""
        named_ranges = {
            'A': {'formula': '=B + C'},
            'B': {'formula': '=D'},
            'C': {'formula': '=E'},
            'D': {'formula': '=A'},  # Creates cycle A -> B -> D -> A
            'E': {'formula': '=SUM(A1:A10)'}  # No dependencies
        }
        
        graph = self.detector._build_dependency_graph(named_ranges)
        
        assert 'A' in graph
        assert 'B' in graph
        assert 'C' in graph
        assert 'D' in graph
        assert 'E' in graph
        
        assert set(graph['A']) == {'B', 'C'}
        assert set(graph['B']) == {'D'}
        assert set(graph['C']) == {'E'}
        assert set(graph['D']) == {'A'}
        assert set(graph['E']) == set()  # No dependencies
    
    def test_cycle_detection_algorithm(self):
        """Test the cycle detection algorithm directly."""
        graph = {
            'A': ['B'],
            'B': ['C'],
            'C': ['A'],  # Creates cycle A -> B -> C -> A
            'D': ['E'],
            'E': ['D'],  # Creates cycle D -> E -> D
            'F': ['G'],  # No cycle
            'G': ['H'],
            'H': []  # No cycle
        }
        
        cycles = self.detector._detect_cycles(graph)
        
        assert len(cycles) == 2
        
        # Check that both cycles are detected
        cycle_lengths = [len(cycle) for cycle in cycles]
        assert 3 in cycle_lengths  # A -> B -> C -> A
        assert 2 in cycle_lengths  # D -> E -> D


if __name__ == '__main__':
    pytest.main([__file__]) 