#!/usr/bin/env python3
"""
Tests for Cross-Sheet Reference Errors Detector.

Covers:
- Valid cross-sheet references
- References to missing sheets
- References to missing cells/ranges
- Formulas with #REF! errors
- References to empty but valid cells
"""

import pytest
import tempfile
import shutil
from pathlib import Path
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from excel_analyzer.probabilistic_error_detector import CrossSheetReferenceErrorsDetector, ErrorSeverity

class TestCrossSheetReferenceErrorsDetector:
    def setup_method(self):
        self.detector = CrossSheetReferenceErrorsDetector()
        self.temp_dir = Path(tempfile.mkdtemp())
    def teardown_method(self):
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
    def create_test_workbook(self, setup_func):
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = 'Sheet1'
        ws2 = wb.create_sheet('Sheet2')
        setup_func(ws1, ws2, wb)
        return wb
    def test_valid_cross_sheet_reference(self):
        # Reference to a valid cell in another sheet
        def setup(ws1, ws2, wb):
            ws2['A1'].value = 123
            ws1['A1'].value = "=Sheet2!A1"
            ws1['A1'].data_type = 'f'
        wb = self.create_test_workbook(setup)
        results = self.detector.detect(wb)
        assert not any(r.error_type == 'cross_sheet_reference_errors' for r in results)
    def test_reference_to_missing_sheet(self):
        # Reference to a missing sheet
        def setup(ws1, ws2, wb):
            ws1['A1'].value = "=MissingSheet!A1"
            ws1['A1'].data_type = 'f'
        wb = self.create_test_workbook(setup)
        # Remove Sheet2 to simulate only Sheet1 present
        del wb['Sheet2']
        results = self.detector.detect(wb)
        assert any(r.error_type == 'cross_sheet_reference_errors' for r in results)
        for r in results:
            assert r.probability >= 0.9
            assert 'missing_sheet' in r.details
    def test_reference_to_missing_cell(self):
        # Reference to a cell outside the used range
        def setup(ws1, ws2, wb):
            ws1['A1'].value = "=Sheet2!Z99"
            ws1['A1'].data_type = 'f'
        wb = self.create_test_workbook(setup)
        results = self.detector.detect(wb)
        assert any(r.error_type == 'cross_sheet_reference_errors' for r in results)
        for r in results:
            assert r.probability >= 0.7
            assert 'missing_cell' in r.details
    def test_ref_error_in_formula(self):
        # Formula with #REF! error
        def setup(ws1, ws2, wb):
            ws1['A1'].value = "=#REF!"
            ws1['A1'].data_type = 'f'
        wb = self.create_test_workbook(setup)
        results = self.detector.detect(wb)
        assert any(r.error_type == 'cross_sheet_reference_errors' for r in results)
        for r in results:
            assert r.probability >= 0.9
    def test_reference_to_empty_cell(self):
        # Reference to an empty but valid cell
        def setup(ws1, ws2, wb):
            ws2['A1'].value = None
            ws1['A1'].value = "=Sheet2!A1"
            ws1['A1'].data_type = 'f'
        wb = self.create_test_workbook(setup)
        results = self.detector.detect(wb)
        assert any(r.error_type == 'cross_sheet_reference_errors' for r in results)
        for r in results:
            assert r.probability <= 0.3
            assert r.severity == ErrorSeverity.LOW

if __name__ == '__main__':
    pytest.main([__file__]) 