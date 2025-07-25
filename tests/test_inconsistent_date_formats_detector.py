#!/usr/bin/env python3
"""
Tests for Inconsistent Date Formats Detector.

Covers:
- All Excel dates
- All text dates
- Mixed Excel and text dates
- Different date formats
- Columns with numbers and dates
- Edge cases (empty cells, non-date text)
"""

import pytest
import tempfile
import shutil
from pathlib import Path
import openpyxl
from openpyxl.utils.datetime import from_excel
from datetime import datetime

from excel_analyzer.probabilistic_error_detector import InconsistentDateFormatsDetector, ErrorSeverity

class TestInconsistentDateFormatsDetector:
    def setup_method(self):
        self.detector = InconsistentDateFormatsDetector()
        self.temp_dir = Path(tempfile.mkdtemp())
    def teardown_method(self):
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
    def create_test_workbook(self, data, col='A', number_format=None):
        wb = openpyxl.Workbook()
        ws = wb.active
        for i, value in enumerate(data, 1):
            cell = ws[f'{col}{i}']
            cell.value = value
            if number_format:
                cell.number_format = number_format
        return wb
    def test_all_excel_dates(self):
        # All cells are Excel dates
        dates = [datetime(2023, 1, i+1) for i in range(10)]
        wb = self.create_test_workbook(dates, number_format='YYYY-MM-DD')
        results = self.detector.detect(wb)
        assert len(results) == 0  # Should not flag
    def test_all_text_dates(self):
        # All cells are text dates
        dates = [f'2023-01-{i+1:02d}' for i in range(10)]
        wb = self.create_test_workbook(dates)
        results = self.detector.detect(wb)
        # Should flag, but lower probability
        assert any(r.error_type == 'inconsistent_date_formats' for r in results)
        for r in results:
            assert r.probability <= 0.5
    def test_mixed_excel_and_text_dates(self):
        # Mix of Excel dates and text dates
        data = [datetime(2023, 1, 1), '2023-01-02', datetime(2023, 1, 3), '2023-01-04'] * 3
        wb = self.create_test_workbook(data, number_format='YYYY-MM-DD')
        results = self.detector.detect(wb)
        assert any(r.error_type == 'inconsistent_date_formats' for r in results)
        for r in results:
            assert r.probability >= 0.9
            assert r.severity == ErrorSeverity.HIGH
            assert r.details['mixed_types']
    def test_different_date_formats(self):
        # Mix of different text date formats
        data = ['2023-01-01', '01/02/2023', '1 Jan 2023', '01.04.2023'] * 3
        wb = self.create_test_workbook(data)
        results = self.detector.detect(wb)
        assert any(r.error_type == 'inconsistent_date_formats' for r in results)
        for r in results:
            assert r.details['text_date_count'] == 12
            assert r.details['date_count'] == 0
    def test_numbers_and_dates(self):
        # Mix of numbers and dates
        data = [datetime(2023, 1, 1), 123, '2023-01-02', 456] * 3
        wb = self.create_test_workbook(data, number_format='YYYY-MM-DD')
        results = self.detector.detect(wb)
        # Should flag as mixed types if both Excel and text dates present
        assert any(r.error_type == 'inconsistent_date_formats' for r in results)
    def test_empty_cells(self):
        # Some empty cells
        data = [datetime(2023, 1, 1), None, '2023-01-02', None] * 3
        wb = self.create_test_workbook(data, number_format='YYYY-MM-DD')
        results = self.detector.detect(wb)
        assert any(r.error_type == 'inconsistent_date_formats' for r in results)
    def test_non_date_text(self):
        # Text that does not look like a date
        data = ['foo', 'bar', 'baz', '2023-01-01', datetime(2023, 1, 2)]
        wb = self.create_test_workbook(data, number_format='YYYY-MM-DD')
        results = self.detector.detect(wb)
        # Should only count the real date and text date
        for r in results:
            assert r.details['date_count'] >= 1
            assert r.details['text_date_count'] >= 1
    def test_no_dates(self):
        # No dates at all
        data = [123, 456, 789, 'foo', 'bar']
        wb = self.create_test_workbook(data)
        results = self.detector.detect(wb)
        assert len(results) == 0

if __name__ == '__main__':
    pytest.main([__file__]) 