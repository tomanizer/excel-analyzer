#!/usr/bin/env python3
"""
Tests for External Data Connection Failures Detector.

Covers:
- Valid external links (should not flag)
- Broken/missing external links (should flag, high probability)
- Connections not refreshed recently (should flag, medium probability)
- Connections to inaccessible resources (should flag as unverifiable)
- Error values in cells (#REF!, #VALUE!, #N/A)
"""

import pytest
import tempfile
import shutil
from pathlib import Path
import openpyxl
from types import SimpleNamespace
from datetime import datetime, timedelta

from excel_analyzer.probabilistic_error_detector import ExternalDataConnectionFailuresDetector, ErrorSeverity

class TestExternalDataConnectionFailuresDetector:
    def setup_method(self):
        self.detector = ExternalDataConnectionFailuresDetector()
        self.temp_dir = Path(tempfile.mkdtemp())
    def teardown_method(self):
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
    def create_test_workbook(self, external_links=None, connections=None, error_cells=None):
        wb = openpyxl.Workbook()
        if external_links is not None:
            wb.external_links = external_links
        if connections is not None:
            wb.connections = connections
        ws = wb.active
        if error_cells:
            for coord, value in error_cells.items():
                ws[coord].value = value
        return wb
    def test_valid_external_link(self):
        # Valid external link (file exists)
        temp_file = self.temp_dir / 'external.xlsx'
        temp_file.touch()
        link = SimpleNamespace(target=str(temp_file))
        wb = self.create_test_workbook(external_links=[link])
        results = self.detector.detect(wb)
        assert any(r.probability <= 0.2 for r in results)
    def test_broken_external_link(self):
        # Broken external link (file missing)
        link = SimpleNamespace(target=str(self.temp_dir / 'missing.xlsx'))
        wb = self.create_test_workbook(external_links=[link])
        results = self.detector.detect(wb)
        assert any(r.probability >= 0.9 for r in results)
    def test_outdated_connection(self):
        # Connection not refreshed recently
        conn = SimpleNamespace(name='TestConn', last_refresh=(datetime.now() - timedelta(days=40)).isoformat())
        wb = self.create_test_workbook(connections=[conn])
        results = self.detector.detect(wb)
        assert any(r.probability >= 0.6 for r in results)
    def test_recent_connection(self):
        # Connection refreshed recently
        conn = SimpleNamespace(name='TestConn', last_refresh=(datetime.now() - timedelta(days=5)).isoformat())
        wb = self.create_test_workbook(connections=[conn])
        results = self.detector.detect(wb)
        assert any(r.probability <= 0.2 for r in results)
    def test_error_values_in_cells(self):
        # Error values in cells
        error_cells = {'A1': '#REF!', 'B2': '#VALUE!', 'C3': '#N/A'}
        wb = self.create_test_workbook(error_cells=error_cells)
        results = self.detector.detect(wb)
        assert any(r.probability >= 0.8 for r in results)
        for r in results:
            assert r.severity == ErrorSeverity.HIGH

if __name__ == '__main__':
    pytest.main([__file__]) 