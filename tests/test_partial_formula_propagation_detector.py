import pytest
import openpyxl
from openpyxl.utils import get_column_letter
from src.excel_analyzer.probabilistic_error_detector import PartialFormulaPropagationDetector, ErrorSeverity


def create_sheet_with_partial_formulas(rows, missing_rows=None, edge_missing=False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    missing_rows = set(missing_rows or [])
    for row in range(1, rows + 1):
        cell = ws.cell(row=row, column=1)
        if edge_missing and (row == 1 or row == rows):
            continue
        if row not in missing_rows:
            cell.value = f"=A{row}+1"
            cell.data_type = 'f'
        else:
            cell.value = 42  # hardcoded value
    return wb

def test_all_formulas():
    wb = create_sheet_with_partial_formulas(20)
    detector = PartialFormulaPropagationDetector()
    results = detector.detect(wb)
    assert not results

def test_one_missing_in_middle():
    wb = create_sheet_with_partial_formulas(20, missing_rows=[10])
    detector = PartialFormulaPropagationDetector()
    results = detector.detect(wb)
    assert results
    assert results[0].probability >= 0.7
    assert results[0].details['row'] == 10
    assert results[0].severity == ErrorSeverity.HIGH

def test_one_missing_at_edge():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    # Set row 1 to a hardcoded value (missing formula), rest have formulas
    ws.cell(row=1, column=1).value = 42
    for row in range(2, 21):
        cell = ws.cell(row=row, column=1)
        cell.value = f"=A{row}+1"
        cell.data_type = 'f'
    detector = PartialFormulaPropagationDetector()
    results = detector.detect(wb)
    assert results
    assert results[0].probability == 0.5
    assert results[0].details['row'] == 1
    assert results[0].severity == ErrorSeverity.MEDIUM

def test_multiple_missing():
    wb = create_sheet_with_partial_formulas(20, missing_rows=[5, 10, 15])
    detector = PartialFormulaPropagationDetector()
    results = detector.detect(wb)
    assert len(results) == 3
    for r in results:
        assert r.probability == 0.6
        assert r.severity == ErrorSeverity.MEDIUM

def test_small_range():
    wb = create_sheet_with_partial_formulas(3, missing_rows=[2])
    detector = PartialFormulaPropagationDetector()
    results = detector.detect(wb)
    assert not results

def test_all_non_formulas():
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in range(1, 20):
        ws.cell(row=row, column=1).value = 42
    detector = PartialFormulaPropagationDetector()
    results = detector.detect(wb)
    assert not results 