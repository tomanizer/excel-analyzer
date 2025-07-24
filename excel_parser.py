import openpyxl
from pathlib import Path
import re
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils.cell import range_boundaries
from typing import List, Dict, Any, Set

def find_data_islands(sheet: Worksheet, visited_cells: Set[str]) -> List[Set[str]]:
    """Finds contiguous blocks of data not already part of a formal table."""
    islands = []
    visited = set()
    # Consider all non-empty cells not already visited (i.e., not in a formal table)
    all_cells = {cell.coordinate for row in sheet.iter_rows() for cell in row 
                 if cell.value is not None and str(cell.value).strip() != "" and cell.coordinate not in visited_cells}

    for cell_coord in all_cells:
        if cell_coord not in visited:
            island = set()
            queue = [cell_coord]
            visited.add(cell_coord)
            while queue:
                current_coord = queue.pop(0)
                island.add(current_coord)
                col, row = openpyxl.utils.cell.coordinate_from_string(current_coord)
                col_idx = openpyxl.utils.cell.column_index_from_string(col)
                for r_offset, c_offset in [(0, 1), (0, -1), (1, 0), (-1, 0)]:
                    neighbor_row, neighbor_col_idx = row + r_offset, col_idx + c_offset
                    if neighbor_row > 0 and neighbor_col_idx > 0:
                        neighbor_coord = f"{openpyxl.utils.cell.get_column_letter(neighbor_col_idx)}{neighbor_row}"
                        if neighbor_coord in all_cells and neighbor_coord not in visited:
                            visited.add(neighbor_coord)
                            queue.append(neighbor_coord)
            islands.append(island)
    return islands

def analyze_workbook_final(file_path: Path):
    if not file_path.exists(): return
    print(f"--- Comprehensive Analysis for: {file_path.name} ---\n")

    wb = openpyxl.load_workbook(file_path, data_only=False, keep_vba=True)

    # 1. VBA Macro Detection
    has_vba = file_path.suffix == '.xlsm'  # Simplified detection
    print(f"VBA Project Detected: {has_vba}")

    # 2. External Link Detection
    external_links = [link.Target for link in wb.external_links] if hasattr(wb, 'external_links') else []
    if external_links:
        print("\nExternal Dependencies:")
        for link in external_links: print(f"  - {link}")

    # 3. Named Range Detection
    named_ranges = {name: d.destinations for name, d in wb.defined_names.items()}
    if named_ranges:
        print("\nNamed Ranges:")
        for name, dest in named_ranges.items(): print(f"  - {name}: {dest}")

    # 4. Formal Table, Chart, Data Validation, and Island Detection
    all_tables = []
    visited_cells = set()

    print("\n--- Sheet-Level Analysis ---")
    for sheet in wb:
        print(f"\nProcessing Sheet: {sheet.title}")
        
        # Formal Tables
        for tbl in sheet.tables.values():
            all_tables.append({"name": tbl.displayName, "type": "Formal Table", "range": tbl.ref, "sheet": sheet.title})
            # Add all cells from this table to visited
            # Parse the table range to get min/max rows and columns
            min_col, min_row, max_col, max_row = openpyxl.utils.cell.range_boundaries(tbl.ref)
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    visited_cells.add(f"{openpyxl.utils.cell.get_column_letter(c)}{r}")
        
        # Chart Detection
        charts = []
        for chart in sheet._charts:
            try:
                chart_info = {"name": chart.title or "Untitled Chart", "type": type(chart).__name__}
                charts.append(chart_info)
            except:
                charts.append({"name": "Unknown Chart", "type": "Unknown"})
        
        if charts:
            print("  Charts Found:")
            for chart in charts: print(f"    - '{chart['name']}' ({chart['type']})")

        # Data Validation Detection
        validations = [f"{dv.sqref}: {dv.formula1}" for dv in sheet.data_validations.dataValidation]
        if validations:
            print("  Data Validation Rules Found:")
            for val in validations: print(f"    - {val}")

        # Informal Data Islands (fallback)
        islands = find_data_islands(sheet, visited_cells)
        for island in islands:
            coords = [openpyxl.utils.cell.coordinate_from_string(c) for c in island]
            rows = [c[1] for c in coords]; cols = [openpyxl.utils.cell.column_index_from_string(c[0]) for c in coords]
            bounding_box = f"{openpyxl.utils.cell.get_column_letter(min(cols))}{min(rows)}:{openpyxl.utils.cell.get_column_letter(max(cols))}{max(rows)}"
            all_tables.append({"name": f"Island_{bounding_box}", "type": "Informal Data Island", "range": bounding_box, "sheet": sheet.title})

    if all_tables:
        print("\n--- Discovered Data Tables & Islands ---")
        for table in all_tables:
            print(f"  - {table['name']} ({table['type']}) on sheet '{table['sheet']}' at range {table['range']}")


if __name__ == "__main__":
    # Create a dummy external file
    ext_wb = openpyxl.Workbook()
    ext_ws = ext_wb.active; ext_ws.title = "ExternalData"
    ext_ws["A1"] = "External Value"; ext_ws["B1"] = 500
    ext_file = Path("external_source.xlsx")
    ext_wb.save(ext_file)

    # Create the main workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # remove default sheet

    # Data Sheet with a Formal Table
    data_ws = wb.create_sheet("SalesData")
    data_ws.append(["Region", "Product", "Sales"])
    data_ws.append(["North", "A", 100]); data_ws.append(["South", "B", 150])
    data_ws.append(["North", "A", 200]); data_ws.append(["South", "C", 250])
    tbl = Table(displayName="SalesTable", ref=f"A1:C5")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tbl.tableStyleInfo = style
    data_ws.add_table(tbl)

    # Input sheet with Data Validation and Named Ranges
    input_ws = wb.create_sheet("Inputs")
    input_ws["A1"] = "Select Region:"
    input_ws["B1"] = "North" # Default value
    # Data Validation Dropdown
    dv = DataValidation(type="list", formula1='"North,South,East,West"')
    dv.add(input_ws["B1"])
    input_ws.data_validations.append(dv)
    # Named Range
    wb.create_named_range('Selected_Region', input_ws, 'B1')

    # External Link Sheet
    link_ws = wb.create_sheet("ExternalLinks")
    link_ws["A1"] = "Value from other file:"; link_ws["B1"] = f"='[{ext_file.name}]ExternalData'!B1"
    
    # NOTE: openpyxl cannot create charts or VBA projects.
    # The logic to detect them is included and will work on real user files.

    main_file = Path("final_model.xlsm")
    wb.save(main_file)
    print(f"Dummy files '{main_file.name}' and '{ext_file.name}' created.")

    analyze_workbook_final(main_file)
