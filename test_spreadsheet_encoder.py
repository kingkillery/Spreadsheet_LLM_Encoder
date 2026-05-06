import unittest
import os
import json
import openpyxl
from openpyxl.styles import Font, PatternFill
from Spreadsheet_LLM_Encoder import (
    spreadsheet_llm_encode,
    create_inverted_index,
    find_boundary_candidates,
    aggregate_regions_dfs,
    vanilla_encode,
    is_header_row,
    filter_unreasonable_candidates,
    filter_overlapping_candidates,
)

class TestSpreadsheetEncoder(unittest.TestCase):

    def setUp(self):
        """Set up a test workbook."""
        self.test_file = "test_workbook.xlsx"
        wb = openpyxl.Workbook()

        # Sheet 1: For boundary and aggregation tests
        ws1 = wb.active
        ws1.title = "Sheet1"
        ws1['A1'] = "Header 1"
        ws1['A1'].font = Font(bold=True)
        ws1['B1'] = "Header 2"
        ws1['B1'].font = Font(bold=True)
        ws1['A2'] = 100
        ws1['B2'] = 200
        ws1['A3'] = 150
        ws1['B3'] = 250

        # Add a separate region with a different format
        ws1['D4'] = "Data"
        ws1['D4'].fill = PatternFill("solid", fgColor="FFFF00")
        ws1['E4'] = "More Data"
        ws1['E4'].fill = PatternFill("solid", fgColor="FFFF00")

        # Sheet 2: For vanilla encoding test
        ws2 = wb.create_sheet("Sheet2")
        ws2['A1'] = "Hello"
        ws2['B1'] = "World"

        wb.save(self.test_file)

    def tearDown(self):
        """Remove the test workbook."""
        os.remove(self.test_file)

    def test_vanilla_encode(self):
        result = vanilla_encode(self.test_file)
        self.assertIn("Sheet1", result)
        self.assertIn("Sheet2", result)
        self.assertTrue(result["Sheet1"].startswith("A1,Header 1|B1,Header 2"))
        self.assertTrue(result["Sheet2"].startswith("A1,Hello|B1,World"))

    def test_find_boundary_candidates_advanced(self):
        wb = openpyxl.load_workbook(self.test_file)
        sheet = wb["Sheet1"]
        rows, cols = find_boundary_candidates(sheet)
        # This is a basic check; the full heuristics are complex.
        # We expect a boundary between the two groups of cells.
        self.assertIn(4, rows) # Boundary between row 3 and 4
        self.assertIn(3, cols) # Boundary between col B and C

    def test_plain_text_header_detected_without_style(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Region", "Revenue", "Cost"])
        ws.append(["West", 100, 50])
        ws.append(["East", 200, 90])

        self.assertTrue(is_header_row(ws, 1))
        self.assertFalse(is_header_row(ws, 2))

    def test_year_header_detected(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append([2022, 2023, 2024])
        ws.append([10, 20, 30])

        self.assertTrue(is_header_row(ws, 1))

    def test_sparse_note_candidate_filtered(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Revenue Report"
        ws["A1"].font = Font(bold=True)
        ws["A20"] = "Note"
        ws["D20"] = "Draft"

        filtered = filter_unreasonable_candidates(ws, [(1, 1, 20, 4)])

        self.assertEqual(filtered, [])

    def test_merged_header_contributes_boundaries(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.merge_cells("A1:C1")
        ws["A1"] = "Quarterly Revenue"
        ws["A1"].font = Font(bold=True)
        ws.append(["Region", "Q1", "Q2"])
        ws.append(["West", 10, 20])
        ws.append(["East", 30, 40])

        rows, cols = find_boundary_candidates(ws)

        self.assertIn(1, rows)
        self.assertTrue(any(c in cols for c in (1, 3)))

    def test_multi_table_sheet_keeps_separate_column_boundaries(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Product"
        ws["B1"] = "Sales"
        ws["D1"] = "Region"
        ws["E1"] = "Cost"
        for cell in ("A1", "B1", "D1", "E1"):
            ws[cell].font = Font(bold=True)
        ws["A2"] = "A"
        ws["B2"] = 10
        ws["D2"] = "West"
        ws["E2"] = 4

        rows, cols = find_boundary_candidates(ws)

        self.assertTrue(rows)
        self.assertIn(2, cols)
        self.assertIn(4, cols)

    def test_overlap_resolution_prefers_stronger_candidate(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        for c, value in enumerate(["A", "B", "C", "D"], start=1):
            ws.cell(row=1, column=c, value=value).font = Font(bold=True)
        for r in range(2, 5):
            for c in range(1, 5):
                ws.cell(row=r, column=c, value=r * c)

        filtered = filter_overlapping_candidates(
            ws,
            [(1, 1, 4, 4), (1, 1, 4, 3)],
        )

        self.assertEqual(filtered, [(1, 1, 4, 4)])

    def test_strict_skeleton_retains_homogeneous_rows(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "ID"
        ws["B1"] = "Value"
        for r in [2, 3]:
            ws.cell(row=r, column=1, value=0)
            ws.cell(row=r, column=2, value=0)
        ws["A4"] = "End"
        ws["B4"] = 5
        path = "strict_skeleton.xlsx"
        wb.save(path)

        try:
            result = spreadsheet_llm_encode(path, k=1, compress_homogeneous=False)
        finally:
            os.remove(path)

        cells = result["sheets"]["Sheet"]["cells"]
        refs = [ref for ranges in cells.values() for ref in ranges]
        self.assertTrue(any(ref.endswith("2") or ref.endswith("3") for ref in refs))

    def test_aggregate_regions_dfs(self):
        wb = openpyxl.load_workbook(self.test_file)
        sheet = wb["Sheet1"]

        format_map = {
            json.dumps({"type": "integer", "nfs": "General"}, sort_keys=True): ["A2", "B2", "A3", "B3"],
            json.dumps({"type": "text", "nfs": "General"}, sort_keys=True): ["D4", "E4"]
        }

        aggregated = aggregate_regions_dfs(sheet, format_map)

        # Check that the regions were aggregated correctly
        key1 = json.dumps({"type": "integer", "nfs": "General"}, sort_keys=True)
        key2 = json.dumps({"type": "text", "nfs": "General"}, sort_keys=True)
        self.assertIn("A2:B3", aggregated[key1])
        self.assertIn("D4:E4", aggregated[key2])

    def test_paper_format_grouping_ignores_rich_styles(self):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet["A1"] = 100
        sheet["B1"] = 200
        sheet["A1"].number_format = "#,##0"
        sheet["B1"].number_format = "#,##0"
        sheet["A1"].font = Font(bold=True)
        sheet["B1"].fill = PatternFill("solid", fgColor="FFFF00")

        _, paper_format_map = create_inverted_index(sheet, [1], [1, 2])
        _, rich_format_map = create_inverted_index(
            sheet, [1], [1, 2], format_mode="rich"
        )

        self.assertEqual(1, len(paper_format_map))
        paper_key = json.loads(next(iter(paper_format_map.keys())))
        self.assertEqual("#,##0", paper_key["nfs"])
        self.assertIn(paper_key["type"], {"integer", "numeric"})
        self.assertGreater(len(rich_format_map), 1)

    def test_spreadsheet_llm_encode_runs(self):
        result = spreadsheet_llm_encode(self.test_file)
        self.assertIsNotNone(result)
        self.assertIn("Sheet1", result["sheets"])

if __name__ == '__main__':
    unittest.main()
