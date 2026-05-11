import unittest
import os
import json
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill
from unittest.mock import patch
from Spreadsheet_LLM_Encoder import (
    spreadsheet_llm_encode,
    create_inverted_index,
    extract_formula_graph,
    extract_formula_references,
    find_boundary_candidates,
    aggregate_regions_dfs,
    vanilla_encode,
    main,
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
        self.assertIn(4, rows)  # Boundary between row 3 and 4
        self.assertIn(3, cols)  # Boundary between col B and C

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

    def test_title_rows_are_allowed_but_far_note_wrappers_are_rejected(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.merge_cells("A1:C1")
        ws["A1"] = "Revenue Summary"
        ws["A1"].font = Font(bold=True)
        for c, value in enumerate(["Region", "Revenue", "Owner"], start=1):
            ws.cell(row=2, column=c, value=value).font = Font(bold=True)
        ws.append(["West", 100, "Mia"])
        ws.append(["East", 150, "Leo"])
        ws["A8"] = "Note: preliminary export"
        ws["A8"].font = Font(italic=True)

        filtered = filter_unreasonable_candidates(
            ws,
            [(1, 1, 4, 3), (1, 1, 8, 3)],
        )

        self.assertIn((1, 1, 4, 3), filtered)
        self.assertNotIn((1, 1, 8, 3), filtered)

    def test_find_boundary_candidates_handles_title_header_table_and_notes(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.merge_cells("A1:C1")
        ws["A1"] = "Revenue Summary"
        ws["A1"].font = Font(bold=True)
        for c, value in enumerate(["Region", "Revenue", "Owner"], start=1):
            ws.cell(row=2, column=c, value=value).font = Font(bold=True)
        ws.append(["West", 100, "Mia"])
        ws.append(["East", 150, "Leo"])
        ws["A8"] = "Note: preliminary export"
        ws["A8"].font = Font(italic=True)

        rows, cols = find_boundary_candidates(ws)

        self.assertIn(1, rows)
        self.assertIn(4, rows)
        self.assertNotIn(8, rows)
        self.assertIn(1, cols)
        self.assertIn(3, cols)

    def test_overlap_resolution_prefers_dense_table_over_sparse_wrapper(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Generated by finance"
        for c, value in enumerate(["Region", "Revenue", "Cost"], start=1):
            ws.cell(row=2, column=c, value=value).font = Font(bold=True)
        ws.append(["West", 100, 40])
        ws.append(["East", 150, 60])
        ws["A6"] = "Note: draft"
        ws["A6"].font = Font(italic=True)

        filtered = filter_overlapping_candidates(
            ws,
            [(1, 1, 6, 3), (2, 1, 4, 3)],
        )

        self.assertEqual(filtered, [(2, 1, 4, 3)])

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

    def test_paper_strict_disables_homogeneous_compression(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "ID"
        ws["B1"] = "Value"
        for r in [2, 3]:
            ws.cell(row=r, column=1, value=0)
            ws.cell(row=r, column=2, value=0)
        ws["A4"] = "End"
        ws["B4"] = 5
        path = "paper_strict.xlsx"
        wb.save(path)

        try:
            result = spreadsheet_llm_encode(path, k=1, paper_strict=True)
        finally:
            os.remove(path)

        sheet = result["sheets"]["Sheet"]
        self.assertEqual(sheet["encoding_mode"], "paper_strict")
        refs = [ref for ranges in sheet["cells"].values() for ref in ranges]
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

    def test_sheet_filters_include_single_sheet_and_record_skips(self):
        result = spreadsheet_llm_encode(
            self.test_file,
            include_sheets=["Sheet2"],
        )

        self.assertEqual(list(result["sheets"].keys()), ["Sheet2"])
        selection = result["sheet_processing"]["selection"]
        self.assertEqual(selection["included_sheets"], ["Sheet2"])
        self.assertIn(
            {"sheet_name": "Sheet1", "reason": "sheet not matched by include filters"},
            selection["skipped_sheets"],
        )
        self.assertEqual(
            result["sheet_processing"]["sheets"]["Sheet1"]["reason"],
            "sheet not matched by include filters",
        )

    def test_sheet_filters_can_exclude_sheet_and_keep_rest(self):
        result = spreadsheet_llm_encode(
            self.test_file,
            exclude_sheets=["Sheet2"],
        )

        self.assertIn("Sheet1", result["sheets"])
        self.assertNotIn("Sheet2", result["sheets"])
        self.assertEqual(
            result["sheet_processing"]["sheets"]["Sheet2"]["reason"],
            "sheet excluded by name filter",
        )

    def test_cli_sheet_filters_write_metadata(self):
        out_path = "cli_sheet_filter.json"
        argv = [
            "Spreadsheet_LLM_Encoder.py",
            self.test_file,
            "--output",
            out_path,
            "--include-sheet",
            "Sheet1",
            "--exclude-sheet",
            "Sheet2",
        ]

        try:
            with patch.object(sys, "argv", argv):
                main()
            with open(out_path, encoding="utf-8") as fh:
                encoded = json.load(fh)
        finally:
            if os.path.exists(out_path):
                os.remove(out_path)

        self.assertIn("Sheet1", encoded["sheets"])
        self.assertNotIn("Sheet2", encoded["sheets"])
        self.assertEqual(
            encoded["sheet_processing"]["sheets"]["Sheet2"]["reason"],
            "sheet not matched by include filters",
        )

    def test_extract_formula_references_normalizes_local_and_cross_sheet_refs(self):
        refs = extract_formula_references("=SUM(B2:C3)+'Data Sheet'!D4+Aux!E5", "Sheet1")

        self.assertEqual(
            refs,
            ["Sheet1!B2:C3", "Data Sheet!D4", "Aux!E5"],
        )

    def test_formula_graph_captures_dependencies_errors_and_repeated_patterns(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "Item"
        ws["B1"] = "Units"
        ws["C1"] = "Price"
        ws["D1"] = "Total"
        ws["B2"] = 2
        ws["C2"] = 5
        ws["D2"] = "=B2*C2"
        ws["B3"] = 3
        ws["C3"] = 7
        ws["D3"] = "=B3*C3"
        ws["D4"] = "=SUM(D2:D3)"
        ws["E1"] = "#REF!"
        aux = wb.create_sheet("Aux")
        aux["A1"] = 10
        ws["F1"] = "=Aux!A1+D4"

        graph = extract_formula_graph(ws)

        formulas = {item["cell"]: item for item in graph["formulas"]}
        self.assertEqual(formulas["Sheet1!D2"]["references"], ["Sheet1!B2", "Sheet1!C2"])
        self.assertEqual(formulas["Sheet1!D4"]["references"], ["Sheet1!D2:D3"])
        self.assertEqual(formulas["Sheet1!F1"]["cross_sheet_references"], ["Aux!A1"])
        self.assertIn({"cell": "Sheet1!E1", "error": "#REF!"}, graph["formula_errors"])
        self.assertTrue(
            any(
                summary["kind"] == "pattern"
                and summary["formula_pattern"] == "=<REF>*<REF>"
                and summary["cells"] == ["Sheet1!D2", "Sheet1!D3"]
                for summary in graph["repeated_formula_summaries"]
            )
        )

    def test_spreadsheet_llm_encode_includes_formula_graph(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "Units"
        ws["B1"] = "Price"
        ws["C1"] = "Total"
        for cell in ("A1", "B1", "C1"):
            ws[cell].font = Font(bold=True)
        ws["A2"] = 2
        ws["B2"] = 5
        ws["C2"] = "=A2*B2"
        path = "formula_graph_encoding.xlsx"
        out_path = "formula_graph_encoding.json"
        wb.save(path)

        try:
            result = spreadsheet_llm_encode(path, out_path, k=1, paper_strict=True)
            with open(out_path, encoding="utf-8") as fh:
                saved = json.load(fh)
        finally:
            os.remove(path)
            if os.path.exists(out_path):
                os.remove(out_path)

        graph = result["sheets"]["Sheet1"]["formula_graph"]
        self.assertEqual(graph["formulas"][0]["cell"], "Sheet1!C2")
        self.assertEqual(graph["formulas"][0]["formula"], "=A2*B2")
        self.assertEqual(graph["formulas"][0]["references"], ["Sheet1!A2", "Sheet1!B2"])
        self.assertEqual(
            saved["sheets"]["Sheet1"]["formula_graph"]["formulas"][0]["cell"],
            "Sheet1!C2",
        )

    def test_bounded_mode_truncates_large_sheet_and_records_metadata(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Large"
        for col in range(1, 41):
            cell = ws.cell(row=1, column=col, value=f"H{col}")
            cell.font = Font(bold=True)
        for row in range(2, 302):
            for col in range(1, 41):
                ws.cell(row=row, column=col, value=row * col)
        path = "large_bounded.xlsx"
        out_path = "large_bounded.json"
        wb.save(path)

        try:
            result = spreadsheet_llm_encode(
                path,
                out_path,
                k=1,
                paper_strict=True,
                max_cells_per_sheet=1000,
            )
            with open(out_path, encoding="utf-8") as fh:
                saved = json.load(fh)
        finally:
            os.remove(path)
            if os.path.exists(out_path):
                os.remove(out_path)

        processing = result["sheet_processing"]
        sheet_meta = processing["sheets"]["Large"]
        self.assertEqual(processing["mode"], "bounded")
        self.assertEqual(sheet_meta["status"], "encoded")
        self.assertTrue(sheet_meta["truncated"])
        self.assertEqual(sheet_meta["original_rows"], 301)
        self.assertEqual(sheet_meta["original_cols"], 40)
        self.assertLessEqual(sheet_meta["effective_cells"], 1000)
        self.assertEqual(saved["sheet_processing"]["sheets"]["Large"], sheet_meta)
        self.assertIn("Large", result["sheets"])

    def test_bounded_mode_can_skip_large_sheet_and_record_reason(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Large"
        for col in range(1, 8):
            ws.cell(row=1, column=col, value=f"H{col}").font = Font(bold=True)
        for row in range(2, 60):
            for col in range(1, 8):
                ws.cell(row=row, column=col, value=row + col)
        path = "large_skip.xlsx"
        wb.save(path)

        try:
            result = spreadsheet_llm_encode(
                path,
                k=1,
                max_rows_per_sheet=10,
                sheet_limit_action="skip",
            )
        finally:
            os.remove(path)

        self.assertNotIn("Large", result["sheets"])
        sheet_meta = result["sheet_processing"]["sheets"]["Large"]
        self.assertEqual(sheet_meta["status"], "skipped")
        self.assertTrue(sheet_meta["truncated"])
        self.assertIn("reason", sheet_meta)

    def test_bounded_mode_can_error_on_large_sheet(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Large"
        ws["A1"] = "Header"
        ws["A2"] = "Value"
        path = "large_error.xlsx"
        wb.save(path)

        try:
            with self.assertRaises(ValueError):
                spreadsheet_llm_encode(
                    path,
                    max_rows_per_sheet=1,
                    sheet_limit_action="error",
                )
        finally:
            os.remove(path)


if __name__ == '__main__':
    unittest.main()
