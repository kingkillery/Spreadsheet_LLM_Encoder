import json
import os
import unittest
from evaluation import (
    load_spreadsheet_dataset,
    load_table_detection_manifest,
    load_qa_dataset,
    load_qa_manifest,
    normalize_qa_answer,
    range_to_bbox,
)

class TestEvaluation(unittest.TestCase):

    def setUp(self):
        """Set up a dummy dataset directory for testing."""
        self.test_dir = "test_data"
        os.makedirs(self.test_dir, exist_ok=True)

        # Create dummy spreadsheet file
        self.spreadsheet_path = os.path.join(self.test_dir, "test1.xlsx")
        # In a real scenario, we'd use openpyxl to create a real xlsx file.
        # For this test, we'll just create an empty file.
        with open(self.spreadsheet_path, "w") as f:
            f.write("")

        # Create dummy annotation file for table detection
        self.ann_path_td = os.path.join(self.test_dir, "test1.json")
        with open(self.ann_path_td, "w") as f:
            json.dump({"tables": [{"range": "A1:B2"}, {"range": "D5:E10"}]}, f)

        # Create dummy annotation file for QA
        self.ann_path_qa = os.path.join(self.test_dir, "test_qa.json")
        self.spreadsheet_path_qa = os.path.join(self.test_dir, "test_qa.xlsx")
        with open(self.spreadsheet_path_qa, "w") as f:
            f.write("")
        with open(self.ann_path_qa, "w") as f:
            json.dump({"qa_pairs": [{"question": "Q1", "answer": "A1"}]}, f)


    def tearDown(self):
        """Clean up the dummy dataset directory."""
        import shutil
        shutil.rmtree(self.test_dir)

    def test_range_to_bbox(self):
        self.assertEqual(range_to_bbox("A1:B2"), (1, 1, 2, 2))
        self.assertEqual(range_to_bbox("C5:C5"), (5, 3, 5, 3))

    def test_normalize_qa_answer_by_type(self):
        self.assertEqual(normalize_qa_answer(" a1 ", "cell_address"), "A1")
        self.assertEqual(normalize_qa_answer(" sum( A1 : A3 ) ", "formula"), "SUM(A1:A3)")
        self.assertEqual(normalize_qa_answer("Total   Revenue", "free_text"), "total revenue")
        self.assertEqual(normalize_qa_answer("  Exact Case  ", "literal"), "Exact Case")

    def test_load_spreadsheet_dataset(self):
        dataset = load_spreadsheet_dataset(self.test_dir)
        self.assertEqual(len(dataset), 1)
        item = dataset[0]
        self.assertEqual(item["spreadsheet_path"], self.spreadsheet_path)
        self.assertEqual(len(item["bboxes"]), 2)
        self.assertEqual(item["bboxes"][0], (1, 1, 2, 2))

    def test_load_qa_dataset(self):
        # We need a separate json for qa test
        dataset = load_qa_dataset(self.test_dir)
        self.assertEqual(len(dataset), 1)
        item = dataset[0]
        self.assertEqual(item["spreadsheet_path"], self.spreadsheet_path_qa)
        self.assertEqual(len(item["qa_pairs"]), 1)
        self.assertEqual(item["qa_pairs"][0]["question"], "Q1")

    def test_load_table_detection_manifest(self):
        manifest_path = os.path.join(self.test_dir, "td_manifest.json")
        with open(manifest_path, "w") as f:
            json.dump({
                "dataset_name": "synthetic",
                "dataset_version": "v1",
                "split_name": "test",
                "items": [{
                    "spreadsheet_path": "test1.xlsx",
                    "tables": [{"range": "A1:B2"}],
                }],
            }, f)

        dataset = load_table_detection_manifest(manifest_path)

        self.assertEqual(len(dataset), 1)
        self.assertEqual(dataset[0]["bboxes"], [(1, 1, 2, 2)])
        self.assertEqual(dataset[0]["dataset_name"], "synthetic")
        self.assertEqual(dataset[0]["dataset_version"], "v1")
        self.assertEqual(dataset[0]["split_name"], "test")
        self.assertTrue(os.path.isabs(dataset[0]["spreadsheet_path"]))

    def test_load_qa_manifest(self):
        manifest_path = os.path.join(self.test_dir, "qa_manifest.json")
        with open(manifest_path, "w") as f:
            json.dump({
                "dataset_name": "qa_synth",
                "dataset_version": "v2",
                "split_name": "validation",
                "items": [{
                    "spreadsheet_path": "test_qa.xlsx",
                    "qa_pairs": [{
                        "question": "Q1",
                        "answer": "[A1]",
                        "answer_type": "cell_address",
                    }],
                }],
            }, f)

        dataset = load_qa_manifest(manifest_path)

        self.assertEqual(len(dataset), 1)
        self.assertEqual(dataset[0]["dataset_name"], "qa_synth")
        self.assertEqual(dataset[0]["split_name"], "validation")
        self.assertEqual(dataset[0]["qa_pairs"][0]["answer_type"], "cell_address")

if __name__ == '__main__':
    unittest.main()
