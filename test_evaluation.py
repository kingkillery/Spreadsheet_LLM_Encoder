import os
import json
import unittest
from evaluation import load_spreadsheet_dataset, load_qa_dataset, range_to_bbox

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

if __name__ == '__main__':
    unittest.main()
