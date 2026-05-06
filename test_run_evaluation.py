"""Tests for TableSense-CNN evaluation record behavior."""
import json
import os
import tempfile
import unittest
from unittest.mock import patch

import run_evaluation


class TestRunEvaluation(unittest.TestCase):

    def test_missing_tablesense_writes_skip_record(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            record_path = os.path.join(tmpdir, "tablesense.json")
            with patch.object(run_evaluation, "TableSenseCNN", None):
                run_evaluation.main(tmpdir, out_record=record_path)

            with open(record_path, encoding="utf-8") as fh:
                record = json.load(fh)

        self.assertTrue(record["skipped"])
        self.assertIn("evaluation_metadata", record)
        metadata = record["evaluation_metadata"]
        self.assertEqual(metadata["baseline_name"], "TableSense-CNN")
        self.assertEqual(metadata["skip_reasons"][0]["component"], "tablesense-cnn")


if __name__ == "__main__":
    unittest.main()

