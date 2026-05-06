"""Tests for the --push-to-hub feature in prepare_finetuning_data.py.

These tests do NOT call the real HuggingFace Hub or perform any network I/O.
"""
import json
import os
import sys
import tempfile
import types
import unittest
from unittest.mock import MagicMock, patch, call


# ---------------------------------------------------------------------------
# Helpers for importing the module under test
# ---------------------------------------------------------------------------

sys.path.insert(0, os.path.dirname(__file__))


def _load_module():
    """Import prepare_finetuning_data fresh (avoids cached import state)."""
    import importlib
    import prepare_finetuning_data
    importlib.reload(prepare_finetuning_data)
    return prepare_finetuning_data


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestPushToHubMissingDatasetsPackage(unittest.TestCase):
    """push_to_hub() must raise a clear ImportError when `datasets` is absent."""

    def _write_sample_jsonl(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(json.dumps({"prompt": "p", "completion": "c"}) + "\n")

    def test_raises_import_error_with_helpful_message(self):
        mod = _load_module()
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".jsonl", delete=False, encoding="utf-8"
        ) as tf:
            tf.write(json.dumps({"prompt": "p", "completion": "c"}) + "\n")
            tmp_path = tf.name

        try:
            # Simulate `datasets` not being installed by injecting None into
            # sys.modules, which makes `import datasets` raise ImportError.
            with patch.dict(sys.modules, {"datasets": None}):
                with self.assertRaises(ImportError) as ctx:
                    mod.push_to_hub(
                        output_path=tmp_path,
                        repo_id="user/test-repo",
                        hub_split="train",
                        hub_private=True,
                    )
        finally:
            os.unlink(tmp_path)

        self.assertIn("pip install datasets", str(ctx.exception))


class TestFormatForFinetuningCoordinates(unittest.TestCase):
    """Fine-tuning targets must use the same compact coordinates as prompts."""

    def test_completion_ranges_are_remapped_to_prompt_coordinates(self):
        mod = _load_module()
        int_key = json.dumps({"type": "integer", "nfs": "0"}, sort_keys=True)
        coord_map = {
            "rows": {"2": 1, "3": 2},
            "cols": {"3": 1, "4": 2},
            "rows_inv": {"1": 2, "2": 3},
            "cols_inv": {"1": 3, "2": 4},
        }
        encoding = {
            "sheets": {
                "Sheet1": {
                    "cells": {},
                    "formats": {int_key: ["C2:D3"]},
                    "coord_map": coord_map,
                }
            }
        }

        records = mod.format_for_finetuning(encoding, [(2, 3, 3, 4)])

        self.assertEqual(len(records), 1)
        self.assertIn("(IntNum|A1:B2)", records[0]["prompt"])
        self.assertIn("'range': 'A1:B2'", records[0]["completion"])
        self.assertNotIn("'range': 'C2:D3'", records[0]["completion"])
        self.assertEqual(records[0]["metadata"]["coordinate_mode"], "compact_prompt_ranges")
        self.assertEqual(records[0]["metadata"]["prompt_ranges"], ["A1:B2"])
        self.assertEqual(records[0]["metadata"]["original_ranges"], ["C2:D3"])

    def test_unmapped_ground_truth_box_is_skipped(self):
        mod = _load_module()
        coord_map = {
            "rows": {"2": 1},
            "cols": {"3": 1},
            "rows_inv": {"1": 2},
            "cols_inv": {"1": 3},
        }
        encoding = {
            "sheets": {
                "Sheet1": {
                    "cells": {"x": ["C2"]},
                    "formats": {},
                    "coord_map": coord_map,
                }
            }
        }

        with self.assertLogs("prepare_finetuning_data", level="WARNING"):
            records = mod.format_for_finetuning(encoding, [(2, 3, 3, 4)])

        self.assertEqual(records[0]["completion"], "[]")
        self.assertEqual(records[0]["metadata"]["prompt_ranges"], [])
        self.assertEqual(records[0]["metadata"]["original_ranges"], [])


class TestFinetuneManifest(unittest.TestCase):
    """Fine-tuning JSONL can carry reproducibility sidecar metadata."""

    def test_build_manifest_records_coordinate_contract(self):
        mod = _load_module()

        manifest = mod.build_finetune_manifest(
            dataset_dir="datasets/train",
            output_path="out/train.jsonl",
            k=4,
            record_count=7,
            push_to_hub_repo="user/dataset",
            hub_split="train",
        )

        self.assertEqual(manifest["task"], "table_detection_finetuning_data")
        self.assertEqual(manifest["record_count"], 7)
        self.assertEqual(manifest["encoder_settings"]["k"], 4)
        self.assertEqual(manifest["coordinate_mode"], "compact_prompt_ranges")
        self.assertEqual(
            manifest["completion_coordinate_mode"],
            manifest["coordinate_mode"],
        )
        self.assertIn("prompt_template_sha256", manifest)

    def test_write_manifest_creates_json_file(self):
        mod = _load_module()
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", delete=False, encoding="utf-8"
        ) as tf:
            tmp_path = tf.name

        try:
            mod.write_finetune_manifest({"task": "x"}, tmp_path)
            with open(tmp_path, encoding="utf-8") as fh:
                data = json.load(fh)
        finally:
            os.unlink(tmp_path)

        self.assertEqual(data, {"task": "x"})


class TestPushToHubNotCalledWhenFlagAbsent(unittest.TestCase):
    """When --push-to-hub is None, datasets must never be imported or called."""

    def test_datasets_not_imported_when_push_flag_is_none(self):
        mod = _load_module()

        # Patch push_to_hub so we can detect any accidental call
        with patch.object(mod, "push_to_hub") as mock_push:
            # We also need main() to not actually try to load a real dataset;
            # patch the heavy I/O helpers.
            with patch("prepare_finetuning_data.load_spreadsheet_dataset", return_value=[]):
                mod.main(
                    dataset_dir="/fake/dir",
                    output_path=os.devnull,
                    k=4,
                    push_to_hub_repo=None,
                )

        mock_push.assert_not_called()


class TestPushToHubCallsDatasetCorrectly(unittest.TestCase):
    """When --push-to-hub is set, Dataset.from_list().push_to_hub() must be
    called with the correct repo_id, split, and private values."""

    def test_correct_hub_args(self):
        mod = _load_module()

        sample_records = [
            {"prompt": "hello", "completion": "world"},
            {"prompt": "foo", "completion": "bar"},
        ]

        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".jsonl", delete=False, encoding="utf-8"
        ) as tf:
            for rec in sample_records:
                tf.write(json.dumps(rec) + "\n")
            tmp_path = tf.name

        try:
            # Build a mock datasets module
            mock_ds_instance = MagicMock()
            mock_dataset_cls = MagicMock(return_value=mock_ds_instance)
            mock_dataset_cls.from_list = MagicMock(return_value=mock_ds_instance)

            mock_datasets_mod = MagicMock()
            mock_datasets_mod.Dataset = mock_dataset_cls

            with patch.dict(sys.modules, {"datasets": mock_datasets_mod}):
                mod.push_to_hub(
                    output_path=tmp_path,
                    repo_id="myuser/my-finetune-v1",
                    hub_split="validation",
                    hub_private=False,
                )
        finally:
            os.unlink(tmp_path)

        # Dataset.from_list must have received all records
        mock_dataset_cls.from_list.assert_called_once_with(sample_records)

        # push_to_hub on the dataset instance must have the right kwargs
        mock_ds_instance.push_to_hub.assert_called_once_with(
            "myuser/my-finetune-v1",
            split="validation",
            private=False,
        )


if __name__ == "__main__":
    unittest.main()
