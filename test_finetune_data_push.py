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
