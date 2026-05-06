"""Tests for evaluation parity metadata records."""
import os
import tempfile
import unittest

from evaluation_metadata import (
    REQUIRED_METADATA_FIELDS,
    build_evaluation_metadata,
    validate_finetune_eval_compatibility,
    validate_evaluation_metadata,
    validate_evaluation_record,
    write_evaluation_record,
)


class TestEvaluationMetadata(unittest.TestCase):

    def test_build_metadata_contains_required_fields(self):
        metadata = build_evaluation_metadata(
            dataset_dir="datasets/synthetic",
            task="table_detection",
            spreadsheet_count=2,
            table_count=3,
            qa_item_count=0,
            encoder_settings={"k": 4},
            prompt_serializer="paper_serializers.to_paper_compressed_prompt",
            coordinate_mode="compact_prompt_unmapped_to_original_for_eob0",
            model_backend="echo",
            metric_definition="EoB-0 exact boundary matching; threshold=0.0",
            baseline_name="SpreadsheetLLM table detection",
        )

        for field in REQUIRED_METADATA_FIELDS:
            self.assertIn(field, metadata)
        self.assertEqual(validate_evaluation_metadata(metadata), [])

    def test_validate_record_rejects_missing_metadata(self):
        self.assertEqual(
            validate_evaluation_record({"task": "table_detection_eob0"}),
            ["missing evaluation_metadata"],
        )

    def test_validate_metadata_rejects_missing_field(self):
        metadata = build_evaluation_metadata(
            dataset_dir="datasets/synthetic",
            task="qa",
            prompt_serializer="serializer",
            coordinate_mode="coordinate_mode",
            model_backend="echo",
            metric_definition="exact match",
            baseline_name="SpreadsheetLLM QA",
        )
        del metadata["coordinate_mode"]

        errors = validate_evaluation_metadata(metadata)

        self.assertIn("missing evaluation_metadata.coordinate_mode", errors)

    def test_write_evaluation_record_validates_before_write(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "record.json")
            metadata = build_evaluation_metadata(
                dataset_dir=tmpdir,
                task="table_detection",
                prompt_serializer="serializer",
                coordinate_mode="coordinate_mode",
                model_backend="echo",
                metric_definition="metric",
                baseline_name="baseline",
            )
            write_evaluation_record(
                {"task": "table_detection_eob0", "evaluation_metadata": metadata},
                out_path,
            )
            self.assertTrue(os.path.exists(out_path))

    def test_validate_finetune_eval_compatibility_accepts_matching_contracts(self):
        finetune_manifest = {
            "encoder_settings": {"k": 4},
            "prompt_serializer": "paper_serializers.to_paper_compressed_prompt",
            "coordinate_mode": "compact_prompt_ranges",
        }
        metadata = build_evaluation_metadata(
            dataset_dir="datasets/synthetic",
            task="table_detection",
            encoder_settings={"k": 4},
            prompt_serializer="paper_serializers.to_paper_compressed_prompt",
            coordinate_mode="compact_prompt_unmapped_to_original_for_eob0",
            model_backend="echo",
            metric_definition="EoB-0",
            baseline_name="SpreadsheetLLM table detection",
        )

        self.assertEqual(
            validate_finetune_eval_compatibility(finetune_manifest, metadata),
            [],
        )

    def test_validate_finetune_eval_compatibility_rejects_mismatched_k(self):
        finetune_manifest = {
            "encoder_settings": {"k": 2},
            "prompt_serializer": "paper_serializers.to_paper_compressed_prompt",
            "coordinate_mode": "compact_prompt_ranges",
        }
        metadata = build_evaluation_metadata(
            dataset_dir="datasets/synthetic",
            task="table_detection",
            encoder_settings={"k": 4},
            prompt_serializer="paper_serializers.to_paper_compressed_prompt",
            coordinate_mode="compact_prompt_unmapped_to_original_for_eob0",
            model_backend="echo",
            metric_definition="EoB-0",
            baseline_name="SpreadsheetLLM table detection",
        )

        errors = validate_finetune_eval_compatibility(finetune_manifest, metadata)

        self.assertTrue(any("encoder k mismatch" in err for err in errors))


if __name__ == "__main__":
    unittest.main()
