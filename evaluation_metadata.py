"""Evaluation parity metadata helpers.

Evaluation numbers are only comparable when the dataset, model/backend,
encoding, coordinate mode, and metric definition are explicit. This module
keeps that contract small and JSON-serializable so scripts can attach the same
metadata shape to table-detection and QA result records.
"""
from __future__ import annotations

import datetime as _dt
import json
import os
from typing import Any, Dict, Iterable, List, Optional


REQUIRED_METADATA_FIELDS = (
    "dataset_name",
    "dataset_version",
    "split_name",
    "spreadsheet_count",
    "table_count",
    "qa_item_count",
    "encoder_settings",
    "prompt_serializer",
    "coordinate_mode",
    "model_backend",
    "metric_definition",
    "baseline_name",
    "baseline_version",
    "skip_reasons",
)


def build_evaluation_metadata(
    *,
    dataset_dir: str,
    task: str,
    split_name: str = "unspecified",
    dataset_name: Optional[str] = None,
    dataset_version: str = "unspecified",
    spreadsheet_count: int = 0,
    table_count: int = 0,
    qa_item_count: int = 0,
    encoder_settings: Optional[Dict[str, Any]] = None,
    prompt_serializer: str,
    coordinate_mode: str,
    model_backend: str,
    metric_definition: str,
    baseline_name: str,
    baseline_version: str = "unspecified",
    skip_reasons: Optional[Iterable[Dict[str, str]]] = None,
) -> Dict[str, Any]:
    """Build the common metadata block required for comparable results."""
    dataset_abs = os.path.abspath(dataset_dir)
    return {
        "created_at": _dt.datetime.now(_dt.timezone.utc).isoformat(),
        "task": task,
        "dataset_name": dataset_name or os.path.basename(dataset_abs) or dataset_abs,
        "dataset_version": dataset_version,
        "dataset_dir": dataset_abs,
        "split_name": split_name,
        "spreadsheet_count": int(spreadsheet_count),
        "table_count": int(table_count),
        "qa_item_count": int(qa_item_count),
        "encoder_settings": encoder_settings or {},
        "prompt_serializer": prompt_serializer,
        "coordinate_mode": coordinate_mode,
        "model_backend": model_backend,
        "metric_definition": metric_definition,
        "baseline_name": baseline_name,
        "baseline_version": baseline_version,
        "skip_reasons": list(skip_reasons or []),
    }


def validate_evaluation_metadata(metadata: Dict[str, Any]) -> List[str]:
    """Return validation errors for an evaluation metadata block."""
    errors: List[str] = []
    if not isinstance(metadata, dict):
        return ["evaluation_metadata must be an object"]

    for field in REQUIRED_METADATA_FIELDS:
        if field not in metadata:
            errors.append(f"missing evaluation_metadata.{field}")

    for field in (
        "dataset_name",
        "dataset_version",
        "split_name",
        "prompt_serializer",
        "coordinate_mode",
        "model_backend",
        "metric_definition",
        "baseline_name",
        "baseline_version",
    ):
        value = metadata.get(field)
        if value is None or value == "":
            errors.append(f"empty evaluation_metadata.{field}")

    for field in ("spreadsheet_count", "table_count", "qa_item_count"):
        value = metadata.get(field)
        if not isinstance(value, int) or value < 0:
            errors.append(f"evaluation_metadata.{field} must be a non-negative integer")

    if not isinstance(metadata.get("encoder_settings"), dict):
        errors.append("evaluation_metadata.encoder_settings must be an object")
    if not isinstance(metadata.get("skip_reasons"), list):
        errors.append("evaluation_metadata.skip_reasons must be an array")
    return errors


def validate_evaluation_record(record: Dict[str, Any]) -> List[str]:
    """Return validation errors for a full result record."""
    if not isinstance(record, dict):
        return ["record must be an object"]
    if "evaluation_metadata" not in record:
        return ["missing evaluation_metadata"]
    return validate_evaluation_metadata(record["evaluation_metadata"])


def write_evaluation_record(record: Dict[str, Any], output_path: str) -> None:
    """Validate and write an evaluation record as pretty JSON."""
    errors = validate_evaluation_record(record)
    if errors:
        raise ValueError("; ".join(errors))
    out_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(record, fh, indent=2)


def validate_finetune_eval_compatibility(
    finetune_manifest: Dict[str, Any],
    evaluation_metadata: Dict[str, Any],
) -> List[str]:
    """Return errors when fine-tune data and evaluation settings diverge."""
    errors: List[str] = []
    if not isinstance(finetune_manifest, dict):
        return ["finetune_manifest must be an object"]
    meta_errors = validate_evaluation_metadata(evaluation_metadata)
    if meta_errors:
        errors.extend(meta_errors)
        return errors

    ft_k = (finetune_manifest.get("encoder_settings") or {}).get("k")
    eval_k = (evaluation_metadata.get("encoder_settings") or {}).get("k")
    if ft_k != eval_k:
        errors.append(f"encoder k mismatch: finetune={ft_k!r} evaluation={eval_k!r}")

    ft_serializer = finetune_manifest.get("prompt_serializer")
    eval_serializer = evaluation_metadata.get("prompt_serializer")
    if ft_serializer and eval_serializer and ft_serializer not in eval_serializer:
        errors.append(
            "prompt serializer mismatch: "
            f"finetune={ft_serializer!r} evaluation={eval_serializer!r}"
        )

    ft_coord = finetune_manifest.get("coordinate_mode")
    eval_coord = evaluation_metadata.get("coordinate_mode")
    if ft_coord and eval_coord and "compact" in ft_coord and "compact" not in eval_coord:
        errors.append(
            "coordinate mode mismatch: "
            f"finetune={ft_coord!r} evaluation={eval_coord!r}"
        )
    return errors


__all__ = [
    "REQUIRED_METADATA_FIELDS",
    "build_evaluation_metadata",
    "validate_evaluation_metadata",
    "validate_evaluation_record",
    "validate_finetune_eval_compatibility",
    "write_evaluation_record",
]
