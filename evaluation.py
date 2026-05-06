import json
import logging
import os
from typing import Any, Dict, List, Tuple
import xml.etree.ElementTree as ET
from openpyxl.utils import column_index_from_string, get_column_letter
import re

logger = logging.getLogger(__name__)

BBox = Tuple[int, int, int, int]


def _resolve_manifest_path(manifest_path: str, maybe_relative: str) -> str:
    if os.path.isabs(maybe_relative):
        return maybe_relative
    return os.path.abspath(os.path.join(os.path.dirname(manifest_path), maybe_relative))


def range_to_bbox(range_str: str) -> BBox:
    """Convert an Excel-style range (e.g., 'A1:F9') to a BBox tuple."""
    parts = range_str.split(':')
    start_cell = parts[0]
    end_cell = parts[1] if len(parts) > 1 else start_cell

    col_start_str = ''.join(filter(str.isalpha, start_cell))
    row_start_str = ''.join(filter(str.isdigit, start_cell))

    col_end_str = ''.join(filter(str.isalpha, end_cell))
    row_end_str = ''.join(filter(str.isdigit, end_cell))

    c1 = column_index_from_string(col_start_str)
    r1 = int(row_start_str)
    c2 = column_index_from_string(col_end_str)
    r2 = int(row_end_str)

    return (r1, c1, r2, c2)


def load_spreadsheet_dataset(path: str) -> List[Dict[str, object]]:
    """
    Load a spreadsheet dataset with annotations in JSON format.
    Expects pairs of .xlsx and .json files.
    """
    dataset = []
    for fname in os.listdir(path):
        if fname.endswith(".xlsx"):
            spreadsheet_path = os.path.join(path, fname)
            ann_path = os.path.join(path, fname.replace(".xlsx", ".json"))

            if not os.path.exists(ann_path):
                logger.warning(f"Annotation file not found for {fname}, skipping.")
                continue

            with open(ann_path, 'r') as f:
                annotations = json.load(f)

            tables = annotations.get("tables", [])
            if not tables:
                continue

            bboxes = [range_to_bbox(t['range']) for t in tables]

            dataset.append({
                "spreadsheet_path": spreadsheet_path,
                "bboxes": bboxes,
                "ann_path": ann_path
            })
    return dataset


def load_table_detection_manifest(manifest_path: str) -> List[Dict[str, object]]:
    """Load a manifest-defined spreadsheet table-detection dataset.

    Expected shape::

        {
          "dataset_name": "synthetic_v1",
          "dataset_version": "1",
          "split_name": "test",
          "items": [
            {
              "spreadsheet_path": "book.xlsx",
              "tables": [{"range": "A1:B2"}]
            }
          ]
        }

    Relative spreadsheet paths are resolved from the manifest directory.
    """
    with open(manifest_path, encoding="utf-8") as fh:
        manifest: Dict[str, Any] = json.load(fh)

    dataset_name = manifest.get("dataset_name", "manifest")
    dataset_version = manifest.get("dataset_version", "unspecified")
    split_name = manifest.get("split_name", "unspecified")
    items = manifest.get("items")
    if not isinstance(items, list):
        raise ValueError("table-detection manifest must contain an items array")

    dataset = []
    for idx, item in enumerate(items):
        if not isinstance(item, dict):
            raise ValueError(f"manifest item {idx} must be an object")
        spreadsheet_path = item.get("spreadsheet_path")
        tables = item.get("tables", [])
        if not spreadsheet_path:
            raise ValueError(f"manifest item {idx} missing spreadsheet_path")
        if not isinstance(tables, list):
            raise ValueError(f"manifest item {idx} tables must be an array")
        bboxes = [range_to_bbox(t["range"]) for t in tables if "range" in t]
        dataset.append({
            "spreadsheet_path": _resolve_manifest_path(manifest_path, spreadsheet_path),
            "bboxes": bboxes,
            "manifest_path": os.path.abspath(manifest_path),
            "dataset_name": dataset_name,
            "dataset_version": dataset_version,
            "split_name": split_name,
        })
    return dataset


def load_dong2019_dataset(path: str) -> List[Dict[str, object]]:
    """Load the Dong et al. (2019) table detection dataset.

    The function expects two subdirectories under ``path``:
    ``images`` containing the page images and ``annotations`` with
    Pascal VOC XML files describing table bounding boxes.

    Returns a list of dictionaries with ``image_path`` and ``bboxes``.
    """
    ann_dir = os.path.join(path, "annotations")
    img_dir = os.path.join(path, "images")
    dataset = []
    for fname in os.listdir(ann_dir):
        if not fname.endswith(".xml"):
            continue
        ann_path = os.path.join(ann_dir, fname)
        tree = ET.parse(ann_path)
        root = tree.getroot()
        bboxes: List[BBox] = []
        for obj in root.findall(".//object"):
            bb = obj.find("bndbox")
            xmin = int(bb.find("xmin").text)
            ymin = int(bb.find("ymin").text)
            xmax = int(bb.find("xmax").text)
            ymax = int(bb.find("ymax").text)
            bboxes.append((xmin, ymin, xmax, ymax))
        image_filename = root.findtext("filename")
        img_path = os.path.join(img_dir, image_filename)
        dataset.append({"image_path": img_path, "bboxes": bboxes, "ann_path": ann_path})
    return dataset


def eob(pred: BBox, gt: BBox) -> float:
    """Compute the Error-of-Boundary metric for a bounding box pair."""
    px0, py0, px1, py1 = pred
    gx0, gy0, gx1, gy1 = gt
    width = gx1 - gx0
    height = gy1 - gy0
    if width <= 0 or height <= 0:
        logger.warning("Invalid ground truth box with non-positive size: %s", gt)
        return float("inf")
    return 0.25 * (
        abs(px0 - gx0) / width +
        abs(px1 - gx1) / width +
        abs(py0 - gy0) / height +
        abs(py1 - gy1) / height
    )


def evaluate_detections(
    pred_boxes: List[BBox],
    gt_boxes: List[BBox],
    threshold: float = 0.0,
) -> Tuple[float, float, float]:
    """Evaluate predicted boxes against ground truth using EoB threshold."""
    matches = 0
    used = set()
    for pb in pred_boxes:
        for idx, gb in enumerate(gt_boxes):
            if idx in used:
                continue
            if eob(pb, gb) <= threshold:
                matches += 1
                used.add(idx)
                break
    precision = matches / len(pred_boxes) if pred_boxes else 0.0
    recall = matches / len(gt_boxes) if gt_boxes else 0.0
    f1 = 2 * precision * recall / (precision + recall) if (precision + recall) else 0.0
    return precision, recall, f1


def normalize_qa_answer(answer: object, answer_type: str = "literal") -> str:
    """Normalize QA answers according to their comparison contract."""
    text = "" if answer is None else str(answer).strip()
    if answer_type == "cell_address":
        return text.upper()
    if answer_type == "formula":
        return re.sub(r"\s+", "", text).upper()
    if answer_type == "free_text":
        return re.sub(r"\s+", " ", text).casefold()
    return text


def load_qa_dataset(path: str) -> List[Dict[str, object]]:
    """
    Load a spreadsheet QA dataset.
    Expects pairs of .xlsx and .json files.
    """
    dataset = []
    for fname in os.listdir(path):
        if fname.endswith(".xlsx"):
            spreadsheet_path = os.path.join(path, fname)
            ann_path = os.path.join(path, fname.replace(".xlsx", ".json"))

            if not os.path.exists(ann_path):
                logger.warning(f"Annotation file not found for {fname}, skipping.")
                continue

            with open(ann_path, 'r') as f:
                annotations = json.load(f)

            qa_pairs = annotations.get("qa_pairs", [])
            if qa_pairs:
                dataset.append({
                    "spreadsheet_path": spreadsheet_path,
                    "qa_pairs": qa_pairs
                })
    return dataset


def load_qa_manifest(manifest_path: str) -> List[Dict[str, object]]:
    """Load a manifest-defined spreadsheet QA dataset.

    Each item must include ``spreadsheet_path`` and ``qa_pairs``. QA pairs may
    include optional fields such as ``sheet_name``, ``table_range``, and
    ``answer_type``; they are preserved for downstream evaluation.
    """
    with open(manifest_path, encoding="utf-8") as fh:
        manifest: Dict[str, Any] = json.load(fh)

    dataset_name = manifest.get("dataset_name", "manifest")
    dataset_version = manifest.get("dataset_version", "unspecified")
    split_name = manifest.get("split_name", "unspecified")
    items = manifest.get("items")
    if not isinstance(items, list):
        raise ValueError("QA manifest must contain an items array")

    dataset = []
    for idx, item in enumerate(items):
        if not isinstance(item, dict):
            raise ValueError(f"manifest item {idx} must be an object")
        spreadsheet_path = item.get("spreadsheet_path")
        qa_pairs = item.get("qa_pairs", [])
        if not spreadsheet_path:
            raise ValueError(f"manifest item {idx} missing spreadsheet_path")
        if not isinstance(qa_pairs, list):
            raise ValueError(f"manifest item {idx} qa_pairs must be an array")
        dataset.append({
            "spreadsheet_path": _resolve_manifest_path(manifest_path, spreadsheet_path),
            "qa_pairs": qa_pairs,
            "manifest_path": os.path.abspath(manifest_path),
            "dataset_name": dataset_name,
            "dataset_version": dataset_version,
            "split_name": split_name,
        })
    return dataset

__all__ = [
    "load_dong2019_dataset",
    "load_spreadsheet_dataset",
    "load_table_detection_manifest",
    "load_qa_dataset",
    "load_qa_manifest",
    "range_to_bbox",
    "normalize_qa_answer",
    "eob",
    "evaluate_detections",
]
