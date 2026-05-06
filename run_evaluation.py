import argparse
import json
import logging
import os
from typing import Optional

from evaluation import load_dong2019_dataset, evaluate_detections
from evaluation_metadata import build_evaluation_metadata, write_evaluation_record

try:
    from tablesense_cnn import TableSenseCNN
except Exception:  # pragma: no cover - model might not be available
    TableSenseCNN = None

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def main(dataset_dir: str, out_record: Optional[str] = None):
    if TableSenseCNN is None:
        logger.error("TableSense-CNN package is not installed.")
        if out_record:
            record = {
                "task": "tablesense_cnn_detection",
                "dataset_dir": os.path.abspath(dataset_dir),
                "skipped": True,
                "evaluation_metadata": build_evaluation_metadata(
                    dataset_dir=dataset_dir,
                    task="table_detection",
                    spreadsheet_count=0,
                    table_count=0,
                    qa_item_count=0,
                    encoder_settings={},
                    prompt_serializer="image_baseline_not_applicable",
                    coordinate_mode="image_bounding_boxes",
                    model_backend="tablesense-cnn",
                    metric_definition="EoB-0 exact boundary matching; threshold=0.0",
                    baseline_name="TableSense-CNN",
                    skip_reasons=[{
                        "component": "tablesense-cnn",
                        "reason": "TableSense-CNN package is not installed.",
                    }],
                ),
            }
            write_evaluation_record(record, out_record)
        return

    data = load_dong2019_dataset(dataset_dir)
    model = TableSenseCNN.pretrained()

    total_f1 = 0.0
    total_ann_size = 0
    total_pred_size = 0

    for item in data:
        preds = model.predict_tables(item["image_path"])
        _, _, f1 = evaluate_detections(preds, item["bboxes"])
        total_f1 += f1

        with open(item["ann_path"], "rb") as f:
            ann_bytes = f.read()
            total_ann_size += len(ann_bytes)
        pred_json = json.dumps(preds).encode("utf-8")
        total_pred_size += len(pred_json)

    avg_f1 = total_f1 / len(data) if data else 0.0
    compression_ratio = total_pred_size / total_ann_size if total_ann_size else 0.0

    logger.info("Average F1 (EoB-0): %.4f", avg_f1)
    logger.info("Bounding box compression ratio: %.4f", compression_ratio)

    if out_record:
        table_count = sum(len(item["bboxes"]) for item in data)
        record = {
            "task": "tablesense_cnn_detection",
            "dataset_dir": os.path.abspath(dataset_dir),
            "n_items": len(data),
            "avg_f1_eob0": avg_f1,
            "bbox_compression_ratio": compression_ratio,
            "evaluation_metadata": build_evaluation_metadata(
                dataset_dir=dataset_dir,
                task="table_detection",
                spreadsheet_count=len(data),
                table_count=table_count,
                qa_item_count=0,
                encoder_settings={},
                prompt_serializer="image_baseline_not_applicable",
                coordinate_mode="image_bounding_boxes",
                model_backend="tablesense-cnn",
                metric_definition="EoB-0 exact boundary matching; threshold=0.0",
                baseline_name="TableSense-CNN",
            ),
        }
        write_evaluation_record(record, out_record)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run TableSense-CNN evaluation.")
    parser.add_argument("dataset_dir", help="Path to Dong et al. 2019 dataset")
    parser.add_argument(
        "--out-record",
        default=None,
        help="Optional path for a structured JSON evaluation record.",
    )
    args = parser.parse_args()
    main(args.dataset_dir, out_record=args.out_record)
