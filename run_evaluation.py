import argparse
import json
import logging
import os

from evaluation import load_dong2019_dataset, evaluate_detections

try:
    from tablesense_cnn import TableSenseCNN
except Exception:  # pragma: no cover - model might not be available
    TableSenseCNN = None

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def main(dataset_dir: str):
    data = load_dong2019_dataset(dataset_dir)
    if TableSenseCNN is None:
        logger.error("TableSense-CNN package is not installed.")
        return

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


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run TableSense-CNN evaluation.")
    parser.add_argument("dataset_dir", help="Path to Dong et al. 2019 dataset")
    args = parser.parse_args()
    main(args.dataset_dir)
