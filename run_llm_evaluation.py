import argparse
import datetime as _dt
import logging
import os
import re
from typing import List, Dict, Optional

import paper_serializers
from evaluation import (
    load_spreadsheet_dataset,
    load_table_detection_manifest,
    evaluate_detections,
    range_to_bbox,
    BBox,
)
from evaluation_metadata import build_evaluation_metadata, write_evaluation_record
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

_RANGE_RE = re.compile(r"""(?:['"]([A-Z]+\d+:[A-Z]+\d+)['"]|\b([A-Z]+\d+:[A-Z]+\d+)\b)""")


def extract_ranges(text: str) -> List[str]:
    """Extract quoted or bare Excel ranges from an LLM response."""
    ranges: List[str] = []
    seen = set()
    for match in _RANGE_RE.finditer(text or ""):
        rng = match.group(1) or match.group(2)
        if rng not in seen:
            ranges.append(rng)
            seen.add(rng)
    return ranges

TABLE_DETECTION_PROMPT_TEMPLATE = """
INSTRUCTION:
Given an input that is a string denoting data of cells in an Excel spreadsheet. The input spreadsheet contains many tuples, describing the cells with content in the spreadsheet. Each tuple consists of two elements separated by a '|': the cell content and the cell address/region, like (Year|A1), ( |A1) or (IntNum|A1:B3). The content in some cells such as '#,##0'/'d-mmm-yy'/'H:mm:ss',etc., represents the CELL DATA FORMATS of Excel. The content in some cells such as 'IntNum'/'DateData'/'EmailData',etc., represents a category of data with the same format and similar semantics. For example, 'IntNum' represents integer type data, and 'ScientificNum' represents scientific notation type data. 'A1:B3' represents a region in a spreadsheet, from the first row to the third row and from column A to column B. Some cells with empty content in the spreadsheet are not entered. Now you should tell me the range of the table in a format like A2:D5, and the range of the table should only CONTAIN HEADER REGION and the data region. DON'T include the title or comments. Note that there can be more than one table in a string, so you should return all the RANGE. DON'T ADD OTHER WORDS OR EXPLANATION.

INPUT:
[Encoded Spreadsheet]
"""


def predict_tables_with_llm(encoding: Dict, llm_callable) -> List[BBox]:
    """
    Predicts table boundaries in a spreadsheet using an LLM.
    """
    predicted_boxes = []
    for sheet_name, sheet_data in encoding.get("sheets", {}).items():
        coord_map = sheet_data.get("coord_map")
        prompt_input = paper_serializers.to_paper_compressed_prompt(
            sheet_data, coord_map=coord_map
        )
        prompt = TABLE_DETECTION_PROMPT_TEMPLATE.replace("[Encoded Spreadsheet]", prompt_input)

        llm_response = llm_callable(prompt)

        # Parse ranges like A1:F9, 'A1:F9', or "A1:F9" from the response.
        ranges = extract_ranges(llm_response)
        for r in ranges:
            if coord_map:
                unmapped = paper_serializers.unremap_range(r, coord_map)
                if unmapped is None:
                    logger.warning("Could not unremap range '%s'; skipping.", r)
                    continue
                r = unmapped
            try:
                predicted_boxes.append(range_to_bbox(r))
            except Exception as e:
                logger.error("Could not parse range '%s' from LLM response: %s", r, e)

    return predicted_boxes


def main(
    dataset_dir: Optional[str],
    k: int,
    llm_callable,
    out_record: Optional[str] = None,
    backend_name: str = "unknown",
    extra_meta: Optional[Dict] = None,
    manifest_path: Optional[str] = None,
):
    """Main function to run the LLM-based table detection evaluation.

    When ``out_record`` is given, also write a structured JSON record with
    timestamp, dataset, k, backend, per-item F1, and average F1 — feeds the
    spreadsheet-llm-fidelity skill_runs log without log-string parsing.
    """
    if not dataset_dir and not manifest_path:
        raise ValueError("dataset_dir is required unless manifest_path is provided")

    data = (
        load_table_detection_manifest(manifest_path)
        if manifest_path
        else load_spreadsheet_dataset(dataset_dir)
    )
    total_f1 = 0.0
    per_item: List[Dict] = []

    if not data:
        logger.error("No data found in the specified dataset directory.")
        return

    for item in data:
        logger.info("Processing %s...", item["spreadsheet_path"])

        # 1. Encode the spreadsheet
        encoding = spreadsheet_llm_encode(item["spreadsheet_path"], k=k)
        if not encoding:
            continue

        # 2. Predict tables with LLM
        pred_boxes = predict_tables_with_llm(encoding, llm_callable)

        # 3. Evaluate
        gt_boxes = item["bboxes"]
        precision, recall, f1 = evaluate_detections(pred_boxes, gt_boxes)

        logger.info("  GT: %d tables, Pred: %d tables", len(gt_boxes), len(pred_boxes))
        logger.info("  Precision: %.4f, Recall: %.4f, F1: %.4f", precision, recall, f1)

        total_f1 += f1
        per_item.append({
            "spreadsheet_path": item["spreadsheet_path"],
            "gt_count": len(gt_boxes),
            "pred_count": len(pred_boxes),
            "precision": precision,
            "recall": recall,
            "f1": f1,
        })

    avg_f1 = total_f1 / len(data) if data else 0.0
    logger.info("\n---------------------------------")
    logger.info("Average F1 Score (EoB-0): %.4f", avg_f1)
    logger.info("---------------------------------")

    if out_record:
        table_count = sum(len(item["bboxes"]) for item in data)
        dataset_meta = data[0] if data else {}
        evaluation_metadata = build_evaluation_metadata(
            dataset_dir=manifest_path or dataset_dir,
            task="table_detection",
            dataset_name=dataset_meta.get("dataset_name"),
            dataset_version=dataset_meta.get("dataset_version", "unspecified"),
            claim_level=dataset_meta.get("claim_level", "synthetic"),
            split_name=dataset_meta.get("split_name", "unspecified"),
            spreadsheet_count=len(data),
            table_count=table_count,
            qa_item_count=0,
            encoder_settings={"k": k},
            prompt_serializer="paper_serializers.to_paper_compressed_prompt",
            coordinate_mode="compact_prompt_unmapped_to_original_for_eob0",
            model_backend=backend_name,
            metric_definition="EoB-0 exact boundary matching; threshold=0.0",
            baseline_name="SpreadsheetLLM table detection",
            baseline_status="run",
            skip_reasons=[],
        )
        record = {
            "timestamp": _dt.datetime.now(_dt.timezone.utc).isoformat(),
            "task": "table_detection_eob0",
            "dataset_dir": os.path.abspath(dataset_dir) if dataset_dir else None,
            "manifest_path": os.path.abspath(manifest_path) if manifest_path else None,
            "k": k,
            "backend": backend_name,
            "n_items": len(data),
            "avg_f1_eob0": avg_f1,
            "per_item": per_item,
            "meta": extra_meta or {},
            "evaluation_metadata": evaluation_metadata,
        }
        write_evaluation_record(record, out_record)
        logger.info("Wrote evaluation record to %s", os.path.abspath(out_record))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run LLM-based table detection evaluation.")
    parser.add_argument(
        "dataset_dir",
        nargs="?",
        default=None,
        help="Path to the spreadsheet dataset directory. Optional when --manifest is provided.",
    )
    parser.add_argument(
        "--k", type=int, default=4,
        help="Neighborhood distance for structural anchors (default: 4)"
    )
    parser.add_argument(
        "--backend", choices=["openai", "echo"], default="echo",
        help="LLM backend to use (default: echo)"
    )
    parser.add_argument(
        "--echo-response", default="",
        help="Fixed response string returned by the echo backend (default: empty)"
    )
    parser.add_argument(
        "--openai-model", default="gpt-4o-mini",
        help="OpenAI model name when --backend=openai (default: gpt-4o-mini)"
    )
    parser.add_argument(
        "--out-record", default=None,
        help="Optional path. When set, write a structured JSON record of "
             "the evaluation (timestamp, k, backend, per-item F1, avg F1).",
    )
    parser.add_argument(
        "--manifest",
        default=None,
        help="Optional table-detection manifest JSON. When set, it overrides dataset_dir scanning.",
    )
    args = parser.parse_args()
    if not args.dataset_dir and not args.manifest:
        parser.error("dataset_dir is required unless --manifest is provided")

    from llm_backend import EchoBackend, OpenAIBackend

    if args.backend == "openai":
        llm_callable = OpenAIBackend(model=args.openai_model)
        backend_name = f"openai:{args.openai_model}"
    else:
        llm_callable = EchoBackend(response=args.echo_response)
        backend_name = "echo"

    main(
        args.dataset_dir,
        args.k,
        llm_callable,
        out_record=args.out_record,
        backend_name=backend_name,
        manifest_path=args.manifest,
    )
