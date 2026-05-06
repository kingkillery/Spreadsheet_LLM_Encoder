import argparse
import json
import logging
from typing import List, Dict, Optional

import paper_serializers
from evaluation import load_spreadsheet_dataset, BBox
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TABLE_DETECTION_PROMPT_TEMPLATE = """
INSTRUCTION:
Given an input that is a string denoting data of cells in an Excel spreadsheet. The input spreadsheet contains many tuples, describing the cells with content in the spreadsheet. Each tuple consists of two elements separated by a '|': the cell content and the cell address/region, like (Year|A1), ( |A1) or (IntNum|A1:B3). The content in some cells such as '#,##0'/'d-mmm-yy'/'H:mm:ss',etc., represents the CELL DATA FORMATS of Excel. The content in some cells such as 'IntNum'/'DateData'/'EmailData',etc., represents a category of data with the same format and similar semantics. For example, 'IntNum' represents integer type data, and 'ScientificNum' represents scientific notation type data. 'A1:B3' represents a region in a spreadsheet, from the first row to the third row and from column A to column B. Some cells with empty content in the spreadsheet are not entered. Now you should tell me the range of the table in a format like A2:D5, and the range of the table should only CONTAIN HEADER REGION and the data region. DON'T include the title or comments. Note that there can be more than one table in a string, so you should return all the RANGE. DON'T ADD OTHER WORDS OR EXPLANATION.

INPUT:
[Encoded Spreadsheet]
"""


def bbox_to_range(bbox: BBox) -> str:
    """Converts a BBox tuple to an Excel-style range string."""
    r1, c1, r2, c2 = bbox
    start_cell = f"{get_column_letter(c1)}{r1}"
    end_cell = f"{get_column_letter(c2)}{r2}"
    if start_cell == end_cell:
        return start_cell
    return f"{start_cell}:{end_cell}"


def bbox_to_prompt_range(bbox: BBox, coord_map: Optional[Dict] = None) -> Optional[str]:
    """Convert a ground-truth box to the coordinate space shown in the prompt."""
    original_range = bbox_to_range(bbox)
    if not coord_map:
        return original_range
    return paper_serializers.remap_range(original_range, coord_map)


def format_for_finetuning(encoding: Dict, gt_boxes: List[BBox]) -> List[Dict]:
    """
    Formats the encoded spreadsheet and ground truth into a list of dicts,
    one per sheet, suitable for fine-tuning (e.g., as JSONL lines).
    """
    records: List[Dict] = []
    for sheet_name, sheet_data in encoding.get("sheets", {}).items():
        coord_map = sheet_data.get("coord_map")
        prompt_input = paper_serializers.to_paper_compressed_prompt(
            sheet_data, coord_map=coord_map
        )
        prompt = TABLE_DETECTION_PROMPT_TEMPLATE.replace("[Encoded Spreadsheet]", prompt_input)

        gt_ranges = []
        for bbox in gt_boxes:
            prompt_range = bbox_to_prompt_range(bbox, coord_map)
            if prompt_range is None:
                logger.warning(
                    "Skipping ground-truth box %s because it is not present in "
                    "the compressed prompt coordinate map for sheet %s.",
                    bbox,
                    sheet_name,
                )
                continue
            gt_ranges.append(prompt_range)
        range_parts = ["'range': '" + r + "'" for r in gt_ranges]
        completion = "[" + ", ".join(range_parts) + "]"

        records.append({"prompt": prompt, "completion": completion})

    return records


def push_to_hub(
    output_path: str,
    repo_id: str,
    hub_split: str,
    hub_private: bool,
) -> None:
    """Uploads the JSONL at *output_path* to HuggingFace Hub as a dataset."""
    try:
        import datasets as _datasets  # noqa: PLC0415
    except ImportError as exc:
        raise ImportError(
            "Install with `pip install datasets` to push to the HuggingFace Hub."
        ) from exc

    with open(output_path, encoding="utf-8") as fh:
        records = [json.loads(line) for line in fh if line.strip()]

    ds = _datasets.Dataset.from_list(records)
    ds.push_to_hub(repo_id, split=hub_split, private=hub_private)
    logger.info(
        "Pushed %d rows to https://huggingface.co/datasets/%s (split=%s, private=%s)",
        len(records),
        repo_id,
        hub_split,
        hub_private,
    )


def main(
    dataset_dir: str,
    output_path: str,
    k: int,
    push_to_hub_repo: Optional[str] = None,
    hub_split: str = "train",
    hub_private: bool = True,
) -> None:
    """
    Main function to prepare data for fine-tuning.
    """
    data = load_spreadsheet_dataset(dataset_dir)

    if not data:
        logger.error("No data found in the specified dataset directory.")
        return

    with open(output_path, "w", encoding="utf-8") as f:
        for item in data:
            logger.info("Processing %s for fine-tuning...", item["spreadsheet_path"])

            # 1. Encode the spreadsheet
            encoding = spreadsheet_llm_encode(item["spreadsheet_path"], k=k)
            if not encoding or not encoding.get("sheets"):
                logger.warning(
                    "Skipping %s due to encoding error or empty sheets.",
                    item["spreadsheet_path"],
                )
                continue

            # 2. Format for fine-tuning (one record per sheet)
            ft_records = format_for_finetuning(encoding, item["bboxes"])

            # 3. Write each record as a JSONL line
            for record in ft_records:
                f.write(json.dumps(record) + "\n")

    logger.info("Fine-tuning data successfully prepared and saved to %s", output_path)

    if push_to_hub_repo:
        push_to_hub(output_path, push_to_hub_repo, hub_split, hub_private)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Prepare spreadsheet table detection data for fine-tuning."
    )
    parser.add_argument("dataset_dir", help="Path to the spreadsheet dataset directory")
    parser.add_argument("output_path", help="Path to save the output JSONL file for fine-tuning")
    parser.add_argument(
        "--k", type=int, default=4,
        help="Neighborhood distance for structural anchors (default: 4)"
    )
    parser.add_argument(
        "--push-to-hub",
        metavar="REPO_ID",
        default=None,
        help=(
            "HuggingFace Hub dataset repo to push results to, e.g. "
            "username/spreadsheetllm-finetune-v1. "
            "Requires `pip install datasets`. "
            "Authenticate via HF_TOKEN env var or `huggingface-cli login`."
        ),
    )
    parser.add_argument(
        "--hub-private",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Whether the Hub dataset repo should be private (default: True).",
    )
    parser.add_argument(
        "--hub-split",
        default="train",
        help="Dataset split name to push (default: train).",
    )
    args = parser.parse_args()
    main(
        args.dataset_dir,
        args.output_path,
        args.k,
        push_to_hub_repo=args.push_to_hub,
        hub_split=args.hub_split,
        hub_private=args.hub_private,
    )
