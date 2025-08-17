import argparse
import json
import logging
import re
from typing import List, Dict

from evaluation import load_spreadsheet_dataset, evaluate_detections, range_to_bbox, BBox
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- Placeholder for a real LLM client ---
def _call_llm(prompt: str) -> str:
    """A placeholder function to simulate an LLM call for table detection."""
    logger.warning("Using placeholder LLM. This will not produce real results.")
    # Simulate identifying a table
    return "['range': 'A1:F9', 'range': 'A12:F18']"  # Hardcoded response

TABLE_DETECTION_PROMPT_TEMPLATE = """
INSTRUCTION:
Given an input that is a string denoting data of cells in an Excel spreadsheet. The input spreadsheet contains many tuples, describing the cells with content in the spreadsheet. Each tuple consists of two elements separated by a '|': the cell content and the cell address/region, like (Year|A1), ( |A1) or (IntNum|A1:B3). The content in some cells such as '#,##0'/'d-mmm-yy'/'H:mm:ss',etc., represents the CELL DATA FORMATS of Excel. The content in some cells such as 'IntNum'/'DateData'/'EmailData',etc., represents a category of data with the same format and similar semantics. For example, 'IntNum' represents integer type data, and 'ScientificNum' represents scientific notation type data. 'A1:B3' represents a region in a spreadsheet, from the first row to the third row and from column A to column B. Some cells with empty content in the spreadsheet are not entered. Now you should tell me the range of the table in a format like A2:D5, and the range of the table should only CONTAIN HEADER REGION and the data region. DON’T include the title or comments. Note that there can be more than one table in a string, so you should return all the RANGE. DON’T ADD OTHER WORDS OR EXPLANATION.

INPUT:
[Encoded Spreadsheet]
"""

def predict_tables_with_llm(encoding: Dict) -> List[BBox]:
    """
    Predicts table boundaries in a spreadsheet using an LLM.
    """
    # We'll evaluate one sheet at a time for simplicity
    predicted_boxes = []
    for sheet_name, sheet_data in encoding.get("sheets", {}).items():
        prompt_input = json.dumps(sheet_data, ensure_ascii=False)
        prompt = TABLE_DETECTION_PROMPT_TEMPLATE.replace("[Encoded Spreadsheet]", prompt_input)

        llm_response = _call_llm(prompt)

        # Parse ranges like 'A1:F9' from the response
        ranges = re.findall(r"\'([A-Z]+\d+:[A-Z]+\d+)\'", llm_response)
        for r in ranges:
            try:
                predicted_boxes.append(range_to_bbox(r))
            except Exception as e:
                logger.error(f"Could not parse range '{r}' from LLM response: {e}")

    return predicted_boxes


def main(dataset_dir: str, k: int):
    """Main function to run the LLM-based table detection evaluation."""
    data = load_spreadsheet_dataset(dataset_dir)
    total_f1 = 0.0

    if not data:
        logger.error("No data found in the specified dataset directory.")
        return

    for item in data:
        logger.info(f"Processing {item['spreadsheet_path']}...")

        # 1. Encode the spreadsheet
        encoding = spreadsheet_llm_encode(item["spreadsheet_path"], k=k)
        if not encoding:
            continue

        # 2. Predict tables with LLM
        pred_boxes = predict_tables_with_llm(encoding)

        # 3. Evaluate
        gt_boxes = item["bboxes"]
        precision, recall, f1 = evaluate_detections(pred_boxes, gt_boxes)

        logger.info(f"  GT: {len(gt_boxes)} tables, Pred: {len(pred_boxes)} tables")
        logger.info(f"  Precision: {precision:.4f}, Recall: {recall:.4f}, F1: {f1:.4f}")

        total_f1 += f1

    avg_f1 = total_f1 / len(data) if data else 0.0
    logger.info("\n---------------------------------")
    logger.info(f"Average F1 Score (EoB-0): {avg_f1:.4f}")
    logger.info("---------------------------------")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run LLM-based table detection evaluation.")
    parser.add_argument("dataset_dir", help="Path to the spreadsheet dataset directory")
    parser.add_argument("--k", type=int, default=2, help="Neighborhood distance for structural anchors (default: 2)")
    args = parser.parse_args()
    main(args.dataset_dir, args.k)
