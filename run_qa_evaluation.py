import argparse
import logging
from typing import List, Dict

from evaluation import load_qa_dataset
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode
from chain_of_spreadsheet import identify_table, generate_response, table_split_qa

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def run_tape_baseline(encoding: Dict, query: str) -> str:
    """Placeholder for TaPEx baseline."""
    logger.info("Running placeholder TaPEx baseline...")
    return "[C5]" # Dummy response

def run_binder_baseline(encoding: Dict, query: str) -> str:
    """Placeholder for Binder baseline."""
    logger.info("Running placeholder Binder baseline...")
    return "[SUM(D1:D5)]" # Dummy response

def main(dataset_dir: str, k: int):
    """Main function to run the Spreadsheet QA evaluation."""
    dataset = load_qa_dataset(dataset_dir)

    if not dataset:
        logger.error("No QA data found in the specified dataset directory.")
        return

    total_questions = 0
    correct_spreadsheetllm = 0
    correct_tape = 0
    correct_binder = 0

    for item in dataset:
        spreadsheet_path = item["spreadsheet_path"]
        qa_pairs = item["qa_pairs"]

        logger.info(f"\n--- Processing {spreadsheet_path} ---")

        encoding = spreadsheet_llm_encode(spreadsheet_path, k=k)
        if not encoding:
            continue

        for qa in qa_pairs:
            total_questions += 1
            query = qa["question"]
            ground_truth = qa["answer"]

            logger.info(f"Q: {query} (GT: {ground_truth})")

            # --- SpreadsheetLLM Evaluation ---
            table_range = identify_table(encoding, query)
            if table_range:
                # In a real scenario, we'd pass the specific sheet data
                sheet_name = next(iter(encoding["sheets"])) # Assume first sheet
                sheet_data = encoding["sheets"][sheet_name]

                # Using table_split_qa to handle large tables
                pred_answer_llm = table_split_qa(sheet_data, table_range, query)

                logger.info(f"  - SpreadsheetLLM Predicted: {pred_answer_llm}")
                if pred_answer_llm.strip() == ground_truth.strip():
                    correct_spreadsheetllm += 1
            else:
                logger.warning("  - SpreadsheetLLM could not identify a relevant table.")

            # --- Baseline Evaluations ---
            pred_answer_tape = run_tape_baseline(encoding, query)
            logger.info(f"  - TaPEx Predicted: {pred_answer_tape}")
            if pred_answer_tape.strip() == ground_truth.strip():
                correct_tape += 1

            pred_answer_binder = run_binder_baseline(encoding, query)
            logger.info(f"  - Binder Predicted: {pred_answer_binder}")
            if pred_answer_binder.strip() == ground_truth.strip():
                correct_binder += 1

    # --- Report Results ---
    logger.info("\n--- QA Evaluation Summary ---")
    if total_questions > 0:
        acc_llm = (correct_spreadsheetllm / total_questions) * 100
        acc_tape = (correct_tape / total_questions) * 100
        acc_binder = (correct_binder / total_questions) * 100

        logger.info(f"SpreadsheetLLM Accuracy: {acc_llm:.2f}% ({correct_spreadsheetllm}/{total_questions})")
        logger.info(f"TaPEx Baseline Accuracy: {acc_tape:.2f}% ({correct_tape}/{total_questions})")
        logger.info(f"Binder Baseline Accuracy: {acc_binder:.2f}% ({correct_binder}/{total_questions})")
    else:
        logger.info("No questions were evaluated.")
    logger.info("--------------------------")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run Spreadsheet QA evaluation.")
    parser.add_argument("dataset_dir", help="Path to the QA dataset directory")
    parser.add_argument("--k", type=int, default=2, help="Neighborhood distance for structural anchors (default: 2)")
    args = parser.parse_args()
    main(args.dataset_dir, args.k)
