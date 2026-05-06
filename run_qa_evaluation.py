import argparse
import datetime as _dt
import json
import logging
import os
from typing import List, Dict, Optional

import chain_of_spreadsheet
from evaluation import load_qa_dataset
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode
from chain_of_spreadsheet import identify_table, table_split_qa

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def run_tape_placeholder(encoding: Dict, query: str) -> str:
    """Placeholder for TaPEx baseline.

    # TODO: Real TaPEx baseline requires separate integration work; this is
    # intentionally a stub flagged by the gap analysis.
    """
    logger.info("Running placeholder TaPEx baseline...")
    return "[C5]"  # Dummy response


def run_binder_placeholder(encoding: Dict, query: str) -> str:
    """Placeholder for Binder baseline.

    # TODO: Real Binder baseline requires separate integration work; this is
    # intentionally a stub flagged by the gap analysis.
    """
    logger.info("Running placeholder Binder baseline...")
    return "[SUM(D1:D5)]"  # Dummy response


def main(
    dataset_dir: str,
    k: int,
    out_record: Optional[str] = None,
    backend_name: str = "unknown",
    extra_meta: Optional[Dict] = None,
):
    """Main function to run the Spreadsheet QA evaluation.

    When ``out_record`` is given, also persist a structured JSON record
    (timestamp, dataset, k, backend, per-question results, accuracies)
    so the spreadsheet-llm-fidelity skill_runs log can track QA accuracy
    iteration-over-iteration without log-string parsing.
    """
    dataset = load_qa_dataset(dataset_dir)

    if not dataset:
        logger.error("No QA data found in the specified dataset directory.")
        return

    total_questions = 0
    correct_spreadsheetllm = 0
    correct_tape = 0
    correct_binder = 0
    per_question: List[Dict] = []

    for item in dataset:
        spreadsheet_path = item["spreadsheet_path"]
        qa_pairs = item["qa_pairs"]

        logger.info("\n--- Processing %s ---", spreadsheet_path)

        encoding = spreadsheet_llm_encode(spreadsheet_path, k=k)
        if not encoding:
            continue

        for qa in qa_pairs:
            total_questions += 1
            query = qa["question"]
            ground_truth = qa["answer"]

            logger.info("Q: %s (GT: %s)", query, ground_truth)

            pred_answer_llm: Optional[str] = None
            llm_correct = False

            # --- SpreadsheetLLM Evaluation ---
            try:
                table_range = identify_table(encoding, query)
            except NotImplementedError:
                logger.error(
                    "No LLM backend configured. Run with --backend openai "
                    "or chain_of_spreadsheet.configure_backend(...)."
                )
                table_range = None

            if table_range:
                # find_relevant_sheet picks the sheet; fall back to first sheet
                sheet_name: Optional[str] = chain_of_spreadsheet.find_relevant_sheet(
                    encoding, query
                )
                if sheet_name is None:
                    sheet_name = next(iter(encoding["sheets"]))
                sheet_data = encoding["sheets"][sheet_name]

                pred_answer_llm = table_split_qa(
                    sheet_data,
                    table_range,
                    query,
                    workbook_path=spreadsheet_path,
                    sheet_name=sheet_name,
                    coord_map=sheet_data.get("coord_map"),
                )

                logger.info("  - SpreadsheetLLM Predicted: %s", pred_answer_llm)
                if pred_answer_llm.strip() == ground_truth.strip():
                    correct_spreadsheetllm += 1
                    llm_correct = True
            else:
                logger.warning("  - SpreadsheetLLM could not identify a relevant table.")

            # --- Baseline Evaluations (placeholder) ---
            pred_answer_tape = run_tape_placeholder(encoding, query)
            logger.info("  - TaPEx Predicted: %s", pred_answer_tape)
            tape_correct = pred_answer_tape.strip() == ground_truth.strip()
            if tape_correct:
                correct_tape += 1

            pred_answer_binder = run_binder_placeholder(encoding, query)
            logger.info("  - Binder Predicted: %s", pred_answer_binder)
            binder_correct = pred_answer_binder.strip() == ground_truth.strip()
            if binder_correct:
                correct_binder += 1

            per_question.append({
                "spreadsheet_path": spreadsheet_path,
                "question": query,
                "ground_truth": ground_truth,
                "spreadsheetllm": {
                    "predicted": pred_answer_llm,
                    "correct": llm_correct,
                },
                "tapex_placeholder": {
                    "predicted": pred_answer_tape,
                    "correct": tape_correct,
                },
                "binder_placeholder": {
                    "predicted": pred_answer_binder,
                    "correct": binder_correct,
                },
            })

    # --- Report Results ---
    logger.info("\n--- QA Evaluation Summary ---")
    acc_llm = acc_tape = acc_binder = 0.0
    if total_questions > 0:
        acc_llm = (correct_spreadsheetllm / total_questions) * 100
        acc_tape = (correct_tape / total_questions) * 100
        acc_binder = (correct_binder / total_questions) * 100

        logger.info(
            "SpreadsheetLLM Accuracy: %.2f%% (%d/%d)",
            acc_llm, correct_spreadsheetllm, total_questions,
        )
        logger.info(
            "TaPEx Baseline Accuracy: %.2f%% (%d/%d)",
            acc_tape, correct_tape, total_questions,
        )
        logger.info(
            "Binder Baseline Accuracy: %.2f%% (%d/%d)",
            acc_binder, correct_binder, total_questions,
        )
    else:
        logger.info("No questions were evaluated.")
    logger.info("--------------------------")

    if out_record:
        record = {
            "timestamp": _dt.datetime.now(_dt.timezone.utc).isoformat(),
            "task": "spreadsheet_qa",
            "dataset_dir": os.path.abspath(dataset_dir),
            "k": k,
            "backend": backend_name,
            "n_questions": total_questions,
            "spreadsheetllm_accuracy_pct": acc_llm,
            "tapex_placeholder_accuracy_pct": acc_tape,
            "binder_placeholder_accuracy_pct": acc_binder,
            "baselines_are_placeholders": True,
            "per_question": per_question,
            "meta": extra_meta or {},
        }
        out_path = os.path.abspath(out_record)
        os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as fh:
            json.dump(record, fh, indent=2)
        logger.info("Wrote QA evaluation record to %s", out_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run Spreadsheet QA evaluation.")
    parser.add_argument("dataset_dir", help="Path to the QA dataset directory")
    parser.add_argument(
        "--k", type=int, default=4,
        help="Neighborhood distance for structural anchors (default: 4)"
    )
    parser.add_argument(
        "--backend", choices=["openai", "echo"], default="echo",
        help="LLM backend to use (default: echo)"
    )
    parser.add_argument(
        "--openai-model", default="gpt-4o-mini",
        help="OpenAI model name when --backend=openai (default: gpt-4o-mini)"
    )
    parser.add_argument(
        "--echo-response", default="['range': 'A1:D5']",
        help="Fixed response for the echo backend (default: a sample range so "
             "Stage 1 returns something parseable).",
    )
    parser.add_argument(
        "--out-record", default=None,
        help="Optional path. When set, write a structured JSON record of "
             "the QA evaluation (timestamp, k, backend, per-question results, "
             "accuracies) so it can be diffed across iterations.",
    )
    args = parser.parse_args()

    from llm_backend import EchoBackend, OpenAIBackend

    if args.backend == "openai":
        backend = OpenAIBackend(model=args.openai_model)
        backend_name = f"openai:{args.openai_model}"
    else:
        backend = EchoBackend(response=args.echo_response)
        backend_name = "echo"
    # Always configure the backend so identify_table / table_split_qa don't
    # raise NotImplementedError under --backend=echo (previously a bug).
    chain_of_spreadsheet.configure_backend(backend)

    main(
        args.dataset_dir,
        args.k,
        out_record=args.out_record,
        backend_name=backend_name,
    )
