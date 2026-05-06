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
    """Placeholder TaPEx fallback.

    Used when --real-tapex is off or transformers isn't installed. Real TaPEx
    requires the workbook path and identified table range; the placeholder
    only sees the encoding, which is why the real path is wired into ``main``
    rather than this function.
    """
    logger.info("Running placeholder TaPEx baseline...")
    return "[C5]"  # Dummy response


def run_tape_real(
    tapex,  # baselines.TaPExBaseline
    workbook_path: str,
    sheet_name: Optional[str],
    table_range: Optional[str],
    query: str,
) -> str:
    """Invoke a real TaPEx model on the identified table sub-range.

    Falls back to ``[]`` when the SpreadsheetLLM Stage 1 didn't produce a
    table_range / sheet_name (TaPEx needs both).
    """
    if not table_range or not sheet_name:
        return "[]"
    try:
        return tapex.answer(workbook_path, sheet_name, table_range, query)
    except Exception as exc:  # pragma: no cover - defensive
        logger.warning("TaPEx baseline raised %s; recording empty answer.", exc)
        return "[]"


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
    tapex=None,
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
            tapex_table_range: Optional[str] = None
            tapex_sheet_name: Optional[str] = None

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

                # The real TaPEx baseline needs the same identified
                # (sheet, table_range), but in original-workbook coords.
                tapex_sheet_name = sheet_name
                cm = sheet_data.get("coord_map")
                if cm:
                    from paper_serializers import unremap_range
                    tapex_table_range = unremap_range(table_range, cm) or table_range
                else:
                    tapex_table_range = table_range
            else:
                logger.warning("  - SpreadsheetLLM could not identify a relevant table.")

            # --- Baseline Evaluations ---
            if tapex is not None:
                pred_answer_tape = run_tape_real(
                    tapex, spreadsheet_path,
                    tapex_sheet_name, tapex_table_range, query,
                )
                tapex_kind = "real"
            else:
                pred_answer_tape = run_tape_placeholder(encoding, query)
                tapex_kind = "placeholder"
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
                f"tapex_{tapex_kind}": {
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
        tapex_real = tapex is not None
        record = {
            "timestamp": _dt.datetime.now(_dt.timezone.utc).isoformat(),
            "task": "spreadsheet_qa",
            "dataset_dir": os.path.abspath(dataset_dir),
            "k": k,
            "backend": backend_name,
            "n_questions": total_questions,
            "spreadsheetllm_accuracy_pct": acc_llm,
            "tapex_accuracy_pct": acc_tape,
            "tapex_kind": "real" if tapex_real else "placeholder",
            "binder_placeholder_accuracy_pct": acc_binder,
            "baselines_are_placeholders": not tapex_real,  # Binder still placeholder
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
    parser.add_argument(
        "--real-tapex", action="store_true",
        help="Use the real TaPEx baseline (microsoft/tapex-base-finetuned-wtq "
             "via HuggingFace transformers). Requires `pip install transformers`. "
             "If omitted, the placeholder is used.",
    )
    parser.add_argument(
        "--tapex-model", default="microsoft/tapex-base-finetuned-wtq",
        help="HF model id for --real-tapex (default: microsoft/tapex-base-finetuned-wtq).",
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

    tapex_instance = None
    if args.real_tapex:
        from baselines import TaPExBaseline
        tapex_instance = TaPExBaseline(model=args.tapex_model)

    main(
        args.dataset_dir,
        args.k,
        out_record=args.out_record,
        backend_name=backend_name,
        tapex=tapex_instance,
    )
