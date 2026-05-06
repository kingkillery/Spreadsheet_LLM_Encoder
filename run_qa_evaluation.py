import argparse
import datetime as _dt
import logging
import os
from typing import List, Dict, Optional

import chain_of_spreadsheet
from evaluation import load_qa_dataset, load_qa_manifest, normalize_qa_answer
from evaluation_metadata import build_evaluation_metadata, write_evaluation_record
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode
from chain_of_spreadsheet import identify_table, table_split_qa
from baselines import BINDER_UNAVAILABLE_REASON, BinderBaseline

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


def run_binder_unavailable() -> Optional[str]:
    """Return no Binder prediction until a real adapter exists."""
    logger.info("Skipping Binder baseline: %s", BinderBaseline().reason)
    return None


QA_METRIC_DEFINITION = (
    "Per-item answer-type aware exact match via evaluation.normalize_qa_answer "
    "(cell_address: trimmed + uppercased; "
    "formula: whitespace-stripped + uppercased; "
    "free_text: casefolded + collapsed whitespace; "
    "literal: stripped only). Default answer_type when unspecified is 'literal'."
)


def main(
    dataset_dir: Optional[str],
    k: int,
    out_record: Optional[str] = None,
    backend_name: str = "unknown",
    extra_meta: Optional[Dict] = None,
    tapex=None,
    manifest_path: Optional[str] = None,
):
    """Main function to run the Spreadsheet QA evaluation.

    Either ``dataset_dir`` or ``manifest_path`` must be supplied. When
    ``out_record`` is given, also persist a structured JSON record
    (timestamp, dataset, k, backend, per-question results, accuracies)
    so the spreadsheet-llm-fidelity skill_runs log can track QA accuracy
    iteration-over-iteration without log-string parsing.
    """
    if not dataset_dir and not manifest_path:
        raise ValueError("either dataset_dir or manifest_path must be provided")
    dataset = load_qa_manifest(manifest_path) if manifest_path else load_qa_dataset(dataset_dir)

    if not dataset:
        logger.error("No QA data found in the specified dataset directory.")
        return

    total_questions = 0
    correct_spreadsheetllm = 0
    correct_tape = 0
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
            answer_type = qa.get("answer_type", "literal")

            logger.info("Q: %s (GT: %s)", query, ground_truth)

            pred_answer_llm: Optional[str] = None
            llm_correct = False
            tapex_table_range: Optional[str] = None
            tapex_sheet_name: Optional[str] = None
            stage2_mode = "not_run"
            stage2_original_range: Optional[str] = None

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
                norm_pred = normalize_qa_answer(pred_answer_llm, answer_type)
                norm_gt = normalize_qa_answer(ground_truth, answer_type)
                if norm_pred == norm_gt:
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
                stage2_original_range = tapex_table_range
                stage2_mode = "original_workbook_uncompressed"
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
            tape_correct = (
                normalize_qa_answer(pred_answer_tape, answer_type)
                == normalize_qa_answer(ground_truth, answer_type)
            )
            if tape_correct:
                correct_tape += 1

            pred_answer_binder = run_binder_unavailable()
            logger.info("  - Binder Predicted: unavailable")
            binder_correct = False

            per_question.append({
                "spreadsheet_path": spreadsheet_path,
                "question": query,
                "ground_truth": ground_truth,
                "answer_type": answer_type,
                "spreadsheetllm": {
                    "predicted": pred_answer_llm,
                    "correct": llm_correct,
                    "sheet_name": tapex_sheet_name,
                    "table_range": table_range,
                    "original_table_range": stage2_original_range,
                    "stage2_mode": stage2_mode,
                },
                f"tapex_{tapex_kind}": {
                    "predicted": pred_answer_tape,
                    "correct": tape_correct,
                },
                "binder_unavailable": {
                    "predicted": pred_answer_binder,
                    "correct": binder_correct,
                    "skip_reason": BINDER_UNAVAILABLE_REASON,
                },
            })

    # --- Report Results ---
    logger.info("\n--- QA Evaluation Summary ---")
    acc_llm = acc_tape = 0.0
    acc_binder = None
    if total_questions > 0:
        acc_llm = (correct_spreadsheetllm / total_questions) * 100
        acc_tape = (correct_tape / total_questions) * 100
        logger.info(
            "SpreadsheetLLM Accuracy: %.2f%% (%d/%d)",
            acc_llm, correct_spreadsheetllm, total_questions,
        )
        logger.info(
            "TaPEx Baseline Accuracy: %.2f%% (%d/%d)",
            acc_tape, correct_tape, total_questions,
        )
        logger.info("Binder Baseline: unavailable (%s)", BINDER_UNAVAILABLE_REASON)
    else:
        logger.info("No questions were evaluated.")
    logger.info("--------------------------")

    if out_record:
        tapex_real = tapex is not None
        skip_reasons = [BinderBaseline().skip_reason()]
        if not tapex_real:
            skip_reasons.append({
                "component": "tapex",
                "reason": "Run omitted --real-tapex; placeholder baseline used.",
            })
        evaluation_metadata = build_evaluation_metadata(
            dataset_dir=manifest_path or dataset_dir,
            task="spreadsheet_qa",
            dataset_name=dataset[0].get("dataset_name") if dataset else None,
            dataset_version=dataset[0].get("dataset_version", "unspecified") if dataset else "unspecified",
            split_name=dataset[0].get("split_name", "unspecified") if dataset else "unspecified",
            spreadsheet_count=len(dataset),
            table_count=0,
            qa_item_count=total_questions,
            encoder_settings={"k": k},
            prompt_serializer="paper_serializers.to_paper_compressed_prompt + stage2_uncompressed_pairs_when_available",
            coordinate_mode="compact_stage1_original_stage2_when_workbook_available",
            model_backend=backend_name,
            metric_definition=QA_METRIC_DEFINITION,
            baseline_name="SpreadsheetLLM QA with TaPEx/Binder baselines",
            skip_reasons=skip_reasons,
        )
        record = {
            "timestamp": _dt.datetime.now(_dt.timezone.utc).isoformat(),
            "task": "spreadsheet_qa",
            "dataset_dir": os.path.abspath(dataset_dir) if dataset_dir else None,
            "manifest_path": os.path.abspath(manifest_path) if manifest_path else None,
            "k": k,
            "backend": backend_name,
            "n_questions": total_questions,
            "spreadsheetllm_accuracy_pct": acc_llm,
            "tapex_accuracy_pct": acc_tape,
            "tapex_kind": "real" if tapex_real else "placeholder",
            "binder_accuracy_pct": acc_binder,
            "binder_status": "unavailable",
            "binder_skip_reason": BINDER_UNAVAILABLE_REASON,
            "baselines_are_placeholders": not tapex_real,
            "per_question": per_question,
            "meta": extra_meta or {},
            "evaluation_metadata": evaluation_metadata,
        }
        write_evaluation_record(record, out_record)
        logger.info("Wrote QA evaluation record to %s", os.path.abspath(out_record))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run Spreadsheet QA evaluation.")
    parser.add_argument(
        "dataset_dir",
        nargs="?",
        default=None,
        help="Path to the QA dataset directory. Optional when --manifest is provided.",
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
        "--manifest",
        default=None,
        help="Optional QA manifest JSON. When set, it overrides dataset_dir scanning.",
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

    if not args.dataset_dir and not args.manifest:
        parser.error("either dataset_dir or --manifest is required")

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
        manifest_path=args.manifest,
    )
