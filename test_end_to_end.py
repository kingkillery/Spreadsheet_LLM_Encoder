"""End-to-end smoke test for the SpreadsheetLLM pipeline.

Wires the synthetic benchmark generator → encoder → LLM evaluation script
together so that a contract mismatch between any pair surfaces here, not in
production. Uses the deterministic ``EchoBackend`` so the test is hermetic
(no network, no GPUs, no datasets).
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

import paper_serializers
from llm_backend import EchoBackend
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode


REPO_ROOT = Path(__file__).resolve().parent


def _run(cmd: list[str]) -> subprocess.CompletedProcess:
    """Run a subprocess from the repo root, capturing stdout/stderr."""
    return subprocess.run(
        cmd, cwd=REPO_ROOT, capture_output=True, text=True, check=False
    )


class TestEndToEnd(unittest.TestCase):

    def setUp(self):
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="sllm_e2e_"))

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_synth_then_encode_then_eval_smoke(self):
        # 1. Synthesize a tiny benchmark.
        synth_dir = self.tmp_dir / "bench"
        result = _run([
            sys.executable, "scripts/synthesize_benchmark.py",
            str(synth_dir), "--n", "3", "--seed", "1",
        ])
        self.assertEqual(
            result.returncode, 0,
            msg=f"synthesize_benchmark failed: {result.stderr}",
        )
        xlsx_files = sorted(synth_dir.glob("*.xlsx"))
        json_files = sorted(synth_dir.glob("*.json"))
        self.assertEqual(len(xlsx_files), 3)
        self.assertEqual(len(json_files), 3)

        # 2. Encode the first workbook through the paper-aligned pipeline.
        encoding = spreadsheet_llm_encode(str(xlsx_files[0]), k=4)
        self.assertIsNotNone(encoding)
        self.assertIn("sheets", encoding)
        self.assertIn("compression_metrics", encoding)

        # Each sheet must carry coord_map (paper-aligned remapping).
        for sheet_name, sheet_data in encoding["sheets"].items():
            self.assertIn("coord_map", sheet_data, msg=f"sheet={sheet_name}")
            cm = sheet_data["coord_map"]
            for k in ("rows", "cols", "rows_inv", "cols_inv"):
                self.assertIn(k, cm)

        # 3. The compressed prompt for that sheet must be a non-empty string
        #    of paper-style tuples — i.e. starts with '(' or contains '|'.
        first_sheet = next(iter(encoding["sheets"].values()))
        prompt = paper_serializers.to_paper_compressed_prompt(
            first_sheet, coord_map=first_sheet["coord_map"]
        )
        self.assertGreater(len(prompt), 0)
        self.assertTrue(
            "(" in prompt and "|" in prompt,
            msg=f"compressed prompt does not look paper-formatted: {prompt[:120]!r}",
        )

        # 4. Run the LLM evaluation script with the EchoBackend so we exercise
        #    the full glue: encoder + paper-compressed prompt + range parsing
        #    + unremap_range + EoB scoring. Echo returns a single hardcoded
        #    range so we don't care about the F1 — only that the script
        #    completes without raising.
        result = _run([
            sys.executable, "run_llm_evaluation.py",
            str(synth_dir),
            "--backend", "echo",
            "--echo-response", "['range': 'A1:D5']",
            "--k", "4",
        ])
        self.assertEqual(
            result.returncode, 0,
            msg=(
                "run_llm_evaluation failed:\n"
                f"stdout=\n{result.stdout}\n"
                f"stderr=\n{result.stderr}"
            ),
        )
        # Average F1 line should appear in the log output (logger writes to
        # stderr by default for INFO).
        combined = result.stdout + result.stderr
        self.assertIn("F1", combined)

    def test_eval_writes_out_record(self):
        """run_llm_evaluation --out-record persists structured metrics."""
        synth_dir = self.tmp_dir / "rec_bench"
        record_path = self.tmp_dir / "eval.json"

        result = _run([
            sys.executable, "scripts/synthesize_benchmark.py",
            str(synth_dir), "--n", "2", "--seed", "3",
        ])
        self.assertEqual(result.returncode, 0, msg=result.stderr)

        result = _run([
            sys.executable, "run_llm_evaluation.py",
            str(synth_dir),
            "--backend", "echo",
            "--echo-response", "['range': 'A1:D5']",
            "--k", "4",
            "--out-record", str(record_path),
        ])
        self.assertEqual(
            result.returncode, 0,
            msg=f"eval failed:\n{result.stdout}\n{result.stderr}",
        )

        self.assertTrue(record_path.exists())
        record = json.loads(record_path.read_text(encoding="utf-8"))
        for key in (
            "timestamp", "task", "dataset_dir", "k", "backend",
            "n_items", "avg_f1_eob0", "per_item", "meta",
        ):
            self.assertIn(key, record)
        self.assertEqual(record["task"], "table_detection_eob0")
        self.assertEqual(record["k"], 4)
        self.assertEqual(record["backend"], "echo")
        self.assertEqual(record["n_items"], 2)
        self.assertEqual(len(record["per_item"]), 2)
        for item in record["per_item"]:
            for sub_key in ("spreadsheet_path", "gt_count", "pred_count",
                            "precision", "recall", "f1"):
                self.assertIn(sub_key, item)

    def test_qa_eval_writes_out_record(self):
        """run_qa_evaluation --out-record persists structured QA metrics
        and the echo-backend path no longer raises NotImplementedError."""
        synth_dir = self.tmp_dir / "qa_rec_bench"
        record_path = self.tmp_dir / "qa_eval.json"

        # 1. Synthesize annotated workbooks + QA pairs.
        r = _run([
            sys.executable, "scripts/synthesize_benchmark.py",
            str(synth_dir), "--n", "2", "--seed", "11",
        ])
        self.assertEqual(r.returncode, 0, msg=r.stderr)
        r = _run([
            sys.executable, "scripts/synthesize_qa.py",
            str(synth_dir), "--per-table", "2", "--seed", "11",
            "--out-suffix", "",
        ])
        self.assertEqual(r.returncode, 0, msg=r.stderr)

        # 2. Run QA eval with echo backend + out-record.
        r = _run([
            sys.executable, "run_qa_evaluation.py",
            str(synth_dir),
            "--backend", "echo",
            "--echo-response", "['range': 'A1:D5']",
            "--k", "4",
            "--out-record", str(record_path),
        ])
        self.assertEqual(
            r.returncode, 0,
            msg=f"qa eval failed:\n{r.stdout}\n{r.stderr}",
        )

        self.assertTrue(record_path.exists())
        record = json.loads(record_path.read_text(encoding="utf-8"))
        for key in (
            "timestamp", "task", "dataset_dir", "k", "backend",
            "n_questions", "spreadsheetllm_accuracy_pct",
            "tapex_accuracy_pct", "tapex_kind",
            "binder_placeholder_accuracy_pct",
            "baselines_are_placeholders", "per_question", "meta",
        ):
            self.assertIn(key, record, msg=f"missing key: {key}")
        self.assertEqual(record["task"], "spreadsheet_qa")
        self.assertEqual(record["k"], 4)
        self.assertEqual(record["backend"], "echo")
        # Without --real-tapex the run uses the placeholder.
        self.assertEqual(record["tapex_kind"], "placeholder")
        self.assertTrue(record["baselines_are_placeholders"])
        self.assertGreater(record["n_questions"], 0)
        self.assertEqual(len(record["per_question"]), record["n_questions"])
        for pq in record["per_question"]:
            for sub in ("question", "ground_truth", "spreadsheetllm",
                        "tapex_placeholder", "binder_placeholder"):
                self.assertIn(sub, pq)

    def test_synth_qa_round_trips_through_qa_dataset_loader(self):
        # 1. Synthesize annotated workbooks.
        synth_dir = self.tmp_dir / "qa_bench"
        r = _run([
            sys.executable, "scripts/synthesize_benchmark.py",
            str(synth_dir), "--n", "2", "--seed", "2",
        ])
        self.assertEqual(r.returncode, 0, msg=r.stderr)

        # 2. Generate QA pairs in-place (--out-suffix "" overwrites .json).
        r = _run([
            sys.executable, "scripts/synthesize_qa.py",
            str(synth_dir), "--per-table", "3", "--seed", "2",
            "--out-suffix", "",
        ])
        self.assertEqual(r.returncode, 0, msg=r.stderr)

        # 3. Loader must accept the result and surface qa_pairs.
        from evaluation import load_qa_dataset
        qa_records = load_qa_dataset(str(synth_dir))
        self.assertGreaterEqual(len(qa_records), 1)
        for rec in qa_records:
            self.assertIn("qa_pairs", rec)
            self.assertGreater(len(rec["qa_pairs"]), 0)
            for pair in rec["qa_pairs"]:
                self.assertIn("question", pair)
                self.assertIn("answer", pair)
                # Answers must be wrapped in [...] per the paper's contract.
                self.assertTrue(
                    pair["answer"].startswith("[") and pair["answer"].endswith("]"),
                    msg=f"unwrapped answer: {pair['answer']!r}",
                )


if __name__ == "__main__":
    unittest.main()
