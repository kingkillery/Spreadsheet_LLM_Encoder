"""
test_synthesize.py
------------------
Tests for scripts/synthesize_benchmark.py and scripts/synthesize_qa.py.
"""
import json
import subprocess
import sys
import tempfile
from pathlib import Path


def _run(script, *args):
    result = subprocess.run(
        [sys.executable, script, *args],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, (
        f"{script} exited {result.returncode}\n"
        f"stdout: {result.stdout}\n"
        f"stderr: {result.stderr}"
    )
    return result


def test_synthesize_benchmark():
    N = 5
    with tempfile.TemporaryDirectory() as tmp:
        script = str(Path(__file__).parent / "scripts" / "synthesize_benchmark.py")
        result = _run(script, tmp, "--n", str(N), "--seed", "7")

        # Summary line check
        assert f"Wrote {N} files to" in result.stdout

        xlsx_files = list(Path(tmp).glob("*.xlsx"))
        json_files = list(Path(tmp).glob("*.json"))

        assert len(xlsx_files) == N, f"Expected {N} xlsx, got {len(xlsx_files)}"
        assert len(json_files) == N, f"Expected {N} json, got {len(json_files)}"

        # Each JSON has non-empty 'tables' with valid ranges
        for jf in json_files:
            data = json.loads(jf.read_text())
            assert "tables" in data and len(data["tables"]) > 0, f"{jf} missing tables"
            for t in data["tables"]:
                assert "range" in t and ":" in t["range"], f"{jf} bad range: {t}"
                assert "layout" in t
                assert "size_class" in t


def test_synthesize_qa():
    N = 5
    with tempfile.TemporaryDirectory() as tmp:
        bench_script = str(Path(__file__).parent / "scripts" / "synthesize_benchmark.py")
        qa_script = str(Path(__file__).parent / "scripts" / "synthesize_qa.py")

        _run(bench_script, tmp, "--n", str(N), "--seed", "13")
        result = _run(qa_script, tmp, "--per-table", "3", "--seed", "13")

        assert "Wrote" in result.stdout

        qa_files = list(Path(tmp).glob("*_qa.json"))
        assert len(qa_files) == N, f"Expected {N} QA files, got {len(qa_files)}"

        for qf in qa_files:
            data = json.loads(qf.read_text())
            assert "qa_pairs" in data and len(data["qa_pairs"]) > 0, (
                f"{qf} missing qa_pairs"
            )
            for pair in data["qa_pairs"]:
                assert "question" in pair and "answer" in pair
                # Answers must be wrapped in [...]
                assert pair["answer"].startswith("[") and pair["answer"].endswith("]"), (
                    f"Answer not wrapped in brackets: {pair['answer']!r}"
                )


def test_load_spreadsheet_dataset_compat():
    """Verify load_spreadsheet_dataset can consume synthesize_benchmark output."""
    import sys, importlib
    sys.path.insert(0, str(Path(__file__).parent))
    from evaluation import load_spreadsheet_dataset

    N = 5
    with tempfile.TemporaryDirectory() as tmp:
        bench_script = str(Path(__file__).parent / "scripts" / "synthesize_benchmark.py")
        _run(bench_script, tmp, "--n", str(N), "--seed", "99")
        dataset = load_spreadsheet_dataset(tmp)
        assert len(dataset) == N, f"Expected {N} entries, got {len(dataset)}"
        for entry in dataset:
            assert "spreadsheet_path" in entry
            assert "bboxes" in entry and len(entry["bboxes"]) > 0


def test_load_qa_dataset_compat():
    """Verify load_qa_dataset can consume synthesize_qa output."""
    import sys
    sys.path.insert(0, str(Path(__file__).parent))
    from evaluation import load_qa_dataset

    N = 5
    with tempfile.TemporaryDirectory() as tmp:
        bench_script = str(Path(__file__).parent / "scripts" / "synthesize_benchmark.py")
        qa_script = str(Path(__file__).parent / "scripts" / "synthesize_qa.py")
        _run(bench_script, tmp, "--n", str(N), "--seed", "77")
        _run(qa_script, tmp, "--per-table", "3", "--seed", "77", "--out-suffix", "")

        dataset = load_qa_dataset(tmp)
        assert len(dataset) > 0, "load_qa_dataset returned empty list"
        for entry in dataset:
            assert "qa_pairs" in entry and len(entry["qa_pairs"]) > 0
