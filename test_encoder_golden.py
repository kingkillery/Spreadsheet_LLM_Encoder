import json
from pathlib import Path
from unittest.mock import patch

import openpyxl

import paper_serializers
import tokenizer
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode


ROOT = Path(__file__).resolve().parent
FIXTURES_DIR = ROOT / "tests" / "fixtures"
GOLDEN_DIR = ROOT / "tests" / "golden"

FIXTURE_NAMES = [
    "simple_table",
    "sparse_sheet",
    "merged_headers",
    "multi_table_sheet",
    "date_currency_percent",
    "formula_cells",
    "hidden_rows_cols",
    "wide_sparse_sheet",
]

ROUND_TRIP_RANGES = {
    "simple_table": ["A1:C3"],
    "sparse_sheet": ["B6:D8"],
    "merged_headers": ["A1:C4"],
    "multi_table_sheet": ["A1:B3", "D1:E3"],
    "date_currency_percent": ["A1:E3"],
    "formula_cells": ["A1:D4"],
    "hidden_rows_cols": ["A1:C3"],
    "wide_sparse_sheet": ["K1:N3"],
}


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _stable_encoding(encoding):
    return {
        "file_name": encoding["file_name"],
        "sheets": encoding["sheets"],
        "sheet_processing": encoding["sheet_processing"],
    }


def _json_snapshot(value) -> str:
    return json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False) + "\n"


def _vanilla_snapshot(workbook_path: Path) -> str:
    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    prompts = {
        sheet_name: paper_serializers.to_paper_vanilla_prompt(workbook[sheet_name])
        for sheet_name in workbook.sheetnames
    }
    if len(prompts) == 1:
        return next(iter(prompts.values()))
    return "\n\n".join(
        f"# {sheet_name}\n{prompt}" for sheet_name, prompt in prompts.items()
    )


def _compressed_snapshot(encoding) -> str:
    rendered = []
    for sheet_name, sheet_data in encoding["sheets"].items():
        prompt = paper_serializers.to_paper_compressed_prompt(
            sheet_data, coord_map=sheet_data.get("coord_map")
        )
        if len(encoding["sheets"]) == 1:
            rendered.append(prompt)
        else:
            rendered.append(f"# {sheet_name}\n{prompt}")
    return "\n\n".join(rendered)


def _encode_with_fallback(workbook_path: Path):
    with patch.object(tokenizer, "_TIKTOKEN_AVAILABLE", False):
        tokenizer._FALLBACK_WARNED = False
        return spreadsheet_llm_encode(str(workbook_path), k=1, paper_strict=True)


def test_golden_fixture_files_exist():
    for name in FIXTURE_NAMES:
        assert (FIXTURES_DIR / f"{name}.xlsx").exists()
        assert (GOLDEN_DIR / f"{name}.vanilla.txt").exists()
        assert (GOLDEN_DIR / f"{name}.compressed.txt").exists()
        assert (GOLDEN_DIR / f"{name}.encoding.json").exists()
        assert (GOLDEN_DIR / f"{name}.metrics.json").exists()


def test_vanilla_prompt_snapshots_match_golden_files():
    for name in FIXTURE_NAMES:
        workbook_path = FIXTURES_DIR / f"{name}.xlsx"
        assert _vanilla_snapshot(workbook_path) == _read(GOLDEN_DIR / f"{name}.vanilla.txt")


def test_compressed_prompt_snapshots_match_golden_files():
    for name in FIXTURE_NAMES:
        encoding = _encode_with_fallback(FIXTURES_DIR / f"{name}.xlsx")
        assert _compressed_snapshot(encoding) == _read(GOLDEN_DIR / f"{name}.compressed.txt")


def test_stable_encoding_snapshots_match_golden_files():
    for name in FIXTURE_NAMES:
        encoding = _encode_with_fallback(FIXTURES_DIR / f"{name}.xlsx")
        assert _json_snapshot(_stable_encoding(encoding)) == _read(
            GOLDEN_DIR / f"{name}.encoding.json"
        )


def test_metric_snapshots_include_tokenizer_metadata_and_match_golden_files():
    for name in FIXTURE_NAMES:
        encoding = _encode_with_fallback(FIXTURES_DIR / f"{name}.xlsx")
        metrics = encoding["compression_metrics"]
        assert metrics["tokenizer"]["backend"] == "char_approximation"
        assert metrics["tokenizer"]["fallback"] is True
        assert _json_snapshot(metrics) == _read(GOLDEN_DIR / f"{name}.metrics.json")


def test_golden_coord_maps_round_trip_expected_ranges():
    for name in FIXTURE_NAMES:
        encoding = _encode_with_fallback(FIXTURES_DIR / f"{name}.xlsx")
        sheet_data = encoding["sheets"]["Sheet1"]
        coord_map = sheet_data["coord_map"]
        for original_range in ROUND_TRIP_RANGES[name]:
            compact = paper_serializers.remap_range(original_range, coord_map)
            assert compact is not None, (name, original_range)
            assert paper_serializers.unremap_range(compact, coord_map) == original_range
