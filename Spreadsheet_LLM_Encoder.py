import os
import openpyxl
import json
import logging
import re
from fnmatch import fnmatch
from copy import copy
from temp_helpers import (
    infer_cell_data_type,
    categorize_number_format,
    get_number_format_string,
    detect_semantic_type,
)
from collections import defaultdict
from openpyxl.utils import get_column_letter

import sys

import paper_serializers
from tokenizer import count_tokens, DEFAULT_MODEL, tokenizer_metadata

logger = logging.getLogger(__name__)

EXCEL_ERROR_VALUES = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
SPARSE_TEXT_MAX_RATIO = 0.5
MAX_TRAILING_NOTE_ROWS = 2
MAX_SPARSE_COLUMN_RATIO = 0.5
MAX_DETAILED_ANCHOR_DIMENSION = 10
HEADER_AT_TOP_BONUS = 30
HEADER_AFTER_TITLE_BONUS = 24
MAX_BODY_ROW_BONUS = 10
MAX_WIDTH_BONUS = 8
BODY_DENSITY_WEIGHT = 25
RANGE_DENSITY_WEIGHT = 8
YEAR_OR_DATE_WEIGHT = 4
MAX_POPULATED_CELL_BONUS = 12
BODY_ROW_BONUS = 2
TITLE_ROW_BONUS = 2
NOTE_ROW_BONUS = 1
EXTRA_ROW_PENALTY = 2
OVERLAP_IOU_SUPPRESSION_THRESHOLD = 0.5
OVERLAP_CONTAINMENT_SUPPRESSION_THRESHOLD = 0.85
_FORMULA_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"(?:(?:'(?P<quoted_sheet>[^']+)'|(?P<sheet>[A-Za-z_][A-Za-z0-9_ .]*))!)?"
    r"\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)"
    r"(?::\$?(?P<end_col>[A-Z]{1,3})\$?(?P<end_row>\d+))?",
    re.IGNORECASE,
)


def calculate_compression_ratio(original_tokens: int, compressed_tokens: int) -> float:
    """Return the compression ratio given original and compressed token counts."""
    if compressed_tokens == 0:
        return 0.0
    if original_tokens == 0:
        return 1.0
    return original_tokens / compressed_tokens


def _normalize_formula_reference(match, current_sheet: str) -> str:
    sheet = match.group("quoted_sheet") or match.group("sheet") or current_sheet
    col = match.group("col").upper()
    row = match.group("row")
    end_col = match.group("end_col")
    end_row = match.group("end_row")
    if end_col and end_row:
        return f"{sheet}!{col}{row}:{end_col.upper()}{end_row}"
    return f"{sheet}!{col}{row}"


def extract_formula_references(formula: str, current_sheet: str) -> list:
    """Return normalized workbook references used by an Excel formula."""
    refs = []
    seen = set()
    for match in _FORMULA_REF_RE.finditer(formula or ""):
        ref = _normalize_formula_reference(match, current_sheet)
        if ref not in seen:
            refs.append(ref)
            seen.add(ref)
    return refs


def _formula_pattern(formula: str, current_sheet: str) -> str:
    """Collapse references in a formula so fill-down families can be grouped."""
    def repl(match):
        ref = _normalize_formula_reference(match, current_sheet)
        return "<RANGE>" if ":" in ref else "<REF>"

    return _FORMULA_REF_RE.sub(repl, formula or "")


def _json_safe_value(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def extract_formula_graph(formula_sheet, cached_sheet=None) -> dict:
    """Extract lightweight formula dependencies and spreadsheet error cells.

    ``formula_sheet`` must be loaded with ``data_only=False`` so formula text is
    available. ``cached_sheet`` should be the same worksheet loaded with
    ``data_only=True`` when cached formula results are available.
    """
    formulas = []
    formula_errors = []
    exact_groups = defaultdict(list)
    pattern_groups = defaultdict(list)
    sheet_name = formula_sheet.title

    for row in range(1, formula_sheet.max_row + 1):
        for col in range(1, formula_sheet.max_column + 1):
            formula_cell = formula_sheet.cell(row=row, column=col)
            value = formula_cell.value
            ref = f"{get_column_letter(col)}{row}"
            qualified_ref = f"{sheet_name}!{ref}"

            cached_value = None
            if cached_sheet is not None:
                cached_value = cached_sheet.cell(row=row, column=col).value

            if isinstance(value, str) and value in EXCEL_ERROR_VALUES:
                formula_errors.append({"cell": qualified_ref, "error": value})

            if not (isinstance(value, str) and value.startswith("=")):
                continue

            references = extract_formula_references(value, sheet_name)
            cross_sheet_references = [
                reference
                for reference in references
                if reference.split("!", 1)[0] != sheet_name
            ]
            errors = []
            if isinstance(cached_value, str) and cached_value in EXCEL_ERROR_VALUES:
                errors.append(cached_value)

            formulas.append({
                "cell": qualified_ref,
                "formula": value,
                "cached_value": _json_safe_value(cached_value),
                "references": references,
                "cross_sheet_references": cross_sheet_references,
                "errors": errors,
            })
            exact_groups[value].append(qualified_ref)
            pattern_groups[_formula_pattern(value, sheet_name)].append(qualified_ref)

    repeated_formula_summaries = []
    for formula, cells in sorted(exact_groups.items()):
        if len(cells) > 1:
            repeated_formula_summaries.append({
                "kind": "exact",
                "formula": formula,
                "count": len(cells),
                "cells": cells,
            })
    for pattern, cells in sorted(pattern_groups.items()):
        if len(cells) > 1:
            repeated_formula_summaries.append({
                "kind": "pattern",
                "formula_pattern": pattern,
                "count": len(cells),
                "cells": cells,
            })

    return {
        "formulas": formulas,
        "formula_errors": formula_errors,
        "repeated_formula_summaries": repeated_formula_summaries,
    }


def _limit_to_positive_int(value, label):
    if value is None:
        return None
    value = int(value)
    if value <= 0:
        raise ValueError(f"{label} must be a positive integer")
    return value


def _bounded_dimensions(
    rows,
    cols,
    max_rows_per_sheet=None,
    max_cols_per_sheet=None,
    max_cells_per_sheet=None,
):
    """Return bounded ``(rows, cols)`` while preserving full size by default."""
    effective_rows = rows
    effective_cols = cols
    if max_rows_per_sheet is not None:
        effective_rows = min(effective_rows, max_rows_per_sheet)
    if max_cols_per_sheet is not None:
        effective_cols = min(effective_cols, max_cols_per_sheet)

    if max_cells_per_sheet is not None and effective_rows * effective_cols > max_cells_per_sheet:
        if effective_cols > max_cells_per_sheet:
            effective_cols = max_cells_per_sheet
            effective_rows = 1
        else:
            effective_rows = max(1, max_cells_per_sheet // max(1, effective_cols))

    return max(1, effective_rows), max(1, effective_cols)


def _validate_and_normalize_filter_list(values, parameter_name):
    if values is None:
        return []
    if isinstance(values, str):
        values = [values]
    normalized = []
    for value in values:
        text = str(value).strip()
        if not text:
            raise ValueError(f"{parameter_name} entries must be non-empty")
        normalized.append(text)
    return normalized


def _compile_sheet_regexes(patterns, parameter_name):
    """Compile sheet-name regex filters as ``(pattern, compiled_regex)`` tuples."""
    compiled = []
    for pattern in patterns:
        try:
            compiled.append((pattern, re.compile(pattern)))
        except re.error as exc:
            raise ValueError(f"Invalid {parameter_name} pattern '{pattern}': {exc}") from exc
    return compiled


def _sheet_selection_decision(
    sheet_name,
    include_names,
    include_globs,
    include_regexes,
    exclude_names,
    exclude_globs,
    exclude_regexes,
):
    include_filters_active = bool(include_names or include_globs or include_regexes)
    include_matches = (
        sheet_name in include_names
        or any(fnmatch(sheet_name, pattern) for pattern in include_globs)
        or any(regex.search(sheet_name) for _, regex in include_regexes)
    )
    if include_filters_active and not include_matches:
        return False, "sheet not matched by include filters"

    if sheet_name in exclude_names:
        return False, "sheet excluded by name filter"

    for pattern in exclude_globs:
        if fnmatch(sheet_name, pattern):
            return False, f"sheet excluded by glob filter '{pattern}'"

    for pattern, regex in exclude_regexes:
        if regex.search(sheet_name):
            return False, f"sheet excluded by regex filter '{pattern}'"

    return True, None


def _copy_bounded_sheet(source_sheet, max_row, max_col):
    """Copy a bounded top-left worksheet region into a normal worksheet."""
    wb = openpyxl.Workbook()
    target = wb.active
    target.title = source_sheet.title

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            source_cell = source_sheet.cell(row=row, column=col)
            target_cell = target.cell(row=row, column=col, value=source_cell.value)
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.fill = copy(source_cell.fill)
                target_cell.border = copy(source_cell.border)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.protection = copy(source_cell.protection)
            target_cell.number_format = source_cell.number_format

    for merged_range in source_sheet.merged_cells.ranges:
        if merged_range.max_row <= max_row and merged_range.max_col <= max_col:
            target.merge_cells(str(merged_range))
    return target


def _load_xlsb_workbook(excel_path):
    """Load ``.xlsb`` into an in-memory openpyxl workbook with cell values only."""
    try:
        import pyxlsb
    except ImportError as exc:
        raise ImportError(
            "Reading .xlsb files requires optional dependency 'pyxlsb'. "
            "Install it with `pip install pyxlsb`."
        ) from exc

    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    with pyxlsb.open_workbook(excel_path) as source_workbook:
        for sheet_name in source_workbook.sheets:
            target_sheet = workbook.create_sheet(title=sheet_name)
            with source_workbook.get_sheet(sheet_name) as source_sheet:
                for row_idx, row in enumerate(source_sheet.rows(), start=1):
                    for col_idx, cell in enumerate(row, start=1):
                        value = getattr(cell, "v", None)
                        if value is not None:
                            target_sheet.cell(row=row_idx, column=col_idx, value=value)
    return workbook


def _load_workbooks(excel_path, data_only=True):
    extension = os.path.splitext(str(excel_path))[1].lower()
    if extension == ".xlsb":
        workbook = _load_xlsb_workbook(excel_path)
        logger.warning(
            "Loaded .xlsb workbook via pyxlsb value-only fallback; style metadata, merged-cell regions, "
            "and formula text/cached-value distinction may differ from .xlsx loading."
        )
        return workbook, workbook, workbook

    workbook = openpyxl.load_workbook(excel_path, data_only=data_only)
    formula_workbook = openpyxl.load_workbook(excel_path, data_only=False)
    cached_workbook = workbook if data_only else openpyxl.load_workbook(excel_path, data_only=True)
    return workbook, formula_workbook, cached_workbook


def _sheet_processing_plan(
    sheet,
    *,
    max_rows_per_sheet=None,
    max_cols_per_sheet=None,
    max_cells_per_sheet=None,
    sheet_limit_action="truncate",
):
    if sheet_limit_action not in {"truncate", "skip", "error"}:
        raise ValueError("sheet_limit_action must be 'truncate', 'skip', or 'error'")

    original_rows = sheet.max_row or 1
    original_cols = sheet.max_column or 1
    effective_rows, effective_cols = _bounded_dimensions(
        original_rows,
        original_cols,
        max_rows_per_sheet=max_rows_per_sheet,
        max_cols_per_sheet=max_cols_per_sheet,
        max_cells_per_sheet=max_cells_per_sheet,
    )
    truncated = effective_rows < original_rows or effective_cols < original_cols
    metadata = {
        "status": "encoded",
        "limit_action": sheet_limit_action,
        "truncated": truncated,
        "original_rows": original_rows,
        "original_cols": original_cols,
        "original_cells": original_rows * original_cols,
        "effective_rows": effective_rows,
        "effective_cols": effective_cols,
        "effective_cells": effective_rows * effective_cols,
        "encoded_range": f"A1:{get_column_letter(effective_cols)}{effective_rows}",
    }
    if truncated:
        metadata["reason"] = "sheet exceeds configured row/column/cell limits"
        if sheet_limit_action == "skip":
            metadata["status"] = "skipped"
            metadata["encoded_range"] = None
        elif sheet_limit_action == "error":
            raise ValueError(
                f"Sheet '{sheet.title}' exceeds configured limits: "
                f"{original_rows}x{original_cols} -> {effective_rows}x{effective_cols}"
            )
    return effective_rows, effective_cols, metadata


def spreadsheet_llm_encode(
    excel_path,
    output_path=None,
    k=4,
    vanilla=False,
    compress_homogeneous=True,
    paper_strict=False,
    data_only=True,
    tokenizer_model=DEFAULT_MODEL,
    max_rows_per_sheet=None,
    max_cols_per_sheet=None,
    max_cells_per_sheet=None,
    sheet_limit_action="truncate",
    include_sheets=None,
    exclude_sheets=None,
    include_sheet_globs=None,
    exclude_sheet_globs=None,
    include_sheet_regexes=None,
    exclude_sheet_regexes=None,
):
    """
    Convert an Excel file to SpreadsheetLLM format or a vanilla markdown-like format.

    Args:
        excel_path (str): Path to the Excel file (.xlsx or .xlsb).
        output_path (str, optional): Path to save the output. Defaults to None.
        k (int, optional): Neighborhood distance for structural anchors.
            Defaults to 4 (paper's best ablation setting).
        vanilla (bool, optional): If True, produce vanilla encoding instead of compressed.
                                Defaults to False.
        compress_homogeneous (bool, optional): Drop fully-homogeneous rows/cols
            after anchor extraction. Defaults to True. Set False for strict
            paper-aligned skeleton retention.
        paper_strict (bool, optional): Apply paper-faithful behavior where it
            differs from pragmatic defaults. Currently this disables
            post-anchor homogeneous row/column pruning. Defaults to False.
        data_only (bool, optional): Load cached formula values instead of formula
            text. Defaults to True (paper expects user-visible values).
        tokenizer_model (str, optional): Model name for tokenizer-based
            compression metrics. Defaults to ``"gpt-4"``.
        max_rows_per_sheet (int, optional): When set, cap each sheet to this
            many rows in bounded mode.
        max_cols_per_sheet (int, optional): When set, cap each sheet to this
            many columns in bounded mode.
        max_cells_per_sheet (int, optional): When set, cap each sheet to this
            many cells by reducing the effective row count after row/column
            caps are applied.
        sheet_limit_action (str, optional): What to do when a sheet exceeds
            the configured caps: ``"truncate"`` (default), ``"skip"``, or
            ``"error"``.
        include_sheets (Iterable[str] | str, optional): Exact sheet names to
            include. When provided, only matching sheets are encoded.
        exclude_sheets (Iterable[str] | str, optional): Exact sheet names to
            exclude from encoding.
        include_sheet_globs (Iterable[str] | str, optional): Glob patterns
            for sheets to include.
        exclude_sheet_globs (Iterable[str] | str, optional): Glob patterns
            for sheets to exclude.
        include_sheet_regexes (Iterable[str] | str, optional): Regex patterns
            for sheets to include.
        exclude_sheet_regexes (Iterable[str] | str, optional): Regex patterns
            for sheets to exclude.

    Returns:
        dict: The SpreadsheetLLM encoding of the Excel file.
    """
    if paper_strict:
        compress_homogeneous = False
    max_rows_per_sheet = _limit_to_positive_int(max_rows_per_sheet, "max_rows_per_sheet")
    max_cols_per_sheet = _limit_to_positive_int(max_cols_per_sheet, "max_cols_per_sheet")
    max_cells_per_sheet = _limit_to_positive_int(max_cells_per_sheet, "max_cells_per_sheet")
    include_sheets = _validate_and_normalize_filter_list(include_sheets, "include_sheets")
    exclude_sheets = _validate_and_normalize_filter_list(exclude_sheets, "exclude_sheets")
    include_sheet_globs = _validate_and_normalize_filter_list(include_sheet_globs, "include_sheet_globs")
    exclude_sheet_globs = _validate_and_normalize_filter_list(exclude_sheet_globs, "exclude_sheet_globs")
    include_sheet_regexes = _validate_and_normalize_filter_list(include_sheet_regexes, "include_sheet_regexes")
    exclude_sheet_regexes = _validate_and_normalize_filter_list(exclude_sheet_regexes, "exclude_sheet_regexes")
    if vanilla:
        return vanilla_encode(
            excel_path,
            output_path,
            include_sheets=include_sheets,
            exclude_sheets=exclude_sheets,
            include_sheet_globs=include_sheet_globs,
            exclude_sheet_globs=exclude_sheet_globs,
            include_sheet_regexes=include_sheet_regexes,
            exclude_sheet_regexes=exclude_sheet_regexes,
        )
    include_sheet_regexes_compiled = _compile_sheet_regexes(
        include_sheet_regexes,
        "include_sheet_regexes",
    )
    exclude_sheet_regexes_compiled = _compile_sheet_regexes(
        exclude_sheet_regexes,
        "exclude_sheet_regexes",
    )
    if sheet_limit_action not in {"truncate", "skip", "error"}:
        raise ValueError("sheet_limit_action must be 'truncate', 'skip', or 'error'")
    logger.info(f"Processing Excel file: {excel_path}")

    try:
        # `data_only=True` returns cached values from formulas (paper-aligned).
        # Number-format strings are still preserved on the cell metadata.
        workbook, formula_workbook, cached_workbook = _load_workbooks(excel_path, data_only=data_only)
        logger.info(
            f"Found {len(workbook.sheetnames)} sheets: {', '.join(workbook.sheetnames)}"
        )
    except FileNotFoundError:
        logger.warning(f"Error: File not found: {excel_path}")
        return None
    except ImportError as e:
        logger.warning(f"Error loading Excel file: {e}")
        return None
    except Exception as e:
        logger.warning(f"Error loading Excel file: {e}")
        return None

    sheets_encoding = {}
    compression_metrics = {
        "tokenizer": tokenizer_metadata(tokenizer_model),
        "sheets": {},
    }
    sheet_processing = {
        "mode": (
            "bounded"
            if any(v is not None for v in (max_rows_per_sheet, max_cols_per_sheet, max_cells_per_sheet))
            else "full"
        ),
        "limits": {
            "max_rows_per_sheet": max_rows_per_sheet,
            "max_cols_per_sheet": max_cols_per_sheet,
            "max_cells_per_sheet": max_cells_per_sheet,
            "sheet_limit_action": sheet_limit_action,
        },
        "selection": {
            "include_sheets": include_sheets,
            "exclude_sheets": exclude_sheets,
            "include_sheet_globs": include_sheet_globs,
            "exclude_sheet_globs": exclude_sheet_globs,
            "include_sheet_regexes": include_sheet_regexes,
            "exclude_sheet_regexes": exclude_sheet_regexes,
            "included_sheets": [],
            "skipped_sheets": [],
        },
        "sheets": {},
    }
    overall_orig = overall_anchor = overall_index = overall_format = overall_final = 0

    for sheet_name in workbook.sheetnames:
        logger.info(f"\\nProcessing sheet: {sheet_name}")
        original_sheet = workbook[sheet_name]
        include_sheet, selection_reason = _sheet_selection_decision(
            sheet_name,
            include_sheets,
            include_sheet_globs,
            include_sheet_regexes_compiled,
            exclude_sheets,
            exclude_sheet_globs,
            exclude_sheet_regexes_compiled,
        )
        if not include_sheet:
            sheet_processing["sheets"][sheet_name] = {
                "status": "skipped",
                "reason": selection_reason,
                "limit_action": sheet_limit_action,
                "truncated": False,
                "original_rows": original_sheet.max_row or 1,
                "original_cols": original_sheet.max_column or 1,
                "original_cells": (original_sheet.max_row or 1) * (original_sheet.max_column or 1),
                "effective_rows": 0,
                "effective_cols": 0,
                "effective_cells": 0,
                "encoded_range": None,
            }
            sheet_processing["selection"]["skipped_sheets"].append(
                {"sheet_name": sheet_name, "reason": selection_reason}
            )
            logger.info("Skipping sheet '%s': %s", sheet_name, selection_reason)
            continue

        if original_sheet.max_row <= 1 and original_sheet.max_column <= 1:
            logger.info(f"Sheet '{sheet_name}' appears to be empty. Skipping.")
            sheet_processing["sheets"][sheet_name] = {
                "status": "skipped",
                "reason": "sheet appears empty",
                "limit_action": sheet_limit_action,
                "truncated": False,
                "original_rows": original_sheet.max_row or 1,
                "original_cols": original_sheet.max_column or 1,
                "original_cells": (original_sheet.max_row or 1) * (original_sheet.max_column or 1),
                "effective_rows": 0,
                "effective_cols": 0,
                "effective_cells": 0,
                "encoded_range": None,
            }
            sheet_processing["selection"]["skipped_sheets"].append(
                {"sheet_name": sheet_name, "reason": "sheet appears empty"}
            )
            continue

        effective_rows, effective_cols, processing_meta = _sheet_processing_plan(
            original_sheet,
            max_rows_per_sheet=max_rows_per_sheet,
            max_cols_per_sheet=max_cols_per_sheet,
            max_cells_per_sheet=max_cells_per_sheet,
            sheet_limit_action=sheet_limit_action,
        )
        sheet_processing["sheets"][sheet_name] = processing_meta
        if processing_meta["status"] == "skipped":
            sheet_processing["selection"]["skipped_sheets"].append(
                {
                    "sheet_name": sheet_name,
                    "reason": processing_meta.get(
                        "reason",
                        "sheet skipped (reason not recorded)",
                    ),
                }
            )
            logger.info(
                "Skipping sheet '%s' because it exceeds configured limits: %s rows x %s cols",
                sheet_name,
                processing_meta["original_rows"],
                processing_meta["original_cols"],
            )
            continue

        sheet = original_sheet
        formula_sheet = formula_workbook[sheet_name] if sheet_name in formula_workbook.sheetnames else None
        cached_sheet = cached_workbook[sheet_name] if sheet_name in cached_workbook.sheetnames else None
        if processing_meta["truncated"]:
            logger.info(
                "Truncating sheet '%s' from %s rows x %s cols to %s rows x %s cols",
                sheet_name,
                processing_meta["original_rows"],
                processing_meta["original_cols"],
                effective_rows,
                effective_cols,
            )
            sheet = _copy_bounded_sheet(original_sheet, effective_rows, effective_cols)
            if formula_sheet is not None:
                formula_sheet = _copy_bounded_sheet(formula_sheet, effective_rows, effective_cols)
            if cached_sheet is not None:
                cached_sheet = _copy_bounded_sheet(cached_sheet, effective_rows, effective_cols)

        logger.info(
            f"Sheet dimensions: {sheet.max_row} rows × {sheet.max_column} columns"
        )
        # print memory usage
        logger.info(f"Estimated memory usage: {sys.getsizeof(sheet)} bytes")

        # --- gather original tokens via the paper's vanilla prompt format ---
        # The paper baseline encodes every cell (including empty ones) in the
        # bounding box as ``A1,value|...`` row-major pairs, then counts tokens
        # with the model tokenizer.
        vanilla_prompt = paper_serializers.to_paper_vanilla_prompt(sheet)
        original_tokens = count_tokens(vanilla_prompt, model=tokenizer_model)

        row_anchors, col_anchors = find_structural_anchors(sheet, k)
        logger.info(
            f"Found {len(row_anchors)} row anchors and {len(col_anchors)} column anchors"
        )

        kept_rows, kept_cols = extract_cells_near_anchors(sheet, row_anchors, col_anchors, 0)

        if compress_homogeneous:
            kept_rows, kept_cols = compress_homogeneous_regions(sheet, kept_rows, kept_cols)
            logger.info(
                f"After compression: {len(kept_rows)} rows and {len(kept_cols)} columns kept"
            )

        # Anchor-stage tokens: the vanilla pair-string restricted to retained
        # rows/cols. Empty cells inside the retained skeleton are still emitted
        # so the count is comparable to the paper's vanilla baseline.
        anchor_parts = []
        for r in kept_rows:
            for c in kept_cols:
                ref = f"{get_column_letter(c)}{r}"
                val = sheet.cell(row=r, column=c).value
                text = "" if val is None else str(val).replace("|", " ").replace("\n", " ")
                anchor_parts.append(f"{ref},{text}")
        anchor_prompt = "|".join(anchor_parts)
        anchor_tokens = count_tokens(anchor_prompt, model=tokenizer_model)

        inverted_index, format_map = create_inverted_index(
            sheet, kept_rows, kept_cols, format_mode="paper"
        )
        logger.info(
            f"Created inverted index with {len(inverted_index)} unique values"
        )

        merged_index = create_inverted_index_translation(inverted_index)
        logger.info(
            f"Merged values into {len(merged_index)} range groups"
        )
        # Inverted-index stage tokens: rendered as paper tuples
        # ``(value|range)`` (no format substitution yet).
        index_only_encoding = {"cells": merged_index, "formats": {}}
        index_prompt = paper_serializers.to_paper_compressed_prompt(index_only_encoding)
        index_tokens = count_tokens(index_prompt, model=tokenizer_model)

        # Create a paper-format map from semantic keys to cell references. Older
        # callers may still pass rich-style keys, so keep a compatibility path.
        type_nfs_map = defaultdict(list)
        for fmt_key, cells in format_map.items():
            try:
                fmt = json.loads(fmt_key)
            except Exception:
                fmt = {}
            if set(("type", "nfs")).issubset(fmt.keys()):
                type_nfs_map[fmt_key].extend(cells)
                continue
            for cell_ref in cells:
                try:
                    cell = sheet[cell_ref]
                except Exception:
                    continue
                type_nfs_map[_paper_format_key(cell)].append(cell_ref)

        aggregated_formats = aggregate_regions_dfs(sheet, type_nfs_map)
        logger.info(
            f"Aggregated {len(aggregated_formats)} format regions"
        )

        numeric_map = {
            fmt: cells
            for fmt, cells in type_nfs_map.items()
            if json.loads(fmt).get("type") in ["numeric", "integer", "float"]
        }
        numeric_ranges = aggregate_regions_dfs(sheet, numeric_map)
        logger.info(f"Clustered {len(numeric_ranges)} numeric format ranges")

        # Coordinate remapping (paper Section 3.3.1): retained rows/cols are
        # remapped to a continuous compact grid so the LLM sees A1, A2, … with
        # no gaps. The inverse map lets predicted compact ranges round-trip
        # back to original workbook addresses.
        coord_map = paper_serializers.build_coord_map(kept_rows, kept_cols)

        sheet_encoding = {
            "structural_anchors": {
                "rows": row_anchors,
                "columns": [get_column_letter(c) for c in col_anchors]
            },
            "cells": merged_index,
            "formats": aggregated_formats,
            "numeric_ranges": numeric_ranges,
            "coord_map": coord_map,
            "encoding_mode": "paper_strict" if paper_strict else "pragmatic",
        }
        if formula_sheet is not None:
            formula_graph = extract_formula_graph(
                formula_sheet,
                cached_sheet,
            )
            if (
                formula_graph["formulas"]
                or formula_graph["formula_errors"]
                or formula_graph["repeated_formula_summaries"]
            ):
                sheet_encoding["formula_graph"] = formula_graph

        # Final stage tokens: the paper-faithful compressed prompt with format
        # substitution and compact-coordinate remapping applied.
        final_prompt = paper_serializers.to_paper_compressed_prompt(
            sheet_encoding, coord_map=coord_map
        )
        format_tokens = count_tokens(
            paper_serializers.to_paper_compressed_prompt(sheet_encoding),
            model=tokenizer_model,
        )
        final_tokens = count_tokens(final_prompt, model=tokenizer_model)

        ratio_anchor = calculate_compression_ratio(original_tokens, anchor_tokens)
        ratio_index = calculate_compression_ratio(original_tokens, index_tokens)
        ratio_format = calculate_compression_ratio(original_tokens, format_tokens)
        ratio_final = calculate_compression_ratio(original_tokens, final_tokens)

        compression_metrics["sheets"][sheet_name] = {
            "original_tokens": original_tokens,
            "after_anchor_tokens": anchor_tokens,
            "after_inverted_index_tokens": index_tokens,
            "after_format_tokens": format_tokens,
            "final_tokens": final_tokens,
            "anchor_ratio": ratio_anchor,
            "inverted_index_ratio": ratio_index,
            "format_ratio": ratio_format,
            "overall_ratio": ratio_final,
        }

        logger.info(
            f"{sheet_name} compression - Anchors: {ratio_anchor:.2f}x, "
            f"Index: {ratio_index:.2f}x, Formats: {ratio_format:.2f}x, "
            f"Overall: {ratio_final:.2f}x"
        )

        sheets_encoding[sheet_name] = sheet_encoding
        sheet_processing["selection"]["included_sheets"].append(sheet_name)

        overall_orig += original_tokens
        overall_anchor += anchor_tokens
        overall_index += index_tokens
        overall_format += format_tokens
        overall_final += final_tokens

    compression_metrics["overall"] = {
        "original_tokens": overall_orig,
        "after_anchor_tokens": overall_anchor,
        "after_inverted_index_tokens": overall_index,
        "after_format_tokens": overall_format,
        "final_tokens": overall_final,
        "anchor_ratio": calculate_compression_ratio(overall_orig, overall_anchor),
        "inverted_index_ratio": calculate_compression_ratio(overall_orig, overall_index),
        "format_ratio": calculate_compression_ratio(overall_orig, overall_format),
        "overall_ratio": calculate_compression_ratio(overall_orig, overall_final),
    }

    logger.info(
        f"Overall compression: {compression_metrics['overall']['overall_ratio']:.2f}x"
    )

    full_encoding = {
        "file_name": os.path.basename(excel_path),
        "sheets": sheets_encoding,
        "compression_metrics": compression_metrics,
        "sheet_processing": sheet_processing,
    }

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_encoding, f, indent=2, ensure_ascii=False)
        logger.info(f"Saved SpreadsheetLLM encoding to {output_path}")

    return full_encoding


def get_cell_style_key(cell):
    """Creates a hashable key representing a cell's style for comparison."""
    if not cell:
        return "no_cell"

    font = cell.font
    border = cell.border
    fill = cell.fill
    alignment = cell.alignment

    # Create a tuple of style attributes. Tuples are hashable.
    style_tuple = (
        (font.bold, font.italic, font.underline, font.sz, str(font.color.rgb if font.color else None)),
        (border.left.style, border.right.style, border.top.style, border.bottom.style),
        (fill.patternType, str(fill.fgColor.rgb if fill.fgColor else None)),
        (alignment.horizontal, alignment.vertical, alignment.wrap_text)
    )
    return style_tuple


def _is_year_like(value):
    """Return True for common spreadsheet header years."""
    if isinstance(value, int):
        return 1900 <= value <= 2100
    if isinstance(value, float) and value.is_integer():
        return 1900 <= int(value) <= 2100
    if isinstance(value, str) and value.strip().isdigit():
        return 1900 <= int(value.strip()) <= 2100
    return False


def is_header_row(sheet, row_idx):
    """More robust heuristics to detect header rows, as per Appendix C."""
    num_populated = 0
    num_bold = 0
    num_all_caps = 0
    num_strings = 0
    num_centered = 0
    num_numeric = 0
    num_year_or_date = 0
    unique_values = set()

    for c in range(1, sheet.max_column + 1):
        cell = sheet.cell(row=row_idx, column=c)
        if cell.value is None or str(cell.value).strip() == "":
            continue

        num_populated += 1
        unique_values.add(str(cell.value).strip())
        if cell.font and cell.font.bold:
            num_bold += 1
        if cell.alignment and cell.alignment.horizontal == 'center':
            num_centered += 1

        sem_type = detect_semantic_type(cell)
        if sem_type in {"numeric", "integer", "float", "percentage", "currency"}:
            num_numeric += 1
        if sem_type in {"year", "date", "datetime"} or _is_year_like(cell.value):
            num_year_or_date += 1

        if isinstance(cell.value, str):
            num_strings += 1
            if cell.value.isupper() and len(cell.value) > 1:
                num_all_caps += 1

    if num_populated == 0:
        return False

    # A high proportion of bolded, centered, or all-caps text cells are strong indicators.
    if num_bold / num_populated > 0.6:
        return True
    if num_centered / num_populated > 0.6:
        return True
    if num_strings > 0 and num_all_caps / num_strings > 0.6:
        return True

    # Plain text headers in benchmark spreadsheets are often not styled.
    # Require at least two populated cells so single-cell titles/notes do not
    # become table headers just because they contain text.
    if (
        num_populated >= 2
        and num_strings / num_populated >= 0.5
        # Plain text headers should not accept ordinary data rows such as
        # ["West", 100, "Mia"], but should still accept rare numeric/date labels.
        and (
            num_numeric == 0
            or num_numeric / num_populated <= 0.1
            or (num_year_or_date > 0 and num_numeric / num_populated <= 0.5)
        )
        and len(unique_values) > 1
    ):
        return True

    # Year/date rows are common spreadsheet headers even when values are typed
    # as numbers or dates rather than strings.
    if num_populated >= 2 and num_year_or_date / num_populated >= 0.5:
        return True

    return False


def _cell_profile(cell, merged_coordinates):
    """Compact profile for structural boundary comparisons."""
    value = cell.value
    populated = value is not None and str(value).strip() != ""
    text_shape = None
    if isinstance(value, str):
        stripped = value.strip()
        if stripped.isupper() and len(stripped) > 1:
            text_shape = "upper"
        elif stripped.istitle():
            text_shape = "title"
        elif stripped:
            text_shape = "text"
    return (
        populated,
        detect_semantic_type(cell) if populated else "empty",
        text_shape,
        cell.coordinate in merged_coordinates,
        get_cell_style_key(cell),
    )


def _range_stats(sheet, r1, c1, r2, c2):
    """Return density and text/number proportions for a candidate range."""
    total = (r2 - r1 + 1) * (c2 - c1 + 1)
    populated = text = numeric = year_or_date = 0
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = sheet.cell(row=r, column=c)
            if cell.value is None or str(cell.value).strip() == "":
                continue
            populated += 1
            sem_type = detect_semantic_type(cell)
            if sem_type in {"numeric", "integer", "float", "percentage", "currency"}:
                numeric += 1
            if sem_type in {"year", "date", "datetime"} or _is_year_like(cell.value):
                year_or_date += 1
            if isinstance(cell.value, str):
                text += 1
    return {
        "density": populated / total if total else 0,
        "populated": populated,
        "text_ratio": text / populated if populated else 0,
        "numeric_ratio": numeric / populated if populated else 0,
        "year_or_date_ratio": year_or_date / populated if populated else 0,
    }


def _edge_density(sheet, r1, c1, r2, c2):
    edge_cells = []
    for c in range(c1, c2 + 1):
        edge_cells.append(sheet.cell(row=r1, column=c))
        if r2 != r1:
            edge_cells.append(sheet.cell(row=r2, column=c))
    for r in range(r1 + 1, r2):
        edge_cells.append(sheet.cell(row=r, column=c1))
        if c2 != c1:
            edge_cells.append(sheet.cell(row=r, column=c2))
    if not edge_cells:
        return 0
    populated = sum(
        1 for cell in edge_cells
        if cell.value is not None and str(cell.value).strip() != ""
    )
    return populated / len(edge_cells)


def is_populated_cell(cell):
    return cell.value is not None and str(cell.value).strip() != ""


def _populated_count_in_row(sheet, row_idx, c1=None, c2=None):
    start = c1 if c1 is not None else 1
    end = c2 if c2 is not None else sheet.max_column
    return sum(1 for c in range(start, end + 1) if is_populated_cell(sheet.cell(row=row_idx, column=c)))


def _populated_count_in_col(sheet, col_idx, r1=None, r2=None):
    start = r1 if r1 is not None else 1
    end = r2 if r2 is not None else sheet.max_row
    return sum(1 for r in range(start, end + 1) if is_populated_cell(sheet.cell(row=r, column=col_idx)))


def _row_density(sheet, row_idx, c1, c2):
    width = c2 - c1 + 1
    return _populated_count_in_row(sheet, row_idx, c1, c2) / width if width else 0


def _col_density(sheet, col_idx, r1, r2):
    height = r2 - r1 + 1
    return _populated_count_in_col(sheet, col_idx, r1, r2) / height if height else 0


def _cell_is_in_merged_range(sheet, cell):
    return any(cell.coordinate in merged_range for merged_range in sheet.merged_cells.ranges)


def _row_text_numeric_counts(sheet, row_idx, c1, c2):
    text = numeric = populated = 0
    for c in range(c1, c2 + 1):
        cell = sheet.cell(row=row_idx, column=c)
        if not is_populated_cell(cell):
            continue
        populated += 1
        sem_type = detect_semantic_type(cell)
        if sem_type in {"numeric", "integer", "float", "percentage", "currency"}:
            numeric += 1
        if isinstance(cell.value, str):
            text += 1
    return populated, text, numeric


def _looks_like_title_or_note_row(sheet, row_idx, c1, c2):
    """Return True for sparse descriptive rows that should not be table headers."""
    populated, text, numeric = _row_text_numeric_counts(sheet, row_idx, c1, c2)
    if populated == 0:
        return False
    width = c2 - c1 + 1
    if numeric > 0:
        return False
    if text == 0:
        return False
    # Title/note rows are normally sparse descriptive text spanning less than
    # half the table width, unless style/merge cues make the role explicit.
    sparse_text = populated <= max(1, int(width * SPARSE_TEXT_MAX_RATIO))
    styled_or_merged = False
    for c in range(c1, c2 + 1):
        cell = sheet.cell(row=row_idx, column=c)
        if not is_populated_cell(cell):
            continue
        if (
            (cell.font and (cell.font.bold or cell.font.italic))
            or (cell.alignment and cell.alignment.horizontal == "center")
            or _cell_is_in_merged_range(sheet, cell)
        ):
            styled_or_merged = True
            break
    return sparse_text or styled_or_merged


def _header_rows_in_range(sheet, r1, c1, r2, c2):
    """Find header rows whose populated cells align with a candidate rectangle."""
    headers = []
    for r in range(r1, r2 + 1):
        if _populated_count_in_row(sheet, r, c1, c2) == 0:
            continue
        if not is_header_row(sheet, r):
            continue
        if (
            _looks_like_title_or_note_row(sheet, r, c1, c2)
            and _row_density(sheet, r, c1, c2) < SPARSE_TEXT_MAX_RATIO
        ):
            continue
        headers.append(r)
    return headers


def _candidate_table_profile(sheet, r1, c1, r2, c2):
    """Classify candidate rows into title/note, header, and body evidence."""
    headers = _header_rows_in_range(sheet, r1, c1, r2, c2)
    if not headers:
        return None

    header_row = headers[0]
    prefix_rows = range(r1, header_row)
    if any(
        _populated_count_in_row(sheet, r, c1, c2) > 0
        and not _looks_like_title_or_note_row(sheet, r, c1, c2)
        for r in prefix_rows
    ):
        return None

    data_rows = [
        r
        for r in range(header_row + 1, r2 + 1)
        if _populated_count_in_row(sheet, r, c1, c2) > 0
    ]
    if not data_rows:
        if _populated_count_in_row(sheet, header_row, c1, c2) < 2:
            return None
        return {
            "header_row": header_row,
            "body_rows": [header_row],
            "note_rows": [],
            "title_rows": list(prefix_rows),
            "populated_cols": _populated_cols_in_row(sheet, header_row),
        }

    body_rows = []
    note_rows = []
    for r in data_rows:
        if (
            _looks_like_title_or_note_row(sheet, r, c1, c2)
            and _row_density(sheet, r, c1, c2) < SPARSE_TEXT_MAX_RATIO
        ):
            note_rows.append(r)
        else:
            body_rows.append(r)

    if not body_rows:
        return None
    if note_rows:
        first_note = min(note_rows)
        last_body = max(body_rows)
        if any(r > first_note for r in body_rows):
            return None
        if first_note > last_body + 1 and any(
            _populated_count_in_row(sheet, r, c1, c2) == 0
            for r in range(last_body + 1, first_note)
        ):
            return None
        if len(note_rows) > MAX_TRAILING_NOTE_ROWS:
            # More than two trailing notes usually means the rectangle swallowed
            # unrelated prose rather than a compact table footnote.
            return None

    populated_cols = [
        c
        for c in range(c1, c2 + 1)
        if _populated_count_in_col(sheet, c, header_row, max(body_rows)) > 0
    ]
    if len(populated_cols) < 2:
        return None

    # Reject wrappers where more than half of the candidate width is blank
    # between the header and body; those usually bridge separate side-by-side tables.
    sparse_internal_cols = [
        c
        for c in range(c1, c2 + 1)
        if _col_density(sheet, c, header_row, max(body_rows)) == 0
    ]
    if len(sparse_internal_cols) / (c2 - c1 + 1) > MAX_SPARSE_COLUMN_RATIO:
        return None

    return {
        "header_row": header_row,
        "body_rows": body_rows,
        "note_rows": note_rows,
        "title_rows": list(prefix_rows),
        "populated_cols": populated_cols,
    }


def _populated_cols_in_row(sheet, row_idx):
    cols = []
    for col_idx in range(1, sheet.max_column + 1):
        value = sheet.cell(row=row_idx, column=col_idx).value
        if value is not None and str(value).strip() != "":
            cols.append(col_idx)
    return cols


def _contiguous_groups(indices):
    if not indices:
        return []
    groups = []
    start = prev = indices[0]
    for idx in indices[1:]:
        if idx == prev + 1:
            prev = idx
            continue
        groups.append((start, prev))
        start = prev = idx
    groups.append((start, prev))
    return groups


def _bounded_header_region_candidates(sheet):
    """Find table-like rectangles without composing every boundary pair.

    The Appendix C-inspired boundary search can become expensive when many
    rows have unique profiles. For larger candidate grids, use styled/header
    rows as seeds and grow each contiguous header band downward until the band
    becomes blank.
    """
    candidates = []
    for row_idx in range(1, sheet.max_row + 1):
        if not is_header_row(sheet, row_idx):
            continue
        for c1, c2 in _contiguous_groups(_populated_cols_in_row(sheet, row_idx)):
            end_row = row_idx
            for data_row in range(row_idx + 1, sheet.max_row + 1):
                populated = False
                for col_idx in range(c1, c2 + 1):
                    value = sheet.cell(row=data_row, column=col_idx).value
                    if value is not None and str(value).strip() != "":
                        populated = True
                        break
                if not populated:
                    break
                end_row = data_row
            if end_row > row_idx and c2 > c1:
                candidates.append((row_idx, c1, end_row, c2))
    return candidates


def _header_region_candidates(sheet):
    """Compose table rectangles from header/title bands and contiguous columns."""
    candidates = []
    for row_idx in range(1, sheet.max_row + 1):
        if not is_header_row(sheet, row_idx):
            continue
        populated_cols = _populated_cols_in_row(sheet, row_idx)
        if len(populated_cols) < 2:
            continue
        for c1, c2 in _contiguous_groups(populated_cols):
            if c2 <= c1:
                continue

            start_row = row_idx
            for title_row in range(row_idx - 1, 0, -1):
                if _populated_count_in_row(sheet, title_row, c1, c2) == 0:
                    break
                if not _looks_like_title_or_note_row(sheet, title_row, c1, c2):
                    break
                start_row = title_row

            end_row = row_idx
            for data_row in range(row_idx + 1, sheet.max_row + 1):
                populated = _populated_count_in_row(sheet, data_row, c1, c2)
                if populated == 0:
                    break
                end_row = data_row
                if (
                    _looks_like_title_or_note_row(sheet, data_row, c1, c2)
                    and _row_density(sheet, data_row, c1, c2) < SPARSE_TEXT_MAX_RATIO
                ):
                    break
            if end_row > row_idx:
                candidates.append((start_row, c1, end_row, c2))
    return candidates


def find_boundary_candidates(sheet):
    """
    Identify row/column boundary candidates using enhanced heterogeneity heuristics
    from Appendix C, including cell value, merged status, and style.
    """
    merged_coordinates = {
        coord for merged_range in sheet.merged_cells.ranges for coord in merged_range
    }

    row_profiles = []
    for r in range(1, sheet.max_row + 1):
        profile = []
        for c in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=r, column=c)
            profile.append(_cell_profile(cell, merged_coordinates))
        row_profiles.append(profile)

    col_profiles = []
    for c in range(1, sheet.max_column + 1):
        profile = []
        for r in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=r, column=c)
            profile.append(_cell_profile(cell, merged_coordinates))
        col_profiles.append(profile)

    row_candidates = set()
    for r in range(1, len(row_profiles)):
        current_populated = _populated_count_in_row(sheet, r)
        next_populated = _populated_count_in_row(sheet, r + 1)
        if row_profiles[r] != row_profiles[r - 1]:
            # Add both sides of the boundary
            row_candidates.add(r)
            row_candidates.add(r + 1)
        if current_populated == 0 and next_populated != 0:
            row_candidates.add(r + 1)
        if current_populated != 0 and next_populated == 0:
            row_candidates.add(r)

    col_candidates = set()
    for c in range(1, len(col_profiles)):
        current_populated = _populated_count_in_col(sheet, c)
        next_populated = _populated_count_in_col(sheet, c + 1)
        if col_profiles[c] != col_profiles[c - 1]:
            col_candidates.add(c)
            col_candidates.add(c + 1)
        if current_populated == 0 and next_populated != 0:
            col_candidates.add(c + 1)
        if current_populated != 0 and next_populated == 0:
            col_candidates.add(c)

    for merged_range in sheet.merged_cells.ranges:
        row_candidates.update([merged_range.min_row, merged_range.max_row])
        col_candidates.update([merged_range.min_col, merged_range.max_col])

    # Step 2: Compose candidate boundaries
    candidates = _header_region_candidates(sheet)
    if row_candidates and col_candidates:
        rows = sorted(list(row_candidates))
        cols = sorted(list(col_candidates))
        candidate_count = (
            (len(rows) * (len(rows) - 1) // 2)
            * (len(cols) * (len(cols) - 1) // 2)
        )
        if candidate_count > 2_000:
            candidates.extend(_bounded_header_region_candidates(sheet))
        else:
            for row_start_pos in range(len(rows)):
                for row_end_pos in range(row_start_pos + 1, len(rows)):
                    for col_start_pos in range(len(cols)):
                        for col_end_pos in range(col_start_pos + 1, len(cols)):
                            candidates.append((
                                rows[row_start_pos],
                                cols[col_start_pos],
                                rows[row_end_pos],
                                cols[col_end_pos],
                            ))
    candidates = sorted(set(candidates))

    # Step 3: Filter unreasonable candidates
    candidates = filter_unreasonable_candidates(sheet, candidates)

    # Step 4: Filter overlapping candidates
    candidates = filter_overlapping_candidates(sheet, candidates)

    # Step 5: Derive anchors from final candidates
    final_row_anchors = set()
    final_col_anchors = set()
    for r1, c1, r2, c2 in candidates:
        final_row_anchors.add(r1)
        final_row_anchors.add(r2)
        final_col_anchors.add(c1)
        final_col_anchors.add(c2)
        profile = _candidate_table_profile(sheet, r1, c1, r2, c2)
        if profile is not None:
            if (r2 - r1 + 1) <= MAX_DETAILED_ANCHOR_DIMENSION:
                final_row_anchors.update(profile["title_rows"])
                final_row_anchors.add(profile["header_row"])
                final_row_anchors.update(profile["body_rows"])
                final_row_anchors.update(profile["note_rows"])
            if (c2 - c1 + 1) <= MAX_DETAILED_ANCHOR_DIMENSION:
                final_col_anchors.update(profile["populated_cols"])
        if r1 > 1 and _populated_count_in_row(sheet, r1 - 1, c1, c2) == 0:
            final_row_anchors.add(r1 - 1)
        if r2 < sheet.max_row and _populated_count_in_row(sheet, r2 + 1, c1, c2) == 0:
            final_row_anchors.add(r2 + 1)
        if c1 > 1 and _populated_count_in_col(sheet, c1 - 1, r1, r2) == 0:
            final_col_anchors.add(c1 - 1)
        if c2 < sheet.max_column and _populated_count_in_col(sheet, c2 + 1, r1, r2) == 0:
            final_col_anchors.add(c2 + 1)

    return sorted(list(final_row_anchors)), sorted(list(final_col_anchors))


def filter_unreasonable_candidates(sheet, candidates):
    """Filter out candidates based on size, sparsity, and header presence."""
    filtered = []
    for r1, c1, r2, c2 in candidates:
        # Size filter
        if (r2 - r1 < 1) or (c2 - c1 < 1):
            continue  # Must have at least 2 rows/cols

        stats = _range_stats(sheet, r1, c1, r2, c2)

        profile = _candidate_table_profile(sheet, r1, c1, r2, c2)
        if profile is None:
            continue

        header_row = profile["header_row"]
        body_end = max(profile["body_rows"])
        body_stats = _range_stats(sheet, header_row, c1, body_end, c2)

        # Internal sparsity filter. The full candidate may include sparse title
        # or trailing note rows, so measure table density from header through
        # body and keep a lower whole-range threshold for contextual rows.
        if body_stats["density"] < 0.25 or stats["density"] < 0.1:
            continue

        # Edge sparsity filter. Real table boundaries generally have visible
        # content near at least one edge; this rejects huge sparse rectangles
        # formed by distant notes or isolated cells.
        if _edge_density(sheet, header_row, c1, body_end, c2) < 0.2:
            continue

        sparse_body_rows = [
            r
            for r in range(header_row, body_end + 1)
            if _row_density(sheet, r, c1, c2) == 0
        ]
        if sparse_body_rows:
            continue

        sparse_body_cols = [
            c
            for c in range(c1, c2 + 1)
            if _col_density(sheet, c, header_row, body_end) == 0
        ]
        if sparse_body_cols:
            continue

        # Header and proportion filters. Keep text-header tables, date/year
        # header tables, and numeric-heavy tables only when a header row exists.
        if body_stats["text_ratio"] == 0 and body_stats["year_or_date_ratio"] == 0:
            continue

        filtered.append((r1, c1, r2, c2))

    return filtered


def calculate_iou(box1, box2):
    """Calculate Intersection over Union (IoU) for two bounding boxes."""
    r1_1, c1_1, r2_1, c2_1 = box1
    r1_2, c1_2, r2_2, c2_2 = box2

    inter_r1 = max(r1_1, r1_2)
    inter_c1 = max(c1_1, c1_2)
    inter_r2 = min(r2_1, r2_2)
    inter_c2 = min(c2_1, c2_2)

    inter_area = max(0, inter_r2 - inter_r1 + 1) * max(0, inter_c2 - inter_c1 + 1)

    area1 = (r2_1 - r1_1 + 1) * (c2_1 - c1_1 + 1)
    area2 = (r2_2 - r1_2 + 1) * (c2_2 - c1_2 + 1)

    union_area = area1 + area2 - inter_area

    return inter_area / union_area if union_area > 0 else 0


def _candidate_area(candidate):
    r1, c1, r2, c2 = candidate
    return (r2 - r1 + 1) * (c2 - c1 + 1)


def _intersection_area(box1, box2):
    r1_1, c1_1, r2_1, c2_1 = box1
    r1_2, c1_2, r2_2, c2_2 = box2
    inter_r1 = max(r1_1, r1_2)
    inter_c1 = max(c1_1, c1_2)
    inter_r2 = min(r2_1, r2_2)
    inter_c2 = min(c2_1, c2_2)
    return max(0, inter_r2 - inter_r1 + 1) * max(0, inter_c2 - inter_c1 + 1)


def _overlap_ratio(candidate, other):
    area = _candidate_area(candidate)
    return _intersection_area(candidate, other) / area if area else 0


def _candidate_score(sheet, candidate):
    r1, c1, r2, c2 = candidate
    stats = _range_stats(sheet, r1, c1, r2, c2)
    profile = _candidate_table_profile(sheet, r1, c1, r2, c2)
    if profile is None:
        return -1

    header_row = profile["header_row"]
    body_end = max(profile["body_rows"])
    body_stats = _range_stats(sheet, header_row, c1, body_end, c2)
    width = c2 - c1 + 1
    height = r2 - r1 + 1
    populated = stats["populated"]
    # Prefer candidates with an early header, dense body, enough body rows and
    # populated cells, while only lightly rewarding contextual title/note rows.
    score = 0
    score += HEADER_AT_TOP_BONUS if header_row == r1 else HEADER_AFTER_TITLE_BONUS
    score += min(MAX_BODY_ROW_BONUS, len(profile["body_rows"]) * BODY_ROW_BONUS)
    score += min(MAX_WIDTH_BONUS, width)
    score += body_stats["density"] * BODY_DENSITY_WEIGHT
    score += stats["density"] * RANGE_DENSITY_WEIGHT
    score += stats["year_or_date_ratio"] * YEAR_OR_DATE_WEIGHT
    score += min(MAX_POPULATED_CELL_BONUS, populated)
    score += TITLE_ROW_BONUS * len(profile["title_rows"])
    score += NOTE_ROW_BONUS * len(profile["note_rows"])
    score -= (
        max(0, height - len(profile["body_rows"]) - len(profile["title_rows"]) - 1)
        * EXTRA_ROW_PENALTY
    )
    return score


def filter_overlapping_candidates(sheet, candidates):
    """Filter overlapping candidates using heuristics from Appendix C."""
    if not candidates:
        return []

    scores = [_candidate_score(sheet, candidate) for candidate in candidates]

    # Non-maximum suppression based on IoU and scores
    indices = list(range(len(candidates)))
    indices.sort(key=lambda i: (scores[i], _candidate_area(candidates[i])), reverse=True)

    keep = []
    while indices:
        current_idx = indices.pop(0)
        keep.append(current_idx)

        remaining_indices = []
        for idx in indices:
            iou = calculate_iou(candidates[current_idx], candidates[idx])
            overlap_current = _overlap_ratio(candidates[current_idx], candidates[idx])
            overlap_other = _overlap_ratio(candidates[idx], candidates[current_idx])
            # Suppress either broad IoU overlaps or near-containment in either
            # direction so sparse wrappers do not coexist with their inner table.
            if (
                iou < OVERLAP_IOU_SUPPRESSION_THRESHOLD
                and overlap_current < OVERLAP_CONTAINMENT_SUPPRESSION_THRESHOLD
                and overlap_other < OVERLAP_CONTAINMENT_SUPPRESSION_THRESHOLD
            ):
                remaining_indices.append(idx)
        indices = remaining_indices

    kept = [candidates[i] for i in keep]
    return sorted(kept, key=lambda box: (box[0], box[1], box[2], box[3]))


def extract_k_neighborhood(indices, k, max_index):
    """Expand indices with a k-neighborhood within bounds."""
    expanded = set()
    for idx in indices:
        for i in range(max(1, idx - k), min(max_index + 1, idx + k + 1)):
            expanded.add(i)
    return sorted(expanded)


def find_structural_anchors(sheet, k=4):
    """Find structural anchors using boundary candidates and k-neighborhood.

    Default ``k=4`` matches the paper's best ablation setting (Section 3.3.1).
    """
    row_candidates, col_candidates = find_boundary_candidates(sheet)
    row_anchors = extract_k_neighborhood(row_candidates, k, sheet.max_row)
    col_anchors = extract_k_neighborhood(col_candidates, k, sheet.max_column)
    return row_anchors, col_anchors


def extract_cells_near_anchors(sheet, row_anchors, col_anchors, k):
    """Extract cells within k units of any anchor."""
    rows_to_keep = set()
    cols_to_keep = set()

    for r in row_anchors:
        for i in range(max(1, r - k), min(sheet.max_row + 1, r + k + 1)):
            rows_to_keep.add(i)

    for c in col_anchors:
        for i in range(max(1, c - k), min(sheet.max_column + 1, c + k + 1)):
            cols_to_keep.add(i)

    return sorted(list(rows_to_keep)), sorted(list(cols_to_keep))


def compress_homogeneous_regions(sheet, rows, cols):
    """Remove rows and columns that are homogeneous in value and format."""
    def row_homogeneous(r):
        vals = []
        fmts = []
        for c in cols:
            cell = sheet.cell(row=r, column=c)
            vals.append(cell.value)
            fmts.append(cell.number_format)
        return len(set(vals)) <= 1 and len(set(fmts)) <= 1

    def col_homogeneous(c):
        vals = []
        fmts = []
        for r in rows:
            cell = sheet.cell(row=r, column=c)
            vals.append(cell.value)
            fmts.append(cell.number_format)
        return len(set(vals)) <= 1 and len(set(fmts)) <= 1

    filtered_rows = [r for r in rows if not row_homogeneous(r)]
    filtered_cols = [c for c in cols if not col_homogeneous(c)]
    return filtered_rows, filtered_cols


def _paper_format_key(cell):
    """Return the paper-faithful semantic format key for a cell."""
    nfs = get_number_format_string(cell)
    sem_type = detect_semantic_type(cell)
    return json.dumps({"type": sem_type, "nfs": nfs}, sort_keys=True)


def _rich_format_key(cell, merged_range=None):
    """Return the legacy rich-style format key for experimental analysis."""
    format_info = {}

    # 1. Font Styles
    font = cell.font
    format_info["font"] = {
        "bold": font.bold,
        "italic": font.italic,
        "underline": font.underline,
        "name": font.name,
        "size": font.sz,
        "color": str(font.color.rgb) if font.color and font.color.rgb else None,
    }

    # 2. Alignment
    alignment = cell.alignment
    format_info["alignment"] = {
        "horizontal": alignment.horizontal,
        "vertical": alignment.vertical,
    }

    # 3. Borders
    border = cell.border
    format_info["border"] = {
        side: {
            "style": getattr(border, side).style,
            "color": (
                str(getattr(border, side).color.rgb)
                if getattr(border, side).color and getattr(border, side).color.rgb
                else None
            ),
        }
        for side in ["left", "right", "top", "bottom"]
    }

    # 4. Fill (Background Color)
    fill = cell.fill
    if hasattr(fill, 'patternType') and fill.patternType == "solid":
        format_info["fill"] = {"color": str(fill.start_color.index)
                               if fill.start_color and fill.start_color.index else None}
    else:
        format_info["fill"] = {"color": None}

    # 5. Number Format (Original, Inferred Type, Category)
    original_number_format = cell.number_format
    inferred_type = infer_cell_data_type(cell)
    category = categorize_number_format(original_number_format, cell)

    format_info["original_number_format"] = original_number_format
    format_info["inferred_data_type"] = inferred_type
    format_info["number_format_category"] = category

    if merged_range is not None:
        format_info["merged"] = True
        format_info["merged_range"] = str(merged_range)
    else:
        format_info["merged"] = False

    return json.dumps(format_info, sort_keys=True)


def create_inverted_index(sheet, kept_rows, kept_cols, format_mode="paper"):
    """Create an inverted index, handling merged cells.

    ``format_mode="paper"`` groups cells only by semantic type and Excel
    number-format string, which matches SheetCompressor's data-format-aware
    aggregation. ``format_mode="rich"`` preserves the older style-heavy key for
    experiments, but rich style metadata is intentionally not the default
    paper path.
    """
    if format_mode not in {"paper", "rich"}:
        raise ValueError("format_mode must be 'paper' or 'rich'")

    inverted_index = defaultdict(list)
    format_map = defaultdict(list)
    merged_ranges = sheet.merged_cells.ranges  # get all merged cell ranges

    for row in kept_rows:
        for col in kept_cols:
            cell = sheet.cell(row=row, column=col)
            cell_ref = f"{get_column_letter(col)}{row}"

            # Merged Cell Handling
            merged_value = None
            merged_range = None
            merged_anchor_cell = None
            for m_range in merged_ranges:
                if cell_ref in m_range:
                    try:
                        merged_anchor_cell = sheet[m_range.start_cell.coordinate]
                        merged_value = merged_anchor_cell.value
                        merged_range = m_range
                        break
                    except Exception:
                        pass  # Skip if there's an issue with the merged range

            # Use merged value if available, otherwise cell value
            try:
                if merged_value is not None:
                    cell_value = str(merged_value) if merged_value is not None else ""
                    inverted_index[cell_value].append(cell_ref)
                elif cell.value is not None:
                    if isinstance(cell.value, (int, float)):
                        cell_value = f"{cell.value}"
                    else:
                        cell_value = str(cell.value)
                    inverted_index[cell_value].append(cell_ref)
            except Exception as e:
                # Handle error for problematic cell values
                logger.warning(f"Error processing cell {cell_ref}: {e}")
                cell_value = "ERROR_VALUE"
                inverted_index[cell_value].append(cell_ref)

            # Format Handling
            try:
                format_cell = merged_anchor_cell if merged_anchor_cell is not None else cell
                if format_mode == "paper":
                    format_key = _paper_format_key(format_cell)
                else:
                    format_key = _rich_format_key(format_cell, merged_range)
                format_map[format_key].append(cell_ref)
            except Exception as e:
                # Handle error for problematic cell formats
                logger.warning(f"Error processing format for cell {cell_ref}: {e}")

    return dict(inverted_index), dict(format_map)


def create_inverted_index_translation(inverted_index):
    """Merge cell references for identical values into ranges.

    Args:
        inverted_index (dict): Mapping of values to lists of cell references.

    Returns:
        dict: Mapping of values to merged cell ranges.
    """

    def _merge_refs(refs):
        coords = []
        for ref in sorted(set(refs)):
            try:
                col_letter, row = split_cell_ref(ref)
                col = openpyxl.utils.cell.column_index_from_string(col_letter)
                coords.append((row, col))
            except Exception:
                continue

        cell_set = set(coords)
        processed = set()
        ranges = []

        for row, col in sorted(coords):
            if (row, col) in processed:
                continue

            width = 1
            while (row, col + width) in cell_set and (row, col + width) not in processed:
                width += 1

            height = 1
            expanding = True
            while expanding:
                next_row = row + height
                for w in range(width):
                    if (next_row, col + w) not in cell_set or (next_row, col + w) in processed:
                        expanding = False
                        break
                if expanding:
                    height += 1

            end_col = col + width - 1
            end_row = row + height - 1
            start_ref = f"{get_column_letter(col)}{row}"
            end_ref = f"{get_column_letter(end_col)}{end_row}"

            if width == 1 and height == 1:
                ranges.append(start_ref)
            else:
                ranges.append(f"{start_ref}:{end_ref}")

            for r in range(row, row + height):
                for c in range(col, col + width):
                    processed.add((r, c))

        return ranges

    merged_index = {}
    for value, refs in inverted_index.items():
        if value is None or str(value).strip() == "":
            continue
        merged_index[value] = _merge_refs(refs)

    return merged_index


def aggregate_regions_dfs(sheet, format_map):
    """Aggregate connected cells by semantic key using DFS.

    Implements Algorithm 1 from Appendix M.1: for each ``{type, nfs}`` group,
    find 4-connected components in the address grid and emit each component
    as one or more rectangle ranges. Rows that share an identical column band
    with the previous row are merged vertically.
    """
    aggregated_regions = {}

    for key, cells in format_map.items():
        coords = set()
        for cell_ref in cells:
            try:
                col_letter, row = split_cell_ref(cell_ref)
                col = get_column_index(col_letter)
            except Exception:
                continue
            coords.add((row, col))

        if not coords:
            aggregated_regions[key] = []
            continue

        visited = set()
        regions = []

        for start in sorted(coords):
            if start in visited:
                continue

            stack = [start]
            component = set()

            while stack:
                row, col = stack.pop()
                if (row, col) in visited or (row, col) not in coords:
                    continue
                visited.add((row, col))
                component.add((row, col))
                stack.extend([
                    (row - 1, col),
                    (row + 1, col),
                    (row, col - 1),
                    (row, col + 1),
                ])

            component_rows = {}
            for row, col in sorted(component):
                component_rows.setdefault(row, []).append(col)

            pending = []
            for row in sorted(component_rows):
                cols = sorted(component_rows[row])
                start_col = cols[0]
                prev_col = cols[0]
                for col in cols[1:]:
                    if col == prev_col + 1:
                        prev_col = col
                        continue
                    pending.append([row, row, start_col, prev_col])
                    start_col = col
                    prev_col = col
                pending.append([row, row, start_col, prev_col])

            merged = []
            for row_start, row_end, col_start, col_end in pending:
                extended = False
                for existing in merged:
                    if (
                        existing[2] == col_start
                        and existing[3] == col_end
                        and existing[1] == row_start - 1
                    ):
                        existing[1] = row_end
                        extended = True
                        break
                if not extended:
                    merged.append([row_start, row_end, col_start, col_end])

            for row_start, row_end, col_start, col_end in merged:
                start_ref = f"{get_column_letter(col_start)}{row_start}"
                end_ref = f"{get_column_letter(col_end)}{row_end}"
                if start_ref == end_ref:
                    regions.append(start_ref)
                else:
                    regions.append(f"{start_ref}:{end_ref}")

        aggregated_regions[key] = regions

    return aggregated_regions


def get_column_index(col_letter):
    """Convert column letter to index (A => 1, AA => 27)."""
    return openpyxl.utils.cell.column_index_from_string(col_letter)


def split_cell_ref(cell_ref):
    """Split cell reference (e.g., 'A1') into column letter and row number."""
    col_str = ''.join(filter(str.isalpha, cell_ref))
    row_str = ''.join(filter(str.isdigit, cell_ref))

    # convert row to integer
    return col_str, int(row_str)


def main():
    """Console script entry point for SpreadsheetLLM encoder."""
    logging.basicConfig(level=logging.INFO)
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert Excel files to SpreadsheetLLM format"
    )
    parser.add_argument("excel_file", help="Path to the Excel file")
    parser.add_argument(
        "--output",
        "-o",
        help="Output JSON file path (default: same as input with .json extension)",
    )
    parser.add_argument(
        "--k",
        type=int,
        default=4,
        help="Neighborhood distance parameter (default: 4, paper's best ablation).",
    )
    parser.add_argument(
        "--vanilla",
        action="store_true",
        help="Produce vanilla markdown-like encoding instead of compressed JSON.",
    )
    parser.add_argument(
        "--no-compress-homogeneous",
        action="store_true",
        help="Skip the homogeneous-row/col compression step (paper-strict skeleton).",
    )
    parser.add_argument(
        "--paper-strict",
        action="store_true",
        help=(
            "Use paper-faithful behavior where it differs from pragmatic "
            "defaults; currently disables homogeneous row/column pruning."
        ),
    )
    parser.add_argument(
        "--tokenizer-model",
        default=DEFAULT_MODEL,
        help=f"Model name passed to tiktoken for token counts (default: {DEFAULT_MODEL}).",
    )
    parser.add_argument(
        "--max-rows-per-sheet",
        type=int,
        default=None,
        help="Bounded mode: encode at most this many rows from each sheet.",
    )
    parser.add_argument(
        "--max-cols-per-sheet",
        type=int,
        default=None,
        help="Bounded mode: encode at most this many columns from each sheet.",
    )
    parser.add_argument(
        "--max-cells-per-sheet",
        type=int,
        default=None,
        help=(
            "Bounded mode: encode at most this many cells per sheet by reducing "
            "the effective row count after row/column caps are applied."
        ),
    )
    parser.add_argument(
        "--sheet-limit-action",
        choices=["truncate", "skip", "error"],
        default="truncate",
        help=(
            "Bounded mode behavior for sheets over configured limits: "
            "truncate, skip, or error (default: truncate)."
        ),
    )
    parser.add_argument(
        "--include-sheet",
        action="append",
        default=[],
        help="Include only this exact sheet name. Repeat flag for multiple sheets.",
    )
    parser.add_argument(
        "--exclude-sheet",
        action="append",
        default=[],
        help="Exclude this exact sheet name. Repeat flag for multiple sheets.",
    )
    parser.add_argument(
        "--include-sheet-glob",
        action="append",
        default=[],
        help="Include sheets matching this glob pattern. Repeatable.",
    )
    parser.add_argument(
        "--exclude-sheet-glob",
        action="append",
        default=[],
        help="Exclude sheets matching this glob pattern. Repeatable.",
    )
    parser.add_argument(
        "--include-sheet-regex",
        action="append",
        default=[],
        help="Include sheets whose names match this regex. Repeatable.",
    )
    parser.add_argument(
        "--exclude-sheet-regex",
        action="append",
        default=[],
        help="Exclude sheets whose names match this regex. Repeatable.",
    )

    args = parser.parse_args()

    if not args.output:
        if args.vanilla:
            args.output = os.path.splitext(args.excel_file)[0] + "_vanilla.txt"
        else:
            args.output = os.path.splitext(args.excel_file)[0] + "_spreadsheetllm.json"

    result = spreadsheet_llm_encode(
        args.excel_file,
        args.output,
        k=args.k,
        vanilla=args.vanilla,
        compress_homogeneous=not args.no_compress_homogeneous,
        paper_strict=args.paper_strict,
        tokenizer_model=args.tokenizer_model,
        max_rows_per_sheet=args.max_rows_per_sheet,
        max_cols_per_sheet=args.max_cols_per_sheet,
        max_cells_per_sheet=args.max_cells_per_sheet,
        sheet_limit_action=args.sheet_limit_action,
        include_sheets=args.include_sheet,
        exclude_sheets=args.exclude_sheet,
        include_sheet_globs=args.include_sheet_glob,
        exclude_sheet_globs=args.exclude_sheet_glob,
        include_sheet_regexes=args.include_sheet_regex,
        exclude_sheet_regexes=args.exclude_sheet_regex,
    )

    if result is not None and not args.vanilla:
        metrics = result.get("compression_metrics", {})
        for sheet_name, sm in metrics.get("sheets", {}).items():
            print(
                f"{sheet_name}: {sm.get('overall_ratio', 0.0):.2f}x compression "
                f"(anchors {sm.get('anchor_ratio', 0.0):.2f}x, "
                f"index {sm.get('inverted_index_ratio', 0.0):.2f}x, "
                f"formats {sm.get('format_ratio', 0.0):.2f}x)"
            )
        overall = metrics.get("overall", {})
        if overall:
            print(f"Overall: {overall.get('overall_ratio', 0.0):.2f}x compression")


def vanilla_encode(
    excel_path,
    output_path=None,
    include_sheets=None,
    exclude_sheets=None,
    include_sheet_globs=None,
    exclude_sheet_globs=None,
    include_sheet_regexes=None,
    exclude_sheet_regexes=None,
):
    """Vanilla markdown-like encoding (paper Section 3.1).

    Produces a ``{sheet_name: pair_string}`` dict where each sheet is the
    paper's row-major ``A1,value|A2,value|...`` baseline. When written to
    disk, all sheets are emitted under ``# {sheet_name}`` headers so
    multi-sheet workbooks aren't silently truncated.
    """
    logger.info(f"Producing vanilla encoding for {excel_path}")
    include_sheets = _validate_and_normalize_filter_list(include_sheets, "include_sheets")
    exclude_sheets = _validate_and_normalize_filter_list(exclude_sheets, "exclude_sheets")
    include_sheet_globs = _validate_and_normalize_filter_list(include_sheet_globs, "include_sheet_globs")
    exclude_sheet_globs = _validate_and_normalize_filter_list(exclude_sheet_globs, "exclude_sheet_globs")
    include_sheet_regexes = _validate_and_normalize_filter_list(include_sheet_regexes, "include_sheet_regexes")
    exclude_sheet_regexes = _validate_and_normalize_filter_list(exclude_sheet_regexes, "exclude_sheet_regexes")
    include_sheet_regexes_compiled = _compile_sheet_regexes(
        include_sheet_regexes,
        "include_sheet_regexes",
    )
    exclude_sheet_regexes_compiled = _compile_sheet_regexes(
        exclude_sheet_regexes,
        "exclude_sheet_regexes",
    )
    try:
        workbook, _, _ = _load_workbooks(excel_path, data_only=True)
    except ImportError as e:
        logger.error(f"Error loading Excel file for vanilla encoding: {e}")
        return None
    except Exception as e:
        logger.error(f"Error loading Excel file for vanilla encoding: {e}")
        return None

    vanilla_content = {}
    for sheet_name in workbook.sheetnames:
        include_sheet, _ = _sheet_selection_decision(
            sheet_name,
            include_sheets,
            include_sheet_globs,
            include_sheet_regexes_compiled,
            exclude_sheets,
            exclude_sheet_globs,
            exclude_sheet_regexes_compiled,
        )
        if include_sheet:
            vanilla_content[sheet_name] = paper_serializers.to_paper_vanilla_prompt(
                workbook[sheet_name]
            )

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            for i, (sheet_name, content) in enumerate(vanilla_content.items()):
                if i:
                    f.write("\n\n")
                f.write(f"# {sheet_name}\n")
                f.write(content)
        logger.info(f"Saved vanilla encoding to {output_path}")

    return vanilla_content


if __name__ == "__main__":
    main()
