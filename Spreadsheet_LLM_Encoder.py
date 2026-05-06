import os
import openpyxl
import json
import logging
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
from tokenizer import count_tokens, DEFAULT_MODEL

logger = logging.getLogger(__name__)


def calculate_compression_ratio(original_tokens: int, compressed_tokens: int) -> float:
    """Return the compression ratio given original and compressed token counts."""
    if compressed_tokens == 0:
        return 0.0
    if original_tokens == 0:
        return 1.0
    return original_tokens / compressed_tokens


def spreadsheet_llm_encode(
    excel_path,
    output_path=None,
    k=4,
    vanilla=False,
    compress_homogeneous=True,
    paper_strict=False,
    data_only=True,
    tokenizer_model=DEFAULT_MODEL,
):
    """
    Convert an Excel file to SpreadsheetLLM format or a vanilla markdown-like format.

    Args:
        excel_path (str): Path to the Excel file.
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

    Returns:
        dict: The SpreadsheetLLM encoding of the Excel file.
    """
    if vanilla:
        return vanilla_encode(excel_path, output_path)
    if paper_strict:
        compress_homogeneous = False
    logger.info(f"Processing Excel file: {excel_path}")

    try:
        # `data_only=True` returns cached values from formulas (paper-aligned).
        # Number-format strings are still preserved on the cell metadata.
        workbook = openpyxl.load_workbook(excel_path, data_only=data_only)
        logger.info(
            f"Found {len(workbook.sheetnames)} sheets: {', '.join(workbook.sheetnames)}"
        )
    except FileNotFoundError:
        logger.warning(f"Error: File not found: {excel_path}")
        return None
    except Exception as e:
        logger.warning(f"Error loading Excel file: {e}")
        return None

    sheets_encoding = {}
    compression_metrics = {"sheets": {}}
    overall_orig = overall_anchor = overall_index = overall_format = overall_final = 0

    for sheet_name in workbook.sheetnames:
        logger.info(f"\\nProcessing sheet: {sheet_name}")
        sheet = workbook[sheet_name]

        if sheet.max_row <= 1 and sheet.max_column <= 1:
            logger.info(f"Sheet '{sheet_name}' appears to be empty. Skipping.")
            continue

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
        and num_numeric / num_populated <= 0.5
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
        if row_profiles[r] != row_profiles[r - 1]:
            # Add both sides of the boundary
            row_candidates.add(r)
            row_candidates.add(r + 1)

    col_candidates = set()
    for c in range(1, len(col_profiles)):
        if col_profiles[c] != col_profiles[c - 1]:
            col_candidates.add(c)
            col_candidates.add(c + 1)

    # Step 2: Compose candidate boundaries
    candidates = []
    if row_candidates and col_candidates:
        rows = sorted(list(row_candidates))
        cols = sorted(list(col_candidates))
        for i in range(len(rows)):
            for j in range(i + 1, len(rows)):
                for k in range(len(cols)):
                    for l in range(k + 1, len(cols)):
                        candidates.append((rows[i], cols[k], rows[j], cols[l]))

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

    return sorted(list(final_row_anchors)), sorted(list(final_col_anchors))


def filter_unreasonable_candidates(sheet, candidates):
    """Filter out candidates based on size, sparsity, and header presence."""
    filtered = []
    for r1, c1, r2, c2 in candidates:
        # Size filter
        if (r2 - r1 < 1) or (c2 - c1 < 1):
            continue  # Must have at least 2 rows/cols

        stats = _range_stats(sheet, r1, c1, r2, c2)

        # Internal sparsity filter.
        if stats["density"] < 0.1:
            continue

        # Edge sparsity filter. Real table boundaries generally have visible
        # content near at least one edge; this rejects huge sparse rectangles
        # formed by distant notes or isolated cells.
        if _edge_density(sheet, r1, c1, r2, c2) < 0.08:
            continue

        # Header and proportion filters. Keep text-header tables, date/year
        # header tables, and numeric-heavy tables only when a header row exists.
        has_header = any(is_header_row(sheet, r) for r in range(r1, r2 + 1))
        if not has_header:
            continue
        if stats["text_ratio"] == 0 and stats["year_or_date_ratio"] == 0:
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


def filter_overlapping_candidates(sheet, candidates):
    """Filter overlapping candidates using heuristics from Appendix C."""
    if not candidates:
        return []

    # Score candidates (higher is better)
    scores = []
    for r1, c1, r2, c2 in candidates:
        score = 0
        # Header score
        for r in range(r1, min(r1 + 3, r2 + 1)):  # Check top 3 rows for header
            if is_header_row(sheet, r):
                score += 10
        stats = _range_stats(sheet, r1, c1, r2, c2)
        score += stats["density"] * 5
        score += stats["year_or_date_ratio"] * 3
        # Area score
        score += (r2 - r1 + 1) * (c2 - c1 + 1)
        scores.append(score)

    # Non-maximum suppression based on IoU and scores
    indices = list(range(len(candidates)))
    indices.sort(key=lambda i: scores[i], reverse=True)

    keep = []
    while indices:
        current_idx = indices.pop(0)
        keep.append(current_idx)

        remaining_indices = []
        for idx in indices:
            iou = calculate_iou(candidates[current_idx], candidates[idx])
            # If high overlap, discard the one with the lower score (which is the current `idx` because of sorting)
            if iou < 0.5:
                remaining_indices.append(idx)
        indices = remaining_indices

    return [candidates[i] for i in keep]


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
            for m_range in merged_ranges:
                if cell_ref in m_range:
                    try:
                        merged_value = sheet[m_range.start_cell.coordinate].value
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
                if format_mode == "paper":
                    format_key = _paper_format_key(cell)
                else:
                    format_key = _rich_format_key(cell, merged_range)
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


def vanilla_encode(excel_path, output_path=None):
    """Vanilla markdown-like encoding (paper Section 3.1).

    Produces a ``{sheet_name: pair_string}`` dict where each sheet is the
    paper's row-major ``A1,value|A2,value|...`` baseline. When written to
    disk, all sheets are emitted under ``# {sheet_name}`` headers so
    multi-sheet workbooks aren't silently truncated.
    """
    logger.info(f"Producing vanilla encoding for {excel_path}")
    try:
        workbook = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception as e:
        logger.error(f"Error loading Excel file for vanilla encoding: {e}")
        return None

    vanilla_content = {
        sheet_name: paper_serializers.to_paper_vanilla_prompt(workbook[sheet_name])
        for sheet_name in workbook.sheetnames
    }

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
