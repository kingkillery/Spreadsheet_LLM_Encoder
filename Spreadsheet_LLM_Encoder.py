import os
import pandas as pd  # Used for some data manipulation
import openpyxl
import json
import logging
from temp_helpers import infer_cell_data_type, categorize_number_format
from collections import defaultdict
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
import sys

logger = logging.getLogger(__name__)

def spreadsheet_llm_encode(excel_path, output_path=None, k=2):
    """
    Convert an Excel file to SpreadsheetLLM format, handling multiple sheets and detailed formats.
    Identical cell values are merged into address ranges for a compact inverted index.

    Args:
        excel_path (str): Path to the Excel file.
        output_path (str, optional): Path to save the JSON output. Defaults to None.
        k (int, optional): Neighborhood distance for structural anchors. Defaults to 2.

    Returns:
        dict: The SpreadsheetLLM encoding of the Excel file.
    """
    logger.info(f"Processing Excel file: {excel_path}")

    try:
        workbook = openpyxl.load_workbook(excel_path, data_only=False) # Changed data_only to False
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

    for sheet_name in workbook.sheetnames:
        logger.info(f"\nProcessing sheet: {sheet_name}")
        sheet = workbook[sheet_name]

        if sheet.max_row <= 1 and sheet.max_column <= 1:
            logger.info(f"Sheet '{sheet_name}' appears to be empty. Skipping.")
            continue

        logger.info(
            f"Sheet dimensions: {sheet.max_row} rows × {sheet.max_column} columns"
        )
        # print memory usage
        logger.info(f"Estimated memory usage: {sys.getsizeof(sheet)} bytes")

        row_anchors, col_anchors = find_structural_anchors(sheet)
        logger.info(
            f"Found {len(row_anchors)} row anchors and {len(col_anchors)} column anchors"
        )

        kept_rows, kept_cols = extract_cells_near_anchors(sheet, row_anchors, col_anchors, k)
        logger.info(
            f"Keeping {len(kept_rows)} rows and {len(kept_cols)} columns"
        )

        inverted_index, format_map = create_inverted_index(sheet, kept_rows, kept_cols)
        logger.info(
            f"Created inverted index with {len(inverted_index)} unique values"
        )

        merged_index = create_inverted_index_translation(inverted_index)
        logger.info(
            f"Merged values into {len(merged_index)} range groups"
        )

        aggregated_formats = aggregate_formats(sheet, format_map)
        logger.info(
            f"Aggregated {len(aggregated_formats)} format regions"
        )

        sheets_encoding[sheet_name] = {
            "structural_anchors": {
                "rows": row_anchors,
                "columns": [get_column_letter(c) for c in col_anchors]
            },
            "cells": merged_index,
            "formats": aggregated_formats
        }

    full_encoding = {
        "file_name": os.path.basename(excel_path),
        "sheets": sheets_encoding
    }

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(full_encoding, f, indent=2, ensure_ascii=False)
        logger.info(f"Saved SpreadsheetLLM encoding to {output_path}")

    return full_encoding

def find_structural_anchors(sheet):
    """Find structural anchors based on cell count and format changes."""
    row_counts = [0] * (sheet.max_row + 1)
    col_counts = [0] * (sheet.max_column + 1)
    row_formats = [set() for _ in range(sheet.max_row + 1)]
    col_formats = [set() for _ in range(sheet.max_column + 1)]

    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.value is not None:
                row_counts[row] += 1
                col_counts[col] += 1

                # --- Format-based Anchors ---
                format_key = get_cell_format_key(cell)
                row_formats[row].add(format_key)
                col_formats[col].add(format_key)

    row_anchors = []
    for r in range(1, len(row_counts)):
        if r == 1 and row_counts[r] > 0:
            row_anchors.append(r)
        elif r > 1 and r < len(row_counts) - 1:
            # Check for significant changes in *both* count and format variety
            if (abs(row_counts[r] - row_counts[r - 1]) > 2 or
                abs(row_counts[r] - row_counts[r + 1]) > 2 or
                len(row_formats[r]) != len(row_formats[r-1]) or
                len(row_formats[r]) != len(row_formats[r+1])):
                row_anchors.append(r)

    col_anchors = []
    for c in range(1, len(col_counts)):
        if c == 1 and col_counts[c] > 0:
            col_anchors.append(c)
        elif c > 1 and c < len(col_counts) - 1:
            if (abs(col_counts[c] - col_counts[c - 1]) > 2 or
                abs(col_counts[c] - col_counts[c + 1]) > 2 or
                len(col_formats[c]) != len(col_formats[c-1]) or
                len(col_formats[c]) != len(col_formats[c+1])):
                col_anchors.append(c)

    return row_anchors, col_anchors

def get_cell_format_key(cell):
    """Helper function to create a consistent format key for a cell."""
    format_info = {}
    try:
        # 1. Font Styles
        font = cell.font
        format_info["font"] = {
            "bold": font.bold,
            "italic": font.italic,
            "underline": font.underline,
            "name": font.name,
            "size": font.sz,
            "color": str(font.color.rgb) if font.color and font.color.rgb else None,  # Get RGB color
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
                "color": str(getattr(border, side).color.rgb) if getattr(border, side).color and getattr(border, side).color.rgb else None,
            } for side in ["left", "right", "top", "bottom"]
        }

        # 4. Fill (Background Color)
        fill = cell.fill
        if hasattr(fill, 'patternType') and fill.patternType == "solid":
            format_info["fill"] = {"color": str(fill.start_color.index) if fill.start_color and fill.start_color.index else None}
        else:
            format_info["fill"] = {"color": None}

        # 5. Number Format (Original, Inferred Type, Category)
        original_number_format = cell.number_format
        inferred_type = infer_cell_data_type(cell) # Call new helper
        category = categorize_number_format(original_number_format, inferred_type) # Call new helper

        format_info["original_number_format"] = original_number_format
        format_info["inferred_data_type"] = inferred_type
        format_info["number_format_category"] = category
        # format_info["number_format"] = cell.number_format # Keep original for now, or decide if it's redundant
    except Exception as e:
        # If there's an error extracting format, use a simplified format key
        format_info = {"error": str(e)}

    return json.dumps(format_info, sort_keys=True)

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

def create_inverted_index(sheet, kept_rows, kept_cols):
    """Create an inverted index, handling merged cells."""
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
                        "color": str(getattr(border, side).color.rgb) if getattr(border, side).color and getattr(border, side).color.rgb else None,
                    } for side in ["left", "right", "top", "bottom"]
                }

                # 4. Fill (Background Color)
                fill = cell.fill
                if hasattr(fill, 'patternType') and fill.patternType == "solid":
                    format_info["fill"] = {"color": str(fill.start_color.index) if fill.start_color and fill.start_color.index else None}
                else:
                    format_info["fill"] = {"color": None}

                # 5. Number Format (Original, Inferred Type, Category)
                original_number_format = cell.number_format
                inferred_type = infer_cell_data_type(cell) # Call new helper
                category = categorize_number_format(original_number_format, inferred_type) # Call new helper

                format_info["original_number_format"] = original_number_format
                format_info["inferred_data_type"] = inferred_type
                format_info["number_format_category"] = category
                # format_info["number_format"] = cell.number_format # Redundant

                # Store the format (handle merged ranges specially in format key)
                if merged_range is not None:
                    format_info["merged"] = True
                    format_info["merged_range"] = str(merged_range)
                else:
                    format_info["merged"] = False

                format_key = json.dumps(format_info, sort_keys=True)
                format_map[format_key].append(cell_ref)
            except Exception as e:
                # Handle error for problematic cell formats
                logger.warning(f"Error processing format for cell {cell_ref}: {e}")

    return dict(inverted_index), dict(format_map)

def aggregate_formats(sheet, format_map):
    """Aggregate cells with the same format into rectangular regions."""
    aggregated_formats = defaultdict(list)
    processed_cells = set()

    for fmt, cells in format_map.items():
        try:
            format_data = json.loads(fmt)
            # Handle merged cells first
            if format_data.get('merged') is True and 'merged_range' in format_data:
                aggregated_formats[fmt].append(format_data['merged_range'])
                # Mark all cells in the merged range as processed
                try:
                    merged_range = format_data['merged_range']
                    start_ref, end_ref = merged_range.split(':')
                    start_col_letter, start_row = split_cell_ref(start_ref)
                    end_col_letter, end_row = split_cell_ref(end_ref)
                    
                    start_col = openpyxl.utils.cell.column_index_from_string(start_col_letter)
                    end_col = openpyxl.utils.cell.column_index_from_string(end_col_letter)
                    
                    for r in range(start_row, end_row + 1):
                        for c in range(start_col, end_col + 1):
                            cell_ref = f"{get_column_letter(c)}{r}"
                            processed_cells.add(cell_ref)
                except Exception as e:
                    logger.warning(
                        f"Error processing merged range {format_data['merged_range']}: {e}"
                    )
                
                continue  # Skip to next format

            if len(cells) < 3:  # Skip very small format groups
                continue

            cells_set = set(cells)  # for faster lookup

            for start_cell in cells:
                if start_cell in processed_cells:
                    continue

                # Parse starting cell
                try:
                    start_col_letter, start_row = split_cell_ref(start_cell)
                    start_col = openpyxl.utils.cell.column_index_from_string(start_col_letter)
                except Exception:
                    continue  # Skip if cell reference can't be parsed

                # Initialize best rectangle
                best_width = 1
                best_height = 1
                best_area = 1
                best_end_cell = start_cell

                # Try expanding in all directions
                max_width = min(20, sheet.max_column - start_col + 1)  # Limit search to reasonable size
                max_height = min(20, sheet.max_row - start_row + 1)    # Limit search to reasonable size
                
                for width in range(1, max_width + 1):
                    for height in range(1, max_height + 1):
                        valid_rectangle = True
                        for r in range(start_row, start_row + height):
                            for c in range(start_col, start_col + width):
                                cell_ref = f"{get_column_letter(c)}{r}"
                                if cell_ref not in cells_set or cell_ref in processed_cells:
                                    valid_rectangle = False
                                    break
                            if not valid_rectangle:
                                break

                        if valid_rectangle:
                            # Update best rectangle if this one is larger
                            area = width * height
                            if area > best_area:
                                best_width = width
                                best_height = height
                                best_area = area
                                best_end_cell = f"{get_column_letter(start_col + width - 1)}{start_row + height - 1}"

                # Add the best rectangle found
                if best_width > 1 or best_height > 1:
                    region = f"{start_cell}:{best_end_cell}"
                    aggregated_formats[fmt].append(region)

                    # Mark cells as processed
                    for r in range(start_row, start_row + best_height):
                        for c in range(start_col, start_col + best_width):
                            processed_cells.add(f"{get_column_letter(c)}{r}")
        except Exception as e:
            logger.warning(f"Error aggregating format {fmt}: {e}")

    return dict(aggregated_formats)

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

def get_column_index(col_letter):
    """Convert column letter to index (A => 1, AA => 27)."""
    return openpyxl.utils.cell.column_index_from_string(col_letter)

def split_cell_ref(cell_ref):
    """Split cell reference (e.g., 'A1') into column letter and row number."""
    col_str = ''.join(filter(str.isalpha, cell_ref))
    row_str = ''.join(filter(str.isdigit, cell_ref))

    # convert row to integer
    return col_str, int(row_str)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    import argparse

    parser = argparse.ArgumentParser(description='Convert Excel files to SpreadsheetLLM format')
    parser.add_argument('excel_file', help='Path to the Excel file')
    parser.add_argument('--output', '-o', help='Output JSON file path (default: same as input with .json extension)')
    parser.add_argument('--k', type=int, default=2, help='Neighborhood distance parameter (default: 2)')

    args = parser.parse_args()

    if not args.output:
        args.output = os.path.splitext(args.excel_file)[0] + '_spreadsheetllm.json'

    spreadsheet_llm_encode(args.excel_file, args.output, args.k)
