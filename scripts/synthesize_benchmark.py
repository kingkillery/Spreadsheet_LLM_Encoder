"""
synthesize_benchmark.py
-----------------------
Generate synthetic annotated XLSX benchmarks compatible with
evaluation.py:load_spreadsheet_dataset.

Usage:
    python scripts/synthesize_benchmark.py <output_dir> [--n 50] [--seed 42]
         [--small-only] [--include-multi-table] [--include-merged]
         [--include-headers]
"""

import argparse
import json
import random
import string
import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, numbers as xl_numbers

# Bold font applied to header cells so the encoder's header heuristic
# (Spreadsheet_LLM_Encoder.is_header_row, requires ~60% of populated header
# cells to be bold/centered/all-caps) can identify them. Real-world
# spreadsheets style their headers; the synthesizer must too, or
# find_structural_anchors filters out every candidate and the encoder
# returns an empty skeleton.
_HEADER_FONT = Font(bold=True)

# ---------------------------------------------------------------------------
# Constants / wordlists
# ---------------------------------------------------------------------------

WORDLIST = [
    "alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel",
    "india", "juliet", "kilo", "lima", "mike", "november", "oscar", "papa",
    "quebec", "romeo", "sierra", "tango", "uniform", "victor", "whiskey",
    "xray", "yankee", "zulu", "red", "blue", "green", "yellow", "orange",
    "north", "south", "east", "west", "prime", "alpha2", "beta", "gamma",
]

FIRST_NAMES = [
    "Alice", "Bob", "Carol", "Dave", "Eve", "Frank", "Grace", "Hank",
    "Iris", "Jack", "Karen", "Leo", "Mia", "Ned", "Olivia", "Pete",
]

DOMAINS = ["example.com", "test.org", "sample.net", "data.io", "bench.co"]

COLUMN_NAMES = [
    "ID", "Name", "Revenue", "Cost", "Profit", "Score", "Date", "Email",
    "Region", "Category", "Amount", "Price", "Quantity", "Rate", "Value",
    "Count", "Status", "Code", "Type", "Rank",
]

# Size classes: (min_rows, max_rows, min_cols, max_cols)
SIZE_CLASSES = {
    "small":  (5, 8, 3, 4),
    "medium": (15, 25, 6, 8),
    "large":  (40, 60, 10, 12),
}

LAYOUTS = ["simple_grid", "row_headers", "merged_title", "blank_row_separator"]

CONTENT_TYPES = ["integer", "float", "date", "currency", "email", "text"]

# ---------------------------------------------------------------------------
# Cell-value generators
# ---------------------------------------------------------------------------


def _gen_integer(rng):
    return rng.randint(1, 99999)


def _gen_float(rng):
    return round(rng.uniform(0.5, 9999.99), 2)


def _gen_date(rng):
    base = datetime.date(2020, 1, 1)
    offset = rng.randint(0, 1460)
    return base + datetime.timedelta(days=offset)


def _gen_currency(rng):
    return round(rng.uniform(10.0, 50000.0), 2)


def _gen_email(rng):
    name = rng.choice(FIRST_NAMES).lower()
    suffix = rng.randint(1, 99)
    domain = rng.choice(DOMAINS)
    return f"{name}{suffix}@{domain}"


def _gen_text(rng):
    return rng.choice(WORDLIST)


GENERATORS = {
    "integer": _gen_integer,
    "float": _gen_float,
    "date": _gen_date,
    "currency": _gen_currency,
    "email": _gen_email,
    "text": _gen_text,
}

NUMBER_FORMATS = {
    "integer": "0",
    "float": "0.00",
    "date": "yyyy-mm-dd",
    "currency": '$#,##0.00',
    "email": "@",
    "text": "@",
}

# ---------------------------------------------------------------------------
# Table builder
# ---------------------------------------------------------------------------


def _pick_col_types(rng, n_cols):
    """Return a list of n_cols content-type strings with at least 2 distinct types."""
    # Ensure at least 2 different types
    base_types = rng.sample(CONTENT_TYPES, min(2, len(CONTENT_TYPES)))
    types = []
    for i in range(n_cols):
        types.append(base_types[i % len(base_types)])
    # shuffle so types aren't always in the same column order
    rng.shuffle(types)
    return types


def _col_header(name_pool, idx):
    if idx < len(name_pool):
        return name_pool[idx]
    return f"Col{idx + 1}"


def _write_simple_grid(ws, rng, start_row, start_col, n_rows, n_cols):
    """1 header row + n_rows body rows. Returns (actual_start_row, actual_end_row)."""
    col_names = rng.sample(COLUMN_NAMES, min(n_cols, len(COLUMN_NAMES)))
    col_types = _pick_col_types(rng, n_cols)

    # Header row (bold so the encoder's header heuristic recognizes it)
    for ci in range(n_cols):
        cell = ws.cell(row=start_row, column=start_col + ci, value=col_names[ci])
        cell.font = _HEADER_FONT

    # Body rows
    for ri in range(1, n_rows + 1):
        for ci, ctype in enumerate(col_types):
            cell = ws.cell(
                row=start_row + ri,
                column=start_col + ci,
                value=GENERATORS[ctype](rng),
            )
            cell.number_format = NUMBER_FORMATS[ctype]

    return start_row, start_row + n_rows, start_col, start_col + n_cols - 1


def _write_row_headers(ws, rng, start_row, start_col, n_rows, n_cols):
    """First row + first column are headers."""
    col_names = rng.sample(COLUMN_NAMES, min(n_cols - 1, len(COLUMN_NAMES)))
    col_types = _pick_col_types(rng, n_cols - 1)

    # Top-left corner blank
    ws.cell(row=start_row, column=start_col, value="")

    # Column headers (first row, bold)
    for ci, name in enumerate(col_names):
        cell = ws.cell(row=start_row, column=start_col + 1 + ci, value=name)
        cell.font = _HEADER_FONT

    # Row headers (first column, bold) + body
    for ri in range(1, n_rows + 1):
        row_label = f"R{ri}"
        cell = ws.cell(row=start_row + ri, column=start_col, value=row_label)
        cell.font = _HEADER_FONT
        for ci, ctype in enumerate(col_types):
            cell = ws.cell(
                row=start_row + ri,
                column=start_col + 1 + ci,
                value=GENERATORS[ctype](rng),
            )
            cell.number_format = NUMBER_FORMATS[ctype]

    return start_row, start_row + n_rows, start_col, start_col + n_cols - 1


def _write_merged_title(ws, rng, start_row, start_col, n_rows, n_cols):
    """Rows 1-2 merged for a title, then header row, then body."""
    title = f"{rng.choice(WORDLIST).capitalize()} Report"
    end_col = start_col + n_cols - 1

    # Merge title across columns (rows 1 and 2)
    ws.merge_cells(
        start_row=start_row,
        start_column=start_col,
        end_row=start_row + 1,
        end_column=end_col,
    )
    ws.cell(row=start_row, column=start_col, value=title)

    # Header row at row 3 (bold)
    header_row = start_row + 2
    col_names = rng.sample(COLUMN_NAMES, min(n_cols, len(COLUMN_NAMES)))
    col_types = _pick_col_types(rng, n_cols)
    for ci, name in enumerate(col_names):
        cell = ws.cell(row=header_row, column=start_col + ci, value=name)
        cell.font = _HEADER_FONT

    # Body rows
    actual_end_row = header_row + n_rows
    for ri in range(1, n_rows + 1):
        for ci, ctype in enumerate(col_types):
            cell = ws.cell(
                row=header_row + ri,
                column=start_col + ci,
                value=GENERATORS[ctype](rng),
            )
            cell.number_format = NUMBER_FORMATS[ctype]

    return start_row, actual_end_row, start_col, end_col


def _write_blank_row_separator(ws, rng, start_row, start_col, n_rows, n_cols):
    """Table preceded by 2 blank rows so anchors must locate it."""
    # 2 blank rows before the actual table
    actual_start = start_row + 2
    r1, r2, c1, c2 = _write_simple_grid(ws, rng, actual_start, start_col, n_rows, n_cols)
    return r1, r2, c1, c2


LAYOUT_WRITERS = {
    "simple_grid": _write_simple_grid,
    "row_headers": _write_row_headers,
    "merged_title": _write_merged_title,
    "blank_row_separator": _write_blank_row_separator,
}

# ---------------------------------------------------------------------------
# Main generation logic
# ---------------------------------------------------------------------------


def _excel_range(r1, c1, r2, c2):
    return f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"


def generate_one(idx, rng, args):
    """Generate one XLSX + JSON annotation pair. Returns (xlsx_path, json_path)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    n_tables = 1
    if args.include_multi_table:
        n_tables = rng.randint(1, 3)

    tables_meta = []
    current_row = 1

    size_cycle = list(SIZE_CLASSES.keys())
    layout_cycle = LAYOUTS

    for t in range(n_tables):
        # Pick layout
        layout = layout_cycle[(idx * n_tables + t) % len(layout_cycle)]
        if args.small_only:
            size_class = "small"
        else:
            size_class = size_cycle[(idx * n_tables + t) % len(size_cycle)]

        min_r, max_r, min_c, max_c = SIZE_CLASSES[size_class]
        n_rows = rng.randint(min_r, max_r)
        n_cols = rng.randint(min_c, max_c)
        start_col = 1

        writer = LAYOUT_WRITERS[layout]
        r1, r2, c1, c2 = writer(ws, rng, current_row, start_col, n_rows, n_cols)

        tables_meta.append({
            "range": _excel_range(r1, c1, r2, c2),
            "sheet": "Sheet1",
            "layout": layout,
            "size_class": size_class,
        })

        # Leave gap before next table
        current_row = r2 + rng.randint(2, 4)

    return wb, tables_meta


def main():
    parser = argparse.ArgumentParser(description="Generate synthetic XLSX benchmark files.")
    parser.add_argument("output_dir", help="Directory to write generated files")
    parser.add_argument("--n", type=int, default=50, help="Number of files to generate")
    parser.add_argument("--seed", type=int, default=42, help="Random seed")
    parser.add_argument("--small-only", action="store_true", help="Only generate small tables")
    parser.add_argument("--include-multi-table", action="store_true",
                        help="Allow 1-3 tables per file (default: 1)")
    parser.add_argument("--include-merged", action="store_true",
                        help="Include merged-title layout (always included by default)")
    parser.add_argument("--include-headers", action="store_true",
                        help="Include row-headers layout (always included by default)")
    args = parser.parse_args()

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    rng = random.Random(args.seed)

    for i in range(args.n):
        wb, tables_meta = generate_one(i, rng, args)

        stem = f"synth_{i:04d}"
        xlsx_path = out_dir / f"{stem}.xlsx"
        json_path = out_dir / f"{stem}.json"

        wb.save(str(xlsx_path))

        annotation = {"tables": tables_meta}
        with open(json_path, "w") as f:
            json.dump(annotation, f, indent=2)

    print(f"Wrote {args.n} files to {out_dir}")


if __name__ == "__main__":
    main()
