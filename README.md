DEMO: https://spreadsheetllmencoder.streamlit.app/

# Spreadsheet LLM Encoder

streamlit

This repository contains an implementation of the SpreadsheetLLM encoding method introduced by Microsoft Research in July 2024. The encoder transforms Excel spreadsheets into a specialized JSON format that preserves both content and structural relationships, making them suitable for processing by Large Language Models (LLMs).

## About SpreadsheetLLM

SpreadsheetLLM is a novel approach to encoding spreadsheets that addresses the limitations of traditional methods when working with LLMs. Instead of converting spreadsheets into simple tables or flattened structures, this method:

1. **Preserves structural relationships** between cells using anchor points
2. **Maintains formatting information** for better visual understanding
3. **Creates compact representations** through inverted indexing
4. **Handles merged cells and complex layouts** effectively

This approach significantly improves an LLM's ability to understand, reason about, and manipulate spreadsheet data.

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/Spreadsheet_LLM_Encoder.git
cd Spreadsheet_LLM_Encoder

# Install dependencies
pip install -r requirements.txt

# For development
pip install -r requirements-dev.txt
```

Required dependencies:
- pandas
- openpyxl

Recommended optional dependencies:
- `tiktoken` — for paper-aligned compression metrics (char/4 fallback used if unavailable)
- `openai` — to use `OpenAIBackend` in the Chain-of-Spreadsheet pipeline

## Usage

### Command Line Interface

```bash
python Spreadsheet_LLM_Encoder.py path/to/your/spreadsheet.xlsx --output output.json --k 4

# Or install and use the CLI entry point
pip install -e .
spreadsheet-llm-encode path/to/your/spreadsheet.xlsx --output output.json --k 4
```

Parameters:
- `excel_file`: Path to the Excel file you want to encode (required)
- `--output`, `-o`: Path to save the JSON output (optional, defaults to input filename with '_spreadsheetllm.json' suffix)
- `--k`: Neighborhood distance parameter for structural anchors (optional, default=4, paper's best ablation)
- `--vanilla`: Produce vanilla pair-string encoding instead of compressed (optional)
- `--no-compress-homogeneous`: Retain all structural-anchor cells without pruning homogeneous rows/columns (optional)
- `--tokenizer-model`: Model name for tokenizer-based compression metrics (optional, default="gpt-4")

The CLI prints compression ratios for each sheet and overall to stdout. These metrics are also stored in the output JSON under `compression_metrics` and emitted via the logger at INFO level.

### Python API

```python
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode

# Basic usage (default k=4, paper's best ablation)
encoding = spreadsheet_llm_encode("path/to/your/spreadsheet.xlsx")

# With custom output path and tokenizer
encoding = spreadsheet_llm_encode(
    excel_path="path/to/your/spreadsheet.xlsx", 
    output_path="output.json",
    k=4,
    tokenizer_model="gpt-4"
)

# Vanilla encoding (pair-string format)
encoding = spreadsheet_llm_encode(
    excel_path="path/to/your/spreadsheet.xlsx",
    vanilla=True,
    output_path="vanilla.txt"
)
```


## Chain-of-Spreadsheet (CoS) Pipeline

The module `chain_of_spreadsheet.py` implements the **Chain of Spreadsheet (CoS)** pipeline structure from the paper, broken into the following stages:

1.  **Table Identification (Stage 1)**: Given a query, the system identifies the most relevant sheet (via `find_relevant_sheet`) and calls an LLM to determine the precise boundaries of the table within that sheet (via `identify_table`).
2.  **Response Generation (Stage 2)**: With `workbook_path`, `sheet_name`, and `table_range` provided, the paper-faithful **uncompressed pair-string** is read directly from the original workbook and passed to the LLM along with the query. Without those parameters, it falls back to the compressed encoding with a deprecation warning.
3.  **Table Splitting for Large Tables (Algorithm 2)**: For tables exceeding `token_limit`, `table_split_qa` performs real **body-row chunking**: it greedily partitions rows so each `header + chunk` pair fits within the token limit (measured by `tokenizer.count_tokens`), then calls the LLM per chunk and aggregates answers.

### Configuring an LLM Backend

Configure via `configure_backend()` (preferred) or monkey-patch `_call_llm`:

```python
import chain_of_spreadsheet as cos
from llm_backend import OpenAIBackend, EchoBackend

# Option 1: use OpenAIBackend
cos.configure_backend(OpenAIBackend(model="gpt-4o-mini"))

# Option 2: use EchoBackend for testing
cos.configure_backend(EchoBackend(response="['range': 'A1:B5']"))

# Option 3: monkey-patch (legacy, still supported)
cos._call_llm = lambda p: my_llm_client.complete(p)
```

The `example_chain_usage.py` script demonstrates end-to-end usage with a real LLM backend.

## How It Works: The `SheetCompressor`

The `SheetCompressor` is at the heart of SpreadsheetLLM, using three sophisticated modules to create a compact and semantically rich representation of a spreadsheet.

### 1. Structural Anchor Detection

The encoder uses advanced heuristics (as described in Appendix C of the paper) to find structural anchors. This multi-step process involves:
- **Enumerating Boundaries**: Identifying changes in cell values, styles (borders, fills), and merged regions.
- **Composing Candidates**: Forming all possible rectangular table candidates from these boundaries.
- **Filtering**: Removing unreasonable candidates based on size and sparsity, and resolving overlaps using an IoU-based non-maximum suppression approach.

This produces a highly accurate "skeleton" of the spreadsheet's structure.

### 2. Inverted Index Creation

A lossless inverted index is created, mapping cell content to cell addresses. This is highly efficient for spreadsheets with repetitive data or many empty cells, as identical values are merged into address ranges (`A1:C1`) and empty cells are omitted.

### 3. Data-Format-Aware Aggregation

This module intelligently groups cells to reduce redundancy and enhance semantic meaning.
- **Semantic Type Detection**: The encoder now recognizes a wider range of semantic types, including **Integer, Float, and Email**, by inspecting both the number format string and the cell value itself.
- **DFS-based Aggregation**: Instead of a simple greedy search, the encoder uses a Depth-First Search (DFS) algorithm (as described in Appendix M.1 of the paper) to find all contiguous regions of cells that share the same semantic type and number format. This correctly aggregates complex, non-rectangular shapes.

The final output is a structured JSON document containing the structural anchors, the inverted index, aggregated format regions, and numeric ranges.

## Output Formats

The encoder produces a JSON with this structure:

```json
{
  "file_name": "example.xlsx",
  "sheets": {
    "Sheet1": {
      "structural_anchors": {
        "rows": [1, 5, 10],
        "columns": ["A", "C", "F"]
      },
      "cells": {
        "Header": ["A1:C1"],
        "42": ["B5"],
        "Total": ["A10"]
      },
      "formats": {
        "{format_definition}": ["A1:C1", "A10:F10"]
      },
      "numeric_ranges": {
        "{format_definition}": ["B2:B8"]
      },
      "coord_map": {
        "rows": {1: 1, 5: 2, 10: 3},
        "cols": {1: 1, 3: 2, 6: 3},
        "rows_inv": {1: 1, 2: 5, 3: 10},
        "cols_inv": {1: 1, 2: 3, 3: 6}
      }
    }
  }
}
```

The `coord_map` field enables round-trip conversion between original workbook coordinates and the compact remapped space used in the paper-faithful prompts.

### Compression Metrics

The encoder reports token counts before and after each stage using `tokenizer.count_tokens`, which uses **tiktoken** when available and falls back to a deterministic **char/4** approximation. These metrics are stored under `compression_metrics` in the JSON output:

```json
"compression_metrics": {
  "overall": {
    "original_tokens": 1200,
    "after_anchor_tokens": 400,
    "after_inverted_index_tokens": 320,
    "after_format_tokens": 350,
    "final_tokens": 340,
    "anchor_ratio": 3.0,
    "inverted_index_ratio": 3.75,
    "format_ratio": 3.43,
    "overall_ratio": 3.53
  },
  "sheets": {
    "Sheet1": {
      "original_tokens": 1200,
      "after_anchor_tokens": 400,
      "anchor_ratio": 3.0,
      "inverted_index_ratio": 3.75,
      "format_ratio": 3.43,
      "overall_ratio": 3.53
    }
  }
}
```

The baseline (`original_tokens`) uses the paper's vanilla pair-string format (including empty cells in the bounding box), not non-empty JSON cells. Install `tiktoken` for metrics aligned with the paper's reported numbers.

## Evaluation

The repository now includes a comprehensive framework for evaluating SpreadsheetLLM, as described in the paper.

### Table Detection Benchmark

-   **Dataset**: The framework uses spreadsheet files (`.xlsx`) and corresponding JSON annotations, which is the correct format for evaluating SpreadsheetLLM. A new data loader `load_spreadsheet_dataset` is included in `evaluation.py`.
-   **Evaluation Script**: The `run_llm_evaluation.py` script runs the full table detection benchmark. It encodes spreadsheets, uses a (placeholder) LLM to predict table boundaries, and evaluates the predictions against ground truth using the EoB-0 metric.
-   **Fine-tuning Preparation**: The `prepare_finetuning_data.py` script can be used to convert a dataset into the JSONL format required for fine-tuning LLMs on the table detection task.

### Spreadsheet QA Benchmark

-   **Dataset**: A new data loader `load_qa_dataset` is included for the Spreadsheet QA benchmark described in Appendix H of the paper.
-   **Evaluation Script**: The `run_qa_evaluation.py` script evaluates the performance of the full CoS pipeline on the QA task. It calculates the accuracy of the generated answers and includes placeholders for running baseline models like `TaPEx` and `Binder`.

## Paper-faithful Prompts

The `paper_serializers` module exports three canonical prompt formats aligned with the paper (arXiv:2407.09025):

**Vanilla pair-string (Section 3.1, Stage 2):**
```python
from paper_serializers import to_paper_vanilla_prompt
import openpyxl

wb = openpyxl.load_workbook("data.xlsx", data_only=True)
prompt = to_paper_vanilla_prompt(wb["Sheet1"])
# Output: "A1,Year|A2,2020|A3,2021|B1,Profit|B2,100|B3,150"
```

**Compressed pair-string with format substitution (Stage 1):**
```python
from paper_serializers import to_paper_compressed_prompt

sheet_data = encoding["sheets"]["Sheet1"]
prompt = to_paper_compressed_prompt(sheet_data, coord_map=sheet_data.get("coord_map"))
# Output: "(Year|A1)(IntNum|A2:A3)(Profit|B1)(IntNum|B2:B3)"
```

**Stage 2 uncompressed sub-range (paper-faithful, requires original workbook):**
```python
from paper_serializers import to_stage2_uncompressed_prompt

prompt = to_stage2_uncompressed_prompt(
    workbook_path="data.xlsx",
    sheet_name="Sheet1",
    table_range="A1:B3"
)
# Output: "A1,Year|A2,2020|A3,2021|B1,Profit|B2,100|B3,150"
```

## Vanilla Encoding

For baseline comparisons and debugging, the encoder can produce a simple "vanilla" pair-string encoding (as described in Section 3.1 of the paper). Use the `--vanilla` flag in the CLI:

```bash
spreadsheet-llm-encode path/to/your/spreadsheet.xlsx --vanilla --output output.txt
```

The vanilla encoding includes **all sheets** in the workbook, each preceded by a sheet-name header (`# {sheet_name}`), in row-major order.

## Gap Analysis Remediation

This section documents changes made to align the implementation with the paper (arXiv:2407.09025):

- **Default k**: Changed from 2 to 4 (paper's best ablation; see Appendix E).
- **Real Stage 2**: `generate_response` now reads the original workbook and emits the uncompressed pair-string for identified sub-ranges (Section 4.2), not the compressed encoding. Call with `workbook_path`, `sheet_name`, and `table_range` for paper-faithful behavior.
- **Real Algorithm 2**: `table_split_qa` now performs actual body-row chunking (Appendix M.2), partitioning rows greedily so each `header + chunk` fits the token limit.
- **Tokenizer-based metrics**: Compression ratios now use `tiktoken` (when available) instead of character counts, matching the paper's reported numbers. Fallback to char/4 when tiktoken is unavailable.
- **Format-region substitution**: In the compressed prompt, compressible types (integer, float, date, datetime, time, year, email, scientific notation, percentage, currency) are substituted with semantic labels (`IntNum`, `FloatNum`, `DateData`, etc.) in the rendered prompt.
- **Coordinate remapping**: Each sheet encoding now includes a `coord_map` field mapping original↔compact coordinates, enabling round-trip conversion between predicted compact ranges and original workbook addresses.
- **Paper serializers**: New `paper_serializers` module exports `to_paper_vanilla_prompt`, `to_paper_compressed_prompt`, and `to_stage2_uncompressed_prompt` for faithful prompt generation.
- **LLM backend abstraction**: `chain_of_spreadsheet.configure_backend(backend)` accepts any `LLMBackend` (any `Callable[[str], str]`). Reference implementations provided: `OpenAIBackend`, `EchoBackend`.

**Not yet implemented** (per the original paper):
- Paper datasets are not bundled.
- TaPEx and Binder baseline comparisons remain placeholder stubs.
- Fine-tuning training loop is not included (data preparation only).

## Research Background

This implementation is based on the paper "[SpreadsheetLLM: Enabling LLMs to Understand Spreadsheets](https://www.microsoft.com/en-us/research/)" published by Microsoft Research in July 2024. The paper introduces a novel approach to encode spreadsheets for LLM comprehension that preserves structural integrity and visual semantics.

## Development

Before submitting, ensure that the code passes `flake8` checks:

```bash
pip install -r requirements-dev.txt
flake8 .
```

## License

This project is licensed under the [MIT License](LICENSE).

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
