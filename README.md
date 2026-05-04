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

## Usage

### Command Line Interface

```bash
python Spreadsheet_LLM_Encoder.py path/to/your/spreadsheet.xlsx --output output.json --k 2

# Or install and use the CLI entry point
pip install -e .
spreadsheet-llm-encode path/to/your/spreadsheet.xlsx --output output.json --k 2
```

Parameters:
- `excel_file`: Path to the Excel file you want to encode (required)
- `--output`, `-o`: Path to save the JSON output (optional, defaults to input filename with '_spreadsheetllm.json' suffix)
- `--k`: Neighborhood distance parameter for structural anchors (optional, default=2)

The CLI prints compression ratios for each sheet and overall to stdout. These metrics are also stored in the output JSON under `compression_metrics` and emitted via the logger at INFO level.

### Python API

```python
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode

# Basic usage
encoding = spreadsheet_llm_encode("path/to/your/spreadsheet.xlsx")

# With custom output path and neighborhood parameter
encoding = spreadsheet_llm_encode(
    excel_path="path/to/your/spreadsheet.xlsx", 
    output_path="output.json",
    k=3
)
```


## Chain-of-Spreadsheet (CoS) Pipeline

The module `chain_of_spreadsheet.py` implements the **Chain of Spreadsheet (CoS)** pipeline structure from the paper, broken into the following stages:

1.  **Table Identification**: Given a query, the system first identifies the most relevant sheet and then calls an LLM to determine the precise boundaries of the table within that sheet that contains the answer.
2.  **Response Generation**: The compressed sheet encoding is passed to the LLM along with the original query to generate a final response.  Note that, unlike the paper (Section 4.2), this implementation feeds the **already-compressed** representation rather than an uncompressed re-encoding of the identified sub-range; a fully-faithful implementation would require access to the original file.
3.  **Table Splitting for Large Tables**: For tables that exceed the LLM's context window the pipeline falls back to the **Table Split QA Algorithm** (Appendix M.2).  The current implementation is a **simplified placeholder**: it does not actually split the table into row-based chunks — it passes the same data for every sub-query.  A production implementation must partition the body rows so that each chunk fits within `token_limit`.

> **Important – LLM integration required**: `chain_of_spreadsheet.py` does **not** ship with a built-in LLM client.  The internal `_call_llm()` function raises `NotImplementedError` by default.  Before using the CoS pipeline you must supply your own LLM backend:
>
> ```python
> import chain_of_spreadsheet as cos
>
> def my_llm(prompt: str) -> str:
>     # e.g. call OpenAI, Anthropic, or a local model
>     ...
>
> cos._call_llm = my_llm
> ```

The `example_chain_usage.py` script demonstrates how to use this pipeline once an LLM backend is configured.

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
      }
    }
  }
}
```

### Compression Metrics

The encoder reports token counts before and after each stage. These values are stored under `compression_metrics` in the JSON output. Example:

```json
"compression_metrics": {
  "overall": {
    "overall_ratio": 3.5
  },
  "sheets": {
    "Sheet1": {
      "anchor_ratio": 2.1,
      "inverted_index_ratio": 3.0,
      "format_ratio": 3.4,
      "overall_ratio": 3.5
    }
  }
}
```

## Evaluation

The repository now includes a comprehensive framework for evaluating SpreadsheetLLM, as described in the paper.

### Table Detection Benchmark

-   **Dataset**: The framework uses spreadsheet files (`.xlsx`) and corresponding JSON annotations, which is the correct format for evaluating SpreadsheetLLM. A new data loader `load_spreadsheet_dataset` is included in `evaluation.py`.
-   **Evaluation Script**: The `run_llm_evaluation.py` script runs the full table detection benchmark. It encodes spreadsheets, uses a (placeholder) LLM to predict table boundaries, and evaluates the predictions against ground truth using the EoB-0 metric.
-   **Fine-tuning Preparation**: The `prepare_finetuning_data.py` script can be used to convert a dataset into the JSONL format required for fine-tuning LLMs on the table detection task.

### Spreadsheet QA Benchmark

-   **Dataset**: A new data loader `load_qa_dataset` is included for the Spreadsheet QA benchmark described in Appendix H of the paper.
-   **Evaluation Script**: The `run_qa_evaluation.py` script evaluates the performance of the full CoS pipeline on the QA task. It calculates the accuracy of the generated answers and includes placeholders for running baseline models like `TaPEx` and `Binder`.

## Vanilla Encoding

For baseline comparisons and debugging, the encoder can produce a simple "vanilla" markdown-like encoding (as described in Section 3.1 of the paper). Use the `--vanilla` flag in the CLI:

```bash
spreadsheet-llm-encode path/to/your/spreadsheet.xlsx --vanilla --output output.txt
```

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
