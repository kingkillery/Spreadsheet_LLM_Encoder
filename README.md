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

The CLI prints compression ratios for each sheet and overall. These metrics are also stored in the output JSON under `compression_metrics`.

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


## Chain-of-Spreadsheet Pipeline

The module `chain_of_spreadsheet.py` implements a simple two-stage flow:
1. **Table selection** with `identify_table` uses your query to choose the most relevant sheet from the compressed encoding.
2. **Response generation** with `generate_response` reuses your query along with the selected sheet data to produce a textual answer.

The helper script `example_chain_usage.py` shows how to run this pipeline:

```bash
python example_chain_usage.py workbook.xlsx "What were the totals?"
```

## How It Works

The SpreadsheetLLM encoder works through several key steps:

### 1. Structural Anchor Detection

The encoder identifies key structural points in the spreadsheet (rows and columns) that define the layout. Boundary candidates are generated where cell type patterns change markedly, then filtered if they overlap with detected header rows. Anchors are thus based on:
- Cell density changes
- Format transitions
- Content type boundaries

### 2. Cell Neighborhood Extraction

Using the identified anchors, the encoder extracts cells within a configurable distance (parameter `k`). This creates a neighborhood around important structural elements while ignoring less relevant areas.

### 3. Inverted Index Creation

Instead of storing each cell individually, the encoder creates an inverted index mapping content values to cell references. Identical values are merged into contiguous address ranges and empty cells are omitted, resulting in a compact representation.

### 4. Format Region Aggregation

Cell formats are aggregated into rectangular regions. During this step the encoder
infers a semantic type (e.g. year, percentage, date) from each cell's number format
string and groups contiguous cells that share both the type and the raw format
string. This greatly reduces repeated style data in the output and improves
compression.

### 5. Compression

Heterogeneous rows and columns around anchors are retained while uniform regions are skipped. Numeric cells with identical formatting are clustered into aggregated ranges.

### 6. JSON Encoding

The final output is a structured JSON document containing:
- File metadata
- Sheet information
- Structural anchors
- Cell values (inverted index with merged ranges)
- Format regions
- Numeric ranges

## Output Format

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

Utilities for evaluating table detection are included:

- `evaluation.py` implements the Error-of-Boundary (EoB-0) metric and provides
  a loader for the Dong et al. (2019) dataset.
- `run_evaluation.py` executes TableSense-CNN on this dataset and prints the
  resulting F1 score and basic compression statistics.

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
