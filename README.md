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
```

Required dependencies:
- pandas
- openpyxl

## Usage

### Command Line Interface

```bash
python Spreadsheet_LLM_Encoder.py path/to/your/spreadsheet.xlsx --output output.json --k 2
```

Parameters:
- `excel_file`: Path to the Excel file you want to encode (required)
- `--output`, `-o`: Path to save the JSON output (optional, defaults to input filename with '_spreadsheetllm.json' suffix)
- `--k`: Neighborhood distance parameter for structural anchors (optional, default=2)

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

## How It Works

The SpreadsheetLLM encoder works through several key steps:

### 1. Structural Anchor Detection

The encoder identifies key structural points in the spreadsheet (rows and columns) that define the layout, based on:
- Cell density changes
- Format transitions
- Content type boundaries

### 2. Cell Neighborhood Extraction

Using the identified anchors, the encoder extracts cells within a configurable distance (parameter `k`). This creates a neighborhood around important structural elements while ignoring less relevant areas.

### 3. Inverted Index Creation

Instead of storing each cell individually, the encoder creates an inverted index mapping content values to cell references. This significantly reduces redundancy and creates a more compact representation.

### 4. Format Region Aggregation

Cell formats are aggregated into rectangular regions. During this step the encoder
infers a semantic type (e.g. year, percentage, date) from each cell's number format
string and groups contiguous cells that share both the type and the raw format
string. This greatly reduces repeated style data in the output and improves
compression.

### 5. JSON Encoding

The final output is a structured JSON document containing:
- File metadata
- Sheet information
- Structural anchors
- Cell values (inverted index)
- Format regions

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
        "Header": ["A1", "B1", "C1"],
        "42": ["B5"],
        "Total": ["A10"]
      },
      "formats": {
        "{format_definition}": ["A1:C1", "A10:F10"]
      }
    }
  }
}
```

## Research Background

This implementation is based on the paper "[SpreadsheetLLM: Enabling LLMs to Understand Spreadsheets](https://www.microsoft.com/en-us/research/)" published by Microsoft Research in July 2024. The paper introduces a novel approach to encode spreadsheets for LLM comprehension that preserves structural integrity and visual semantics.

## License

This project is licensed under the [MIT License](LICENSE).

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
