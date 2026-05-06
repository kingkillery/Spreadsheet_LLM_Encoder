# Datasets for SpreadsheetLLM evaluation

## Why this directory exists

The SpreadsheetLLM paper (arXiv:2407.09025, Microsoft Research) introduced:
- A **188-spreadsheet table-detection benchmark** with quality-improved annotations
- A **64-spreadsheet QA dataset** for question-answering evaluation

These datasets were not publicly released. To make the evaluation pipeline runnable **end-to-end without waiting for paper dataset release**, this repository provides:

1. **Synthetic data generators** that emit the same on-disk format as `evaluation.py` expects
2. **Guidance for plugging in real datasets** (public or in-house)

This allows you to:
- Smoke-test the entire pipeline immediately
- Bootstrap fine-tuning with synthetic data
- Substitute real datasets when available

## Annotation format

### Table detection: `.json` paired with `.xlsx`

Each spreadsheet has a corresponding JSON annotation file with table bounds:

```json
{
  "tables": [
    {
      "range": "A1:D5",
      "sheet": "Sheet1"
    },
    {
      "range": "A8:F12",
      "sheet": "Sheet1"
    }
  ]
}
```

The `range` format is Excel-style: `A1:D5` means columns A–D, rows 1–5. This is loaded by `evaluation.py:load_spreadsheet_dataset`:

```python
from evaluation import load_spreadsheet_dataset

dataset = load_spreadsheet_dataset("datasets/my_benchmark/")
# Returns: [
#   {
#     "spreadsheet_path": "...",
#     "bboxes": [(1, 1, 5, 4), (8, 1, 12, 6)],  # (row_start, col_start, row_end, col_end)
#     "ann_path": "..."
#   },
#   ...
# ]
```

### QA task: `.json` paired with `.xlsx`

For spreadsheet QA evaluation, annotations are pairs of questions and answers:

```json
{
  "qa_pairs": [
    {
      "question": "What is the total revenue in Q1?",
      "answer": "[45000]"
    },
    {
      "question": "Which region has the highest profit?",
      "answer": "[North]"
    }
  ]
}
```

Answer strings are wrapped in `[...]` to match the paper's QA evaluation protocol. Loaded by `evaluation.py:load_qa_dataset`:

```python
from evaluation import load_qa_dataset

dataset = load_qa_dataset("datasets/my_qa_set/")
# Returns: [
#   {
#     "spreadsheet_path": "...",
#     "qa_pairs": [...]
#   },
#   ...
# ]
```

## Synthesizing your own benchmark

### Generate table-detection annotations

Use `scripts/synthesize_benchmark.py` to create annotated synthetic spreadsheets:

```bash
python scripts/synthesize_benchmark.py datasets/synthetic_v1/ --n 50 --seed 42
```

**Parameters**:
- `output_dir`: Directory to write `.xlsx` and `.json` files
- `--n`: Number of spreadsheets to generate (default: 50)
- `--seed`: Random seed for reproducibility (default: 42)
- `--small-only`: Generate only small tables (5–8 rows, 3–4 columns)
- `--include-multi-table`: Allow 1–3 tables per spreadsheet (default: 1)

**Output**: Creates `synth_0000.xlsx` + `synth_0000.json`, `synth_0001.xlsx` + `synth_0001.json`, etc.

**Layouts and content types**:
- **Layouts**: simple grid, row headers, merged title, blank-row separator
- **Content types**: integer, float, date, currency, email, text
- **Size classes**: small (5–8 rows), medium (15–25 rows), large (40–60 rows)

The synthetic generator cycles through layouts and sizes, ensuring variety while remaining deterministic under a fixed seed.

### Generate QA pairs from spreadsheets

To create question-answer pairs from your spreadsheets (synthetic or real):

```bash
python scripts/synthesize_qa.py datasets/synthetic_v1/ --per-table 4 --seed 42
```

**Parameters**:
- `dataset_dir`: Directory containing `.xlsx` files (and optionally existing `.json` annotations)
- `--per-table`: Number of QA pairs to generate per table (default: 4)
- `--seed`: Random seed (default: 42)

**Output**: Creates or updates `.json` files with `"qa_pairs"` arrays. Existing table annotations are preserved.

**Question types**:
- **Cell lookup**: "What is the value in B3?"
- **Column aggregation**: "What is the sum of the Revenue column?"
- **Cross-column comparison**: "Which row has the highest Score?"
- **Header lookup**: "What is the column header for column C?"

Answers are wrapped in `[...]` format (e.g., `[42]` or `[North]`) to match the paper's evaluation contract.

## Plugging in real datasets

Any directory with paired `.xlsx` and `.json` files following the annotation schema above will work with `load_spreadsheet_dataset` and `load_qa_dataset`.

### Real dataset candidates

1. **TableSense (Dong et al., 2019, AAAI)**
   - Public table detection benchmark
   - Request from authors or find on AAAI Proceedings
   - Format: Images + Pascal VOC XML (use `evaluation.py:load_dong2019_dataset` for the image-based format)

2. **HuggingFace Hub or Kaggle**
   - Search for "spreadsheet benchmark", "table detection", or "Excel annotation"
   - Examples: TableNet, PubTabNet (though primarily for PDF/image tables)
   - May require format conversion to `.xlsx` + `.json`

3. **Custom in-house datasets**
   - Annotate with the `.json` schema above
   - Spreadsheet format: any `.xlsx` workbook
   - One JSON per spreadsheet, with `"tables"` and/or `"qa_pairs"` arrays

### Format conversion example

If you have a different annotation format (e.g., CSV with column names `filepath`, `table_range`):

```python
import json
from pathlib import Path

def convert_csv_to_json(csv_path, output_dir):
    import pandas as pd
    df = pd.read_csv(csv_path)
    for _, row in df.iterrows():
        json_path = Path(output_dir) / row['filepath'].replace('.xlsx', '.json')
        annotation = {
            "tables": [
                {"range": row['table_range']}
            ]
        }
        with open(json_path, 'w') as f:
            json.dump(annotation, f)

convert_csv_to_json('my_dataset.csv', 'datasets/my_benchmark/')
```

## Known gaps vs. the paper

### Synthetic data limitations

1. **Narrower vocabulary**: Synthetic data uses a fixed wordlist (alpha, bravo, ..., north, south, etc.); real spreadsheets have arbitrary text, domain-specific terms, multi-language content.

2. **Layout diversity**: Synthetic layouts are templated (simple grid, row headers, merged title, blank separators); real spreadsheets have nested headers, multiple title regions, inconsistent formatting, and artistic layouts.

3. **No quality improvement pipeline**: The paper re-annotated the original 188 spreadsheets using human reviewers and LLM refinement. Synthetic data skips this; annotations are deterministic and rule-based, not human-validated.

4. **Synthetic QA**: Questions are template-based with simple answer formats. Real QA from the paper includes:
   - Multi-step reasoning (e.g., "Find the profit in Q2, then calculate its percentage of annual total")
   - Ambiguity resolution (e.g., cell values that match multiple columns)
   - Contextual answers (e.g., "The highest value is in row 5" vs. `[12345]`)

### Evaluation differences

- **F1 scores not directly comparable**: Without the paper's exact data splits and quality-improvement annotations, F1 values from your test set cannot be compared to the paper's reported 0.85 EoB-0 metric.
- **Distribution mismatch**: Fine-tuning on synthetic data may underperform on real spreadsheets; use real data for production models.

### Recommendations

- **Smoke testing**: Synthetic data is sufficient
- **Bootstrapping fine-tuning**: Use synthetic data for initial exploration, then transition to real data for quality
- **Reproduction**: For claims matching the paper's results, use the paper's dataset (request from authors or wait for potential future release)

## Next steps

1. Start with synthetic data: `python scripts/synthesize_benchmark.py datasets/synthetic_v1/ --n 20`
2. Run the evaluation pipeline: `python run_llm_evaluation.py datasets/synthetic_v1/`
3. Prepare fine-tuning data: `python prepare_finetuning_data.py datasets/synthetic_v1/ output/finetune.jsonl`
4. Substitute real datasets when available (keep the same `.xlsx` + `.json` structure)
