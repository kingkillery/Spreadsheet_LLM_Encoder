# Golden Workbook Fixtures and Snapshot Tests Scope

## Goal

Add persistent golden workbook fixtures and serializer/encoding snapshot tests
that lock the repository's core paper-facing format contract.

This is the current highest-leverage next step because tokenizer metadata and
claim guardrails are now in place. The biggest remaining risk is silent
behavioral drift in prompts, coordinate maps, anchors, and metrics.

## Fixture Directory

Create:

```text
tests/fixtures/
  simple_table.xlsx
  sparse_sheet.xlsx
  merged_headers.xlsx
  multi_table_sheet.xlsx
  date_currency_percent.xlsx
  formula_cells.xlsx
  hidden_rows_cols.xlsx
  wide_sparse_sheet.xlsx
```

## Golden Directory

Create:

```text
tests/golden/
  simple_table.vanilla.txt
  simple_table.compressed.txt
  simple_table.encoding.json
  simple_table.metrics.json
```

Repeat equivalent golden files for each fixture where useful.

## Snapshot Coverage

For each fixture, verify:

- `to_paper_vanilla_prompt` exactly matches the golden pair-string.
- `to_paper_compressed_prompt` exactly matches the golden compressed prompt.
- `coord_map` round-trips compact/original ranges.
- Encoding JSON matches stable structural fields.
- Metrics include tokenizer metadata.
- Metrics are stable under a known tokenizer backend or documented fallback.

## Acceptance Criteria

- Golden fixtures are committed to the repo, not generated only at test time.
- Snapshot tests fail on prompt-format drift.
- Snapshot tests fail on unexpected coordinate-map drift.
- Snapshot tests fail if tokenizer metadata is missing.
- Known nondeterministic fields are normalized before comparison.
- Existing synthetic fixture tests remain useful as broader behavioral coverage.

## Recommended Test File

Create:

```text
test_encoder_golden.py
```

Suggested test groups:

- vanilla prompt snapshots
- compressed prompt snapshots
- stable encoding JSON snapshots
- metric/tokenizer metadata snapshots
- coordinate remap round-trip properties

## Out of Scope

- Reproducing paper-original metrics.
- Adding original paper datasets.
- Refactoring the encoder into packages.
- Adding real LLM benchmark runs.
