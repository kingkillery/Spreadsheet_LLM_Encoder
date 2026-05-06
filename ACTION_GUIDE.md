# SpreadsheetLLM Encoder PRD / Action Guide

## Final Pass Verdict

`Spreadsheet_LLM_Encoder` is now best described as a **paper-facing
SpreadsheetLLM implementation scaffold**, not a simple spreadsheet serializer.
It has meaningful implementations for the core encoding surface: vanilla
pair-string baseline, structural-anchor extraction, inverted-index translation,
data-format aggregation, coordinate remapping, paper-style serializers,
Chain-of-Spreadsheet Stage 1/Stage 2, row chunking, evaluation metadata,
optional baselines, bounded large-workbook encoding, and formula dependency
metadata.

The repo correctly warns that it is **not a paper-metric reproduction** unless
the original datasets, splits, models, procedures, and baselines are supplied.

The highest-value work now is not "add the basic paper ideas." Those are mostly
present. The next work is:

1. harden correctness,
2. prove paper-fidelity claims,
3. modularize the implementation,
4. add missing benchmark assets / baselines,
5. deepen spreadsheet reasoning capabilities such as formula lineage and
   formula-aware prompting.

## 1. Verified Implementation Map

### Encoder Core

The main encoder loads workbooks, computes the paper vanilla prompt token
baseline, extracts structural anchors, optionally prunes homogeneous
rows/columns, builds an inverted index, merges ranges, aggregates data-format
regions, builds compact coordinate maps, serializes paper-compressed prompts,
records per-stage compression metrics, and can run in bounded mode for large
workbooks.

```text
Workbook
-> vanilla pair-string token baseline
-> structural anchors
-> retained row/column skeleton
-> optional homogeneous pruning
-> inverted index
-> merged address ranges
-> semantic type / number-format aggregation
-> numeric range clustering
-> compact coordinate remapping
-> compressed prompt
-> compression metrics
-> sheet_processing metadata
```

### Paper Serializers

`paper_serializers.py` provides the important paper-facing prompt surfaces:

- `to_paper_vanilla_prompt`
- `to_paper_compressed_prompt`
- `to_stage2_uncompressed_prompt`
- coordinate remapping and unremapping helpers

It also enforces a max range size to avoid pathological range expansion.

### Chain-of-Spreadsheet

`chain_of_spreadsheet.py` implements the CoS flow:

- configurable LLM backend,
- Stage 1 compressed prompt table identification,
- multi-range extraction helper,
- sheet selection with LLM-first / keyword fallback,
- Stage 2 original-workbook uncompressed prompt path,
- legacy compressed fallback,
- Algorithm 2-style row chunking and final synthesis.

### Evaluation

The repo has a reproducibility protocol that separates **synthetic**,
**reconstructed**, and **paper-original** claims. Evaluation records now carry a
machine-readable `claim_level`, and `paper-original` claims fail validation
unless required parity metadata is present.

`run_llm_evaluation.py` implements compressed-prompt table detection,
compact-to-original coordinate unremapping, EoB-0 evaluation, and structured
output records.

`evaluation_metadata.py` defines required metadata fields and validates
evaluation records.

### Baselines and Backends

The repo includes a real optional TaPEx wrapper and explicitly marks Binder
unavailable rather than silently faking it.

`llm_backend.py` provides `EchoBackend`, `CallableBackend`, and `OpenAIBackend`.

`tokenizer.py` uses `tiktoken` when available and falls back to deterministic
char/4 estimates.

## 2. Current Capability Matrix

| Area | Current state | Action needed |
|---|---:|---|
| Vanilla pair-string baseline | Implemented | Add persisted golden tests against fixture files |
| Paper compressed prompt | Implemented | Add snapshot tests |
| Structural anchors | Approximate / heuristic | Build rule-by-rule parity matrix |
| Inverted index | Mostly aligned | Add edge-case tests for merged / repeated / sparse cells |
| Data-format aggregation | Mostly aligned | Expand fixtures for date, time, currency, percent, email, scientific notation |
| Coordinate remapping | Implemented | Expand property tests |
| CoS Stage 1 | Implemented | Harden LLM output parsing |
| CoS Stage 2 | Implemented with original workbook path | Make fallback opt-in / explicit |
| Algorithm 2 row chunking | Implemented | Add stress tests and token-limit tests |
| Evaluation metadata | Implemented | Keep validation in CI |
| Claim-level validation | Implemented | Add CLI flags where missing |
| Table detection evaluation | Scaffold implemented | Add real/synthetic fixture pack |
| QA evaluation | Scaffold implemented | Add reliable local smoke dataset |
| TaPEx | Optional real wrapper | Add docs and CI skip behavior |
| Binder | Explicitly unavailable | Implement adapter or keep as structured skip |
| Formula semantics | Lightweight graph implemented | Add formula-aware prompt mode |
| Large workbook handling | Bounded mode implemented | Add semantic sampling / streaming later |
| Architecture | Working but monolithic | Refactor into package modules |
| Paper reproduction | Not established | Add `paper_reproduction/` gated workflow |

## 3. Priority Action Guide

### P0 - Protect Claims and Stabilize the Repo

#### 1. Claim Mode for Evaluation Outputs

Status: **mostly implemented**.

Evaluation metadata now includes `claim_level`, and `paper-original` fails
without concrete parity metadata.

Remaining work:

- Add explicit CLI flags where useful:

```bash
--claim-level synthetic|reconstructed|paper-original
```

Validation rules:

| Claim level | Required fields |
|---|---|
| `synthetic` | dataset name/version, generator/source |
| `reconstructed` | reconstruction source, split method, known divergences |
| `paper-original` | original data ID, split ID, model/backend, metric definition, baseline status |

Acceptance criteria:

- `paper-original` fails unless required metadata is present.
- output record includes `claim_level`.
- README examples use `synthetic` by default.
- CI has one test proving invalid `paper-original` metadata fails.

#### 2. Make `tiktoken` Status Explicit in Output

The tokenizer module already falls back to char/4 when `tiktoken` is
unavailable. The output JSON should record which tokenizer path was used.

Add to `compression_metrics`:

```json
{
  "tokenizer": {
    "model": "gpt-4",
    "backend": "tiktoken",
    "fallback": false
  }
}
```

Acceptance criteria:

- metrics record whether fallback was used.
- tests cover both `tiktoken` and forced-fallback modes.
- paper-comparable claim fails if fallback is used.

#### 3. Binder Status as Mandatory Evaluation Metadata

Status: **partially implemented**.

Binder is explicitly unavailable and QA records include Binder status/skip
reason. Keep this invariant in future evaluation paths.

Acceptance criteria:

- every QA run includes `binder_status`.
- every QA run includes `skip_reasons`.
- no output shows Binder as `0%` accuracy unless a real Binder adapter exists.

## 4. P1 - Prove Encoder Correctness

### 4.1 Add Golden Workbook Fixtures

Current status: `test_paper_parity_fixtures.py` creates deterministic synthetic
fixtures at test time. The next step is to persist golden files.

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

tests/golden/
  simple_table.vanilla.txt
  simple_table.compressed.txt
  simple_table.encoding.json
  simple_table.metrics.json
```

Acceptance criteria:

- `to_paper_vanilla_prompt` exactly matches golden pair-string.
- `to_paper_compressed_prompt` exactly matches golden compressed prompt.
- `coord_map` round-trips compact -> original ranges.
- compression metrics are stable within a known tokenizer backend.

### 4.2 Add Structural-Anchor Parity Tests

Current structural anchor detection is intentionally approximate. The README
says it is Appendix C-inspired but not complete.

Create `docs/STRUCTURAL_ANCHOR_PARITY.md`:

| Rule | Paper behavior | Current implementation | Test fixture | Status |
|---|---|---|---|---|
| boundary changes | detects row/col discontinuities | implemented via profile changes | `simple_table.xlsx` | pass |
| sparse rectangle rejection | rejects low-density candidates | implemented via density threshold | `sparse_sheet.xlsx` | pass |
| edge sparsity | rejects huge sparse rectangles | implemented | `notes_far_away.xlsx` | pass |
| header/title/note handling | detailed paper rules | partial | `title_note_table.xlsx` | partial |
| overlap resolution | table relationship heuristics | IoU heuristic | `overlap_tables.xlsx` | pass/partial |

Acceptance criteria:

- Each rule has a fixture.
- Current non-parity is documented, not hidden.
- Test names distinguish paper-strict behavior from pragmatic heuristics.

### 4.3 Add Coordinate Property Tests

Paper serializers already include remap/unremap helpers.

Property:

```python
original_range == unremap_range(remap_range(original_range, coord_map), coord_map)
```

Add cases for:

- single cell,
- rectangle,
- non-contiguous retained rows,
- non-contiguous retained columns,
- invalid unmappable ranges,
- JSON-reloaded maps with string keys.

Acceptance criteria:

- all compact ranges predicted in Stage 1 can be unmapped before Stage 2.
- unmappable ranges fail gracefully with structured warnings.

## 5. P2 - Refactor Architecture

The main encoder is doing too much in one file: workbook loading, anchors,
compression, indexing, formatting, metrics, CLI, bounded processing, formula
metadata, and serialization orchestration.

Move toward:

```text
spreadsheet_llm_encoder/
  __init__.py
  cli.py
  encode.py
  workbook.py
  anchors.py
  compression.py
  inverted_index.py
  format_aggregation.py
  coordinate_map.py
  serializers.py
  metrics.py
  formulas.py
  cos.py
  eval/
    table_detection.py
    qa.py
    metadata.py
  baselines/
    tapex.py
    binder.py
  tests/
```

Suggested PR split:

1. Package skeleton, no behavior change.
2. Extract anchors.
3. Extract compression/indexing.
4. Extract metrics and tokenizer metadata.
5. Extract formulas.

Acceptance criteria:

- `spreadsheet-llm-encode` still works.
- old imports remain backward-compatible:

```python
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode
```

- all tests pass.

## 6. P3 - Harden CoS Behavior

### 6.1 Make Stage 2 Fallback Explicit

`generate_response` supports a compressed JSON fallback when workbook path /
sheet / range are not provided. That is useful for legacy use, but it should
not happen silently in paper-facing runs.

Implement:

```python
generate_response(..., allow_compressed_fallback=False)
```

Acceptance criteria:

- default paper path raises a clear error if original workbook data is missing.
- legacy callers can pass `allow_compressed_fallback=True`.
- evaluation scripts never use compressed fallback for paper-comparable claims.

### 6.2 Improve LLM Range Parsing

Current regex accepts quoted and bare ranges. Add support for common variants:

```text
Sheet1!A1:B5
'Sheet 1'!A1:B5
range=A1:B5
{"range": "A1:B5"}
[{"range": "A1:B5"}]
```

Acceptance criteria:

- parser returns structured objects.
- handles multiple ranges.
- deduplicates while preserving order.
- rejects impossible ranges cleanly.

### 6.3 Strengthen Multi-Sheet Selection

Current sheet selection uses LLM if available, then keyword fallback.

Add a deterministic ranking fallback:

```text
score =
  sheet-name match
+ header/value token match
+ number-format relevance
+ table-size prior
+ anchor density prior
```

Acceptance criteria:

- fallback returns ranked candidates, not just one.
- Stage 1 can evaluate top N sheets when ambiguity exists.
- evaluation record includes selected sheet and fallback reason.

## 7. P4 - Evaluation and Benchmark Plan

### 7.1 Add Synthetic Benchmark Pack

Create a small bundled dataset that is **not** paper-comparable, but is useful
for regression.

```text
datasets/synthetic/
  table_detection/
    manifest.json
    simple_table.xlsx
    multi_table.xlsx
    merged_headers.xlsx
  qa/
    manifest.json
    revenue_lookup.xlsx
    formula_answer.xlsx
```

Acceptance criteria:

- `python run_llm_evaluation.py --manifest ... --backend echo` runs locally.
- `python run_qa_evaluation.py --manifest ... --backend echo` runs locally.
- output records pass metadata validation.

### 7.2 Add Real LLM Smoke Workflow

Add a CI/manual workflow that can run only when `OPENAI_API_KEY` is set.

```text
.github/workflows/llm-smoke.yml
```

Acceptance criteria:

- skipped safely without key.
- runs 2-3 tiny synthetic examples.
- records model name, prompt serializer, coordinate mode, token backend.

### 7.3 Add Paper-Reproduction Gate

Create:

```text
paper_reproduction/
  README.md
  manifest.schema.json
  run_table_detection.sh
  run_qa.sh
  REQUIRED_ASSETS.md
```

This path should **not** ship fake paper data. It should document what must be
supplied:

```text
original paper spreadsheets
original table annotations
original QA items
exact split metadata
model/backend config
baseline availability
metric definitions
fine-tuning recipe/version
```

Acceptance criteria:

- cannot run `--claim-level paper-original` without manifest fields.
- docs explain how to label results accurately.
- result record includes all required metadata.

## 8. P5 - Formula Lineage Extension

Status: **partially implemented**.

The encoder now emits a lightweight `formula_graph` when formulas or spreadsheet
error cells are present. The next phase is formula-aware prompt integration.

Current graph captures:

- formula cell,
- formula string,
- cached value,
- same-sheet references,
- cross-sheet references,
- formula errors,
- repeated-formula summaries.

Future API:

```python
spreadsheet_llm_encode(
    excel_path,
    preserve_formulas=True,
    formula_prompt_mode=True,
)
```

Future output shape:

```json
{
  "formulas": {
    "Sheet1!D8": {
      "formula": "=SUM(D2:D7)",
      "cached_value": 12345,
      "references": ["Sheet1!D2:D7"],
      "depends_on": ["Sheet1!D2", "Sheet1!D3"]
    }
  }
}
```

Acceptance criteria:

- default prompt remains paper-style visible values.
- formula-aware prompt mode is opt-in.
- formulas do not affect paper-comparable metrics unless explicitly enabled.
- tests cover same-sheet, cross-sheet, ranges, formula errors, and repeated
  formula families.

## 9. P6 - Baseline Roadmap

### TaPEx

Already implemented as an optional wrapper with lazy loading.

Action items:

- add `--tapex-max-rows`;
- include TaPEx model ID in metadata;
- include truncation flag when body rows exceed max rows;
- add docs for installing optional baseline dependencies.

### Binder

Currently unavailable by design.

Options:

- Option A: keep unavailable. This is acceptable if the repo is not claiming
  full paper reproduction.
- Option B: implement a real adapter. This is needed for paper-comparable
  baseline claims.

Minimum interface:

```python
class BinderBaseline:
    def answer(
        self,
        workbook_path: str,
        sheet_name: str,
        table_range: str,
        query: str,
    ) -> str:
        ...
```

Acceptance criteria:

- no placeholder predictions;
- records Binder version/config;
- supports skip reasons for failed optional dependencies;
- tests use a mocked execution loop.

## 10. P7 - Documentation Upgrades

### Add `PAPER_FIDELITY.md`

Include:

```text
paper concept
repo implementation
file/function
current status
known divergence
test coverage
claim impact
```

### Add `ARCHITECTURE.md`

Include:

```text
encoder pipeline
prompt serializers
CoS flow
evaluation flow
metadata contract
baseline contract
```

### Add `LIMITATIONS.md`

State clearly:

- original paper datasets are not bundled;
- Binder is unavailable;
- structural anchors are approximate;
- formula-aware prompts are not currently modeled;
- paper-comparable claims require strict metadata and assets.

## 11. Suggested Issue Backlog

### Milestone 1 - Correctness Hardening

1. Add tokenizer backend metadata to compression metrics.
2. Add golden workbook fixtures.
3. Add coordinate remap property tests.
4. Add structural-anchor parity matrix.
5. Add CLI claim-level flags.

### Milestone 2 - Evaluation Readiness

6. Add bundled synthetic table-detection dataset.
7. Add bundled synthetic QA dataset.
8. Add manifest schemas.
9. Add real LLM smoke workflow.
10. Add result-record validation to CI.

### Milestone 3 - Architecture Cleanup

11. Convert flat modules into package layout.
12. Extract anchors module.
13. Extract inverted-index module.
14. Extract format aggregation module.
15. Extract metrics module.
16. Maintain backward-compatible imports.

### Milestone 4 - Paper-Fidelity Path

17. Add `paper_reproduction/` directory.
18. Add paper-original manifest requirements.
19. Add paper-comparable run guard.
20. Add fine-tune/eval compatibility checks to CLI.
21. Add baseline status checks.

### Milestone 5 - Advanced Reasoning

22. Add formula-aware prompt mode.
23. Expand formula dependency graph into explicit precedent cells.
24. Add cross-sheet reference edge cases.
25. Add tests for formula-heavy workbooks.

### Milestone 6 - Operational Scale

26. Add semantic sampling for bounded mode.
27. Add per-sheet include/exclude controls.
28. Add `.xlsb` reader support.
29. Add streaming/token-budget prompt generation.

## 12. Recommended Next Three PRs

### PR 1: Tokenizer and Claim Guardrails

Files likely touched:

```text
tokenizer.py
Spreadsheet_LLM_Encoder.py
evaluation_metadata.py
run_llm_evaluation.py
run_qa_evaluation.py
EVALUATION.md
test_evaluation_metadata.py
```

Goal: record tokenizer backend/fallback status and prevent accidental
paper-comparable claims with fallback tokenization.

### PR 2: Golden Fixtures and Serializer Tests

Files likely touched:

```text
tests/fixtures/
tests/golden/
test_paper_serializers.py
test_encoder_golden.py
```

Goal: lock down prompt format and coordinate behavior.

### PR 3: Encoder Modularization, No Behavior Change

Files likely touched:

```text
spreadsheet_llm_encoder/
Spreadsheet_LLM_Encoder.py
pyproject.toml
tests/
```

Goal: reduce risk and make future paper-parity work easier.

## 13. Bottom Line

The repo is in a good state for a **paper-facing implementation**, but the next
phase should be about proof and guardrails:

```text
Current state:
  meaningful implementation of SpreadsheetLLM-style encoding and CoS flow

Next state:
  tested, modular, reproducible, claim-safe implementation

Later state:
  paper-comparable reproduction path + optional formula-aware reasoning
```

The most important immediate action is to **turn the current implementation
into a verifiable system**: golden fixtures, tokenizer metadata, claim-level
enforcement at every CLI boundary, metadata validation, and structural-anchor
parity documentation.
