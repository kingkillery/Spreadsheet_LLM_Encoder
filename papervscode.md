# SpreadsheetLLM Paper vs. `Spreadsheet_LLM_Encoder`

This document compares the SpreadsheetLLM paper with the local
`Spreadsheet_LLM_Encoder` codebase. It also provides graph-ready node and edge
specifications that can be pasted into a knowledge-graph tool.

## Sources and Verification

External paper source:

- arXiv: [SpreadsheetLLM: Encoding Spreadsheets for Large Language Models](https://arxiv.org/abs/2407.09025)

Local repository evidence:

- `README.md`
- `EVALUATION.md`
- `FINAL_POLISH_SPEC.md`
- `Spreadsheet_LLM_Encoder.py`
- `paper_serializers.py`
- `chain_of_spreadsheet.py`
- `evaluation.py`
- `evaluation_metadata.py`
- `baselines.py`
- `prepare_finetuning_data.py`
- `pyproject.toml`
- Existing test suite: `test_*.py`

The original draft of this file was speculative because it was not verified
against the repository. This version is based on the current local files.

## 1. Knowledge Graph: SpreadsheetLLM Paper

### Core nodes

| Node | Type | Meaning |
|---|---|---|
| SpreadsheetLLM | Project / paper | Framework for encoding spreadsheets so LLMs can reason over workbook structure. |
| SheetCompressor | Method | Main spreadsheet compression framework introduced by the paper. |
| Chain of Spreadsheet | Method | Multi-stage spreadsheet QA pipeline that identifies relevant table regions before answer generation. |
| Vanilla serialization | Baseline | Pair-string spreadsheet representation without SheetCompressor compression. |
| Structural anchors | Technique | Row and column skeleton extraction for layout preservation. |
| Inverted index translation | Technique | Token-efficient representation that groups repeated cell values by address ranges. |
| Data-format aggregation | Technique | Compression of cells that share meaningful number formats or semantic data classes. |
| Coordinate remapping | Technique | Compact coordinate system used after structural compression. |
| Table detection | Task | Detect table boundaries in messy worksheets. |
| Spreadsheet QA | Task | Answer questions whose answers are cell addresses, formulas, or values in a spreadsheet. |
| EoB-0 exact boundary matching | Metric | Exact table-boundary evaluation metric used for table detection. |
| Fine-tuning | Procedure | Training LLMs on compressed spreadsheet prompts for table detection. |
| Baseline comparison | Evaluation | Comparison against vanilla serialization and prior table/spreadsheet methods. |

### Key edges

```json
[
  {
    "source": "SpreadsheetLLM",
    "target": "SheetCompressor",
    "relationship": "uses",
    "description": "SpreadsheetLLM uses SheetCompressor as its main encoding and compression framework for turning spreadsheet grids into LLM-readable prompts."
  },
  {
    "source": "SheetCompressor",
    "target": "Structural anchors",
    "relationship": "contains",
    "description": "SheetCompressor contains structural-anchor extraction to preserve rows and columns that define spreadsheet layout and table boundaries."
  },
  {
    "source": "SheetCompressor",
    "target": "Inverted index translation",
    "relationship": "contains",
    "description": "SheetCompressor contains inverted index translation to represent repeated cell values with compact address regions instead of repeated literal text."
  },
  {
    "source": "SheetCompressor",
    "target": "Data-format aggregation",
    "relationship": "contains",
    "description": "SheetCompressor contains data-format aggregation to compress cells with shared numeric or semantic formats while retaining spreadsheet meaning."
  },
  {
    "source": "Structural anchors",
    "target": "Coordinate remapping",
    "relationship": "enables",
    "description": "Structural anchors enable coordinate remapping because only the retained spreadsheet skeleton needs compact row and column addresses."
  },
  {
    "source": "Chain of Spreadsheet",
    "target": "Table detection",
    "relationship": "uses",
    "description": "Chain of Spreadsheet uses table detection as its first stage so later question answering can focus on the relevant table region."
  },
  {
    "source": "Chain of Spreadsheet",
    "target": "Spreadsheet QA",
    "relationship": "supports",
    "description": "Chain of Spreadsheet supports spreadsheet QA by narrowing the workbook to a relevant table and then asking the LLM for the answer cell or formula."
  },
  {
    "source": "Table detection",
    "target": "EoB-0 exact boundary matching",
    "relationship": "evaluated by",
    "description": "Table detection is evaluated by EoB-0 exact boundary matching, which requires predicted table boundaries to match the ground truth exactly."
  },
  {
    "source": "Fine-tuning",
    "target": "Table detection",
    "relationship": "improves",
    "description": "Fine-tuning improves table detection by training an LLM on SheetCompressor prompt formats and table-boundary targets."
  },
  {
    "source": "Vanilla serialization",
    "target": "SheetCompressor",
    "relationship": "contrasts with",
    "description": "Vanilla serialization contrasts with SheetCompressor because it preserves raw grid contents but does not apply the paper's compression stages."
  }
]
```

## 2. Knowledge Graph: Local Codebase

### Core nodes

| Node | Type | Verified local evidence |
|---|---|---|
| `Spreadsheet_LLM_Encoder.py` | Runtime module | Main CLI and encoder implementation. |
| `spreadsheet_llm_encode` | Function | Encodes workbooks, computes compression metrics, and writes JSON artifacts. |
| `find_structural_anchors` | Function | Runs boundary-candidate extraction and k-neighborhood expansion. |
| `compress_homogeneous_regions` | Function | Optional pragmatic pruning after anchor extraction. |
| `create_inverted_index` | Function | Builds value-to-address maps and format maps. |
| `aggregate_regions_dfs` | Function | Groups contiguous cells with shared format keys. |
| `paper_serializers.py` | Runtime module | Exports paper-style prompt serializers and coordinate remapping helpers. |
| `to_paper_vanilla_prompt` | Function | Emits vanilla pair-string prompts. |
| `to_paper_compressed_prompt` | Function | Emits compressed pair-string prompts with format-region substitution. |
| `to_stage2_uncompressed_prompt` | Function | Emits uncompressed pair-string prompts for a Stage 2 table range. |
| `coord_map` | Artifact field | Stores original-to-compact and compact-to-original coordinate maps. |
| `chain_of_spreadsheet.py` | Runtime module | Implements CoS-style table selection, Stage 2 response generation, and table splitting. |
| `evaluation.py` | Runtime module | Provides table-detection and QA dataset loading and scoring. |
| `evaluation_metadata.py` | Runtime module | Records reproducibility metadata for evaluation outputs. |
| `baselines.py` | Runtime module | Provides optional TaPEx baseline wrapper and explicit Binder-unavailable adapter. |
| `prepare_finetuning_data.py` | Runtime module | Creates JSONL training records and metadata sidecars. |
| `pyproject.toml` | Packaging metadata | Declares package module list, entry point, and optional dependency groups. |

### Key edges

```json
[
  {
    "source": "spreadsheet_llm_encode",
    "target": "find_structural_anchors",
    "relationship": "calls",
    "description": "spreadsheet_llm_encode calls find_structural_anchors to identify rows and columns that should survive the compression skeleton."
  },
  {
    "source": "spreadsheet_llm_encode",
    "target": "compress_homogeneous_regions",
    "relationship": "optionally calls",
    "description": "spreadsheet_llm_encode optionally calls compress_homogeneous_regions for pragmatic row and column pruning unless strict paper mode is requested."
  },
  {
    "source": "spreadsheet_llm_encode",
    "target": "create_inverted_index",
    "relationship": "calls",
    "description": "spreadsheet_llm_encode calls create_inverted_index to map repeated values and format groups into compact address-range representations."
  },
  {
    "source": "spreadsheet_llm_encode",
    "target": "aggregate_regions_dfs",
    "relationship": "calls",
    "description": "spreadsheet_llm_encode calls aggregate_regions_dfs to group contiguous cells that share semantic type and number-format keys."
  },
  {
    "source": "spreadsheet_llm_encode",
    "target": "coord_map",
    "relationship": "produces",
    "description": "spreadsheet_llm_encode produces coord_map so compact prompt coordinates can be remapped back to original workbook coordinates."
  },
  {
    "source": "paper_serializers.py",
    "target": "to_paper_vanilla_prompt",
    "relationship": "exports",
    "description": "paper_serializers.py exports to_paper_vanilla_prompt for the paper's vanilla pair-string baseline and uncompressed prompt surface."
  },
  {
    "source": "paper_serializers.py",
    "target": "to_paper_compressed_prompt",
    "relationship": "exports",
    "description": "paper_serializers.py exports to_paper_compressed_prompt for compressed prompts with value tuples and data-format aggregation tuples."
  },
  {
    "source": "chain_of_spreadsheet.py",
    "target": "to_paper_compressed_prompt",
    "relationship": "uses",
    "description": "chain_of_spreadsheet.py uses to_paper_compressed_prompt during Stage 1 table identification when an LLM backend is configured."
  },
  {
    "source": "chain_of_spreadsheet.py",
    "target": "to_stage2_uncompressed_prompt",
    "relationship": "uses",
    "description": "chain_of_spreadsheet.py uses to_stage2_uncompressed_prompt for paper-faithful Stage 2 prompts when the original workbook, sheet, and table range are available."
  },
  {
    "source": "evaluation.py",
    "target": "EoB-0 exact boundary matching",
    "relationship": "implements",
    "description": "evaluation.py implements EoB-style boundary scoring for table-detection predictions and ground-truth table ranges."
  },
  {
    "source": "evaluation_metadata.py",
    "target": "Evaluation reproducibility",
    "relationship": "supports",
    "description": "evaluation_metadata.py supports reproducible experiments by recording dataset, split, encoder, prompt, backend, baseline, and metric settings."
  },
  {
    "source": "baselines.py",
    "target": "TaPEx baseline",
    "relationship": "implements",
    "description": "baselines.py implements an optional HuggingFace TaPEx wrapper for table question answering over extracted table ranges."
  },
  {
    "source": "baselines.py",
    "target": "Binder baseline",
    "relationship": "marks unavailable",
    "description": "baselines.py marks Binder as explicitly unavailable with a machine-readable skip reason until a real adapter is implemented."
  },
  {
    "source": "prepare_finetuning_data.py",
    "target": "Fine-tuning metadata",
    "relationship": "produces",
    "description": "prepare_finetuning_data.py produces JSONL records and optional sidecar metadata for auditable fine-tuning data generation."
  }
]
```

## 3. Capability Matrix

| Capability | Paper expectation | Local codebase status | Gap severity |
|---|---|---|---|
| Vanilla pair-string prompt | Required baseline and Stage 2 format | Implemented in `paper_serializers.to_paper_vanilla_prompt` and CLI `--vanilla` | Low |
| Structural anchors | Core SheetCompressor stage | Implemented as Appendix C-inspired heuristics, documented as approximate | Medium |
| Strict paper skeleton | Preserve anchor-neighborhood skeleton without extra pruning | Supported through `--no-compress-homogeneous` / strict mode behavior | Low-medium |
| Inverted index translation | Core SheetCompressor stage | Implemented by `create_inverted_index` and range merging | Low |
| Data-format aggregation | Core SheetCompressor stage | Implemented by semantic type plus Excel number-format keys and DFS region grouping | Low-medium |
| Coordinate remapping | Needed for compact prompts and range conversion | Implemented through `coord_map`, remap, and unremap helpers | Low |
| Stage 1 table identification | Compressed prompt plus LLM table-range prediction | Implemented when a backend is configured; fallback sheet selection remains pragmatic | Medium |
| Stage 2 QA prompt | Original workbook uncompressed sub-range | Implemented when `workbook_path`, `sheet_name`, and `table_range` are available | Low |
| Large-table splitting | Chunk large tables within token limits | Implemented in `table_split_qa` with row chunking and synthesis | Low-medium |
| Table-detection evaluation | EoB-0 exact boundary matching on paper dataset | EoB-style scaffold exists; paper dataset is not bundled | High for paper-comparable claims |
| Spreadsheet QA evaluation | Paper QA dataset and answer normalization | Manifest scaffold and answer normalization exist; paper dataset is not bundled | High for paper-comparable claims |
| Fine-tuning data generation | Training records for table detection | Implemented as JSONL preparation with metadata sidecar support | Medium |
| Paper metric reproduction | Original datasets, splits, model procedures, baselines | Not claimed; explicitly blocked without paper assets and procedures | High |
| TaPEx baseline | Runnable comparison baseline where dependencies exist | Optional wrapper implemented | Medium |
| Binder baseline | Prior baseline comparison | Explicitly unavailable until a real adapter is added | High |
| Formula dependency graph | Useful for spreadsheet reasoning | Formula text/value handling is limited; no explicit dependency graph is implemented | Medium-high |
| Rich visual style emission | Visual semantics can matter, but paper favors efficient format aggregation | Rich style keys exist in lower-level API, but default prompt does not emit full style metadata | Low-medium |
| Multi-table messy-sheet robustness | Important for real spreadsheets | Heuristics and tests exist, but not proven paper-equivalent | Medium-high |

## 4. What the Codebase Already Does Well

### Paper-facing prompt surfaces

The repository has a dedicated `paper_serializers.py` module instead of leaving
paper prompt formats scattered across runtime code. This is a strong design
choice because it creates a single contract for vanilla prompts, compressed
prompts, Stage 2 uncompressed prompts, coordinate remapping, and range parsing.

### Compression pipeline shape

The main encoder follows the same broad SheetCompressor sequence:

```text
Workbook loading
-> vanilla token baseline
-> structural-anchor extraction
-> optional homogeneous row/column pruning
-> inverted index translation
-> data-format aggregation
-> compact coordinate mapping
-> compressed prompt metrics
```

This does not prove paper equivalence, but it is much closer than a simple
spreadsheet-to-text serializer.

### Reproducibility posture

The README, `EVALUATION.md`, and `FINAL_POLISH_SPEC.md` are explicit that the
repo does not reproduce paper metrics unless the original datasets, splits,
model procedures, and baselines are available. That distinction is important:
the project can support SpreadsheetLLM-style experiments without overstating
paper-comparable results.

### Test coverage breadth

The repository includes tests for serializers, coordinate maps, compression,
Chain-of-Spreadsheet behavior, LLM backends, evaluations, fine-tuning data,
baselines, packaging metadata, tokenizer behavior, Streamlit helpers, and
end-to-end flows.

## 5. Remaining Gaps

### Gap A: Structural-anchor equivalence

The local anchor extraction is documented as Appendix C-inspired, not as a full
paper-equivalent implementation. The code considers values, styles, merged
regions, boundary candidates, filtering, and overlap handling, but the remaining
risk is methodological equivalence: there is no bundled paper-original fixture
set proving the same table skeletons as the paper.

Recommended next work:

1. Add fixture workbooks for title rows, notes, sparse sheets, merged headers,
   side-by-side tables, date/year headers, and overlapping candidates.
2. Record expected anchors and expected compressed prompts for each fixture.
3. Separate "paper-strict" and "pragmatic default" results in test names and
   documentation.

### Gap B: Benchmark parity

The codebase has evaluation scaffolding, but it does not bundle the paper's
table-detection or QA datasets. `EVALUATION.md` correctly separates synthetic,
reconstructed, and paper-original claims.

Recommended next work:

1. Keep synthetic CI tests local and deterministic.
2. Add manifest validation for reconstructed and paper-original benchmark runs.
3. Fail or warn when users attempt paper-comparable claims without required
   dataset, split, coordinate, prompt, backend, baseline, and metric metadata.

### Gap C: Binder baseline

TaPEx has an optional wrapper, but Binder is intentionally represented as
unavailable. This is acceptable if the project keeps the skip reason explicit,
but it remains a paper-comparison gap.

Recommended next work:

1. Keep Binder marked unavailable until an actual implementation exists.
2. If implemented, isolate the adapter behind the same baseline interface as
   TaPEx.
3. Require versioned baseline metadata in evaluation outputs.

### Gap D: Formula reasoning

The codebase can preserve cell values and prompt ranges, but it does not expose
a full formula dependency graph. For spreadsheet reasoning, formula precedents,
dependents, cross-sheet references, cached values, and errors can matter.

Recommended graph object:

```json
{
  "cell": "Sheet1!D10",
  "formula": "=SUM(D2:D9)",
  "cached_value": "12800",
  "references": ["Sheet1!D2:D9"],
  "dependents": [],
  "role": "total",
  "semantic_label": "Quarterly revenue total"
}
```

### Gap E: Multi-table robustness

The repository has table detection and Chain-of-Spreadsheet support, but robust
multi-table behavior depends on fixture coverage and benchmark evidence. Messy
spreadsheets can include notes above tables, blank separators, nested headers,
side-by-side regions, hidden rows, merged labels, and partial totals.

Recommended next work:

1. Add a synthetic multi-table benchmark manifest.
2. Include exact boundary targets and compact-coordinate targets.
3. Test both backend-driven table identification and fallback behavior.

## 6. Priority Roadmap

### Phase 1: Lock the verified paper-facing contract

1. Keep `paper_serializers.py` as the canonical prompt surface.
2. Keep strict and pragmatic compression modes visibly separate.
3. Continue asserting that saved `coord_map` artifacts round-trip after JSON
   reload.
4. Keep packaging metadata aligned with top-level modules in `pyproject.toml`.

### Phase 2: Expand methodological fixtures

1. Add structural-anchor fixtures for irregular layouts.
2. Add compressed-prompt golden outputs for number formats and semantic labels.
3. Add Stage 1 and Stage 2 Chain-of-Spreadsheet fixtures with original and
   compact ranges.
4. Add formula-heavy workbooks for future dependency-graph work.

### Phase 3: Strengthen evaluation claims

1. Make evaluation manifest validation stricter.
2. Emit machine-readable claim levels in every result.
3. Keep paper-original, reconstructed, and synthetic results separate.
4. Require baseline skip reasons instead of silent omissions.

### Phase 4: Add missing research capabilities where justified

1. Implement a formula dependency graph if downstream QA needs it.
2. Implement Binder only if a real adapter and reproducible baseline procedure
   are available.
3. Add richer visual-style analysis only where tests show it improves
   structure detection or QA, because full style emission can hurt token
   efficiency.

## 7. Bottom Line

The current repository is not merely a basic spreadsheet serializer. It already
implements many SpreadsheetLLM-inspired and paper-facing components:
paper-style serializers, structural-anchor heuristics, inverted-index
translation, data-format aggregation, coordinate remapping, CoS-style prompting,
evaluation scaffolds, fine-tuning data preparation, packaging metadata, and
baseline adapters.

The main remaining gap is not basic implementation, but proof and parity:
structural-anchor equivalence, paper-original datasets, exact paper procedures,
and unavailable Binder support. The repo is best described as a
methodologically careful, paper-inspired implementation with several
paper-aligned prompt and compression surfaces, not as a full reproduction of the
paper's reported metrics.
