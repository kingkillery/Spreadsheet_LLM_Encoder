# Final Polish PRD / Implementation Spec

## Status

Draft for final paper-alignment, bug-fix, and reproducibility work.

## Product Goal

Move this repository from a paper-inspired SpreadsheetLLM implementation to a reproducible, installable, and methodologically defensible implementation whose behavior is explicit about what is paper-faithful, what is an approximation, and what remains unavailable because the original datasets or baselines are not bundled.

The final polished version should let a user:

- Install the package from a clean checkout.
- Encode workbooks with stable paper-style prompt surfaces.
- Save and reload encoded artifacts without coordinate drift.
- Prepare fine-tuning records whose prompts and completions use the same coordinate system.
- Run table-detection and QA evaluations with recorded settings.
- Understand which parts match the paper and which are documented approximations.

## Non-Goals

- Do not claim reproduction of the paper's reported metrics unless the original datasets, splits, model procedures, and baselines are actually reconstructed.
- Do not bundle proprietary or unavailable benchmark data.
- Do not hide approximation paths behind paper-faithful naming.
- Do not require paid LLM APIs for basic tests or local synthetic benchmarks.

## Current Baseline

Already implemented:

- `k=4` default for structural-anchor extraction.
- Paper-style prompt serializers.
- Tokenizer-based compression counting with fallback.
- Coordinate remapping in compressed prompts.
- Semantic substitution in compressed prompts.
- LLM backend adapter.
- Chain-of-Spreadsheet Stage 2 path that can read uncompressed ranges from the original workbook.
- Real HuggingFace TaPEx baseline wrapper.
- JSON-reloaded coordinate-map normalization.
- Fine-tuning completion ranges remapped into prompt coordinates.
- Expanded packaging metadata for top-level modules and optional dependencies.

Known remaining gaps:

- Structural-anchor extraction is still simpler than Appendix C.
- Data-format aggregation does not yet choose between Excel number-format strings and semantic labels in a paper-faithful way.
- CoS Stage 1 table selection and large-table synthesis are still approximations.
- Evaluation lacks paper dataset splits, full reproducibility metadata, Binder support, and fine-tuning run records.
- Documentation still overstates some paper-faithful behavior in places.

## Users

- Researchers comparing spreadsheet encoding methods.
- Developers integrating spreadsheet compression into LLM workflows.
- Maintainers validating future paper-alignment changes.
- Benchmark authors building synthetic or reconstructed table-detection / QA datasets.

## Success Criteria

- Clean install works with `pip install -e .`.
- Core test suite passes without real network, OpenAI, HuggingFace model downloads, or paper datasets.
- Every evaluation result includes dataset, split, model/backend, encoding settings, coordinate mode, prompt serializer, and metric definition.
- Fine-tuning JSONL records never mix compact-coordinate prompts with original-coordinate completions.
- Saved encodings can be reloaded and used for remap/unremap without manual key conversion.
- README and spec language distinguish paper-faithful paths from approximations.
- All TODO baseline paths either run, skip with machine-readable reasons, or are explicitly documented as unavailable.

## Phase 1: Reproducibility Foundation

### 1.1 Packaging Metadata

Problem:
Top-level runtime modules and optional dependencies must be reflected in package metadata.

Implementation:
- Keep `pyproject.toml` `py-modules` aligned with all importable top-level modules.
- Keep optional dependency groups for tokenizer, OpenAI backend, fine-tuning, QLoRA, and baselines.
- Add a lightweight package metadata regression test if practical.

Acceptance:
- `python -m pip install -e . --dry-run` succeeds.
- Import smoke test succeeds for `paper_serializers`, `chain_of_spreadsheet`, `llm_backend`, `tokenizer`, `evaluation`, and `baselines`.

### 1.2 Artifact Reload Contract

Problem:
Saved JSON artifacts must behave the same as in-memory encodings.

Implementation:
- Treat `coord_map` JSON keys as untrusted persisted data.
- Normalize coordinate maps before all remap/unremap operations.
- Add a public helper for artifact loading if multiple scripts read saved encodings directly.

Acceptance:
- A saved encoding can be reloaded with `json.load()` and passed to prompt serializers.
- Compact coordinates in saved artifacts remap/unremap identically to in-memory coordinates.

### 1.3 Fine-Tuning Coordinate Parity

Problem:
Fine-tuning prompts use compact coordinates, so completions must also use compact coordinates.

Implementation:
- Remap ground-truth table boxes into prompt coordinates during JSONL generation.
- Skip or flag ground-truth boxes not represented in the compressed coordinate map.
- Include enough metadata in future records to reconstruct original ranges if needed.

Acceptance:
- Tests assert prompt ranges and completion ranges match.
- Unmapped annotations are logged and do not silently produce inconsistent training targets.

## Phase 2: Method Fidelity

### 2.1 Structural-Anchor Extraction

Problem:
Current anchor extraction is directionally aligned but does not fully implement the paper's Appendix C heuristics.

Implementation:
- Expand boundary proposal features to include neighboring discrepancies across values, merged cells, borders, fill colors, text, font, and style where available.
- Add candidate filtering for edge sparsity, internal sparsity, row/column text-number proportions, size, and header-likeness.
- Improve overlap resolution using relative-position, title/header/date/year cues, and table relationship patterns.
- Add a strict paper-mode option that disables post-anchor homogeneous row/column pruning by default.

Acceptance:
- Fixture workbooks cover title rows, notes, multi-table sheets, sparse tables, merged headers, date/year headers, and overlapping candidates.
- Strict paper mode preserves skeleton rows/columns within the selected anchor neighborhood.
- README clearly distinguishes strict paper mode from pragmatic default mode if both remain.

### 2.2 Data-Format-Aware Aggregation

Problem:
The compressed prompt currently favors semantic labels, while the paper also uses informative Excel number-format strings.

Implementation:
- Define a deterministic selection rule:
  - Use semantic labels for paper-listed semantic classes when no informative NFS is needed.
  - Use actual Excel number-format strings when the format string carries meaningful semantics.
  - Keep fallback labels for generic numeric/date/email/scientific regions.
- Normalize vocabulary to paper-observed labels where possible.
- Add tests for integer, float, date, year, time, percentage, currency, scientific notation, email, and custom NFS strings.

Acceptance:
- Prompt examples include both semantic labels and NFS strings where expected.
- Vocabulary drift is documented or removed.
- Tests prove literal cell values are suppressed only when the replacement tuple is emitted.

### 2.3 Chain-of-Spreadsheet Fidelity

Problem:
CoS has the right broad shape, but table selection and large-table synthesis remain simplified.

Implementation:
- Make Stage 1 table identification consistently use the compressed paper-style prompt when a backend is configured.
- Make keyword sheet selection an explicit fallback, not the default paper-faithful path.
- Add structured parsing for multiple predicted ranges.
- Add a final synthesis call for chunked large-table QA instead of returning joined per-chunk answers.
- Record whether Stage 2 used original-workbook uncompressed pairs or fallback compressed JSON.

Acceptance:
- Tests cover backend-driven sheet/table selection, fallback paths, multiple range extraction, compact-to-original range conversion, and final synthesis.
- Logs or result metadata make fallback behavior visible.
- Stage 2 paper-faithful path requires `workbook_path`, `sheet_name`, and original table range.

## Phase 3: Evaluation Parity

### 3.1 Evaluation Metadata Contract

Problem:
Evaluation outputs are not comparable unless settings are recorded.

Implementation:
- Define a result schema containing:
  - dataset name and version
  - split name
  - spreadsheet count
  - table / QA item count
  - encoder settings
  - prompt serializer
  - coordinate mode
  - model/backend
  - metric definition
  - baseline name and version
  - skip reasons
- Write records to `skill_runs/` or a dedicated `runs/` directory with stable filenames.

Acceptance:
- Evaluation scripts produce machine-readable metadata.
- Missing parity metadata causes a warning or failed validation.

### 3.2 Table Detection Benchmark

Problem:
The repo has EoB-style evaluation but not the paper's full dataset/split reconstruction.

Implementation:
- Document the required paper benchmark shape: spreadsheet count, table count, token-size partitions, and EoB-0 exact boundary matching.
- Add synthetic fixture partitions for CI.
- Add a benchmark manifest format for real or reconstructed datasets.
- Add validation that predictions and ground truth use the same coordinate system.

Acceptance:
- Synthetic benchmark runs end to end without external APIs.
- Real benchmark runs require a manifest and fail clearly if unavailable.
- Reports separate synthetic, reconstructed, and paper-original claims.

### 3.3 Spreadsheet QA Benchmark

Problem:
QA evaluation is usable but not yet comparable to the paper.

Implementation:
- Add a QA manifest format with workbook path, sheet, question, answer, answer type, and table range if known.
- Record whether answers are cell addresses, formulas, literal values, or free text.
- Add exact-match normalization rules by answer type.
- Keep TaPEx optional and lazy-loaded.
- Keep Binder marked unavailable until a real implementation exists.

Acceptance:
- QA scripts produce accuracy plus per-item records.
- Baselines are either runnable or skipped with explicit reason codes.
- Paper-comparison reports require compatible answer-type handling.

### 3.4 Fine-Tuning Execution Records

Problem:
The repo prepares JSONL but does not yet make fine-tuning runs reproducible.

Implementation:
- Add a `finetune_manifest.json` schema with dataset, split, base model, adapter settings, prompt template hash, coordinate mode, and command.
- Have `prepare_finetuning_data.py` optionally emit metadata next to JSONL.
- Document local LoRA and hosted job execution paths as reproducibility recipes, not guaranteed paper reproduction.

Acceptance:
- Fine-tuning JSONL has adjacent metadata.
- Evaluation can reject fine-tuned model outputs if metadata does not match evaluation settings.

## Phase 4: Documentation and UX Polish

### 4.1 README Corrections

Problem:
Some README language implies stronger paper fidelity than the implementation currently supports.

Implementation:
- Replace overbroad claims with precise status language.
- Add a "Paper Fidelity Matrix" table.
- Move implementation gaps into a visible section with links to this spec.

Acceptance:
- README does not claim full reproduction of paper metrics.
- Users can identify strict paper paths, pragmatic defaults, and unavailable baselines.

### 4.2 CLI and Script Ergonomics

Problem:
Users need clearer modes and failure messages.

Implementation:
- Add CLI flags for strict paper mode where behavior differs.
- Add explicit warnings for fallback compressed Stage 2, missing backend, missing original workbook, unavailable Binder, and non-tokenizer fallback metrics.
- Ensure Windows path examples use safe quoting in TOML docs and config snippets.

Acceptance:
- Common failure paths produce actionable messages.
- CLI help names paper-faithful vs fallback behavior.

### 4.3 Test Stability on Windows

Problem:
Full pytest currently runs tests successfully but can exit nonzero on Windows due to temp symlink cleanup.

Implementation:
- Identify which test or fixture creates symlink-style `current` temp paths.
- Avoid symlink cleanup issues by configuring pytest temp behavior or using plain directories for affected tests.
- Document any Python / pytest version constraints if needed.

Acceptance:
- `python -m pytest` exits zero on Windows after all tests pass.
- No generated `.pytest_tmp` or temp artifacts remain in the repo after normal test runs.

## Implementation Priority

1. Fix Windows pytest cleanup exit.
2. Add import/package smoke tests.
3. Add evaluation metadata schema and validator.
4. Correct README paper-fidelity claims.
5. Implement data-format NFS-vs-label selection.
6. Improve CoS Stage 1 table selection and chunk synthesis.
7. Expand structural-anchor fixtures and heuristics.
8. Add benchmark manifests and synthetic CI benchmark.
9. Add fine-tuning metadata sidecar.
10. Decide whether Binder remains documented unavailable or gets a real implementation.

## Release Gate Checklist

- `python -m pytest` exits zero.
- `python -m pip install -e . --dry-run` succeeds.
- Import smoke tests pass after editable install.
- README has a paper-fidelity matrix.
- Saved artifact reload test passes.
- Fine-tuning coordinate parity test passes.
- CoS Stage 2 original-workbook path test passes.
- Evaluation metadata validator passes on generated table-detection and QA records.
- Optional dependency paths skip cleanly when dependencies are absent.
- No result claims compare against the paper unless dataset/model/baseline parity is documented.

## Suggested Issue Breakdown

- Issue 1: Stabilize pytest cleanup on Windows.
- Issue 2: Add package import smoke tests.
- Issue 3: Add evaluation metadata schema and validator.
- Issue 4: Update README with paper-fidelity matrix.
- Issue 5: Implement NFS-vs-semantic-label aggregation policy.
- Issue 6: Add CoS final synthesis for split tables.
- Issue 7: Make backend-driven CoS table selection the paper-faithful path.
- Issue 8: Expand structural-anchor Appendix C fixtures.
- Issue 9: Add benchmark manifests and synthetic parity runs.
- Issue 10: Add fine-tuning metadata sidecars.
- Issue 11: Document Binder as unavailable or implement a real adapter.

