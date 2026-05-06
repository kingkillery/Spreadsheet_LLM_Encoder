---
name: spreadsheet-llm-fidelity
description: Project-scoped skill for keeping this repo faithful to the SpreadsheetLLM paper (arXiv:2407.09025). Triggers on requests to "check fidelity to the paper", "remediate gap analysis", "verify paper alignment", "rerun the eval pipeline", or to add a new gap-remediation iteration.
type: project
version: 1
---

# spreadsheet-llm-fidelity

## Purpose

Keep this codebase faithful to the SpreadsheetLLM paper (arXiv:2407.09025) and turn each gap-remediation iteration into a tracked, evaluable amendment instead of a one-off patch. Every run produces a structured record under `skill_runs/` so future iterations can replay or compare.

## Trigger Conditions

Activate when the user:

- Provides a paper-fidelity gap analysis and asks to "plan and implement", "remediate", or "fill gaps".
- Asks to "rerun the eval pipeline" or "evaluate against the paper".
- Asks to "check what's still missing vs. the paper".
- Asks to add another iteration to the improvement loop.
- Says "loop", "fill gaps as we go", "iterate on remediation".

Do NOT activate for unrelated repo work (e.g., Streamlit cosmetics, dependency bumps that don't touch the paper-aligned modules).

## Input Schema

```yaml
inputs:
  goal: string              # required — one-line objective for this iteration
  evidence:                 # required — what makes us think there is a gap
    - failure_logs: optional
    - paper_section: optional   # e.g. "Section 3.3.3 data-format aggregation"
    - test_failures: optional
  scope:                    # required — what files/modules are in scope
    files: list[string]
  acceptance_criteria:      # required — how we know the iteration is done
    - tests: list[string]
    - manual_check: optional
  out_of_scope: optional list[string]
```

## Tool Calls

Order matters here.

1. **Observe** — read the gap analysis or failing tests verbatim. Do NOT skim. Capture VERIFIED / INFERRED / UNKNOWN claims.
2. **Plan** — write the smallest sufficient amendment. Prefer per-iteration tasks via `TaskCreate`.
3. **Dispatch** — for multi-file or independent edits, use parallel `Agent` calls (executor + writer + test-engineer). Provide each a self-contained brief that names the new module's contract, not just behavior.
4. **Run tests** — `python -m pytest` from repo root. Treat the pre-existing Windows `OSError WinError 448` pytest temp-cleanup as noise, not a failure (see Examples).
5. **Validate** — for changes touching paper-aligned modules, run the architect / security-reviewer / code-reviewer triad in parallel via `Agent`. Address every CRITICAL and HIGH before committing.
6. **Log run** — append a JSON record to `skill_runs/` (see Output Format).

The eval scripts (`run_llm_evaluation.py`, `run_qa_evaluation.py`) are part of step 4 ONLY when the iteration touches encoder, CoS, or evaluator wiring — they're slow without a real LLM backend, so default to `--backend echo` for smoke tests.

## Output Format

Every iteration emits one record at `skill_runs/<iso8601>-<short-slug>.json`:

```json
{
  "iteration": 7,
  "timestamp": "2026-05-06T00:00:00Z",
  "goal": "Substitute compressible format regions with paper labels.",
  "verified": ["paper section 3.3.3 calls for label substitution, not annotation"],
  "inferred": ["our previous JSON 'formats' field was an annotation, not substitution"],
  "unknown": ["whether the paper's prompt joins tuples with comma, no separator, or newline"],
  "amendment": {
    "files_changed": ["paper_serializers.py", "Spreadsheet_LLM_Encoder.py"],
    "tests_added": ["test_paper_serializers.py::test_to_paper_compressed_prompt_substitutes_label"],
    "lines_delta": "+312/-48"
  },
  "evaluation": {
    "tests_total": 111,
    "tests_passed": 111,
    "validators": {
      "architect": "approved",
      "security": "approved-after-fixes",
      "code_review": "approved-after-fixes"
    }
  },
  "verdict": "promote",
  "next_gap": "Real TaPEx baseline integration"
}
```

Always include `verdict` ∈ `{promote, rollback, deferred}` and at least one `next_gap` line so the loop has somewhere to go.

## Amendment Loop

```
observe -> inspect -> amend -> evaluate -> commit-or-rollback -> log -> next_gap
```

Hard rules:

1. Separate VERIFIED, INFERRED, UNKNOWN. Don't promote a finding to VERIFIED without a log/test/citation.
2. Make the smallest sufficient change.
3. Try to falsify the leading hypothesis before patching.
4. Run the validator triad before declaring an iteration done if the change touches paper-aligned modules.
5. Every iteration writes ONE `skill_runs/` record. No record → no promotion.

## Examples

### Trigger: "rerun the eval pipeline"

Smallest sufficient action: synthesize a small benchmark, run encoder + LLM eval with `EchoBackend`, capture metrics in a run record.

```bash
python scripts/synthesize_benchmark.py datasets/synth_smoke/ --n 5 --seed 1
python scripts/synthesize_qa.py datasets/synth_smoke/ --per-table 3 --seed 1 --out-suffix ""
python run_llm_evaluation.py datasets/synth_smoke/ --backend echo --k 4
```

Then write `skill_runs/<ts>-eval-smoke.json` with the captured metrics.

### Trigger: "fill gap N from the gap analysis"

1. Quote the gap analysis line in `verified[]` of the run record.
2. Plan the smallest patch — usually a single module + tests.
3. Dispatch executors in parallel only when files are non-overlapping.
4. Run validators in parallel when paper-aligned modules are touched.
5. Log + name `next_gap`.

### Anti-trigger: "fix this typo"

Don't activate. This skill is for paper-fidelity iterations, not micro-edits.

## Known Pitfalls (capture as you go)

- The Windows pytest temp-symlink cleanup raises `OSError [WinError 448]` after a successful run. Treat as noise.
- `tiktoken` is optional. Without it, compression metrics use a `len // 4` fallback that is NOT comparable to paper numbers.
- `categorize_number_format` historically had a duplicate caller passing the wrong arg type (a string instead of a Cell). Removed in iteration 5; if you see it return again, it's dead code resurfacing.
- Reversed-endpoint ranges (`C3:A1`) used to silently normalize. They now log a warning. Check for that warning when an LLM returns malformed ranges.
