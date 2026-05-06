# skill_runs

Structured iteration records for the `spreadsheet-llm-fidelity` skill.

## Why

The improvement loop is observation-driven, not memory-driven. Every iteration that touches paper-aligned modules writes one JSON record here so we can:

- replay a specific gap-remediation iteration months later
- compare metrics between iterations (compression ratio, tests added/passing)
- audit what the architect / security / code-review validators verified vs. left UNKNOWN
- pick up the loop from the last `next_gap` instead of re-deriving it

## Filename convention

`<iso8601>-<short-slug>.json`, e.g. `2026-05-06T10-30-00Z-format-substitution.json`.

ISO timestamp first so a directory listing is chronological. Slug should be 2–4 hyphenated words tied to the goal.

## Schema

The canonical schema is documented in `.claude/skills/spreadsheet-llm-fidelity/SKILL.md` under "Output Format". Required fields:

- `iteration` (int) — monotonic counter; bump from the previous run record's value.
- `timestamp` (RFC3339)
- `goal` (one line)
- `verified` / `inferred` / `unknown` (arrays of strings)
- `amendment` — files changed, tests added, lines delta
- `evaluation` — test totals + validator verdicts
- `verdict` ∈ `{promote, rollback, deferred}`
- `next_gap` (one line)

Anything else can be added but those fields must always be present.

## Reading the log

```bash
ls skill_runs/*.json | sort
python -c "import json; from pathlib import Path; \
  recs = sorted(Path('skill_runs').glob('*.json')); \
  [print(json.loads(r.read_text())['iteration'], '-', json.loads(r.read_text())['goal']) for r in recs]"
```

## When NOT to write a record

- Pure cosmetic edits (typos, formatting, README polish) that don't change behavior.
- Dependency bumps that don't touch paper-aligned modules.
- Test-only changes that fix a flaky test without changing the system under test.

For everything else — write the record.
