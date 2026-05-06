# Top 3 High-Leverage Tasks

## 1. Build a Paper-Parity Fixture Suite

Create 10-20 synthetic `.xlsx` fixtures that cover messy spreadsheet cases:
merged headers, side-by-side tables, notes above tables, sparse sheets,
date/year headers, totals, formulas, and multi-table layouts.

For each fixture, add expected outputs for:

- structural anchors
- compressed prompts
- detected ranges
- coordinate remapping

Why this is highest leverage: it turns vague "paper-inspired" claims into
measurable behavior and protects future changes from regressions.

## 2. Implement Formula Dependency Extraction

Add a lightweight formula graph that captures:

- formula cell
- formula string
- cached value
- referenced ranges
- cross-sheet references
- formula errors
- repeated-formula summaries

Why this matters: formulas are one of the biggest remaining
spreadsheet-specific reasoning gaps, and this work is relatively isolated
compared with full benchmark reproduction.

## 3. Tighten Evaluation Claim Validation

Make evaluation records explicitly label results as:

- `synthetic`
- `reconstructed`
- `paper-original`

Fail or warn if a run attempts to claim paper-comparable results without:

- dataset metadata
- split metadata
- model/backend metadata
- prompt serializer
- coordinate mode
- baseline status
- metric definition

Why this is leveraged: it prevents unsupported research claims and makes the
repo methodologically defensible without requiring the original paper datasets.
