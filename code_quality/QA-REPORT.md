# QA Report — otlmow_template

> GENERATED FILE — produced by `code_quality/scripts/qa_report.py`.
> Do not edit by hand; edit the thresholds in the script instead.

Generated: 2026-10-20 · commit `543d681`

## Thresholds (gate)

| Metric | Threshold | Current | Status |
|---|---|---|---|
| ruff issues | 0 | 181 | 🔴 |
| files needing `ruff format` | 0 | 0 | 🟢 |
| mypy errors | 0 | 18 | 🔴 |
| test coverage (lines) | ≥ 80.0 % | 95.9 % | 🟢 |
| test coverage (branches) | ≥ 75.0 % | 86.4 % | 🟢 |
| max CRAP score | ≤ 10.0 | 35.1 | 🔴 |
| max cyclomatic complexity | ≤ 15 | 35 | 🔴 |
| bandit HIGH issues | 0 | 0 | 🟢 |
| dead code items | 0 | 0 | 🟢 |

## Lead refactor target (highest CRAP)

`otlmow_template/SubsetTemplateCreator.py:491` → `alter_excel_sheet` — CRAP **35.1** — complexity **35** — coverage **95.2 %**

Verify with `code_quality/README.md` §7 before touching it: the behaviour must be pinned by a golden-file test first.

## Change hot spots (churn x complexity)

| Score | Commits | Lines changed | Max complexity | File |
|---:|---:|---:|---:|---|
| 3780 | 7 | 108 | 35 | `otlmow_template/SubsetTemplateCreator.py` |

## ruff findings per rule

| Rule | Count |
|---|---:|
| `I001` | 29 |
| `Q000` | 29 |
| `UP009` | 21 |
| `PLR0913` | 16 |
| `PLR0917` | 16 |
| `UP006` | 13 |
| `PTH123` | 7 |
| `C901` | 5 |
| `RUF013` | 5 |
| `TRY003` | 5 |
| `F401` | 3 |
| `PLR0912` | 3 |
| `UP035` | 3 |
| `E401` | 2 |
| `E712` | 2 |

## History

| Date | Coverage % | Branch % | Ruff | MyPy | Max CRAP | Max complexity | Result |
|---|---:|---:|---:|---:|---:|---:|---|
| 2026-10-20 | 95.9 | 86.4 | 181 | 18 | 35.1 | 35 | red |
