# QA Report — otlmow_template

> GENERATED FILE — produced by `code_quality/scripts/qa_report.py`.
> Do not edit by hand; edit the thresholds in the script instead.

Generated: 2026-10-20 · commit `14260fe`

## Thresholds (gate)

| Metric | Threshold | Current | Status |
|---|---|---|---|
| ruff issues | 0 | 171 | 🔴 |
| files needing `ruff format` | 0 | 31 | 🔴 |
| mypy errors | 0 | 18 | 🔴 |
| test coverage (lines) | ≥ 80.0 % | 98.7 % | 🟢 |
| test coverage (branches) | ≥ 75.0 % | 90.4 % | 🟢 |
| max CRAP score | ≤ 6.0 | 5.0 | 🟢 |
| max cyclomatic complexity | ≤ 6 | 5 | 🟢 |
| bandit HIGH issues | 0 | 0 | 🟢 |
| dead code items | 0 | 0 | 🟢 |

## Enforced gates

These gates block a merge — all met. The gates that are not listed are reported only; the ratchet is enforced by every run.

| Gate | Threshold | Current | Status |
|---|---|---|---|
| `coverage` | 80.0 | 98.7 | 🟢 |
| `coverage-branch` | 75.0 | 90.4 | 🟢 |
| `crap` | 6.0 | 5.0 | 🟢 |
| `complexity` | 6 | 5 | 🟢 |
| `bandit` | 0 | 0 | 🟢 |
| `dead-code` | 0 | 0 | 🟢 |

## Lead refactor target (highest CRAP)

`otlmow_template/SubsetTemplateCreator.py:818` → `add_choice_list_validation` — CRAP **5.0** — complexity **5** — coverage **100.0 %**

Verify with `code_quality/README.md` §7 before touching it: the behaviour must be pinned by a golden-file test first.

## Change hot spots (churn x complexity)

| Score | Commits | Lines changed | Max complexity | File |
|---:|---:|---:|---:|---|
| 4230 | 8 | 846 | 5 | `otlmow_template/SubsetTemplateCreator.py` |

## ruff findings per rule

| Rule | Count |
|---|---:|
| `I001` | 29 |
| `Q000` | 27 |
| `UP009` | 21 |
| `PLR0913` | 16 |
| `PLR0917` | 16 |
| `UP006` | 13 |
| `PTH123` | 7 |
| `RUF013` | 5 |
| `TRY003` | 5 |
| `F401` | 3 |
| `UP035` | 3 |
| `C901` | 2 |
| `E401` | 2 |
| `E712` | 2 |
| `E721` | 2 |

## History

| Date | Coverage % | Branch % | Ruff | MyPy | Max CRAP | Max complexity | Result |
|---|---:|---:|---:|---:|---:|---:|---|
| 2026-10-20 | 95.9 | 86.4 | 181 | 18 | 35.1 | 35 | red |
| 2026-10-20 | 98.7 | 90.4 | 171 | 18 | 5.0 | 5 | red |
