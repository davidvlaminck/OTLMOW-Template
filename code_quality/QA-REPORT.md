# QA Report — otlmow_template

> GENERATED FILE — produced by `code_quality/scripts/qa_report.py`.
> Do not edit by hand; edit the thresholds in the script instead.

Generated: 2026-10-21 · commit `ef7153c`

## Thresholds (gate)

| Metric | Threshold | Current | Status |
|---|---|---|---|
| ruff issues | 0 | 62 | 🔴 |
| files needing `ruff format` | 0 | 4 | 🔴 |
| mypy errors | 0 | 0 | 🟢 |
| test coverage (lines) | ≥ 80.0 % | 98.7 % | 🟢 |
| test coverage (branches) | ≥ 75.0 % | 91.3 % | 🟢 |
| max CRAP score | ≤ 6.0 | 5.0 | 🟢 |
| max cyclomatic complexity | ≤ 6 | 5 | 🟢 |
| bandit HIGH issues | 0 | 0 | 🟢 |
| dead code items | 0 | 0 | 🟢 |
| god classes | 0 | 0 | 🟢 |

## Lead refactor target (highest CRAP)

`otlmow_template/excel_choice_list.py:20` → `add_choice_list_validation` — CRAP **5.0** — complexity **5** — coverage **100.0 %**

Verify with `code_quality/README.md` §7 before touching it: the behaviour must be pinned by a golden-file test first.

## Change hot spots (churn x complexity)

| Score | Commits | Lines changed | Max complexity | File |
|---:|---:|---:|---:|---|
| 5556 | 9 | 1389 | 4 | `otlmow_template/SubsetTemplateCreator.py` |

## ruff findings per rule

| Rule | Count |
|---|---:|
| `PLR0913` | 16 |
| `PLR0917` | 16 |
| `PLC0415` | 7 |
| `I001` | 3 |
| `N806` | 3 |
| `TRY003` | 3 |
| `BLE001` | 2 |
| `C901` | 2 |
| `PLR2004` | 2 |
| `ARG003` | 1 |
| `F401` | 1 |
| `G004` | 1 |
| `LOG015` | 1 |
| `N818` | 1 |
| `N999` | 1 |

## History

| Date | Coverage % | Branch % | Ruff | MyPy | Max CRAP | Max complexity | Result |
|---|---:|---:|---:|---:|---:|---:|---|
| 2026-10-20 | 95.9 | 86.4 | 181 | 18 | 35.1 | 35 | red |
| 2026-10-20 | 98.7 | 90.4 | 171 | 18 | 5.0 | 5 | red |
| 2026-10-21 | 99.0 | 0.0 | 0 | 0 | 5.0 | 5 | red |
| 2026-10-21 | 99.0 | 91.3 | 0 | 0 | 5.0 | 5 | green |
| 2026-10-21 | 99.0 | 91.3 | 0 | 0 | 5.0 | 5 | red |
| 2026-10-21 | 99.0 | 91.3 | 63 | 18 | 5.0 | 5 | red |
| 2026-10-21 | 98.7 | 91.3 | 62 | 0 | 5.0 | 5 | red |
| 2026-10-21 | 98.7 | 91.3 | 62 | 0 | 5.0 | 5 | red |
| 2026-10-21 | 98.7 | 91.3 | 62 | 0 | 5.0 | 5 | red |
| 2026-10-21 | 98.7 | 91.3 | 62 | 0 | 5.0 | 5 | red |
