# QA Report — otlmow_template

> GENERATED FILE — produced by `code_quality/scripts/qa_report.py`.
> Do not edit by hand; edit the thresholds in the script instead.

Generated: 2026-10-28 · commit `9cb193b`

## Thresholds (gate)

| Metric | Threshold | Current | Status |
|---|---|---|---|
| ruff issues | ≤ 0 | 0 | 🟢 |
| files needing `ruff format` | ≤ 0 | crashed | 🔴 |
| mypy errors | ≤ 0 | 103 | 🔴 |
| test coverage (lines, owned code) | ≥ 80.0 | 85.4 % | 🟢 |
| test coverage (branches, owned code) | ≥ 75.0 | 0.0 % | 🔴 |
| max CRAP score | ≤ 6.0 | 45.8 | 🔴 |
| max cyclomatic complexity | ≤ 6 | 35 | 🔴 |
| bandit HIGH issues | ≤ 0 | 0 | 🟢 |
| dead code items | ≤ 0 | 0 | 🟢 |
| pip-audit vulnerabilities | ≤ 0 | 0 | 🟢 |
| god classes | ≤ 0 | 1 | 🔴 |
| files radon cannot parse | ≤ 0 | 0 | 🟢 |
| manual conventions | ≤ 0 | 3 | 🔴 |
| import cycles/architecture | ≤ — | not enabled | ⚪ |

## Tools not available

The metrics of these tools are **not measured**: a missing tool reports 0 findings, which looks like a clean run. Install them before trusting the numbers above.

- `lint-imports`

## Coverage split — generated vs hand-written

The gate above measures the **hand-written** code: that is the code a refactor touches. The whole package — generated model included — covers 85.4 % of the lines and 0.0 % of the branches. The generated model lifts the headline number and would hide the gap in the code we own.

| Code | Statements | Line % | Branches | Branch % |
|---|---:|---:|---:|---:|
| generated (excluded) | 0 | 0.0 % | 0 | 0.0 % |
| hand-written | 967 | 85.4 % | 0 | 0.0 % |

## Lead refactor target (highest CRAP)

`otlmow_template/InfotemplateGenerator.py:152` → `_get_attributes_for_class` — CRAP **45.8** — complexity **15** — coverage **48.5 %**

Verify with `code_quality/README.md` §7 before touching it: the behaviour must be pinned by a golden-file test first.

## Change hot spots (churn x complexity)

| Score | Commits | Lines changed | Max complexity | File |
|---:|---:|---:|---:|---|
| 25236 | 116 | 6309 | 4 | `otlmow_template/SubsetTemplateCreator.py` |
| 22050 | 16 | 1050 | 21 | `otlmow_template/InfotemplateGenerator.py` |
| 20650 | 1 | 590 | 35 | `otlmow_template/AanlevertemplateGenerator.py` |
| 1600 | 2 | 100 | 16 | `otlmow_template/TemplateGeneratorBase.py` |
| 0 | 11 | 590 | 0 | `otlmow_template/CsvTemplateCreator.py` |

## History

| Date | Coverage % | Branch % | Ruff | MyPy | Max CRAP | Max complexity | Result |
|---|---:|---:|---:|---:|---:|---:|---|
| 2026-10-28 | 85.4 | 0.0 | 0 | 103 | 45.8 | 35 | red |
