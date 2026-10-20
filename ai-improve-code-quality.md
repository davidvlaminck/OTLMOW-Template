# AI-Improve-Code-Quality

Measure → QA → refactor → check → loop cycle for full Python repositories, with **CRAP score** as the lead driver for refactoring.

This document is written for AI agents (and humans) to apply consistently across any pure-Python repo.

---

## 1. Philosophy

- **Measure** code quality objectively (lint, types, tests, coverage, complexity, security).
- **Publish** a QA report with current numbers vs thresholds.
- **Refactor** the worst offender first — always the item with the highest CRAP score.
- **Re-measure** and update the QA report.
- **Loop** until every metric is green. Do not merge until the gate is green.

CRAP score (Change Risk Anti-Patterns) is the lead driver because it fuses cyclomatic complexity and test coverage into one number: "How dangerous is this code to change?"

---

## 2. Tool mapping (Python equivalents)

| Stage | Tool | Purpose |
|---|---|---|
| Lint / anti-patterns | `ruff` | Style, bugs, imports, banned constructs (bans live in `PY-ANTI-PATTERNS.md`) |
| Format | `ruff format` | Deterministic formatting |
| Static types | `mypy` (strict) or `pyright` | Type errors |
| Tests | `pytest` + `pytest-cov` | Unit tests + coverage |
| Complexity | `radon` | Cyclomatic complexity (A–F + raw score) |
| CRAP score | Computed from `radon` + `coverage` | Lead refactor driver (see section 4) |
| Security | `bandit` | Common security issues |

Install (dev dependencies):

```bash
uv pip install ruff mypy pytest pytest-cov radon bandit
```

Or add to `[project.optional-dependencies] test` in `pyproject.toml`.

---

## 3. Configuration

### 3.1 pyproject.toml

Append (adjust `target-version`, `src`, and thresholds to the repo):

```toml
[tool.ruff]
line-length = 100
target-version = "py39"
src = ["<package>", "tests"]

[tool.ruff.lint]
select = ["E", "F", "I", "B", "C4", "UP", "PYI", "PT", "RUF", "S", "BLE", "TRY", "FLY", "A", "C", "P", "Q", "RET", "N", "W"]
ignore = ["E501"]  # line-length handled by formatter

[tool.ruff.lint.flake8-tidy-imports]
ban-relative-imports = "all"

[tool.ruff.format]
docstring-code-format = true

[tool.mypy]
python_version = "3.9"
strict = true
warn_return_any = true
warn_unused_ignores = true
disallow_any_generics = true
check_untyped_defs = true

[tool.pytest.ini_options]
testpaths = ["tests"]
addopts = "-v --cov=<package> --cov-report=term --cov-report=html:tests/htmlcov --cov-fail-under=80"
```

### 3.2 QA Report: `QA-REPORT.md`

Commit a markdown dashboard. Rewrite it every cycle. Template:

```markdown
# QA Report — <repo-name>
Baseline date: YYYY-MM-DD

## Thresholds (gate)
| Metric | Threshold | Current | Status |
|---|---|---|---|
| ruff issues | 0 | — | ⬜ |
| mypy errors | 0 | — | ⬜ |
| test coverage | ≥ 80 % | — | ⬜ |
| max CRAP score | ≤ 10 | — | ⬜ |
| bandit HIGH issues | 0 | — | ⬜ |

## Lead refactor target (highest CRAP)
File: <path>:<function> — CRAP: <n> — Complexity: <n> — Coverage: <n>%

## Last cycle notes
- Worst offender: <file> — <reason>

## History
| Cycle | Date | Coverage | Ruff | MyPy | Max CRAP | Result |
|---|---|---|---|---|---|---|
| 0 | YYYY-MM-DD | — | — | — | — | baseline |
```

### 3.3 Anti-patterns: `PY-ANTI-PATTERNS.md`

Document banned constructs mapped to ruff rules:

| Anti-pattern | Why | Ruff rule | Fix |
|---|---|---|---|
| Bare `except:` | swallows `KeyboardInterrupt` / `SystemExit` | `BLE001` | `except Exception:` |
| Mutable default argument (`def f(x=[])`) | shared mutable state | `B006` | `def f(x=None): ...` |
| `print()` in library code | breaks logging / captures | `T201` | `logging` |
| f-string in `logging.debug(f"...")` | eagerly evaluates string | `G004` | lazy `%`-formatting |
| `import *` | namespace pollution | `F401` / `F403` | explicit imports |
| `eval(` / `exec(` | injection risk | `S307` / `S302` | avoid |

---

## 4. CRAP score — the lead driver

### Formula

```
CRAP = mccabe² × (1 − coverage/100)³ + mccabe
```

- `mccabe` = cyclomatic complexity (from `radon cc`)
- `coverage` = line coverage % (from `pytest-cov` / `coverage json`)

### Interpretation

| Coverage | Effect on CRAP |
|---|---|
| 100 % | CRAP = mccabe (complexity alone matters) |
| 0 % | CRAP = mccabe² + mccabe (complexity explodes when untested) |

A simple but untested function can outrank a complex but fully-tested one.

### Threshold bands

| CRAP | Meaning | Action |
|---|---|---|
| ≤ 5 | Clean | Leave alone |
| 5–10 | Acceptable | Monitor |
| 10–30 | Risky | Refactor soon |
| > 30 | Dangerous | Refactor now |
| > 60 | Critical | Block merge until fixed |

### Computing CRAP in Python

There is no mainstream `crap` CLI for Python. Compute it by joining two JSON outputs:

1. `radon cc -j <package>` → per-function `complexity`
2. `coverage json` → per-file/per-line execution count

A small script (~10–20 lines) maps functions to covered lines and applies the formula. The agent should produce `CRAP_SCORES.md` or append a CRAP column to `QA-REPORT.md`.

Example output column:

```markdown
| Module | Function | Complexity | Coverage | CRAP | Status |
|---|---|---|---|---|---|
| otlmow_template/SubsetTemplateCreator.py | generate_objects_for_template | 14 | 62 % | 48 | ❌ Refactor |
```

---

## 5. The cycle (measure → QA → refactor → check → loop)

```
┌─────────────┐
│ 1. Measure  │────┐
└──────┬──────┘     │
       ▼           │
┌─────────────┐    │
│ 2. QA report│    │
│   update    │    │
└──────┬──────┘    │
       ▼           │
┌─────────────┐    │
│ 3. Refactor │    │
│   highest   │    │
│   CRAP      │    │
└──────┬──────┘    │
       ▼           │
┌─────────────┐    │
│ 4. Re-measure│───┘
│ thresholds OK?│
└──────┬──────┘
   yes / no
    ▼    ▼
 merge  loop
```

### Step-by-step commands

```bash
# 1. Measure
ruff check <package> tests
ruff format --check <package> tests
mypy <package> tests
python -m pytest --cov=<package> --cov-report=term --cov-report=html:tests/htmlcov
radon cc -a -n C <package> tests   # fail if worse than C
bandit -q -r <package>

# 2. Compute CRAP (join radon + coverage JSON)
# Update QA-REPORT.md with all numbers

# 3. Refactor the file/function with the highest CRAP score

# 4. Repeat from step 1 until QA-REPORT.md is fully green
```

**Loop rule:** Do not merge until every row in `QA-REPORT.md` is green. If one metric fails, refactor that file and re-measure.

---

## 6. CI enforcement (GitHub Actions)

Add a `quality` job that runs on every PR and fails if any threshold is breached. This keeps the loop honest.

Example `.github/workflows/quality.yml`:

```yaml
name: Quality gate
on:
  push:
    branches: [master]
  pull_request:
    branches: [master]

jobs:
  quality:
    name: Quality gate
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - uses: actions/checkout@v7
      - uses: astral-sh/setup-uv@v7
        with: { python-version: '3.12' }
      - run: uv venv --python 3.12
      - run: uv pip install -r pyproject.toml --extra test ruff mypy radon bandit
      - name: Lint
        run: uv run ruff check <package> tests
      - name: Format check
        run: uv run ruff format --check <package> tests
      - name: Type check
        run: uv run mypy <package> tests
      - name: Tests + coverage
        run: uv run python -m pytest --cov=<package> --cov-report=term --cov-fail-under=80
      - name: Complexity
        run: uv run radon cc -a -n C <package> tests
      - name: Security
        run: uv run bandit -q -r <package>
```

Any failure blocks the merge. Refactor until the gate is green.

---

## 7. How an agent applies this to a new repo

1. Read `pyproject.toml` and existing test/CI setup.
2. Add the `[tool.ruff]`, `[tool.mypy]`, `[tool.pytest.ini_options]` sections.
3. Run `ruff check`, `mypy`, `pytest --cov`, `radon cc`, `bandit` locally for a baseline.
4. Create `QA-REPORT.md` with baseline numbers; mark all red.
5. Create `PY-ANTI-PATTERNS.md` listing banned constructs.
6. Compute CRAP scores; identify the highest-CRAP function/module.
7. Refactor that target; re-measure; update `QA-REPORT.md`.
8. Repeat until green.
9. Add the CI `quality` workflow so the gate is enforced on every PR.

---

## 8. Repo-specific note (OTLMOW-Template)

- Primary package: `otlmow_template`
- Tests: `UnitTests/`
- Active high-CRAP candidate: `otlmow_template/SubsetTemplateCreator.py` (749 lines, sync + async duplicates, `ThreadPoolExecutor`, nested `while True` retry loops) — measure its CRAP first.
- Existing CI: `.github/workflows/unittest.yml` (runs pytest on 3.9–3.14 across Ubuntu/Windows/macOS).

Replace `<package>` and `tests` with `otlmow_template` and `UnitTests` when applying to this repo.
