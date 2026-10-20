"""Structure snapshots of generated template files.

A generated template is not byte-comparable: the dummy values are random and the
order of the sheets, of the columns and of the choice lists depends on the order
in which the parallel generation finishes. These helpers normalise all of that,
so two runs can be compared on *structure* — which sheet exists, which columns it
has, which cells are filled, what type of value they hold, which data validations
are attached — instead of on luck.

Values keep their type, so a number that turns into a string is still detected,
but the random content itself is not compared.
"""

from __future__ import annotations

import csv
import datetime
import json
from pathlib import Path
from typing import Any

import openpyxl

# The model picks one branch of a union type at random, so the columns of a union
# type differ from run to run. Their headers are collapsed into a single key.
UNSTABLE_PREFIXES = ('testUnionType',)

GOLDEN_DIR = Path(__file__).parent / 'snapshots'


def normalise_value(value: Any) -> Any:
    """Replace a value with a marker that keeps its type."""
    if value is None:
        return None
    if isinstance(value, bool):
        return '<bool>'
    if isinstance(value, (datetime.datetime, datetime.date)):
        return '<date>'
    if isinstance(value, (int, float)):
        return '<number>'
    return '<text>'


def transpose(rows: list[list[Any]]) -> list[list[Any]]:
    """Turn a list of rows into a list of columns."""
    if not rows:
        return []
    width = max(len(row) for row in rows)
    return [[row[index] if index < len(row) else None for row in rows] for index in range(width)]


def sort_choice_list_columns(rows: list[list[Any]]) -> list[list[Any]]:
    """Sort the columns of the choice-list sheet by their header.

    The choice lists are added in the order the attributes are encountered, which
    depends on the order in which the classes were generated.
    """
    columns = transpose(rows)
    columns.sort(key=lambda column: (column[0] is None, str(column[0])))
    return transpose(columns)


def collapse_unstable_columns(
    columns: dict[str, list[Any]], unstable_prefixes: list[str]
) -> dict[str, list[Any]]:
    """Collapse the columns of a union type into a single, order-independent key."""
    for prefix in unstable_prefixes:
        matching = [header for header in columns if header.startswith(prefix)]
        if not matching:
            continue
        markers = {marker for header in matching for marker in columns[header]}
        for header in matching:
            del columns[header]
        columns[prefix + '*'] = sorted(markers)
    return columns


def snapshot_workbook(
    file_path: Path, unstable_prefixes: list[str] | None = None
) -> dict[str, Any]:
    """Return a comparable snapshot of a generated workbook."""
    if unstable_prefixes is None:
        unstable_prefixes = list(UNSTABLE_PREFIXES)
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    snapshot: dict[str, Any] = {}
    try:
        for sheet in workbook.worksheets:
            rows = [[cell.value for cell in row] for row in sheet.rows]
            if sheet.title == 'Keuzelijsten':
                # the choice lists themselves are deterministic, only their order is not
                rows = sort_choice_list_columns(rows)
                columns = {
                    str(column[0]): list(column[1:])
                    for column in transpose(rows)
                    if column and column[0] is not None
                }
            else:
                headers = (
                    [str(value) if value is not None else '' for value in rows[0]] if rows else []
                )
                data = [[normalise_value(value) for value in row] for row in rows[1:]]
                columns = {}
                for index, header in enumerate(headers):
                    markers = {
                        row[index] for row in data if index < len(row) and row[index] is not None
                    }
                    columns[header] = sorted(markers)
                columns = collapse_unstable_columns(columns, unstable_prefixes)
            snapshot[sheet.title] = {
                'columns': columns,
                'validations': sorted(
                    str(validation.sqref) for validation in sheet.data_validations.dataValidation
                ),
            }
    finally:
        workbook.close()
    return dict(sorted(snapshot.items()))


def snapshot_csv(file_path: Path) -> list[list[Any]]:
    """Return a comparable snapshot of a generated csv file."""
    with Path(file_path).open(encoding='utf-8', newline='') as handle:
        rows = list(csv.reader(handle, delimiter=';'))
    return [[normalise_value(cell) if cell != '' else None for cell in row] for row in rows]


def golden_path(name: str) -> Path:
    """Return the path of the golden snapshot of a template."""
    return GOLDEN_DIR / f'{name}.structure.json'


def save_golden(name: str, snapshot: dict[str, Any]) -> Path:
    """Write the golden snapshot of a template."""
    path = golden_path(name)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8') as handle:
        json.dump(snapshot, handle, indent=2, sort_keys=True)
        handle.write('\n')
    return path


def load_golden(name: str) -> dict[str, Any]:
    """Read the golden snapshot of a template."""
    with golden_path(name).open(encoding='utf-8') as handle:
        snapshot: dict[str, Any] = json.load(handle)
    return snapshot
