"""Behaviour preservation: the structure of a generated template is pinned.

The generated template contains random dummy values and its sheet and column
order depends on the order in which the parallel generation finishes, so it
cannot be compared byte by byte. The snapshot in ``template_snapshot`` normalises
that away and compares the *structure*: the sheets, their columns, which cells
are filled, the type of every value and the attached data validations.

This is the test that makes a refactor safe: it fails when the generated
template changes, and it does not fail when only the random values change.

Regenerate the golden file with::

    UPDATE_GOLDEN=1 python -m pytest UnitTests/Subset/SubsetTemplateCreator_structure_test.py
"""

from __future__ import annotations

import os
import random
from pathlib import Path

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator
from template_snapshot import load_golden, save_golden, snapshot_workbook

current_dir = Path(__file__).parent
model_directory_path = Path(__file__).parent.parent / 'TestModel'

GOLDEN_NAME = 'OTL_AllCasesTestClass'
SEED = 42


def generate_template(template_path: Path, **options) -> None:
    """Generate a template with a fixed seed, so the run is reproducible."""
    random.seed(SEED)
    SubsetTemplateCreator().generate_template_from_subset(
        subset_path=current_dir / 'OTL_AllCasesTestClass.db',
        template_file_path=template_path,
        model_directory=model_directory_path,
        **options,
    )


def test_structure_of_the_generated_template_is_pinned(tmp_path):
    """The structure of the generated template must not change unnoticed."""
    template_path = tmp_path / 'OTL_AllCasesTestClass.xlsx'
    generate_template(template_path)
    snapshot = snapshot_workbook(template_path)

    if os.environ.get('UPDATE_GOLDEN') == '1':
        save_golden(GOLDEN_NAME, snapshot)

    assert snapshot == load_golden(GOLDEN_NAME)


def test_generation_is_reproducible(tmp_path):
    """Two runs with the same seed produce the same template."""
    first_path = tmp_path / 'first.xlsx'
    second_path = tmp_path / 'second.xlsx'
    generate_template(first_path)
    generate_template(second_path)

    assert snapshot_workbook(first_path) == snapshot_workbook(second_path)


def test_every_sheet_has_a_type_uri_column(tmp_path):
    """Each asset sheet starts with the typeURI column."""
    template_path = tmp_path / 'OTL_AllCasesTestClass.xlsx'
    generate_template(template_path)
    snapshot = snapshot_workbook(template_path)

    asset_sheets = [name for name in snapshot if name != 'Keuzelijsten']
    assert asset_sheets
    for name in asset_sheets:
        assert 'typeURI' in snapshot[name]['columns']
        assert 'assetId.identificator' in snapshot[name]['columns']


def test_choice_lists_are_generated(tmp_path):
    """The choice lists of the test model end up in the template."""
    template_path = tmp_path / 'OTL_AllCasesTestClass.xlsx'
    generate_template(template_path)
    snapshot = snapshot_workbook(template_path)

    choice_lists = snapshot['Keuzelijsten']['columns']
    assert 'KlAIMToestand' in choice_lists
    assert 'KlTestKeuzelijst' in choice_lists
    assert 'geannuleerd' in choice_lists['KlAIMToestand']
    assert 'in-gebruik' in choice_lists['KlAIMToestand']


def test_choice_lists_have_data_validations(tmp_path):
    """The columns with a choice list get a data validation attached."""
    template_path = tmp_path / 'OTL_AllCasesTestClass.xlsx'
    generate_template(template_path)
    snapshot = snapshot_workbook(template_path)

    validations = snapshot['onderdeel#AllCasesTestClass']['validations']
    assert validations
