"""The synchronous and the asynchronous implementation must behave the same.

``generate_template_from_subset`` and ``generate_template_from_subset_async``
are two copies of the same logic. The happy-path suite only exercises a small
part of the options of the asynchronous one, so the two copies can drift apart
without any test noticing: the asynchronous variant would silently produce a
template without relations, without deprecated markers or with different sheet
titles.

Every option of the public api is therefore generated twice — once synchronously
and once asynchronously — and the two results are compared on structure.
"""

from __future__ import annotations

import random
from pathlib import Path

import pytest
from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator
from template_snapshot import snapshot_workbook

current_dir = Path(__file__).parent
model_directory_path = Path(__file__).parent.parent / 'TestModel'

SEED = 42

OPTIONS = [
    pytest.param({}, id='default'),
    pytest.param({'ignore_relations': False}, id='with-relations'),
    pytest.param({'add_deprecated': True}, id='with-deprecated'),
    pytest.param({'add_attribute_info': True}, id='with-attribute-info'),
    pytest.param({'abbreviate_excel_sheettitles': True}, id='with-abbreviated-sheet-titles'),
    pytest.param({'filter_attributes_by_subset': False}, id='without-subset-filter'),
    pytest.param({'generate_choice_list': False}, id='without-choice-list'),
    pytest.param({'add_geometry': False}, id='without-geometry'),
    pytest.param({'dummy_data_rows': 0}, id='without-data-rows'),
]


def generate_sync(template_path: Path, **options) -> None:
    """Generate a template synchronously."""
    random.seed(SEED)
    SubsetTemplateCreator().generate_template_from_subset(
        subset_path=current_dir / 'OTL_AllCasesTestClass.db',
        template_file_path=template_path,
        model_directory=model_directory_path,
        **options,
    )


async def generate_async(template_path: Path, **options) -> None:
    """Generate a template asynchronously."""
    random.seed(SEED)
    await SubsetTemplateCreator().generate_template_from_subset_async(
        subset_path=current_dir / 'OTL_AllCasesTestClass.db',
        template_file_path=template_path,
        model_directory=model_directory_path,
        **options,
    )


@pytest.mark.parametrize('options', OPTIONS)
@pytest.mark.asyncio
async def test_sync_and_async_generate_the_same_template(tmp_path, options):
    """Both implementations must produce the same template."""
    sync_path = tmp_path / 'sync.xlsx'
    async_path = tmp_path / 'async.xlsx'

    generate_sync(sync_path, **options)
    await generate_async(async_path, **options)

    assert snapshot_workbook(sync_path) == snapshot_workbook(async_path)


@pytest.mark.asyncio
async def test_async_with_relations_contains_the_relation_sheets(tmp_path):
    """The asynchronous variant must add the relations as well."""
    template_path = tmp_path / 'async_relations.xlsx'
    await generate_async(template_path, ignore_relations=False)
    snapshot = snapshot_workbook(template_path)

    assert 'onderdeel#Bevestiging' in snapshot
    assert 'bron.typeURI' in snapshot['onderdeel#Bevestiging']['columns']


@pytest.mark.asyncio
async def test_async_with_abbreviated_sheet_titles_has_short_titles(tmp_path):
    """The asynchronous variant must abbreviate the sheet titles as well."""
    template_path = tmp_path / 'async_abbreviated.xlsx'
    await generate_async(template_path, abbreviate_excel_sheettitles=True)
    snapshot = snapshot_workbook(template_path)

    assert all(len(name) <= 31 for name in snapshot)
