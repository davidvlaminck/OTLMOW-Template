
import os
import shutil
from pathlib import Path

import pytest

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

ROOT_DIR = Path(__file__).parent


# Toegangsprocedure en Agent zijn niet meer in de subset aanwezig
def test_func1(subtests):
    subset_tool = SubsetTemplateCreator()
    subset_location = ROOT_DIR / 'Flitspaal_noAgent3.0.db'
    xls_location = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.xlsx'
    csv_location = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.csv'

    with subtests.test(msg='xls'):
        subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                  path_to_template_file_and_extension=xls_location,
                                                  )
        assert Path(ROOT_DIR / 'testFileStorage' / 'template_file_text.xlsx').exists()

    with subtests.test(msg='csv, split per type'):
        subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                  path_to_template_file_and_extension=csv_location,
                                                  split_per_type=True)

    with subtests.test(msg='csv, not split per type'):
        subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                  path_to_template_file_and_extension=csv_location,
                                                  split_per_type=False)
        assert Path(ROOT_DIR / 'testFileStorage' / 'template_file_text.csv').exists()

    shutil.rmtree(Path(ROOT_DIR) / 'testFileStorage')
    os.makedirs(Path(ROOT_DIR) / 'testFileStorage')


def test_subset_actual_subset():
    subset_tool = SubsetTemplateCreator()
    csv_location = ROOT_DIR / 'testFileStorage' / 'camera_steun.csv'
    subset_tool.generate_template_from_subset(path_to_subset=ROOT_DIR / 'camera_steun.db',
                                              path_to_template_file_and_extension=csv_location,
                                              split_per_type=True)
    csv1 = ROOT_DIR / 'testFileStorage' / 'camera_steun_onderdeel_Bevestiging.csv'
    csv2 = ROOT_DIR / 'testFileStorage' / 'camera_steun_onderdeel_Camera.csv'
    csv3 = ROOT_DIR / 'testFileStorage' / 'camera_steun_onderdeel_RechteSteun.csv'
    assert csv1.exists()
    assert csv2.exists()
    assert csv3.exists()
    shutil.rmtree(Path(ROOT_DIR) / 'testFileStorage')
    os.makedirs(Path(ROOT_DIR) / 'testFileStorage')


def test_filter():
    db_location = ROOT_DIR / 'flitspaal_noAgent3.0.db'
    list_of_filter_uri = ['https://wegenenverkeer.data.vlaanderen.be/ns/installatie#Flitspaal']
    filtered = SubsetTemplateCreator.filters_assets_by_subset(db_location, list_of_filter_uri)
    assert len(filtered) == 1
    assert filtered[0].name == 'Flitspaal'


def test_empty_filter_list_removes_all_entries():
    db_location = ROOT_DIR / 'flitspaal_noAgent3.0.db'
    list_of_filter_uri = []
    filtered = SubsetTemplateCreator.filters_assets_by_subset(db_location, list_of_filter_uri)
    assert len(filtered) == 0