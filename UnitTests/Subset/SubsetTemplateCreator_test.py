import os
import shutil
from pathlib import Path

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))


# Toegangsprocedure en Agent zijn niet meer in de subset aanwezig
def test_func1(subtests):
    subset_tool = SubsetTemplateCreator()
    subset_location = Path(ROOT_DIR) / 'Flitspaal_noAgent3.0.db'
    xls_location = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.xlsx'
    csv_location = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.csv'

    with subtests.test(msg='xls'):
        subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                  path_to_template_file_and_extension=xls_location,
                                                  )
        template_path = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.xlsx'
        assert template_path.exists()

    with subtests.test(msg='csv, split per type'):
        subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                  path_to_template_file_and_extension=csv_location,
                                                  split_per_type=True)

    with subtests.test(msg='csv, not split per type'):
        subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                  path_to_template_file_and_extension=csv_location,
                                                  split_per_type=False)
        template_path = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.csv'
        assert template_path.exists()

    path = Path(ROOT_DIR) / 'testFileStorage'
    [f.unlink() for f in Path(path).glob("*") if f.is_file()]
    # Add an __init__.py file to the testFileStorage folder to make it a package
    open(Path(ROOT_DIR) / 'testFileStorage' / '__init__.py', 'a').close()


def test_subset_actual_subset():
    subset_tool = SubsetTemplateCreator()
    csv_location = Path(ROOT_DIR) / 'testFileStorage' / 'camera_steun.csv'
    subset_tool.generate_template_from_subset(path_to_subset=Path(ROOT_DIR) / 'camera_steun.db',
                                              path_to_template_file_and_extension=csv_location,
                                              split_per_type=True)
    csv1 = Path(ROOT_DIR) / 'testFileStorage' / 'camera_steun_onderdeel_Bevestiging.csv'
    csv2 = Path(ROOT_DIR) / 'testFileStorage' / 'camera_steun_onderdeel_Camera.csv'
    csv3 = Path(ROOT_DIR) / 'testFileStorage' / 'camera_steun_onderdeel_RechteSteun.csv'
    assert csv1.exists()
    assert csv2.exists()
    assert csv3.exists()
    path = Path(ROOT_DIR) / 'testFileStorage'
    [f.unlink() for f in Path(path).glob("*") if f.is_file()]
    open(Path(ROOT_DIR) / 'testFileStorage' / '__init__.py', 'a').close()


def test_filter():
    db_location = Path(ROOT_DIR) / 'Flitspaal_noAgent3.0.db'
    list_of_filter_uri = ['https://wegenenverkeer.data.vlaanderen.be/ns/installatie#Flitspaal']
    filtered = SubsetTemplateCreator.filters_assets_by_subset(db_location, list_of_otl_objectUri=list_of_filter_uri)
    assert len(filtered) == 1
    assert filtered[0].name == 'Flitspaal'


def test_empty_filter_list_removes_all_entries():
    db_location = Path(ROOT_DIR) / 'Flitspaal_noAgent3.0.db'
    list_of_filter_uri = []
    filtered = SubsetTemplateCreator.filters_assets_by_subset(db_location, list_of_otl_objectUri=list_of_filter_uri)
    assert len(filtered) == 0
