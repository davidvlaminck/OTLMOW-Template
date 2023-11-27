import os
import tempfile
from pathlib import Path

from otlmow_template.CsvTemplateCreator import CsvTemplateCreator
from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))


def test_files_get_generated(subtests):
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


def test_filter_returns_filtered_list():
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


def test_no_filter_list_returns_all_entries():
    db_location = Path(ROOT_DIR) / 'Flitspaal_noAgent3.0.db'
    filtered = SubsetTemplateCreator.filters_assets_by_subset(db_location)
    assert len(filtered) == 11


def test_remove_mockdata_csv_clears_data_if_no_examples_wanted():
    data = ['test1', 'test2', 'test3']
    new_data = CsvTemplateCreator().remove_mock_data_csv(data=data, rows_of_examples=0)
    assert new_data == []


def test_remove_mockdata_csv_leaves_data_intact_if_examples_wanted():
    data = ['test1', 'test2', 'test3']
    new_data = CsvTemplateCreator().remove_mock_data_csv(data=data, rows_of_examples=1)
    assert new_data == data


def test_find_uri_in_csv_returns_index_of_uri():
    data = ['test1', 'typeURI', 'test3']
    index = CsvTemplateCreator().find_uri_in_csv(header=data)
    assert index == 1


def test_find_uri_in_csv_returns_none_if_uri_not_found():
    data = ['test1', 'test2', 'test3']
    index = CsvTemplateCreator().find_uri_in_csv(header=data)
    assert index is None


def test_return_temp_path_returns_path_to_temporary_file():
    path_to_template_file_and_extension = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.xlsx'
    temp_path = SubsetTemplateCreator.return_temp_path(path_to_template_file_and_extension)
    assert temp_path == Path(tempfile.gettempdir()) / 'temp-otlmow' / 'template_file_text.xlsx'



