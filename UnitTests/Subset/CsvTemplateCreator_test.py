import os

from otlmow_template.CsvTemplateCreator import CsvTemplateCreator

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))


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
