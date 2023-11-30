import os

from UnitTests.TestModel.OtlmowModel.Classes.Onderdeel.AllCasesTestClass import AllCasesTestClass
from UnitTests.TestModel.OtlmowModel.Classes.Onderdeel.AnotherTestClass import AnotherTestClass
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


def test_add_attribute_info_csv():
    header = ['typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'testBooleanField']
    data = [['https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#AllCasesTestClass',
             '0000', 'AWV', 'True']]
    instantiated_objects = [AllCasesTestClass()]
    info_data = CsvTemplateCreator.add_attribute_info_csv(header, data, instantiated_objects)
    assert info_data == ['De URI van het object volgens https://www.w3.org/2001/XMLSchema#anyURI .',
                         'Een groep van tekens om een AIM object te identificeren of te benoemen.',
                         'Gegevens van de organisatie die de toekenning deed.',
                         'Test attribuut voor BooleanField']


def test_highlight_deprecated_attributes_csv():
    header = ['typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'deprecatedString']
    data = [['https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#AnotherTestClass',
             '0000', 'AWV', 'deprecated']]
    instantiated_objects = [AnotherTestClass()]
    info_data = CsvTemplateCreator.highlight_deprecated_attributes_csv(header, data, instantiated_objects)
    assert info_data == ['typeURI',
                         'assetId.identificator',
                         'assetId.toegekendDoor',
                         '[DEPRECATED] deprecatedString']


def test_find_uri_in_csv_returns_index_of_uri():
    data = ['test1', 'typeURI', 'test3']
    index = CsvTemplateCreator().get_type_uri_index_in_row(header=data)
    assert index == 1


def test_find_uri_in_csv_returns_none_if_uri_not_found():
    data = ['test1', 'test2', 'test3']
    index = CsvTemplateCreator().get_type_uri_index_in_row(header=data)
    assert index is None
