import csv
from pathlib import Path

import openpyxl
import pytest

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

current_dir = Path(__file__).parent
model_directory_path = Path(__file__).parent.parent / 'TestModel'
class_uri = "https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#AllCasesTestClass"
sheet_name = 'onderdeel#AllCasesTestClass'


def generate_template(path_to_template_file: Path) -> None:
    SubsetTemplateCreator().generate_template_from_subset(
        subset_path=current_dir / 'OTL_AllCasesTestClass.db', template_file_path=path_to_template_file,
        class_uris_filter=[class_uri], model_directory=model_directory_path, add_attribute_info=True)


def get_excel_attribute_info_by_header(path_to_template_file: Path) -> dict:
    book = openpyxl.load_workbook(path_to_template_file, read_only=True, data_only=True)
    attribute_info_row, header_row = list(book[sheet_name].iter_rows(min_row=1, max_row=2, values_only=True))
    book.close()
    return dict(zip(header_row, attribute_info_row))


def get_csv_attribute_info_by_header(path_to_template_file: Path) -> dict:
    with open(path_to_template_file, encoding='utf-8', newline='\n') as output_file:
        rows = list(csv.reader(output_file, delimiter=';'))
    return dict(zip(rows[1], rows[0]))


@pytest.fixture(scope='module')
def excel_attribute_info_by_header():
    path_to_template_file = current_dir / 'OTL_AllCasesTestClass_attribute_info.xlsx'
    generate_template(path_to_template_file)
    yield get_excel_attribute_info_by_header(path_to_template_file)
    path_to_template_file.unlink()


@pytest.fixture(scope='module')
def csv_attribute_info_by_header():
    path_to_template_file = current_dir / 'OTL_AllCasesTestClass_attribute_info.csv'
    generate_template(path_to_template_file)
    split_path_to_template_file = current_dir / 'OTL_AllCasesTestClass_attribute_info_onderdeel_AllCasesTestClass.csv'
    yield get_csv_attribute_info_by_header(split_path_to_template_file)
    path_to_template_file.unlink(missing_ok=True)
    split_path_to_template_file.unlink(missing_ok=True)


# the definition of a quantitative value refers to the 'waarde' attribute of the datatype, which does not
# explain the attribute itself. For those datatypes the definition of the attribute itself is expected,
# combined with the standard unit of the datatype.
# datatypes with a waarde shortcut but without a standard unit (ex. DteTestEenvoudigType) keep the definition
# of the value.
@pytest.mark.parametrize('header, expected_definition', [
    ('theoretischeLevensduur',
     'De levensduur in aantal maanden die theoretisch mag verwacht worden voor een object. Standaard eenheid: mo'),
    ('testKwantWrd', 'Test attribuut voor een kwantitatieve waarde. Standaard eenheid: %'),
    ('testKwantWrdMetKard[]',
     'Test attribuut voor een kwantitatieve waarde met kardinaliteit > 1. Standaard eenheid: %'),
    ('testComplexType.testKwantWrd',
     'Test attribuut voor Kwantitatieve waarde in een complex datatype. Standaard eenheid: %'),
    ('testComplexType.testComplexType2.testKwantWrd',
     'Test attribuut voor Kwantitatieve waarde in een complex datatype. Standaard eenheid: %'),
    ('testComplexTypeMetKard[].testKwantWrd',
     'Test attribuut voor Kwantitatieve waarde in een complex datatype. Standaard eenheid: %'),
    ('testEenvoudigType', 'De string die het eenvoudige test datatype voorstelt.'),
    ('testEenvoudigTypeMetKard[]', 'De string die het eenvoudige test datatype voorstelt.'),
    ('notitie', 'Extra notitie voor het object.'),
])
def test_attribute_info_in_excel(excel_attribute_info_by_header, header, expected_definition):
    assert excel_attribute_info_by_header[header] == expected_definition


@pytest.mark.parametrize('header, expected_definition', [
    ('theoretischeLevensduur',
     'De levensduur in aantal maanden die theoretisch mag verwacht worden voor een object. Standaard eenheid: mo'),
    ('testKwantWrd', 'Test attribuut voor een kwantitatieve waarde. Standaard eenheid: %'),
    ('testComplexType.testKwantWrd',
     'Test attribuut voor Kwantitatieve waarde in een complex datatype. Standaard eenheid: %'),
    ('testEenvoudigType', 'De string die het eenvoudige test datatype voorstelt.'),
    ('notitie', 'Extra notitie voor het object.'),
])
def test_attribute_info_in_csv(csv_attribute_info_by_header, header, expected_definition):
    assert csv_attribute_info_by_header[header] == expected_definition
