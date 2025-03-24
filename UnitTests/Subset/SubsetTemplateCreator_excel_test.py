from itertools import product
from pathlib import Path

import openpyxl
import pytest

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

current_dir = Path(__file__).parent
model_directory_path = Path(__file__).parent.parent / 'TestModel'


# create one unit test for generating a excel template (see the generic_csv test) for every combination of the
# following parameters
# dummy_data_rows: 0, 1, 2
# add_geometry: True, False
# add_attribute_info: True, False
# add_deprecated: True, False
# generate_choice_list: True, False


@pytest.mark.parametrize(
    "index, dummy_data_rows, add_geometry, add_attribute_info, add_deprecated, generate_choice_list",
    [
        (i, amount, geo, attr, dep, choice) for i, (amount, geo, attr, dep, choice) in enumerate(
        product(
            [0, 1, 2],
            [True, False],
            [True, False],
            [True, False],
            [True, False])
    )
    ]
)
def test_generate_excel_template(index, dummy_data_rows, add_geometry, add_attribute_info,
                               add_deprecated, generate_choice_list):
    # Arrange
    subset_path = current_dir / 'OTL_AllCasesTestClass.db'
    path_to_template_file = current_dir / f'OTL_AllCasesTestClass_{index}.xlsx'
    kwargs = {
        'dummy_data_rows': dummy_data_rows,
        'add_geometry': add_geometry,
        'add_attribute_info': add_attribute_info,
        'add_deprecated': add_deprecated,
        'generate_choice_list': generate_choice_list,
    }

    # Act
    subset_tool = SubsetTemplateCreator()
    subset_tool.generate_template_from_subset(
        subset_path=subset_path, template_file_path=path_to_template_file,
        class_uris_filter=["https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#AnotherTestClass"],
        model_directory=model_directory_path, **kwargs)

    # Assert
    # sourcery skip: no-conditionals-in-tests

    book = openpyxl.load_workbook(path_to_template_file, read_only=True, data_only=True)
    output_rows = []
    choice_list_sheet = False
    choice_list_data = []

    for sheet in book.worksheets:
        sheet_name = sheet.title
        if sheet_name == 'Keuzelijsten':
            choice_list_sheet = True
            for row in sheet.rows:
                choice_list_data.append([cell.value for cell in row])
            continue
        if sheet_name != 'onderdeel#AnotherTestClass':
            continue
        for row in sheet.rows:
            output_rows.append([cell.value for cell in row])
    book.close()

    expected_attribute_info_row = [
        'De URI van het object volgens https://www.w3.org/2001/XMLSchema#anyURI .',
        'Een groep van tekens om een AIM object te identificeren of te benoemen.',
        'Gegevens van de organisatie die de toekenning deed.',
        'Een verwijzing naar een postnummer uit het specifieke bestek waar het object mee verband houdt.',
        'Datum van de oprichting van het object.',
        'Tekstveld dat niet meer gebruikt wordt',
        'Geeft aan of het object actief kan gebruikt worden of (zacht) verwijderd is uit het asset beheer systeem.',
        'Extra notitie voor het object.',
        'Een verwijzing naar een postnummer uit het standaardbestek waar het object mee verband houdt. De notatie van het postnummer moet overeenkomen met de notatie die gebruikt is in de catalogi van standaardbestekken, bijvoorbeeld postnummer 0701.20404G.',
        'Bevat een getal die bij het datatype hoort.', 'Geeft de actuele stand in de levenscyclus van het object.']
    expected_header_row = ['typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'bestekPostNummer[]',
                           'datumOprichtingObject', 'deprecatedString', 'isActief', 'notitie',
                           'standaardBestekPostNummer[]', 'theoretischeLevensduur', 'toestand']
    expected_deprecated_header_row = [None, None, None, None, None, 'DEPRECATED', None, None, None, None, None]

    expected_row_count = dummy_data_rows + 1

    expected_column_count = 11

    if add_geometry:
        expected_header_row.insert(6, 'geometry')
        expected_attribute_info_row.insert(6, 'geometry voor DAVIE')
        expected_deprecated_header_row.insert(6, None)
        expected_column_count += 1

    header_index = 0
    deprecated_index = 0

    if add_attribute_info:
        assert output_rows[0] == expected_attribute_info_row
        header_index += 1
        expected_row_count += 1
        deprecated_index += 1

    if add_deprecated:
        assert output_rows[deprecated_index] == expected_deprecated_header_row
        header_index += 1
        expected_row_count += 1

    assert len(output_rows) == expected_row_count

    if generate_choice_list:
        assert choice_list_sheet
        assert choice_list_data[0][0] == 'KlAIMToestand'

    path_to_template_file.unlink()


@pytest.mark.parametrize(
    "index, dummy_data_rows, add_geometry, add_attribute_info, add_deprecated, generate_choice_list",
    [
        (i, amount, geo, attr, dep, choice) for i, (amount, geo, attr, dep, choice) in enumerate(
        product(
            [0, 1, 2],
            [True, False],
            [True, False],
            [True, False],
            [True, False])
    )
    ]
)
@pytest.mark.asyncio
async def test_generate_excel_template_async(index, dummy_data_rows, add_geometry, add_attribute_info,
                               add_deprecated, generate_choice_list):
    # Arrange
    subset_path = current_dir / 'OTL_AllCasesTestClass.db'
    path_to_template_file = current_dir / f'OTL_AllCasesTestClass_{index}.xlsx'
    kwargs = {
        'dummy_data_rows': dummy_data_rows,
        'add_geometry': add_geometry,
        'add_attribute_info': add_attribute_info,
        'add_deprecated': add_deprecated,
        'generate_choice_list': generate_choice_list,
    }

    # Act
    subset_tool = SubsetTemplateCreator()
    await subset_tool.generate_template_from_subset_async(
        subset_path=subset_path, template_file_path=path_to_template_file,
        class_uris_filter=["https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#AnotherTestClass"],
        model_directory=model_directory_path, **kwargs)

    # Assert
    # sourcery skip: no-conditionals-in-tests

    book = openpyxl.load_workbook(path_to_template_file, read_only=True, data_only=True)
    output_rows = []
    choice_list_sheet = False
    choice_list_data = []

    for sheet in book.worksheets:
        sheet_name = sheet.title
        if sheet_name == 'Keuzelijsten':
            choice_list_sheet = True
            for row in sheet.rows:
                choice_list_data.append([cell.value for cell in row])
            continue
        if sheet_name != 'onderdeel#AnotherTestClass':
            continue
        for row in sheet.rows:
            output_rows.append([cell.value for cell in row])
    book.close()

    expected_attribute_info_row = [
        'De URI van het object volgens https://www.w3.org/2001/XMLSchema#anyURI .',
        'Een groep van tekens om een AIM object te identificeren of te benoemen.',
        'Gegevens van de organisatie die de toekenning deed.',
        'Een verwijzing naar een postnummer uit het specifieke bestek waar het object mee verband houdt.',
        'Datum van de oprichting van het object.',
        'Tekstveld dat niet meer gebruikt wordt',
        'Geeft aan of het object actief kan gebruikt worden of (zacht) verwijderd is uit het asset beheer systeem.',
        'Extra notitie voor het object.',
        'Een verwijzing naar een postnummer uit het standaardbestek waar het object mee verband houdt. De notatie van het postnummer moet overeenkomen met de notatie die gebruikt is in de catalogi van standaardbestekken, bijvoorbeeld postnummer 0701.20404G.',
        'Bevat een getal die bij het datatype hoort.', 'Geeft de actuele stand in de levenscyclus van het object.']
    expected_header_row = ['typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'bestekPostNummer[]',
                           'datumOprichtingObject', 'deprecatedString', 'isActief', 'notitie',
                           'standaardBestekPostNummer[]', 'theoretischeLevensduur', 'toestand']
    expected_deprecated_header_row = [None, None, None, None, None, 'DEPRECATED', None, None, None, None, None]

    expected_row_count = dummy_data_rows + 1

    expected_column_count = 11

    if add_geometry:
        expected_header_row.insert(6, 'geometry')
        expected_attribute_info_row.insert(6, 'geometry voor DAVIE')
        expected_deprecated_header_row.insert(6, None)
        expected_column_count += 1

    header_index = 0
    deprecated_index = 0

    if add_attribute_info:
        assert output_rows[0] == expected_attribute_info_row
        header_index += 1
        expected_row_count += 1
        deprecated_index += 1

    if add_deprecated:
        assert output_rows[deprecated_index] == expected_deprecated_header_row
        header_index += 1
        expected_row_count += 1

    assert len(output_rows) == expected_row_count

    if generate_choice_list:
        assert choice_list_sheet
        assert choice_list_data[0][0] == 'KlAIMToestand'

    path_to_template_file.unlink()
