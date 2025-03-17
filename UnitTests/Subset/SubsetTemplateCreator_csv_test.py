import csv
from pathlib import Path

import pytest

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

current_dir = Path(__file__).parent
model_directory_path = Path(__file__).parent.parent / 'TestModel'


@pytest.mark.parametrize(
    "index, amount_of_examples, add_geo_artefact, add_attribute_info, add_deprecated_info",
    [
        (i, amount, geo, attr, dep) for i, (amount, geo, attr, dep) in enumerate(
        [
            (0, True, True, True),
            (0, True, True, False),
            (0, True, False, True),
            (0, True, False, False),
            (0, False, True, True),
            (0, False, True, False),
            (0, False, False, True),
            (0, False, False, False),
            (1, True, True, True),
            (1, True, True, False),
            (1, True, False, True),
            (1, True, False, False),
            (1, False, True, True),
            (1, False, True, False),
            (1, False, False, True),
            (1, False, False, False),
            (2, True, True, True),
            (2, True, True, False),
            (2, True, False, True),
            (2, True, False, False),
            (2, False, True, True),
            (2, False, True, False),
            (2, False, False, True),
            (2, False, False, False),
        ]
    )
    ]
)
def test_generate_csv_template(index, amount_of_examples, add_geo_artefact, add_attribute_info,
                               add_deprecated_info):
    # Arrange
    path_to_subset = current_dir / 'OTL_AllCasesTestClass.db'
    path_to_template_file = current_dir / f'OTL_AllCasesTestClass_{index}.csv'
    kwargs = {
        'amount_of_examples': amount_of_examples,
        'add_geo_artefact': add_geo_artefact,
        'add_attribute_info': add_attribute_info,
        'highlight_deprecated_attributes': add_deprecated_info,
    }

    # Act
    subset_tool = SubsetTemplateCreator()
    subset_tool.generate_template_from_subset(
        path_to_subset=path_to_subset, path_to_template_file_and_extension=path_to_template_file, split_per_type=False,
        list_of_otl_objectUri=["https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#AnotherTestClass"],
        model_directory=model_directory_path, **kwargs)

    # Assert
    # sourcery skip: no-conditionals-in-tests
    with open(path_to_template_file, newline='') as output_file:
        output_reader = csv.reader(output_file, delimiter=';')
        output_rows = list(output_reader)

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

    expected_row_count = amount_of_examples + 1

    if add_geo_artefact:
        expected_header_row.insert(6, 'geometry')
        expected_attribute_info_row.insert(6, 'geometry voor DAVIE')

    if add_deprecated_info:
        expected_header_row.insert(5, f'[DEPRECATED] {expected_header_row.pop(5)}')

    if add_attribute_info:
        assert output_rows[0] == expected_attribute_info_row
        assert output_rows[1] == expected_header_row
        expected_row_count += 1
    else:
        assert output_rows[0] == expected_header_row

    assert len(output_rows) == expected_row_count

    path_to_template_file.unlink()


@pytest.mark.parametrize(
    "index, amount_of_examples, add_geo_artefact, add_attribute_info, add_deprecated_info",
    [
        (i, amount, geo, attr, dep) for i, (amount, geo, attr, dep) in enumerate(
        [
            (0, True, True, True),
            (0, True, True, False),
            (0, True, False, True),
            (0, True, False, False),
            (0, False, True, True),
            (0, False, True, False),
            (0, False, False, True),
            (0, False, False, False),
            (1, True, True, True),
            (1, True, True, False),
            (1, True, False, True),
            (1, True, False, False),
            (1, False, True, True),
            (1, False, True, False),
            (1, False, False, True),
            (1, False, False, False),
            (2, True, True, True),
            (2, True, True, False),
            (2, True, False, True),
            (2, True, False, False),
            (2, False, True, True),
            (2, False, True, False),
            (2, False, False, True),
            (2, False, False, False),
        ]
    )
    ]
)
@pytest.mark.asyncio
async def test_generate_csv_template_async(index, amount_of_examples, add_geo_artefact, add_attribute_info,
                               add_deprecated_info):
    # Arrange
    path_to_subset = current_dir / 'OTL_AllCasesTestClass.db'
    path_to_template_file = current_dir / f'OTL_AllCasesTestClass_async_{index}.csv'
    kwargs = {
        'amount_of_examples': amount_of_examples,
        'add_geo_artefact': add_geo_artefact,
        'add_attribute_info': add_attribute_info,
        'highlight_deprecated_attributes': add_deprecated_info,
    }

    # Act
    subset_tool = SubsetTemplateCreator()
    await subset_tool.generate_template_from_subset_async(
        path_to_subset=path_to_subset, path_to_template_file_and_extension=path_to_template_file, split_per_type=False,
        list_of_otl_objectUri=["https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#AnotherTestClass"],
        model_directory=model_directory_path, **kwargs)

    # Assert
    # sourcery skip: no-conditionals-in-tests
    with open(path_to_template_file, newline='') as output_file:
        output_reader = csv.reader(output_file, delimiter=';')
        output_rows = list(output_reader)

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

    expected_row_count = amount_of_examples + 1

    if add_geo_artefact:
        expected_header_row.insert(6, 'geometry')
        expected_attribute_info_row.insert(6, 'geometry voor DAVIE')

    if add_deprecated_info:
        expected_header_row.insert(5, f'[DEPRECATED] {expected_header_row.pop(5)}')

    if add_attribute_info:
        assert output_rows[0] == expected_attribute_info_row
        assert output_rows[1] == expected_header_row
        expected_row_count += 1
    else:
        assert output_rows[0] == expected_header_row

    assert len(output_rows) == expected_row_count

    path_to_template_file.unlink()
