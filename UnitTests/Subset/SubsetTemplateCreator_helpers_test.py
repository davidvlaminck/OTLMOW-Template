"""The edge paths of the sheet and csv helpers of SubsetTemplateCreator.

The template tests only generate templates for the happy path. The helpers that
the refactor split out of ``alter_excel_sheet`` and ``alter_csv_file`` also
decide what is *skipped*: a column without a header, the asset id of the agent
sheet, a sheet whose title is only the beginning of the type uri. A change in
those rules moves the columns of a template without failing a single test, so
they are pinned here.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

current_dir = Path(__file__).parent
model_directory_path = Path(__file__).parent.parent / 'TestModel'

TYPE_URI_DESCRIPTION = 'De URI van het object volgens https://www.w3.org/2001/XMLSchema#anyURI .'


class FakeInstance:
    """An instance with nothing but the type of the sheet it belongs to."""

    def __init__(self, type_uri: str):
        self.typeURI = type_uri


class FakeStandaardEenheid:
    """The standard unit of a datatype, with its two notations."""

    def __init__(self, usagenote: str, constraints: str):
        self.usagenote = usagenote
        self.constraints = constraints


class FakeWaarde:
    """The value of a quantitative attribute."""

    def __init__(self, standaard_eenheid=None):
        self._standaardEenheid = standaard_eenheid


class FakeAttribuut:
    """A quantitative attribute with a value."""

    def __init__(self, waarde):
        self.waarde = waarde

    def add_empty_value(self):
        self.waarde = FakeWaarde()


def build_sheet(title: str, headers: list[str]) -> tuple:
    """Return a workbook with one sheet that has the given headers."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = title
    for index, header in enumerate(headers, start=1):
        sheet.cell(row=1, column=index).value = header
    return workbook, sheet


def test_a_column_without_a_header_is_skipped():
    """A column without a header gets no information and no validation."""
    workbook, sheet = build_sheet('onderdeel#Test', ['typeURI', None])
    # a cell that is emptied keeps the column in the sheet
    sheet['B1'] = 'leeg'
    sheet['B1'] = None

    SubsetTemplateCreator.alter_excel_sheet(
        add_attribute_info=True,
        choice_list_dict={},
        generate_choice_list=False,
        dummy_data_rows=0,
        instances=[FakeInstance('https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#Test')],
        sheet=sheet,
        add_deprecated=True,
        workbook=workbook,
    )

    # the attribute info row is inserted in front of the header row
    assert sheet['A1'].value == TYPE_URI_DESCRIPTION
    assert sheet['B1'].value is None
    assert sheet['A2'].value == 'typeURI'
    assert sheet['B2'].value is None


def test_the_asset_id_of_the_agent_sheet_is_skipped():
    """The agent sheet has no asset id, so that column is left alone."""
    workbook, sheet = build_sheet('Agent', ['typeURI', 'assetId.identificator'])

    SubsetTemplateCreator.alter_excel_sheet(
        add_attribute_info=True,
        choice_list_dict={},
        generate_choice_list=False,
        dummy_data_rows=0,
        instances=[FakeInstance('http://purl.org/dc/terms/Agent')],
        sheet=sheet,
        add_deprecated=True,
        workbook=workbook,
    )

    # only the type uri gets a definition, the asset id column is skipped
    assert sheet['A1'].value == TYPE_URI_DESCRIPTION
    assert sheet['B1'].value is None
    assert sheet['A2'].value == 'typeURI'
    assert sheet['B2'].value == 'assetId.identificator'


def test_a_sheet_title_that_is_the_beginning_of_the_type_uri_finds_its_instance():
    """A sheet title that is shorter than the type uri still finds its instance."""
    workbook, sheet = build_sheet('onderdeel#Test', ['typeURI'])

    SubsetTemplateCreator.alter_excel_sheet(
        add_attribute_info=False,
        choice_list_dict={},
        generate_choice_list=False,
        dummy_data_rows=0,
        instances=[
            FakeInstance('https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#TestKlasse')
        ],
        sheet=sheet,
        add_deprecated=False,
        workbook=workbook,
    )

    assert sheet['A1'].value == 'typeURI'


def test_the_standard_unit_is_read_from_the_usagenote():
    """The standard unit of a datatype is the value between the quotes of its usagenote."""
    attribute = FakeAttribuut(FakeWaarde(FakeStandaardEenheid(usagenote='"m"', constraints='"km"')))

    assert SubsetTemplateCreator.get_standard_unit(attribute) == 'm'


def test_the_standard_unit_is_read_from_the_constraints():
    """A datatype without a usagenote keeps its standard unit in the constraints."""
    attribute = FakeAttribuut(
        FakeWaarde(FakeStandaardEenheid(usagenote='geen', constraints='"km"'))
    )

    assert SubsetTemplateCreator.get_standard_unit(attribute) == 'km'


def test_a_datatype_without_a_standard_unit_has_none():
    """A datatype without a standard unit gives None."""
    attribute = FakeAttribuut(
        FakeWaarde(FakeStandaardEenheid(usagenote='geen', constraints='ook geen'))
    )

    assert SubsetTemplateCreator.get_standard_unit(attribute) is None


def test_a_value_without_a_standard_unit_has_none():
    """A value without a standard unit gives None."""
    attribute = FakeAttribuut(FakeWaarde())

    assert SubsetTemplateCreator.get_standard_unit(attribute) is None


def test_an_empty_value_gets_an_empty_value_first():
    """A quantitative attribute without a value gets one before the unit is looked up."""
    attribute = FakeAttribuut(None)

    assert SubsetTemplateCreator.get_standard_unit(attribute) is None
    assert attribute.waarde is not None
