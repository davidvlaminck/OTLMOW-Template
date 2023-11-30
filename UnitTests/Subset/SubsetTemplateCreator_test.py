import os
import tempfile
from pathlib import Path

from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook

from otlmow_template.CsvTemplateCreator import CsvTemplateCreator
from otlmow_template.ExcelTemplateCreator import ExcelTemplateCreator
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
    db_location = Path(ROOT_DIR) / 'OTL_AllCasesTestClass.db'
    list_of_filter_uri = ['https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#AllCasesTestClass']
    filtered = SubsetTemplateCreator.filters_assets_by_subset(db_location, list_of_otl_objectUri=list_of_filter_uri)
    assert len(filtered) == 1
    assert filtered[0].name == 'AllCasesTestClass'


def test_empty_filter_list_removes_all_entries():
    db_location = Path(ROOT_DIR) / 'OTL_AllCasesTestClass.db'
    list_of_filter_uri = []
    filtered = SubsetTemplateCreator.filters_assets_by_subset(db_location, list_of_otl_objectUri=list_of_filter_uri)
    assert len(filtered) == 0


def test_no_filter_list_returns_all_entries():
    db_location = Path(ROOT_DIR) / 'OTL_AllCasesTestClass.db'
    filtered = SubsetTemplateCreator.filters_assets_by_subset(db_location)
    assert len(filtered) == 5


def test_return_temp_path_returns_path_to_temporary_file():
    path_to_template_file_and_extension = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.xlsx'
    temp_path = SubsetTemplateCreator.return_temp_path(path_to_template_file_and_extension)
    assert temp_path == Path(tempfile.gettempdir()) / 'temp-otlmow' / 'template_file_text.xlsx'


def test_xlsx_geo_artefact_column_is_removed_when_present():
    wb = Workbook()
    ws = wb.active
    ws.append(['geometry', 'test', 'test'])
    ws.append(['geotest', 'test', 'test'])
    ExcelTemplateCreator.remove_geo_artefact_excel(workbook=wb)
    assert ws.max_column == 2


def test_xlsx_no_column_removed_removed_when_no_geo_artefact_present():
    wb = Workbook()
    ws = wb.active
    ws.append(['test', 'test', 'test'])
    ws.append(['test', 'test', 'test'])
    ExcelTemplateCreator.remove_geo_artefact_excel(workbook=wb)
    assert ws.max_column == 3


def test_xlsx_find_uri_returns_uri():
    wb = Workbook()
    ws = wb.active
    ws.append(['typeURI', 'test', 'test'])
    ws.append(['eggshells', 'and', 'bombshells'])
    uri = ExcelTemplateCreator().find_uri_in_sheet(sheet=ws)
    assert uri == "eggshells"


def test_xlsx_find_uri_returns_none_if_no_uri_present():
    wb = Workbook()
    ws = wb.active
    ws.append(['TypUrI', 'uriType', 'testURI'])
    ws.append(['eggshells', 'and', 'bombshells'])
    uri = ExcelTemplateCreator().find_uri_in_sheet(sheet=ws)
    assert uri is None


def test_choice_list_options_get_added_to_seperate_sheet():
    wb = Workbook()
    wb.create_sheet('Keuzelijsten')
    options = ['test1', 'test2', 'test3']
    name = 'options_test'
    choice_list_dict = {}
    choice_list_dict = ExcelTemplateCreator().add_choice_list_to_sheet(workbook=wb, choice_list_dict=choice_list_dict,
                                                                       name=name, options=options)
    assert choice_list_dict[name] == "A"
    ws = wb['Keuzelijsten']
    assert ws['A1'].value == 'options_test'
    assert ws['A2'].value == 'test1'
    assert ws['A3'].value == 'test2'
    assert ws['A4'].value == 'test3'


def test_return_column_letter_returns_correct_letter():
    wb = Workbook()
    wb.create_sheet('Keuzelijsten')
    options = ['test1', 'test2', 'test3']
    name = 'options_test'
    choice_list_dict = {}
    column = ExcelTemplateCreator().return_column_letter_of_choice_list(workbook=wb, choice_list_dict=choice_list_dict,
                                                                        name=name, options=options)
    assert column == "A"


def test_return_column_letter_returns_correct_letter_if_column_already_exists():
    wb = Workbook()
    wb.create_sheet('Keuzelijsten')
    options = ['test1', 'test2', 'test3']
    name = 'options_test'
    choice_list_dict = {"options_test": "A", name: "A"}
    column = ExcelTemplateCreator().return_column_letter_of_choice_list(workbook=wb, choice_list_dict=choice_list_dict,
                                                                        name=name, options=options)
    assert column == "A"
