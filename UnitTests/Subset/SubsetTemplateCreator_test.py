import os
import tempfile
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook import Workbook

from otlmow_template.ExcelTemplateCreator import ExcelTemplateCreator
from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
model_directory_path = Path(__file__).parent.parent / 'TestModel'

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


def test_subset_with_AllCasesTestClass_no_double_kard_excel():
    subset_tool = SubsetTemplateCreator()
    excel_path = Path(ROOT_DIR) / 'testFileStorage' / 'OTL_AllCasesTestClass_no_double_kard.xlsx'
    subset_tool.generate_template_from_subset(path_to_subset=Path(ROOT_DIR) / 'OTL_AllCasesTestClass_no_double_kard.db',
                                              path_to_template_file_and_extension=excel_path, amount_of_examples=1,
                                              split_per_type=True, model_directory=model_directory_path,
                                              add_geo_artefact=True)
    assert excel_path.exists()

    book = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    header_row_list = []
    for sheet in book.worksheets:
        sheet_name = sheet.title
        if sheet_name != 'onderdeel#AllCasesTestClass':
            continue
        for row in sheet.rows:
            header_row_list = [cell.value for cell in row]
            break
    book.close()
    
    union_headers = [header for header in header_row_list if header.startswith('testUnionType')]
    header_row_list = [header for header in header_row_list if not header.startswith('testUnionType')]

    assert header_row_list == ['typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'bestekPostNummer[]',
        'datumOprichtingObject',  'geometry', 'isActief', 'notitie', 'standaardBestekPostNummer[]',
        'testBooleanField', 'testComplexType.testBooleanField',
        'testComplexType.testComplexType2.testKwantWrd',
        'testComplexType.testComplexType2.testStringField',
        'testComplexType.testComplexType2MetKard[].testKwantWrd',
        'testComplexType.testComplexType2MetKard[].testStringField',
        'testComplexType.testKwantWrd',
        'testComplexType.testKwantWrdMetKard[]',
        'testComplexType.testStringField',
        'testComplexType.testStringFieldMetKard[]',
        'testComplexTypeMetKard[].testBooleanField', 'testComplexTypeMetKard[].testComplexType2.testKwantWrd',
        'testComplexTypeMetKard[].testComplexType2.testStringField',
        'testComplexTypeMetKard[].testKwantWrd',
        'testComplexTypeMetKard[].testStringField',
        'testDateField', 'testDateTimeField', 'testDecimalField',
        'testDecimalFieldMetKard[]', 'testEenvoudigType', 'testEenvoudigTypeMetKard[]',
        'testIntegerField', 'testIntegerFieldMetKard[]', 'testKeuzelijst',
        'testKeuzelijstMetKard[]', 'testKwantWrd', 'testKwantWrdMetKard[]',
        'testStringField', 'testStringFieldMetKard[]', 'testTimeField', 'theoretischeLevensduur', 'toestand']

    assert union_headers[0].startswith('testUnionType.')
    assert union_headers[1].startswith('testUnionTypeMetKard[].')

    path = Path(ROOT_DIR) / 'testFileStorage'
    [f.unlink() for f in Path(path).glob("*") if f.is_file()]
    open(Path(ROOT_DIR) / 'testFileStorage' / '__init__.py', 'a').close()


@pytest.mark.asyncio(scope="module")
def test_subset_with_AllCasesTestClass_no_double_kard_excel_async():
    subset_tool = SubsetTemplateCreator()
    excel_path = Path(ROOT_DIR) / 'testFileStorage' / 'OTL_AllCasesTestClass_no_double_kard.xlsx'
    subset_tool.generate_template_from_subset(
        path_to_subset=Path(ROOT_DIR) / 'OTL_AllCasesTestClass_no_double_kard.db', amount_of_examples=1,
        path_to_template_file_and_extension=excel_path, plit_per_type=True, model_directory=model_directory_path,
        add_geo_artefact=True, filter_attributes_by_subset=True, generate_choice_list=True)
    assert excel_path.exists()

    book = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    header_row_list = []
    for sheet in book.worksheets:
        sheet_name = sheet.title
        if sheet_name != 'onderdeel#AllCasesTestClass':
            continue
        for row in sheet.rows:
            header_row_list = [cell.value for cell in row]
            break
    book.close()

    union_headers = [header for header in header_row_list if header.startswith('testUnionType')]
    header_row_list = [header for header in header_row_list if not header.startswith('testUnionType')]

    assert header_row_list == ['typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'bestekPostNummer[]',
                               'datumOprichtingObject', 'geometry', 'isActief', 'notitie',
                               'standaardBestekPostNummer[]',
                               'testBooleanField', 'testComplexType.testBooleanField',
                               'testComplexType.testComplexType2.testKwantWrd',
                               'testComplexType.testComplexType2.testStringField',
                               'testComplexType.testComplexType2MetKard[].testKwantWrd',
                               'testComplexType.testComplexType2MetKard[].testStringField',
                               'testComplexType.testKwantWrd',
                               'testComplexType.testKwantWrdMetKard[]',
                               'testComplexType.testStringField',
                               'testComplexType.testStringFieldMetKard[]',
                               'testComplexTypeMetKard[].testBooleanField',
                               'testComplexTypeMetKard[].testComplexType2.testKwantWrd',
                               'testComplexTypeMetKard[].testComplexType2.testStringField',
                               'testComplexTypeMetKard[].testKwantWrd',
                               'testComplexTypeMetKard[].testStringField',
                               'testDateField', 'testDateTimeField', 'testDecimalField',
                               'testDecimalFieldMetKard[]', 'testEenvoudigType', 'testEenvoudigTypeMetKard[]',
                               'testIntegerField', 'testIntegerFieldMetKard[]', 'testKeuzelijst',
                               'testKeuzelijstMetKard[]', 'testKwantWrd', 'testKwantWrdMetKard[]',
                               'testStringField', 'testStringFieldMetKard[]', 'testTimeField', 'theoretischeLevensduur',
                               'toestand']

    assert union_headers[0].startswith('testUnionType.')
    assert union_headers[1].startswith('testUnionTypeMetKard[].')

    path = Path(ROOT_DIR) / 'testFileStorage'
    [f.unlink() for f in Path(path).glob("*") if f.is_file()]
    open(Path(ROOT_DIR) / 'testFileStorage' / '__init__.py', 'a').close()


def test_subset_with_AllCasesTestClass_fewer_attributes_excel():
    subset_tool = SubsetTemplateCreator()
    excel_path = Path(ROOT_DIR) / 'testFileStorage' / 'OTL_AllCasesTestClass_fewer_attributes.xlsx'
    subset_tool.generate_template_from_subset(path_to_subset=Path(ROOT_DIR) / 'OTL_AllCasesTestClass_fewer_attributes.db',
                                              path_to_template_file_and_extension=excel_path, amount_of_examples=1,
                                              split_per_type=True, model_directory=model_directory_path)
    assert excel_path.exists()

    book = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    header_row_list = []
    for sheet in book.worksheets:
        sheet_name = sheet.title
        if sheet_name != 'onderdeel#AllCasesTestClass':
            continue
        for row in sheet.rows:
            header_row_list = [cell.value for cell in row]
            break
    book.close()

    union_headers = [header for header in header_row_list if header.startswith('testUnionType')]
    header_row_list = [header for header in header_row_list if not header.startswith('testUnionType')]

    assert header_row_list == [
        'typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'bestekPostNummer[]', 'datumOprichtingObject',
        'isActief', 'notitie', 'standaardBestekPostNummer[]', 'testBooleanField', 'testComplexType.testBooleanField',
        'testComplexType.testComplexType2.testKwantWrd', 'testComplexType.testComplexType2.testStringField',
        'testComplexType.testComplexType2MetKard[].testKwantWrd',
        'testComplexType.testComplexType2MetKard[].testStringField', 'testComplexType.testKwantWrd',
        'testComplexType.testKwantWrdMetKard[]', 'testComplexType.testStringField',
        'testComplexType.testStringFieldMetKard[]', 'testComplexTypeMetKard[].testBooleanField',
        'testComplexTypeMetKard[].testComplexType2.testKwantWrd',
        'testComplexTypeMetKard[].testComplexType2.testStringField', 'testComplexTypeMetKard[].testKwantWrd',
        'testComplexTypeMetKard[].testStringField', 'testEenvoudigType', 'testEenvoudigTypeMetKard[]', 'testKeuzelijst',
        'testKeuzelijstMetKard[]', 'testKwantWrd', 'testKwantWrdMetKard[]', 'theoretischeLevensduur', 'toestand']
    assert union_headers[0].startswith('testUnionType.')
    assert union_headers[1].startswith('testUnionTypeMetKard[].')

    path = Path(ROOT_DIR) / 'testFileStorage'
    [f.unlink() for f in Path(path).glob("*") if f.is_file()]
    open(Path(ROOT_DIR) / 'testFileStorage' / '__init__.py', 'a').close()


@pytest.mark.asyncio(scope="module")
async def test_subset_with_AllCasesTestClass_no_double_kard_csv_async():
    subset_tool = SubsetTemplateCreator()
    csv_location = Path(ROOT_DIR) / 'testFileStorage' / 'OTL_AllCasesTestClass_no_double_kard.csv'
    await subset_tool.generate_template_from_subset(
        path_to_subset=Path(ROOT_DIR) / 'OTL_AllCasesTestClass_no_double_kard.db', filter_attributes_by_subset=True,
        path_to_template_file_and_extension=csv_location, amount_of_examples=1, generate_choice_list=True,
        split_per_type=True, model_directory=model_directory_path)
    csv = Path(ROOT_DIR) / 'testFileStorage' / 'OTL_AllCasesTestClass_no_double_kard_onderdeel_AllCasesTestClass.csv'
    assert csv.exists()

    with open(csv, 'r') as f:
        header_row = f.readline()
    header_row_list = header_row.split(';')
    union_headers = [header for header in header_row_list if header.startswith('testUnionType')]
    print(union_headers)
    header_row_list = [header for header in header_row_list if not header.startswith('testUnionType')]

    assert header_row_list == ['typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'bestekPostNummer[]',
        'datumOprichtingObject', 'isActief', 'notitie', 'standaardBestekPostNummer[]',
        'testBooleanField', 'testComplexType.testBooleanField',
        'testComplexType.testComplexType2.testKwantWrd',
        'testComplexType.testComplexType2.testStringField',
        'testComplexType.testComplexType2MetKard[].testKwantWrd',
        'testComplexType.testComplexType2MetKard[].testStringField',
        'testComplexType.testKwantWrd',
        'testComplexType.testKwantWrdMetKard[]',
        'testComplexType.testStringField',
        'testComplexType.testStringFieldMetKard[]',
        'testComplexTypeMetKard[].testBooleanField',
        'testComplexTypeMetKard[].testComplexType2.testKwantWrd',
        'testComplexTypeMetKard[].testComplexType2.testStringField',
        'testComplexTypeMetKard[].testKwantWrd', 'testComplexTypeMetKard[].testStringField',
        'testDateField', 'testDateTimeField', 'testDecimalField',
        'testDecimalFieldMetKard[]', 'testEenvoudigType', 'testEenvoudigTypeMetKard[]',
        'testIntegerField', 'testIntegerFieldMetKard[]', 'testKeuzelijst',
        'testKeuzelijstMetKard[]', 'testKwantWrd', 'testKwantWrdMetKard[]',
        'testStringField', 'testStringFieldMetKard[]', 'testTimeField', 'theoretischeLevensduur', 'toestand\n']

    assert union_headers[0].startswith('testUnionType.')
    assert union_headers[1].startswith('testUnionTypeMetKard[].')

    path = Path(ROOT_DIR) / 'testFileStorage'
    [f.unlink() for f in Path(path).glob("*") if f.is_file()]
    open(Path(ROOT_DIR) / 'testFileStorage' / '__init__.py', 'a').close()


def test_subset_with_AllCasesTestClass_no_double_kard_csv():
    subset_tool = SubsetTemplateCreator()
    csv_location = Path(ROOT_DIR) / 'testFileStorage' / 'OTL_AllCasesTestClass_no_double_kard.csv'
    subset_tool.generate_template_from_subset(path_to_subset=Path(ROOT_DIR) / 'OTL_AllCasesTestClass_no_double_kard.db',
                                              path_to_template_file_and_extension=csv_location, amount_of_examples=1,
                                              split_per_type=True, model_directory=model_directory_path)
    csv = Path(ROOT_DIR) / 'testFileStorage' / 'OTL_AllCasesTestClass_no_double_kard_onderdeel_AllCasesTestClass.csv'
    assert csv.exists()

    with open(csv, 'r') as f:
        header_row = f.readline()
    header_row_list = header_row.split(';')
    union_headers = [header for header in header_row_list if header.startswith('testUnionType')]
    print(union_headers)
    header_row_list = [header for header in header_row_list if not header.startswith('testUnionType')]

    assert header_row_list == ['typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'bestekPostNummer[]',
        'datumOprichtingObject', 'isActief', 'notitie', 'standaardBestekPostNummer[]',
        'testBooleanField', 'testComplexType.testBooleanField',
        'testComplexType.testComplexType2.testKwantWrd',
        'testComplexType.testComplexType2.testStringField',
        'testComplexType.testComplexType2MetKard[].testKwantWrd',
        'testComplexType.testComplexType2MetKard[].testStringField',
        'testComplexType.testKwantWrd',
        'testComplexType.testKwantWrdMetKard[]',
        'testComplexType.testStringField',
        'testComplexType.testStringFieldMetKard[]',
        'testComplexTypeMetKard[].testBooleanField',
        'testComplexTypeMetKard[].testComplexType2.testKwantWrd',
        'testComplexTypeMetKard[].testComplexType2.testStringField',
        'testComplexTypeMetKard[].testKwantWrd', 'testComplexTypeMetKard[].testStringField',
        'testDateField', 'testDateTimeField', 'testDecimalField',
        'testDecimalFieldMetKard[]', 'testEenvoudigType', 'testEenvoudigTypeMetKard[]',
        'testIntegerField', 'testIntegerFieldMetKard[]', 'testKeuzelijst',
        'testKeuzelijstMetKard[]', 'testKwantWrd', 'testKwantWrdMetKard[]',
        'testStringField', 'testStringFieldMetKard[]', 'testTimeField', 'theoretischeLevensduur', 'toestand\n']

    assert union_headers[0].startswith('testUnionType.')
    assert union_headers[1].startswith('testUnionTypeMetKard[].')

    path = Path(ROOT_DIR) / 'testFileStorage'
    [f.unlink() for f in Path(path).glob("*") if f.is_file()]
    open(Path(ROOT_DIR) / 'testFileStorage' / '__init__.py', 'a').close()


def test_subset_actual_subset_excel():
    subset_tool = SubsetTemplateCreator()
    excel_path = Path(ROOT_DIR) / 'testFileStorage' / 'camera_steun.xlsx'
    subset_tool.generate_template_from_subset(path_to_subset=Path(ROOT_DIR) / 'camera_steun_2.14.db',
                                              path_to_template_file_and_extension=excel_path,
                                              split_per_type=True, amount_of_examples=1)

    book = openpyxl.load_workbook(excel_path, data_only=True)
    header_row_lists = []
    asset_versie_values = []
    for sheet in book.worksheets:
        sheet_name = sheet.title
        if sheet_name == 'Keuzelijsten':
            continue
        header_row_lists.append([r.value for r in next(sheet.rows)])
        asset_versie_values.append(sheet['D2'].value)
        asset_versie_values.append(sheet['E2'].value)
        asset_versie_values.append(sheet['F2'].value)
    book.close()

    assert asset_versie_values == [None, None, None, None, None, None]
    assert header_row_lists == [[
        'typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'assetVersie.context', 'assetVersie.timestamp',
        'assetVersie.versienummer', 'beeldverwerkingsinstelling[].configBestand.bestandsnaam',
        'beeldverwerkingsinstelling[].configBestand.mimeType',
        'beeldverwerkingsinstelling[].configBestand.omschrijving',
        'beeldverwerkingsinstelling[].configBestand.opmaakdatum', 'beeldverwerkingsinstelling[].configBestand.uri',
        'beeldverwerkingsinstelling[].typeBeeldverwerking', 'bestekPostNummer[]', 'datumOprichtingObject', 'dnsNaam',
        'heeftFlits', 'ipAdres', 'isActief', 'isPtz', 'merk', 'modelnaam', 'naam', 'notitie', 'opstelhoogte',
        'opstelwijze', 'rijrichting', 'serienummer', 'spectrum', 'standaardBestekPostNummer[]',
        'technischeFiche[].bestandsnaam', 'technischeFiche[].mimeType', 'technischeFiche[].omschrijving',
        'technischeFiche[].opmaakdatum', 'technischeFiche[].uri', 'theoretischeLevensduur', 'toestand'],
        ['typeURI', 'assetId.identificator', 'assetId.toegekendDoor', 'assetVersie.context', 'assetVersie.timestamp',
         'assetVersie.versienummer', 'beschermendeLaag.beschermlaag', 'bestekPostNummer[]', 'bijzonderTransport',
         'datumOprichtingObject', 'elektrischeBeveiliging', 'fabrikant', 'hoogteBovenkant', 'isActief', 'kleur', 'naam',
         'notitie', 'standaardBestekPostNummer[]', 'theoretischeLevensduur', 'toestand', 'type']]
    excel_path.unlink()


def test_subset_actual_subset_csv():
    subset_tool = SubsetTemplateCreator()
    csv_location = Path(ROOT_DIR) / 'testFileStorage' / 'camera_steun.csv'
    subset_tool.generate_template_from_subset(path_to_subset=Path(ROOT_DIR) / 'camera_steun_2.14.db',
                                              path_to_template_file_and_extension=csv_location,
                                              split_per_type=True)
    csv1 = Path(ROOT_DIR) / 'testFileStorage' / 'camera_steun_onderdeel_Bevestiging.csv'
    csv2 = Path(ROOT_DIR) / 'testFileStorage' / 'camera_steun_onderdeel_Camera.csv'
    csv3 = Path(ROOT_DIR) / 'testFileStorage' / 'camera_steun_onderdeel_RechteSteun.csv'
    assert not csv1.exists()
    assert csv2.exists()
    assert csv3.exists()

    subset_tool.generate_template_from_subset(path_to_subset=Path(ROOT_DIR) / 'camera_steun_2.14.db',
                                              path_to_template_file_and_extension=csv_location,
                                              split_per_type=True, ignore_relations=False)
    assert csv1.exists()
    assert csv2.exists()
    assert csv3.exists()

    path = Path(ROOT_DIR) / 'testFileStorage'
    [f.unlink() for f in Path(path).glob("*") if f.is_file()]
    open(Path(ROOT_DIR) / 'testFileStorage' / '__init__.py', 'a').close()


def test_filter_returns_filtered_list():
    db_location = Path(ROOT_DIR) / 'OTL_AllCasesTestClass.db'
    list_of_filter_uri = ['https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#AllCasesTestClass']
    filtered = SubsetTemplateCreator.filters_classes_by_subset(
        collector = SubsetTemplateCreator._load_collector_from_subset_path(db_location),
        list_of_otl_objectUri=list_of_filter_uri)
    assert len(filtered) == 1
    assert filtered[0].name == 'AllCasesTestClass'


def test_empty_filter_list_returns_all_entries():
    db_location = Path(ROOT_DIR) / 'OTL_AllCasesTestClass.db'
    list_of_filter_uri = []
    filtered = SubsetTemplateCreator.filters_classes_by_subset(
        collector = SubsetTemplateCreator._load_collector_from_subset_path(db_location),
        list_of_otl_objectUri=list_of_filter_uri)
    assert len(filtered) == 11


def test_no_filter_list_returns_all_entries():
    db_location = Path(ROOT_DIR) / 'OTL_AllCasesTestClass.db'
    filtered = SubsetTemplateCreator.filters_classes_by_subset(
        collector = SubsetTemplateCreator._load_collector_from_subset_path(db_location))
    assert len(filtered) == 11


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
