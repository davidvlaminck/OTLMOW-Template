import gc
from pathlib import Path

import openpyxl
import pytest

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

current_dir = Path(__file__).parent
model_directory_path = Path(__file__).parent.parent / 'TestModel'


def test_subset_with_AllCasesTestClass_excel_generic():
    subset_tool = SubsetTemplateCreator()
    excel_path = current_dir / 'OTL_AllCasesTestClass.xlsx'
    subset_tool.generate_template_from_subset(subset_path=current_dir / 'OTL_AllCasesTestClass.db',
                                              template_file_path=excel_path, model_directory=model_directory_path)
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

    gc.collect()

    excel_path.unlink()


@pytest.mark.asyncio
async def test_subset_with_AllCasesTestClass_excel_generic_async():
    subset_tool = SubsetTemplateCreator()
    excel_path = current_dir / 'async_OTL_AllCasesTestClass_async.xlsx'
    await subset_tool.generate_template_from_subset_async(subset_path=current_dir / 'OTL_AllCasesTestClass.db',
                                                          template_file_path=excel_path,
                                                          model_directory=model_directory_path)
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

    gc.collect()

    excel_path.unlink()


def test_subset_with_AllCasesTestClass_generic_csv():
    subset_tool = SubsetTemplateCreator()
    csv_path = current_dir / 'OTL_AllCasesTestClass.csv'
    subset_tool.generate_template_from_subset(subset_path=current_dir / 'OTL_AllCasesTestClass.db',
                                              template_file_path=csv_path, model_directory=model_directory_path)

    csv_allcases_path = csv_path.parent / csv_path.name.replace('.csv', '_onderdeel_AllCasesTestClass.csv')
    csv_another_path = csv_path.parent / csv_path.name.replace('.csv', '_onderdeel_AnotherTestClass.csv')
    csv_deprecated_path = csv_path.parent / csv_path.name.replace('.csv', '_onderdeel_DeprecatedTestClass.csv')

    assert csv_allcases_path.exists()

    with open(csv_allcases_path, 'r') as f:
        header_row = f.readline()
    header_row_list = header_row.split(';')

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
                               'toestand\n']

    assert union_headers[0].startswith('testUnionType.')
    assert union_headers[1].startswith('testUnionTypeMetKard[].')

    csv_allcases_path.unlink()
    csv_another_path.unlink()
    csv_deprecated_path.unlink()


@pytest.mark.asyncio
async def test_subset_with_AllCasesTestClass_generic_csv_async():
    subset_tool = SubsetTemplateCreator()
    csv_path = current_dir / 'generic_OTL_AllCasesTestClass_async.csv'
    await subset_tool.generate_template_from_subset_async(subset_path=current_dir / 'OTL_AllCasesTestClass.db',
                                                          template_file_path=csv_path,
                                                          model_directory=model_directory_path)

    csv_allcases_path = csv_path.parent / csv_path.name.replace('.csv', '_onderdeel_AllCasesTestClass.csv')
    csv_another_path = csv_path.parent / csv_path.name.replace('.csv', '_onderdeel_AnotherTestClass.csv')
    csv_deprecated_path = csv_path.parent / csv_path.name.replace('.csv', '_onderdeel_DeprecatedTestClass.csv')


    assert csv_allcases_path.exists()

    with open(csv_allcases_path, 'r') as f:
        header_row = f.readline()
    header_row_list = header_row.split(';')

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
                               'toestand\n']

    assert union_headers[0].startswith('testUnionType.')
    assert union_headers[1].startswith('testUnionTypeMetKard[].')

    csv_allcases_path.unlink()
    csv_another_path.unlink()
    csv_deprecated_path.unlink()
