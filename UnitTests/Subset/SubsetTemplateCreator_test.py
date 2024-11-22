from pathlib import Path

from UnitTests.TestClasses.OtlmowModel.Classes.Onderdeel.AllCasesTestClass import AllCasesTestClass
from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

ROOT_DIR = Path(__file__).parent
model_directory_path = ROOT_DIR.parent / 'TestModel'


def test_generate_template_from_subset_different_formats(subtests):
    subset_tool = SubsetTemplateCreator()
    subset_location = Path(ROOT_DIR) / 'Flitspaal_noAgent3.0.db'
    xls_location = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.xlsx'
    csv_location = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.csv'

    with subtests.test(msg='xls'):
        subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                  path_to_template_file_and_extension=xls_location)
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


def test_filter():
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


def test_generate_template_from_subset_test_classes():
    subset_tool = SubsetTemplateCreator()
    subset_location = Path(ROOT_DIR) / 'OTL_AllCasesTestClass.db'
    xls_location = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.xlsx'

    subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                              path_to_template_file_and_extension=xls_location,
                                              model_directory=model_directory_path)
    template_path = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.xlsx'
    assert template_path.exists()


    path = Path(ROOT_DIR) / 'testFileStorage'
    [f.unlink() for f in Path(path).glob("*") if f.is_file()]
    # Add an __init__.py file to the testFileStorage folder to make it a package
    open(Path(ROOT_DIR) / 'testFileStorage' / '__init__.py', 'a').close()


def test_clear_list_of_list_attributes():
    instance = AllCasesTestClass()
    instance.testStringField = 'test1'
    instance.testComplexType.testStringField = 'test2'
    instance.testDecimalFieldMetKard = [1.1]
    instance.testKwantWrdMetKard[0].waarde = 1.2
    instance.testComplexTypeMetKard[0].testStringField = 'test3'
    instance.testComplexTypeMetKard[0].testComplexType2.testStringField = 'test4'
    instance.testComplexTypeMetKard[0].testComplexType2MetKard[0].testStringField = 'test5'
    instance.testComplexTypeMetKard[0].testStringFieldMetKard = ['test6']

    SubsetTemplateCreator().clear_list_of_list_attributes(instance)

    assert instance.testStringField == 'test1'
    assert instance.testComplexType.testStringField == 'test2'
    assert instance.testDecimalFieldMetKard == [1.1]
    assert instance.testKwantWrdMetKard[0].waarde == 1.2
    assert instance.testComplexTypeMetKard[0].testStringField == 'test3'
    assert instance.testComplexTypeMetKard[0].testComplexType2.testStringField == 'test4'
    assert instance.testComplexTypeMetKard[0].testComplexType2MetKard[0].testStringField is None
    assert instance.testComplexTypeMetKard[0].testStringFieldMetKard is None
