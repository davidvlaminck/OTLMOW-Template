import ntpath
import os
import site
import tempfile
from pathlib import Path
from typing import List
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.dimensions import DimensionHolder, ColumnDimension
from otlmow_converter.DotnotationHelper import DotnotationHelper

from otlmow_converter.OtlmowConverter import OtlmowConverter
from otlmow_model.Helpers.AssetCreator import dynamic_create_instance_from_uri
from otlmow_modelbuilder.DatatypeBuilderFunctions import get_type_link_from_attribuut, get_single_field_from_type_uri
from otlmow_modelbuilder.OSLOCollector import OSLOCollector
from otlmow_modelbuilder.OTLEnumerationCreator import OTLEnumerationCreator

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

enumeration_validation_rules = {
    "valid_uri_and_types": {},
    "valid_regexes": [
        "^https://wegenenverkeer.data.vlaanderen.be/ns/.+"]
}


class SubsetTemplateCreator:
    def __init__(self):
        pass

    @staticmethod
    def _load_collector_from_subset_path(path_to_subset: Path) -> OSLOCollector:
        collector = OSLOCollector(path_to_subset)
        collector.collect_all(include_abstract=True)
        return collector

    def generate_template_from_subset(self, path_to_subset: Path, path_to_template_file_and_extension: Path,
                                      **kwargs):
        tempdir = tempfile.gettempdir()
        test = ntpath.basename(path_to_template_file_and_extension)
        temporary_path = Path(tempdir) / test
        instantiated_attributes = self.generate_basic_template(path_to_subset=path_to_subset, temp_path = temporary_path,
                                                               path_to_template_file_and_extension=path_to_template_file_and_extension,
                                                               **kwargs)
        extension = os.path.splitext(path_to_template_file_and_extension)[-1].lower()
        tempdir = tempfile.gettempdir()
        test = ntpath.basename(path_to_template_file_and_extension)
        temporary_path = Path(tempdir) / test
        print(test)
        if extension == '.xlsx':
            self.alter_excel_template(path_to_template_file_and_extension=path_to_template_file_and_extension, temp_path= temporary_path,
                                      path_to_subset=path_to_subset, instantiated_attributes=instantiated_attributes,
                                      **kwargs)

    def generate_basic_template(self, path_to_subset, path_to_template_file_and_extension, temp_path: Path, **kwargs):
        collector = self._load_collector_from_subset_path(path_to_subset=path_to_subset)
        otl_objects = []

        for class_object in list(filter(lambda cl: cl.abstract == 0, collector.classes)):
            model_directory = None
            if kwargs is not None:
                model_directory = kwargs.get('model_directory', None)
            instance = dynamic_create_instance_from_uri(class_object.objectUri, model_directory=model_directory)
            if instance is None:
                continue
            instance.fill_with_dummy_data()
            otl_objects.append(instance)

            attributen = collector.find_attributes_by_class(class_object)
            for attribute_object in attributen:
                attr = getattr(instance, '_' + attribute_object.name)
                attr.fill_with_dummy_data()
        converter = OtlmowConverter()
        converter.create_file_from_assets(filepath=temp_path,
                                          list_of_objects=otl_objects, **kwargs)
        path_is_split = kwargs.get('split_per_type', True)
        extension = os.path.splitext(path_to_template_file_and_extension)[-1].lower()
        instantiated_attributes = []
        if path_is_split is False or extension == '.xlsx':
            instantiated_attributes = converter.create_assets_from_file(filepath=temp_path,
                                                                        path_to_subset=path_to_subset)
        return instantiated_attributes

    # TODO: Verschillende methodes voor verschillende documenten excel, csv
    @classmethod
    def alter_excel_template(cls, path_to_template_file_and_extension: Path, path_to_subset: Path,
                             instantiated_attributes: List, temp_path: Path, **kwargs):
        generate_choice_list = kwargs.get('generate_choice_list', False)
        add_geo_artefact = kwargs.get('add_geo_artefact', False)
        add_attribute_info = kwargs.get('add_attribute_info', False)
        highlight_deprecated_attributes = kwargs.get('highlight_deprecated_attributes', False)
        amount_of_examples = kwargs.get('amount_of_examples', 0)
        wb = load_workbook(temp_path)
        # Volgorde is belangrijk! Eerst rijen verwijderen indien nodig dan choice list toevoegen,
        # staat namelijk vast op de kolom en niet het attribuut in die kolom
        if add_geo_artefact is False:
            cls.remove_geo_artefact_excel(workbook=wb)
        if generate_choice_list:
            cls.add_choice_list_excel(workbook=wb, instantiated_attributes=instantiated_attributes,
                                      path_to_subset=path_to_subset)
        cls.add_mock_data_excel(workbook=wb, rows_of_examples=amount_of_examples)
        if highlight_deprecated_attributes:
            cls.check_for_deprecated_attributes(workbook=wb, instantiated_attributes=instantiated_attributes)
        if add_attribute_info:
            cls.add_attribute_info_excel(workbook=wb, instantiated_attributes=instantiated_attributes)
        cls.design_workbook_excel(workbook=wb)
        wb.save(path_to_template_file_and_extension)
        [f.unlink() for f in Path(temp_path).glob("*") if f.is_file()]

    def alter_single_csv_template(self, path_to_template_file_and_extension: Path, path_to_subset: Path,
                                  instantiated_attributes: List, **kwargs):
        add_geo_artefact = kwargs.get('add_geo_artefact', False)
        add_attribute_info = kwargs.get('add_attribute_info', False)
        highlight_deprecated_attributes = kwargs.get('highlight_deprecated_attributes', False)
        amount_of_examples = kwargs.get('amount_of_examples', 0)
        if add_geo_artefact is False:
            self.remove_geo_artefact_csv(path_to_template_file_and_extension=path_to_template_file_and_extension)
        pass

    @classmethod
    def filters_assets_by_subset(cls, path_to_subset: Path, **kwargs):
        list_of_otl_objectUri = kwargs.get('list_of_otl_objectUri', [])
        collector = cls._load_collector_from_subset_path(path_to_subset=path_to_subset)
        filtered_list = [x for x in collector.classes if x.objectUri in list_of_otl_objectUri]
        return filtered_list

    @staticmethod
    def _try_getting_settings_of_converter() -> Path:
        converter_path = Path(site.getsitepackages()[0]) / 'otlmow_converter'
        return converter_path / 'settings_otlmow_converter.json'

    @classmethod
    def design_workbook_excel(cls, workbook):
        for sheet in workbook:
            dim_holder = DimensionHolder(worksheet=sheet)
            for col in range(sheet.min_column, sheet.max_column + 1):
                dim_holder[get_column_letter(col)] = ColumnDimension(sheet, min=col, max=col, width=20)
            sheet.column_dimensions = dim_holder

    @classmethod
    def add_attribute_info_excel(cls, workbook, instantiated_attributes: List):
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            single_attribute = next(x for x in instantiated_attributes if x.typeURI == filter_uri)
            sheet.insert_rows(1)
            for rows in sheet.iter_rows(min_row=2, max_row=2, min_col=1):
                for cell in rows:
                    if cell.value == 'typeURI':
                        value = 'De URI van het object volgens https://www.w3.org/2001/XMLSchema#anyURI .'
                    elif cell.value.find('[DEPRECATED]') != -1:
                        strip = cell.value.split(' ')
                        dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute,
                                                                                                strip[1])
                        value = dotnotation_attribute.definition
                    else:
                        dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute,
                                                                                                cell.value)
                        value = dotnotation_attribute.definition

                    sheet.cell(row=1, column=cell.column, value=value)
                    sheet.cell(row=1, column=cell.column).fill = PatternFill(start_color="808080", end_color="808080",
                                                                             fill_type="solid")

    @classmethod
    def check_for_deprecated_attributes(cls, workbook, instantiated_attributes: List):
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            single_attribute = next(x for x in instantiated_attributes if x.typeURI == filter_uri)
            for rows in sheet.iter_rows(min_row=1, max_row=1, min_col=2):
                for cell in rows:
                    is_deprecated = False
                    if cell.value.count('.') == 1:
                        dot_split = cell.value.split('.')
                        attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute,
                                                                                    dot_split[0])

                        if len(attribute.deprecated_version) > 0:
                            is_deprecated = True
                    dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute,
                                                                                            cell.value)
                    if len(dotnotation_attribute.deprecated_version) > 0:
                        is_deprecated = True

                    if is_deprecated:
                        cell.value = '[DEPRECATED] ' + cell.value

    @classmethod
    def find_uri_in_sheet(cls, sheet):
        filter_uri = None
        for row in sheet.iter_rows(min_row=1, max_row=1):
            for cell in row:
                if cell.value == 'typeURI':
                    row_index = cell.row
                    column_index = cell.column
                    filter_uri = sheet.cell(row=row_index + 1, column=column_index).value
        return filter_uri

    @classmethod
    # TODO: remove geometry row from excel
    def remove_geo_artefact_excel(cls, workbook):
        for sheet in workbook:
            for row in sheet.iter_rows(min_row=1, max_row=1):
                for cell in row:
                    if cell.value == 'geometry':
                        sheet.delete_cols(cell.column)

    @classmethod
    def add_choice_list_excel(cls, workbook, instantiated_attributes: List, path_to_subset: Path):
        collector = cls._load_collector_from_subset_path(path_to_subset=path_to_subset)
        creator = OTLEnumerationCreator(collector)
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            single_attribute = next(x for x in instantiated_attributes if x.typeURI == filter_uri)
            for rows in sheet.iter_rows(min_row=1, max_row=1, min_col=2):
                for cell in rows:
                    if cell.value.find('[DEPRECATED]') != -1:
                        strip = cell.value.split(' ')
                        dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute,
                                                                                                strip[1])
                    else:
                        dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute,
                                                                                                cell.value)
                    attributes = collector.attributes
                    single = next((x for x in attributes if x.objectUri == dotnotation_attribute.objectUri), None)
                    enums = collector.enumerations
                    enum_check = None
                    if single is not None:
                        enum_check = next((x for x in enums if x.objectUri == single.type), None)
                    if enum_check is not None:
                        choice_list = collector.find_enumeration_by_uri(enum_check.objectUri)

                        choicelist_values = creator.get_keuzelijstwaardes_by_name(choice_list.name)
                        list = []
                        for options in choicelist_values:
                            list.append(options.invulwaarde)
                        values = ','.join(list)
                        data_val = DataValidation(type="list", formula1=f'"{values}"', allowBlank=True)
                        sheet.add_data_validation(data_val)
                        data_val.add(f'{get_column_letter(cell.column)}2:{get_column_letter(cell.column)}1000')
                    # TODO: change how this works because it doesn't work for all cases
                    attribute = [x for x in attributes if x.label == dotnotation_attribute.label]
                    if len(attribute) == 0:
                        continue
                    else:
                        field_type = get_single_field_from_type_uri(attribute[0].type)
                        if field_type == 'BooleanField':
                            data_validation = DataValidation(type="list", formula1='"TRUE,FALSE,-"', allow_blank=True)
                            column = cell.column
                            sheet.add_data_validation(data_validation)
                            data_validation.add(f'{get_column_letter(column)}2:{get_column_letter(column)}1000')
                            sheet.add_data_validation(data_validation)

    @classmethod
    def add_mock_data_excel(cls, workbook, rows_of_examples: int):
        for sheets in workbook:
            mock_values = []
            for rows in sheets.iter_rows(min_row=2, max_row=2):
                for cell in rows:
                    mock_values.append(cell.value)
            if rows_of_examples == 0:
                for rows in sheets.iter_rows(min_row=2, max_row=2):
                    for cell in rows:
                        cell.value = ''
            else:
                for rows in sheets.iter_rows(min_row=2, max_row=rows_of_examples + 1):
                    for cell in rows:
                        cell.value = mock_values[cell.column - 1]

    def remove_geo_artefact_csv(self, path_to_template_file_and_extension):
        with open(path_to_template_file_and_extension, 'w') as file:
            print(file)
        pass


if __name__ == '__main__':
    subset_tool = SubsetTemplateCreator()
    subset_location = Path(ROOT_DIR) / 'UnitTests' / 'Subset' / 'Flitspaal_noAgent3.0.db'
    # directory = Path(ROOT_DIR) / 'UnitTests' / 'TestClasses'
    # Slash op het einde toevoegen verandert weinig of niks aan het resultaat
    # directory = os.path.join(directory, '')
    xls_location = Path(ROOT_DIR) / 'UnitTests' / 'Subset' / 'testFileStorage' / 'template_file.xlsx'
    subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                              path_to_template_file_and_extension=xls_location, add_attribute_info=True,
                                              highlight_deprecated_attributes=True,
                                              amount_of_examples=5,
                                              generate_choice_list=True,
                                              )
