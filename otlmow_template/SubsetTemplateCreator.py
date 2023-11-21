import logging
import os
import site
from pathlib import Path
from typing import List
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import DimensionHolder, ColumnDimension
from otlmow_converter.DotnotationHelper import DotnotationHelper

from otlmow_converter.OtlmowConverter import OtlmowConverter
from otlmow_model.Helpers.AssetCreator import dynamic_create_instance_from_uri
from otlmow_modelbuilder.OSLOCollector import OSLOCollector

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))


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
        collector = self._load_collector_from_subset_path(path_to_subset=path_to_subset)
        otl_objects = []

        for class_object in list(filter(lambda cl: cl.abstract == 0, collector.classes)):
            model_directory = None
            if kwargs is not None and 'model_directory' in kwargs:
                model_directory = kwargs['model_directory']
            instance = dynamic_create_instance_from_uri(class_object.objectUri, model_directory=model_directory)
            if instance is None:
                continue
            instance._assetId.fill_with_dummy_data()
            otl_objects.append(instance)

            attributen = collector.find_attributes_by_class(class_object)
            for attribute_object in attributen:
                attr = getattr(instance, '_' + attribute_object.name)
                attr.fill_with_dummy_data()
        converter = OtlmowConverter()
        converter.create_file_from_assets(filepath=path_to_template_file_and_extension,
                                          list_of_objects=otl_objects, **kwargs)
        path_is_split = kwargs.get('split_per_type', False)
        instantiated_attributes = []
        if not path_is_split:
            instantiated_attributes = converter.create_assets_from_file(filepath=path_to_template_file_and_extension,
                                                                        path_to_subset=path_to_subset)
        self.alter_template(changes=kwargs, path_to_template_file_and_extension=path_to_template_file_and_extension,
                            path_to_subset=path_to_subset, instantiated_attributes=instantiated_attributes)

    @classmethod
    def alter_template(cls, changes, path_to_template_file_and_extension: Path, path_to_subset: Path,
                       instantiated_attributes: List):
        # use **kwargs to pass changes
        generate_choice_list = changes.get('generate_choice_list', False)
        add_geo_artefact = changes.get('add_geo_artefact', False)
        add_attribute_info = changes.get('add_attribute_info', False)
        highlight_deprecated_attributes = changes.get('highlight_deprecated_attributes', False)
        amount_of_examples = changes.get('amount_of_examples', 0)
        if generate_choice_list:
            raise NotImplementedError("generate_choice_list is not implemented yet")
        if add_geo_artefact:
            raise NotImplementedError("add_geo_artefact is not implemented yet")
        if amount_of_examples > 0:
            raise NotImplementedError("amount_of_examples is not implemented yet")
        if highlight_deprecated_attributes:
            cls.check_for_deprecated_attributes(path_to_workbook=path_to_template_file_and_extension,
                                                instantiated_attributes=instantiated_attributes,
                                                path_to_subset=path_to_subset)
        if add_attribute_info:
            cls.add_attribute_info_excel(path_to_workbook=path_to_template_file_and_extension,
                                         instantiated_attributes=instantiated_attributes)

    @classmethod
    def filters_assets_by_subset(cls, path_to_subset: Path, list_of_otl_objectUri: List):
        collector = cls._load_collector_from_subset_path(path_to_subset=path_to_subset)
        filtered_list = [x for x in collector.classes if x.objectUri in list_of_otl_objectUri]
        return filtered_list

    @staticmethod
    def _try_getting_settings_of_converter() -> Path:
        converter_path = Path(site.getsitepackages()[0]) / 'otlmow_converter'
        return converter_path / 'settings_otlmow_converter.json'

    @classmethod
    def design_workbook_excel(cls, path_to_workbook: Path):
        wb = load_workbook(path_to_workbook)
        for sheet in wb:
            for rows in sheet.iter_rows(min_row=1, max_row=1):
                for cell in rows:
                    cell.fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
            dim_holder = DimensionHolder(worksheet=sheet)
            for col in range(sheet.min_column, sheet.max_column + 1):
                dim_holder[get_column_letter(col)] = ColumnDimension(sheet, min=col, max=col, width=20)
            sheet.column_dimensions = dim_holder
        wb.save(path_to_workbook)

    @classmethod
    def add_attribute_info_excel(cls, path_to_workbook: Path, instantiated_attributes: List):
        dotnotation_module = DotnotationHelper()
        workbook = load_workbook(path_to_workbook)
        for sheet in workbook:
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            single_attribute = [x for x in instantiated_attributes if x.typeURI == filter_uri]
            for rows in sheet.iter_rows(min_row=1, max_row=1, min_col=2):
                for cell in rows:
                    dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute[0],
                                                                                            cell.value)
                    cell.value = dotnotation_attribute.definition + "\n\n" + " " + cell.value
        workbook.save(path_to_workbook)

    @classmethod
    def check_for_deprecated_attributes(cls, path_to_workbook: Path, instantiated_attributes: List,
                                        path_to_subset: Path):
        dotnotation_module = DotnotationHelper()
        workbook = load_workbook(path_to_workbook)
        for sheet in workbook:
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            single_attribute = [x for x in instantiated_attributes if x.typeURI == filter_uri]
            for rows in sheet.iter_rows(min_row=1, max_row=1, min_col=2):
                for cell in rows:
                    is_deprecated = False
                    if cell.value.count('.') == 1:
                        dot_split = cell.value.split('.')
                        attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute[0],
                                                                                    dot_split[0])

                        if len(attribute.deprecated_version) > 0:
                            is_deprecated = True
                    dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute[0],
                                                                                            cell.value)
                    if len(dotnotation_attribute.deprecated_version) > 0:
                        is_deprecated = True

                    if is_deprecated:
                        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        workbook.save(path_to_workbook)

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


if __name__ == '__main__':
    subset_tool = SubsetTemplateCreator()
    subset_location = Path(ROOT_DIR) / 'UnitTests' / 'Subset' / 'Flitspaal_noAgent3.0.db'
    print(subset_location)
    xls_location = Path(ROOT_DIR) / 'UnitTests' / 'Subset' / 'testFileStorage' / 'template_file.xlsx'
    subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                              path_to_template_file_and_extension=xls_location, add_attribute_info=True,
                                              highlight_deprecated_attributes=True)
