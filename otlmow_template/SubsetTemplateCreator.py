import contextlib
import csv
import logging
import ntpath
import os
import site
import tempfile
from asyncio import sleep
from pathlib import Path


from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.dimensions import DimensionHolder, ColumnDimension
from otlmow_converter.DotnotationHelper import DotnotationHelper
from otlmow_converter.OtlmowConverter import OtlmowConverter
from otlmow_model.OtlmowModel.BaseClasses.BooleanField import BooleanField
from otlmow_model.OtlmowModel.BaseClasses.KeuzelijstField import KeuzelijstField
from otlmow_model.OtlmowModel.BaseClasses.OTLObject import dynamic_create_instance_from_uri
from otlmow_model.OtlmowModel.Helpers.generated_lists import get_hardcoded_relation_dict
from otlmow_modelbuilder.OSLOCollector import OSLOCollector
from otlmow_modelbuilder.SQLDataClasses.OSLOClass import OSLOClass

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

enumeration_validation_rules = {
    "valid_uri_and_types": {},
    "valid_regexes": [
        "^https://wegenenverkeer.data.vlaanderen.be/ns/.+"]
}


class SubsetTemplateCreator:
    @classmethod
    def _load_collector_from_subset_path(cls, subset_path: Path) -> OSLOCollector:
        collector = OSLOCollector(subset_path)
        collector.collect_all(include_abstract=True)
        return collector

    @classmethod
    def generate_template_from_subset(
            cls, subset_path: Path, template_file_path: Path, ignore_relations: bool = True,
            filter_attributes_by_subset: bool = True, class_uris_filter: [str] = None, **kwargs):

        """
        Generate a template from a subset file
        :param subset_path: Path to the subset file
        :param template_file_path: Path to where the template file should be created
        :param ignore_relations: Whether to ignore relations when creating the template
        :param filter_attributes_by_subset: Whether to filter by the attributes in the subset
        :param class_uris_filter: List of class URIs to filter by. If not None, only classes with these URIs will be included

        :return:
        """
        tempdir = Path(tempfile.gettempdir()) / 'temp-otlmow'
        if not tempdir.exists():
            os.makedirs(tempdir)

        temporary_path = Path(tempdir) / template_file_path.name
        instantiated_attributes = cls.generate_basic_template(
            subset_path=subset_path, temporary_path=temporary_path, ignore_relations=ignore_relations,
            template_file_path=template_file_path,
            filter_attributes_by_subset=filter_attributes_by_subset, **kwargs)

        # TODO split altering and moving the temp file to final location
        extension = template_file_path.suffix.lower()
        if extension == '.xlsx':
            cls.alter_excel_template(template_file_path=template_file_path,
                                      temporary_path=temporary_path,
                                      subset_path=subset_path, instantiated_attributes=instantiated_attributes,
                                      **kwargs)
        elif extension == '.csv':
            cls.determine_multiplicity_csv(
                template_file_path=template_file_path,
                                            subset_path=subset_path,
                                            instantiated_attributes=instantiated_attributes,
                                            temporary_path=temporary_path,
                                            **kwargs)

    @classmethod
    async def generate_template_from_subset_async(cls, subset_path: Path, template_file_path: Path,
                                      ignore_relations: bool = True, filter_attributes_by_subset: bool = True,
                                      **kwargs):
        tempdir = Path(tempfile.gettempdir()) / 'temp-otlmow'
        if not tempdir.exists():
            os.makedirs(tempdir)
        test = ntpath.basename(template_file_path)
        temporary_path = Path(tempdir) / test
        await sleep(0)
        instantiated_attributes = (await cls.generate_basic_template_async(
            subset_path=subset_path, temporary_path=temporary_path, ignore_relations=ignore_relations,
            template_file_path=template_file_path,
            filter_attributes_by_subset=filter_attributes_by_subset, **kwargs))
        await sleep(0)
        extension = os.path.splitext(template_file_path)[-1].lower()
        if extension == '.xlsx':
            await cls.alter_excel_template_async(
                template_file_path=template_file_path,
                                      temporary_path=temporary_path,
                                      subset_path=subset_path, instantiated_attributes=instantiated_attributes,
                                      **kwargs)
        elif extension == '.csv':
            await cls.determine_multiplicity_csv_async(
                template_file_path=template_file_path,
                                            subset_path=subset_path,
                                            instantiated_attributes=instantiated_attributes,
                                            temporary_path=temporary_path,
                                            **kwargs)

    @classmethod
    def generate_basic_template(cls, subset_path: Path, template_file_path: Path,
                                temporary_path: Path, ignore_relations: bool = True, **kwargs):
        class_uris_filter = None
        if kwargs is not None:
            class_uris_filter = kwargs.get('class_uris_filter', None)
        collector = cls._load_collector_from_subset_path(subset_path=subset_path)
        filtered_class_list = cls.filters_classes_by_subset(
            collector=collector, class_uris_filter=class_uris_filter)
        otl_objects = []
        amount_of_examples = kwargs.get('amount_of_examples', 0)
        model_directory = None
        if kwargs is not None:
            model_directory = kwargs.get('model_directory', None)
        relation_dict = get_hardcoded_relation_dict(model_directory=model_directory)

        generate_dummy_records = 1
        if amount_of_examples > 1:
            generate_dummy_records = amount_of_examples

        for class_object in [cl for cl in filtered_class_list if cl.abstract == 0]:
            if ignore_relations and class_object.objectUri in relation_dict:
                continue
            for _ in range(generate_dummy_records):
                instance = dynamic_create_instance_from_uri(class_object.objectUri, model_directory=model_directory)
                if instance is None:
                    continue
                attributen = collector.find_attributes_by_class(class_object)
                for attribute_object in attributen:
                    attr = getattr(instance, f'_{attribute_object.name}')
                    attr.fill_with_dummy_data()
                with contextlib.suppress(AttributeError):
                    geo_attr = getattr(instance, '_geometry')
                    geo_attr.fill_with_dummy_data()
                otl_objects.append(instance)

                DotnotationHelper.clear_list_of_list_attributes(instance)

        converter = OtlmowConverter()
        converter.from_objects_to_file(file_path=temporary_path, sequence_of_objects=otl_objects, **kwargs)
        path_is_split = kwargs.get('split_per_type', True)
        extension = os.path.splitext(template_file_path)[-1].lower()
        instantiated_attributes = []
        if path_is_split is False or extension == '.xlsx':
            instantiated_attributes = converter.from_file_to_objects(
                file_path=temporary_path, subset_path=subset_path)
        return list(instantiated_attributes)

    @classmethod
    async def generate_basic_template_async(cls, subset_path: Path, template_file_path: Path,
                                temporary_path: Path, ignore_relations: bool = True, **kwargs):
        class_uris_filter = None
        if kwargs is not None:
            class_uris_filter = kwargs.get('class_uris_filter', None)
        collector = cls._load_collector_from_subset_path(subset_path=subset_path)
        filtered_class_list = cls.filters_classes_by_subset(
            collector=collector, class_uris_filter=class_uris_filter)
        otl_objects = []
        amount_of_examples = kwargs.get('amount_of_examples', 0)
        model_directory = None
        if kwargs is not None:
            model_directory = kwargs.get('model_directory', None)
        relation_dict = get_hardcoded_relation_dict(model_directory=model_directory)

        generate_dummy_records = 1
        if amount_of_examples > 1:
            generate_dummy_records = amount_of_examples

        for class_object in [cl for cl in filtered_class_list if cl.abstract == 0]:
            if ignore_relations and class_object.objectUri in relation_dict:
                continue
            for _ in range(generate_dummy_records):
                instance = dynamic_create_instance_from_uri(class_object.objectUri, model_directory=model_directory)
                await sleep(0)
                if instance is None:
                    continue
                attributen = collector.find_attributes_by_class(class_object)
                for attribute_object in attributen:
                    attr = getattr(instance, f'_{attribute_object.name}')
                    attr.fill_with_dummy_data()
                with contextlib.suppress(AttributeError):
                    geo_attr = getattr(instance, '_geometry')
                    geo_attr.fill_with_dummy_data()
                otl_objects.append(instance)

                DotnotationHelper.clear_list_of_list_attributes(instance)

        await sleep(0)
        converter = OtlmowConverter()
        await converter.from_objects_to_file_async(file_path=temporary_path, sequence_of_objects=otl_objects, **kwargs)
        path_is_split = kwargs.get('split_per_type', True)
        extension = os.path.splitext(template_file_path)[-1].lower()
        instantiated_attributes = []
        if path_is_split is False or extension == '.xlsx':
            instantiated_attributes = await converter.from_file_to_objects_async(
                file_path=temporary_path, subset_path=subset_path)
        return list(instantiated_attributes)

    @classmethod
    def alter_excel_template(cls, template_file_path: Path, subset_path: Path,
                             instantiated_attributes: list, temporary_path, **kwargs):
        generate_choice_list = kwargs.get('generate_choice_list', False)
        add_geo_artefact = kwargs.get('add_geo_artefact', False)
        add_attribute_info = kwargs.get('add_attribute_info', False)
        highlight_deprecated_attributes = kwargs.get('highlight_deprecated_attributes', False)
        amount_of_examples = kwargs.get('amount_of_examples', 0)
        original_amount_of_examples = amount_of_examples
        if add_attribute_info and amount_of_examples == 0:
            amount_of_examples = 1
        wb = load_workbook(temporary_path)
        wb.create_sheet('Keuzelijsten')
        # Volgorde is belangrijk! Eerst rijen verwijderen indien nodig dan choice list toevoegen,
        # staat namelijk vast op de kolom en niet het attribuut in die kolom
        if add_geo_artefact is False:
            cls.remove_geo_artefact_excel(workbook=wb)
        if generate_choice_list:
            cls.add_choice_list_excel(workbook=wb, instantiated_attributes=instantiated_attributes,
                                      subset_path=subset_path, add_attribute_info=add_attribute_info)
        cls.add_mock_data_excel(workbook=wb, rows_of_examples=amount_of_examples) # remove dummy rows if needed

        cls.custom_exel_fixes(workbook=wb, instantiated_attributes=instantiated_attributes,
                                    add_attribute_info=add_attribute_info)
        if highlight_deprecated_attributes:
            cls.check_for_deprecated_attributes(workbook=wb, instantiated_attributes=instantiated_attributes)
        if add_attribute_info:
            cls.add_attribute_info_excel(workbook=wb, instantiated_attributes=instantiated_attributes)
        if original_amount_of_examples == 0 and add_attribute_info:
            cls.remove_examples_from_excel_again(workbook=wb)
        wb.save(template_file_path)
        file_location = os.path.dirname(temporary_path)
        [f.unlink() for f in Path(file_location).glob("*") if f.is_file()]


    @classmethod
    async def alter_excel_template_async(cls, template_file_path: Path, subset_path: Path,
                             instantiated_attributes: list, temporary_path, **kwargs):
        await sleep(0)
        generate_choice_list = kwargs.get('generate_choice_list', False)
        add_geo_artefact = kwargs.get('add_geo_artefact', False)
        add_attribute_info = kwargs.get('add_attribute_info', False)
        highlight_deprecated_attributes = kwargs.get('highlight_deprecated_attributes', False)
        amount_of_examples = kwargs.get('amount_of_examples', 0)
        original_amount_of_examples = amount_of_examples
        if add_attribute_info and amount_of_examples == 0:
            amount_of_examples = 1
        await sleep(0)
        wb = load_workbook(temporary_path)
        wb.create_sheet('Keuzelijsten')
        # Volgorde is belangrijk! Eerst rijen verwijderen indien nodig dan choice list toevoegen,
        # staat namelijk vast op de kolom en niet het attribuut in die kolom
        if add_geo_artefact is False:
            await sleep(0)
            cls.remove_geo_artefact_excel(workbook=wb)
        if generate_choice_list:
            await sleep(0)
            await cls.add_choice_list_excel_async(workbook=wb, instantiated_attributes=instantiated_attributes,
                                      subset_path=subset_path, add_attribute_info=add_attribute_info)
        await sleep(0)
        cls.add_mock_data_excel(workbook=wb, rows_of_examples=amount_of_examples) # remove dummy rows if needed

        await cls.custom_exel_fixes_async(workbook=wb, instantiated_attributes=instantiated_attributes,
                                    add_attribute_info=add_attribute_info)
        await sleep(0)
        if highlight_deprecated_attributes:
            await sleep(0)
            cls.check_for_deprecated_attributes(workbook=wb, instantiated_attributes=instantiated_attributes)
        if add_attribute_info:
            await sleep(0)
            await cls.add_attribute_info_excel_async(workbook=wb, instantiated_attributes=instantiated_attributes)
        if original_amount_of_examples == 0 and add_attribute_info:
            cls.remove_examples_from_excel_again(workbook=wb)
        await sleep(0)
        wb.save(template_file_path)
        file_location = os.path.dirname(temporary_path)
        [f.unlink() for f in Path(file_location).glob("*") if f.is_file()]

    @classmethod
    def determine_multiplicity_csv(cls, template_file_path: Path, subset_path: Path,
                                   instantiated_attributes: list, temporary_path: Path, **kwargs):
        path_is_split = kwargs.get('split_per_type', True)
        if path_is_split is False:
            cls.alter_csv_template(template_file_path=template_file_path,
                                    temporary_path=temporary_path, subset_path=subset_path, **kwargs)
        else:
            cls.multiple_csv_template(
                template_file_path=template_file_path,
                                       temporary_path=temporary_path,
                                       subset_path=subset_path, instantiated_attributes=instantiated_attributes,
                                       **kwargs)
        file_location = os.path.dirname(temporary_path)
        [f.unlink() for f in Path(file_location).glob("*") if f.is_file()]

    @classmethod
    async def determine_multiplicity_csv_async(cls, template_file_path: Path, subset_path: Path,
                                   instantiated_attributes: list, temporary_path: Path, **kwargs):
        path_is_split = kwargs.get('split_per_type', True)
        await sleep(0)
        if path_is_split is False:
            await cls.alter_csv_template_async(template_file_path=template_file_path,
                                    temporary_path=temporary_path, subset_path=subset_path, **kwargs)
        else:
            await cls.multiple_csv_template_async(
                template_file_path=template_file_path,
                                       temporary_path=temporary_path,
                                       subset_path=subset_path, instantiated_attributes=instantiated_attributes,
                                       **kwargs)
        file_location = os.path.dirname(temporary_path)
        [f.unlink() for f in Path(file_location).glob("*") if f.is_file()]

    @classmethod
    def filters_classes_by_subset(cls, collector: OSLOCollector,
                                  class_uris_filter: [str] = None) -> list[OSLOClass]:
        if class_uris_filter is None:
            return collector.classes
        return [x for x in collector.classes if x.objectUri in class_uris_filter]

    @classmethod
    def _try_getting_settings_of_converter(cls) -> Path:
        converter_path = Path(site.getsitepackages()[0]) / 'otlmow_converter'
        return converter_path / 'settings_otlmow_converter.json'

    @classmethod
    def add_type_uri_choice_list_in_excel(cls, sheet, instantiated_attributes, add_attribute_info: bool):
        starting_row = '3' if add_attribute_info else '2'
        if sheet.title == 'Keuzelijsten':
            return
        type_uri_found = False
        for row in sheet.iter_rows(min_row=1, max_row=1):
            for cell in row:
                if cell.value == 'typeURI':
                    type_uri_found = True
                    break
            if type_uri_found:
                break
        if not type_uri_found:
            return

        sheet_name = sheet.title
        type_uri = ''
        if sheet_name.startswith('http'):
            type_uri = sheet_name
        else:
            split_name = sheet_name.split("#")
            subclass_name = split_name[1]

            possible_classes = [x for x in instantiated_attributes if x.typeURI.endswith(subclass_name)]
            if len(possible_classes) == 1:
                type_uri = possible_classes[0].typeURI

        if type_uri == '':
            return

        data_validation = DataValidation(type="list", formula1=f'"{type_uri}"', allow_blank=True)
        for rows in sheet.iter_rows(min_row=1, max_row=1, min_col=1, max_col=1):
            for cell in rows:
                column = cell.column
                sheet.add_data_validation(data_validation)
                data_validation.add(f'{get_column_letter(column)}{starting_row}:{get_column_letter(column)}1000')

    @classmethod
    async def add_type_uri_choice_list_in_excel_async(cls, sheet, instantiated_attributes, add_attribute_info: bool):
        starting_row = '3' if add_attribute_info else '2'
        await sleep(0)
        if sheet.title == 'Keuzelijsten':
            return
        type_uri_found = False
        for row in sheet.iter_rows(min_row=1, max_row=1):
            for cell in row:
                if cell.value == 'typeURI':
                    type_uri_found = True
                    break
            if type_uri_found:
                break
        if not type_uri_found:
            return

        await sleep(0)
        sheet_name = sheet.title
        type_uri = ''
        if sheet_name.startswith('http'):
            type_uri = sheet_name
        else:
            split_name = sheet_name.split("#")
            subclass_name = split_name[1]

            possible_classes = [x for x in instantiated_attributes if x.typeURI.endswith(subclass_name)]
            if len(possible_classes) == 1:
                type_uri = possible_classes[0].typeURI

        if type_uri == '':
            return

        data_validation = DataValidation(type="list", formula1=f'"{type_uri}"', allow_blank=True)
        await sleep(0)
        for rows in sheet.iter_rows(min_row=1, max_row=1, min_col=1, max_col=1):
            for cell in rows:
                await sleep(0)
                column = cell.column
                sheet.add_data_validation(data_validation)
                data_validation.add(f'{get_column_letter(column)}{starting_row}:{get_column_letter(column)}1000')

    @classmethod
    def custom_exel_fixes(cls, workbook, instantiated_attributes, add_attribute_info: bool):
        for sheet in workbook:
            cls.set_fixed_column_width(sheet=sheet, width=25)
            cls.add_type_uri_choice_list_in_excel(sheet=sheet, instantiated_attributes=instantiated_attributes,
                                                        add_attribute_info=add_attribute_info)
            cls.remove_asset_versie(sheet=sheet)

    @classmethod
    async def custom_exel_fixes_async(cls, workbook, instantiated_attributes, add_attribute_info: bool):
        for sheet in workbook:
            await sleep(0)
            await cls.set_fixed_column_width_async(sheet=sheet, width=25)
            await cls.add_type_uri_choice_list_in_excel_async(sheet=sheet,
                                                            instantiated_attributes=instantiated_attributes,
                                                        add_attribute_info=add_attribute_info)
            await cls.remove_asset_versie_async(sheet=sheet)

    @classmethod
    def remove_asset_versie(cls, sheet):
        for row in sheet.iter_rows(min_row=1, max_row=1, min_col=4):
            for cell in row:
                if cell.value is None or not cell.value.startswith('assetVersie'):
                    continue
                for rows in sheet.iter_rows(min_col=cell.column, max_col=cell.column, min_row=2, max_row=1000):
                    for c in rows:
                        c.value = ''

    @classmethod
    async def remove_asset_versie_async(cls, sheet):
        for row in sheet.iter_rows(min_row=1, max_row=1, min_col=4):
            for cell in row:
                await sleep(0)
                if cell.value is None or not cell.value.startswith('assetVersie'):
                    continue
                for rows in sheet.iter_rows(min_col=cell.column, max_col=cell.column, min_row=2, max_row=1000):
                    for c in rows:
                        await sleep(0)
                        c.value = ''

    @classmethod
    def set_fixed_column_width(cls, sheet, width: int):
        dim_holder = DimensionHolder(worksheet=sheet)
        for col in range(sheet.min_column, sheet.max_column + 1):
            dim_holder[get_column_letter(col)] = ColumnDimension(sheet, min=col, max=col, width=width)
        sheet.column_dimensions = dim_holder

    @classmethod
    async def set_fixed_column_width_async(cls, sheet, width: int):
        dim_holder = DimensionHolder(worksheet=sheet)
        for col in range(sheet.min_column, sheet.max_column + 1):
            await sleep(0)
            dim_holder[get_column_letter(col)] = ColumnDimension(sheet, min=col, max=col, width=width)
        sheet.column_dimensions = dim_holder

    @classmethod
    def add_attribute_info_excel(cls, workbook, instantiated_attributes: list):
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            if sheet == workbook['Keuzelijsten']:
                break
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
                    newcell = sheet.cell(row=1, column=cell.column, value=value)
                    newcell.alignment = Alignment(wrapText=True, vertical='top')
                    newcell.fill = PatternFill(start_color="808080", end_color="808080",
                                               fill_type="solid")

    @classmethod
    async def add_attribute_info_excel_async(cls, workbook, instantiated_attributes: list):
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            if sheet == workbook['Keuzelijsten']:
                break
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            await sleep(0)
            single_attribute = next(x for x in instantiated_attributes if x.typeURI == filter_uri)
            sheet.insert_rows(1)
            for rows in sheet.iter_rows(min_row=2, max_row=2, min_col=1):
                for cell in rows:
                    await sleep(0)
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
                    newcell = sheet.cell(row=1, column=cell.column, value=value)
                    newcell.alignment = Alignment(wrapText=True, vertical='top')
                    newcell.fill = PatternFill(start_color="808080", end_color="808080",
                                               fill_type="solid")

    @classmethod
    def check_for_deprecated_attributes(cls, workbook, instantiated_attributes: list):
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            if sheet == workbook['Keuzelijsten']:
                break
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
                        cell.value = f'[DEPRECATED] {cell.value}'

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
    def remove_geo_artefact_excel(cls, workbook):
        for sheet in workbook:
            for row in sheet.iter_rows(min_row=1, max_row=1):
                for cell in row:
                    if cell.value == 'geometry':
                        sheet.delete_cols(cell.column)

    @classmethod
    def add_choice_list_excel(cls, workbook, instantiated_attributes: list, subset_path: Path,
                                    add_attribute_info: bool=False):
        choice_list_dict = {}
        starting_row = '3' if add_attribute_info else '2'
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            if sheet == workbook['Keuzelijsten']:
                break
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
                    if issubclass(dotnotation_attribute.field, KeuzelijstField):
                        name = dotnotation_attribute.field.naam
                        valid_options = [v.invulwaarde for k, v in dotnotation_attribute.field.options.items()
                                         if v.status != 'verwijderd']
                        if (dotnotation_attribute.field.naam not in choice_list_dict):
                            choice_list_dict = cls.add_choice_list_to_sheet(
                                workbook=workbook, name=name,  options=valid_options, choice_list_dict=choice_list_dict)
                        column = choice_list_dict[dotnotation_attribute.field.naam]
                        start_range = f"${column}$2"
                        end_range = f"${column}${len(valid_options) + 1}"
                        data_val = DataValidation(type="list", formula1=f"Keuzelijsten!{start_range}:{end_range}",
                                                  allowBlank=True)
                        sheet.add_data_validation(data_val)
                        data_val.add(f'{get_column_letter(cell.column)}{starting_row}:'
                                     f'{get_column_letter(cell.column)}1000')

                    if issubclass(dotnotation_attribute.field, BooleanField):
                        data_validation = DataValidation(type="list", formula1='"TRUE,FALSE,"', allow_blank=True)
                        column = cell.column
                        sheet.add_data_validation(data_validation)
                        data_validation.add(f'{get_column_letter(column)}{starting_row}:'
                                            f'{get_column_letter(column)}1000')

    @classmethod
    async def add_choice_list_excel_async(cls, workbook, instantiated_attributes: list, subset_path: Path,
                                    add_attribute_info: bool = False):
        choice_list_dict = {}
        starting_row = '3' if add_attribute_info else '2'
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            await sleep(0)
            if sheet == workbook['Keuzelijsten']:
                break
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            single_attribute = next(x for x in instantiated_attributes if x.typeURI == filter_uri)
            for rows in sheet.iter_rows(min_row=1, max_row=1, min_col=2):
                for cell in rows:
                    await sleep(0)
                    if cell.value.find('[DEPRECATED]') != -1:
                        strip = cell.value.split(' ')
                        dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute,
                                                                                                strip[1])
                    else:
                        dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(single_attribute,
                                                                                                cell.value)
                    await sleep(0)
                    if issubclass(dotnotation_attribute.field, KeuzelijstField):
                        name = dotnotation_attribute.field.naam
                        valid_options = [v.invulwaarde for k, v in dotnotation_attribute.field.options.items()
                                         if v.status != 'verwijderd']
                        if (dotnotation_attribute.field.naam not in choice_list_dict):
                            choice_list_dict = cls.add_choice_list_to_sheet(
                                workbook=workbook, name=name, options=valid_options, choice_list_dict=choice_list_dict)
                        column = choice_list_dict[dotnotation_attribute.field.naam]
                        start_range = f"${column}$2"
                        end_range = f"${column}${len(valid_options) + 1}"
                        data_val = DataValidation(type="list", formula1=f"Keuzelijsten!{start_range}:{end_range}",
                                                  allowBlank=True)
                        sheet.add_data_validation(data_val)
                        data_val.add(f'{get_column_letter(cell.column)}{starting_row}:'
                                     f'{get_column_letter(cell.column)}1000')

                    await sleep(0)
                    if issubclass(dotnotation_attribute.field, BooleanField):
                        data_validation = DataValidation(type="list", formula1='"TRUE,FALSE,"', allow_blank=True)
                        column = cell.column
                        sheet.add_data_validation(data_validation)
                        data_validation.add(f'{get_column_letter(column)}{starting_row}:'
                                            f'{get_column_letter(column)}1000')


    @classmethod
    def add_mock_data_excel(cls, workbook, rows_of_examples: int):
        for sheet in workbook:
            if sheet == workbook["Keuzelijsten"]:
                break
            if rows_of_examples == 0:
                for rows in sheet.iter_rows(min_row=2, max_row=2):
                    for cell in rows:
                        cell.value = ''

    @classmethod
    def remove_geo_artefact_csv(cls, header, data):
        if 'geometry' in header:
            deletion_index = header.index('geometry')
            header.remove('geometry')
            for d in data:
                d.pop(deletion_index)
        return [header, data]

    @classmethod
    def multiple_csv_template(cls, template_file_path, subset_path, temporary_path,
                              instantiated_attributes, **kwargs):
        file_location = os.path.dirname(template_file_path)
        tempdir = Path(tempfile.gettempdir()) / 'temp-otlmow'
        logging.debug(file_location)
        file_name = ntpath.basename(template_file_path)
        split_file_name = file_name.split('.')
        things_in_there = os.listdir(tempdir)
        csv_templates = [x for x in things_in_there if x.startswith(f'{split_file_name[0]}_')]
        for file in csv_templates:
            test_template_loc = Path(os.path.dirname(template_file_path)) / file
            temp_loc = Path(tempdir) / file
            cls.alter_csv_template(template_file_path=test_template_loc, temporary_path=temp_loc,
                                   subset_path=subset_path, **kwargs)

    @classmethod
    async def multiple_csv_template_async(cls, template_file_path, subset_path, temporary_path,
                              instantiated_attributes, **kwargs):
        file_location = os.path.dirname(template_file_path)
        tempdir = Path(tempfile.gettempdir()) / 'temp-otlmow'
        logging.debug(file_location)
        file_name = ntpath.basename(template_file_path)
        split_file_name = file_name.split('.')
        things_in_there = os.listdir(tempdir)
        csv_templates = [x for x in things_in_there if x.startswith(f'{split_file_name[0]}_')]
        for file in csv_templates:
            test_template_loc = Path(os.path.dirname(template_file_path)) / file
            temp_loc = Path(tempdir) / file
            await sleep(0)
            await cls.alter_csv_template_async(
                template_file_path=test_template_loc, temporary_path=temp_loc,
                subset_path=subset_path, **kwargs)

    @classmethod
    def alter_csv_template(cls, template_file_path, subset_path, temporary_path,
                           **kwargs):
        converter = OtlmowConverter()
        instantiated_attributes = converter.from_file_to_objects(file_path=temporary_path,
                                                                 subset_path=subset_path)
        header = []
        data = []
        delimiter = ';'
        add_geo_artefact = kwargs.get('add_geo_artefact', False)
        add_attribute_info = kwargs.get('add_attribute_info', False)
        highlight_deprecated_attributes = kwargs.get('highlight_deprecated_attributes', False)
        amount_of_examples = kwargs.get('amount_of_examples', 0)
        quote_char = '"'
        with open(temporary_path, 'r+', encoding='utf-8') as csvfile:
            with open(template_file_path, 'w', encoding='utf-8') as new_file:
                reader = csv.reader(csvfile, delimiter=delimiter, quotechar=quote_char)
                for row_nr, row in enumerate(reader):
                    if row_nr == 0:
                        header = row
                    else:
                        data.append(row)
                if add_geo_artefact is False:
                    [header, data] = cls.remove_geo_artefact_csv(header=header, data=data)
                if add_attribute_info:
                    [info, header] = cls.add_attribute_info_csv(header=header, data=data,
                                                                instantiated_attributes=instantiated_attributes)
                    new_file.write(delimiter.join(info) + '\n')
                data = cls.add_mock_data_csv(header=header, data=data, rows_of_examples=amount_of_examples)
                if highlight_deprecated_attributes:
                    header = cls.highlight_deprecated_attributes_csv(header=header, data=data,
                                                                     instantiated_attributes=instantiated_attributes)
                new_file.write(delimiter.join(header) + '\n')
                for d in data:
                    new_file.write(delimiter.join(d) + '\n')

    @classmethod
    async def alter_csv_template_async(cls, template_file_path, subset_path, temporary_path,
                           **kwargs):
        converter = OtlmowConverter()
        instantiated_attributes = await converter.from_file_to_objects_async(
            file_path=temporary_path, subset_path=subset_path)
        header = []
        data = []
        delimiter = ';'
        add_geo_artefact = kwargs.get('add_geo_artefact', False)
        add_attribute_info = kwargs.get('add_attribute_info', False)
        highlight_deprecated_attributes = kwargs.get('highlight_deprecated_attributes', False)
        amount_of_examples = kwargs.get('amount_of_examples', 0)
        quote_char = '"'
        with open(temporary_path, 'r+', encoding='utf-8') as csvfile:
            with open(template_file_path, 'w', encoding='utf-8') as new_file:
                reader = csv.reader(csvfile, delimiter=delimiter, quotechar=quote_char)
                for row_nr, row in enumerate(reader):
                    if row_nr == 0:
                        header = row
                    else:
                        data.append(row)
                        await sleep(0)
                if add_geo_artefact is False:
                    [header, data] = cls.remove_geo_artefact_csv(header=header, data=data)
                if add_attribute_info:
                    [info, header] = cls.add_attribute_info_csv(header=header, data=data,
                                                                instantiated_attributes=instantiated_attributes)
                    new_file.write(delimiter.join(info) + '\n')
                data = cls.add_mock_data_csv(header=header, data=data, rows_of_examples=amount_of_examples)
                if highlight_deprecated_attributes:
                    header = cls.highlight_deprecated_attributes_csv(header=header, data=data,
                                                                     instantiated_attributes=instantiated_attributes)
                new_file.write(delimiter.join(header) + '\n')
                for d in data:
                    new_file.write(delimiter.join(d) + '\n')

    @classmethod
    def add_attribute_info_csv(cls, header, data, instantiated_attributes):
        info_data = []
        info_data.extend(header)
        found_uri = []
        dotnotation_module = DotnotationHelper()
        uri_index = cls.find_uri_in_csv(header)
        for d in data:
            if d[uri_index] not in found_uri:
                found_uri.append(d[uri_index])
        for uri in found_uri:
            single_object = next(x for x in instantiated_attributes if x.typeURI == uri)
            for dotnototation_title in info_data:
                if dotnototation_title == 'typeURI':
                    index = info_data.index(dotnototation_title)
                    info_data[index] = 'De URI van het object volgens https://www.w3.org/2001/XMLSchema#anyURI .'
                else:
                    index = info_data.index(dotnototation_title)
                    try:
                        dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(
                            single_object, dotnototation_title)
                    except AttributeError as e:
                        continue
                    info_data[index] = dotnotation_attribute.definition
        return [info_data, header]

    @classmethod
    def add_mock_data_csv(cls, header, data, rows_of_examples):
        if rows_of_examples == 0:
            data = []
        return data

    @classmethod
    def highlight_deprecated_attributes_csv(cls, header, data, instantiated_attributes):
        found_uri = []
        dotnotation_module = DotnotationHelper()
        uri_index = cls.find_uri_in_csv(header)
        for d in data:
            if d[uri_index] not in found_uri:
                found_uri.append(d[uri_index])
        for uri in found_uri:
            single_object = next(x for x in instantiated_attributes if x.typeURI == uri)
            for dotnototation_title in header:
                if dotnototation_title == 'typeURI':
                    continue

                index = header.index(dotnototation_title)
                value = header[index]
                try:
                    is_deprecated = False
                    if dotnototation_title.count('.') == 1:
                        dot_split = dotnototation_title.split('.')
                        attribute = dotnotation_module.get_attribute_by_dotnotation(single_object,
                                                                                    dot_split[0])

                        if len(attribute.deprecated_version) > 0:
                            is_deprecated = True
                    dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(single_object,
                                                                                            dotnototation_title)
                    if len(dotnotation_attribute.deprecated_version) > 0:
                        is_deprecated = True
                except AttributeError:
                    continue
                if is_deprecated:
                    header[index] = f"[DEPRECATED] {value}"
        return header

    @classmethod
    def find_uri_in_csv(cls, header):
        return header.index('typeURI') if 'typeURI' in header else None

    @classmethod
    def add_choice_list_to_sheet(cls, workbook, name, options, choice_list_dict):
        active_sheet = workbook['Keuzelijsten']
        row_nr = 2
        for rows in active_sheet.iter_rows(min_row=1, max_row=1, min_col=1, max_col=700):
            for cell in rows:
                if cell.value is None:
                    cell.value = name
                    column_nr = cell.column
                    for option in options:
                        active_sheet.cell(row=row_nr, column=column_nr, value=option)
                        row_nr += 1
                    choice_list_dict[name] = get_column_letter(column_nr)
                    break
        return choice_list_dict

    @classmethod
    def remove_examples_from_excel_again(cls, workbook):
        first_value_row_i = 3 #with a description the values only start at row 2 (third row)
        for sheet in workbook:
            sheet.delete_rows(idx=first_value_row_i, amount=1)
