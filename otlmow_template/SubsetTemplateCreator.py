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
from openpyxl.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.dimensions import DimensionHolder, ColumnDimension
from otlmow_converter.DotnotationHelper import DotnotationHelper
from otlmow_converter.OtlmowConverter import OtlmowConverter
from otlmow_model.OtlmowModel.BaseClasses.BooleanField import BooleanField
from otlmow_model.OtlmowModel.BaseClasses.KeuzelijstField import KeuzelijstField
from otlmow_model.OtlmowModel.BaseClasses.OTLObject import dynamic_create_instance_from_uri, OTLObject, \
    get_attribute_by_name
from otlmow_model.OtlmowModel.Helpers.generated_lists import get_hardcoded_relation_dict
from otlmow_modelbuilder.OSLOCollector import OSLOCollector
from otlmow_modelbuilder.SQLDataClasses.OSLOClass import OSLOClass

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

enumeration_validation_rules = {
    "valid_uri_and_types": {},
    "valid_regexes": [
        "^https://wegenenverkeer.data.vlaanderen.be/ns/.+"]
}

short_to_long_ns = {
    'ond': 'https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#',
    'onderdeel': 'https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#',
    'ins': 'https://wegenenverkeer.data.vlaanderen.be/ns/installatie#',
    'installatie': 'https://wegenenverkeer.data.vlaanderen.be/ns/installatie#',
    'imp': 'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#',
    'implementatieelement': 'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#',
}


class SubsetTemplateCreator:
    @classmethod
    def _load_collector_from_subset_path(cls, subset_path: Path) -> OSLOCollector:
        collector = OSLOCollector(subset_path)
        collector.collect_all(include_abstract=True)
        return collector

    @classmethod
    def generate_template_from_subset(
            cls,
            subset_path: Path,
            template_file_path: Path,
            ignore_relations: bool = True,
            filter_attributes_by_subset: bool = True,
            class_uris_filter: [str] = None,
            dummy_data_rows: int = 1,
            add_geometry: bool = True,
            add_attribute_info: bool = False,
            tag_deprecated: bool = False,
            generate_choice_list: bool = True,
            split_per_type: bool = True,
            model_directory: Path = None):
        """
        Generate a template from a subset file
        :param subset_path: Path to the subset file
        :param template_file_path: Path to where the template file should be created
        :param ignore_relations: Whether to ignore relations when creating the template, defaults to True
        :param filter_attributes_by_subset: Whether to filter by the attributes in the subset, defaults to True
        :param class_uris_filter: List of class URIs to filter by. If not None, only classes with these URIs will be
        included, defaults to None
        :param dummy_data_rows: Amount of dummy data rows to add to the template, defaults to 1
        :param add_geometry: Whether to include the geometry attribute in the template, defaults to True
        :param add_attribute_info: Whether to add attribute information to the template (colored grey in Excel), defaults to False
        :param tag_deprecated: Whether to tag deprecated attributes in the template, defaults to False
        :param generate_choice_list: Whether to generate a choice list in the template (only for Excel), defaults to True
        :param split_per_type: Whether to split the template into a file per type (only for CSV), defaults to True
        :param model_directory: Path to the model directory, defaults to None
        :return:
        """
        tempdir = Path(tempfile.gettempdir()) / 'temp-otlmow'
        if not tempdir.exists():
            os.makedirs(tempdir)

        # generate objects to write to file
        objects = cls.generate_objects_for_template(
            subset_path=subset_path, ignore_relations=ignore_relations, class_uris_filter=class_uris_filter,
            add_geometry=add_geometry, filter_attributes_by_subset=filter_attributes_by_subset,
            dummy_data_rows=dummy_data_rows, model_directory=model_directory)

        # write the file
        temporary_path = Path(tempdir) / template_file_path.name
        OtlmowConverter.from_objects_to_file(file_path=temporary_path, sequence_of_objects=objects,
                                             split_per_type=split_per_type)

        # alter the file if needed
        extension = template_file_path.suffix.lower()
        if extension == '.xlsx':
            cls.alter_excel_template(template_file_path=template_file_path, generate_choice_list=generate_choice_list,
                                     temporary_path=temporary_path, dummy_data_rows=dummy_data_rows,
                                     subset_path=subset_path, instances=objects, tag_deprecated=tag_deprecated,
                                     add_geometry=add_geometry, add_attribute_info=add_attribute_info)

        elif extension == '.csv':
            cls.determine_multiplicity_csv(
                template_file_path=template_file_path, dummy_data_rows=dummy_data_rows, add_geometry=add_geometry,
                subset_path=subset_path, split_per_type=split_per_type, add_attribute_info=add_attribute_info,
                instances=objects,
                temporary_path=temporary_path, tag_deprecated=tag_deprecated)

    # TODO move the file

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
        instances = (await cls.generate_basic_template_async(
            subset_path=subset_path, temporary_path=temporary_path, ignore_relations=ignore_relations,
            template_file_path=template_file_path,
            filter_attributes_by_subset=filter_attributes_by_subset, **kwargs))
        await sleep(0)
        extension = os.path.splitext(template_file_path)[-1].lower()
        if extension == '.xlsx':
            await cls.alter_excel_template_async(
                template_file_path=template_file_path,
                                      temporary_path=temporary_path,
                                      subset_path=subset_path, instances=instances,
                                      **kwargs)
        elif extension == '.csv':
            await cls.determine_multiplicity_csv_async(
                template_file_path=template_file_path,
                                            subset_path=subset_path,
                                            instances=instances,
                                            temporary_path=temporary_path,
                                            **kwargs)

    @classmethod
    def generate_objects_for_template(
            cls, subset_path: Path, class_uris_filter: [str], filter_attributes_by_subset: bool,
            dummy_data_rows: int, add_geometry: bool, ignore_relations: bool, model_directory: Path = None
    ) -> [OTLObject]:
        """
        This method is used to generate objects for the template. It will generate objects based on the subset file
        """
        collector = cls._load_collector_from_subset_path(subset_path=subset_path)
        filtered_class_list = cls.filters_classes_by_subset(collector=collector, class_uris_filter=class_uris_filter)
        relation_dict = get_hardcoded_relation_dict(model_directory=model_directory)

        amount_objects_to_create = max(1, dummy_data_rows)
        otl_objects = []

        for oslo_class in [cl for cl in filtered_class_list if cl.abstract == 0]:
            if ignore_relations and oslo_class.objectUri in relation_dict:
                continue

            otl_objects.extend(cls.generate_objects_from_oslo_class(
                oslo_class=oslo_class, amount_objects_to_create=amount_objects_to_create, add_geometry=add_geometry,
                filter_attributes_by_subset=filter_attributes_by_subset, collector=collector,
                model_directory=model_directory))

        return otl_objects

    @classmethod
    def generate_objects_from_oslo_class(
            cls, oslo_class: OSLOClass, amount_objects_to_create: int, add_geometry: bool,
            filter_attributes_by_subset: bool, collector: OSLOCollector, model_directory: Path = None) -> [OTLObject]:
        """
        Generate a number of objects from a given OSLO class
        """

        otl_objects = []

        for _ in range(amount_objects_to_create):
            instance = dynamic_create_instance_from_uri(oslo_class.objectUri, model_directory=model_directory)
            if instance is None:
                continue

            if filter_attributes_by_subset:
                for attribute_object in collector.find_attributes_by_class(oslo_class):
                    attr = get_attribute_by_name(instance, attribute_object.name)
                    attr.fill_with_dummy_data()
            else:
                for attr in instance:
                    if attr.naam != 'geometry':
                        attr.fill_with_dummy_data()
            with contextlib.suppress(AttributeError):
                if add_geometry:
                    geo_attr = get_attribute_by_name(instance, 'geometry')
                    if geo_attr is not None:
                        geo_attr.fill_with_dummy_data()

            asset_versie = get_attribute_by_name(instance, 'assetVersie')
            if asset_versie is not None:
                asset_versie.set_waarde(None)

            otl_objects.append(instance)

            DotnotationHelper.clear_list_of_list_attributes(instance)

            otl_objects.append(instance)

        return otl_objects

    @classmethod
    async def generate_basic_template_async(cls, subset_path: Path, template_file_path: Path,
                                temporary_path: Path, class_uris_filter: [str] = None, ignore_relations: bool = True,
                                            **kwargs):
        collector = cls._load_collector_from_subset_path(subset_path=subset_path)
        filtered_class_list = cls.filters_classes_by_subset(
            collector=collector, class_uris_filter=class_uris_filter)
        otl_objects = []
        dummy_data_rows = kwargs.get('dummy_data_rows', 0)
        model_directory = None
        if kwargs is not None:
            model_directory = kwargs.get('model_directory', None)
        relation_dict = get_hardcoded_relation_dict(model_directory=model_directory)

        generate_dummy_records = 1
        if dummy_data_rows > 1:
            generate_dummy_records = dummy_data_rows

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
        instances = []
        if path_is_split is False or extension == '.xlsx':
            instances = await converter.from_file_to_objects_async(
                file_path=temporary_path, subset_path=subset_path)
        return list(instances)

    @classmethod
    def alter_excel_template(cls, template_file_path: Path, subset_path: Path, add_geometry: bool,
                             instances: list, temporary_path: Path, add_attribute_info: bool,
                             generate_choice_list: bool, dummy_data_rows: int, tag_deprecated: bool):
        wb = load_workbook(temporary_path, read_only=False)
        wb.create_sheet('Keuzelijsten')

        # tag_deprecated (loop over attributes by header row) alter header row
        # add_attribute_info (loop over attributes by header row) collect attribute info and insert first row when done
        # generate_choice_list (loop over attributes by header row) alter Keuzelijsten sheet (use iter_cols)
        # type_uri choice_list in Excel (loop over attributes) (use iter_cols)

        # set_fixed_column_width (25)

        choice_list_dict = {}
        for sheet in wb:
            if sheet.title == 'Keuzelijsten':
                break

            type_uri = cls.get_uri_from_sheet_name(sheet.title)
            instance = next(x for x in instances if x.typeURI == type_uri)

            boolean_validation = DataValidation(type="list", formula1=f'"{type_uri}"', allow_blank=True)
            sheet.add_data_validation(boolean_validation)

            collected_attribute_info = []
            header_row_nr = 2 if add_attribute_info else 1

            header_row = next(sheet.iter_rows(min_row=header_row_nr, max_row=header_row_nr))
            for index, header_cell in enumerate(header_row):
                header = header_cell.value
                if header is None or header == '':
                    continue

                # add type_uri
                if header == 'typeURI':
                    data_validation = DataValidation(type="list", formula1=f'"{type_uri}"', allow_blank=True)
                    sheet.add_data_validation(data_validation)
                    data_validation.add(f'{header_cell.column_letter}{(header_row_nr + 1)}:'
                                        f'{header_cell.column_letter}1000')
                    continue

                attribute = DotnotationHelper.get_attribute_by_dotnotation(instance, header)

                if add_attribute_info:
                    collected_attribute_info.append(attribute.definition)

                if tag_deprecated and attribute.deprecated_version:
                    sheet.cell(row=header_row_nr, column=index + 1, value=f'[DEPRECATED] {header}')

                if generate_choice_list:
                    if issubclass(attribute.field, BooleanField):
                        boolean_validation.add(f'{header_cell.column_letter}{(header_row_nr + 1)}:'
                                            f'{header_cell.column_letter}1000')
                        continue

                    if issubclass(attribute.field, KeuzelijstField):
                        choice_list_values = [cv for cv in attribute.field.options.values()
                                              if cv.status != 'verwijderd']
                        if attribute.field.naam not in choice_list_dict:
                            # add choice_list to sheet Keuzelijsten and save its column
                            cls.add_choice_list_to_sheet(workbook=wb, name=attribute.field.naam,
                                                         options=choice_list_values, choice_list_dict=choice_list_dict)

                        column_in_choice_sheet = choice_list_dict[attribute.field.naam]
                        start_range = f"${column_in_choice_sheet}$2"
                        end_range = f"${column_in_choice_sheet}${len(choice_list_values) + 1}"
                        data_val = DataValidation(type="list", formula1=f"Keuzelijsten!{start_range}:{end_range}",
                                                  allowBlank=True)
                        sheet.add_data_validation(data_val)
                        data_val.add(f'{get_column_letter(header_cell.column)}{header_row_nr + 1}:'
                                     f'{get_column_letter(header_cell.column)}1000')

        wb.save(template_file_path)
        file_location = os.path.dirname(temporary_path)
        [f.unlink() for f in Path(file_location).glob("*") if f.is_file()]


    @classmethod
    async def alter_excel_template_async(cls, template_file_path: Path, subset_path: Path, add_geometry: bool,
                             instances: list, temporary_path, **kwargs):
        await sleep(0)
        generate_choice_list = kwargs.get('generate_choice_list', False)
        add_attribute_info = kwargs.get('add_attribute_info', False)
        tag_deprecated = kwargs.get('tag_deprecated', False)
        dummy_data_rows = kwargs.get('dummy_data_rows', 0)
        original_dummy_data_rows = dummy_data_rows
        if add_attribute_info and dummy_data_rows == 0:
            dummy_data_rows = 1
        await sleep(0)
        wb = load_workbook(temporary_path)
        wb.create_sheet('Keuzelijsten')
        # Volgorde is belangrijk! Eerst rijen verwijderen indien nodig dan choice list toevoegen,
        # staat namelijk vast op de kolom en niet het attribuut in die kolom
        if add_geometry is False:
            await sleep(0)
            cls.remove_geo_artefact_excel(workbook=wb)
        if generate_choice_list:
            await sleep(0)
            await cls.add_choice_list_excel_async(workbook=wb, instances=instances,
                                      subset_path=subset_path, add_attribute_info=add_attribute_info)
        await sleep(0)
        cls.add_mock_data_excel(workbook=wb, rows_of_examples=dummy_data_rows) # remove dummy rows if needed

        await cls.custom_exel_fixes_async(workbook=wb, instances=instances,
                                    add_attribute_info=add_attribute_info)
        await sleep(0)
        if tag_deprecated:
            await sleep(0)
            cls.check_for_deprecated_attributes(workbook=wb, instances=instances)
        if add_attribute_info:
            await sleep(0)
            await cls.add_attribute_info_excel_async(workbook=wb, instances=instances)
        if original_dummy_data_rows == 0 and add_attribute_info:
            cls.remove_examples_from_excel_again(workbook=wb)
        await sleep(0)
        wb.save(template_file_path)
        file_location = os.path.dirname(temporary_path)
        [f.unlink() for f in Path(file_location).glob("*") if f.is_file()]

    @classmethod
    def determine_multiplicity_csv(cls, template_file_path: Path, subset_path: Path,
                                   instances: list, temporary_path: Path, **kwargs):
        path_is_split = kwargs.get('split_per_type', True)
        if path_is_split is False:
            cls.alter_csv_template(template_file_path=template_file_path,
                                    temporary_path=temporary_path, subset_path=subset_path, **kwargs)
        else:
            cls.multiple_csv_template(
                template_file_path=template_file_path,
                                       temporary_path=temporary_path,
                                       subset_path=subset_path, instances=instances,
                                       **kwargs)
        file_location = os.path.dirname(temporary_path)
        [f.unlink() for f in Path(file_location).glob("*") if f.is_file()]

    @classmethod
    async def determine_multiplicity_csv_async(cls, template_file_path: Path, subset_path: Path,
                                   instances: list, temporary_path: Path, **kwargs):
        path_is_split = kwargs.get('split_per_type', True)
        await sleep(0)
        if path_is_split is False:
            await cls.alter_csv_template_async(template_file_path=template_file_path,
                                    temporary_path=temporary_path, subset_path=subset_path, **kwargs)
        else:
            await cls.multiple_csv_template_async(
                template_file_path=template_file_path,
                                       temporary_path=temporary_path,
                                       subset_path=subset_path, instances=instances,
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
    def add_type_uri_choice_list_in_excel(cls, sheet, instances, add_attribute_info: bool):
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

            possible_classes = [x for x in instances if x.typeURI.endswith(subclass_name)]
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
    async def add_type_uri_choice_list_in_excel_async(cls, sheet, instances, add_attribute_info: bool):
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

            possible_classes = [x for x in instances if x.typeURI.endswith(subclass_name)]
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
    def custom_exel_fixes(cls, workbook, instances, add_attribute_info: bool):
        for sheet in workbook:
            cls.set_fixed_column_width(sheet=sheet, width=25)
            cls.add_type_uri_choice_list_in_excel(sheet=sheet, instances=instances,
                                                        add_attribute_info=add_attribute_info)
            cls.remove_asset_versie(sheet=sheet)

    @classmethod
    async def custom_exel_fixes_async(cls, workbook, instances, add_attribute_info: bool):
        for sheet in workbook:
            await sleep(0)
            await cls.set_fixed_column_width_async(sheet=sheet, width=25)
            await cls.add_type_uri_choice_list_in_excel_async(sheet=sheet,
                                                            instances=instances,
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
    def add_attribute_info_excel(cls, workbook, instances: list):
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            if sheet == workbook['Keuzelijsten']:
                break
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            single_attribute = next(x for x in instances if x.typeURI == filter_uri)
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
    async def add_attribute_info_excel_async(cls, workbook, instances: list):
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            if sheet == workbook['Keuzelijsten']:
                break
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            await sleep(0)
            single_attribute = next(x for x in instances if x.typeURI == filter_uri)
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
    def check_for_deprecated_attributes(cls, workbook, instances: list):
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            if sheet == workbook['Keuzelijsten']:
                break
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            single_attribute = next(x for x in instances if x.typeURI == filter_uri)
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
    def add_choice_list_excel(cls, workbook, instances: list, subset_path: Path,
                                    add_attribute_info: bool=False):
        choice_list_dict = {}
        starting_row = '3' if add_attribute_info else '2'
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            if sheet == workbook['Keuzelijsten']:
                break
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            single_attribute = next(x for x in instances if x.typeURI == filter_uri)
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
    async def add_choice_list_excel_async(cls, workbook, instances: list, subset_path: Path,
                                    add_attribute_info: bool = False):
        choice_list_dict = {}
        starting_row = '3' if add_attribute_info else '2'
        dotnotation_module = DotnotationHelper()
        for sheet in workbook:
            await sleep(0)
            if sheet == workbook['Keuzelijsten']:
                break
            filter_uri = SubsetTemplateCreator.find_uri_in_sheet(sheet)
            single_attribute = next(x for x in instances if x.typeURI == filter_uri)
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
    def remove_dummy_row_from_excel(cls, workbook: Workbook, add_attribute_info: bool):
        row_nr = 3 if add_attribute_info else 2
        for sheet in workbook:
            if sheet == workbook["Keuzelijsten"]:
                break
            sheet.delete_rows(row_nr, 1)

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
                              instances, **kwargs):
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
                              instances, **kwargs):
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
        instances = converter.from_file_to_objects(file_path=temporary_path,
                                                                 subset_path=subset_path)
        header = []
        data = []
        delimiter = ';'
        add_geometry = kwargs.get('add_geometry', False)
        add_attribute_info = kwargs.get('add_attribute_info', False)
        tag_deprecated = kwargs.get('tag_deprecated', False)
        dummy_data_rows = kwargs.get('dummy_data_rows', 0)
        quote_char = '"'
        with open(temporary_path, 'r+', encoding='utf-8') as csvfile:
            with open(template_file_path, 'w', encoding='utf-8') as new_file:
                reader = csv.reader(csvfile, delimiter=delimiter, quotechar=quote_char)
                for row_nr, row in enumerate(reader):
                    if row_nr == 0:
                        header = row
                    else:
                        data.append(row)
                if add_geometry is False:
                    [header, data] = cls.remove_geo_artefact_csv(header=header, data=data)
                if add_attribute_info:
                    [info, header] = cls.add_attribute_info_csv(header=header, data=data,
                                                                instances=instances)
                    new_file.write(delimiter.join(info) + '\n')
                data = cls.add_mock_data_csv(header=header, data=data, rows_of_examples=dummy_data_rows)
                if tag_deprecated:
                    header = cls.tag_deprecated_csv(header=header, data=data,
                                                                     instances=instances)
                new_file.write(delimiter.join(header) + '\n')
                for d in data:
                    new_file.write(delimiter.join(d) + '\n')

    @classmethod
    async def alter_csv_template_async(cls, template_file_path, subset_path, temporary_path,
                           **kwargs):
        converter = OtlmowConverter()
        instances = await converter.from_file_to_objects_async(
            file_path=temporary_path, subset_path=subset_path)
        header = []
        data = []
        delimiter = ';'
        add_geometry = kwargs.get('add_geometry', False)
        add_attribute_info = kwargs.get('add_attribute_info', False)
        tag_deprecated = kwargs.get('tag_deprecated', False)
        dummy_data_rows = kwargs.get('dummy_data_rows', 0)
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
                if add_geometry is False:
                    [header, data] = cls.remove_geo_artefact_csv(header=header, data=data)
                if add_attribute_info:
                    [info, header] = cls.add_attribute_info_csv(header=header, data=data,
                                                                instances=instances)
                    new_file.write(delimiter.join(info) + '\n')
                data = cls.add_mock_data_csv(header=header, data=data, rows_of_examples=dummy_data_rows)
                if tag_deprecated:
                    header = cls.tag_deprecated_csv(header=header, data=data,
                                                                     instances=instances)
                new_file.write(delimiter.join(header) + '\n')
                for d in data:
                    new_file.write(delimiter.join(d) + '\n')

    @classmethod
    def add_attribute_info_csv(cls, header, data, instances):
        info_data = []
        info_data.extend(header)
        found_uri = []
        dotnotation_module = DotnotationHelper()
        uri_index = cls.find_uri_in_csv(header)
        for d in data:
            if d[uri_index] not in found_uri:
                found_uri.append(d[uri_index])
        for uri in found_uri:
            single_object = next(x for x in instances if x.typeURI == uri)
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
    def tag_deprecated_csv(cls, header, data, instances):
        found_uri = []
        dotnotation_module = DotnotationHelper()
        uri_index = cls.find_uri_in_csv(header)
        for d in data:
            if d[uri_index] not in found_uri:
                found_uri.append(d[uri_index])
        for uri in found_uri:
            single_object = next(x for x in instances if x.typeURI == uri)
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
        column_nr = choice_list_dict.keys().__len__() + 1
        row_nr = 1
        new_header = active_sheet.cell(row=row_nr, column=column_nr)
        if new_header.value is not None:
            raise ValueError(f'Header already exists at column {column_nr}: {new_header.value}')
        new_header.value = name
        for index, option in enumerate(options, start=1):
            cell = active_sheet.cell(row=row_nr + index, column=column_nr)
            cell.value = option.invulwaarde

        choice_list_dict[name] = new_header.column_letter

    @classmethod
    def remove_examples_from_excel_again(cls, workbook):
        first_value_row_i = 3 #with a description the values only start at row 2 (third row)
        for sheet in workbook:
            sheet.delete_rows(idx=first_value_row_i, amount=1)

    @classmethod
    def get_uri_from_sheet_name(cls, title: str) -> str:
        if title == 'Agent':
            return 'http://purl.org/dc/terms/Agent'
        if '#' not in title:
            raise ValueError('Sheet title does not contain a #')
        class_ns, class_name = title.split('#', maxsplit=1)
        return short_to_long_ns.get(class_ns, class_ns) + class_name
