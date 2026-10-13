import asyncio
import os
from asyncio import sleep
from pathlib import Path

from openpyxl.worksheet.worksheet import Worksheet
from otlmow_converter.OtlmowConverter import OtlmowConverter
from otlmow_model.OtlmowModel.BaseClasses.OTLObject import OTLObject
from otlmow_modelbuilder.OSLOCollector import OSLOCollector
from otlmow_modelbuilder.SQLDataClasses.OSLOClass import OSLOClass

from otlmow_template.AanlevertemplateGenerator import AanlevertemplateGenerator

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))


class SubsetTemplateCreator:
    """Legacy facade for generating delivery templates from a subset.
    
    Delegates to AanlevertemplateGenerator which contains the actual implementation.
    """
    
    green_fill = AanlevertemplateGenerator.green_fill
    _generator: AanlevertemplateGenerator = None
    
    @classmethod
    def _get_generator(cls) -> AanlevertemplateGenerator:
        if cls._generator is None:
            cls._generator = AanlevertemplateGenerator()
        return cls._generator
    
    @classmethod
    def _load_collector_from_subset_path(cls, subset_path: Path) -> OSLOCollector:
        return cls._get_generator()._load_collector_from_subset_path(subset_path)
    
    @classmethod
    async def generate_template_from_subset_async(
            cls,
            subset_path: Path,
            template_file_path: Path,
            ignore_relations: bool = True,
            filter_attributes_by_subset: bool = True,
            class_uris_filter: [str] = None,
            dummy_data_rows: int = 1,
            add_geometry: bool = True,
            add_attribute_info: bool = False,
            add_deprecated: bool = False,
            generate_choice_list: bool = True,
            split_per_type: bool = True,
            model_directory: Path = None, **kwargs):
        """
        Generate a template from a subset file, async version.
        Await this function!
        """
        generator = cls._get_generator()
        
        objects = await generator.generate_objects_for_template_async(
            subset_path=subset_path, ignore_relations=ignore_relations, class_uris_filter=class_uris_filter,
            add_geometry=add_geometry, filter_attributes_by_subset=filter_attributes_by_subset,
            dummy_data_rows=dummy_data_rows, model_directory=model_directory)
        
        abbreviate_excel_after = False
        if kwargs.get('abbreviate_excel_sheettitles') == True:
            kwargs['abbreviate_excel_sheettitles'] = False
            abbreviate_excel_after = True
        
        await OtlmowConverter.from_objects_to_file_async(
            file_path=template_file_path, sequence_of_objects=objects, split_per_type=split_per_type, **kwargs)
        
        extension = template_file_path.suffix.lower()
        if extension == '.xlsx':
            await generator.alter_excel_template_async(
                generate_choice_list=generate_choice_list, file_path=template_file_path, dummy_data_rows=dummy_data_rows,
                instances=objects, add_deprecated=add_deprecated, add_attribute_info=add_attribute_info,
                abbreviate_excel_sheettitles=abbreviate_excel_after)
        
        elif extension == '.csv':
            await generator.alter_csv_template_async(
                split_per_type=split_per_type, file_path=template_file_path, dummy_data_rows=dummy_data_rows,
                instances=objects, add_deprecated=add_deprecated, add_attribute_info=add_attribute_info)
    
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
            add_deprecated: bool = False,
            generate_choice_list: bool = True,
            split_per_type: bool = True,
            model_directory: Path = None, **kwargs):
        """
         Generate a template from a subset file.
        """
        generator = cls._get_generator()
        
        objects = generator.generate_objects_for_template(
            subset_path=subset_path, ignore_relations=ignore_relations, class_uris_filter=class_uris_filter,
            add_geometry=add_geometry, filter_attributes_by_subset=filter_attributes_by_subset,
            dummy_data_rows=dummy_data_rows, model_directory=model_directory)
        
        abbreviate_excel_after = False
        if kwargs.get('abbreviate_excel_sheettitles') == True:
            kwargs['abbreviate_excel_sheettitles'] = False
            abbreviate_excel_after = True
        
        OtlmowConverter.from_objects_to_file(
            file_path=template_file_path, sequence_of_objects=objects, split_per_type=split_per_type,
            model_directory=model_directory, **kwargs)
        
        extension = template_file_path.suffix.lower()
        if extension == '.xlsx':
            generator.alter_excel_template(
                generate_choice_list=generate_choice_list, file_path=template_file_path, dummy_data_rows=dummy_data_rows,
                instances=objects, add_deprecated=add_deprecated, add_attribute_info=add_attribute_info,
                abbreviate_excel_sheettitles=abbreviate_excel_after)
        elif extension == '.csv':
            generator.alter_csv_template(
                split_per_type=split_per_type, file_path=template_file_path, dummy_data_rows=dummy_data_rows,
                instances=objects, add_deprecated=add_deprecated, add_attribute_info=add_attribute_info)
    
    @classmethod
    def create_x_objects(cls, oslo_class, add_geometry, collector, filter_attributes_by_subset, model_directory,
                         amount_objects_to_create):
        return cls._get_generator().create_x_objects(
            oslo_class=oslo_class, add_geometry=add_geometry, collector=collector,
            filter_attributes_by_subset=filter_attributes_by_subset, model_directory=model_directory,
            amount_objects_to_create=amount_objects_to_create)
    
    @classmethod
    def generate_objects_for_template(
            cls, subset_path: Path, class_uris_filter: [str], filter_attributes_by_subset: bool,
            dummy_data_rows: int, add_geometry: bool, ignore_relations: bool, model_directory: Path = None
    ) -> [OTLObject]:
        return cls._get_generator().generate_objects_for_template(
            subset_path=subset_path, class_uris_filter=class_uris_filter,
            filter_attributes_by_subset=filter_attributes_by_subset, dummy_data_rows=dummy_data_rows,
            add_geometry=add_geometry, ignore_relations=ignore_relations, model_directory=model_directory)
    
    @classmethod
    async def create_x_objects_async(cls, oslo_class, add_geometry, collector, filter_attributes_by_subset,
                                  model_directory, amount_objects_to_create):
        return await cls._get_generator().create_x_objects_async(
            oslo_class=oslo_class, add_geometry=add_geometry, collector=collector,
            filter_attributes_by_subset=filter_attributes_by_subset, model_directory=model_directory,
            amount_objects_to_create=amount_objects_to_create)
    
    @classmethod
    async def generate_objects_for_template_async(
            cls, subset_path: Path, class_uris_filter: [str], filter_attributes_by_subset: bool,
            dummy_data_rows: int, add_geometry: bool, ignore_relations: bool, model_directory: Path = None
    ) -> [OTLObject]:
        return await cls._get_generator().generate_objects_for_template_async(
            subset_path=subset_path, class_uris_filter=class_uris_filter,
            filter_attributes_by_subset=filter_attributes_by_subset, dummy_data_rows=dummy_data_rows,
            add_geometry=add_geometry, ignore_relations=ignore_relations, model_directory=model_directory)
    
    @classmethod
    def generate_object_from_oslo_class(
            cls, oslo_class: OSLOClass, add_geometry: bool,
            filter_attributes_by_subset: bool, collector: OSLOCollector, model_directory: Path = None) -> [OTLObject]:
        return cls._get_generator().generate_object_from_oslo_class(
            oslo_class=oslo_class, add_geometry=add_geometry,
            filter_attributes_by_subset=filter_attributes_by_subset, collector=collector,
            model_directory=model_directory)
    
    @classmethod
    def alter_excel_template(cls, instances: list, file_path: Path, add_attribute_info: bool,
                             generate_choice_list: bool, dummy_data_rows: int, add_deprecated: bool,
                             abbreviate_excel_sheettitles: bool):
        cls._get_generator().alter_excel_template(
            instances=instances, file_path=file_path, add_attribute_info=add_attribute_info,
            generate_choice_list=generate_choice_list, dummy_data_rows=dummy_data_rows,
            add_deprecated=add_deprecated, abbreviate_excel_sheettitles=abbreviate_excel_sheettitles)
    
    @classmethod
    def fill_class_dict(cls, instances: list) -> dict:
        return cls._get_generator().fill_class_dict(instances)
    
    @classmethod
    def alter_csv_template(cls, instances: list, file_path: Path, add_attribute_info: bool,
                           split_per_type: bool, dummy_data_rows: int, add_deprecated: bool):
        cls._get_generator().alter_csv_template(
            instances=instances, file_path=file_path, add_attribute_info=add_attribute_info,
            split_per_type=split_per_type, dummy_data_rows=dummy_data_rows, add_deprecated=add_deprecated)
    
    @classmethod
    async def alter_csv_template_async(cls, instances: list, file_path: Path, add_deprecated: bool,
                                       add_attribute_info: bool, split_per_type: bool, dummy_data_rows: int):
        await cls._get_generator().alter_csv_template_async(
            instances=instances, file_path=file_path, add_deprecated=add_deprecated,
            add_attribute_info=add_attribute_info, split_per_type=split_per_type, dummy_data_rows=dummy_data_rows)
    
    @classmethod
    def alter_csv_file(cls, add_attribute_info: bool, instances: [OTLObject], add_deprecated: bool, file_path: Path,
                       dummy_data_rows: int):
        cls._get_generator().alter_csv_file(
            add_attribute_info=add_attribute_info, instances=instances, add_deprecated=add_deprecated,
            file_path=file_path, dummy_data_rows=dummy_data_rows)
    
    @classmethod
    async def alter_excel_template_async(cls, instances: list, file_path: Path, add_attribute_info: bool,
                                          generate_choice_list: bool, dummy_data_rows: int, add_deprecated: bool,
                                          abbreviate_excel_sheettitles: bool):
        await cls._get_generator().alter_excel_template_async(
            instances=instances, file_path=file_path, add_attribute_info=add_attribute_info,
            generate_choice_list=generate_choice_list, dummy_data_rows=dummy_data_rows,
            add_deprecated=add_deprecated, abbreviate_excel_sheettitles=abbreviate_excel_sheettitles)
    
    @classmethod
    def alter_excel_sheet(cls, add_attribute_info: bool, choice_list_dict: dict, generate_choice_list: bool,
                          instances: [OTLObject], sheet: Worksheet, add_deprecated: bool, workbook: Worksheet,
                          dummy_data_rows: int, abbreviate_excel_sheettitle: bool = True):
        cls._get_generator().alter_excel_sheet(
            add_attribute_info=add_attribute_info, choice_list_dict=choice_list_dict,
            generate_choice_list=generate_choice_list, instances=instances, sheet=sheet,
            add_deprecated=add_deprecated, workbook=workbook, dummy_data_rows=dummy_data_rows,
            abbreviate_excel_sheettitle=abbreviate_excel_sheettitle)
    
    @classmethod
    def color_choice_lists_green(cls, sheet: Worksheet, header_cell_column):
        cls._get_generator().color_choice_lists_green(sheet=sheet, header_cell_column=header_cell_column)
    
    @classmethod
    def add_deprecated_row_to_sheet(cls, deprecated_attributes_row, sheet):
        cls._get_generator().add_deprecated_row_to_sheet(deprecated_attributes_row, sheet)
    
    @classmethod
    def add_attribute_info_to_sheet(cls, collected_attribute_info: list[str], sheet: Worksheet):
        cls._get_generator().add_attribute_info_to_sheet(collected_attribute_info, sheet)
    
    @classmethod
    def get_attribute_info(cls, instance: OTLObject, header: str, attribute) -> str:
        return cls._get_generator().get_attribute_info(instance=instance, header=header, attribute=attribute)
    
    @classmethod
    def get_quantitative_attribute(cls, instance: OTLObject, header: str):
        return cls._get_generator().get_quantitative_attribute(instance=instance, header=header)
    
    @classmethod
    def get_standard_unit(cls, attribute) -> str:
        return cls._get_generator().get_standard_unit(attribute=attribute)
    
    @classmethod
    def generate_choice_list_in_excel(cls, attribute, choice_list_dict, column, row_nr, sheet: Worksheet,
                                      workbook: Worksheet):
        cls._get_generator().generate_choice_list_in_excel(
            attribute=attribute, choice_list_dict=choice_list_dict, column=column, row_nr=row_nr,
            sheet=sheet, workbook=workbook)
    
    @classmethod
    def filters_classes_by_subset(cls, collector: OSLOCollector,
                                  class_uris_filter: [str] = None) -> list[OSLOClass]:
        return cls._get_generator().filters_classes_by_subset(collector=collector, class_uris_filter=class_uris_filter)
    
    @classmethod
    def set_fixed_column_width(cls, sheet, width: int):
        cls._get_generator().set_fixed_column_width(sheet=sheet, width=width)
    
    @classmethod
    def add_choice_list_to_sheet(cls, workbook, name, options, choice_list_dict):
        cls._get_generator().add_choice_list_to_sheet(
            workbook=workbook, name=name, options=options, choice_list_dict=choice_list_dict)
    
    @classmethod
    def get_uri_from_sheet_name(cls, title: str) -> str:
        return cls._get_generator().get_uri_from_sheet_name(title=title)
    
    @classmethod
    def append_relations_to_objects(cls, otl_objects: [OTLObject], collector: OSLOCollector, class_uris_filter: [str],
                                    model_directory: Path):
        cls._get_generator().append_relations_to_objects(
            otl_objects=otl_objects, collector=collector, class_uris_filter=class_uris_filter,
            model_directory=model_directory)
    
    @classmethod
    def abbreviate_excel_sheettitle(cls, sheet):
        cls._get_generator().abbreviate_excel_sheettitle(sheet=sheet)