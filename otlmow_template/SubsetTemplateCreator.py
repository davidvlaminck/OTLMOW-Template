"""The public entry point of the template generation.

The class is the facade of the parts of the generation:

- :mod:`otlmow_template.object_generator` creates the objects of a template
- :mod:`otlmow_template.relations` adds the relations between those objects
- :mod:`otlmow_template.excel_sheet` fills the sheets of an Excel template
- :mod:`otlmow_template.excel_choice_list` adds the choice lists of an Excel template
- :mod:`otlmow_template.csv_writer` writes the csv template
- :mod:`otlmow_template.attribute_info` describes the attributes of a template
"""

from pathlib import Path
from typing import Optional

from otlmow_converter.OtlmowConverter import OtlmowConverter
from otlmow_template.attribute_info import (
    AttributeInfo,
    AttributeInfoMixin,
    find_standard_unit,
    get_attribute_info,
    get_quantitative_attribute,
    get_standard_unit,
)
from otlmow_template.csv_writer import CsvWriterMixin
from otlmow_template.excel_base import ExcelBaseMixin
from otlmow_template.excel_choice_list import ExcelChoiceListMixin
from otlmow_template.excel_context import ExcelSheetContext
from otlmow_template.excel_sheet import ExcelSheetMixin
from otlmow_template.object_generator import GenerationContext, ObjectGeneratorMixin

__all__ = [
    'AttributeInfo',
    'AttributeInfoMixin',
    'CsvWriterMixin',
    'ExcelBaseMixin',
    'ExcelChoiceListMixin',
    'ExcelSheetContext',
    'ExcelSheetMixin',
    'GenerationContext',
    'ObjectGeneratorMixin',
    'SubsetTemplateCreator',
    'find_standard_unit',
    'get_attribute_info',
    'get_quantitative_attribute',
    'get_standard_unit',
]


class SubsetTemplateCreator(
    ExcelSheetMixin,
    CsvWriterMixin,
    ObjectGeneratorMixin,
    AttributeInfoMixin,
):
    """Generates a template from a subset file."""

    @classmethod
    def generate_template_from_subset(
        cls,
        subset_path: Path,
        template_file_path: Path,
        ignore_relations: bool = True,
        filter_attributes_by_subset: bool = True,
        class_uris_filter: Optional[list[str]] = None,
        dummy_data_rows: int = 1,
        add_geometry: bool = True,
        add_attribute_info: bool = False,
        add_deprecated: bool = False,
        generate_choice_list: bool = True,
        split_per_type: bool = True,
        model_directory: Optional[Path] = None,
        **kwargs,
    ):
        """
        Generate a template from a subset file.

        :param subset_path: Path to the subset file
        :param template_file_path: Path to where the template file should be created
        :param ignore_relations: Whether to ignore relations when creating the template, defaults to True
        :param filter_attributes_by_subset: Whether to filter by the attributes in the subset, defaults to True
        :param class_uris_filter: List of class URIs to filter by. If not None, only classes with these URIs will be included, defaults to None
        :param dummy_data_rows: Amount of dummy data rows to add to the template, defaults to 1
        :param add_geometry: Whether to include the geometry attribute in the template, defaults to True
        :param add_attribute_info: Whether to add attribute information to the template (colored grey in Excel), defaults to False
        :param add_deprecated: Whether to tag deprecated attributes in the template, defaults to False
        :param generate_choice_list: Whether to generate a choice list in the template (only for Excel), defaults to True
        :param split_per_type: Whether to split the template into a file per type (only for CSV), defaults to True
        :param model_directory: Path to the model directory, defaults to None

        :return: None
        """
        # generate objects to write to file
        objects = cls.generate_objects_for_template(
            subset_path=subset_path,
            ignore_relations=ignore_relations,
            class_uris_filter=class_uris_filter,
            add_geometry=add_geometry,
            filter_attributes_by_subset=filter_attributes_by_subset,
            dummy_data_rows=dummy_data_rows,
            model_directory=model_directory,
        )

        abbreviate_excel_after = False
        if kwargs.get('abbreviate_excel_sheettitles'):
            kwargs['abbreviate_excel_sheettitles'] = False
            abbreviate_excel_after = True

        # write the file
        OtlmowConverter.from_objects_to_file(
            file_path=template_file_path,
            sequence_of_objects=objects,
            split_per_type=split_per_type,
            model_directory=model_directory,
            **kwargs,
        )

        # alter the file if needed
        extension = template_file_path.suffix.lower()
        if extension == '.xlsx':
            cls.alter_excel_template(
                generate_choice_list=generate_choice_list,
                file_path=template_file_path,
                dummy_data_rows=dummy_data_rows,
                instances=objects,
                add_deprecated=add_deprecated,
                add_attribute_info=add_attribute_info,
                abbreviate_excel_sheettitles=abbreviate_excel_after,
            )
        elif extension == '.csv':
            cls.alter_csv_template(
                split_per_type=split_per_type,
                file_path=template_file_path,
                dummy_data_rows=dummy_data_rows,
                instances=objects,
                add_deprecated=add_deprecated,
                add_attribute_info=add_attribute_info,
            )

    @classmethod
    async def generate_template_from_subset_async(
        cls,
        subset_path: Path,
        template_file_path: Path,
        ignore_relations: bool = True,
        filter_attributes_by_subset: bool = True,
        class_uris_filter: Optional[list[str]] = None,
        dummy_data_rows: int = 1,
        add_geometry: bool = True,
        add_attribute_info: bool = False,
        add_deprecated: bool = False,
        generate_choice_list: bool = True,
        split_per_type: bool = True,
        model_directory: Optional[Path] = None,
        **kwargs,
    ):
        """
        Generate a template from a subset file, async version.
        Await this function!

        :param subset_path: Path to the subset file
        :param template_file_path: Path to where the template file should be created
        :param ignore_relations: Whether to ignore relations when creating the template, defaults to True
        :param filter_attributes_by_subset: Whether to filter by the attributes in the subset, defaults to True
        :param class_uris_filter: List of class URIs to filter by. If not None, only classes with these URIs will be included, defaults to None
        :param dummy_data_rows: Amount of dummy data rows to add to the template, defaults to 1
        :param add_geometry: Whether to include the geometry attribute in the template, defaults to True
        :param add_attribute_info: Whether to add attribute information to the template (colored grey in Excel), defaults to False
        :param add_deprecated: Whether to add a deprecated row to the template (colored red in Excel), defaults to False
        :param generate_choice_list: Whether to generate a choice list in the template (only for Excel), defaults to True
        :param split_per_type: Whether to split the template into a file per type (only for CSV), defaults to True
        :param model_directory: Path to the model directory, defaults to None

        :return: None
        """
        # generate objects to write to file
        objects = await cls.generate_objects_for_template_async(
            subset_path=subset_path,
            ignore_relations=ignore_relations,
            class_uris_filter=class_uris_filter,
            add_geometry=add_geometry,
            filter_attributes_by_subset=filter_attributes_by_subset,
            dummy_data_rows=dummy_data_rows,
            model_directory=model_directory,
        )

        abbreviate_excel_after = False
        if kwargs.get('abbreviate_excel_sheettitles'):
            kwargs['abbreviate_excel_sheettitles'] = False
            abbreviate_excel_after = True

        # write the file
        await OtlmowConverter.from_objects_to_file_async(
            file_path=template_file_path,
            sequence_of_objects=objects,
            split_per_type=split_per_type,
            **kwargs,
        )

        # alter the file if needed
        extension = template_file_path.suffix.lower()
        if extension == '.xlsx':
            await cls.alter_excel_template_async(
                generate_choice_list=generate_choice_list,
                file_path=template_file_path,
                dummy_data_rows=dummy_data_rows,
                instances=objects,
                add_deprecated=add_deprecated,
                add_attribute_info=add_attribute_info,
                abbreviate_excel_sheettitles=abbreviate_excel_after,
            )

        elif extension == '.csv':
            await cls.alter_csv_template_async(
                split_per_type=split_per_type,
                file_path=template_file_path,
                dummy_data_rows=dummy_data_rows,
                instances=objects,
                add_deprecated=add_deprecated,
                add_attribute_info=add_attribute_info,
            )
