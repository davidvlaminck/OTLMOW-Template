import asyncio
import concurrent
import contextlib
import csv
import logging
import os
from asyncio import sleep
from collections import defaultdict
from concurrent.futures import ALL_COMPLETED, ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from openpyxl.cell import Cell
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.worksheet.worksheet import Worksheet
from otlmow_converter.DotnotationHelper import DotnotationHelper
from otlmow_converter.Exceptions.UnknownExcelError import UnknownExcelError
from otlmow_converter.OtlmowConverter import OtlmowConverter
from otlmow_model.OtlmowModel.BaseClasses.OTLObject import (
    OTLAttribuut,
    OTLObject,
    dynamic_create_instance_from_uri,
    dynamic_create_type_from_uri,
    get_attribute_by_name,
)
from otlmow_model.OtlmowModel.Helpers.generated_lists import get_hardcoded_relation_dict
from otlmow_model.OtlmowModel.Helpers.GenericHelper import get_ns_and_name_from_uri
from otlmow_model.OtlmowModel.Helpers.RelationCreator import (
    create_betrokkenerelation,
    create_relation,
)
from otlmow_modelbuilder.OSLOCollector import OSLOCollector
from otlmow_modelbuilder.SQLDataClasses.OSLOClass import OSLOClass

ROOT_DIR = (Path(__file__).parent.parent).resolve()


SHEET_TITLE_HAS_NO_HASH = 'Sheet title does not contain a #'
HEADER_EXISTS_AT_COLUMN = 'Header already exists at column {column_nr}: {value}'

short_to_long_ns = {
    'ond': 'https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#',
    'onderdeel': 'https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#',
    'ins': 'https://wegenenverkeer.data.vlaanderen.be/ns/installatie#',
    'installatie': 'https://wegenenverkeer.data.vlaanderen.be/ns/installatie#',
    'imp': 'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#',
    'implementatieelement': 'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#',
    'proefenmeting': 'https://wegenenverkeer.data.vlaanderen.be/ns/proefenmeting#',
    'pro': 'https://wegenenverkeer.data.vlaanderen.be/ns/proefenmeting#',
    'lev': 'https://wegenenverkeer.data.vlaanderen.be/ns/levenscyclus#',
    'levenscyclus': 'https://wegenenverkeer.data.vlaanderen.be/ns/levenscyclus#',
}

AGENT_TYPE_URI = 'http://purl.org/dc/terms/Agent'
HEEFT_BETROKKENE_URI = 'https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#HeeftBetrokkene'
TYPE_URI_DESCRIPTION = 'De URI van het object volgens https://www.w3.org/2001/XMLSchema#anyURI .'
QUOTE_CHAR = '"'
# the validations are only added up to this row, so a template stays usable
MAX_VALIDATION_ROW = 1000


@dataclass
class AttributeInfo:
    """The information about the attributes of a template, collected per column.

    The Excel and the csv writer both add a row with the definition of every
    attribute and a row that marks the deprecated attributes, so the collection
    of those rows is shared.
    """

    add_attribute_info: bool
    add_deprecated: bool
    attribute_info_row: list[str] = field(default_factory=list)
    deprecated_row: list[str] = field(default_factory=list)

    def add_attribute(self, instance, header: str, attribute) -> None:
        """Add the definition and the deprecated marker of an attribute."""
        if self.add_attribute_info:
            self.attribute_info_row.append(
                SubsetTemplateCreator.get_attribute_info(
                    instance=instance, header=header, attribute=attribute
                )
            )
        if self.add_deprecated:
            marker = 'DEPRECATED' if attribute.deprecated_version else ''
            self.deprecated_row.append(marker)

    def add_type_uri(self) -> None:
        """Add the definition and the deprecated marker of the typeURI column."""
        if self.add_attribute_info:
            self.attribute_info_row.append(TYPE_URI_DESCRIPTION)
        if self.add_deprecated:
            self.deprecated_row.append('')

    def add_missing_attribute(self) -> None:
        """Add the definition and the deprecated marker of a column without attribute."""
        if self.add_attribute_info:
            self.attribute_info_row.append('')
        if self.add_deprecated:
            self.deprecated_row.append('')


@dataclass
class ExcelSheetContext:
    """The state of a sheet while its headers are processed."""

    sheet: Worksheet
    workbook: Workbook
    instance: OTLObject
    type_uri: str
    first_data_row: int
    boolean_validation: DataValidation
    generate_choice_list: bool
    choice_list_dict: dict[str, str]
    info: AttributeInfo
    string_columns: list[int] = field(default_factory=list)


@dataclass
class GenerationContext:
    """The input of the generation of the objects of a template."""

    collector: OSLOCollector
    filtered_class_list: list[OSLOClass]
    add_geometry: bool
    filter_attributes_by_subset: bool
    dummy_data_rows: int
    model_directory: Path


class SubsetTemplateCreator:
    green_fill = PatternFill(start_color='90EE90', fill_type='solid')

    @classmethod
    def _load_collector_from_subset_path(cls, subset_path: Path) -> OSLOCollector:
        collector = OSLOCollector(subset_path)
        collector.collect_all(include_abstract=True)
        return collector

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
    def create_x_objects(
        cls,
        oslo_class,
        add_geometry,
        collector,
        filter_attributes_by_subset,
        model_directory,
        amount_objects_to_create,
    ):
        if oslo_class.objectUri in cls.relation_dict:
            return []

        otl_objects = []
        for _ in range(amount_objects_to_create):
            otl_object = cls.generate_object_from_oslo_class(
                oslo_class=oslo_class,
                add_geometry=add_geometry,
                collector=collector,
                filter_attributes_by_subset=filter_attributes_by_subset,
                model_directory=model_directory,
            )
            if otl_object is not None:
                otl_objects.append(otl_object)
        return otl_objects

    @classmethod
    def generate_objects_for_template(
        cls,
        subset_path: Path,
        class_uris_filter: [str],
        filter_attributes_by_subset: bool,
        dummy_data_rows: int,
        add_geometry: bool,
        ignore_relations: bool,
        model_directory: Optional[Path] = None,
    ) -> [OTLObject]:
        """
        This method is used to generate objects for the template. It will generate objects based on the subset file
        """
        collector = cls._load_collector_from_subset_path(subset_path=subset_path)
        filtered_class_list = cls.filters_classes_by_subset(
            collector=collector, class_uris_filter=class_uris_filter
        )
        if filtered_class_list == []:
            raise ValueError('Something went wrong, as the class_uri filter list is empty')
        cls.relation_dict = get_hardcoded_relation_dict(model_directory=model_directory)

        context = GenerationContext(
            collector=collector,
            filtered_class_list=filtered_class_list,
            add_geometry=add_geometry,
            filter_attributes_by_subset=filter_attributes_by_subset,
            dummy_data_rows=dummy_data_rows,
            model_directory=model_directory,
        )
        otl_objects = cls.create_objects_with_unique_ids(context=context)

        if not ignore_relations:
            cls.append_relations_to_objects(
                otl_objects=otl_objects,
                collector=collector,
                class_uris_filter=cls.get_non_relation_class_uris(filtered_class_list),
                model_directory=model_directory,
            )

        return otl_objects

    @classmethod
    def create_objects_with_unique_ids(cls, context: GenerationContext) -> list:
        """Creates the objects of a template, until every object has its own id."""
        while True:
            otl_objects = cls.create_objects_in_parallel(context=context)
            if cls.has_unique_ids(otl_objects=otl_objects):
                return otl_objects

    @classmethod
    def create_objects_in_parallel(cls, context: GenerationContext) -> list:
        """Creates the objects of every class of a template, in parallel."""
        otl_objects = []
        with ThreadPoolExecutor() as executor:
            futures = cls.submit_object_creation(executor=executor, context=context)
            while futures:
                done, not_done = concurrent.futures.wait(
                    futures, return_when=ALL_COMPLETED, timeout=60
                )
                for future in as_completed(done):
                    otl_objects.extend(future.result())
                futures = not_done
                cls.report_progress(done=done, not_done=not_done)
        return otl_objects

    @classmethod
    def submit_object_creation(cls, executor, context: GenerationContext) -> list:
        """Submits the creation of the objects of every class of a template."""
        amount_objects_to_create = max(1, context.dummy_data_rows)
        return [
            executor.submit(
                cls.create_x_objects,
                oslo_class,
                context.add_geometry,
                context.collector,
                context.filter_attributes_by_subset,
                context.model_directory,
                amount_objects_to_create,
            )
            for oslo_class in cls.get_classes_to_create(context=context)
        ]

    @classmethod
    def get_classes_to_create(cls, context: GenerationContext) -> list:
        """Returns the concrete classes of a template."""
        return [cl for cl in context.filtered_class_list if cl.abstract == 0]

    @classmethod
    def report_progress(cls, done, not_done) -> None:
        """Reports how many objects are created and how many are still running."""
        if len(not_done) > 0:
            pass

    @classmethod
    def has_unique_ids(cls, otl_objects: list) -> bool:
        """Returns whether every object of a template has its own id."""
        created = len(otl_objects)
        unique_ids = len(
            {
                obj.assetId.identificator
                if obj.typeURI != AGENT_TYPE_URI
                else obj.agentId.identificator
                for obj in otl_objects
            }
        )
        return created == unique_ids

    @classmethod
    def get_non_relation_class_uris(cls, filtered_class_list: list) -> list:
        """Returns the uris of the classes of a template that are not a relation."""
        return [
            cl.objectUri
            for cl in filtered_class_list
            if cl.abstract == 0 and cl.objectUri not in cls.relation_dict
        ]

    @classmethod
    async def create_x_objects_async(
        cls,
        oslo_class,
        add_geometry,
        collector,
        filter_attributes_by_subset,
        model_directory,
        amount_objects_to_create,
    ):
        if oslo_class.objectUri in cls.relation_dict:
            return []

        otl_objects = []
        for _ in range(amount_objects_to_create):
            await sleep(0)
            otl_object = cls.generate_object_from_oslo_class(
                oslo_class=oslo_class,
                add_geometry=add_geometry,
                collector=collector,
                filter_attributes_by_subset=filter_attributes_by_subset,
                model_directory=model_directory,
            )
            if otl_object is not None:
                otl_objects.append(otl_object)
        return otl_objects

    @classmethod
    async def generate_objects_for_template_async(
        cls,
        subset_path: Path,
        class_uris_filter: [str],
        filter_attributes_by_subset: bool,
        dummy_data_rows: int,
        add_geometry: bool,
        ignore_relations: bool,
        model_directory: Optional[Path] = None,
    ) -> [OTLObject]:
        """
        This method is used to generate objects for the template. It will generate objects based on the subset file
        """
        await sleep(0)
        collector = cls._load_collector_from_subset_path(subset_path=subset_path)
        await sleep(0)
        filtered_class_list = cls.filters_classes_by_subset(
            collector=collector, class_uris_filter=class_uris_filter
        )
        if filtered_class_list == []:
            raise ValueError('Something went wrong, as the class_uri filter list is empty')
        await sleep(0)
        cls.relation_dict = get_hardcoded_relation_dict(model_directory=model_directory)

        context = GenerationContext(
            collector=collector,
            filtered_class_list=filtered_class_list,
            add_geometry=add_geometry,
            filter_attributes_by_subset=filter_attributes_by_subset,
            dummy_data_rows=dummy_data_rows,
            model_directory=model_directory,
        )
        otl_objects = await cls.create_objects_with_unique_ids_async(context=context)

        if not ignore_relations:
            cls.append_relations_to_objects(
                otl_objects=otl_objects,
                collector=collector,
                class_uris_filter=cls.get_non_relation_class_uris(filtered_class_list),
                model_directory=model_directory,
            )

        return otl_objects

    @classmethod
    async def create_objects_with_unique_ids_async(cls, context: GenerationContext) -> list:
        """Creates the objects of a template, until every object has its own id."""
        while True:
            otl_objects = await cls.create_objects_in_parallel_async(context=context)
            if cls.has_unique_ids(otl_objects=otl_objects):
                return otl_objects

    @classmethod
    async def create_objects_in_parallel_async(cls, context: GenerationContext) -> list:
        """Creates the objects of every class of a template, at the same time."""
        amount_objects_to_create = max(1, context.dummy_data_rows)
        tasks = [
            cls.create_x_objects_async(
                oslo_class,
                context.add_geometry,
                context.collector,
                context.filter_attributes_by_subset,
                context.model_directory,
                amount_objects_to_create,
            )
            for oslo_class in cls.get_classes_to_create(context=context)
        ]
        results = await asyncio.gather(*tasks)
        return [item for sublist in results for item in sublist]

    @classmethod
    def generate_object_from_oslo_class(
        cls,
        oslo_class: OSLOClass,
        add_geometry: bool,
        filter_attributes_by_subset: bool,
        collector: OSLOCollector,
        model_directory: Optional[Path] = None,
    ) -> [OTLObject]:
        """
        Generate an object from a given OSLO class
        """
        instance = dynamic_create_instance_from_uri(
            oslo_class.objectUri, model_directory=model_directory
        )
        if instance is None:
            return None

        if filter_attributes_by_subset:
            cls.fill_attributes_of_subset(
                instance=instance, collector=collector, oslo_class=oslo_class
            )
        else:
            cls.fill_all_attributes(instance=instance)

        cls.fill_asset_id(instance=instance, add_geometry=add_geometry)
        cls.clear_asset_version(instance=instance)

        DotnotationHelper.clear_list_of_list_attributes(instance)

        return instance

    @classmethod
    def fill_attributes_of_subset(
        cls, instance, collector: OSLOCollector, oslo_class: OSLOClass
    ) -> None:
        """Fills the attributes that are in the subset with dummy data."""
        for attribute_object in collector.find_attributes_by_class(oslo_class):
            attr = get_attribute_by_name(instance, attribute_object.name)
            if attr is None:
                logging.warning(
                    f'Attribute {attribute_object.name} not found in class {oslo_class.objectUri}'
                )
            elif attr.naam == 'isActief':
                attr.set_waarde(True)
            else:
                attr.fill_with_dummy_data()

    @classmethod
    def fill_all_attributes(cls, instance) -> None:
        """Fills every attribute of an object with dummy data."""
        for attr in instance:
            if attr.naam == 'isActief':
                attr.set_waarde(True)
            elif attr.naam != 'geometry':
                attr.fill_with_dummy_data()

    @classmethod
    def fill_asset_id(cls, instance, add_geometry: bool) -> None:
        """Fills the asset id and the geometry of an object with dummy data."""
        with contextlib.suppress(AttributeError):
            assetId_attr = get_attribute_by_name(instance, 'assetId')
            if assetId_attr is not None:
                assetId_attr.fill_with_dummy_data()
            if add_geometry:
                geo_attr = get_attribute_by_name(instance, 'geometry')
                if geo_attr is not None:
                    geo_attr.fill_with_dummy_data()

    @classmethod
    def clear_asset_version(cls, instance) -> None:
        """Clears the asset version of an object."""
        asset_versie = get_attribute_by_name(instance, 'assetVersie')
        if asset_versie is not None:
            asset_versie.set_waarde(None)

    @classmethod
    def alter_excel_template(
        cls,
        instances: list,
        file_path: Path,
        add_attribute_info: bool,
        generate_choice_list: bool,
        dummy_data_rows: int,
        add_deprecated: bool,
        abbreviate_excel_sheettitles: bool,
    ):
        wb = load_workbook(file_path)
        wb.create_sheet('Keuzelijsten')

        choice_list_dict = {}
        for sheet in wb:
            if sheet.title == 'Keuzelijsten':
                break

            cls.alter_excel_sheet(
                add_attribute_info=add_attribute_info,
                choice_list_dict=choice_list_dict,
                generate_choice_list=generate_choice_list,
                dummy_data_rows=dummy_data_rows,
                instances=instances,
                sheet=sheet,
                add_deprecated=add_deprecated,
                workbook=wb,
                abbreviate_excel_sheettitle=abbreviate_excel_sheettitles,
            )

        wb.save(file_path)
        wb.close()

    @classmethod
    def fill_class_dict(cls, instances: list) -> dict:
        class_dict = defaultdict(list)
        for instance in instances:
            class_dict[instance.typeURI].append(instance)
        return class_dict

    @classmethod
    def alter_csv_template(
        cls,
        instances: list,
        file_path: Path,
        add_attribute_info: bool,
        split_per_type: bool,
        dummy_data_rows: int,
        add_deprecated: bool,
    ):
        classes_dict = cls.fill_class_dict(instances)
        if split_per_type:
            for type_uri, typed_instances in classes_dict.items():
                ns, name = get_ns_and_name_from_uri(type_uri)
                ns_name = f'{ns}_{name}'
                if name == 'Agent':
                    ns_name = 'Agent'
                class_file_path = file_path.parent / f'{file_path.stem}_{ns_name}.csv'
                cls.alter_csv_file(
                    add_attribute_info=add_attribute_info,
                    add_deprecated=add_deprecated,
                    dummy_data_rows=dummy_data_rows,
                    instances=typed_instances,
                    file_path=class_file_path,
                )
        else:
            cls.alter_csv_file(
                add_attribute_info=add_attribute_info,
                add_deprecated=add_deprecated,
                dummy_data_rows=dummy_data_rows,
                instances=instances,
                file_path=file_path,
            )

    @classmethod
    async def alter_csv_template_async(
        cls,
        instances: list,
        file_path: Path,
        add_deprecated: bool,
        add_attribute_info: bool,
        split_per_type: bool,
        dummy_data_rows: int,
    ):
        classes_dict = cls.fill_class_dict(instances)
        if split_per_type:
            for type_uri, typed_instances in classes_dict.items():
                await sleep(0)
                ns, name = get_ns_and_name_from_uri(type_uri)
                ns_name = f'{ns}_{name}'
                if name == 'Agent':
                    ns_name = 'Agent'
                class_file_path = file_path.parent / f'{file_path.stem}_{ns_name}.csv'
                cls.alter_csv_file(
                    add_attribute_info=add_attribute_info,
                    dummy_data_rows=dummy_data_rows,
                    instances=typed_instances,
                    add_deprecated=add_deprecated,
                    file_path=class_file_path,
                )
        else:
            await sleep(0)
            cls.alter_csv_file(
                add_attribute_info=add_attribute_info,
                dummy_data_rows=dummy_data_rows,
                instances=instances,
                add_deprecated=add_deprecated,
                file_path=file_path,
            )

    @classmethod
    def alter_csv_file(
        cls,
        add_attribute_info: bool,
        instances: [OTLObject],
        add_deprecated: bool,
        file_path: Path,
        dummy_data_rows: int,
    ):
        info = AttributeInfo(add_attribute_info=add_attribute_info, add_deprecated=add_deprecated)
        instance = instances[0]

        header_row, csv_data = cls.read_csv_file(file_path=file_path)
        cls.collect_csv_header_info(instance=instance, info=info, header_row=header_row)
        cls.write_csv_file(
            info=info,
            file_path=file_path,
            header_row=header_row,
            csv_data=csv_data,
            dummy_data_rows=dummy_data_rows,
        )

    @classmethod
    def read_csv_file(cls, file_path: Path) -> tuple[list, list]:
        """Reads the header row and the data rows of a csv file."""
        with Path(file_path).open(encoding='utf-8') as file:
            csv_reader = csv.reader(file, delimiter=';', quotechar=QUOTE_CHAR)
            header_row = next(csv_reader)
            csv_data = list(csv_reader)
        return header_row, csv_data

    @classmethod
    def collect_csv_header_info(cls, instance, info: AttributeInfo, header_row: list) -> None:
        """Collects the information of every column of a csv file."""
        for header in header_row:
            if not header:
                continue
            if header == 'typeURI':
                info.add_type_uri()
            else:
                cls.collect_csv_attribute_info(instance=instance, info=info, header=header)

    @classmethod
    def collect_csv_attribute_info(cls, instance, info: AttributeInfo, header: str) -> None:
        """Collects the information of one attribute column of a csv file."""
        try:
            attribute = DotnotationHelper.get_attribute_by_dotnotation(instance, header)
        except AttributeError:
            info.add_missing_attribute()
            return
        info.add_attribute(instance=instance, header=header, attribute=attribute)

    @classmethod
    def write_csv_file(
        cls,
        info: AttributeInfo,
        file_path: Path,
        header_row: list,
        csv_data: list,
        dummy_data_rows: int,
    ) -> None:
        """Writes the information rows and the data of a csv file."""
        with Path(file_path).open('w', newline='', encoding='utf-8') as file:
            csv_writer = csv.writer(
                file, delimiter=';', quotechar=QUOTE_CHAR, quoting=csv.QUOTE_MINIMAL
            )
            cls.write_csv_info_rows(info=info, csv_writer=csv_writer)
            csv_writer.writerow(header_row)
            cls.write_csv_data_rows(
                csv_writer=csv_writer, csv_data=csv_data, dummy_data_rows=dummy_data_rows
            )

    @classmethod
    def write_csv_info_rows(cls, info: AttributeInfo, csv_writer) -> None:
        """Writes the attribute info row and the deprecated row of a csv file."""
        if info.add_attribute_info:
            csv_writer.writerow(info.attribute_info_row)
        if info.add_deprecated and any(info.deprecated_row):
            csv_writer.writerow(info.deprecated_row)

    @classmethod
    def write_csv_data_rows(cls, csv_writer, csv_data: list, dummy_data_rows: int) -> None:
        """Writes the data rows of a csv file, unless the template has no data rows."""
        if dummy_data_rows == 0:
            return
        for line in csv_data:
            csv_writer.writerow(line)

    @classmethod
    async def alter_excel_template_async(
        cls,
        instances: list,
        file_path: Path,
        add_attribute_info: bool,
        generate_choice_list: bool,
        dummy_data_rows: int,
        add_deprecated: bool,
        abbreviate_excel_sheettitles: bool,
    ):
        wb = load_workbook(file_path)
        wb.create_sheet('Keuzelijsten')

        choice_list_dict = {}
        for sheet in wb:
            if sheet.title == 'Keuzelijsten':
                break

            cls.alter_excel_sheet(
                add_attribute_info=add_attribute_info,
                choice_list_dict=choice_list_dict,
                generate_choice_list=generate_choice_list,
                dummy_data_rows=dummy_data_rows,
                instances=instances,
                sheet=sheet,
                add_deprecated=add_deprecated,
                workbook=wb,
                abbreviate_excel_sheettitle=abbreviate_excel_sheettitles,
            )
            await sleep(0)

        wb.save(file_path)
        wb.close()

    @classmethod
    def alter_excel_sheet(
        cls,
        add_attribute_info: bool,
        choice_list_dict: dict,
        generate_choice_list: bool,
        instances: [OTLObject],
        sheet: Worksheet,
        add_deprecated: bool,
        workbook: Workbook,
        dummy_data_rows: int,
        abbreviate_excel_sheettitle: bool = True,
    ):
        type_uri = cls.get_uri_from_sheet_name(sheet.title)
        instance = cls.get_instance_for_type_uri(instances=instances, type_uri=type_uri)

        boolean_validation = DataValidation(
            type='list', formula1='"TRUE,FALSE,-"', allow_blank=True
        )
        sheet.add_data_validation(boolean_validation)

        context = ExcelSheetContext(
            sheet=sheet,
            workbook=workbook,
            instance=instance,
            type_uri=type_uri,
            first_data_row=cls.get_first_data_row(add_attribute_info=add_attribute_info),
            boolean_validation=boolean_validation,
            generate_choice_list=generate_choice_list,
            choice_list_dict=choice_list_dict,
            info=AttributeInfo(
                add_attribute_info=add_attribute_info, add_deprecated=add_deprecated
            ),
        )

        for header_cell in next(sheet.iter_rows(min_row=1, max_row=1)):
            cls.process_header_cell(context=context, header_cell=header_cell)

        cls.format_string_columns(context=context)
        cls.delete_dummy_rows(context=context, instances=instances, dummy_data_rows=dummy_data_rows)
        cls.add_deprecated_row(context=context)
        cls.add_attribute_info_row(context=context)

        cls.set_fixed_column_width(sheet=sheet, width=25)

        if abbreviate_excel_sheettitle:
            cls.abbreviate_excel_sheettitle(sheet=sheet)

    @classmethod
    def get_instance_for_type_uri(cls, instances: list, type_uri: str):
        """Returns the instance of the type of a sheet, raises when there is none."""
        instance = cls.find_instance_with_type_uri(instances=instances, type_uri=type_uri)
        if instance is None:
            instance = cls.find_instance_with_type_uri_prefix(
                instances=instances, type_uri=type_uri
            )
        if instance is None:
            raise UnknownExcelError(
                f'When creating a template, no instance could be created for {type_uri}'
            )
        return instance

    @classmethod
    def find_instance_with_type_uri(cls, instances: list, type_uri: str):
        """Returns the instance with exactly the given type uri, None when there is none."""
        for instance in instances:
            if instance.typeURI == type_uri:
                return instance
        return None

    @classmethod
    def find_instance_with_type_uri_prefix(cls, instances: list, type_uri: str):
        """Returns the instance whose type uri starts with the given one, None when there is none."""
        for instance in instances:
            if instance.typeURI.startswith(type_uri):
                return instance
        return None

    @classmethod
    def get_first_data_row(cls, add_attribute_info: bool) -> int:
        """Returns the first row of a sheet that holds data.

        The first row is the header row. The attribute info row is added in front
        of it, so the data starts one row lower. The deprecated row is added
        after the validations are applied, so it does not move the data.
        """
        first_data_row = 2
        if add_attribute_info:
            first_data_row += 1
        return first_data_row

    @classmethod
    def process_header_cell(cls, context: ExcelSheetContext, header_cell) -> None:
        """Adds the validations and the information of one header cell to the sheet."""
        header = header_cell.value
        if not header:
            return
        if header == 'typeURI':
            cls.process_type_uri_header(context=context, header_cell=header_cell)
            return
        if cls.is_agent_asset_id_header(context=context, header=header):
            return
        cls.process_attribute_header(context=context, header_cell=header_cell, header=header)

    @classmethod
    def is_agent_asset_id_header(cls, context: ExcelSheetContext, header: str) -> bool:
        """Returns whether a header is the asset id of the agent sheet."""
        return context.type_uri == AGENT_TYPE_URI and header.startswith('assetId.')

    @classmethod
    def process_type_uri_header(cls, context: ExcelSheetContext, header_cell) -> None:
        """Adds the validation that fixes the type of the objects of a sheet."""
        data_validation = DataValidation(
            type='list', formula1=f'"{context.type_uri}"', allow_blank=True
        )
        context.sheet.add_data_validation(data_validation)
        data_validation.add(
            cls.get_validation_range(header_cell=header_cell, first_data_row=context.first_data_row)
        )
        context.info.add_type_uri()

    @classmethod
    def process_attribute_header(cls, context: ExcelSheetContext, header_cell, header: str) -> None:
        """Adds the validations and the information of one attribute column to the sheet."""
        attribute = DotnotationHelper.get_attribute_by_dotnotation(context.instance, header)

        context.info.add_attribute(instance=context.instance, header=header, attribute=attribute)
        cls.add_choice_list_validation(
            context=context, attribute=attribute, header_cell=header_cell
        )

        if attribute.field.native_type is str:
            context.string_columns.append(header_cell.column)

    @classmethod
    def add_choice_list_validation(cls, context: ExcelSheetContext, attribute, header_cell) -> None:
        """Adds the validation of a boolean or a choice list column."""
        if not context.generate_choice_list:
            return
        if attribute.field.native_type is bool:
            cls.add_boolean_validation(context=context, header_cell=header_cell)
        elif hasattr(attribute.field, 'options') and attribute.field.options is not None:
            cls.add_option_validation(context=context, attribute=attribute, header_cell=header_cell)

    @classmethod
    def add_boolean_validation(cls, context: ExcelSheetContext, header_cell) -> None:
        """Restricts a boolean column to the values of a boolean list."""
        context.boolean_validation.add(
            cls.get_validation_range(header_cell=header_cell, first_data_row=context.first_data_row)
        )
        cls.color_choice_lists_green(sheet=context.sheet, header_cell_column=header_cell)

    @classmethod
    def add_option_validation(cls, context: ExcelSheetContext, attribute, header_cell) -> None:
        """Restricts a column to the values of its choice list."""
        cls.generate_choice_list_in_excel(
            attribute=attribute,
            choice_list_dict=context.choice_list_dict,
            column=header_cell.column,
            row_nr=1,
            sheet=context.sheet,
            workbook=context.workbook,
        )
        cls.color_choice_lists_green(sheet=context.sheet, header_cell_column=header_cell)

    @classmethod
    def get_validation_range(cls, header_cell, first_data_row: int) -> str:
        """Returns the range of the rows a validation is added to."""
        column_letter = header_cell.column_letter
        return f'{column_letter}{first_data_row}:{column_letter}{MAX_VALIDATION_ROW}'

    @classmethod
    def format_string_columns(cls, context: ExcelSheetContext) -> None:
        """Formats the string columns of a sheet as text, so Excel keeps the leading zeros."""
        for col_idx in context.string_columns:
            for row_idx in range(context.first_data_row, MAX_VALIDATION_ROW + 1):
                context.sheet.cell(row=row_idx, column=col_idx).number_format = '@'

    @classmethod
    def delete_dummy_rows(
        cls, context: ExcelSheetContext, instances: list, dummy_data_rows: int
    ) -> None:
        """Deletes the dummy rows of a template without data rows."""
        if dummy_data_rows != 0:
            return
        instance_count = len([x for x in instances if x.typeURI == context.type_uri])
        if instance_count > 0:
            context.sheet.delete_rows(idx=2, amount=instance_count)

    @classmethod
    def add_deprecated_row(cls, context: ExcelSheetContext) -> None:
        """Adds the row that marks the deprecated attributes of a sheet."""
        if context.info.add_deprecated and any(context.info.deprecated_row):
            cls.add_deprecated_row_to_sheet(context.info.deprecated_row, context.sheet)

    @classmethod
    def add_attribute_info_row(cls, context: ExcelSheetContext) -> None:
        """Adds the row with the definition of the attributes of a sheet."""
        if context.info.add_attribute_info:
            cls.add_attribute_info_to_sheet(context.info.attribute_info_row, context.sheet)

    @classmethod
    def color_choice_lists_green(cls, sheet: Worksheet, header_cell_column: Cell):
        for cell in sheet.iter_rows(
            min_col=header_cell_column.column,
            max_col=header_cell_column.column,
            min_row=2,
            max_row=MAX_VALIDATION_ROW,
        ):
            cell[0].fill = cls.green_fill

    @classmethod
    def add_deprecated_row_to_sheet(cls, deprecated_attributes_row, sheet):
        sheet.insert_rows(idx=1)
        for index, depr_info in enumerate(deprecated_attributes_row, start=1):
            if depr_info != '':
                cell = sheet.cell(row=1, column=index)
                cell.value = depr_info
                cell.fill = PatternFill(start_color='FF7276', end_color='FF7276', fill_type='solid')

    @classmethod
    def add_attribute_info_to_sheet(cls, collected_attribute_info: list[str], sheet: Worksheet):
        sheet.insert_rows(idx=1)
        fill = PatternFill(start_color='808080', fill_type='solid')
        alignment = Alignment(wrapText=True, vertical='top')
        for index, attr_info in enumerate(collected_attribute_info, start=1):
            cell = sheet.cell(row=1, column=index, value=attr_info)
            cell.alignment = alignment
            cell.fill = fill

    @classmethod
    def get_attribute_info(cls, instance: OTLObject, header: str, attribute: OTLAttribuut) -> str:
        """Returns the definition to show in the attribute info row for a column of a template.

        The definition of the 'waarde' attribute of a quantitative value is always the same and does not explain
        the attribute itself. For quantitative values the definition of the attribute itself is combined with the
        standard unit of the datatype.
        """
        quantitative_attribute = cls.get_quantitative_attribute(instance=instance, header=header)
        if quantitative_attribute is None:
            return attribute.definition
        standard_unit = cls.get_standard_unit(quantitative_attribute)
        if standard_unit is None:
            return attribute.definition
        definition = quantitative_attribute.definition.strip()
        if not definition.endswith('.'):
            definition = f'{definition}.'
        return f'{definition} Standaard eenheid: {standard_unit}'

    @classmethod
    def get_quantitative_attribute(cls, instance: OTLObject, header: str) -> OTLAttribuut:
        """Returns the attribute a column header refers to when it points to the value of a quantitative value,
        None when the header does not point to a quantitative value."""
        attribute = DotnotationHelper.get_attribute_by_dotnotation(
            instance_or_attribute=instance, dotnotation=header, waarde_shortcut=False
        )
        if not attribute.field.waarde_shortcut_applicable:
            return None
        if not hasattr(attribute.field.waardeObject, 'standaardEenheid'):
            return None
        return attribute

    @classmethod
    def get_standard_unit(cls, attribute: OTLAttribuut) -> str:
        """Returns the standard unit (ex. 'mo') of a quantitative value attribute,
        None when the standard unit of the datatype is unknown."""
        waarde = attribute.waarde
        if waarde is None:
            attribute.add_empty_value()
            waarde = attribute.waarde
        if isinstance(waarde, list):
            waarde = next(iter(waarde), None)
        standaard_eenheid = get_attribute_by_name(waarde, 'standaardEenheid')
        if standaard_eenheid is None:
            return None
        return cls.find_standard_unit(standaard_eenheid=standaard_eenheid)

    @classmethod
    def find_standard_unit(cls, standaard_eenheid) -> str:
        """Returns the standard unit in the notations of a datatype,
        None when the notations do not contain one."""
        for notation in [standaard_eenheid.usagenote, standaard_eenheid.constraints]:
            if '"' in notation:
                return notation.split('"')[1]
        return None

    @classmethod
    def generate_choice_list_in_excel(
        cls, attribute, choice_list_dict, column, row_nr, sheet: Worksheet, workbook: Workbook
    ):
        choice_list_values = [
            cv for cv in attribute.field.options.values() if cv.status != 'verwijderd'
        ]
        if attribute.field.naam not in choice_list_dict:
            cls.add_choice_list_to_sheet(
                workbook=workbook,
                name=attribute.field.naam,
                options=choice_list_values,
                choice_list_dict=choice_list_dict,
            )
        column_in_choice_sheet = choice_list_dict[attribute.field.naam]
        start_range = f'${column_in_choice_sheet}$2'
        end_range = f'${column_in_choice_sheet}${len(choice_list_values) + 2}'
        data_val = DataValidation(
            type='list', formula1=f'Keuzelijsten!{start_range}:{end_range}', allowBlank=True
        )
        sheet.add_data_validation(data_val)
        data_val.add(
            f'{get_column_letter(column)}{row_nr + 1}:'
            f'{get_column_letter(column)}{MAX_VALIDATION_ROW}'
        )

    @classmethod
    def filters_classes_by_subset(
        cls, collector: OSLOCollector, class_uris_filter: [str] = None
    ) -> list[OSLOClass]:
        if class_uris_filter is None:
            return collector.classes
        return [x for x in collector.classes if x.objectUri in class_uris_filter]

    @classmethod
    def set_fixed_column_width(cls, sheet, width: int):
        dim_holder = DimensionHolder(worksheet=sheet)
        for col in range(sheet.min_column, sheet.max_column + 1):
            dim_holder[get_column_letter(col)] = ColumnDimension(
                sheet, min=col, max=col, width=width
            )
        sheet.column_dimensions = dim_holder

    @classmethod
    def add_choice_list_to_sheet(cls, workbook, name, options, choice_list_dict):
        active_sheet = workbook['Keuzelijsten']
        column_nr = choice_list_dict.keys().__len__() + 1
        row_nr = 1
        new_header = active_sheet.cell(row=row_nr, column=column_nr)
        if new_header.value is not None:
            raise ValueError(HEADER_EXISTS_AT_COLUMN.format(column_nr=column_nr, value=new_header.value))
        new_header.value = name
        cell = active_sheet.cell(row=row_nr + 1, column=column_nr)
        cell.value = '-'
        for index, option in enumerate(options, start=2):
            cell = active_sheet.cell(row=row_nr + index, column=column_nr)
            cell.value = option.invulwaarde

        choice_list_dict[name] = new_header.column_letter

    @classmethod
    def get_uri_from_sheet_name(cls, title: str) -> str:
        if title == 'Agent':
            return 'http://purl.org/dc/terms/Agent'
        if '#' not in title:
            raise ValueError(SHEET_TITLE_HAS_NO_HASH)
        class_ns, class_name = title.split('#', maxsplit=1)
        return short_to_long_ns.get(class_ns, class_ns) + class_name

    @classmethod
    def append_relations_to_objects(
        cls,
        otl_objects: [OTLObject],
        collector: OSLOCollector,
        class_uris_filter: [str],
        model_directory: Path,
    ):
        class_dict = cls.fill_class_dict(otl_objects)

        for class_uri in class_uris_filter:
            cls.append_relations_of_class(
                class_dict=class_dict,
                otl_objects=otl_objects,
                class_uri=class_uri,
                collector=collector,
                model_directory=model_directory,
            )

    @classmethod
    def append_relations_of_class(
        cls,
        class_dict: dict,
        otl_objects: [OTLObject],
        class_uri: str,
        collector: OSLOCollector,
        model_directory: Path,
    ) -> None:
        """Adds the relations of one class to the objects of a template."""
        for relation in collector.find_all_concrete_relations(objectUri=class_uri):
            if class_uri != relation.bron_uri:
                continue
            cls.append_relations_of_relation(
                class_dict=class_dict,
                otl_objects=otl_objects,
                relation=relation,
                model_directory=model_directory,
            )

    @classmethod
    def append_relations_of_relation(
        cls, class_dict: dict, otl_objects: [OTLObject], relation, model_directory: Path
    ) -> None:
        """Adds the relations of one relation type to the objects of a template."""
        for index, bron_instance in enumerate(class_dict[relation.bron_uri]):
            if relation.doel_uri not in class_dict:
                continue
            doel_instance = class_dict[relation.doel_uri][index]
            relation_instance = cls.create_relation_instance(
                relation=relation,
                bron_instance=bron_instance,
                doel_instance=doel_instance,
                model_directory=model_directory,
            )
            if relation_instance is not None:
                otl_objects.append(relation_instance)

    @classmethod
    def create_relation_instance(
        cls, relation, bron_instance, doel_instance, model_directory: Path
    ):
        """Creates the relation between two objects, None when the relation is not created."""
        if relation.objectUri == HEEFT_BETROKKENE_URI:
            return create_betrokkenerelation(
                rol='toezichter',
                source=bron_instance,
                target=doel_instance,
                model_directory=model_directory,
            )
        if (
            relation.richting == 'Unspecified'
            and bron_instance.assetId.identificator > doel_instance.assetId.identificator
        ):
            return None
        relation_type = dynamic_create_type_from_uri(
            class_uri=relation.objectUri, model_directory=model_directory
        )
        return create_relation(
            relation_type=relation_type,
            source=bron_instance,
            target=doel_instance,
            model_directory=model_directory,
        )

    @classmethod
    def abbreviate_excel_sheettitle(cls, sheet):
        class_name = sheet.title
        # abbreviates the class_name so it doesn't exceed the 31 character limit of sheet titles in excel
        if '#' not in class_name:
            sheet.title = class_name[:31]
            return
        split_name = class_name.split('#')
        namespace_name = split_name[0]
        subclass_name = split_name[1]
        class_name = f'{namespace_name[:3]}#{subclass_name}'
        sheet.title = class_name[:31]
