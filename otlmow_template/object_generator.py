"""The generator that creates the objects of a template from a subset file."""

import asyncio
import concurrent
import contextlib
import logging
from asyncio import sleep
from concurrent.futures import ALL_COMPLETED, Future, ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, cast

from otlmow_converter.DotnotationHelper import DotnotationHelper
from otlmow_model.OtlmowModel.BaseClasses.OTLObject import (
    OTLObject,
    dynamic_create_instance_from_uri,
    get_attribute_by_name,
)
from otlmow_model.OtlmowModel.Helpers.generated_lists import get_hardcoded_relation_dict
from otlmow_modelbuilder.OSLOCollector import OSLOCollector
from otlmow_modelbuilder.SQLDataClasses.OSLOClass import OSLOClass
from otlmow_template.constants import AGENT_TYPE_URI
from otlmow_template.relations import RelationsMixin


@dataclass
class GenerationContext:
    """The input of the generation of the objects of a template."""

    collector: OSLOCollector
    filtered_class_list: list[OSLOClass]
    add_geometry: bool
    filter_attributes_by_subset: bool
    dummy_data_rows: int
    model_directory: Optional[Path]


class ObjectGeneratorMixin(RelationsMixin):
    """The class methods that create the objects of a template."""

    # the uris of the classes that are a relation, set when the objects are generated
    relation_dict: dict[str, str]

    @classmethod
    def _load_collector_from_subset_path(cls, subset_path: Path) -> OSLOCollector:
        collector = OSLOCollector(subset_path)
        collector.collect_all(include_abstract=True)
        return collector

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
        class_uris_filter: Optional[list[str]],
        filter_attributes_by_subset: bool,
        dummy_data_rows: int,
        add_geometry: bool,
        ignore_relations: bool,
        model_directory: Optional[Path] = None,
    ) -> list[OTLObject]:
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
    def create_objects_with_unique_ids(cls, context: GenerationContext) -> list[OTLObject]:
        """Creates the objects of a template, until every object has its own id."""
        while True:
            otl_objects = cls.create_objects_in_parallel(context=context)
            if cls.has_unique_ids(otl_objects=otl_objects):
                return otl_objects

    @classmethod
    def create_objects_in_parallel(cls, context: GenerationContext) -> list[OTLObject]:
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
                futures = list(not_done)
                cls.report_progress(done=done, not_done=not_done)
        return otl_objects

    @classmethod
    def submit_object_creation(
        cls, executor, context: GenerationContext
    ) -> list[Future[list[OTLObject]]]:
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
    def get_classes_to_create(cls, context: GenerationContext) -> list[OSLOClass]:
        """Returns the concrete classes of a template."""
        return [cl for cl in context.filtered_class_list if cl.abstract == 0]

    @classmethod
    def report_progress(cls, done, not_done) -> None:
        """Reports how many objects are created and how many are still running."""
        if len(not_done) > 0:
            pass

    @classmethod
    def has_unique_ids(cls, otl_objects: list[OTLObject]) -> bool:
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
    def get_non_relation_class_uris(cls, filtered_class_list: list[OSLOClass]) -> list[str]:
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
        class_uris_filter: Optional[list[str]],
        filter_attributes_by_subset: bool,
        dummy_data_rows: int,
        add_geometry: bool,
        ignore_relations: bool,
        model_directory: Optional[Path] = None,
    ) -> list[OTLObject]:
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
    async def create_objects_with_unique_ids_async(
        cls, context: GenerationContext
    ) -> list[OTLObject]:
        """Creates the objects of a template, until every object has its own id."""
        while True:
            otl_objects = await cls.create_objects_in_parallel_async(context=context)
            if cls.has_unique_ids(otl_objects=otl_objects):
                return otl_objects

    @classmethod
    async def create_objects_in_parallel_async(cls, context: GenerationContext) -> list[OTLObject]:
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
    ) -> Optional[OTLObject]:
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
    def filters_classes_by_subset(
        cls, collector: OSLOCollector, class_uris_filter: Optional[list[str]] = None
    ) -> list[OSLOClass]:
        if class_uris_filter is None:
            return cast(list[OSLOClass], collector.classes)
        return [x for x in collector.classes if x.objectUri in class_uris_filter]
