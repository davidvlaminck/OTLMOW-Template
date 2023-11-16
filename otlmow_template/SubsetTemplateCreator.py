import logging
import site
from pathlib import Path
from typing import List

from otlmow_converter.OtlmowConverter import OtlmowConverter
from otlmow_model.Helpers.AssetCreator import dynamic_create_instance_from_uri
from otlmow_modelbuilder.OSLOCollector import OSLOCollector


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
            class_directory = None
            if kwargs is not None and 'class_directory' in kwargs:
                class_directory = kwargs['class_directory']
            instance = dynamic_create_instance_from_uri(class_object.objectUri, directory=class_directory)
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

    @classmethod
    def filters_assets_by_subset(cls, path_to_subset: Path, list_of_otl_objectUri: List):
        collector = cls._load_collector_from_subset_path(path_to_subset=path_to_subset)
        filtered_list = [x for x in collector.classes if x.objectUri in list_of_otl_objectUri]
        return filtered_list

    @staticmethod
    def _try_getting_settings_of_converter() -> Path:
        converter_path = Path(site.getsitepackages()[0]) / 'otlmow_converter'
        return converter_path / 'settings_otlmow_converter.json'
