import ntpath
import os
import site
import tempfile
from pathlib import Path
from otlmow_converter.OtlmowConverter import OtlmowConverter
from otlmow_model.OtlmowModel.Helpers.AssetCreator import dynamic_create_instance_from_uri
from otlmow_modelbuilder.OSLOCollector import OSLOCollector

from otlmow_template.CsvTemplateCreator import CsvTemplateCreator
from otlmow_template.ExcelTemplateCreator import ExcelTemplateCreator

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
        temporary_path = self.return_temp_path(path_to_template_file_and_extension=path_to_template_file_and_extension)
        instantiated_attributes = self.generate_basic_template(path_to_subset=path_to_subset,
                                                               temporary_path=temporary_path,
                                                               path_to_template_file_and_extension=path_to_template_file_and_extension,
                                                               **kwargs)
        extension = os.path.splitext(path_to_template_file_and_extension)[-1].lower()
        if extension == '.xlsx':
            ExcelTemplateCreator().alter_excel_template(
                path_to_template_file_and_extension=path_to_template_file_and_extension,
                temporary_path=temporary_path, instantiated_attributes=instantiated_attributes, **kwargs)
        elif extension == '.csv':
            CsvTemplateCreator().determine_multiplicity_csv(
                path_to_template_file_and_extension=path_to_template_file_and_extension,
                path_to_subset=path_to_subset,
                temporary_path=temporary_path,
                **kwargs)

    def generate_basic_template(self, path_to_subset: Path, path_to_template_file_and_extension: Path,
                                temporary_path: Path, **kwargs):
        collector = self._load_collector_from_subset_path(path_to_subset=path_to_subset)
        otl_objects = []
        amount_of_examples = kwargs.get('amount_of_examples', 0)
        class_list = self.filters_assets_by_subset(path_to_subset=path_to_subset, **kwargs)

        for class_object in list(class_list):
            model_directory = None
            if kwargs is not None:
                model_directory = kwargs.get('model_directory', None)
            if amount_of_examples != 0:
                for i in range(amount_of_examples):
                    instance = dynamic_create_instance_from_uri(class_object.objectUri, model_directory=model_directory)
                    if instance is None:
                        continue
                    instance.fill_with_dummy_data()
                    otl_objects.append(instance)
            else:
                instance = dynamic_create_instance_from_uri(class_object.objectUri, model_directory=model_directory)
                if instance is None:
                    continue
                instance.fill_with_dummy_data()
                otl_objects.append(instance)

        converter = OtlmowConverter()
        converter.create_file_from_assets(filepath=temporary_path,
                                          list_of_objects=otl_objects, **kwargs)
        path_is_split = kwargs.get('split_per_type', True)
        extension = os.path.splitext(path_to_template_file_and_extension)[-1].lower()
        instantiated_attributes = []
        if path_is_split is False or extension == '.xlsx':
            instantiated_attributes = converter.create_assets_from_file(filepath=temporary_path,
                                                                        path_to_subset=path_to_subset)
        return instantiated_attributes

    @classmethod
    def filters_assets_by_subset(cls, path_to_subset: Path, **kwargs):
        list_of_otl_object_uri = kwargs.get('list_of_otl_objectUri', None)
        collector = cls._load_collector_from_subset_path(path_to_subset=path_to_subset)
        if list_of_otl_object_uri is None:
            return [x for x in collector.classes if x.abstract == 0]
        else:
            collector = cls._load_collector_from_subset_path(path_to_subset=path_to_subset)
            filtered_list = [x for x in collector.classes if x.objectUri in list_of_otl_object_uri]
            return filtered_list

    @staticmethod
    def _try_getting_settings_of_converter() -> Path:
        converter_path = Path(site.getsitepackages()[0]) / 'otlmow_converter'
        return converter_path / 'settings_otlmow_converter.json'

    @classmethod
    def return_temp_path(cls, path_to_template_file_and_extension: Path):
        tempdir = Path(tempfile.gettempdir()) / 'temp-otlmow'
        if not tempdir.exists():
            os.makedirs(tempdir)
        test = ntpath.basename(path_to_template_file_and_extension)
        temporary_path = Path(tempdir) / test
        return temporary_path


if __name__ == '__main__':
    subset_tool = SubsetTemplateCreator()
    subset_location = Path(ROOT_DIR) / 'UnitTests' / 'Subset' / 'Flitspaal_noAgent3.0.db'
    xls_location = Path(ROOT_DIR) / 'UnitTests' / 'Subset' / 'testFileStorage' / 'template_file.xlsx'
    subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                              path_to_template_file_and_extension=xls_location, add_attribute_info=True,
                                              highlight_deprecated_attributes=True,
                                              amount_of_examples=5,
                                              generate_choice_list=True,
                                              )
