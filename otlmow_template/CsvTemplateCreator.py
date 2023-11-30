import csv
import ntpath
import os
import tempfile
from pathlib import Path
from typing import List, Sequence, Optional

from otlmow_converter.DotnotationHelper import DotnotationHelper
from otlmow_converter.OtlmowConverter import OtlmowConverter


class CsvTemplateCreator:

    @classmethod
    def determine_multiplicity_csv(cls, path_to_template_file_and_extension: Path, path_to_subset: Path,
                                   temporary_path: Path, **kwargs):
        path_is_split = kwargs.get('split_per_type', True)
        if path_is_split is False:
            cls.alter_csv_template(path_to_template_file_and_extension=path_to_template_file_and_extension,
                                   temporary_path=temporary_path, path_to_subset=path_to_subset, **kwargs)
        else:
            cls.multiple_csv_template(path_to_template_file_and_extension=path_to_template_file_and_extension,
                                      path_to_subset=path_to_subset, **kwargs)
        file_location = os.path.dirname(temporary_path)
        [f.unlink() for f in Path(file_location).glob("*") if f.is_file()]

    @classmethod
    def multiple_csv_template(cls, path_to_template_file_and_extension, path_to_subset, **kwargs):
        file_location = os.path.dirname(path_to_template_file_and_extension)
        tempdir = Path(tempfile.gettempdir()) / 'temp-otlmow'
        file_name = ntpath.basename(path_to_template_file_and_extension)
        split_file_name = file_name.split('.')
        things_in_there = os.listdir(tempdir)
        csv_templates = [x for x in things_in_there if x.startswith(split_file_name[0] + '_')]
        for file in csv_templates:
            test_template_loc = Path(os.path.dirname(path_to_template_file_and_extension)) / file
            temp_loc = Path(tempdir) / file
            cls.alter_csv_template(path_to_template_file_and_extension=test_template_loc, temporary_path=temp_loc,
                                   path_to_subset=path_to_subset, **kwargs)

    @classmethod
    def alter_csv_template(cls, path_to_template_file_and_extension, path_to_subset, temporary_path,
                           **kwargs):
        converter = OtlmowConverter()
        instantiated_attributes = converter.create_assets_from_file(filepath=temporary_path,
                                                                    path_to_subset=path_to_subset)
        header = []
        data = []
        delimiter = ';'
        add_geo_artefact = kwargs.get('add_geo_artefact', False)
        add_attribute_info = kwargs.get('add_attribute_info', False)
        highlight_deprecated_attributes = kwargs.get('highlight_deprecated_attributes', False)
        amount_of_examples = kwargs.get('amount_of_examples', 0)
        quote_char = '"'
        with open(temporary_path, 'r+', encoding='utf-8') as csvfile:
            new_file = open(path_to_template_file_and_extension, 'w', encoding='utf-8')
            reader = csv.reader(csvfile, delimiter=delimiter, quotechar=quote_char)
            for row_nr, row in enumerate(reader):
                if row_nr == 0:
                    header = row
                else:
                    data.append(row)
            if add_geo_artefact is False:
                [header, data] = cls.remove_geo_artefact_csv(header=header, data=data)
            if add_attribute_info:
                info = cls.add_attribute_info_csv(header=header, data=data,
                                                  instantiated_objects=instantiated_attributes)
                new_file.write(delimiter.join(info) + '\n')
            data = cls.remove_mock_data_csv(data=data, rows_of_examples=amount_of_examples)
            if highlight_deprecated_attributes:
                header = cls.highlight_deprecated_attributes_csv(header=header, data=data,
                                                                 instantiated_attributes=instantiated_attributes)
            new_file.write(delimiter.join(header) + '\n')
            for d in data:
                new_file.write(delimiter.join(d) + '\n')
            new_file.close()

    @classmethod
    def add_attribute_info_csv(cls, header: List[str], data: List[List[str]], instantiated_objects: List) -> List[str]:
        info_data = []
        info_data.extend(header)

        dotnotation_module = DotnotationHelper()

        uri_index = cls.get_type_uri_index_in_row(header)
        found_uris = set(d[uri_index] for d in data)

        for uri in found_uris:
            single_object = next(x for x in instantiated_objects if x.typeURI == uri)
            for index, dotnototation_title in enumerate(info_data):
                if dotnototation_title == 'typeURI':
                    info_data[index] = 'De URI van het object volgens https://www.w3.org/2001/XMLSchema#anyURI .'
                else:
                    dotnotation_attribute = dotnotation_module.get_attribute_by_dotnotation(
                        single_object, dotnototation_title)
                    info_data[index] = dotnotation_attribute.definition
        return info_data

    @classmethod
    def remove_mock_data_csv(cls, data, rows_of_examples):
        if rows_of_examples == 0:
            data = []
        return data

    @classmethod
    def highlight_deprecated_attributes_csv(cls, header, data, instantiated_attributes):
        found_uri = []
        dotnotation_module = DotnotationHelper()
        uri_index = cls.get_type_uri_index_in_row(header)
        for d in data:
            if d[uri_index] not in found_uri:
                found_uri.append(d[uri_index])
        for uri in found_uri:
            single_object = next(x for x in instantiated_attributes if x.typeURI == uri)
            for dotnototation_title in header:
                if dotnototation_title == 'typeURI':
                    continue
                else:
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
                        header[index] = "[DEPRECATED] " + value
        return header

    @classmethod
    def get_type_uri_index_in_row(cls, header: Sequence[str]) -> Optional[int]:
        try:
            return header.index('typeURI')
        except ValueError:
            return None

    @classmethod
    def remove_geo_artefact_csv(cls, header, data):
        if 'geometry' in header:
            deletion_index = header.index('geometry')
            header.remove('geometry')
            for d in data:
                d.pop(deletion_index)
        return [header, data]

