"""The writer that turns the objects of a template into a csv file.

It adds the row with the definition of every attribute and the row that marks
the deprecated attributes.
"""

import csv
from asyncio import sleep
from pathlib import Path

from otlmow_converter.DotnotationHelper import DotnotationHelper
from otlmow_model.OtlmowModel.BaseClasses.OTLObject import OTLObject
from otlmow_model.OtlmowModel.Helpers.GenericHelper import get_ns_and_name_from_uri
from otlmow_template.attribute_info import AttributeInfo
from otlmow_template.common import CommonMixin
from otlmow_template.constants import QUOTE_CHAR


class CsvWriterMixin(CommonMixin):
    """The class methods that write the csv template."""

    @classmethod
    def alter_csv_template(
        cls,
        instances: list[OTLObject],
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
        instances: list[OTLObject],
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
        instances: list[OTLObject],
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
    def read_csv_file(cls, file_path: Path) -> tuple[list[str], list[list[str]]]:
        """Reads the header row and the data rows of a csv file."""
        with Path(file_path).open(encoding='utf-8') as file:
            csv_reader = csv.reader(file, delimiter=';', quotechar=QUOTE_CHAR)
            header_row = next(csv_reader)
            csv_data = list(csv_reader)
        return header_row, csv_data

    @classmethod
    def collect_csv_header_info(cls, instance, info: AttributeInfo, header_row: list[str]) -> None:
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
        header_row: list[str],
        csv_data: list[list[str]],
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
    def write_csv_data_rows(
        cls, csv_writer, csv_data: list[list[str]], dummy_data_rows: int
    ) -> None:
        """Writes the data rows of a csv file, unless the template has no data rows."""
        if dummy_data_rows == 0:
            return
        for line in csv_data:
            csv_writer.writerow(line)
