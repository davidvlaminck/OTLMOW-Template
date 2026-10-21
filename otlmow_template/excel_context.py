"""The state of a sheet while its headers are processed."""

from dataclasses import dataclass, field

from openpyxl.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet
from otlmow_model.OtlmowModel.BaseClasses.OTLObject import OTLObject
from otlmow_template.attribute_info import AttributeInfo


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
