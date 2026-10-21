"""The description of the attributes of a template, collected per column.

The Excel and the csv writer both add a row with the definition of every
attribute and a row that marks the deprecated attributes, so the collection of
those rows is shared.
"""

from dataclasses import dataclass, field
from typing import Optional, cast

from otlmow_converter.DotnotationHelper import DotnotationHelper
from otlmow_model.OtlmowModel.BaseClasses.OTLObject import (
    OTLAttribuut,
    OTLObject,
    get_attribute_by_name,
)
from otlmow_template.constants import TYPE_URI_DESCRIPTION


class AttributeInfoMixin:
    """The class methods that describe the attributes of a template."""

    @classmethod
    def get_attribute_info(cls, instance: OTLObject, header: str, attribute: OTLAttribuut) -> str:
        """Returns the definition to show in the attribute info row for a column of a template.

        The definition of the 'waarde' attribute of a quantitative value is always the same and does not explain
        the attribute itself. For quantitative values the definition of the attribute itself is combined with the
        standard unit of the datatype.
        """
        quantitative_attribute = cls.get_quantitative_attribute(instance=instance, header=header)
        if quantitative_attribute is None:
            return cast(str, attribute.definition)
        standard_unit = cls.get_standard_unit(quantitative_attribute)
        if standard_unit is None:
            return cast(str, attribute.definition)
        definition = quantitative_attribute.definition.strip()
        if not definition.endswith('.'):
            definition = f'{definition}.'
        return f'{definition} Standaard eenheid: {standard_unit}'

    @classmethod
    def get_quantitative_attribute(cls, instance: OTLObject, header: str) -> Optional[OTLAttribuut]:
        """Returns the attribute a column header refers to when it points to the value of a
        quantitative value, None when the header does not point to a quantitative value."""
        attribute = DotnotationHelper.get_attribute_by_dotnotation(
            instance_or_attribute=instance, dotnotation=header, waarde_shortcut=False
        )
        if not attribute.field.waarde_shortcut_applicable:
            return None
        if not hasattr(attribute.field.waardeObject, 'standaardEenheid'):
            return None
        return attribute

    @classmethod
    def get_standard_unit(cls, attribute: OTLAttribuut) -> Optional[str]:
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
    def find_standard_unit(cls, standaard_eenheid) -> Optional[str]:
        """Returns the standard unit in the notations of a datatype,
        None when the notations do not contain one."""
        for notation in [standaard_eenheid.usagenote, standaard_eenheid.constraints]:
            if '"' in notation:
                return cast(str, notation.split('"')[1])
        return None


def get_attribute_info(instance: OTLObject, header: str, attribute: OTLAttribuut) -> str:
    """Returns the definition to show in the attribute info row for a column of a template."""
    return AttributeInfoMixin.get_attribute_info(
        instance=instance, header=header, attribute=attribute
    )


def get_quantitative_attribute(instance: OTLObject, header: str) -> Optional[OTLAttribuut]:
    """Returns the attribute a column header refers to when it points to the value of a
    quantitative value, None when the header does not point to a quantitative value."""
    return AttributeInfoMixin.get_quantitative_attribute(instance=instance, header=header)


def get_standard_unit(attribute: OTLAttribuut) -> Optional[str]:
    """Returns the standard unit (ex. 'mo') of a quantitative value attribute,
    None when the standard unit of the datatype is unknown."""
    return AttributeInfoMixin.get_standard_unit(attribute=attribute)


def find_standard_unit(standaard_eenheid) -> Optional[str]:
    """Returns the standard unit in the notations of a datatype,
    None when the notations do not contain one."""
    return AttributeInfoMixin.find_standard_unit(standaard_eenheid=standaard_eenheid)


@dataclass
class AttributeInfo:
    """The information about the attributes of a template, collected per column."""

    add_attribute_info: bool
    add_deprecated: bool
    attribute_info_row: list[str] = field(default_factory=list)
    deprecated_row: list[str] = field(default_factory=list)

    def add_attribute(self, instance, header: str, attribute) -> None:
        """Add the definition and the deprecated marker of an attribute."""
        if self.add_attribute_info:
            self.attribute_info_row.append(
                get_attribute_info(instance=instance, header=header, attribute=attribute)
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
