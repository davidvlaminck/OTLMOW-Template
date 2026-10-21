"""The helpers that more than one part of the template generation needs."""

from collections import defaultdict

from otlmow_model.OtlmowModel.BaseClasses.OTLObject import OTLObject


class CommonMixin:
    """The class methods that group the objects of a template."""

    @classmethod
    def fill_class_dict(cls, instances: list[OTLObject]) -> dict[str, list[OTLObject]]:
        """Groups the objects of a template by their type uri."""
        class_dict = defaultdict(list)
        for instance in instances:
            class_dict[instance.typeURI].append(instance)
        return class_dict
