"""The builder that adds the relations between the objects of a template."""

from pathlib import Path
from typing import Optional

from otlmow_model.OtlmowModel.BaseClasses.OTLObject import (
    OTLObject,
    dynamic_create_type_from_uri,
)
from otlmow_model.OtlmowModel.Helpers.RelationCreator import (
    create_betrokkenerelation,
    create_relation,
)
from otlmow_modelbuilder.OSLOCollector import OSLOCollector
from otlmow_template.common import CommonMixin
from otlmow_template.constants import HEEFT_BETROKKENE_URI


class RelationsMixin(CommonMixin):
    """The class methods that add the relations between the objects of a template."""

    @classmethod
    def append_relations_to_objects(
        cls,
        otl_objects: list[OTLObject],
        collector: OSLOCollector,
        class_uris_filter: list[str],
        model_directory: Optional[Path],
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
        class_dict: dict[str, list[OTLObject]],
        otl_objects: list[OTLObject],
        class_uri: str,
        collector: OSLOCollector,
        model_directory: Optional[Path],
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
        cls,
        class_dict: dict[str, list[OTLObject]],
        otl_objects: list[OTLObject],
        relation,
        model_directory: Optional[Path],
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
        cls, relation, bron_instance, doel_instance, model_directory: Optional[Path]
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
