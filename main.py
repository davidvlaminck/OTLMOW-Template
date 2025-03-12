from pathlib import Path

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator


if __name__ == '__main__':
    stc = SubsetTemplateCreator()
    p = Path('simpele_vergelijkings_subset2.db')
    stc.generate_template_from_subset(path_to_subset=p, path_to_template_file_and_extension=Path('test.xlsx'),
                                      add_geo_artefact=True, amount_of_examples=1, add_attribute_info=True)
