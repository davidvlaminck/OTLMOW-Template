from pathlib import Path

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator


if __name__ == '__main__':
    stc = SubsetTemplateCreator()
    p = Path('simpele_vergelijkings_subset2.db')
    stc.generate_template_from_subset(subset_path=p, template_file_path=Path('test.xlsx'), dummy_data_rows=3,
                                      ignore_relations=False)
