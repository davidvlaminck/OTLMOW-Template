from pathlib import Path

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator


if __name__ == '__main__':
    stc = SubsetTemplateCreator()
    p = Path('UnitTests/Subset/Kast_Agent.db')
    stc.generate_template_from_subset(subset_path=p, template_file_path=Path('test.csv'), dummy_data_rows=3,
                                      ignore_relations=False)
