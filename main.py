from pathlib import Path

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator


if __name__ == '__main__':
    stc = SubsetTemplateCreator()
    p = Path('/home/davidlinux/Downloads/Subset VVOP-installatie.db')
    stc.generate_template_from_subset(path_to_subset=p, path_to_template_file_and_extension=Path('VVOP.xlsx'))
