from pathlib import Path

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator


if __name__ == '__main__':
    SubsetTemplateCreator()

    print (Path('/home/davidlinux/PycharmProjects/OTLMOW-ModelBuilder/UnitTests/../../OTLMOW-Template/UnitTests/TestClasses').resolve())