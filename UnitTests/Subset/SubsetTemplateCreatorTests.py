import os
import unittest
from pathlib import Path
from unittest import TestCase

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))


class SubsetTemplateCreatorTests(TestCase):
    def test_func1(self):
        subset_tool = SubsetTemplateCreator()
        subset_location = Path(ROOT_DIR) / 'OTL_AllCasesTestClass.db'
        xls_location = Path(ROOT_DIR) / 'template_file_text.xlsx'
        csv_location = Path(ROOT_DIR) / 'template_file_text.csv'

        with self.subTest('xlsx'):
            subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                      path_to_template_file_and_extension=xls_location,
                                                      class_directory='UnitTests.TestClasses.Classes')

        with self.subTest('csv, split per type'):
            subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                      path_to_template_file_and_extension=csv_location,
                                                      class_directory='UnitTests.TestClasses.Classes',
                                                      split_per_type=True)
        with self.subTest('csv, not split per type'):
            subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                      path_to_template_file_and_extension=csv_location,
                                                      class_directory='UnitTests.TestClasses.Classes',
                                                      split_per_type=False)

        self.assertTrue(True)

    def test_subset_actual_subset(self):
        subset_tool = SubsetTemplateCreator()
        csv_location = Path(ROOT_DIR) / 'camera_steun.csv'
        subset_tool.generate_template_from_subset(path_to_subset=Path(ROOT_DIR) / 'camera_steun.db',
                                                  path_to_template_file_and_extension=csv_location,
                                                  split_per_type=True)
        self.assertTrue(True)

    @unittest.skip
    def test_func2(self):
        list_of_otl_objects = []
        SubsetTemplateCreator.filters_assets_by_subset(Path('OTL_AllCasesTestClass.db'), list_of_otl_objects)
