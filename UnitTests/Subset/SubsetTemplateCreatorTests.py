import os
import shutil
import unittest
from pathlib import Path
from unittest import TestCase

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

# ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

ROOT_DIR = Path(__file__).parent


class SubsetTemplateCreatorTests(TestCase):

    # Toegangsprocedure en Agent zijn niet meer in de subset aanwezig
    def test_func1(self):
        subset_tool = SubsetTemplateCreator()
        subset_location = ROOT_DIR / 'Flitspaal_noAgent3.0.db'
        xls_location = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.xlsx'
        csv_location = Path(ROOT_DIR) / 'testFileStorage' / 'template_file_text.csv'

        with self.subTest('xlsx'):
            subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                      path_to_template_file_and_extension=xls_location,
                                                      )
            assert Path(ROOT_DIR / 'testFileStorage' / 'template_file_text.xlsx').exists()

        with self.subTest('csv, split per type'):
            subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                      path_to_template_file_and_extension=csv_location,
                                                      split_per_type=True)

        with self.subTest('csv, not split per type'):
            subset_tool.generate_template_from_subset(path_to_subset=subset_location,
                                                      path_to_template_file_and_extension=csv_location,
                                                      split_per_type=False)
            assert Path(ROOT_DIR / 'testFileStorage' / 'template_file_text.csv').exists()

        self.assertTrue(True)
        shutil.rmtree(Path(ROOT_DIR) / 'testFileStorage')
        os.makedirs(Path(ROOT_DIR) / 'testFileStorage')


    def test_subset_actual_subset(self):
        subset_tool = SubsetTemplateCreator()
        csv_location = ROOT_DIR / 'testFileStorage' / 'camera_steun.csv'
        subset_tool.generate_template_from_subset(path_to_subset=ROOT_DIR / 'camera_steun.db',
                                                  path_to_template_file_and_extension=csv_location,
                                                  split_per_type=True)
        csv1 = ROOT_DIR / 'testFileStorage' / 'camera_steun_onderdeel_Bevestiging.csv'
        csv2 = ROOT_DIR / 'testFileStorage' / 'camera_steun_onderdeel_Camera.csv'
        csv3 = ROOT_DIR / 'testFileStorage' / 'camera_steun_onderdeel_RechteSteun.csv'
        assert csv1.exists()
        assert csv2.exists()
        assert csv3.exists()
        shutil.rmtree(Path(ROOT_DIR) / 'testFileStorage')
        os.makedirs(Path(ROOT_DIR) / 'testFileStorage')

    @unittest.skip
    def test_func2(self):
        list_of_otl_objects = []
        SubsetTemplateCreator.filters_assets_by_subset(Path('OTL_AllCasesTestClass.db'), list_of_otl_objects)
