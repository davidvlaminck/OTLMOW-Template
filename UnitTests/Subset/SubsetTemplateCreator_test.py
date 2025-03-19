import os
import tempfile
from pathlib import Path

import openpyxl
from _pytest.fixtures import fixture
from openpyxl.workbook import Workbook

from otlmow_template.ExcelTemplateCreator import ExcelTemplateCreator
from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
model_directory_path = Path(__file__).parent.parent / 'TestModel'


def test_subset_actual_subset():
    subset_tool = SubsetTemplateCreator()
    csv_location = Path(ROOT_DIR) / 'camera_steun.csv'
    subset_tool.generate_template_from_subset(subset_path=Path(ROOT_DIR) / 'camera_steun_2.14.db',
                                              template_file_path=csv_location)
    csv1 = Path(ROOT_DIR) / 'camera_steun_onderdeel_Bevestiging.csv'
    csv2 = Path(ROOT_DIR) / 'camera_steun_onderdeel_Camera.csv'
    csv3 = Path(ROOT_DIR) / 'camera_steun_onderdeel_RechteSteun.csv'
    assert not csv1.exists()
    assert csv2.exists()
    assert csv3.exists()

    subset_tool.generate_template_from_subset(subset_path=Path(ROOT_DIR) / 'camera_steun_2.14.db',
                                              template_file_path=csv_location, ignore_relations=False)
    assert csv1.exists()
    assert csv2.exists()
    assert csv3.exists()

    csv1.unlink()
    csv2.unlink()
    csv3.unlink()
