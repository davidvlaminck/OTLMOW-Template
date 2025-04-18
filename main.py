import time
from pathlib import Path

from otlmow_template.SubsetTemplateCreator import SubsetTemplateCreator


if __name__ == '__main__':
    start_time = time.time()
    stc = SubsetTemplateCreator()
    p = Path('OTL (3).db')
    stc.generate_template_from_subset(subset_path=p, template_file_path=Path('test.xlsx'), dummy_data_rows=1,
                                      ignore_relations=True)
    end_time = time.time()
    print(f"Execution time: {end_time - start_time} seconds")
