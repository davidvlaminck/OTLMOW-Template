import time
from pathlib import Path

from otlmow_template.InfotemplateGenerator import InfotemplateGenerator


if __name__ == '__main__':
    start_time = time.time()
    generator = InfotemplateGenerator()
    
    # Path to your subset database
    subset_path = Path('1.db')
    output_path = Path('infotemplate_output.xlsx')
    
    if not subset_path.exists():
        print(f"Subset file not found: {subset_path}")
        print("Please update the subset_path variable to point to your subset database.")
    else:
        generator.generate_infotemplate_from_subset(
            subset_path=subset_path,
            template_file_path=output_path
        )
    
    end_time = time.time()
    print(f"Execution time: {end_time - start_time} seconds")