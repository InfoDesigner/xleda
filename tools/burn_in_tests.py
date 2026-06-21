from pathlib import Path
import pandas as pd
import numpy as np
from faker import Faker
import shutil
from xleda import wb
from global_utils import time_function

# ROOT_DIR = Path(__file__).resolve().parents[1]
# if str(ROOT_DIR) not in sys.path:
#     sys.path.insert(0, str(ROOT_DIR))




# ------------------------------------------------------
# Recreate output directory

output_path =  Path().cwd() / 'tests' / 'all_sizes'
output_path.mkdir(parents=True, exist_ok=True)

if output_path.exists() and output_path.is_dir():
    shutil.rmtree(output_path)

@time_function
def create_fake_df(rows: int, cols: int) -> pd.DataFrame:
    """Creates a df of fake mixed numeric/string data

    Args:
        rows (int): Number of rows
        cols (int): Number of cols

    Returns:
        pd.DataFrame: A pandas df
    """


    # Initialize Faker
    fake = Faker()

    categories = [fake.word() for _ in range(20)]
    names = [fake.name() for _ in range(20)]
    numbers = np.random.randint(0, 10, size=20)

    # Fake categorical data
    initial_data = {'category': np.random.choice(categories, size=rows),
                    'name': np.random.choice(names, size=rows),}
    
    # Fake numeric data
    for i in range(3, cols):
        initial_data[fake.word()] = np.random.choice(numbers, size=rows)

 
    # Return DataFrame
    return pd.DataFrame(initial_data) 





@time_function
def create_all_sizes():

    
    specs = {
        'extra_large_df': {'rows': 1_000_001, 'cols': 16001},
        'extra_large_row_df': {'rows': 1_000_001, 'cols': 5},
        'extra_large_col_df': {'rows': 50_000, 'cols': 16001},
        'large_df': {'rows': 100_001, 'cols': 101},
        'large_row_df': {'rows': 100_001, 'cols': 5},
        'large_col_df': {'rows': 50_000, 'cols': 101},
        'below_default_df': {'rows': 50_000, 'cols': 16},
    }

    for name, kwargs in specs.items():
        df = create_fake_df(**kwargs)


        # Create without large report flag
        wb(data=df,
           theme_color='random',
           file_name=name + '_small',
           wb_path=output_path)
        

        # Create with large report flag
        wb(data=df,
           theme_color='random',
           file_name=name + '_large',
           wb_path=output_path,
           large_report=True)
        




create_all_sizes()


