import pandas as pd
import numpy as np
from faker import Faker
import time

from ..src.xleda import create_workbook


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

    # Parameters
    mean = 0
    std_dev = 1

    # Generate Data
    initial_data = {
        
        # Fake categorical data
        'category': [fake.word() for _ in range(rows)],

        # Fake name data
        'fake_name': [fake.name() for _ in range(rows)]
        }

    # Create DataFrame
    df = pd.DataFrame(initial_data)

    # Add extra numeric columns
    for i in range(3, cols):
        df[fake.word()] = np.random.normal(mean, std_dev, rows)



    return df

# Create Sample DFs to test against

extra_large_row_df = create_fake_df(rows=1_000_000, cols=5)
large_row_df = create_fake_df(rows=100_001, cols=5)
extra_large_col_df = create_fake_df(rows=1_000_000, cols=5)
large_col_df = create_fake_df(rows=100_001, cols=101)
above_default_df = create_fake_df(rows=100_001, cols=101)
below_default_df = create_fake_df(rows=50000, cols=16)


timings = []


fake_dfs = [extra_large_row_df , large_row_df , extra_large_col_df , large_col_df , above_default_df , below_default_df]
fake_df_names = ['extra_large_row_df ', 'large_row_df ', 'extra_large_col_df ', 'large_col_df ', 'above_default_df ', 'below_default_df']


for i, fake_df in enumerate(fake_dfs):
    
    timing = {}
    start = time.time()

    timing['name'] = fake_df_names[i]

    timing['start'] = start

    create_workbook(input_df=fake_df,
                    close_wb=True, 
                    large_report=True, 
                    name=fake_df_names[i]
                    )
    
    timing['end'] = time.time() - start

    timings.append(timing)


timings_df = pd.DataFrame.from_records(timings)

print(timings_df)
