import pandas as pd
from pathlib import Path
import seaborn as sns
import openml
from sklearn.datasets import fetch_openml
from xleda import wb


dimensions = []

def get_dimensions(name: str, input_df: pd.DataFrame):
    
    global dimensions

    dimensions.append({'name': name,
                       'rows': len(input_df),
                       'columns': len(input_df.columns)})


def collect_dimensions():

    # Dataset lists
    seaborn_datasets = ['penguins', 'titanic', 'anagrams', 'anscombe', 'attention', 'brain_networks', 'car_crashes', 'diamonds', 'dots', 'dowjones', 'exercise', 'flights', 'fmri', 'geyser', 'glue', 'healthexp', 'iris', 'mpg', 'planets', 'seaice', 'taxis', 'tips']
    aoml_examples = ['african_soil', 'air_bnb', 'mlb', 'nyc_taxi']


    # Path lists
    examples_path = Path() / 'examples'
    example_data_dir = examples_path / 'example_data'

  

    # -----------------------------------------------
    # Get dimentions for basic datasets

    for dataset in aoml_examples:
        get_dimensions(name=dataset, input_df=pd.read_feather(example_data_dir / f'{dataset}.feather'))

    
    # -----------------------------------------------
    # Get dimentions for seaborn datasets

    for dataset in seaborn_datasets:
        get_dimensions(name=dataset, input_df=sns.load_dataset(dataset))



    
    # -----------------------------------------------
    # Get dimentions for openml datasets
    
    # Grab metadata
    suite = openml.study.get_suite('OpenML-CC18')
    openml_datasets_df = openml.datasets.list_datasets(data_id=suite.data, output_format='dataframe')

    
    # Collate metadata
    openml_datasets_df['columns'] = openml_datasets_df['NumberOfFeatures'] + openml_datasets_df['NumberOfClasses']
    openml_datasets_df['rows'] = openml_datasets_df['NumberOfInstances']
    openml_datasets_df = openml_datasets_df[['name', 'rows', 'columns']]


    # Compile all results into a dataframe
    primary_datasets_df = pd.DataFrame.from_records(dimensions)

    return pd.concat([primary_datasets_df, openml_datasets_df], axis=0, ignore_index=True)

    


if __name__ == '__main__':
    
    df = collect_dimensions()

    wb = wb(name="Dataset Dimensions", 
            data=df, 
            overwrite=True)



    



