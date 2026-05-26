from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import missingno as msno
import xlwings as xw

import openml
from sklearn.datasets import fetch_openml

from xleda import FieldAnalysis
from global_utils import time_function


# Dataset lists
seaborn_datasets = ['anagrams', 'anscombe', 'attention', 'brain_networks', 'car_crashes', 'diamonds', 'dots', 'dowjones', 'exercise', 'flights', 'fmri', 'geyser', 'glue', 'healthexp', 'iris', 'mpg', 'planets', 'seaice', 'taxis', 'tips']
aoml_datasets = ['african_soil', 'air_bnb', 'mlb', 'nyc_taxi']


# Path lists
examples_path = Path() / 'examples'
example_data_dir = examples_path / 'example_data'
other_examples_path = examples_path / 'other_examples'


# Data for primary examples
african_soil = pd.read_feather(example_data_dir / 'african_soil.feather').iloc[:, :600]
air_bnb = pd.read_feather(example_data_dir / 'air_bnb.feather')
mlb = pd.read_feather(example_data_dir / 'mlb.feather')
nyc_taxi = pd.read_feather(example_data_dir / 'nyc_taxi.feather')
penguins = sns.load_dataset("penguins")
titanic = sns.load_dataset("titanic")


# Data for Titanic Example
titanic_incompleted: Path = examples_path / 'Titanic.xlsx'
titanic_completed: Path = examples_path / 'Titanic (Completed).xlsx'




@time_function
def create_primary_examples():
    """
    Creates example xleda workbooks from aoml data sets

    """


    # Configure additional penguin plots
    pair_plots = sns.pairplot(penguins, hue="species").figure
    null_matrix = msno.matrix(penguins).get_figure() # type: ignore
    plt.style.use("dark_background")
    null_matrix.set_size_inches(9.35, 4.5) # type: ignore



    aoml_examples = [{'input_df': african_soil,
                      'name': 'African Soil',
                      'theme_color': "#0A7F02",
                      'large_report': True},
                     {'input_df': air_bnb,
                      'name': "Airbnb",
                      'theme_color': "#B30934",},
                     {'input_df': mlb,
                      'name': "MLB",
                      'theme_color': "#031835",},
                     {'input_df': nyc_taxi,
                      'name': "NYC Taxi",
                      'theme_color': "#8E6505"},
                     {'input_df': penguins, 
                      'name': "Penguins", 
                      'theme_color': "#4C4C4C",
                      'add_plots': {'Pair Plots': pair_plots,
                                    'Null Matrix': null_matrix,
                                    }},
                     {'input_df': titanic,
                      'name': "Titanic",
                      'no_vba': True}]
    

    for example in aoml_examples:

        xleda = FieldAnalysis(overwrite=True,
                              wb_path=examples_path,
                              **example)
        
        xleda.create_workbook()

        


@time_function
def download_openml_dataset(download_path: Path):
    """
    Downloads the current top openML datasets

    Parameters
    ----------
    download_path : Path
        Pathlib path for download directory
    """


    # Fetch the OpenML-CC18 benchmark suite
    suite = openml.study.get_suite('OpenML-CC18') 


    for task_id in suite.tasks: # type: ignore

        print("-" * 50 + '\n')
        
        # Get the task, dataset name, and set path for downloaded file
        task = openml.tasks.get_task(task_id)
        dataset_name = openml.datasets.get_dataset(task.dataset_id).name
        full_path = (example_data_dir / dataset_name).with_suffix(".feather")

        if not full_path.exists() and dataset_name:
        
            print(f"Downloading: {dataset_name}")
            
            df = fetch_openml(data_id=task.dataset_id, as_frame=True, parser='auto').frame
            
            # Store in a feather file with the dataset name
            df.to_feather(full_path)
            print(f"Successfully processed: {dataset_name}")


@time_function
def create_openml_examples():
    """
    Creates xleda workbooks from all feather files in a directory
    
    """

    
    for datafile in example_data_dir.iterdir():
        
        # Make sure the datafile is a feather file and the workbook 
        # is not already being created with the primary examples.
        name = datafile.stem

        if datafile.suffix == '.feather' and name not in aoml_datasets:


            df = pd.read_feather(datafile)

            xleda = FieldAnalysis(input_df=df,
                                  name=name.replace("-", " ").title(),
                                  theme_color='random',
                                  wb_path=other_examples_path)
            
            xleda.create_workbook()


@time_function
def complete_titanic_wb():
    """
    Creates a copy of the completed Titanic example from the current template

    """


    


    with xw.App(visible=False, add_book=False) as app:


        source_wb = app.books.open(titanic_completed)
        target_wb = app.books.open(titanic_incompleted)
        
        source_ws = source_wb.sheets("Field Analysis")
        target_ws = target_wb.sheets("Field Analysis")


        # Update data
        completed_df = source_ws.tables['tbl_SourceData'].range.options(pd.DataFrame, index=False).value
        target_ws.tables['tbl_SourceData'].update(completed_df, index=False)

        # Definitions/Notes/Description
        target_ws.range("Notes").value = source_ws.range("Notes").value
        target_ws.range("Definitions").value = source_ws.range("Definitions").value
        target_ws.range("Description").value = source_ws.range("Description").value
        target_ws.range("FieldLists").value = source_ws.range("FieldLists").value


        # Unhide completed sections
        visible_ranges = ['Data_Description', 'Compiled_Lists', 'Field_Notes', 'Field_Lists']

        for range in visible_ranges:
            target_ws.range(range).api.EntireRow.Hidden = False


        # Overwrite targetwb with source wb
        source_wb.close()
        target_wb.save(titanic_completed)


@time_function
def create_other_examples():
    """
    Creates other examples using the current template
    
    """

    for dataset in seaborn_datasets:
        
        df = sns.load_dataset(dataset)
        proper_title = dataset.replace("_", " ")


        # Configure xleda
        xleda = FieldAnalysis(input_df=df,
                              name=proper_title,
                              theme_color='random',
                              wb_path=other_examples_path,
                              overwrite=True)
        
        xleda.create_workbook()


if __name__ == '__main__':
    
    create_primary_examples()
    complete_titanic_wb()
    create_other_examples()
    # create_openml_examples()
