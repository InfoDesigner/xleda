from pathlib import Path
import pandas as pd
import missingno as msno
import xlwings as xw
import pickle
import time
from typing import Any
import platform

import openml
from sklearn.datasets import fetch_openml

import matplotlib
matplotlib.use('Agg')  # Set the silent, non-interactive backend
import matplotlib.pyplot as plt
import seaborn as sns



from xleda import wb

debug = True


os = platform.system()
win = os == 'Windows'
mac = os == 'Darwin'


# --------------------------------------------------
# Set variables

# Dataset lists
seaborn_datasets = ['anagrams', 'anscombe', 'attention', 'brain_networks', 'car_crashes', 'diamonds', 'dots', 'dowjones', 'exercise', 'flights', 'fmri', 'geyser', 'glue', 'healthexp', 'iris', 'mpg', 'planets', 'seaice', 'taxis', 'tips']
aoml_datasets = ['african_soil', 'air_bnb', 'mlb']


# Path lists
examples_path = Path() / 'examples'
example_data_dir = examples_path / 'example_data'
other_examples_path = examples_path / 'other_examples'
pickle_path = Path().cwd() / 'tests//export_dict.pkl'


# Data for primary examples
african_soil = pd.read_feather(example_data_dir / 'african_soil.feather').iloc[:, :600]
air_bnb = pd.read_feather(example_data_dir / 'air_bnb.feather')
mlb = pd.read_feather(example_data_dir / 'mlb.feather')
nyc_taxi = pd.read_feather(example_data_dir / 'nyc_taxi.feather')
titanic = sns.load_dataset("titanic")


# Add df example dfs
og_penguins = pd.read_csv("https://raw.githubusercontent.com/allisonhorst/palmerpenguins/refs/heads/main/inst/extdata/penguins_raw.csv")
penguins = sns.load_dataset('penguins')
seaice = sns.load_dataset('seaice')


# Data for Titanic Example
titanic_incompleted: Path = examples_path / 'Titanic.xlsm'
titanic_completed: Path = examples_path / 'Titanic Completed.xlsm'


section_dfs = []



def save_pickle(pickle_path: Path, cucumber: Any):

    """
    Saves a dictionary as a pickle file

    Parameters
    ----------
    pickle_path : Path
        File path to save the pickle

    cucumber : Any
        A pickleable Python object
    """


    # Save cucumber to a pickle file
    with open(pickle_path, 'wb') as file:
        pickle.dump(cucumber, file)



def create_primary_examples():
    
    """
        Creates example xleda workbooks from aoml data sets

    """

    global performance

    # Configure additional penguin plots
    pair_plots = sns.pairplot(penguins, hue="species").figure
    null_matrix = msno.matrix(penguins).get_figure() # type: ignore
    plt.style.use("dark_background")
    null_matrix.set_size_inches(9.35, 4.5) # type: ignore



    primary_examples = [
                     {'data': {"Airbnb": air_bnb},
                      'theme_color': "#B30934",},
                     {'data': {'African Soil': african_soil},
                      'theme_color': "#0A7F02",
                      'large_report': True},
                     {'data': {"MLB": mlb},
                      'theme_color': "#031835",},
                     {'data': {"NYC Taxi": nyc_taxi},
                      'theme_color': "#8E6505",
                      'no_vba': True},
                     {'data': {"Titanic": titanic},
                      'theme_color': 'random'},
                     {'data': {"Penguins": penguins,
                               'Sea Ice': seaice,
                               'OG Penguins': og_penguins},
                      'theme_color': 'random',
                      'plots': {'Pair Plots': pair_plots,
                                'Null Matrix': null_matrix,}}
                                ]
    

    for example in primary_examples[::-1]:

        xleda = wb(overwrite=True,
                   wb_path=examples_path,
                   open_wb=False,
                   debug=debug,
                   **example)
        
        compile_performance_data(xleda)
        
        time.sleep(2)
        



def create_other_examples():

    """
    Creates other examples using the current template
    
    """
    
    global performance

    for dataset in seaborn_datasets:
        
        df = sns.load_dataset(dataset)
        proper_title = dataset.replace("_", " ").title()


        # Configure xleda
        xleda = wb(data=df,
                   file_name=proper_title,
                   theme_color='random',
                   wb_path=other_examples_path,
                   debug=debug,
                   overwrite=True,
                   open_wb=False)
        
        compile_performance_data(xleda)
        
        time.sleep(2)



def complete_titanic_wb(update_pickle: bool=False):
    """
        
    Creates a copy of the completed Titanic example from the current template 
        optionally exports into a pickle file for testing

    """

    print("Recreating the Titanic Completed example")


    with xw.App(visible=debug, add_book=False) as app:


        
        source_wb = app.books.open(titanic_completed)
        target_wb = app.books.open(titanic_incompleted)
        
        source_ws = source_wb.sheets("Titanic")
        target_ws = target_wb.sheets("Titanic")


        # Update data
        completed_df = source_ws.tables[0].range.options(pd.DataFrame, index=False).value
        target_ws.tables[0].update(completed_df, index=False)

        # Definitions/Notes/Description
        target_ws.range("Notes").value = source_ws.range("Notes").value
        target_ws.range("Definitions").value = source_ws.range("Definitions").value
        target_ws.range("Description").value = source_ws.range("Description").value
                
        # Test me
        target_ws.range("Field_Lists").current_region.value = source_ws.range("Field_Lists").current_region.value


        # Unhide completed sections
        visible_ranges = ['Data_Description', 'Compiled_Lists', 'Field_Lists']

        for range in visible_ranges:
            if win:
                target_ws.range(range).api.EntireRow.Hidden = False
            elif mac:
                target_ws.range(range).api.entire_row.hidden.set(False)



        # Overwrite targetwb with source wb
        source_wb.close()
        
        target_wb.save(titanic_completed)

    
    # Update the pickle file
    if update_pickle:
        
        print("Updating pickle export")
        export_dict = wb(data=titanic,
                         wb_path=examples_path / 'Titanic Completed.xlsm',
                         file_name='Titanic',
                         export=True).export_dicts
        
        save_pickle(pickle_path=pickle_path, cucumber=export_dict)
    
    print("Done")

def download_openml_dataset():
    
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



def create_feather_examples():
    
    """
    Creates xleda workbooks from all feather files in a directory
    
    """

    
    for datafile in (example_data_dir / 'other').iterdir():
        
        # Make sure the datafile is a feather file and the workbook 
        # is not already being created with the primary examples.
        name = datafile.stem

        if datafile.suffix == '.feather':


            df = pd.read_feather(datafile)

            xleda = wb(data=df,
                       file_name=name.replace("-", " ").title(),
                       theme_color='random',
                       wb_path=other_examples_path,
                       overwrite=True,
                       debug=debug,
                       open_wb=False)
            
            compile_performance_data(xleda)


def compile_performance_data(wb: wb):


    global section_dfs


    # Add production timing to section performance
    section_df = wb.logger.section_performance.set_index('Production Section').T
    section_df = section_df.stack().to_frame().T # type: ignore
    
    new_columns = []
    
    
    # Combine mullti-index column names
    for level0, level1 in section_df.columns:
        if 'Seconds' in level0:
            new_columns.append(f"{level1} - Seconds")
        else:
            new_columns.append(f"{level1} - Pct")
    section_df.columns = new_columns

    # Combine section performance and df_overview
    df_overview = pd.concat([ds.df_overview for ds in wb.datasets], ignore_index=True)
    output_df = pd.concat([df_overview, section_df], axis=1)
    

    section_dfs.append(output_df)


            
def create_performance_wb():

    """
    Creates xleda workbooks documenting the production timing of xleda examples.
    
    """  


    section_df = pd.concat(section_dfs, ignore_index=True)



    wb(data=section_df,
       file_name="Performance Timing",
       overwrite=True)


if __name__ == '__main__':
    
    # Create examples
    # download_openml_dataset()
    # create_openml_examples()


    create_feather_examples()
    create_primary_examples()
    complete_titanic_wb(update_pickle=True)
    create_other_examples()
    create_performance_wb()
    

        
    
