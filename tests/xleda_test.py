# Testing imports
from pathlib import Path
import platform
import pandas as pd
import xlwings as xw
import pickle
import shutil
import hashlib
from typer.testing import CliRunner
import pytest

# Plotting imports
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib as mpl
import missingno as msno

# Package imports
from xleda.utilities import DataSetParser, Settings, DataError, Logger
from xleda.cli import cli
from xleda import wb







# -----------------------------------------------------------------
# Runtime Vars

os = platform.system()
win = os == 'Windows'
mac = os == 'Darwin'
debug = False
runner = CliRunner()
update_pickle = False



# -----------------------------------------------------------------
# Path lists

examples_path = Path.cwd() / 'examples'
other_examples_path = examples_path / 'other_examples'
data_dir = examples_path / 'data'
tmp_path = str((Path.cwd() / 'examples' / 'tmp').resolve())
feather_data_path = data_dir / 'feathers'



# -----------------------------------------------------------------
# Local Data

# Dataframes
african_soil = pd.read_feather(data_dir / 'african_soil.feather').iloc[:, :600]
nyc_taxi = pd.read_feather(data_dir / 'nyc_taxi.feather')
air_bnb = pd.read_feather(data_dir / 'air_bnb.feather')
og_penguins = pd.read_csv(data_dir / "penguins_raw.csv")

# File Paths
duck_db = str((data_dir / 'duckdb.duckdb').resolve())
sqlite = str((data_dir / 'chinook.db').resolve())
air_bnb = str((data_dir / 'air_bnb.feather').resolve())
csv = str((data_dir / 'titanic.csv').resolve())
excel = str((data_dir / 'sample_workbook.xlsx').resolve())
parquet = str((data_dir / 'userdata.parquet').resolve())
json = str((data_dir / 'sample.json').resolve())
pickle_source = str((data_dir / 'multiple_dfs.pkl').resolve())

# -----------------------------------------------------------------
# Online Data

titanic = sns.load_dataset("titanic")
penguins = sns.load_dataset('penguins')
seaice = sns.load_dataset('seaice')



# -----------------------------------------------------------------
# Workbook Paths

titanic_incompleted: Path = examples_path / 'Titanic.xlsm'
titanic_completed: Path = examples_path / 'Titanic Completed.xlsm'
pickle_path = data_dir / 'export_dict.pkl'




# -----------------------------------------------------------------
# Additional penguin plots

mpl.use('Agg')  # Set the silent, non-interactive backend
pair_plots = sns.pairplot(penguins, hue="species").figure
null_matrix = msno.matrix(penguins).get_figure() # type: ignore
plt.style.use("dark_background")
null_matrix.set_size_inches(9.35, 4.5) # type: ignore




# --------------------------------------------------------------------
# Invalid argument sets that produce erorrs

invalids = {
    'invalid_json': {'data': 'invalid.json', 
                     'error': 'XML/JSON file not successfully parsed'},
    'missing_file': {'data': 'not_a_real_file.csv', 
                     'error': 'Data file not found'},
    'unsupported_extension': {'data': 'unsupported_extension.xleda', 
                              'error': 'Unsupported file type'},
    'nothing': {'data': None,
                'error': 'No Data Provided'},
    'spoiled_pickle': {'data': 'spoiled_pickle.pkl',
                       'error': 'Pickle file not successfully parsed'},
    'fake_db': {'data': 'fake_db.db',
                'error': 'Database file is not sqlite or duckdb'}}

# --------------------------------------------------------------------
# Valid Python API arguments


valids = {
        'penguins': {'data': {"Penguins": penguins,
                              "OG Penguins": og_penguins},
                     'wb_path': examples_path,
                     'plots': {'Pair Plots': pair_plots,
                               'Null Matrix': null_matrix},
                     'expected_path': examples_path / "Penguins.xlsm"},
        
        'african_soil': {'data': {"African Soil": african_soil},
                         'wb_path':examples_path,
                         'large_report': True,
                         'theme': '#31AC83',
                         'expected_path': examples_path / "African Soil.xlsm"},
                         
        'nyc_xlsx': {'data': {"NYC Taxi": nyc_taxi},
                     'wb_path': examples_path / "NYC Taxi.xlsx",
                     'expected_path': examples_path / "NYC Taxi.xlsx"},
                     
        'diamonds': {'data': {df_name: sns.load_dataset(df_name.lower()) for df_name in ['Diamonds', 'dots', 'dowjones']},
                     'wb_path': examples_path,
                     'theme': 'random',
                     'expected_path': examples_path / "Diamonds.xlsm"}}




# -----------------------------------------------
# Setup variables


cli_args = {
    
    # These are deleted after creation
    'duck_db': ['wb', duck_db],
    'csv': ['wb', csv],
    'excel_file': ['wb', excel],
    'parquet': ['wb', parquet],
    'json': ['wb', json],
    'pickle': ['wb', pickle_source],
                
                
    # These are kept
    'sqlite': ['wb', sqlite, '--theme', "#7124BA", '--wb_path', str(examples_path)], 
    'feather': ['wb', air_bnb, '--file_name', 'Airbnb', '--theme', '#B30934', '--wb_path', str(examples_path)], 
    
    
    # These don't create workbooks
    'toggle_vba': ['cli_settings', 'vba'], 
    'retoggle_vba': ['cli_settings', 'vba'], 
    'theme': ['cli_settings', 'theme', '#262626']
    
    }








def get_df_hash(input_df: pd.DataFrame) -> str:
    
    """
    Converts a dataframe to a sha256 hash in order to test equality

    Parameters
    ----------
    input_df : pd.DataFrame
        A pandas dataframe of any size.

    Returns
    -------
    str
        A SHA 256 Hash

    """

    # Sort to ensure structural consistency
    df_sorted = input_df.sort_index(axis=1).sort_index(axis=0)

    # Convert to CSV and then to bytes for hashing
    return hashlib.sha256(df_sorted.to_csv(index=False).encode()).hexdigest()











@pytest.fixture
def create_completed_example():
    """
        
    Creates a copy of the completed Titanic example from the current template 
        optionally exports into a pickle file for testing

    """

    print("Recreating the Titanic Completed example")


    # Create titanic workbook as input_df
    wb(input_df=sns.load_dataset("titanic"),
       
       # Name isn't being derived from this wb_path
       wb_path=titanic_incompleted,
       open_wb=False,
       no_vba=False,
       theme='random',
       overwrite=True)



    with xw.App(visible=debug, add_book=False) as app:


        # Open both workbooks
        source_wb = app.books.open(titanic_completed)
        target_wb = app.books.open(titanic_incompleted)
        
        
        
        # --------------------------------------------------
        # Copy data from the completed copy to the new copy
        
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
        

        # Save export dict to a pickle file
        with open(pickle_path, 'wb') as file:
            pickle.dump(export_dict, file)

        

# Includes the fixture above as a requirement
def test_completed_example(create_completed_example):
    
    """
    Tests the Titanic Completed example and the export functionality

    """
   
    # Use this to double-check export
    with open(pickle_path, "rb") as f:
        expected = pickle.load(f)

    actual = wb(data=titanic,
                wb_path=examples_path / 'Titanic Completed.xlsm', 
                file_name='Titanic',
                export=True).export_dicts

    for i in range(len(actual)):
        for key, value in actual[i].items():
            if isinstance(value, pd.DataFrame):
                actual[i][key] = get_df_hash(input_df=actual[i][key])
                expected[i][key] =  get_df_hash(input_df=expected[i][key])

    
    assert actual == expected
    
    




# -------------------------------------------------------------------------
# Test Python API

@pytest.mark.parametrize("args", 
                         list(invalids.values()), 
                         ids=list(invalids.keys()))
def test_invalids(args: dict):
    
    """
    Tests the Python API succeeds or fails how it should

    """
    
    data = args.pop('data', None)
    error = args.pop('error', '')
    expected_path = args.pop('expected_path', None)
    

    # Convert data to a path for data files
    if data is not None and isinstance(data, str):
        data = data_dir / data
    
            
    # if an expected path is included, successfully create a workbook
    if expected_path:
        try:
            
            # Delete the old one if it exists
            Path(expected_path).unlink(missing_ok=True)
            
            # Create the valid workbook example
            wb(data=data, overwrite=True, open_wb=False, no_vba=False, **args)
            
            # Ensure it's there
            assert Path(expected_path).is_file()
            
        except Exception as e:
            raise Exception(e)
    
    
    
    # If there's an error argument provided, don't create a workbook but do validate the error
    if error:
        with pytest.raises(DataError, match=error):
            wb(data=data, **args)
    






# -------------------------------------------------------------------------
# Tests that the CLI can create workbooks from several common file types


@pytest.mark.parametrize("command", 
                         list(cli_args.values()), 
                         ids=list(cli_args.keys()))
def test_from_cli(command: list):
    
    """
    Creates an xleda workbook from the CLI endpoint

    """
    
    # Set vars
    temp_folder = Path(tmp_path)
    
    
    # If it's a cli setting, remove that flag before running
    if 'cli_settings' in command:
        command = command[1:]
    
    # Add common options if it's a 'xleda wb' command
    else:
    
        # Add common options
        command += ['--vba', '--no_open_wb', '--overwrite']
    
    
        # Add path when needed
        if 'wb_path' not in str(command):
            command += ['--wb_path', tmp_path]
        
    
        # Show the workbook if debugging
        if debug:
            command += ['--debug']
        
        # Recreate the tmp directory on start
        shutil.rmtree(temp_folder, ignore_errors=True)
        temp_folder.mkdir(parents=True, exist_ok=True)
            


    # Run the cli command and assert that there was no error
    result = runner.invoke(cli, command)
    assert result.exit_code == 0, f"CLI failed with output: {result.output}"

    
    # Remove the tmp directory on end
    shutil.rmtree(temp_folder, ignore_errors=True)




@pytest.mark.parametrize("args", list(valids.values()), ids=list(valids.keys()))
def test_valids(args: dict):

    """ 
    Tests example workbooks for structure and content
    
    """



    # ---------------------------------------------
    # Set vars 

    
    # Extract the expected path
    wb_path = args.pop('expected_path')
    
    
    # Create the workbook
    wb_object = wb(**args, overwrite=True, no_vba=False, open_wb=debug)
    
    
    # Calculate the Datasets externally
    logger = Logger()
    external_datasets = DataSetParser(settings=Settings(locals=args, logger=logger)).datasets
    
    # Use the internal datasets to capture worksheet/table names
    datasets = wb_object.datasets
    
    
    # Set comparison vars
    expected = {}
    actual = {}
    
    # Add paths
    expected['wb_path'] = str(wb_path).lower()
    actual['wb_path'] = str(wb_object.template.path).lower()

    # Test the created workbook for structure
    with xw.App(visible=True, add_book=False) as app:

        
        book = app.books.open(wb_path)
        for i, ds in enumerate(datasets):
            

            # ---------------------------------------------
            # Collect basic expectations from each blueprint

            expected[ds.name] = {'rows': ds.rows,
                                 'columns': ds.columns,
                                 'headers': ds.columns,
                                 'name': ds.name,
                                 'plots': []}
            

            # ---------------------------------------------
            # Collect basic actuals for each blueprint

            ws = book.sheets(ds.name)
            df = external_datasets[i].source_data.iloc[:, 1:-2]
            table = ws.tables(ds.table_name)
            
            # Ensure source data is there
            rows = table.data_body_range.rows.count # pyright: ignore[reportOptionalMemberAccess]
            columns = table.data_body_range.columns.count - 3 # pyright: ignore[reportOptionalMemberAccess]
            
            # Ensure named ranges are correctly reflected
            headers = len(ws.range("Headers"))

            # Ensure name is correct
            name = ws.range("Name").value


            actual[ds.name] = {'rows': rows,
                               'columns': columns,
                               'headers': headers,
                               'name': name,
                               'plots': []}



            # ---------------------------------------------
            # Collect expected plots for each blueprint


            # Loop through columns and identify expected plots
            for col in df.columns:
                expected[ds.name]['plots'].append(f'composition_{col}')

                if pd.api.types.is_numeric_dtype(df[col]):
                    expected[ds.name]['plots'].append(f'histogram_{col}')
            
                

            # ---------------------------------------------
            # Collect actual plots



            if win:
                ws.range("Histogram").api.EntireRow.Hidden = False
                ws.range("Composition").api.EntireRow.Hidden = False
            elif mac:
                ws.range("Histogram").api.entire_row.hidden.set(False)
                ws.range("Composition").api.entire_row.hidden.set(False)


            
            # Collect actuals
            actual[ds.name]['plots'] = [shape.name for shape in ws.shapes
                                        if 'histogram_' in shape.name or 'composition_' in shape.name and
                                        shape.height / 72 > 1.5 and shape.width / 72 > 1.5]




        # ---------------------------------------------
        # Checks the additional plots
        
        plots = wb_object.settings.plots

        if plots:

            expected['plots'] = [plot_name for plot_name, figure in plots.items()]


            actual['plots'] = [plot for plot in expected['plots'] if 
                               plot in book.sheet_names and 
                               book.sheets(plot).shapes(plot) is not None]
            



        # ---------------------------------------------
        # Checks that all expected worksheets are included

        # Start with additonal plots and Overview
        expected['sheets'] = [plot_title for plot_title, figure in plots.items()] + ['Overview']

        
        # Add the rest
        for ds in datasets:
            
            # Add all basic sheets
            expected['sheets'].append(ds.name)


        actual['sheets'] = [sht for sht in expected['sheets'] if book.sheets(sht) is not None]

            
                
    assert actual == expected