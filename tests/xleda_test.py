from pathlib import Path
import hashlib
import pandas as pd
import pickle
import pytest
import xlwings as xw
from matplotlib.figure import Figure


from xleda import wb
from tools.create_examples import (seaborn_datasets, penguins, 
                                   african_soil, titanic, titanic_completed, 
                                   titanic_incompleted, nyc_taxi)


# -----------------------------------------------
# Setup variables

# Path of examples
examples_path = Path().cwd() / 'examples'


# Example workbooks
penguin_dict = {'input_df': penguins, 
                'name': "Penguins",
                'wb_path': examples_path,
                'export': True,
                'add_plots': {'Pair Plots': Figure(),
                              'Null Matrix': Figure()},}

african_soil_dict = {'input_df': african_soil,
                     'name': "African Soil",
                     'wb_path':examples_path,
                     'large_report': True,
                     'export':True}

completed_titanic_dict = {'input_df': titanic,
                          'name': "Titanic",
                          'wb_path': examples_path / "Titanic Completed.xlsm",
                          'export': True}

nyc_xlsx_dict = {'input_df': nyc_taxi,
                 'name': "NYC Taxi",
                 'wb_path': examples_path,
                 'no_vba': True,
                 'export': True}




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



def test_examples_are_created():

    """
        Tests that all examples workbooks are created
    
    """


    primary_examples = ['African Soil', 'Airbnb', 'MLB', 'Penguins']


    complete_wb = {'Titanic': titanic_incompleted, 'Completed Titanic': titanic_completed}
    basic_wb =  {'NYC Taxi (xlsx)': examples_path / 'NYC Taxi.xlsx'}
    primary_wbs = {f'{dataset}': examples_path / (dataset + '.xlsm') for dataset in primary_examples}
    seaborn_wbs = {f'{dataset}': examples_path / 'other_examples' / (dataset.replace("_", " ").title() + '.xlsm') 
                   for dataset in seaborn_datasets}


    expected = primary_wbs | seaborn_wbs | complete_wb | basic_wb

    actual = {k: v for k, v in expected.items() if v.is_file()}

    # Check that all examples are created
    assert actual == expected





@pytest.mark.parametrize("wb_dict", 
                         [penguin_dict, 
                          african_soil_dict, 
                          completed_titanic_dict,
                          nyc_xlsx_dict])
def test_example_workbooks(wb_dict: dict):

    """ 
        Tests example workbooks for structure and content
    
    """



    # ---------------------------------------------
    # Set vars


    wb_object = wb(**wb_dict)
    blueprints = wb_object.blueprints
    wb_path = wb_object.cfg.path
       
    expected = {}
    actual = {}

    

    with xw.App(visible=True, add_book=False) as app:

        
        book = app.books.open(wb_path)


        for bp in blueprints:
            

            # ---------------------------------------------
            # Collect basic expectations from each blueprint

            expected[bp.name] = {'rows': bp.rows,
                                 'columns': bp.columns,
                                 'headers': bp.columns,
                                 'name': bp.title,
                                 'plots': []}
            
            

            # ---------------------------------------------
            # Collect basic actuals for each blueprint

            ws = book.sheets(bp.field_analysis)
            df = bp.source_data.iloc[:, 1:-3]
            table = ws.tables[0]
            
            # Ensure source data is there
            rows = table.data_body_range.rows.count # pyright: ignore[reportOptionalMemberAccess]
            columns = table.data_body_range.columns.count - 4 # pyright: ignore[reportOptionalMemberAccess]
            
            # Ensure named ranges are correctly reflected
            headers = len(ws.range("Headers"))

            # Ensure name is correct
            name = ws.range("Name").value


            actual[bp.name] = {'rows': rows,
                               'columns': columns,
                               'headers': headers,
                               'name': name,
                               'plots': []}
                     



            # ---------------------------------------------
            # Collect expected plots for each blueprint


            # Loop through columns and identify expected plots
            for col in df.columns:
                expected[bp.name]['plots'].append(f'composition_{col}')

                if pd.api.types.is_numeric_dtype(df[col]):
                    expected[bp.name]['plots'].append(f'histogram_{col}')
            
                

            # ---------------------------------------------
            # Collect actual plots


            # Unhide the expected plot ranges to verify size
            ws.range("Histogram").api.EntireRow.Hidden = False
            ws.range("Composition").api.EntireRow.Hidden = False

            
            # Collect actuals
            actual[bp.name]['plots'] = [shape.name for shape in ws.shapes
                                        if 'histogram_' in shape.name or 'composition_' in shape.name and
                                        shape.height / 72 > 1.5 and shape.width / 72 > 1.5]




        # ---------------------------------------------
        # Checks the additional plots

        if wb_object.cfg.additional_plots:

            expected['additional_plots'] = [plot['title'] for plot in wb_object.cfg.additional_plots]


            actual['additional_plots'] = [plot for plot in expected['additional_plots'] if 
                                          plot in book.sheet_names and 
                                          book.sheets(plot).shapes(plot) is not None]
            

        # ---------------------------------------------
        # Checks the pivot tables/sheets


        expected['pivot'] = [blueprints[0].pivot]


        actual['pivot'] = [pivot for pivot in expected['pivot'] if 
                           book.sheets(pivot).api.PivotTables('pvt_Pivot') is not None]



        # ---------------------------------------------
        # Checks that all expected worksheets are included

        # Start with additonal plots
        expected['sheets'] = [plot['title'] for plot in wb_object.cfg.additional_plots]

        
        # Add the rest
        for bp in blueprints:
            
            # Add all basic sheets
            expected['sheets'].append(bp.field_analysis)
            expected['sheets'].append(bp.overview)

            # Add the pivot if it's not an empty string            
            if bp.pivot:
                expected['sheets'].append(bp.pivot)


        actual['sheets'] = [sht for sht in expected['sheets'] if book.sheets(sht) is not None]

            
                
    assert actual == expected



def test_export_dict():
    
    """
        Tests the Titanic Completed example and the export functionality

    """
   
    # Use this to double-check export
    with open(r"tests/export_dict.pkl", "rb") as f:
        expected = pickle.load(f)

    actual_wb = wb(input_df=titanic,
                   wb_path=examples_path / 'Titanic Completed.xlsm', 
                   name='Titanic',
                   no_vba=True,
                   export=True)
    
    actual = actual_wb.export_dicts

    for i in range(len(actual)):
        for key, value in actual[i].items():
            if isinstance(value, pd.DataFrame):
                actual[i][key] = get_df_hash(input_df=actual[i][key])
                expected[i][key] =  get_df_hash(input_df=expected[i][key])

    
    assert actual == expected





