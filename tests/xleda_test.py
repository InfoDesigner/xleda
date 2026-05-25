from pathlib import Path
from typing import Any

import hashlib
import pandas as pd
import pickle
import pytest
import xlwings as xw

from xleda import FieldAnalysis
from tools.create_examples import seaborn_datasets, penguins, african_soil, titanic, titanic_completed


examples_path = Path().cwd() / 'examples'



def get_df_hash(input_df: pd.DataFrame) -> str:
    
    """
    Converts a dataframe to a sha256 hash

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
    Tests that all examples are created
    
    """


    primary_examples = ['African Soil', 'Airbnb', 'MLB', 'Penguins', 'NYC Taxi', 'Titanic']

    primary_wbs = {'file': (examples_path / (dataset + '.xlsm')) for dataset in primary_examples}
    seaborn_wbs = {'file': (examples_path / 'other_examples' / (dataset + '.xlsm')) for dataset in seaborn_datasets}

    expected = primary_wbs | seaborn_wbs

    actual = {k: v for k, v in expected.items() if v.is_file()}

    # Check that all examples are created
    assert actual == expected



@pytest.mark.parametrize("wb_path, input_df, extra_plots", 
                         [(examples_path / "Penguins.xlsm", penguins, ['Pair Plots', 'Null Matrix']), 
                          (examples_path / "African Soil.xlsm", african_soil, None),
                          (titanic_completed, titanic, None)])
def test_example_workbook(wb_path: Path, input_df: pd.DataFrame, extra_plots: list[str] | None):

    """ 
    Tests example workbooks for structure and content
    
    """

    df = input_df


    expected: dict[str, Any] = {'rows': len(df),
                                'columns': len(df.columns)}
    actual: dict[str, Any] = {}


    with xw.App(visible=False, add_book=False) as app:

        # ---------------------------------------------
        # Set vars

        wb = app.books.open(wb_path)
        ws = wb.sheets("Field Analysis")


        # Verify that all columns are represented in the headers
        expected['Header Count'] = expected['columns']
        actual['Header Count'] = len(ws.range("Headers"))
        


        # ---------------------------------------------
        # Verify that the source data table has the correct amount of rows, columns

        table = ws.tables['tbl_SourceData']

        actual['rows'] = table.data_body_range.rows.count # pyright: ignore[reportOptionalMemberAccess]
        actual['columns'] = table.header_row_range.columns.count - 2 # pyright: ignore[reportOptionalMemberAccess]


        # ---------------------------------------------
        # Verify that field plots are correct 

        # Unhide the expected plot ranges 
        ws.range("Histogram").api.EntireRow.Hidden = False
        ws.range("Composition").api.EntireRow.Hidden = False

        
        # Set initial plot vars
        expected['plots'] = []
        actual['plots'] = [shape.name for shape in ws.shapes 
                           if 'histogram_' in shape.name or 'composition_' in shape.name]
        

        # Loop through columns and identify expected plots
        for col in df.columns:
            expected['plots'].append(f'composition_{col}')

            if pd.api.types.is_numeric_dtype(df[col]):
                expected['plots'].append(f'histogram_{col}')
            

        # Identify correctly sized plots
        expected['correctly_sized_plots'] = [plot for plot in expected['plots']]

        actual['correctly_sized_plots'] = [plot for plot in actual['plots'] 
                                           if (ws.shapes(plot).height / 72) > 1.5 
                                           and (ws.shapes(plot).width / 72) > 1.5]
        


        # ---------------------------------------------
        # Checks the additional plots

        if extra_plots:

            expected['extra_plots'] = extra_plots


            actual['extra_plots'] = [plot for plot in extra_plots if  # type: ignore
                                    plot in wb.sheet_names and 
                                    wb.sheets(plot).shapes(plot) is not None]
            

        # ---------------------------------------------
        # Checks the pivot tables/sheets


        expected['pivots'] = ['Blanks', 'Pivot']


        actual['pivots'] = [pivot for pivot in expected if 
                            wb.sheets(pivot).api.PivotTables('pvt_' + pivot) is not None]


        
    assert actual == expected



def test_export_dict():
    
    """
    Tests the Titanic (Completed) and the export_analysis functionality

    """
    

    with open(r"tests/export_dict.pkl", "rb") as f:
        expected = pickle.load(f)

    actual: dict[str, Any] = FieldAnalysis(name = 'Titanic (Completed)', wb_path=examples_path, input_df=titanic).export_analysis()


    for df in ['altered_source_data', 'source_data']:

        actual[df] =  get_df_hash(input_df=actual[df])
        expected[df] =  get_df_hash(input_df=expected[df])
    
    assert actual == expected





