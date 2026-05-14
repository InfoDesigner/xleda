# Imports, vars, config
import seaborn as sns
import pandas as pd
import shutil
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib as mpl
from matplotlib.figure import Figure 

import xlwings as xw
from xlwings import Range


# Field Analysis template
template_file = Path(__file__).parent / "xleda_template.xlsm"


# Set matplotlib theme
mpl.use('Agg') 
plt.style.use("dark_background")



def create_base_analysis(df: pd.DataFrame) -> pd.DataFrame:

    """ Produces the base analysis dataframe for an input dataframe"""

    # Order of output fields
    col_order = ['Data type', 'Memory Usage', 'Memory Usage %', 'Distinct', 'Distinct %', 'Count', 'Count %', 'Missing', 'Missing %', 'Mean', 'Median', 'Mode', 'Standard Deviation', 'Variance', 'Min', '5%', '25%', '50%', '75%', '95%', 'Max', 'Range', 'IQR']


    # Get statistical summary
    rows_count = len(df)
    desc = df.describe(include='all', percentiles=[.05, .25, .5, .75, .95])

    # Add additional components into a DataFrame
    info_df = pd.DataFrame({
        'Data type': df.dtypes.astype(str),
        'Memory Usage': df.memory_usage(deep=True, index=False),
        'Memory Usage %': df.memory_usage(deep=True, index=False)/df.memory_usage(deep=True).sum(),
        'Count': rows_count-df.isnull().sum(),
        'Count %': (rows_count - df.isnull().sum())/ rows_count,
        'Missing': df.isnull().sum(),
        'Missing %': df.isnull().sum() / rows_count,
        'IQR': desc.loc['75%'] - desc.loc['25%'],
        'Median': desc.loc['50%'],
        'Mode': df.mode().iloc[0],
        'Range': desc.loc['max'] - desc.loc['min'],
        'Variance': desc.loc['std']**2,
        'Distinct': df.nunique(),
        'Distinct %': df.nunique() / rows_count
        }).T

    # Combine info and describe dfs
    summary_df = pd.concat([info_df, desc])


    # Old:New field names
    field_map = {'mean': 'Mean',
                 'std': 'Standard Deviation',
                 'min': 'Min',
                 'max': 'Max'}

    # Rename index fields, reorder, and split into three filter
    summary_df = summary_df.rename(index=field_map)
    summary_df = summary_df.loc[col_order]

    return summary_df


def create_composition_table(input_df: pd.DataFrame, plot_color: str) -> Figure:

    # Font size
    font_size = 24

    # Prepare plot values
    counts = input_df.squeeze().value_counts() # type: ignore
    top_5 = counts.head(5)
    total_entries = len(input_df)
    other_count = total_entries - top_5.sum()

    # Assemble plot values
    categories = list(top_5.index) + ['Other']
    values = list(top_5.values) + [other_count]

    # Initialize the plot
    fig, ax = plt.subplots(figsize=(8, 8))


    y_pos = range(len(categories))[::-1]


    # Add bars to plot
    bars = ax.barh(y_pos, values, color=plot_color, height=0.5)  # noqa: F841

    # Remove spines
    for spine in ax.spines.values():
        spine.set_visible(False)

    # Remove other extra plot elements
    ax.set_xticks([])
    ax.set_yticks([])
    ax.set_title('')
    
    
    # Make enough room for text on the left so they don't overlap
    plt.subplots_adjust(left=0.4, right=0.9) 

    max_val = max(values)

    for y, cat, val in zip(y_pos, categories, values):
        pct = (val / total_entries) * 100     
        
        # Truncate long category names
        display_cat = str(cat)
        if len(display_cat) > 6:
            display_cat = display_cat[:5] + '..'
        
        # Add labels and adjust left to prevent overlap
        ax.text(-0.55, y, display_cat, color='white', va='center', ha='left', fontsize=font_size, transform=ax.get_yaxis_transform())

        # Add Percentages
        ax.text(-0.05, y, f'{pct:.0f}%', color='white', va='center', ha='right', fontsize=font_size, transform=ax.get_yaxis_transform())


        # Add Counts to the right of the bars
        ax.text(val + max_val * 0.02, y, str(val), color='white', va='center', ha='left', fontsize=font_size)

    return fig


def create_histogram(input_df: pd.DataFrame, plot_color: str) -> Figure:


    # Setup plot area
    fig, ax = plt.subplots(figsize=(5, 5))
    ax.set_axis_off() 

    # Plot a histogram

    sns.histplot(data=input_df, x=input_df.columns[0], color=plot_color, stat="density", alpha=0.5, ax=ax)

    # Layer the KDE line
    sns.kdeplot(data=input_df, x=input_df.columns[0], color="silver", linewidth=3, ax=ax)

    # Add vertical mean line
    mean_val = input_df[input_df.columns[0]].mean()
    ax.axvline(mean_val, color="silver", linestyle=":", linewidth=2)

    # Remove tick labels
    ax.tick_params(left=False, bottom=False, labelleft=False, labelbottom=False)

    # Add Min and Max text at the bottom corners
    min_val = input_df[input_df.columns[0]].min()
    max_val = input_df[input_df.columns[0]].max()
    
    ax.text(0, -0.05, f"Min {min_val:g}", transform=ax.transAxes, fontsize=16, color="silver", ha="left", va="top")
    ax.text(1, -0.05, f"Max {max_val:g}", transform=ax.transAxes, fontsize=16, color="silver", ha="right", va="top")


    return fig


def add_small_multiple(fig: Figure, target_range: Range):
    """
    Adds a small chart to an Excel cell.  
    The small chart is 90% of the size of the cell and is centered.

    Args:
        fig (Figure): A matplotlib Figure object
        target_range (Range): An Excel cell
    """


    
    # Calculate 90% of cell dimensions
    target_width = target_range.width * 0.9
    target_height = target_range.height * 0.9
    
    # Calculate position to center the picture
    target_left = target_range.left + (target_range.width / 2) - (target_width / 2)
    target_top = target_range.top + (target_range.height / 2) - (target_height / 2)
    
    # Add the picture to the sheet
    pic = target_range.sheet.pictures.add(
        fig,
        left=target_left,
        top=target_top,
        width=target_width,
        height=target_height
    )
    
    # Set placement to xlMoveAndSize
    try:
        pic.api.Placement = 1
    except AttributeError:
        pass


def add_plots(input_df: pd.DataFrame, input_wb: xw.Book, plot_color: str):
    """Adds plots for each column of a pandas dataframe to an Excel range

    Args:
        input_df (pd.DataFrame): A data frame to plot
        input_wb (xw.Workbook): Target Workbook for plots
        plot_color (str): Primary color used for the plots 
    """
    
    # --------------------------------------------------
    # Setup workbook objects
    
    wb = input_wb
    field_analysis_ws = wb.sheets("Field Analysis")


    # Set initial ranges for added plots
    histogram_range = field_analysis_ws.range("Histogram")
    composition_range = field_analysis_ws.range("CompositionTable")


    for col in input_df.columns:


        # --------------------------------------------------
        # Add Composition Multiple

        composition_table = create_composition_table(input_df[[col]], plot_color)
        add_small_multiple(target_range=composition_range, fig=composition_table)


        if pd.api.types.is_numeric_dtype(input_df[col]):
            

            # --------------------------------------------------
            # Add Histogram Multiple

            histogram = create_histogram(input_df[[col]], plot_color)
            add_small_multiple(target_range=histogram_range, fig=histogram)


        # --------------------------------------------------
        # Increment Target Ranges

        histogram_range = histogram_range.offset(0, 1)
        composition_range = composition_range.offset(0, 1)


def create_workbook(input_df: pd.DataFrame, 
                    name: str, 
                    theme_color: str = '#053476', 
                    large_report=False, 
                    overwrite: bool = False,
                    close_wb: bool = False) -> xw.Book | None:
    
    """Creates a Field Analysis workbook from a given dataframe

    Args:
        input_df (pd.DataFrame): Pandas dataframe to base the Field Analysis on
        name (str): Title of the workbook
        theme_color (str, optional): Theme color for the main worksheeet and the plots.  Defaults to '#053476'.
        large_report (bool, optional): Raises default limts of 100 columns/100,000 rows to 16,000 columns/1,000,000 rows. Defaults to False
        overwrite (bool, optional): Whether to overwrite existing workbook with the same name. Defaults to False.
        close_wb: bool = Whether to close the created workbook once created.  Defaults to False
        
        ) -> xw.Book | None:

    Returns:
        xw.Book: An xlwings workbook object

    """


    # --------------------------------------------------
    # Setup Source Data


    # Create a copy and get dimensions
    source_df = input_df.copy()
    rows = len(input_df)
    columns = len(input_df.columns)


    # Evaluate dimensions and subsample if necessary
    above_default = rows > 100_000 or columns > 100
    above_limit = rows > 16_000 or columns > 1_000_000
    warning_msg = ""

    if not above_default:
        warning = False

    elif above_default and not large_report:
        warning = True
        warning_msg = "This is only showing a sample because it is larger than the limits of 100,000 rows/100 columns"
        rows = min(rows, 100_000)
        columns = min(columns, 100)
        source_df = source_df.iloc[:, :columns].sample(n=rows).sort_index()

    elif large_report and not above_limit:
        warning = False

    elif large_report and above_limit:
        warning = True
        warning_msg = "This is only showing a sample because it is larger than the limits of 1,000,000 rows/16,000 columns"
        rows = min(rows, 1_000_000)
        columns = min(columns, 16_000)
        source_df = source_df.iloc[:, :columns].sample(n=rows).sort_index()
    

    
    # --------------------------------------------------
    # Create Field Analysis and split into sections

    summary_df = create_base_analysis(source_df)
    overview_df = summary_df.loc[['Data type', 'Distinct %', 'Missing %', 'Memory Usage %']]
    composition_df = summary_df.loc[['Memory Usage', 'Distinct', 'Count', 'Missing']]
    summary_stats_df = summary_df.loc[['Mean', 'Median', 'Mode', 'Standard Deviation', 'Variance']]
    percentiles_df = summary_df.loc[['Min', '5%', '25%', '50%', '75%', '95%', 'Max', 'Range', 'IQR']]

    

    # --------------------------------------------------
    # Create Field Analysis Workbook

    # Set Template Path
    field_analysis_path = (Path().cwd() / name).with_suffix('.xlsm')

    
    # Handle existing files as appropriate
    if field_analysis_path.is_file() and not overwrite:
        print(f"""There is already a workbook named {field_analysis_path} 
              Use overwrite=True or rename/remove the existing workbook""")
        return
    elif field_analysis_path.is_file() and overwrite:
        try:
            field_analysis_path.unlink(missing_ok=True)
        except PermissionError:
            print("Error: The workbook cannot be overwrittten while open.")
            return

    # Create a copy of the template and open it
    shutil.copy(template_file, field_analysis_path)
    

    # --------------------------------------------------
    # Initialize Template

    app = xw.App(visible=not close_wb, add_book=False)

    wb = app.books.open(field_analysis_path, read_only=False)
    field_analysis_ws = wb.sheets("Field Analysis")
    field_analysis_ws.range("FieldAnalysisTheme").color = theme_color
    field_analysis_ws.range("Dimensions").value = f"rows = {rows}, columns = {columns}"
    
                    

    # --------------------------------------------------
    # Format Field Analysis sections

    # Format placeholders
    format_from = field_analysis_ws.range('FormatRange')
    format_to = field_analysis_ws.range('FormatRange').offset(0,1).resize(None, columns-2)
    format_from.api.Copy()
    format_to.api.Select()
    field_analysis_ws.api.Paste()
    
    
    # Clear clipboard, set selection to top, and add header values
    wb.app.api.CutCopyMode = False
    field_analysis_ws.range('Headers_Start').api.Select()
    headers = source_df.columns.to_list() + ["Record Hash"]
    field_analysis_ws.range('Headers_Start').value = headers






    # --------------------------------------------------
    # Add Field Analysis sections

    field_analysis_ws.range('Overview')[0, 0].value = overview_df.values
    field_analysis_ws.range('Composition')[0, 0].value = composition_df.values
    field_analysis_ws.range('Summary_Stats')[0, 0].value = summary_stats_df.values
    field_analysis_ws.range('Percentiles')[0, 0].value = percentiles_df.values


    # --------------------------------------------------
    # Add Plots

    add_plots(input_df=source_df, input_wb=wb, plot_color=theme_color)



    # --------------------------------------------------
    # Setup Source Data table


    # Add fields to Source Data to track individual records
    source_df.insert(loc=0, column='Mark For Removal', value='False')
    source_df['Record Hash'] = pd.util.hash_pandas_object(source_df, index=False)

    # Update tbl_SourceData with source data, format "Mark For Removal" column
    source_table = field_analysis_ws.tables['tbl_SourceData']
    source_table.update(source_df, index=False)
    source_table.data_body_range[0, 0].copy(destination=source_table.data_body_range[:, 0]) # type: ignore

    # Set "Record Hash" to LightHeader for contrast with source data fields
    last_header = field_analysis_ws.tables('tbl_SourceData').header_row_range.last_cell # type: ignore
    field_analysis_ws.range('LightHeader').copy(destination=last_header)
    last_header.value = "Record Hash"
    last_header.columns.autofit()


    # --------------------------------------------------
    # Adjust Named Ranges for Field Action Formulas 
        
    field_analysis_ws.range('tbl_SourceData[Record Hash]').name = 'RecordHashes'
    field_analysis_ws.range('FieldRange').resize(row_size=1, column_size=len(headers)-1).name = 'FieldRange'

    for i in range(1, 7):
        excel_range = 'FieldAction' + str(i)
        field_analysis_ws.range(excel_range).resize(row_size=1, column_size=len(headers)-1).name = excel_range


    
    # --------------------------------------------------
    # Initialize Field Analysis UI

    # Show/Hide Data Size Warning
    if warning:
        field_analysis_ws.range("Warning").value = warning_msg
        field_analysis_ws.range("Warning").api.EntireRow.Hidden = not warning

    # Collapse subsections
    for excel_range in ['Composition', 'Summary_Stats', 'Percentiles', 'Field_Actions', 'Field_Action_Lists']:
        field_analysis_ws.range(excel_range).api.EntireRow.Hidden = True

    
    
    # --------------------------------------------------
    # Save workbook, close if required, and return it

    wb.save(field_analysis_path)



    if close_wb:
        wb.close()
        app.quit()


    print(f"{field_analysis_path} created")
          
    return wb


def export_analysis(input_wb: xw.Book) -> dict[str, list[str]]:

    """Exports notes and action lists from an xleda field analysis workbook

    Args:
        input_wb (xw.Book): _description_

    Returns:
        dict[str, list[str]]: _description_
    """

    target_ws = input_wb.sheets('Field Analysis')

    export_dict = {}    
    
    actions = target_ws.range('Actions').value
    action_lists = target_ws.range('ActionLists').value

    for i in range(len(actions)):
        export_dict[i] = action_lists[i]


    return export_dict