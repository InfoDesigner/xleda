import seaborn as sns  # type: ignore
import pandas as pd
import shutil
from pathlib import Path
import ast
import sys
import random
import platform
import time
import send2trash
from typing import Any

import matplotlib.pyplot as plt
import matplotlib as mpl 
from matplotlib.figure import Figure

import xlwings as xw
from xlwings.constants import HAlign, VAlign

from .color_shaping import use_black_text, ensure_readable, color_formatter, warn_print

from tqdm.auto import tqdm


template_file = Path(__file__).parent / "xleda_template.xlsm"

separator = "\n" + ("-" * 100)

# Set matplotlib theme
mpl.use("Agg")  # type: ignore
plt.style.use("dark_background")  # type: ignore


default_row_limit = 25_000
default_column_limit = 50
upper_row_limit = 1_000_000
upper_column_limit = 16_000




class wb():

    """
        A class that represents an xleda workbook.

    """

    def __init__(self, 
                 input_df: pd.DataFrame, 
                 name: str = 'xleda', 
                 theme_color: str = "#262626", 
                 large_report: bool = False, 
                 overwrite: bool = False, 
                 wb_path: Path= Path().cwd(),
                 add_plots: dict[str, Figure] = {},
                 no_vba: bool = False,
                 open_wb: bool = True,
                 export: bool = False):

        """
            Creates an xleda workbook

        Parameters
        ----------

        input_df : pd.DataFrame
            * A pandas dataframe of any size.  
            * Will create an xleda workbook that is 25,000 rows/50 columns by default.  
            * Use large_report=True to expand this to Excel's limits.

        name : str
            * Name of the workbook to be created.

        theme_color : str, optional
            * A hexidecimal color used for charts/accent color.  
            * Use theme_color='random' for random colors
            * Defaults to "#262626"

        large_report : bool, optional
            * Used to override default limits of 25,000 rows/50 columns. 
            * Sets limits to 1,000,000 rows/16,000 columns
            * Defaults to False

        overwrite : bool, optional
            * Whether to overwrite existing reports with the same name
            * Defaults to False

        wb_path : Path, optional
            * Pathlib path directory of an xleda workbook
            * Defaults to current working directory

        add_plots : dict[str, Figure], optional
            * Additional plots to be included 
            * Uses "{'plot1_name': Plot1Figure, 'plot2_name': Plot2Figure, ...}" format.  
            * Each entry will get it's own worksheet.  
            * No resizing or syling is done for plots added this way
            * Defaults to None

        no_vba : bool, optional
            * Will create the workbook as an xlsx file that has no VBA.
            
        open_wb: bool, optional
            * Whether to open the workbook after creating.  
            * Set to False if creating multiple workbooks.
            * Defaults to True

        export: bool, optional
            * Exports data from an xleda workbook instead of creating one.
            * Exported data is available as a dictonary and as class properties. 
                  e.g. my_export_dict = wb(df).export_dict
                       my_export_dict['source_data'] == wb(df).source_data # returns True
                                   
            * Defaults to False
            * xleda.wb() includes:

                * `description`: Dataframe description if you've added one
                * `definitions`: Any field definitions you've added.
                * `notes`: Any field notes you've added
                * `lists`: Any lists showing in the compiled lists section
                * 'field_metadata': A basic metadata dataframe, combining information from 
                                    pandas info/describe/quantile.
                * 'overview_metadata': A transposed copy of the field_metadata.
                * `source_data`: A copy of your unaltered source data that includes 
                                `Record Hash`/`Record List`/`HasBlank` columns.
                * `altered_source_data`: source data from the workbook that includes 
                                         any manual edits you've made such as removing 
                                         records, renaming fields, etc. 
                                         
                                         ** Note that data types will likely change in the round-trip translation. **

        """

        # Set base properties
        self.input_df: pd.DataFrame = input_df.copy()
        self.name: str = name
        self.additional_plots: dict[str, Figure] | dict = add_plots
        self.large_report: bool = large_report
        self.overwrite: bool = overwrite
        self.no_vba: bool = no_vba
        self.export: bool = export
        self.open_wb: bool = open_wb
        self.os = platform.system()
        self.entry_message: str = separator
        self.exit_msg: str = ""
        self.start_time = time.time()
        
        
        
        # Set Template Path
        if no_vba:
            self.wb_path: Path = (wb_path / self.name).with_suffix(".xlsx")
        else:
            self.wb_path: Path = (wb_path / self.name).with_suffix(".xlsm")
        

        # Configure Theme
        if theme_color == 'random':
            self.theme_color = "#{:06x}".format(random.randint(0, 0xFFFFFF))
        else:
            self.theme_color: str = theme_color
        self.black_text: bool = use_black_text(self.theme_color)
        self.tqdm_theme: str = ensure_readable(self.theme_color[:7])


        # Add placeholder properties
        self.warning_msg: str = ""
        self.warning: bool = False
        self.rows = len(input_df)
        self.columns = len(input_df.columns)
        
        
        # Configure Source Data
        self.source_data: pd.DataFrame = self._configure_source_data()


        # add field_metadata/overview_metadata
        self.field_metadata: pd.DataFrame = self._create_field_metadata()
        self.df_metadata: dict[str, Any] = self._create_df_metadata()
        self.overview_metadata = self.field_metadata.T


        # Add initial export_dict
        self.export_dict: dict[str, Any] = {'source_data': self.source_data,
                                            'field_metadata': self.field_metadata,
                                            'overview_metadata': self.field_metadata.T}
        

        # Initialize performance logging
        self.performance_section: dict[str, Any] = {}
        self.performance_pivot: dict[str, Any] = {}
        self.performance_plots: dict[str, Any] = {}
        self.performance: dict[str, pd.DataFrame] = {}
        self._log_performance(log_type='section')


        # Create workbook or add export data from workbook to export_dict
        if self.export:
            self._export_analysis()
        else:
            self._create_wb()



    def _configure_source_data(self) -> pd.DataFrame:
        
        """
            Configures source data for analysis and subsamples it if necessary


        Returns:
            source_df: A pandas dataframe object
        
        """

        # --------------------------------------------------
        # Configure variables and create a copy of index as the last column 

        rows = self.rows
        columns = self.columns
        df = self.input_df
        

        above_default = rows > default_row_limit or columns > default_column_limit
        above_limit = rows > upper_row_limit or columns > upper_column_limit
          

        # Source data is above default limits
        if above_default and not self.large_report and not above_limit:

            self.warning = True
            self.warning_msg = "This is only showing a sample because it is larger than the default limits of 25,000 rows/50 columns.  See documentation for details."
            rows = min(rows, default_row_limit)
            columns = min(columns, default_column_limit)
            

        # Source data is larger than Excel's limits
        elif above_limit:

            self.warning = True
            self.warning_msg = "This is only showing a sample because it is larger than Excel's limits of 1,000,000 rows/16,000 columns.  See documentation for details."
            rows = min(rows, upper_row_limit)
            columns = min(columns, upper_column_limit)


        # Subsample df if needed, record adjusted rows/columns
        df = (df.iloc[:, :columns].sample(n=rows).sort_index())

        self.rows = rows
        self.columns = columns

        # Add index, Record List, HasBlank, Record Hash fields to source_data
        df['index'] = df.index
        df.insert(loc=0, column="Record List", value="False")
        df['HasBlank'] = df.isnull().any(axis=1).astype(int)
        df["Record Hash"] = pd.util.hash_pandas_object(df, index=False)


        return df



    def _create_field_metadata(self) -> pd.DataFrame:

        """
            Produces the base analysis dataframe for an input dataframe

        Returns
        -------
        pd.DataFrame
            A dataframe with field metadata 
        
        """


        # --------------------------------------------------
        # Collect metadata

        # Omit added columns
        df = self.source_data.iloc[:, 1: -3]

        # Order of output fields
        row_order = ["Data type", "Memory Usage", "Memory Usage %", "Distinct", "Distinct %", "Count", "Count %", "Missing", "Missing %", "Mean", "Median", "Mode", "Standard Deviation", "Variance", "Min", "5%", "25%", "50%", "75%", "95%", "Max", "Range", "IQR", ]

        # Get statistical summary
        rows_count = len(df)
        desc = df.describe(include="all", percentiles=[0.05, 0.25, 0.5, 0.75, 0.95])


        # Add missing describe entries if they don't exist
        all_describe_fields = ['count', 'mean', 'std', 'min', '25%', '50%', '75%', 'max', 'unique', 'top', 'freq', '5%', '95%']
        desc = desc.reindex(all_describe_fields)

        # Add additional components into a DataFrame
        info_df = pd.DataFrame(
            {
                "Data type": df.dtypes.astype(str),
                "Memory Usage": df.memory_usage(deep=True, index=False),
                "Memory Usage %": df.memory_usage(deep=True, index=False) / df.memory_usage(deep=True).sum(),  # type: ignore
                "Count": rows_count - df.isnull().sum(),
                "Count %": (rows_count - df.isnull().sum()) / rows_count,
                "Missing": df.isnull().sum(),
                "Missing %": df.isnull().sum() / rows_count,
                "IQR": desc.loc["75%"] - desc.loc["25%"],
                "Median": desc.loc["50%"],
                "Mode": df.mode().iloc[0],
                "Range": desc.loc["max"] - desc.loc["min"],
                "Variance": desc.loc["std"] ** 2,
                "Distinct": df.nunique(),
                "Distinct %": df.nunique() / rows_count,
                }).T


        # --------------------------------------------------
        # Combine/Format Metadata

        
        # Combine info and describe dfs
        summary_df = pd.concat([info_df, desc])

        # Old:New field names
        field_map = {
            "mean": "Mean",
            "std": "Standard Deviation",
            "min": "Min",
            "max": "Max",
            }

        # Rename index fields, reorder, filter rows
        summary_df = summary_df.rename(index=field_map)
        summary_df = summary_df.loc[row_order]

        # convert to string to prevent issues with timedelta/random datatypes
        summary_df = summary_df.astype(str)

        return summary_df



    def _create_df_metadata(self) -> dict[str, Any]:

        """
            Creates a dictonary of df-level metadata

        """

        df = self.input_df

        df_metadata = {'Rows': self.rows,
                       'Columns': self.columns,
                       'Memory Usage': f"{df.memory_usage(deep=True).sum():,} bytes",
                       'Distinct Rows %': (len(df.drop_duplicates()) / len(df)),
                       'Missing %': df.isnull().mean().mean(),}
        
        return df_metadata
        


    def _create_blank_template(self, progress_bar: tqdm):

        """
            Creates a blank Field Analysis template, overwriting if necessary

        """


        # Return an error if the OS isn't MacOS or Windows

        if self.os not in ['Windows', 'Darwin']:
            warn_print("Error: xleda requires Windows or MacOS")
            sys.exit()

        
        junk_drawer = 'Recycle Bin' if self.os == 'Windows' else 'Trash'
        
        progress_bar.update(1)


        # Return an error if there's an existing file and no overwrite flag

        if self.wb_path.is_file() and not self.overwrite:

            warn_print(f"Error: There is already a workbook named {self.wb_path}!")
            warn_print("Use overwrite=True or rename/remove the existing workbook")
            
            sys.exit()


        # Delete the file if there's an overwrite flag, return error if it's open

        elif self.wb_path.is_file() and self.overwrite:
            try:

                send2trash.send2trash(self.wb_path)

                self.exit_msg += f"\nThe previously existing file was sent to your {junk_drawer}"
                
            except OSError:
                
                warn_print("Error: The workbook cannot be overwritten while open!")
                sys.exit()

            except Exception:
                
                warn_print(f"[bold red]An unexpected error occurred when deleting {self.wb_path.name}")
                sys.exit()

        progress_bar.update(1)
        
        # Create parent directories if necessary and a blank copy of the template
        self.wb_path.parent.mkdir(parents=True, exist_ok=True)
        progress_bar.update(1)

        # if no_vba has been set, open the default template, remove macro triggers from shapes, and save as xlsx
        if self.no_vba:

            # Hide warnings and convert file
            with xw.App(visible=False, add_book=False) as app:
                
                app.api.DisplayAlerts = False

                wb = app.books.open(template_file, read_only=False)
                ws = wb.sheets('Field Analysis')

                # Loop through all shapes and clear their triggers
                for shp in ws.shapes:
                    shp.api.OnAction = ""

                # Save as .xlsx and reenable alerts
                wb.api.SaveAs(str(self.wb_path.resolve()), FileFormat=51)

                app.api.DisplayAlerts = True
        else:

            # Create a copy of the template as .xlsm
            shutil.copy(template_file, self.wb_path)
        
        progress_bar.update(3)
        
        


    def _add_field_metadata(self, progress_bar: tqdm):

        """
            Adds field metadata to Field Analysis workbook

        Parameters
        ----------

        progress_bar
            A tqdm progress bar object
        
        """

        # --------------------------------------------------
        # Set variables

        df = self.field_metadata
        wb = self.wb
        ws = wb.sheets('Field Analysis')



        # --------------------------------------------------
        # Add basic metadata


        # Set Name
        ws.range("Name").value = self.name

        # Add Data Description metadata
        ws.range("Dimensions").options(transpose=True).value = list(self.df_metadata.values())
        

        # Add all header values except Record List
        headers = self.source_data.columns.to_list()[1:]
        ws.range("Headers_Start").value = headers

        progress_bar.update(1)



        # --------------------------------------------------
        # Set Field Analysis theme

        self._expand_range(name="Theme", ws=ws)

        self._set_theme(ws.range("Theme"))

        progress_bar.update(1)



        # --------------------------------------------------
        # Split base analysis into Field Analysis sections

        composition_df = df.loc[["Data type", "Distinct %", "Missing %", "Memory Usage %", "Memory Usage", "Distinct", "Count", "Missing"]]
        summary_stats_df = df.loc[["Mean", "Median", "Mode", "Standard Deviation", "Variance"]]
        percentiles_df = df.loc[["Min", "5%", "25%", "50%", "75%", "95%", "Max", "Range", "IQR"]]

        progress_bar.update(1)




        # --------------------------------------------------
        # Format metadata placeholders
        
        columns_to_format = self.columns-3
        if columns_to_format > 0:
            format_from = ws.range("FormatRange")
            format_to = (ws.range("FormatRange").offset(0, 1).resize(None, columns_to_format))
            format_from.api.Copy()
            format_to.paste(paste='formats')

        # Clear clipboard and move selection back to upper left
        wb.api.CutCopyMode = False
        ws.range("Headers_Start").api.Select()

        progress_bar.update(1)


        # --------------------------------------------------
        # Add Field Analysis sections to workbook

        ws.range("Composition")[0, 0].offset(0,1).value = composition_df.values
        ws.range("Summary_Stats")[0, 0].offset(0,1).value = summary_stats_df.values
        ws.range("Percentiles")[0, 0].offset(0,1).value = percentiles_df.values

        progress_bar.update(1)



    def _add_overview(self, progress_bar: tqdm):

        """
            Adds overview_metadata to the Overview worksheet
        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object
        
        """


        progress_bar.update(2)


        # --------------------------------------------------
        # Setup variables and worksheet

        df = self.input_df
        wb = self.wb
        ws = wb.sheets("Overview")
        overview_table = ws.tables["tbl_Overview"]

        col_order = ['Field', 'Definition', 'Field Notes', 'Data type', 'Distinct %', 'Missing %', 'Memory Usage %', 'Memory Usage', 'Distinct', 'Count', 'Count %', 'Missing', 'Mean', 'Median', 'Mode', 'Standard Deviation', 'Variance', 'Min', '5%', '25%', '50%', '75%', '95%', 'Max', 'Range', 'IQR']

        # Activate worksheet and set theme
        self._set_theme(ws.range('tbl_Overview[#Headers]'))


        # --------------------------------------------------
        # Configure overview df, reorder columns

        df = self.overview_metadata
        df['Field Notes'], df['Definition'], df['Field'] = None, None, df.index
        df = df[col_order]


        progress_bar.update(2)


        # --------------------------------------------------
        # Add metadata/overview table, configure the Field Notes/Definitions columns

        ws.range("Metadata").options(transpose=True).value = list(self.df_metadata.values())
        overview_table.update(df, index=False)
        
        
        # Add Formulas to pull field definitions, notes into overview table
        ws.range('tbl_Overview[Definition]')[0].formula = r'=IF(INDEX(Definitions,1,MATCH([@Field],Headers,0))="Definition","",INDEX(Definitions,1,MATCH([@Field],Headers,0)))'
        ws.range('tbl_Overview[Field Notes]')[0].formula = r'=IF(INDEX(Notes,1,MATCH([@Field],Headers,0))="Note","",INDEX(Notes,1,MATCH([@Field],Headers,0)))'
        


        progress_bar.update(1)



    def _add_source_data(self, progress_bar: tqdm):

        """
            Adds source_data to the Field Analysis workbook

        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object

        """

        # --------------------------------------------------
        # Set variables, dtype conversion if necessary

        ws = self.wb.sheets("Field Analysis")
        df = self.source_data


        # Convert fields with datatypes that Excel doesn't support to string
        unsupported_dtypes = df.select_dtypes(exclude=['number', 'bool', 'datetime64']).columns
        df[unsupported_dtypes] = df[unsupported_dtypes].astype(str)

        
        progress_bar.update(3)
        


        # --------------------------------------------------
        # Add source data, format cells in Record List/Record Hash/HasBlank columns

        source_table = ws.tables["tbl_SourceData"]
        source_table.update(df, index=False)
        

        progress_bar.update(2)



        
        # --------------------------------------------------
        # Set formatting for tbl_SourceData and added columns

        
        ws.range("tbl_SourceData[[#All]]").api.HorizontalAlignment = HAlign.xlHAlignCenter
        record_list = ws.range('tbl_SourceData[[#All], [Record List]]')
        recordhash_and_hasblank = ws.range('tbl_SourceData[[#All], [index]:[Record Hash]]')

        # Reduce contrast to subdue added fields
        for dimmed_range in [record_list, recordhash_and_hasblank]:
            self._greyscale_range(dimmed_range)


        progress_bar.update(5)



    def _create_composition_plot(self, input_df: pd.Series) -> Figure:
        
        """
            Creates a composition table from a dataframe

        Args:
            input_series (pd.Series): A pandas Series

        Returns:
            Figure: A matplotlib Figure object
        """



        # --------------------------------------------------
        # Setup values


        # Font size
        font_size = 24

        # Prepare plot values
        counts = input_df.squeeze().value_counts()  # type: ignore
        top_5 = counts.head(5)
        total_entries = len(input_df)
        other_count = total_entries - top_5.sum()

        # Assemble plot values
        categories = list(top_5.index) + ["Other"]
        values = list(top_5.values) + [other_count]

        
        
        # --------------------------------------------------
        # Setup plot
        
        # Initialize the plot
        fig, ax = plt.subplots(figsize=(8, 8))  # type: ignore

        y_pos = range(len(categories))[::-1]

        # Add bars to plot
        bars = ax.barh(y_pos, values, color=self.theme_color, height=0.5, edgecolor='silver')  # noqa: F841

        
        # --------------------------------------------------
        # Adjust Formatting
        
        # Remove spines
        for spine in ax.spines.values():
            spine.set_visible(False)

        # Remove other extra plot elements
        ax.set_xticks([])
        ax.set_yticks([])
        ax.set_title("")

        # Make enough room for text on the left so they don't overlap
        plt.subplots_adjust(left=0.4, right=0.9)  # type: ignore

        max_val = max(values)



        # --------------------------------------------------
        # Setup mpl data bars

        for y, cat, val in zip(y_pos, categories, values):
            pct = (val / total_entries) * 100

            # Truncate long category names
            display_cat = str(cat)
            if len(display_cat) > 6:
                display_cat = display_cat[:5] + ".."

            # Add labels and adjust left to prevent overlap
            ax.text( -0.55, y, display_cat, color="white", va="center", ha="left", fontsize=font_size, transform=ax.get_yaxis_transform(), )

            # Add Percentages
            ax.text( -0.05, y, f"{pct:.0f}%", color="white", va="center", ha="right", fontsize=font_size, transform=ax.get_yaxis_transform(), )

            # Add Counts to the right of the bars
            ax.text(val + max_val * 0.02, y, str(val), color="white", va="center", ha="left", fontsize=font_size, )

        return fig



    def _create_histogram_plot(self, input_series: pd.Series) -> Figure:

        """
            Creates a histogram from a dataframe

        Args:
            input_series (pd.Series): A pandas Series

        Returns:
            Figure: A matplotlib Figure object
        """


        # --------------------------------------------------
        # Setup plot area, plot

        fig, ax = plt.subplots(figsize=(5, 5))  # type: ignore
        ax.set_axis_off()

        # Plot a histogram
        sns.histplot(x=input_series, color=self.theme_color, stat="density", alpha=0.5, ax=ax)

        # Layer the KDE line
        sns.kdeplot(x=input_series, color="silver", linewidth=3, ax=ax, warn_singular=False)

        
        # --------------------------------------------------
        # Add additional plot details

        # Add vertical mean line
        mean_val = input_series.mean()
        ax.axvline(mean_val, color="silver", linestyle=":", linewidth=2)

        # Remove tick labels
        ax.tick_params(left=False, bottom=False, labelleft=False, labelbottom=False)

        # Add Min and Max text at the bottom corners
        min_val = input_series.min()
        max_val = input_series.max()

        ax.text(0, -0.05, f"Min {min_val:g}", transform=ax.transAxes, fontsize=16, color="silver", ha="left", va="top",) 
        ax.text( 1, -0.05, f"Max {max_val:g}", transform=ax.transAxes, fontsize=16, color="silver", ha="right", va="top", )

        return fig



    def _add_small_plot(self, fig: Figure, target_range: xw.Range, name: str):

        """
            Adds a small chart to an Excel cell that is centered 
            at 90% of the size of the cell

        Args:
            fig (Figure): A matplotlib Figure object
            target_range (Range): An Excel cell
        """
        # --------------------------------------------------
        # Calculate size/position

        # Calculate 90% of cell dimensions
        target_width = target_range.width * 0.9
        target_height = target_range.height * 0.9

        # Calculate position to center the picture
        target_left = target_range.left + (target_range.width / 2) - (target_width / 2)
        target_top = target_range.top + (target_range.height / 2) - (target_height / 2)


        # --------------------------------------------------
        # Add the picture to the sheet

        pic = target_range.sheet.pictures.add(
            fig,
            name=name,
            left=target_left,
            top=target_top,
            width=target_width,
            height=target_height,
        )

        # Set placement to xlMoveAndSize
        try:
            pic.api.Placement = 1
        except AttributeError:
            pass



    def _add_large_plot(self, title: str, fig: Figure):
            
            """
                Creates a worksheet and adds an additonal_plot

            """

            
            # Set vars
            wb = self.wb
            assert wb is not None


            # Create a copy of the ws template
            ws = wb.sheets("SinglePlot").copy(name=title)
            ws.visible = True

            # Make it visible and set theme
            # ws.activate()
            self._set_theme(ws.range("Theme"))



            # Set target range, title, and autofit title range
            plot_range = ws.range("SinglePlot")
            ws.range("SinglePlotTitle").value = title
            ws.range("SinglePlotTitle").columns.autofit()
            
            ws.pictures.add(fig, 
                            name=title, 
                            update=True,
                            left=plot_range.left, 
                            top=plot_range.top)



    def _add_plots(self, progress_bar: tqdm):
        
        """
            Adds plots to an Excel range
        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object

        """


        # --------------------------------------------------
        # Add additional plots

        for plot in self.additional_plots.keys():
            self._add_large_plot(title=plot, fig=self.additional_plots[plot])

            progress_bar.update(1)
        

        
        # --------------------------------------------------
        # Set vars, start performance logging

        ws = self.wb.sheets('Field Analysis')
        df = self.source_data.iloc[:, 1:-3]
        self._log_performance(log_type='plots')
        
        
        
        # Set initial ranges for added plots and ensure they aren't hidden
        histogram_range = ws.range("Histogram")
        composition_range = ws.range("CompositionTable")
        
        histogram_range.api.EntireRow.Hidden = False
        composition_range.api.EntireRow.Hidden = False


        # --------------------------------------------------
        # Add plots for all except added columns

        for col in df.columns:


            # --------------------------------------------------
            # Add Composition Table Plot

            composition_table = self._create_composition_plot(df[col])
            self._add_small_plot(target_range=composition_range, fig=composition_table, name=f'composition_{col}')
          
            self._log_performance(log_type='plots', details={f"composition_{col}": {'Field': col,
                                                                                    'Dtype': str(df[col].dtype), 
                                                                                    'Plot Type': 'composition', 
                                                                                    'Performance': ''}})


            if pd.api.types.is_numeric_dtype(df[col]):



                # --------------------------------------------------
                # Add Histogram Multiple

                histogram = self._create_histogram_plot(df[col]) 
                self._add_small_plot(target_range=histogram_range, fig=histogram, name=f'histogram_{col}')
                self._log_performance(log_type='plots', details={f"histogram_{col}": {'Field': col,
                                                                                      'Dtype': str(df[col].dtype), 
                                                                                      'Plot Type': 'histogram', 
                                                                                      'Performance': ''}})



                
            # --------------------------------------------------
            # Increment Target Ranges/progress bar

            histogram_range = histogram_range.offset(0, 1)
            composition_range = composition_range.offset(0, 1)
            progress_bar.update(1)

            



    def _configure_pivot(self, progress_bar: tqdm):
            
        """
            Configures the pivot table
        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object

        """


        # --------------------------------------------------
        # Set vars/intialize logging/set theme


        ws = self.wb.sheets('Pivot')
        pt = ws.api.PivotTables('pvt_Pivot')
        df = self.input_df.copy()


        # Initialize logging
        self._log_performance(log_type='pivot')
        
        # Set theme
        self._set_theme(ws.range("Theme"))

        # Fields to be added to the pivot tables
        pivot_fields = df.columns.to_list()[:min(10, self.columns)]
        


        # --------------------------------------------------
        # Refesh pivotcache and set field limit warning

        
        # Refresh the PivotCache
        pt.PivotCache.BackgroundQuery = False
        pt.RefreshTable()
    
        
        # Show field limit warning for datasets with >10 columns
        if len(df.columns) <= 10:
            ws.range("DefaultWarn").value = ''
        else:
            ws.range("DefaultWarn").value = 'Pivots only show first 10 fields by default'


        # --------------------------------------------------
        # Add fields to pivot table


        for field in pivot_fields:

            # Add field to row section of pivot table
            pt.PivotFields(field).Orientation = xw.constants.PivotFieldOrientation.xlRowField
            

            # Remove subtotals if possible
            try:
                pt.PivotFields(field).Subtotals = tuple(False for _ in range(12))
            except Exception:
                pass
            
            # Log performance, update progress bar
            self._log_performance(log_type='pivot', details={field: {'Field': field,
                                                                     'Dtype': str(df[field].dtype), 
                                                                     'Performance': ''}})

            progress_bar.update(1)
            


        # --------------------------------------------------
        # Format pivot tables
        
        # Collapse fields if possible, starting from the end
        for field in pivot_fields[::-1]:
            try:
                pt.PivotFields(field).ShowDetail = False
            except Exception:
                pass

        

        # Set ranges for formatting
        pivot_range = ws.range(pt.TableRange1.Address)
        pivot_headers = pivot_range[0,:]            
        header_column = pivot_range[: , 0]

        # Format headers, first column, and column width
        pivot_headers.api.WrapText = True
        pivot_headers.api.HorizontalAlignment = HAlign.xlHAlignCenter
        pivot_headers.api.VerticalAlignment = VAlign.xlVAlignCenter
        header_column.api.HorizontalAlignment = HAlign.xlHAlignLeft
        pivot_range.columns.autofit()

        # Set last three pivot columns width to 13 and center
        pivot_range[:, -3:].column_width = 13
        self._greyscale_range(pivot_range[:, -3:])
        ws.range(pivot_range[:, -2:].address).api.HorizontalAlignment = HAlign.xlHAlignCenter
        
        progress_bar.update(1)
            


    def _expand_range(self, name: str, ws: xw.Sheet, extra_columns: int = 0):
        
        """
            Expands a named range by the amount of columns in a df +/- extra_columns

        Parameters
        ----------
        name : str
            Name of named range
        ws : xw.Sheet
            Worksheet of named range
        extra_columns : int, optional
            Adjustment to the amount of columns to expand the named range by
            Defaults to 0
        """

        # Adding 1 Chooses the first source data column
        ws.range(name).resize(row_size=None, column_size=self.columns + 1 + extra_columns).name = name



    def _set_theme(self, input_range: xw.Range):

        """
            Sets the background/font colors of a range object to the current xleda theme.

        """

        input_range.color = self.theme_color

        if self.black_text:
            input_range.font.color = '#000000'


        
       


    def _greyscale_range(self, input_range: xw.Range):
        
        """
            Formats a range as grey on grey

        """

        input_range.color = '#262626'
        input_range.font.color = '#898989'
        


    def _configure_workbook(self, progress_bar: tqdm):

        """
            Configures an xleda workbook for use

        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object
            
        """


        # --------------------------------------------------
        # Set variables, activate ws, and set theme

        ws = self.wb.sheets("Field Analysis")
        ws.activate()


        progress_bar.update(1)



        # --------------------------------------------------
        # Adjust Named Ranges to fit dataframe size and add EDA placeholder values 

        
        # Set record hash named range for Record List
        ws.range("tbl_SourceData[Record Hash]").name = "RecordHashes"
        

        # Expand named ranges to fit number of columns and set input prompt for notes/definitions
        expand_ranges = (["FieldList" + str(i) for i in range(1, 9)] + ["FieldRange", "Notes", "Definitions", "Headers"])

        for name_range in expand_ranges:
            
            self._expand_range(name=name_range, ws=ws, extra_columns=-1)

            # Add placeholder values if necessary.
            if name_range in ["Notes", "Definitions"]:    
                ws.range(name_range).value = name_range[:-1]

            elif name_range.startswith("FieldList"):
                ws.range(name_range).value = "FALSE"


        # Set FieldLists range so that it also includes the names on the left
        # Used to recreate completed examples
        self._expand_range(name="FieldLists", ws=ws, extra_columns=1)


        progress_bar.update(1)



        # --------------------------------------------------
        # Configure UI


        # Show/Hide Data Size Warning
        if self.warning:
            ws.range("Warning").value = self.warning_msg
            ws.range("Warning").api.EntireRow.Hidden = not self.warning


        # Orient toggles, and collapse subsections
        ws.range("Toggles").api.Orientation = 0
        ws.range("TopToggle").api.Orientation = -90

        for excel_range in ["Data_Description", "Field_Notes", "Composition", "Summary_Stats", 
                            "Percentiles", "Field_Lists", "Compiled_Lists"]:
                        ws.range(excel_range).api.EntireRow.Hidden = True

        progress_bar.update(1)

    def _color_printer(self, text: str):
        
        """
            Prints to the console using theme_color text. 

        Parameters
        ----------
        text : str
            Text to be printed

        """
        
        print(color_formatter(text=text, theme=self.tqdm_theme))

    def _log_performance(self, log_type: str, details: dict = {}):

        """
            Logs production performance data

        """

        now = time.time()

        # Identify log storage
        log_storage = eval(f"self.performance_{log_type}")

        # If no details are provided, initialize the last variable and exit
        if not details:
            log_storage['last'] = now
            return
        
        # If details are provided, handle each log type

        # Add timing to section
        if log_type == 'section':
            details[list(details.keys())[0]] =  now - log_storage['last']
                               
        else:
               
            # Add timing to plot or pivot
            for k, v in details.items():
                for key, value in v.items():
                    if not value:
                        details[k][key] = now - log_storage['last']

        # Add to log storage, set 'last'
        log_storage |= details


        log_storage['last'] = now


    def _log_close(self):

        """
           Closes performance logging by converting logs to dataframes

        """
        


        # Create dataframes of performance logs
            
        for log_type in ['pivot', 'plots', 'section']:
            
            # Remove 'last' key
            del eval(f"self.performance_{log_type}")['last']

            # Convert section logging to dataframe
            if log_type == 'section':
                self.performance[log_type] = pd.DataFrame.from_records([eval(f"self.performance_{log_type}")])
            
            else:

                # Convert pivot/plots logging to dataframes
                self.performance[log_type] = pd.DataFrame.from_dict(eval(f"self.performance_{log_type}"),
                                                                    orient='index')

            # Add dataset details to the dataframe
            self.performance[log_type]['Dataset'] = self.name
            self.performance[log_type]['Rows'] = self.rows
            self.performance[log_type]['Columns'] = self.columns
            self.performance[log_type]['ProductionTime'] = time.time() - self.start_time





    def _add_progress_bar(self, desc: str, total: float) -> tqdm:
        
        """
            Creates a tqdm progress bar

        Returns
        -------
        tqdm
            A tqdm object
        """
        

        padded_desc = f"{desc}{' ' * (25-len(desc))}"
        fmt = "{desc} | {percentage:3.0f}% | {bar} | {elapsed}"


        # Pad the raw desc before applying color so ANSI codes don't shift bar alignment
        padded_desc = f"{desc:<30}"

        # Create a tqdm instance
        pbar = tqdm(
            total=total,
            desc=padded_desc,
            bar_format=fmt,
            colour=self.tqdm_theme,
            # ncols=100,
            # dynamic_ncols=True
        )
        
        return pbar
    

    def _create_wb(self):
        """
            Creates an xdleda workbook from a given dataframe.  

            Workbook is saved in current directory
        
        """


        # --------------------------------------------------   
        # Construct entry message/Create blank template
        

        # Initial output
        if self.name == 'xleda':
            self.entry_message += f"\nPreparing an xleda workbook located at:\n    {self.wb_path}\n"
        else:
            self.entry_message += f"\nPreparing an xleda workbook with {self.name} data located at:\n    {self.wb_path}"

        self.entry_message += (f"\n\nProcess started at {time.strftime('%H:%M:%S')}\n")


        self._color_printer(self.entry_message)
       

        # --------------------------------------------------
        # Initialize Template
        
        pbar = self._add_progress_bar(desc="Creating Workbook...", total=11)
        self._create_blank_template(progress_bar=pbar)
        
        pbar.update(3)
        

        # Open Excel using a context manager to while creating the workbook
        with xw.App(visible=False, add_book=False) as app:

            

            app.display_alerts = False
            

            # Set vars
            wb = app.books.open(self.wb_path, read_only=False)
            self.wb = wb


            self._log_performance(log_type='section', details={'creating_workbook': ""})
            pbar.update(2)
            pbar.close()


            
            # --------------------------------------------------
            # Add metadata

            # Create Progress Bar

            with self._add_progress_bar(desc="Adding Metadata...", total=10) as pbar:

                self._add_field_metadata(progress_bar=pbar)
                
                self._log_performance(log_type='section', details={'add_field_metadata': {'add_field_metadata': ""}})


                self._add_overview(progress_bar=pbar)
                self._log_performance(log_type='section', details={'add_overview': {'add_overview': ""}})


            # --------------------------------------------------
            # Add source data

            with self._add_progress_bar(desc="Adding Source Data...", total=10) as pbar:

                self._add_source_data(progress_bar=pbar)
                self._log_performance(log_type='section', details={'adding_source_data': {'adding_source_data': ""}})


            # --------------------------------------------------
            # Add Plots
        
            with self._add_progress_bar(desc="Adding Plots...", total=self.columns + len(self.additional_plots.keys())) as pbar:
            
                self._add_plots(progress_bar=pbar)
                self._log_performance(log_type='section', details={'adding_plots': {'adding_plots': ""}})


            # --------------------------------------------------
            # Configure Workbook

            with self._add_progress_bar(desc="Configuring Workbook...", total=min(10, self.columns) + 5) as pbar:

                self._configure_pivot(progress_bar=pbar)
                self._log_performance(log_type='section', details={'configuring_pivots': {'configuring_pivots': ""}})

                self._configure_workbook(progress_bar=pbar)
                self._log_performance(log_type='section', details={'configuring_pivots': {'configuring_workbook': ""}})

                # Save workbook
                app.display_alerts = True
                wb.save(self.wb_path)
                
                duration = time.time() - self.start_time
                pbar.update(1)



        # --------------------------------------------------
        # Compile closing messsage

        
        self.exit_msg += f"\n\nxleda workbook created in {int(duration)} seconds.\n"

        
        if self.open_wb:
            

            # Create a new Excel instance
            app = xw.App(visible=True, add_book=False) 

            # Open the workbook
            wb = app.books.open(self.wb_path)


        if self.additional_plots:
            
            self.exit_msg += "Additional plots included:"
            
            for plot in self.additional_plots.keys():
                self.exit_msg += f"    {plot}"
        
        
        # Close logs
        self._log_close()
        
        # Print closing message
        self._color_printer(self.exit_msg + separator)
        




    def _export_analysis(self):

        """
            Exports data from an xleda workbook into self.export_dict

        """
       
        
        # --------------------------------------------------
        # Setup placeholder vars

        
        definitions = {}
        notes = {}
        lists = {}
        export_order = ['description', 'definitions', 'notes', 'lists', 'field_metadata', 'overview_metadata', 'source_data', 'altered_source_data']


        
        # --------------------------------------------------
        # If the file is not found return messaging

        if not self.wb_path.is_file():

            self._color_printer(f"File not found at {self.wb_path}\n")
            self._color_printer("wb().export_dict will be limited" + separator)

            sys.exit()



        # --------------------------------------------------
        # Configure output


        start_time = time.time()
        self._color_printer(separator + f"\nExport started at {time.strftime('%H:%M:%S')}\n")
        



        # --------------------------------------------------
        # Open Excel/export data using a context manager


        with xw.App(visible=False, add_book=False) as app:

            
            with self._add_progress_bar(desc="Reading workbook...", total=10) as pbar:
                
                pbar.update(2)

                try:
                    book = app.books.open(self.wb_path)
                except FileNotFoundError:
                    self._color_printer(f"File not found at {self.wb_path}")
                    sys.exit()

                ws = book.sheets("Field Analysis")
                    
                pbar.update(3)



                # --------------------------------------------------
                # Export lists/notes/definitions from workbook


                # Definitions/Notes/Fields
                field_notes = ws.range("Notes").value
                field_definitions = ws.range("Definitions").value
                fields = ws.range("FieldRange").value

                # Compiled Lists
                compiled_lists_names = ws.range("Compiled_Lists").value
                compiled_lists = ws.range("Compiled_Lists").offset(0, 1).value
                
                pbar.update(1)



                # --------------------------------------------------
                # Extract description

                description = ws.range("Description").value

                

                # --------------------------------------------------
                # Compile definitions

                for (i, definition) in enumerate(field_definitions):
                    if definition != "Definition":
                        definitions[fields[i]] = definition

                pbar.update(1)



                # --------------------------------------------------
                # Compile notes

                for (i, note) in enumerate(field_notes):
                    if note != "Notes":
                        notes[fields[i]] = note

                pbar.update(1)



                # --------------------------------------------------
                # Compile lists

                for (i, list_name) in enumerate(compiled_lists_names):
                    
                    if list_name:
                        lists[list_name[:-3]] = ast.literal_eval(compiled_lists[i])

                pbar.update(1)



                # --------------------------------------------------
                # Extract altered_source_data
                
                altered_source_data = ws.tables['tbl_SourceData'].range.options(pd.DataFrame, index=False).value

                pbar.update(1)



        # --------------------------------------------------
        # Prepare exports


        # Add/sort export_dict fields
        self.export_dict |= {key: eval(key) for key in export_order if key not in self.export_dict.keys()}
        self.export_dict = {key: self.export_dict[key] for key in export_order if key in self.export_dict}

        # Add class properties
        self.description = description
        self.definitions = definitions
        self.notes = notes
        self.lists = lists
        self.altered_source_data = altered_source_data
        

        
        # --------------------------------------------------
        # Print output
        
        duration = time.time() - start_time
        self._color_printer(f"\nExport completed after {int(duration)} seconds" + separator)



class FieldAnalysis(wb):

    """
    
    """
    def __init__(self, input_df: pd.DataFrame, export: bool = False, **kwargs):

        """ 
            A placeholder for legacy FieldAnalysis support.

            Creating a separate Field Analysis configuration is no longer required.
            
            Use xleda.wb to both configure and create a workbook
             
        """
        warn_print(separator + 
              "\nCreating a separate Field Analysis configuration is no longer required.\n"
              "Use xleda.wb to both configure and create a workbook\n" + separator)

        super().__init__(input_df=input_df,
                         **kwargs)


    def create_workbook(self):

        """
            FieldAnalysis() has been replaced by wb()
            
            Use xleda.wb() to both configure and create a workbook

        """ 

        self._color_printer(separator + "FieldAnalysis() has been replaced by wb()")
        self._color_printer("\nUse xleda.wb() to both configure and create a workbook")


    
    def export_analysis(self) -> dict[str, Any]:

        self._color_printer(separator + "FieldAnalysis() has been replaced by wb()")
        self._color_printer("\nUse xleda.wb(export=True) to export from an xleda workbook")

        self._export_analysis()

        return self.export_dict
    
