import seaborn as sns
import pandas as pd
import shutil
from pathlib import Path
import ast
import sys
import random

import matplotlib.pyplot as plt
import matplotlib as mpl
from matplotlib.figure import Figure

import xlwings as xw
from xlwings import Range

from rich.progress import (
    Progress, BarColumn, TextColumn, 
    TimeElapsedColumn, TaskID)
from rich.console import Console
from rich.style import Style


import time






# --------------------------------------------------
# Setup/global variables

# Rich console
console = Console()

# Field Analysis template
template_file = Path(__file__).parent / "xleda_template.xlsm"

separator = "\n" + ("-" * 100) + "\n"

# Set matplotlib theme
mpl.use("Agg")
plt.style.use("dark_background")



class FieldAnalysis():

    """
    A class that represents an xleda Field Analysis workbook.

    Methods:

        create_workbook: Creates an xdleda workbook from a given dataframe. Workbook 
                         is saved in current directory

        export_analysis: Export notes, lists, data from an xleda field analysis workbook



    """

    def __init__(self, 
                 input_df: pd.DataFrame, 
                 name: str, 
                 theme_color: str = "#05233E", 
                 large_report: bool = False, 
                 overwrite: bool = False, 
                 wb_path: Path= Path().cwd(),
                 add_plots: dict[str, Figure] = {}):

        """Configures an xleda workbook

        Args:

            name (str): Name of the workbook to be created.

            theme_color (str): A hexidecimal color used for charts/accent color.  
                               Dark colors work better.    
                               Use theme_color='random' for random colors.

            large_report (bool): Used to override default limits of 100,000 rows/100 columns. 
                                 Sets limits to 1,000,000 rows/16,000 columns. 
                                 Defaults to False

            wb_path (Path): Pathlib path of xleda workbook.  
                            Defaults to current working directory.

            overwrite (bool): Whether to overwrite existing reports with the same name. 
                              Defaults to False

            add_plots (dict[str, Figure]): Additional plots to be included 
                Uses "{'plot1_name': Plot1Figure, 'plot2_name': Plot2Figure}" format.  
                Each entry will get it's own worksheet.  
                No resizing or syling is done for plots added this way.

        
        """

        # Set base properties
        self.name: str = name
        self.large_report: bool = large_report
        self.overwrite: bool = overwrite
        
        # Set Template Path
        self.wb_path: Path = (wb_path / self.name).with_suffix(".xlsm")


        # Set theme
        if theme_color == 'random':
            self.theme_color = "#{:06x}".format(random.randint(0, 0xFFFFFF))
        else:
            self.theme_color: str = theme_color
        self.input_df: pd.DataFrame = input_df.copy()
        self.additional_plots: dict[str, Figure] | dict = add_plots

        # Configure Theme
        self.theme_style: Style = Style(color=self.theme_color[:7])
        self.silver_style = Style(color='#C0C0C0')
        self.black_text: bool = self._use_black_text()


        # Add placeholder properties
        self.warning_msg: str = ""
        self.warning: bool = False
        self.rows = len(input_df)
        self.columns = len(input_df.columns)
        self.wb: xw.Book | None = None
        self.field_analysis_ws: xw.Sheet | None = None
        
        # Configure Source Data
        self.source_df: pd.DataFrame = self._configure_source_data()

        # Create base analysis
        self.base_analysis: pd.DataFrame = self._create_base_analysis()



    def _use_black_text(self) -> bool:
        """
        Converts theme color to RGB and calculates 
        whether black text is required.

        """

        # Remove '#' if present and convert hex to RGB
        hex_color = self.theme_color.lstrip('#')

        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        # Standard formula for perceived brightness
        brightness = (r * 0.299 + g * 0.587 + b * 0.114) / 255
        
        # Use black text for light backgrounds, white for dark ones

        return brightness > 0.5



    def _configure_source_data(self) -> pd.DataFrame:
        
        """
        Configures source data for analysis and subsamples it if necessary


        Returns:
            source_df: A pandas dataframe object
        
        """


        # --------------------------------------------------
        # Configure variables

        rows = self.rows
        columns = self.columns
        df = self.input_df

        above_default = rows > 100_000 or columns > 100
        above_limit = rows > 1_000_000 or columns > 16_000
          

        # Source data is above default limits
        if above_default and not self.large_report and not above_limit:

            self.warning = True
            self.warning_msg = "This is only showing a sample because it is larger than the default limits of 100,000 rows/100 columns.  See documentation for details."
            rows = min(rows, 100_000)
            columns = min(columns, 100)
            

        # Source data is larger than Excel's limits
        elif above_limit:

            self.warning = True
            self.warning_msg = "This is only showing a sample because it is larger than Excel's limits of 1,000,000 rows/16,000 columns.  See documentation for details."
            rows = min(rows, 1_000_000)
            columns = min(columns, 16_000)


        # Subsample df if needed, record adjusted rows/columns
        df = (df.iloc[:, :columns].sample(n=rows).sort_index())

        self.rows = rows
        self.columns = columns

        # Add fields to Source Data to track individual records
        df.insert(loc=0, column="Record List", value="False")
        df["Record Hash"] = pd.util.hash_pandas_object(df, index=False)

        return df



    def _create_base_analysis(self) -> pd.DataFrame:
        """Produces the base analysis dataframe for an input dataframe"""


        # --------------------------------------------------
        # Collect metadata

        df = self.source_df.iloc[:, 1:-1]

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
                "Memory Usage %": df.memory_usage(deep=True, index=False) / df.memory_usage(deep=True).sum(),
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



    def _create_blank_template(self):

        """
        Creates a blank Field Analysis template, overwriting if necessary

        """

        
        # Return an error if there's an existing file and no overwrite flag

        if self.wb_path.is_file() and not self.overwrite:

            console.print(f"[bold red]Error: There is already a workbook named {self.wb_path}![/bold red]")
            console.print("Use overwrite=True or rename/remove the existing workbook")
            
            sys.exit()


        # Delete the file if there's an overwrite flag, return error if it's open

        elif self.wb_path.is_file() and self.overwrite:
            try:
                self.wb_path.unlink(missing_ok=True)

            except PermissionError:
                
                console.print("[bold red]Error: The workbook cannot be overwritten while open![/bold red]")

                sys.exit()

        
        # Create parent directories if necessary and a blank copy of the template
        self.wb_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(template_file, self.wb_path)



    def _configure_template(self, progress: Progress, task_id: TaskID):
        """
        Set template theme and format placeholders

        """

        # Set vars
        wb = self.wb
        assert wb is not None
        
        self.field_analysis_ws = wb.sheets("Field Analysis")
        ws = self.field_analysis_ws


        # Set Name
        ws.range("Name").value = self.name
        

        # Set Theme
        for sheet in wb.sheets:
            sheet.range("Theme").color = self.theme_color
            if self.black_text:
                sheet.range("Theme").font.color = '#000000'


        # Add dimentions to workbook
        ws.range("Dimensions").options(transpose=True).value = [self.rows, self.columns]



        progress.update(task_id, completed=6, refresh=True)


        # Format metadata placeholders 
        columns_to_format = self.columns-3
        if columns_to_format > 0:
            format_from = ws.range("FormatRange")
            format_to = (ws.range("FormatRange").offset(0, 1).resize(None, self.columns-3))
            format_from.api.Copy()
            format_to.api.Select()
            ws.api.Paste()

        progress.update(task_id, completed=8, refresh=True)


        # Clear clipboard, set selection to top, and add header values
        wb.api.CutCopyMode = False
        ws.range("Headers_Start").api.Select()
        headers = self.source_df.columns.to_list()[1:]
        ws.range("Headers_Start").value = headers

        progress.update(task_id, completed=10, refresh=True)



    def _add_overview(self, progress: Progress, task_id: TaskID):

        """
        Adds an Overview worksheet that includes a transposed copy of the base analysis
        
        """


        progress.update(task_id, completed=1, refresh=True)


        # --------------------------------------------------
        # Setup variables and overview data


        # Set variables
        wb = self.wb
        assert wb is not None
        ws = wb.sheets("Overview")
        df = self.input_df


        col_order = ['Field', 'Definition', 'Field Notes', 'Data type', 'Distinct %', 'Missing %', 'Memory Usage %', 'Memory Usage', 'Distinct', 'Count', 'Count %', 'Missing', 'Mean', 'Median', 'Mode', 'Standard Deviation', 'Variance', 'Min', '5%', '25%', '50%', '75%', '95%', 'Max', 'Range', 'IQR']

        
        progress.update(task_id, completed=2, refresh=True)

        overview_metadata = {'Rows': self.rows,
                             'Columns': self.columns,
                             'Memory Usage': f"{df.memory_usage(deep=True).sum():,} bytes",
                             'Distinct Rows %': (len(df.drop_duplicates()) / len(df)),
                             'Missing %': float(df.isnull().mean().mean()),
                             }

        
        # Configure overview df, reorder columns
        df = self.base_analysis.T
        df['Field Notes'], df['Definition'], df['Field'] = None, None, df.index
        df = df[col_order]

        progress.update(task_id, completed=3, refresh=True)


        # --------------------------------------------------
        # Add metadata/overview table, configure the Field Notes/Definitions columns

        ws.range("Metadata").options(transpose=True).value = list(overview_metadata.values())
        overview_table = ws.tables["tbl_Overview"]
        overview_table.update(df, index=False)
        
        # Add Formulas to pull field definitions, notes into overview table
        ws.range('tbl_Overview[Definition]')[0].formula = r'=IF(INDEX(Definitions,1,MATCH([@Field],Headers,0))="Definition","",INDEX(Definitions,1,MATCH([@Field],Headers,0)))'
        ws.range('tbl_Overview[Field Notes]')[0].formula = r'=IF(INDEX(Notes,1,MATCH([@Field],Headers,0))="Notes","",INDEX(Notes,1,MATCH([@Field],Headers,0)))'
        
        
        progress.update(task_id, completed=4, refresh=True)



    def _add_metadata(self, progress: Progress, task_id: TaskID):

        """
        Add base analysis to Field Analysis workbook

        """

        # Set variables
        ws = self.field_analysis_ws
        df = self.base_analysis
        assert ws is not None


        # Split base analysis into Field Analysis sections
        composition_df = df.loc[["Data type", "Distinct %", "Missing %", "Memory Usage %", "Memory Usage", "Distinct", "Count", "Missing"]]
        summary_stats_df = df.loc[["Mean", "Median", "Mode", "Standard Deviation", "Variance"]]
        percentiles_df = df.loc[["Min", "5%", "25%", "50%", "75%", "95%", "Max", "Range", "IQR"]]

        progress.update(task_id, completed=5, refresh=True)



        # Add Field Analysis sections to workbook
        ws.range("Composition")[0, 0].value = composition_df.values
        ws.range("Summary_Stats")[0, 0].value = summary_stats_df.values
        ws.range("Percentiles")[0, 0].value = percentiles_df.values

        progress.update(task_id, completed=10, refresh=True)



    def _create_composition_plot(self, input_df: pd.DataFrame) -> Figure:
        """Creates a composition table from a dataframe

        Args:
            input_df (pd.DataFrame): A 1 column dataframe.

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
        fig, ax = plt.subplots(figsize=(8, 8))

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
        plt.subplots_adjust(left=0.4, right=0.9)

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



    def _create_histogram_plot(self, input_df: pd.DataFrame) -> Figure:

        """Creates a histogram from a dataframe

        Args:
            input_df (pd.DataFrame): A 1 column dataframe.

        Returns:
            Figure: A matplotlib Figure object
        """


        # --------------------------------------------------
        # Setup plot area, plot

        fig, ax = plt.subplots(figsize=(5, 5))
        ax.set_axis_off()

        # Plot a histogram
        sns.histplot( data=input_df, x=input_df.columns[0], color=self.theme_color, stat="density", alpha=0.5, ax=ax)

        # Layer the KDE line
        sns.kdeplot(data=input_df, x=input_df.columns[0], color="silver", linewidth=3, ax=ax, warn_singular=False)

        
        # --------------------------------------------------
        # Add additional plot details

        # Add vertical mean line
        mean_val = input_df[input_df.columns[0]].mean()
        ax.axvline(mean_val, color="silver", linestyle=":", linewidth=2)

        # Remove tick labels
        ax.tick_params(left=False, bottom=False, labelleft=False, labelbottom=False)

        # Add Min and Max text at the bottom corners
        min_val = input_df[input_df.columns[0]].min()
        max_val = input_df[input_df.columns[0]].max()

        ax.text(0, -0.05, f"Min {min_val:g}", transform=ax.transAxes, fontsize=16, color="silver", ha="left", va="top",) 
        ax.text( 1, -0.05, f"Max {max_val:g}", transform=ax.transAxes, fontsize=16, color="silver", ha="right", va="top", )

        return fig



    def _add_small_plot(self, fig: Figure, target_range: Range):
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


            # Create a copy of the ws template and make it visible
            ws = wb.sheets("SinglePlot").copy(name=title)
            ws.visible = True

            # Set target range, title, and autofit title range
            plot_range = ws.range("SinglePlot")
            ws.range("SinglePlotTitle").value = title
            ws.range("SinglePlotTitle").columns.autofit()
            
            ws.pictures.add(fig, 
                            name=title, 
                            update=True,
                            left=plot_range.left, 
                            top=plot_range.top)



    def _add_plots(self, progress: Progress, task_id: TaskID):
        
        """Adds plots to an Excel range

        Args:
            input_wb (xw.Workbook): Target Workbook for plots
            progress (Progress): A rich Progress object
            task_id: A rich TaskID object

        """


        # --------------------------------------------------
        # Add additional plots

        for plot in self.additional_plots.keys():
            self._add_large_plot(title=plot, fig=self.additional_plots[plot])

            progress.update(task_id, advance=1, refresh=True)
        

        
        # --------------------------------------------------
        # Set vars, activate fields analysis ws

        
        ws = self.field_analysis_ws
        assert ws is not None
        
        
        # Set initial ranges for added plots and ensure they aren't hidden
        histogram_range = ws.range("Histogram")
        composition_range = ws.range("CompositionTable")
        
        histogram_range.api.EntireRow.Hidden = False
        composition_range.api.EntireRow.Hidden = False


        # --------------------------------------------------
        # Add plots per column

        for col in self.source_df.iloc[:, 1:-1].columns:


            # --------------------------------------------------
            # Add Composition Multiple

            composition_table = self._create_composition_plot(self.source_df[[col]])
            
            self._add_small_plot(target_range=composition_range, fig=composition_table)

            if pd.api.types.is_numeric_dtype(self.source_df[col]):



                # --------------------------------------------------
                # Add Histogram Multiple

                histogram = self._create_histogram_plot(self.source_df[[col]])
                self._add_small_plot(target_range=histogram_range, fig=histogram)



            # --------------------------------------------------
            # Increment Target Ranges

            histogram_range = histogram_range.offset(0, 1)
            composition_range = composition_range.offset(0, 1)

            progress.update(task_id, advance=1, refresh=True)



    def _add_source_data(self, progress: Progress, task_id: TaskID):

        """
        Adds source data to the Field Analysis workbook

        """

        # --------------------------------------------------
        # Set variables

        ws = self.field_analysis_ws
        assert ws is not None
        progress.update(task_id, completed=1, refresh=True)

        
        # Convert df to string before writing to Excel to prevent 
        # issues with timedelta/random datatypes
        df = self.source_df.astype(str)

        


        # --------------------------------------------------
        # Add source data, format cells in Record List column

        source_table = ws.tables["tbl_SourceData"]
        source_table.update(df, index=False)
        if source_table.data_body_range is not None:
            source_table.data_body_range[0, 0].copy(destination=source_table.data_body_range[:, 0]) 


        progress.update(task_id, completed=10, refresh=True)



    def _iniialize_ui(self, progress: Progress, task_id: TaskID):

        """
        Initializes the Field Analysis UI for use

        """

        # --------------------------------------------------
        # Set variables

        ws = self.field_analysis_ws
        wb = self.wb
        assert ws is not None
        assert wb is not None


        progress.update(task_id, completed=1, refresh=True)


        # --------------------------------------------------
        # Adjust Named Ranges for Field List Formulas

        # Set record hash named range for Record List
        ws.range("tbl_SourceData[Record Hash]").name = "RecordHashes"

        # Expand Record List ranges to fit number of columns
        for name_range in ["FieldRange", "Notes", "Definitions", "Headers"]:
            ws.range(name_range).resize(row_size=1, column_size=self.columns).name = name_range


        # Expand FieldList ranges to fit number of columns
        for i in range(1, 9):
            excel_range = "FieldList" + str(i)
            ws.range(excel_range).resize(row_size=1, column_size=self.columns).name = excel_range

        progress.update(task_id, completed=5, refresh=True)


        # --------------------------------------------------
        # Format Record Hash columns for contrast with source data

        # Set Record Hash in lower section to LightHeader        
        lower_hash = ws.tables["tbl_SourceData"].header_row_range.last_cell  # type: ignore
        ws.range("LightHeader").copy(destination=lower_hash)
        lower_hash.value = "Record Hash"

        
        # Copy logo to last cell in FieldAnalysis Headers
        wb.sheets("Overview").range("Logo").copy(destination=ws.range((3, lower_hash.column)))
        
        # Autofit
        lower_hash.columns.autofit()


        # --------------------------------------------------
        # Configure UI

        # Show/Hide Data Size Warning
        if self.warning:
            ws.range("Warning").value = self.warning_msg
            ws.range("Warning").api.EntireRow.Hidden = not self.warning


        # Activate Field Analysis Worksheet and collapse field analysis subsections
        ws.activate()
        for excel_range in ["Data_Description", "Field_Notes", "Composition", "Summary_Stats", "Percentiles", "Field_Lists", "Compiled_Lists"]:
            
            ws.range(excel_range).offset(-2, -2).api.Orientation = 0
            ws.range(excel_range).api.EntireRow.Hidden = True


        progress.update(task_id, completed=9, refresh=True)


        # --------------------------------------------------
        # Save workbook

        wb.save(self.wb_path)

        progress.update(task_id, completed=10, refresh=True)



    def create_workbook(self):
        """
        Creates an xdleda workbook from a given dataframe.  

        Workbook is saved in current directory
        
        Returns:
            wb: An xlwings Book object
        
        """



        self._create_blank_template()

        # --------------------------------------------------
        # Setup progress bars

        start_time = time.time()
        
      
        # Initial output
        console.print(separator + f"\nPreparing an xleda workbook with {self.name} data", style=self.theme_style)
        console.print(f"\nProcess started at {time.strftime('%H:%M:%S')}\n", style=self.theme_style)
        

        # Configure progress bar table
        with Progress(TextColumn("{task.description}", style=self.theme_style),
                      BarColumn(pulse_style=self.silver_style, style=self.silver_style, complete_style=self.theme_style),
                      TextColumn("{task.percentage:>3.0f}%", style=self.theme_style),
                      "|",
                      TimeElapsedColumn(), # Displays elapsed time
                      console=console,
                      ) as progress:
            

            # Register tasks to get IDs
            task_create_template = progress.add_task("[self.theme_color]Creating Template...", total=10)
            task_add_metadata = progress.add_task("[self.theme_color]Adding Metadata...", total=10)
            task_add_plots = progress.add_task("[self.theme_color]Adding Plots...", total=self.columns + len(self.additional_plots.keys()))
            task_add_source_data = progress.add_task("[self.theme_color]Adding Source Data...", total=10)
            task_initialize_ui = progress.add_task("[self.theme_color]Initializing UI...", total=10)


            progress.update(task_id=task_create_template, completed=3, refresh=True)


            # --------------------------------------------------
            # Initialize Template

            with xw.App(visible=False, add_book=False) as app:

                
                # Set vars
                wb = app.books.open(self.wb_path, read_only=False)
                self.wb = wb
                self.field_analysis_ws = self.wb.sheets('Field Analysis')
                
                progress.update(task_id=task_create_template, completed=5, refresh=True)

                
                # --------------------------------------------------
                # Configure the template

                self._configure_template(progress=progress, task_id=task_create_template)


                #---------------------------------------------------

                self._add_overview(progress=progress, task_id=task_add_metadata)



                # --------------------------------------------------
                # Add Field Analysis sections

                self._add_metadata(progress=progress, task_id=task_add_metadata)



                # --------------------------------------------------
                # Add Plots

                self._add_plots(task_id=task_add_plots, progress=progress)



                # --------------------------------------------------
                # Setup Source Data table

                self._add_source_data(progress=progress, task_id=task_add_source_data)



                # --------------------------------------------------
                # Initialize the UI

                self._iniialize_ui(progress=progress, task_id=task_initialize_ui)

        
        # Print exit message/file path
        duration = time.time() - start_time

        console.print(f"\nxleda workbook created in {int(duration)} seconds \n", style=self.theme_style)
        
        if self.additional_plots:
            console.print("Additional plots included:", style=self.theme_style)
            for plot in self.additional_plots.keys():
                console.print(f"    {plot}", style=self.theme_style)
        
        
        console.print(f"\nlocation: \n    {self.wb_path} \n" + separator, style=self.theme_style)
        


    


    def export_analysis(self) -> dict[str, dict[str, list] | pd.DataFrame]:
        
        """
        Export an xleda field analysis

        Returns:

            Dictionary of exported items that includes:

            * `description`: Dataframe description if you've added one
            * `definitions`: Any field definitions you've added.
            * `notes`: Any field notes you've added
            * `lists`: Any lists showing in the compiled lists section
            * `source_data`: A copy of your unaltered source data that includes 
                             `Record Hash`/`Record List` columns.
            * `altered_source_data`: source data from the workbook that includes 
                                     any manual edits you've made such as removing 
                                     records, renaming fields, etc. *Note that 
                                     data types will likely change in the round-trip
                                       translation.* **

        """
        
        
        # --------------------------------------------------
        # Setup placeholder vars

        export_dict = {}
        definitions = {}
        notes = {}
        lists = {}


        # --------------------------------------------------
        # Configure output


        start_time = time.time()
        console.print(f"Process started at {time.strftime('%H:%M:%S')}", style=self.theme_style)
        

        # Configure progress bar table
        with Progress(TextColumn("{task.description}", style=self.theme_style),
                      BarColumn(pulse_style=self.silver_style, style=self.silver_style, complete_style=self.theme_style),
                      TextColumn("{task.percentage:>3.0f}%", style=self.theme_style),
                      "|",
                      TimeElapsedColumn(), # Displays elapsed time
                      console=console,
                      ) as progress:
            
            # Register tasks to get IDs, show progress bar
            export_analysis = progress.add_task("Exporting Analysis...", total=10)
            progress.update(export_analysis, completed=1, refresh=True)



            # --------------------------------------------------
            # Setup vars

            with xw.App(visible=False, add_book=False) as app:

                wb = app.books.open(self.wb_path)
                ws = wb.sheets("Field Analysis")


                progress.update(export_analysis, completed=5, refresh=True)



                # --------------------------------------------------
                # Export lists/notes/definitions from workbook


                # Definitions/Notes/Fields
                field_notes = ws.range("Notes").value
                field_definitions = ws.range("Definitions").value
                fields = ws.range("FieldRange").value

                # Compiled Lists
                compiled_lists_names = ws.range("Compiled_Lists").value
                compiled_lists = ws.range("Compiled_Lists").offset(0, 1).value
                
                progress.update(export_analysis, completed=6, refresh=True)



                # --------------------------------------------------
                # Add description to export_dict
                export_dict['description'] = ws.range("Description").value

                
                # --------------------------------------------------
                # Clean up, add definitions to export_dict
                for (i, definition) in enumerate(field_definitions):
                    if definition != "Definition":
                        definitions[fields[i]] = definition

                export_dict['definitions'] = definitions

                progress.update(export_analysis, completed=7, refresh=True)



                # --------------------------------------------------
                # Clean up/add notes to export_dict

                for (i, note) in enumerate(field_notes):
                    if note != "Notes":
                        notes[fields[i]] = note

                export_dict['notes'] = notes

                progress.update(export_analysis, completed=8, refresh=True)



                # --------------------------------------------------
                # Clean up/add lists to export_dict

                for (i, list_name) in enumerate(compiled_lists_names):
                    
                    if list_name:
                        lists[list_name[:-3]] = ast.literal_eval(compiled_lists[i])

                export_dict['lists'] = lists

                progress.update(export_analysis, completed=9, refresh=True)



                # --------------------------------------------------
                # Add data sets to export_dict
                
                # Source Data
                export_dict['source_data'] = self.source_df

                # Altered Source Data
                export_dict['altered_source_data'] = ws.tables['tbl_SourceData'].range.options(pd.DataFrame, index=False).value


                progress.update(export_analysis, completed=10, refresh=True)

        
        
        # --------------------------------------------------
        # Print output
        
        duration = time.time() - start_time
        console.print(separator + f"Export completed after {int(duration)} seconds", style=self.theme_style)


        return export_dict



    

    