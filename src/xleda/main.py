import seaborn as sns
import pandas as pd
import shutil
from pathlib import Path
import ast

import matplotlib.pyplot as plt
import matplotlib as mpl
from matplotlib.figure import Figure

import xlwings as xw
from xlwings import Range

from rich.progress import (
    Progress, BarColumn, TextColumn, 
    TimeElapsedColumn, TaskProgressColumn,
    TaskID)
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
                 theme_color: str = "#053476", 
                 large_report: bool = False, 
                 overwrite: bool = False, 
                 close_wb: bool = False,):

        """Configures an xleda workbook

        Args:

            name (str): Name of the workbook to be created.

            theme_color (str): A hexidecimal color used for charts/accent color.  Dark colors work better.

            large_report (bool): Used to override default limits of 100,000 rows/100 columns. Sets limits to 1,000,000 rows/16,000 columns. 

            overwrite (bool): Whether to overwrite existing reports with the same name.

            close_wb (bool): Whether to close the workbook after it has been created.
        
        """

        self.name = name
        self.theme_color = theme_color
        self.large_report = large_report
        self.overwrite = overwrite
        self.close_wb = close_wb
        self.theme_style = Style(color=self.theme_color)

        # Set Template Path
        self.field_analysis_path = (Path().cwd() / self.name).with_suffix(".xlsm")
        
        # Configure Source Data
        self.source_df: pd.DataFrame = self._configure_source_data(input_df.copy())

        # Create base analysis
        self.base_analysis: pd.DataFrame = self._create_base_analysis()



    def _create_base_analysis(self) -> pd.DataFrame:
        """Produces the base analysis dataframe for an input dataframe"""


        # --------------------------------------------------
        # Collect metadata

        df = self.source_df.iloc[:, 1:-1]

        # Order of output fields
        col_order = ["Data type", "Memory Usage", "Memory Usage %", "Distinct", "Distinct %", "Count", "Count %", "Missing", "Missing %", "Mean", "Median", "Mode", "Standard Deviation", "Variance", "Min", "5%", "25%", "50%", "75%", "95%", "Max", "Range", "IQR", ]

        # Get statistical summary
        rows_count = len(df)
        desc = df.describe(include="all", percentiles=[0.05, 0.25, 0.5, 0.75, 0.95])

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

        # Rename index fields, reorder, and filter return rows
        summary_df = summary_df.rename(index=field_map)
        summary_df = summary_df.loc[col_order]

        return summary_df



    def _create_composition_table(self, input_df: pd.DataFrame) -> Figure:
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



    def _create_histogram(self, input_df: pd.DataFrame) -> Figure:

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
        sns.histplot( data=input_df, x=input_df.columns[0], color=self.theme_color, stat="density", alpha=0.5, ax=ax, )

        # Layer the KDE line
        sns.kdeplot(data=input_df, x=input_df.columns[0], color="silver", linewidth=3, ax=ax)

        
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



    def _add_small_multiple(self, fig: Figure, target_range: Range):
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



    def _add_plots(self, progress: Progress, task_id: TaskID):
        
        """Adds plots to an Excel range

        Args:
            input_wb (xw.Workbook): Target Workbook for plots
            progress (Progress): A rich Progress object
            task_id: A rich TaskID object

        """

        # --------------------------------------------------
        # Setup workbook objects

        
        

        field_analysis_ws = self.wb.sheets("Field Analysis")

        # Set initial ranges for added plots
        histogram_range = field_analysis_ws.range("Histogram")
        composition_range = field_analysis_ws.range("CompositionTable")

        for col in self.source_df.iloc[:, 1:-1].columns:


            # --------------------------------------------------
            # Add Composition Multiple

            composition_table = self._create_composition_table(self.source_df[[col]])
            
            self._add_small_multiple(target_range=composition_range, fig=composition_table)

            if pd.api.types.is_numeric_dtype(self.source_df[col]):



                # --------------------------------------------------
                # Add Histogram Multiple

                histogram = self._create_histogram(self.source_df[[col]])
                self._add_small_multiple(target_range=histogram_range, fig=histogram)



            # --------------------------------------------------
            # Increment Target Ranges

            histogram_range = histogram_range.offset(0, 1)
            composition_range = composition_range.offset(0, 1)

            progress.update(task_id, advance=1, refresh=True)



    def _configure_source_data(self, input_df: pd.DataFrame) -> pd.DataFrame:
        
        """
        Configures source data for analysis and subsamples it if necessary


        Returns:
            source_df: A pandas dataframe object
        
        """


        # --------------------------------------------------
        # Setup Source Data


        # Initialize Warning Vars
        self.warning_msg = ""
        self.warning = False

        # Evaluate dimensions and subsample if necessary
        rows = len(input_df)
        columns = len(input_df.columns)

        above_default = rows > 100_000 or columns > 100
        above_limit = rows > 16_000 or columns > 1_000_000
          

        # Source data is being sub sampled to fit below defaults
        if above_default and not self.large_report and not above_limit:
            self.warning = True
            self.warning_msg = "This is only showing a sample because it is larger than the limits of 100,000 rows/100 columns"
            rows = min(rows, 100_000)
            columns = min(columns, 100)
            input_df = (input_df.iloc[:, :columns].sample(n=rows).sort_index())

        # Source data is being sub sampled to fit within Excel's limits
        elif above_limit:
            self.warning = True
            self.warning_msg = "This is only showing a sample because it is larger than the limits of 1,000,000 rows/16,000 columns"
            rows = min(rows, 1_000_000)
            columns = min(columns, 16_000)
            self.source_df = (input_df.iloc[:, :columns].sample(n=rows).sort_index())

        # Record adjusted rows/columns 
        self.rows = rows
        self.columns = columns

        # Add fields to Source Data to track individual records
        input_df.insert(loc=0, column="Record List", value="False")
        input_df["Record Hash"] = pd.util.hash_pandas_object(input_df, index=False)

        return input_df



    def create_workbook(self) -> xw.Book | None:
        """
        Creates an xdleda workbook from a given dataframe.  

        Workbook is saved in current directory
        
        Returns:
            wb: An xlwings Book object
        
        """

        # --------------------------------------------------
        # Setup progress bars


        start_time = time.time()

        
        silver_style = Style(color='#C0C0C0')
      

        console.print(separator + f"\n Process started at {time.strftime("%H:%M:%S")}", style=self.theme_style)
        console.print(f"\n Preparing an xleda workbook with {self.name} data \n", style=self.theme_style)
        

        with Progress(
            TextColumn("{task.description}", style=self.theme_style),
            BarColumn(pulse_style=silver_style, style=silver_style, complete_style=self.theme_style),
            TextColumn("{task.percentage:>3.0f}%", style=self.theme_style),
            "•",
            TimeElapsedColumn(), # Displays elapsed time
            console=console,
            ) as progress:
            

            
            # 1. Register tasks to get IDs
            task_create_template = progress.add_task("[self.theme_color]Creating Template...", total=10)
            task_add_metadata = progress.add_task("[self.theme_color]Adding Metadata...", total=10)
            task_add_plots = progress.add_task("[self.theme_color]Adding Plots...", total=self.columns)
            task_add_source_data = progress.add_task("[self.theme_color]Adding Source Data...", total=10)
            task_initialize_ui = progress.add_task("[self.theme_color]Initializing UI...", total=10)

        

            # --------------------------------------------------
            # Create Field Analysis Workbook

            progress.update(task_create_template, completed=1, refresh=True)
            
            # Handle existing files as appropriate
            if self.field_analysis_path.is_file() and not self.overwrite:
                
                console.print(f"[bold red]Error: There is already a workbook named {self.field_analysis_path}![/bold red]")
                console.print("Use overwrite=True or rename/remove the existing workbook")

                return None

            elif self.field_analysis_path.is_file() and self.overwrite:
                try:
                    self.field_analysis_path.unlink(missing_ok=True)

                except PermissionError:
                    console.print("[bold red]Error: The workbook cannot be overwritten while open![/bold red]")
                    return None

            
            # Create a copy of the template and open it
            shutil.copy(template_file, self.field_analysis_path)
            
            progress.update(task_create_template, completed=5, refresh=True)


            # --------------------------------------------------
            # Initialize Template

            with xw.App(visible=False, add_book=False) as app:

                self.wb = app.books.open(self.field_analysis_path, read_only=False)

                field_analysis_ws = self.wb.sheets("Field Analysis")
                field_analysis_ws.range("FieldAnalysisTheme").color = self.theme_color
                field_analysis_ws.range("Dimensions").value = f"rows = {self.rows}, columns = {self.columns}"

                progress.update(task_create_template, completed=10, refresh=True)



                # --------------------------------------------------
                # Split base analysis into Field Analysis sections

                progress.update(task_add_metadata, completed=1, refresh=True)

                overview_df = self.base_analysis.loc[["Data type", "Distinct %", "Missing %", "Memory Usage %"]]
                composition_df = self.base_analysis.loc[["Memory Usage", "Distinct", "Count", "Missing"]]
                summary_stats_df = self.base_analysis.loc[["Mean", "Median", "Mode", "Standard Deviation", "Variance"]]
                percentiles_df = self.base_analysis.loc[["Min", "5%", "25%", "50%", "75%", "95%", "Max", "Range", "IQR"]]

                progress.update(task_add_metadata, completed=2, refresh=True)



                # --------------------------------------------------
                # Format Field Analysis sections

                # Format placeholders
                format_from = field_analysis_ws.range("FormatRange")
                format_to = (field_analysis_ws.range("FormatRange").offset(0, 1).resize(None, self.columns-2))
                format_from.api.Copy()
                format_to.api.Select()
                field_analysis_ws.api.Paste()

                # Clear clipboard, set selection to top, and add header values
                self.wb.app.api.CutCopyMode = False
                field_analysis_ws.range("Headers_Start").api.Select()
                headers = self.source_df.columns.to_list()[1:]
                field_analysis_ws.range("Headers_Start").value = headers

                progress.update(task_add_metadata, completed=6, refresh=True)



                # --------------------------------------------------
                # Add Field Analysis sections

                field_analysis_ws.range("Overview")[0, 0].value = overview_df.values
                field_analysis_ws.range("Composition")[0, 0].value = composition_df.values
                field_analysis_ws.range("Summary_Stats")[0, 0].value = summary_stats_df.values
                field_analysis_ws.range("Percentiles")[0, 0].value = percentiles_df.values

                progress.update(task_add_metadata, completed=10, refresh=True)



                # --------------------------------------------------
                # Add Plots

                self._add_plots(task_id=task_add_plots, progress=progress)



                # --------------------------------------------------
                # Setup Source Data table

                progress.update(task_add_source_data, completed=1, refresh=True)

                # Update tbl_SourceData with source data, format "Mark For Removal" column
                source_table = field_analysis_ws.tables["tbl_SourceData"]
                source_table.update(self.source_df, index=False)
                source_table.data_body_range[0, 0].copy(destination=source_table.data_body_range[:, 0])  # type: ignore

                progress.update(task_add_source_data, completed=5, refresh=True)

                # Set "Record Hash" to LightHeader for contrast with source data fields
                last_header = field_analysis_ws.tables("tbl_SourceData").header_row_range.last_cell  # type: ignore
                field_analysis_ws.range("LightHeader").copy(destination=last_header)
                last_header.value = "Record Hash"
                last_header.columns.autofit()

                progress.update(task_add_source_data, completed=10, refresh=True)



                # --------------------------------------------------
                # Adjust Named Ranges for Field Action Formulas

                field_analysis_ws.range("tbl_SourceData[Record Hash]").name = "RecordHashes"
                field_analysis_ws.range("FieldRange").resize(row_size=1, column_size=len(headers) - 1).name = "FieldRange"
                field_analysis_ws.range("Notes").resize(row_size=1, column_size=len(headers) - 1).name = "Notes"

                for i in range(1, 7):
                    excel_range = "FieldList" + str(i)
                    field_analysis_ws.range(excel_range).resize(row_size=1, column_size=len(headers) - 1).name = excel_range



                # --------------------------------------------------
                # Initialize Field Analysis UI

                progress.update(task_initialize_ui, completed=1, refresh=True)

                # Show/Hide Data Size Warning
                if self.warning:
                    field_analysis_ws.range("Warning").value = self.warning_msg
                    field_analysis_ws.range("Warning").api.EntireRow.Hidden = not self.warning

                progress.update(task_initialize_ui, completed=5, refresh=True)


                # Collapse subsections
                for excel_range in ["Composition", "Summary_Stats", "Percentiles", "Field_Lists", "Compiled_Lists",]:
                    field_analysis_ws.range(excel_range).api.EntireRow.Hidden = True


                progress.update(task_initialize_ui, completed=7, refresh=True)

                # --------------------------------------------------
                # Save workbook, close if required, and return it

                self.wb.save(self.field_analysis_path)

                progress.update(task_initialize_ui, completed=10, refresh=True)

        
        # Print exit message/file path
        duration = time.time() - start_time

        console.print(f"\n xleda workbook created in {int(duration)} seconds \n \n", style=self.theme_style)
        console.print(f"location: \n {self.field_analysis_path} \n" + separator, style=self.theme_style)
        

        return self.wb
    


    def export_analysis(self) -> dict[str, dict[str, list] | pd.DataFrame]:
        
        """Export notes, lists, data from an xleda field analysis workbook

        Returns:

            export_dict: Dictionary of exported lists, notes, and data.

            export_dict['notes']: Any lists showing in the compiled lists section
            
            export_dict['lists']: Any field notes you've added

            export_dict['source_data']: Unaltered source data with `Record Hash`/`Record List` columns added.

            export_dict['altered_source_data']: Source data from the workbook that includes any edits made such 
                as removing records, renaming fields, etc.  Note that data types will likely change in the 
                round-trip translation.

        """


        # --------------------------------------------------
        # Configure output


        start_time = time.time()


        console.print(f"Process started at {time.strftime("%H:%M:%S")}", style=self.theme_style)
        


        with Progress(
            TextColumn("[bold blue]{task.description}"),
            BarColumn(),
            TaskProgressColumn(),
            "•",
            TimeElapsedColumn(), # Displays elapsed time
            ) as progress:
            
            # 1. Register tasks to get IDs
            export_analysis = progress.add_task("[indigo]Exporting Analysis...", total=10)

            progress.update(export_analysis, completed=1, refresh=True)



            # --------------------------------------------------
            # Setup vars


            app = xw.App(visible=False) 
            wb = app.books.open(self.field_analysis_path)
            target_ws = wb.sheets("Field Analysis")

            export_dict = {}
            notes = {}
            lists = {}


            progress.update(export_analysis, completed=5, refresh=True)



            # --------------------------------------------------
            # Export lists/notes from workbook


            # Notes/Fields
            field_notes = target_ws.range("Notes").value
            fields = target_ws.range("FieldRange").value

            # Compiled Lists
            compiled_lists_names = target_ws.range("Compiled_Lists").value
            compiled_lists = target_ws.range("Compiled_Lists").offset(0, 1).value
            
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
            # Add Data sets to export_dict
            
            # Source Data
            export_dict['source_data'] = self.source_df


            # Altered Source Data
            export_dict['altered_source_data'] = target_ws.tables['tbl_SourceData'].range.options(pd.DataFrame, index=False).value



            # Close app
            app.quit()
            
            progress.update(export_analysis, completed=10, refresh=True)


        # Print outputs
        
        duration = time.time() - start_time

        console.print(separator + f"Export completed after {int(duration)} seconds", style=self.theme_style)

        return export_dict
