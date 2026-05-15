# Imports, vars, config
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

import time





# Field Analysis template
template_file = Path(__file__).parent / "xleda_template.xlsm"


# Set matplotlib theme
mpl.use("Agg")
plt.style.use("dark_background")


class FieldAnalysis():
    def __init__(self, input_df: pd.DataFrame, name: str, theme_color: str = "#053476", large_report: bool = False, overwrite: bool = False, close_wb: bool = False,):

        self.source_df: pd.DataFrame = input_df.copy()
        self.name = name
        self.theme_color = theme_color
        self.large_report = large_report
        self.overwrite = overwrite
        self.close_wb = close_wb


        # Set Template Path
        self.field_analysis_path = (Path().cwd() / self.name).with_suffix(".xlsm")

        # Configure Source Data
        self._configure_source_data()

        # Create base analysis
        self.base_analysis = self._create_base_analysis()


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
        bars = ax.barh(y_pos, values, color=self.theme_color, height=0.5)  # noqa: F841

        
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
        Adds a small chart to an Excel cell.
        The small chart is 90% of the size of the cell and is centered.

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
        """Adds plots for each column of a pandas dataframe to an Excel range

        Args:
            input_wb (xw.Workbook): Target Workbook for plots
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

            progress.update(task_id, advance=1)

    def _configure_source_data(self):

        # --------------------------------------------------
        # Setup Source Data

        # Evaluate dimensions and subsample if necessary
        rows = len(self.source_df)
        columns = len(self.source_df.columns)

        above_default = rows > 100_000 or columns > 100
        above_limit = rows > 16_000 or columns > 1_000_000
        self.warning_msg = ""

        # Source data is below defaults or within Excel limits with large_report = True
        if not above_default or (self.large_report and not above_limit):
            self.warning = False
            self.warning_msg = ""

        # Source data is being sub sampled to fit below defaults
        elif above_default and not self.large_report:
            self.warning = True
            self.warning_msg = "This is only showing a sample because it is larger than the limits of 100,000 rows/100 columns"
            rows = min(rows, 100_000)
            columns = min(columns, 100)
            self.source_df = (self.source_df.iloc[:, :columns].sample(n=rows).sort_index())

        # Source data is being sub sampled to fit within Excel's limits
        elif above_limit:
            self.warning = True
            self.warning_msg = "This is only showing a sample because it is larger than the limits of 1,000,000 rows/16,000 columns"
            rows = min(rows, 1_000_000)
            columns = min(columns, 16_000)
            self.source_df = (self.source_df.iloc[:, :columns].sample(n=rows).sort_index())

        # Record adjusted rows/columns 
        self.rows = rows
        self.columns = columns

        # Add fields to Source Data to track individual records
        self.source_df.insert(loc=0, column="Mark For Removal", value="False")
        self.source_df["Record Hash"] = pd.util.hash_pandas_object(self.source_df, index=False)



    def create_workbook(self) -> None:
        """Creates a Field Analysis workbook from a given dataframe.  

           Workbook is saved in current directory

        """

        # --------------------------------------------------
        # Setup progress bars


        start_time = time.time()
        print(f"Process started at {time.strftime("%H:%M:%S")}")

        with Progress(
            TextColumn("[bold blue]{task.description}"),
            BarColumn(),
            TaskProgressColumn(),
            "•",
            TimeElapsedColumn(), # Displays elapsed time
            ) as progress:
            
            # 1. Register tasks to get IDs
            task_create_template = progress.add_task("[indigo]Creating Template...", total=10)
            task_add_metadata = progress.add_task("[indigo]Adding Metadata...", total=10)
            task_add_plots = progress.add_task("[indigo]Adding Plots...", total=self.columns)
            task_add_source_data = progress.add_task("[indigo]Adding Source Data...", total=10)
            task_initialize_ui = progress.add_task("[indigo]Initializing UI...", total=10)

        

            # --------------------------------------------------
            # Create Field Analysis Workbook

            progress.update(task_create_template, completed=1)
            
            # Handle existing files as appropriate
            if self.field_analysis_path.is_file() and not self.overwrite:
                print(f"""There is already a workbook named {self.field_analysis_path} 
                Use overwrite=True or rename/remove the existing workbook""")
                return None

            elif self.field_analysis_path.is_file() and self.overwrite:
                try:
                    self.field_analysis_path.unlink(missing_ok=True)

                except PermissionError:
                    print("Error: The workbook cannot be overwrittten while open.")
                    return None

            
            # Create a copy of the template and open it
            shutil.copy(template_file, self.field_analysis_path)
            
            progress.update(task_create_template, completed=5)


            # --------------------------------------------------
            # Initialize Template

            app = xw.App(visible=not self.close_wb, add_book=False)

            self.wb = app.books.open(self.field_analysis_path, read_only=False)

            field_analysis_ws = self.wb.sheets("Field Analysis")
            field_analysis_ws.range("FieldAnalysisTheme").color = self.theme_color
            field_analysis_ws.range("Dimensions").value = f"rows = {self.rows}, columns = {self.columns}"

            progress.update(task_create_template, completed=10)



            # --------------------------------------------------
            # Split base analysis into Field Analysis sections

            progress.update(task_add_metadata, completed=1)

            overview_df = self.base_analysis.loc[["Data type", "Distinct %", "Missing %", "Memory Usage %"]]
            composition_df = self.base_analysis.loc[["Memory Usage", "Distinct", "Count", "Missing"]]
            summary_stats_df = self.base_analysis.loc[["Mean", "Median", "Mode", "Standard Deviation", "Variance"]]
            percentiles_df = self.base_analysis.loc[["Min", "5%", "25%", "50%", "75%", "95%", "Max", "Range", "IQR"]]

            progress.update(task_add_metadata, completed=2)



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

            progress.update(task_add_metadata, completed=6)



            # --------------------------------------------------
            # Add Field Analysis sections

            field_analysis_ws.range("Overview")[0, 0].value = overview_df.values
            field_analysis_ws.range("Composition")[0, 0].value = composition_df.values
            field_analysis_ws.range("Summary_Stats")[0, 0].value = summary_stats_df.values
            field_analysis_ws.range("Percentiles")[0, 0].value = percentiles_df.values

            progress.update(task_add_metadata, completed=10)



            # --------------------------------------------------
            # Add Plots

            self._add_plots(task_id=task_add_plots, progress=progress)



            # --------------------------------------------------
            # Setup Source Data table

            progress.update(task_add_source_data, completed=1)

            # Update tbl_SourceData with source data, format "Mark For Removal" column
            source_table = field_analysis_ws.tables["tbl_SourceData"]
            source_table.update(self.source_df, index=False)
            source_table.data_body_range[0, 0].copy(destination=source_table.data_body_range[:, 0])  # type: ignore

            progress.update(task_add_source_data, completed=5)

            # Set "Record Hash" to LightHeader for contrast with source data fields
            last_header = field_analysis_ws.tables("tbl_SourceData").header_row_range.last_cell  # type: ignore
            field_analysis_ws.range("LightHeader").copy(destination=last_header)
            last_header.value = "Record Hash"
            last_header.columns.autofit()

            progress.update(task_add_source_data, completed=10)



            # --------------------------------------------------
            # Adjust Named Ranges for Field Action Formulas

            field_analysis_ws.range("tbl_SourceData[Record Hash]").name = "RecordHashes"
            field_analysis_ws.range("FieldRange").resize(row_size=1, column_size=len(headers) - 1).name = "FieldRange"
            field_analysis_ws.range("Notes").resize(row_size=1, column_size=len(headers) - 1).name = "Notes"

            for i in range(1, 7):
                excel_range = "FieldAction" + str(i)
                field_analysis_ws.range(excel_range).resize(row_size=1, column_size=len(headers) - 1).name = excel_range



            # --------------------------------------------------
            # Initialize Field Analysis UI

            progress.update(task_initialize_ui, completed=1)

            # Show/Hide Data Size Warning
            if self.warning:
                field_analysis_ws.range("Warning").value = self.warning_msg
                field_analysis_ws.range("Warning").api.EntireRow.Hidden = not self.warning

            progress.update(task_initialize_ui, completed=5)


            # Collapse subsections
            for excel_range in ["Composition", "Summary_Stats", "Percentiles", "Field_Actions", "Field_Action_Lists",]:
                field_analysis_ws.range(excel_range).api.EntireRow.Hidden = True


            progress.update(task_initialize_ui, completed=7)

            # --------------------------------------------------
            # Save workbook, close if required, and return it

            self.wb.save(self.field_analysis_path)

            if self.close_wb:
                self.wb.close()
                app.quit()


            progress.update(task_initialize_ui, completed=10)

        
        # Print exit message/file path
        duration = time.time() - start_time

        print(f"Process completed after {int(duration)} seconds")

        print(f"{self.field_analysis_path}")



    def export_analysis(self, close_wb: bool =False) -> dict[str, list[str]]:
        
        """Exports notes and action lists from an xleda field analysis workbook

        Args:
            close_wb: Leave the analysis workbook open after import or not.  Default is False

        Returns:
            dict[str, list[str]]: Dictionary of exported notes and action lists

        """


        # --------------------------------------------------
        # Setup vars


        # app = xw.App(visible=not close_wb, add_book=False)
        wb = xw.books.open(self.field_analysis_path)

        target_ws = wb.sheets("Field Analysis")

        export_dict = {}
        notes_dict = {}


        # --------------------------------------------------
        # Export Notes/action lists from workbook

        field_notes = target_ws.range("Notes").value
        fields = target_ws.range("FieldRange").value

        action_names = target_ws.range("Field_Action_Lists").value
        action_lists = target_ws.range("ActionLists").value[:-1]
        
        
    
        # --------------------------------------------------
        # Add Notes to export_dict

        for (i, note) in enumerate(field_notes):
            if note != "Notes":
                notes_dict[fields[i]] = note

        export_dict['notes'] = notes_dict


        # --------------------------------------------------
        # Add action_lists to export_dict

        for (i, action) in enumerate(action_names):
            
            if action:
                action_name = action[:-3]
                export_dict[action_name] = ast.literal_eval(action_lists[i])


        return export_dict
