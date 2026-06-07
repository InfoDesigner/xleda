import seaborn as sns
import random
import colorsys
import re
import time
import subprocess
from xlwings.constants import Constants

from collections import Counter
from pathlib import Path
import platform
import shutil
import sys
import datetime
import os
import pandas as pd
from typing import Any
import send2trash
from collections import defaultdict

import matplotlib.pyplot as plt
import matplotlib as mpl 
from matplotlib.figure import Figure


from tqdm.auto import tqdm
import xlwings as xw


# Set matplotlib theme
mpl.use("Agg")
plt.style.use("dark_background")

default_row_limit = 25_000
default_column_limit = 50
upper_row_limit = 1_000_000
upper_column_limit = 16_000

xlsm_file = Path(__file__).parent / "xleda_template.xlsm"
xlsx_file = Path(__file__).parent / "xleda_template.xlsx"



separator = "\n" + ("-" * 100)



class Blueprint():

    """ 
    Class that represents an xleda blueprint"""

    def __init__(self, 
                 name: str,
                 title: str,
                 input_df: pd.DataFrame, 
                 large_report: bool,
                 overview:str,
                 field_analysis: str,
                 pivot:str = "") -> None:

        self.name = name
        self.title = title
        self.input_df = input_df
        self.large_report = large_report
        self.overview=overview
        self.field_analysis= field_analysis
        self.pivot=pivot
        
        # Add placeholder properties
        self.warning_msg: str = ""
        self.warning: bool = False
        self.rows = len(input_df)
        self.columns = len(input_df.columns)

        # Configure Source Data
        self.source_data: pd.DataFrame = self.configure_source_data()


        # add field_metadata/overview_metadata
        self.field_metadata: pd.DataFrame = self.create_field_metadata()
        self.df_metadata: dict[str, Any] = self.create_df_metadata()
        self.overview_metadata = self.field_metadata.T


        # Add export placeholders
        self.description = ""
        self.definitions: dict = {}
        self.notes: dict = {}
        self.lists: dict = {}
        self.altered_source_data: pd.DataFrame = pd.DataFrame()


        # Initialize performance logging
        self.perf_pivot: dict[str, Any] = {}
        self.perf_plots: dict[str, Any] = {}


    def configure_source_data(self) -> pd.DataFrame:
        
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
            self.warning_msg = ("This is only showing a sample because it is larger than the default "
                                "limits of 25,000 rows/50 columns.  See documentation for details.")
            rows = min(rows, default_row_limit)
            columns = min(columns, default_column_limit)
            

        # Source data is larger than Excel's limits
        elif above_limit:

            self.warning = True
            self.warning_msg = ("This is only showing a sample because it is larger than Excel's limits "
                                "of 1,000,000 rows/16,000 columns.  See documentation for details.")
            
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



    def create_field_metadata(self) -> pd.DataFrame:

        """
        Produces a field_metadata dataframe

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

        
        # Calculate variance without failing on unsupported types such as timedelta
        try:
            variance = desc.loc["std"] ** 2
        except TypeError:
            variance = None
        
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
                "Variance": variance,
                "Distinct": df.nunique(),
                "Distinct %": df.nunique() / rows_count,
                }).T


        # --------------------------------------------------
        # Combine/Format Metadata

        
        # Combine info and describe dfs
        summary_df = pd.concat([info_df, desc])

        # Field name map
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



    def create_df_metadata(self) -> dict[str, Any]:

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
        


class Environment():
    

    def __init__(self, debug: bool = False) -> None:
        
        """
        Gathers os, python, and terminal details for debugging.

        """    
        
        # Primary environment detail
        
        self.os = platform.system()
        self.win = self.os == 'Windows'
        self.mac = self.os == 'Darwin'
        self.env_type = self.get_env_type()
        self.excel_version = self.get_excel_version()
        self.debug = debug


        self.validate_compatibility()
        

        # Determine file recovery tool
        if self.win:
            self.junk_drawer = 'Recycle Bin'
        elif self.mac:
            self.junk_drawer = 'Trash'
        else:
            self.junk_drawer = 'None'


        # Additional environment details
        terminal_size = shutil.get_terminal_size()
        self.date = datetime.datetime.now().date().strftime("%#m/%#d/%Y")
        self.os_release = platform.release()
        self.os_version = platform.version()
        self.architecture = platform.machine()
        self.processor = platform.processor()
        self.python_version = platform.python_version()
        self.python_implementation = platform.python_implementation()
        self.console_columns = terminal_size.columns
        self.console_lines = terminal_size.lines


    def get_excel_version(self) -> str:

        """
        Uses xlwings to obtain the Excel version

        Returns
        -------
        str
            The version of Excel discovered

        """

        app_version = ""
        
        # Creates an xw.App object and pulls the version details

        try:
            with xw.App() as app:
                app_version = str(app.version)
        except Exception:
            return ""
        
        return app_version



    def get_env_type(self) -> str :
        
        """
        Determines whether the program is running in:
            a notebook, vs code notebook, a terminal, or an IDE

        Returns
        -------
        str
            The type of environment being used
        """

        # Check for Notebook (Jupyter/Colab)
        if 'ipykernel' in sys.modules or 'JPY_PARENT_PID' in os.environ:
            return 'Notebook'
        
        # Check for IDE or terminal interactivity
        if "__vsc_ipynb_file__" in globals():
            return 'VS Code Notebook'
        
        if sys.stdin is None or not sys.stdin.isatty():
            return 'IDE Non Interactive'
        
        # Raw Terminal vs IDE terminal
        if 'TERM_PROGRAM' in os.environ or 'TERMINAL_EMULATOR' in os.environ:
            
            # Common in VS Code, PyCharm, or Terminal tabs
            return 'Terminal or IDE'
        
        return 'Raw Terminal (Standard Python Shell)'



    def validate_compatibility(self):

        """
        Validates that supported OS/Office versions are being used

        """

        self.compatible = True

        # Determine if a supported OS is being used
        if self.mac or self.win:
            os_compatible = "Compatible"
            self.compatible = True
        else:
            self.compatible = False
            os_compatible = "Incompatible"

        # Determine if a supported version of Excel is being used
        if not self.excel_version:
            self.excel_version = "Not Detected"
        elif float(self.excel_version) >= 16:
            excel_compatibility = "Compatible"
        else:
            self.compatible = False
            excel_compatibility = "Incompatible"

        
        if not self.compatible:

            compatibility_msg = (
                separator + "\nxleda requires a full desktop version of Microsoft Excel\n"
                "\nIt has been developed and tested on Windows\n"
                "\nIt should also work on MacOS though this has not yet been tested\n\n"
                f"{'Requires MacOS/Windows':<25} | Detected {self.os:<20} | {os_compatible:<20}\n"
                f"{'Requires Excel >=16.0':<25} | Detected {self.excel_version:<20} | {excel_compatibility:<20}\n"
                + separator)

            print(compatibility_msg)

            sys.exit()



class Theme():

    """
    CLass that represents an xleda theme
    
    """

    def __init__(self, theme_color: str, env: Environment) -> None:

        """
        Configures an xleda theme that includes workbook 
            theme, progress bars, and console printing

        """
    
    
        # ------------------------------------------------------------------------------
        # Configure Theme
        
        if theme_color == 'random':
            self.theme_color = "#{:06x}".format(random.randint(0, 0xFFFFFF))
        else:
            self.theme_color: str = theme_color[:7]
            
        self.black_text: bool = self.use_black_text(self.theme_color)
        self.print_theme: str = self.ensure_readable(self.theme_color[:7])
        self.env = env


    def create_progress_bar(self, desc: str, total: float) -> tqdm:
        
        """
        Creates a tqdm progress bar

        Returns
        -------
        tqdm
            A tqdm object
        """
        

        fmt = "{desc} | {percentage:3.0f}% | {bar} | {elapsed}"


        # Pad the raw desc for alignment
        padded_desc = f"{desc:<30}"

        # Create a tqdm instance
        pbar = tqdm(
            total=total,
            desc=padded_desc,
            bar_format=fmt,
            colour=self.print_theme,
            # ncols=100,
            # dynamic_ncols=True
        )
        
        return pbar



    def greyscale_color(self, iteration: int, windows=True):

        """
        Increments the brightness of a color by 10%

        Returns
        -------
        
        An Excel Index color on Windows or an RGB tuple for MacOS
            
        """

        iteration = iteration % 9
        
        if windows:
            return (iteration*26) + (iteration*26*256)  + (iteration*26*256*256)
        else:
            return ((iteration*26), (iteration*26), (iteration*26))
        


    def print(self, text: str):
        
        """
        Prints to the console using theme_color text. 

        Parameters
        ----------
        text : str
            Text to be printed

        """
        
        print(self.color_formatter(text=text, theme=self.print_theme))

        

    def hex_to_ansi(self, hex_color):
        
        """ 
        Converts a hex color to ansi code for colored terminal text
        
        """

        # Remove '#' and convert hex to RGB integers
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

        ansi_color = f"\033[38;2;{r};{g};{b}m"
        return ansi_color



    def color_formatter(self, text: str, theme: str):
        
        """
        Wraps text in ansi colored code start/stop statements

        Returns
        -------
        str
            
            A terminated ansi colored text string

        """

        color = self.hex_to_ansi(theme)
        reset = "\033[0m"  # Crucial: Resets terminal to default style
        
        return (f"{color}{text}{reset}")



    def warn_print(self, text: str):
        
        """
        Prints text in red bold for warning messages

        """
        
        print(f"\033[1;31m{text}\033[0m")



    def hex_to_rgb(self, hex_str):

        """
        Converts #RRGGBB to (R, G, B) normalized to 0-1.
        
        """

        hex_str = hex_str.lstrip('#')

        return tuple(int(hex_str[i:i+2], 16) / 255.0 for i in (0, 2, 4))



    def rgb_to_hex(self, rgb):

        """
        Converts normalized (R, G, B) back to #RRGGBB.
        
        """

        return '#' + ''.join(f'{int(round(c * 255)):02x}' for c in rgb)



    def get_luminance(self, rgb):

        """
        Calculates relative luminance for WCAG contrast standards.
        
        """

        res = []
        for c in rgb:
            # Standard linearization of sRGB
            res.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
        return 0.2126 * res[0] + 0.7152 * res[1] + 0.0722 * res[2]



    def get_contrast(self, rgb1, rgb2):

        """
        Calculates the contrast ratio between two colors.
        
        """

        l1, l2 = self.get_luminance(rgb1), self.get_luminance(rgb2)
        return (max(l1, l2) + 0.05) / (min(l1, l2) + 0.05)



    def ensure_readable(self, hex_color, target_ratio=4.5):

        """
        Lightens a color until it reaches the target contrast on black.
        
        """

        rgb = self.hex_to_rgb(hex_color)
        black = (0, 0, 0)
        
        if self.get_contrast(rgb, black) >= target_ratio:
            return hex_color
        
        # Convert to HSL to adjust lightness (l) while keeping hue (h) and saturation (s)
        hue, luminance, saturation = colorsys.rgb_to_hls(*rgb)
        
        # Binary search for the minimum lightness adjustment
        low, high = luminance, 1.0
        for _ in range(20):
            mid = (low + high) / 2
            if self.get_contrast(colorsys.hls_to_rgb(hue, mid, saturation), black) >= target_ratio:
                high = mid
            else:
                low = mid
                
        return self.rgb_to_hex(colorsys.hls_to_rgb(hue, high, saturation))


    def use_black_text(self, color: str) -> bool:

        """
        Converts theme color to RGB and calculates 
            whether black text is required.
        
        """

        # Remove '#' if present and convert hex to RGB
        hex_color = color.lstrip('#')

        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        # Standard formula for perceived brightness
        brightness = (r * 0.299 + g * 0.587 + b * 0.114) / 255
        
        # Use black text for light backgrounds, white for dark ones

        return brightness > 0.5



    def set_theme(self, input_range: xw.Range):

        """
        Sets the background/font colors of a range object to the current xleda theme.

        """

        input_range.color = self.theme_color

        if self.black_text:
            input_range.font.color = '#000000'
       


    def greyscale_range(self, input_range: xw.Range):
        
        """
        Formats a range as grey on grey

        """

        input_range.color = '#262626'
        input_range.font.color = '#898989'



    def greyscale_tabs(self, bp: Blueprint, iteration:int, book: xw.Book):

        """
        Colors worksheet tabs to a shade of grey for contrast with adjacent worksheets
            
        """
        

        
        wb = book

        # 52*5 > 255 limit for RGB so limit to 4
        iteration = (iteration + 1) % 4
        multiplier = 52 * iteration


        # Set color for each dataframe's set of worksheets
        for sheet in [bp.overview, bp.field_analysis]:
            if self.env.win:
                color = (multiplier) + (multiplier*256)  + (multiplier*256*256)
                wb.sheets(sheet).api.Tab.Color = color
            
            elif self.env.mac:
                color = ((multiplier), (multiplier), (multiplier))
                wb.sheets(sheet).api.tab_color.set(color) 
        


class Plotter():
    
    """
    Class that represents a xleda plotting object

    """
    
    def __init__(self, theme: Theme, env: Environment) -> None:
        
        """
        Creates theme appropriate plots and optinally writes them to a range

        """
        
        self.theme_color = theme.theme_color
        self.env = env

    # TODO: Test this
    def add_small_plot(self, fig: Figure, target_range: xw.Range, name: str):

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


        if self.env.win:

            pic = target_range.sheet.pictures.add(fig,
                                                  name=name,
                                                  left=target_left,
                                                  top=target_top,
                                                  width=target_width,
                                                  height=target_height)

            # Set placement to xlMoveAndSize
            try:
                pic.api.Placement = 1
            except AttributeError:
                pass
        
        elif self.env.mac:
            
            # Uses anchor instead
            pic = target_range.sheet.pictures.add(fig,
                                                  name=name,
                                                  anchor=target_range,
                                                  width=target_width,
                                                  height=target_height)


    def create_composition_plot(self, input_df: pd.Series) -> Figure:
        
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
        ax.barh(y_pos,
                values,
                color=self.theme_color, 
                height=0.5, 
                edgecolor='silver')

        
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
            ax.text(-0.55, 
                    y, 
                    display_cat, 
                    color="white", 
                    va="center", 
                    ha="left", 
                    fontsize=font_size, 
                    transform=ax.get_yaxis_transform(), 
                    )

            # Add Percentages
            ax.text(-0.05, 
                    y, 
                    f"{pct:.0f}%", 
                    color="white", 
                    va="center", 
                    ha="right", 
                    fontsize=font_size, 
                    transform=ax.get_yaxis_transform(), 
                    )

            # Add Counts to the right of the bars
            ax.text(val + max_val * 0.02, 
                    y, 
                    str(val), 
                    color="white",
                    va="center", 
                    ha="left", 
                    fontsize=font_size, 
                    )

        return fig



    def create_histogram_plot(self, input_series: pd.Series) -> Figure:

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
        sns.histplot(x=input_series, 
                     color=self.theme_color, 
                     stat="density", 
                     alpha=0.5, 
                     ax=ax)

        # Layer the KDE line
        sns.kdeplot(x=input_series, 
                    color="silver", 
                    linewidth=3, 
                    ax=ax, 
                    warn_singular=False
                    )

        
        # --------------------------------------------------
        # Add additional plot details

        # Add vertical mean line
        mean_val = input_series.mean()
        ax.axvline(mean_val, 
                   color="silver", 
                   linestyle=":", 
                   linewidth=2)

        # Remove tick labels
        ax.tick_params(left=False, 
                       bottom=False, 
                       labelleft=False, 
                       labelbottom=False)

        # Add Min and Max text at the bottom corners
        min_val = input_series.min()
        max_val = input_series.max()

        ax.text(0, -0.05, f"Min {min_val:g}", transform=ax.transAxes, fontsize=16, 
                color="silver", ha="left", va="top",) 
        
        ax.text( 1, -0.05, f"Max {max_val:g}", transform=ax.transAxes, fontsize=16, 
                color="silver", ha="right", va="top", )

        return fig


    


class ExportDict(dict):

    """ 
    Class that represents an xleda export dictionary
    
    """

    def __init__(self, bp: Blueprint) -> None:
        
        """
        Generates an ExportDict object that provides access to xleda 
            metadata through both key/val and dot notation

        Parameters
        ----------
        bp : Blueprint
            An xleda blueprint

        """


        for key in ['description', 'definitions', 'notes', 'lists', 'field_metadata', 'overview_metadata', 'altered_source_data']:
            
            val = eval(f"bp.{key}")

            if isinstance(val, pd.DataFrame) or val:
                setattr(self, key, eval(f"bp.{key}"))
                self[key] = eval(f"bp.{key}")
    


class Config():

    """
    Class that represents an xleda configuration

    """

    def __init__(self, 
                 wb_path: str | Path,
                 input_df: pd.DataFrame, 
                 theme: Theme,
                 env: Environment,
                 no_vba: bool, 
                 large_report: bool,
                 name: str,
                 add_plots: dict[str, Figure],
                 add_dfs: dict[str, pd.DataFrame],
                 overwrite: bool,
                 debug: bool,
                 open_wb: bool,
                 ) -> None:
        
        """
        Primary configuration object for an xleda workbook

        Parameters
        ----------
        input_path : str | Path
            The wb_path provided or default

        input_df : pd.DataFrame
            The primary dataframe provided or default

        theme : Theme
            The theme provided or default

        env : Environment
            The operating environment in Environment form

        input_no_vba : bool, optional
            The no_vba flag provided or default
            
        large_report : bool, optional
            The large_report flag provided or default

        name : str, optional
            The name provided or default

        add_plots : dict[str, Figure], optional
            The add_plots provided or default

        add_dfs : dict[str, pd.DataFrame], optional
            add_dfs provided or default

        overwrite : bool, optional
            The overwrite flag provided or default

        debug : bool, optional
            The debug flag provided or default

        open_wb : bool, optional
            The open_wb flag provided or default

        """
        

        self.name = name
        self.title = self.sanitize_name(name, file_name=True)
        self.large_report = large_report
        self.theme: Theme = theme
        self.env = env
        self.no_vba = no_vba
        self.wb_path = wb_path
        self.overwrite: bool = overwrite
        self.debug: bool = debug
        self.open_wb: bool = open_wb
        self.input_df = input_df.copy()
        self.exit_msg = separator


        # Calculate the target file path        
        self.path = self.calculate_full_path()

        # Set placeholder variables
        self.additional_plots: list[dict[str, Any]] = []
        self.blueprints: list[Blueprint] = []
        
        
        # Assemble blueprints for dataframes, name objects, configure plots to add
        self.allocate_components(plots_to_add=add_plots,
                                 dfs_to_add=add_dfs)



    def allocate_components(self,
                            plots_to_add: dict[str, Figure], 
                            dfs_to_add: dict[str, pd.DataFrame]):

        """ 
        Allocates xleda components.

        """        


        # -----------------------------------------------------------------
        # Collect all titles/unique names together

        titles = [self.sanitize_name(self.name, restricted_name=False)]
        unique_names = [self.sanitize_name(self.name)]
        
        if dfs_to_add:
            titles += [self.sanitize_name(name, restricted_name=False) for name in list(dfs_to_add.keys())]
            unique_names += [self.sanitize_name(name) for name in list(dfs_to_add.keys())]
        
        if plots_to_add:
            titles += [self.sanitize_name(name, restricted_name=False) for name in list(plots_to_add.keys())]
            unique_names += [self.sanitize_name(name) for name in list(plots_to_add.keys())]



        # -----------------------------------------------------------------
        # Append an index number to duplicates

        unique_names = self.ensure_unique(unique_names)

        
        # -----------------------------------------------------------------
        # Allocate additional plots
        

        if plots_to_add:
            
            # Get titles/names/figures           
            plot_titles = [titles.pop(-1) for title in range(len(plots_to_add))]
            plot_names = [unique_names.pop(-1) for title in range(len(plots_to_add))]
            figures = plots_to_add.values()


            for i, figure in enumerate(figures):
                self.additional_plots.append({'title': plot_titles[i],
                                               'name': plot_names[i],
                                               'fig': figure})

        
        # -----------------------------------------------------------------
        # Create blueprints for all dataframes

        all_dfs = [self.input_df.copy()] + list(dfs_to_add.values())
        
        for i, df in enumerate(all_dfs):

            
            # use clearer names for the primary dataframe
            if not i:
                fa = 'Field Analysis'
                ov = 'Overview'
                pv = 'Pivot'
            
            else:
                fa = f'Field Analysis | {unique_names[i]}'
                ov = f'Overview | {unique_names[i]}'
                pv = ''


            self.blueprints.append(Blueprint(name=unique_names[i],
                                             title=titles[i],
                                             input_df=df,
                                             large_report=self.large_report,
                                             field_analysis= fa,
                                             overview=ov,
                                             pivot=pv))






    def calculate_full_path(self) -> Path:
    

        """    
        Calculates a full file path for an Excel workbook given 
            a path like string or Path object

        Returns
        -------
        Path
            An absolute pathlib Path object for an Excel workbook

        """
        
        wb_path = Path(self.wb_path)
        wb_directory = Path.cwd()


        # Handle correct extension is passed
        if wb_path.suffix in ['.xlsx', '.xlsm']:

            # If full path is provided
            if wb_path.is_absolute():      

                new_path = wb_path

            # if a full path isn't provided, construct it
            else:
                new_path = Path().cwd() / wb_path
            
            return new_path
            
        # Handle other full paths
        elif wb_path.is_absolute():
            
            # Handle full path with some other extension
            if wb_path.is_absolute() and bool(wb_path.suffix):
                
                wb_directory = wb_path.parent
            
            # Handle full path without a file name/exension
            elif wb_path.is_absolute() and not bool(wb_path.suffix):
                    wb_directory = wb_path
        
        # Handle partial paths
        elif not wb_path.is_absolute():
            
            # with no extension
            if not bool(wb_path.suffix):
                wb_directory = Path().cwd().joinpath(*wb_path._tail)  # type: ignore

            # With an incorrect extension
            if bool(wb_path.suffix) and wb_path.suffix not in ['.xlsx', '.xlsm']:
                wb_directory = Path().cwd() / wb_path.parent
            

        # If only a directory has been calculated, add file name to it and return
        name = self.sanitize_name(self.name, file_name=True)

        if self.no_vba:
            wb_file_name = f"{name}.xlsx"
        else:
            wb_file_name = f"{name}.xlsm"

        new_path = wb_directory / wb_file_name


        return new_path



    def ensure_unique(self, str_list: list) -> list:
        
        """
        Ensures each list item is unique by appending 
            numbers to duplicates

        Parameters
        ----------
        str_list : list
            A list with potentially duplicate items

        Returns
        -------
        list
            A list of unique items
            
        """

        # Track how many times we have seen each item
        seen_counts = Counter()
        result = []

        for item in str_list:
            seen_counts[item] += 1
            if seen_counts[item] > 1:

                
                # Append the occurrence number to the duplicate
                
                # Make room if needed
                if len(item) > 12:
                    limit = len(item) - 14
                    result.append(f"{item[:limit]}_{seen_counts[item]}")
                
                else:
                    result.append(f"{item}_{seen_counts[item]}")

            else:
                
                # Keep the first occurrence as-is
                result.append(item)

        return result


    def create_blank_template(self, progress_bar: tqdm):

        """
        Creates a blank Field Analysis template, overwriting if necessary

        """

        


        # Return an error if there's an existing file and no overwrite flag

        if self.path.is_file() and not self.overwrite:

            self.theme.warn_print(f"Error: There is already a workbook named {self.path}!")
            self.theme.warn_print("Use overwrite=True or rename/remove the existing workbook")
            
            sys.exit()


        # Delete the file if there's an overwrite flag, return error if it's open
    
        elif self.path.is_file() and self.overwrite:
            try:

                send2trash.send2trash(self.path)

                self.exit_msg += f"\nThe previously existing file was sent to your {self.env.junk_drawer}"
                
            except OSError:
                
                self.theme.warn_print("\nError: The workbook cannot be overwritten while open!")
                sys.exit()

            except Exception:
                
                self.theme.warn_print(f"An unexpected error occurred when deleting {self.path.name}")
                sys.exit()

        progress_bar.update(2)
        

        # Create parent directories if necessary and a blank copy of the template
        self.path.parent.mkdir(parents=True, exist_ok=True)
        progress_bar.update(1)



        # ---------------------------------------------------------------------
        # Convert template to xlsx if necessary

        # if no_vba has been set, open the default template, remove macro triggers from shapes, and save as xlsx

        if self.no_vba:

            # Hide warnings and convert file
            if self.debug:
                print("Converting template to .xlsx")

            with xw.App(visible=self.debug, add_book=False) as app:
                
                app.api.DisplayAlerts = self.debug

                wb = app.books.open(xlsm_file, read_only=False)
                ws = wb.sheets('Field Analysis')

                # Loop through all shapes and clear their triggers
                for shp in ws.shapes:
                    shp.api.OnAction = ""

                # Save as .xlsx and reenable alerts
                wb.api.SaveAs(str(self.path.resolve()), FileFormat=51)

                app.api.DisplayAlerts = False
        else:

            # Create a copy of the template as .xlsm
            shutil.copy(xlsm_file, self.path)
        
        progress_bar.update(1)
        



    def expand_range(self, name: str, ws: xw.Sheet, columns: int = 0):
        
        """
        Expands a named range by +/- extra_columns

        Parameters
        ----------

        name : str
            Name of named range

        ws : xw.Sheet
            Worksheet of named range

        columns : int, optional
            The amount of columns to expand the named range by
            Defaults to 0
            
        """

        ws.activate()

        ws.range(name).resize(row_size=None, 
                              column_size=columns).name = name



    def sanitize_name(self, 
                      input_str: str, 
                      restricted_name: bool = True, 
                      file_name: bool = False):

        """
        Strips out all punctuation from a string and optionally removes 
            limits to 14 characters to fit within the 31 character 
            limit of worksheet names
        
        """
        
        
        # Strip out illegal file name characters
        file_name_pattern =   r'[\\/:*?"<>|]'

        if file_name:
            return re.sub(file_name_pattern, '', input_str)
        

        
        # Strip out illegal characters from potential worksheet/table objects/object names

        # Remove all punctuation, leaves spaces
        sanitized_name_patttern = r'[^a-zA-Z0-9 _-]'
        sanitized_name = re.sub(sanitized_name_patttern, '', input_str)


        # Return the lenth limited one if it's restricted    
        if restricted_name:
            return sanitized_name[:14]
        else:
            return sanitized_name

    # TODO: Incorporate this
    def white_list(self, file_path: Path):

        """
        Remove mark of the web from a file
        
        """

        
        # Ensure the file exists
        if file_path.is_file():

            # Remove the quarantine attribute if it exists
            try:
                subprocess.run(["xattr", "-d", "com.apple.quarantine", 
                                str(file_path)], 
                                check=True)
            except subprocess.CalledProcessError:
                pass

        else:
            print("File does not exist.")

# TODO: Incorporate this into the create template function
    def vba_object_model_trusted(self, book: xw.Book) -> bool:
        """
        Determines whether "Trust Access to the VBA Object Model" has been set.

        """

        try:
            if self.env.win:
                _ = book.api.VBProject    
            elif self.env.mac:
                _ = book.api.VBProject.VBComponents.Count
                
            vba_object_model_trust = True
            
        except Exception:
            vba_object_model_trust = False
        
        return vba_object_model_trust

    # TODO: Incorporate this
    def set_cell_alignment(self,
                           input_range: xw.Range,
                           horizontal: str ='', 
                           vertical: str=''):
        
        center = Constants.xlCenter  # noqa: F841
        left = Constants.xlLeft # noqa: F841
        
        if self.env.win:
            if horizontal:
                input_range.api.HorizontalAlignment = eval(horizontal)
            if vertical:
                input_range.api.VerticalAlignment = eval(vertical)

        elif self.env.mac:
            if horizontal:
                input_range.api.horizontal_alignment.set(eval(horizontal))
            if vertical:
                input_range.api.vertical_alignment.set(eval(vertical))


    
    # * `ws.range("Toggles").api.Orientation = 0`
    # TODO: Incorporate this
    def set_text_orientation(self, 
                             input_range: xw.Range, 
                             degrees: int=0):
        
        """
        Sets text orientation (0 = normal horizontal text, -90 = up/down)

        Parameters
        ----------
        input_range : xw.Range
            Range object to orient
        
        degrees : int, optional
            Degrees to set text orientation to.
            Defaults to 0
        """
        # 
        if self.env.win:
            
            input_range.api.Orientation = degrees
        
        elif self.env.mac:
            
            input_range.api.orientation.set(degrees)



    def get_updated_pivot_table(self,
                                ws: xw.Sheet) -> Any:
        """
        Finds and returns a native Pivot Table object by its name 
        on both Windows and macOS.
        """

        if self.env.mac:
            pt = ws.api.pivot_tables['Pivot']
            pt.update_pivot_table()
        
        elif self.env.win:
            ws.activate()
            pt = ws.api.PivotTables('pvt_Pivot')
            pt.PivotCache().Refresh()
        
        return pt


    def hide_rows(self,
                  input_range: xw.Range,
                  hide: bool=True):

        """Cross-platform helper to show/hide Excel rows."""
        
        if self.env.win:
            input_range.api.EntireRow.Hidden = hide
            
        elif self.env.mac:
            input_range.api.entire_row.hidden.set(hide)


class PerformanceLogger():

    """
    Class representing a performance logger

    """
    
    def __init__(self, 
                 input_args: dict,
                 env: Environment) -> None:

        # ------------------------------------------------------------------------------
        # Initialize Performance Logging

        self.start: float = time.time()
        self.section_last: float = time.time()
        self.last: float = time.time()

        self.performance_logs: dict[str, list] = defaultdict(list)
        self.section_performance: pd.DataFrame = pd.DataFrame()
        self.field_performance: pd.DataFrame = pd.DataFrame()
        self.performance_metadata: pd.DataFrame = pd.DataFrame()
        self.config: pd.DataFrame = pd.DataFrame()
        self.env: pd.DataFrame = pd.DataFrame()
        self.errors: pd.DataFrame = pd.DataFrame()
        self.total_production_time: float

        self.log_open(input_args, env=env)
        


    def log_open(self, input_args: dict,
                 env: Environment):
        
        """
        Logs an xleda configuration for loggging

        """

        # ---------------------------------------------------------
        # Set up config df
       
        # Remove large items from input args
        del input_args['input_df']
        del input_args['self']
        input_args['wb_path'] = str(input_args['wb_path'].resolve())
        input_args['add_plots'] = str([k for k in input_args['add_plots'].keys()])
        input_args['add_dfs'] = str([k for k in input_args['add_dfs'].keys()])

        # Convert to dataframe, transpose, set column names, and store
        input = pd.DataFrame.from_records([input_args]).T.astype(str)
        input = input.reset_index()
        input.columns = ['Input Argument', 'Value']
        
        self.config = input
        

        # ---------------------------------------------------------
        # Set up envirnment df

        env_dict = vars(env).copy()
        del env_dict['win']
        del env_dict['mac']

        env_df = pd.DataFrame.from_records([env_dict]).T
        env_df = env_df.reset_index()
        env_df.columns = ['Environment Variable', 'Value']
        self.env = env_df



    def log(self,
            log_type: str, 
            details: dict = {}):

        """
        Logs production performance data

        """

        now = time.time()

        # Add performance timing to details
        for k, v in details.items():

            if not v:
                
                # Use section.last for section logs
                if log_type == 'section':
                    details[k] = now - self.section_last
                    self.section_last = now

                else:
                    # If the log type isn't section, use the newest last
                    details[k] = now - max(self.last, self.section_last)


        # Append to log storage
        self.performance_logs[log_type].append(details)
            
        # Set last values
        self.last = now
      


    def close(self, blueprints: list[Blueprint], additional_plots: int):

        """
        Closes performance logging by converting logs to dataframes

        """


        self.total_production_time = time.time() - self.start

        

        # Create section performance df
        self.section_performance = pd.DataFrame.from_records(self.performance_logs['section'])
        

        # Create field_performance df
        field_logs = self.performance_logs['pivot'] + self.performance_logs['plots']
        self.field_performance = pd.DataFrame.from_records(field_logs)
        

        # Compile performance_metadata df
        self.performance_metadata = pd.DataFrame.from_records(self.performance_logs['performance_metadata'])

        # Compile errors df, adding a blank entry if there are none
        if len(self.performance_logs['error']) == 0:
            self.performance_logs['error'].append({'Error Type': 'No Errors occurred',
                                                   'Value': 'N/A'})
        self.errors = pd.DataFrame.from_records(self.performance_logs['error'])
        

        
        # Add % of Production Time columns
        for df in [self.field_performance, self.section_performance]:
            df['% of Production Time'] = df['Production Time in Seconds']/self.total_production_time






