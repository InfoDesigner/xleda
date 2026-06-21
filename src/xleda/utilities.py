from __future__ import annotations

# Base imports
import random
import colorsys
import re
import time
import subprocess
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

# Plotting imports
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib as mpl 
from matplotlib.figure import Figure


# TQDM Imports
from tqdm.auto import tqdm
import threading

# xlwings imports
import xlwings as xw
from xlwings.constants import Constants


# Import wb for typing
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .main import wb


# ---------------------------------------------
# Set variables

# Set OS variables
win = platform.system() == 'Windows'
mac = platform.system() == 'Darwin'

if win:
    import winreg
    

if mac:
    from appscript import app, k # type: ignore
    

# Plot/print vars
plt.style.use("dark_background")
separator = "\n" + ("-" * 100)


# DataFile variables
db_file_extensions = ['.sqlite', '.sqlite3', '.db', '.db3', '.s3db', '.sl3', '.duckdb', '.ddb']
excel_extension = ['.xlsx', '.xls', '.xlsm', '.xlsb']
standalone_extensions = ['.csv', '.feather', '.parquet', '.xml', '.json', '.rdata', '.pkl', '.pickle', '.pck'] 
supported_extensions = standalone_extensions + db_file_extensions + excel_extension

# Dataset variables
default_row_limit = 25_000
default_column_limit = 50
upper_row_limit = 1_000_000
upper_column_limit = 16_000


# Template Variables
xlsm_file = Path(__file__).parent / "xleda_template.xlsm"
xlsx_file = Path(__file__).parent / "xleda_template.xlsx"


template_objects ={"meta":[],
                   "SinglePlot": [],
                   "Field Analysis": ["tbl_SourceData"],
                   "|": ["tbl_FieldOverview", "tbl_debug_environment", "tbl_debug_config", "tbl_DfOverview",  "tbl_debug_errors", "tbl_debug_section"]}



class Environment():
    

    def __init__(self, debug: bool = False) -> None:
        
        """
        Gathers os, python, and terminal details for debugging.

        """    
        
        # Primary environment detail
        
        self.os = platform.system()
        self.win = win
        self.mac = mac
        self.env_type = self.get_env_type()
        self.excel_version = self.get_excel_version()
        self.debug = debug


        self.validate_compatibility()
        

        # Determine file recovery tool
        if self.win:
            self.junk_drawer = 'Recycle Bin'
        elif self.mac:
            self.junk_drawer = 'Trash'


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
        
        try:
            if win:
                # Read the current Excel version from Classes Root
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer") as key:
                    app_version, _ = winreg.QueryValueEx(key, "")
                    app_version = app_version.split(".")[-1]
            if mac: 
                # Ask macOS to look up the Excel application version
                app_version = app("Microsoft Excel").version.get()
            
            return app_version
        
        except Exception:
            return ""
            
        

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

            self.warn_print(compatibility_msg)

            sys.exit()



    def warn_print(self, text: str):
        
        """
        Prints text in red bold for warning messages

        """
        
        print(f"\033[1;31m{text}\033[0m")
        


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
        
        # Creates a thread to refresh the progres bar
        self.start_auto_refresh(pbar=pbar)

        return pbar


    def start_auto_refresh(self, pbar: tqdm):

        """
        Refreshing tqdm progress bars until they are complete.

        """

        interval=0.1

        def _refresh_loop():
            while not pbar.disable:
                pbar.refresh()
                time.sleep(interval)
                
        thread = threading.Thread(target=_refresh_loop, daemon=True)
        thread.start()
        return thread




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



    def greyscale_tab(self, ws: xw.Sheet, iteration:int):

        """
        Colors worksheet tabs to a shade of grey for contrast with adjacent worksheets
            
        """
        
        # 26*10 > 255 limit for RGB so limit to 9
        iteration = (iteration + 1) % 9
        multiplier = 26 * iteration


        # Set color for field analysis worksheet
        if self.env.win:
            color = (multiplier) + (multiplier*256)  + (multiplier*256*256)
            ws.api.Tab.Color = color
        
        elif self.env.mac:
            color = ((multiplier), (multiplier), (multiplier))
            ws.api.sheet_tab.color.set((color))



class Plotter():
    
    """
    Class that represents a xleda plotting object

    """
    
    """
    Class that represents a xleda plotting object

    """
    
    def __init__(self, theme: Theme, env: Environment) -> None:
        
        """
        Creates theme appropriate plots and optinally writes them to a range

        """
        
        self.theme_color = theme.theme_color
        self.env = env


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
                                                  left=target_left,
                                                  top=target_top,
                                                  width=target_width,
                                                  height=target_height)
            
            pic.api.placement.set(k.placement_move_and_size)


    def create_composition_plot(self, input_series: pd.Series) -> Figure:
        
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
        category_counts = input_series.value_counts()
        top_5_categories = category_counts.head(5)
        total_records = len(input_series)
        other_counts = total_records - top_5_categories.sum()

        # Assemble plot values
        categories = list(top_5_categories.index) + ["Other"]
        values = list(top_5_categories.values) + [other_counts]

        
        
        # --------------------------------------------------
        # Setup plot
        
        # Initialize the plot inside a context manager that creates headless plots
        with mpl.rc_context({'backend': 'Agg'}):
            
            fig, ax = plt.subplots(figsize=(8, 8))

            y_pos = range(len(categories))[::-1]

            # Add bars to plot
            ax.barh(y_pos, values,color=self.theme_color, height=0.5, edgecolor='silver')

            
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
                pct = (val / total_records) * 100

                # Truncate long category name
                display_cat = str(cat)
                if len(display_cat) > 6:
                    display_cat = display_cat[:5] + ".."

                # Add category name and adjust left to prevent overlap
                ax.text(-0.55, y, display_cat, color="white", va="center", 
                        ha="left", fontsize=font_size, transform=ax.get_yaxis_transform(), )

                # Add category percentage
                ax.text(-0.05, y, f"{pct:.0f}%", color="white", va="center", ha="right", 
                        fontsize=font_size, transform=ax.get_yaxis_transform(), )

                # Add category count to the right of the bars
                ax.text(val + max_val * 0.02, y, str(val), color="white",
                        va="center", ha="left", fontsize=font_size, )

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

        # Create the plot inside a context manager that creates headless plots
        with mpl.rc_context({'backend': 'Agg'}):
            fig, ax = plt.subplots(figsize=(5, 5))
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

    def __init__(self, ds: DataSet) -> None:
        
        """
        Generates an ExportDict object that provides access to xleda 
            metadata through both key/val and dot notation

        Parameters
        ----------
        ds : 
            An xleda dataset

        """


        for key in ['description', 'definitions', 'notes', 'lists', 'field_overview', 'df_overview', 'altered_source_data']:
            
            val = eval(f"ds.{key}")

            # If the value exists or is a dataframe, add it to the dictionary and class object
            if isinstance(val, pd.DataFrame) or val:
                setattr(self, key, eval(f"bp.{key}"))
                self[key] = eval(f"bp.{key}")
    


class Config():

    """
    Class that represents an xleda configuration

    """

    def __init__(self,
                 wb: wb,
                 **kwargs: Any) -> None:
        
        """
        Primary configuration object for an xleda workbook

        

        """
        
        # Set file name
        self.file_name = self.sanitize_name(input_str=kwargs['name'], name_type='file')
        self.env = wb.env
        self.no_vba = kwargs['no_vba']
        self.wb_path = kwargs['wb_path']
        self.overwrite: bool = kwargs['overwrite']
        self.debug: bool = kwargs['debug']
        self.open_wb: bool = kwargs['open_wb']
        self.exit_msg = separator
        self.large_report = kwargs['large_report']
        self.export = kwargs['export']



        # Calculate the target file path        
        self.path: Path = self.calculate_full_path()

        
        # Ensure unique names for all worksheets/tables
        self.ensure_unique(wb=wb)





    def calculate_full_path(self) -> Path:
    
        """    
        Calculates a full file path for an Excel workbook given 
            a path like string or Path object

        Returns
        -------
        Path
            An absolute pathlib Path object for an Excel workbook

        """

        # Set vars
        wb_path = Path(self.wb_path)
        wb_directory = Path.cwd()
        new_path = None

                
        # Handle correct extension is passed
        if wb_path.suffix in ['.xlsx', '.xlsm']:

            # If full path is provided, use it
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
                wb_directory = Path().cwd() / wb_path

            # With an incorrect extension
            if bool(wb_path.suffix) and wb_path.suffix not in ['.xlsx', '.xlsm']:
                wb_directory = Path().cwd() / wb_path.parent
            

        
        # If only a directory has been calculated, add file name to it
        if not new_path:
        
            if self.no_vba:
                wb_file_name = f"{self.file_name}.xlsx"
            else:
                wb_file_name = f"{self.file_name}.xlsm"

            new_path = wb_directory / wb_file_name
        

        return new_path



    def ensure_unique(self, wb: wb):
        
        """
        Ensures:
            Each dataframe has a unique name for worksheets and tables
            Each plot has a unique name for worksheet
            
        """
        
        datasets = wb.datasets


        # --------------------------------------------------------
        # Set up counter
        
        # Track how many times each item has been seen
        seen_counts = Counter()
                              

        # Add pre-existing worksheet/table names to counter
        pre_existing_worksheets = list(template_objects.keys()) + ["Pivot"]
        pre_existing_tables = [table for table_list in template_objects.values() for table in table_list] + ["pvt_Pivot"]
        pre_existing_names = pre_existing_worksheets + pre_existing_tables
        
        for name in pre_existing_names:
            seen_counts[name] += 1


        # --------------------------------------------------------
        # Ensure unique names for each prepared dataframe
        
                
        for dataset in datasets:
            
            # Set vars
            title = self.sanitize_name(dataset.name, name_type='title')
            table_name = self.sanitize_name(dataset.name, name_type='table')
            
            # Update seen counts
            seen_counts[title] += 1
            seen_counts[table_name] += 1
            
            
            # ------------------------------------------------------
            # Handle title
            
            # Append the occurrence number to duplicates
            if seen_counts[title] > 1:
                
                # Append the occurrence number to the duplicate
                if len(title) > 31:
                    past_limit = len(title) - 31
                    title = f"{title[:past_limit]}_{seen_counts[title]}"

            dataset.update_name(title)
            

            # ------------------------------------------------------
            # Handle table_name
            
            # Append the occurrence number to duplicates
            if seen_counts[table_name] > 1:
                
                # Append the occurrence number to the duplicate
                if len(table_name) > 31:
                    past_limit = len(table_name) - 31
                    table_name = f"{table_name[:past_limit]}_{seen_counts[table_name]}"

            dataset.table_name = table_name
                

        # ------------------------------------------------------
        # Handle plots
        
        if self.plots:
            
            # Replacement dictionary
            new_plots = {}


            
            for plot_name, figure in self.plots.items():
                
                # Set var, update count
                plot_name = self.sanitize_name(plot_name, name_type='title')
                seen_counts[plot_name] += 1
                                
                
                # Append the occurrence number to duplicates
                if seen_counts[plot_name] > 1:
                    

                    if len(plot_name) > 31:
                        past_limit = len(plot_name) - 31
                        plot_name = f"{plot_name[:past_limit]}_{seen_counts[plot_name]}"
                
                # Add amended plot to the replacement dictionary
                new_plots[plot_name] = figure
                
            
            self.plots = new_plots
                


    def create_blank_template(self, progress_bar: tqdm):

        """
        Creates a blank xleda template, overwriting if necessary

        """


        # Return an error if there's an existing file and no overwrite flag

        if self.path.is_file() and not self.overwrite:

            self.env.warn_print(f"Error: There is already a workbook named {self.path}!")
            self.env.warn_print("Use overwrite=True or rename/remove the existing workbook")
            
            sys.exit()


        # Delete the file if there's an overwrite flag, return error if it's open
    
        elif self.path.is_file() and self.overwrite:
            try:

                send2trash.send2trash(self.path)

                self.exit_msg += f"\nThe previously existing file was sent to your {self.env.junk_drawer}"
                
            except OSError:
                
                self.env.warn_print("\nError: The workbook cannot be overwritten while open!")
                sys.exit()

            except Exception:
                
                self.env.warn_print(f"An unexpected error occurred when deleting {self.path.name}")
                sys.exit()

        progress_bar.update(2)
        

        # Create parent directories if necessary
        self.path.parent.mkdir(parents=True, exist_ok=True)

        progress_bar.update(1)




        

        # --------------------------------------------------------
        # Create a copy of the template

        # Remove MOTW from the templates before using if on MacOS
        if self.env.mac:
            self.white_list_templates()
        

        if self.path.suffix == '.xlsx':
            shutil.copy(xlsx_file, self.path) 
        else:
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
                      name_type: str) -> str:
        
        """
        Prepares names for use as file/worksheet/table names
        
        """

        # Handle file names that only need certain punctuation marks
        #  removed and can retain a long length and spaces
        file_name_pattern =   r'[\\/:*?"<>|]'
        if name_type == 'file':
            return re.sub(file_name_pattern, '', input_str)
        
        else:
            
            # All other names have most punctuation removed
            sanitized_name_patttern = r'[^a-zA-Z0-9 _-]'
            sanitized_name = re.sub(sanitized_name_patttern, '', input_str)[:31]
                     
            # Table names also have spaces removed
            if name_type == 'table':
                sanitized_name = sanitized_name.replace(" ","")
            
            return sanitized_name



    def white_list_templates(self):

        """
        Remove mark of the web from the templates on MacOS
        
        """

        
        # Ensure the templates exist
        if xlsm_file.is_file() and xlsx_file.is_file():

            # Remove the quarantine attribute if it exists
            try:
                subprocess.run(["xattr", "-d", "com.apple.quarantine", 
                                str(xlsm_file)],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
                subprocess.run(["xattr", "-d", "com.apple.quarantine", 
                                str(xlsx_file)],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                
            except subprocess.CalledProcessError:
                pass

        else:
            self.env.warn_print("File does not exist.")



    def vba_object_model_trusted(self) -> bool:
        
        """
        Determines whether "Trust Access to the VBA Object Model" has been set.

        """

        # Use xlwings to create a temporary app/book and test if the VBA object model is trusted
        with xw.App(visible=False):

            
            # Creates a new, temporary, unsaved workbook
            book = xw.Book()  
            
            try:
                if self.env.win:
                    _ = book.api.VBProject    
                elif self.env.mac:
                    _ = book.api.VBProject.VBComponents.Count
                    
                vba_object_model_trust = True
                
            except Exception:
                vba_object_model_trust = False
        
            # Close the workbook when finished
            book.close()


        return vba_object_model_trust



    def set_cell_alignment(self,
                           input_range: xw.Range,
                           horizontal: str ='', 
                           vertical: str=''):
        
        """
        Sets text alignment of a given Excel range

        """
        
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


    
    def set_text_orientation(self, 
                             input_range: xw.Range, 
                             degrees: int=0):
        
        """
        Sets text orientation (0 = normal horizontal text, -90 = up/down)

        """
        
        if self.env.win:
            input_range.api.Orientation = degrees
        
        elif self.env.mac:
            input_range.api.text_orientation.set(degrees)
            
            

    def get_updated_pivot(self,
                          pt_name: str,
                          ws: xw.Sheet) -> Any:
        
        
        if self.env.mac:
            
            # Update pivot table
            pt = ws.api.pivot_tables[pt_name]
            pt.refresh_table()
        
        elif self.env.win:
            
            # Update pivot table
            pt = ws.api.PivotTables(pt_name)
            pt.PivotCache().Refresh()
            
        return pt


        
        

    def get_configured_pivot_range(self,
                                   ws: xw.Sheet,
                                   pivot_fields: list[str]) -> xw.Range:
        """
        Configures a pivot table and returns its range.

        """

        pt = self.get_updated_pivot(ws=ws,
                                    pt_name="pvt_Pivot")

        if self.env.mac:
            
            # Add fields
            pt.add_fields_to_pivot_table(row_fields=pivot_fields)
            
            # Collapse fields/remove subtotals
            script = f'''
            tell application "Microsoft Excel"
                tell workbook "{self.wb_path}"
                    tell sheet "{"Pivot"}"
                        tell pivot table "{"pvt_Pivot"}"
                            set all_fields to pivot fields
                            repeat with pf in all_fields
                                try
                                    set subtotals pf subtotal index 1 without value
                                    set show detail of pf to false
                                end try
                            end repeat
                        end tell
                    end tell
                end tell
            end tell
            '''
            subprocess.run(["osascript", "-e", script], capture_output=True, text=True)

            # Get range of pivot table
            pivot_range = ws.range(pt.table_range1.get_address())
        
        elif self.env.win:
            

            # Add fields
            pt.AddFields(RowFields=pivot_fields)

            # Loop through and configure fields
            for field in pivot_fields[::-1]:
                
                # Collapse pivot fields
                try:
                    pt.PivotFields(field).ShowDetail = False
                except Exception:
                    pass

                # Hide subtotals
                try:
                    if self.env.win:
                        pt.PivotFields(field).Subtotals = tuple([False] * 12)
                except Exception:
                    pass
                
                # Get range of pivot table
                pivot_range = ws.range(pt.TableRange1.Address)
            
        return pivot_range


    def hide_rows(self,
                  input_range: xw.Range,
                  hide: bool=True):

        """
        Hides Excel rows
        
        """
        
        if self.env.win:
            input_range.api.EntireRow.Hidden = hide
            
        elif self.env.mac:
            input_range.api.entire_row.hidden.set(hide)



class PerformanceLogger():

    """
    Class representing a performance logger

    """
    
    def __init__(self,
                 wb: wb) -> None:

        # ------------------------------------------------------------------------------
        # Initialize Performance Logging

        self.start: float = time.time()
        self.last: float = time.time()

        self.performance_logs: dict[str, list] = defaultdict(list)
        self.section_performance: pd.DataFrame = pd.DataFrame()
        self.performance_metadata: pd.DataFrame = pd.DataFrame()
        self.config: pd.DataFrame = pd.DataFrame()
        self.env: pd.DataFrame = pd.DataFrame()
        self.errors: pd.DataFrame = pd.DataFrame()
        self.total_production_time: float       

        self.config  = self.add_config_log(wb=wb)
        self.env = self.add_env_log(wb=wb)
        


    def add_config_log(self,
                       wb: wb) -> pd.DataFrame:
        
        """
        Logs an xleda configuration for loggging

        """

        # ---------------------------------------------------------
        # Set up config df
               
        config = {'file_name': wb.file_name,
                  'dataframes': ''.join([ds.name for ds in wb.datasets]),
                  'plots': ', '.join(wb.plots.keys()), 
                  'theme_color': wb.theme.theme_color, 
                  'large_report': wb.cfg.large_report, 
                  'overwrite': wb.cfg.overwrite, 
                  'wb_path': wb.cfg.wb_path, 
                  'open_wb': wb.cfg.open_wb, 
                  'no_vba': wb.cfg.open_wb, 
                  'export': wb.cfg.export, 
                  'debug': wb.cfg.debug}

        
        # Convert to dataframe, transpose, set column names, and store
        config_df = pd.DataFrame.from_records([config]).T
        config_df = config_df.reset_index()
        config_df.columns = ['Input Argument', 'Value']
        
        return config_df
        

    def add_env_log(self,
                    wb:wb) -> pd.DataFrame:
        
        # ---------------------------------------------------------
        # Set up envirnment df

        env_dict = vars(wb.env).copy()
        del env_dict['win']
        del env_dict['mac']
        
        # Create a dataframe, transpose it, add column names, and save it
        env_df = pd.DataFrame.from_records([env_dict]).T
        env_df = env_df.reset_index()
        env_df.columns = ['Environment Variable', 'Value']
        
        return env_df



    def log(self,
            log_type: str, 
            details: dict = {}):

        """
        Logs production performance data

        """

        now = time.time()

        # Add performance timing to details
        for key, value in details.items():

            if not value:
                details[key] = now - self.last


        # Append to log storage
        self.performance_logs[log_type].append(details)
            
        # Set last values
        self.last = now
      


    def close(self):

        """
        Closes performance logging by converting logs to dataframes

        """


        self.total_production_time = time.time() - self.start

        

        # Create section performance df
        self.section_performance = pd.DataFrame.from_records(self.performance_logs['section'])
        

        # Compile performance_metadata df
        self.performance_metadata = pd.DataFrame.from_records(self.performance_logs['performance_metadata'])

        # Compile errors df, adding a blank entry if there are none
        if len(self.performance_logs['error']) == 0:
            self.performance_logs['error'].append({'Error Type': 'No Errors occurred',
                                                   'Value': 'N/A'})
        self.errors = pd.DataFrame.from_records(self.performance_logs['error'])
        

        
        # Add % of Production Time columns
        df = self.section_performance
        df['% of Production Time'] = df['Production Time in Seconds']/self.total_production_time



class DataError(Exception):
    
    """
    An exception class for capturing file parsing errors

    """
    
    def __init__(self, 
                 message: str):
        
        super().__init__(message)


        
class DataSetParser():

    def __init__(self, 
                 data: Path | dict[str, pd.DataFrame],
                 large_report: bool = False):

        self.data = data
        self.large_report = large_report
        self.datasets: list[DataSet] = []
        

        # If a dict of dataframes is provided, use it to create a list of PreparedDataFrames 
        if isinstance(data, dict) and all(isinstance(v, pd.DataFrame) for v in data.values()):
            self.from_dataframes()

        # If a path is provided, use it to create a list of PreparedDataFrames 
        elif isinstance(data, Path):
            self.file_path = data
            self.file_name = data.stem
            self.read_data_file()

 

    def read_data_file(self):
        
        """
        Creates a list of PreparedDataFrame objects from from supported data files

        """    
        
        assert isinstance(self.file_path, Path)

        extension = self.file_path.suffix.lower()
        is_file = self.file_path.is_file()
        
        # Handle .rdata files that don't have an extension
        if is_file and not extension:
            extension = self.file_path.name.lower()
            
        # if a file with an unsupported extension is provided, raise an error 
        if not is_file or extension not in supported_extensions:
            supported = ", ".join(sorted(supported_extensions))
            raise DataError(f"Unsupported file type '{extension}'. Supported files: {supported}")


        # ----------------------------------------------------
        # Create datasets from tabular files
        
        elif extension == ".csv":
            self.from_csv()

        elif extension == ".feather":
            self.from_feather()
           
        elif extension == ".parquet":
            self.from_parquet()
            
        elif extension in ['.xml', '.json']:
            self.from_xml_json()
        
        # ----------------------------------------------------
        # Create datasets from multidimensional files
        
        elif extension == ".rdata":
            self.from_rdata()
            
        elif extension in db_file_extensions:
            self.from_db_file()
            
        elif extension in excel_extension:
            self.from_excel()

               


    def identify_db_type(self) -> str:
        
        """
        Identifies the correct database type from files that could be duckdb, sqlite, or unknown

        """
        
        db_type = ""
        
        try:
            with open(self.file_path, 'rb') as f:
                header = f.read(16)
                
                # sqlite format
                if header.startswith(b'SQLite format 3'):
                    db_type = "sqlite"
                
                # duckdb format
                elif header[8:12] == b'DUCK':
                    db_type = "duckdb"
                    
        except Exception:
            pass
        
        return db_type



    def from_xml_json(self):

        """
        Creates a list of PreparedDataFrame objects from a supported from xml or json files

        """
        
        # Datafile imports
        import xmltodict
        import json
        

        try: 
            with open(self.file_path) as file: 
                
                # Parse xml 
                if self.file_path.suffix == '.xml': 
                    data = xmltodict.parse(file.read())
                    
                    # Identify the root xml node 
                    if isinstance(data, dict) and len(data) == 1:
                        root_val = list(data.values())[0]
                        
                        # If the root contains a list of items, parse the values
                        if isinstance(root_val, dict) and len(root_val) == 1:
                            data = list(root_val.values())[0]
                        else:
                            data = root_val

                # Parse json 
                elif self.file_path.suffix == '.json': 
                    data = json.load(file) 
                
                # Create a dataframe
                df = pd.json_normalize(data) 
                self.datasets.append(DataSet(input_df=df,
                                                  name=self.file_name,
                                                  large_report=self.large_report))

                
        except Exception: 
            raise DataError("XML/JSON file not successfully parsed") 
            
        

    def from_pickle(self):
        
        """
        Creates a list of PreparedDataFrame objects from dataframes inside a pickle file

        """
        

        # Datafile imports
        import pickle
        
        try: 

            # Open the pickle file
            with open(self.file_path, 'rb') as f:
                while True:
                    
                    # Load the next available object
                    data = pickle.load(f)
                    
                    # Save the the object to df_list if it is a dataframe
                    if isinstance(data, pd.DataFrame):

                        self.datasets.append(DataSet(input_df=data,
                                                          name=self.file_name,
                                                          large_report=self.large_report))
        except Exception: 
            raise DataError("Pickle file not successfully parsed")      


        

    def from_rdata(self):
        """
        Creates a list of PreparedDataFrame objects from an RData file
        
        """
        # Datafile imports
        import rdata
        
        try:
            # Parse the file into pure Python/R objects.
            parsed_file = rdata.parser.parse_file(str(self.file_path))
            
            # Convert the entire workspace to its Python equivalents
            converted_data = rdata.conversion.convert(parsed_file)
            
        except Exception as e:

            raise DataError(f"RData file not successfully parsed: {e}")
        
        # Iterate through variables and extract valid DataFrames
        for name, value in converted_data.items():
            try:
                df = None
                
                # Objects that have already been converted to a dataframe
                if isinstance(value, pd.DataFrame):
                    df = value
                    
                # Convert to a dataframe if possible
                elif isinstance(value, dict) or hasattr(value, '__array__'):
                    try:
                        df = pd.DataFrame(value)
                    
                    # Continue to the next object if it doesn't convert successfully
                    except Exception:
                        continue
                
                # Append successfully converted dfs
                if df is not None and not df.empty:
                    self.datasets.append(DataSet(input_df=df,
                                                          name=name,
                                                          large_report=self.large_report))
                    
            except Exception:
                
                pass





    def from_db_file(self):
        
        """
        Creates a list of PreparedDataFrame objects from a supported database file

        """
        
        # Datafile imports
        import duckdb

        
        db_type = self.identify_db_type()

            
        # Handle unsupported database file types
        if not db_type:
            
            # database file type undetermined
            raise DataError("Database file not successfully parsed")


        # Create sqlite connection/query
        if db_type == 'sqlite':

            # Open a DuckDB instance
            conn = duckdb.connect()

            # Attach SQLite file
            conn.sql(f"ATTACH '{self.file_path}' AS sqlite_db (TYPE sqlite);")
            
            query = """
                SELECT name AS table_name 
                FROM sqlite_schema 
                WHERE type = 'table' AND name NOT LIKE 'sqlite_%'
            """
            
        # Create duckdb connection/query
        elif db_type == 'duckdb':
            
            conn = duckdb.connect(self.file_path)
            
            # Adds schema. to the table name
            query = """
                SELECT table_schema || '.' || table_name 
                FROM information_schema.tables 
                WHERE table_schema NOT IN ('information_schema', 'pg_catalog')
            """
            
        # Create a dataframe from each table
        try:
        
            # Gather schema.table and table lists to test for collissions
            long_table_names = [row[0] for row in conn.execute(query).fetchall()]
            short_table_names = [table.split('.')[-1] for table in long_table_names]
            has_name_collissions = len(short_table_names) != len(set(short_table_names))
            
            # Use the concise name if possible
            if has_name_collissions:
                tables = long_table_names
            else:
                tables = short_table_names

        
            for table in tables:
                
                # Use an underscore format for the dictionary key if necessary
                if has_name_collissions:
                    table = table.replace('.', '_')

                                
                # Create dataframes for each table
                if db_type == 'sqlite':
                    table_query = conn.sql(f"SELECT * FROM sqlite_db.{table}")
                    df = table_query.df()
                                
                # Use a faster read process for duckdb tables
                elif db_type == 'duckdb':
                    df = conn.execute(f"SELECT * FROM {table}").df()
                                        
                    
                self.datasets.append(DataSet(input_df=df,
                                                      name=table,
                                                      large_report=self.large_report))
                    
        except Exception:
            raise DataError("Database file not successfully parsed")
                
        finally:
                conn.close()
            


    def from_excel(self):
        
        """
        Creates a list of PreparedDataFrame objects from an Excel file

        """    
        
        try:
            # Open the target workbook
            with xw.App(visible=False) as app:
                
                book = app.books.open(self.file_path)
                
                # Loop through all worksheets
                for sheet in book.sheets:
                    
                    # Loop through tables
                    for table in sheet.tables:
                        
                        # Create dataframes from each
                        df = table.range.options(pd.DataFrame, index=False, header=True).value
                        
                        self.datasets.append(DataSet(input_df=df,
                                                              name=table.name,
                                                              large_report=self.large_report))
                        
                book.close()
                
        except Exception:
            raise DataError("Excel file not successfully parsed")


    def from_feather(self):
            
            """
            Creates a list of PreparedDataFrame objects from a feather file

            """    
            
            try:
                self.datasets.append(DataSet(input_df=pd.read_feather(self.file_path),
                                                  name=self.file_name,
                                                  large_report=self.large_report))
                    
            except Exception:
                raise DataError("Feather file not successfully parsed")

    def from_csv(self):
            
            """
            Creates a list of PreparedDataFrame objects from a feather file

            """    
            
            try:
                self.datasets.append(DataSet(input_df=pd.read_csv(self.file_path),
                                                  name=self.file_name, 
                                                  large_report=self.large_report))
                    
            except Exception:
                raise DataError("CSV file not successfully parsed")

    def from_parquet(self):
            
            """
            Creates a list of PreparedDataFrame objects from a feather file

            """
            
            try:
                self.datasets.append(DataSet(input_df=pd.read_parquet(self.file_path),
                                                  name=self.file_name, 
                                                  large_report=self.large_report))
                    
            except Exception:
                raise DataError("Parquet file not successfully parsed")


    def from_dataframes(self):
            
        """
        Creates a list of PreparedDataFrame objects from a dictionary of dataframes

        """    
        
        assert isinstance(self.data, dict)
        
        try:
            
            for name, df in self.data.items():

                self.datasets.append(DataSet(input_df=df,
                                             name=name,
                                             large_report=self.large_report))

                
        except Exception:
            raise DataError("Dataframes not successfully parsed")



class DataSet():
    
    """
    A class that represents a dataframe that has been appropriately 
    subsampled and includes all necessary metadata
    
    """
    
    
    def __init__(self, 
                 name: str,
                 large_report: bool = False,
                 input_df: pd.DataFrame = pd.DataFrame()):
        
        # Capture inputs/placeholder values
        self.large_report = large_report
        self.name: str = name
        self.table_name: str = ""
        self.warning: bool = False
        self.warning_msg: str = ""
        
        # Subsample input_df if necessary and record new dimensions
        self.original_df: pd.DataFrame = self.create_original_df(input_df=input_df)
        self.rows = len(self.original_df)
        self.columns = len(self.original_df.columns)
        
        
        # Add Record List, HasBlank, Record Hash fields to source_data
        self.source_data = self.create_source_data()
       

        # add field_overview/df_overview/df_metadata
        self.field_metadata: pd.DataFrame = self.create_field_metadata()
        self.field_overview: pd.DataFrame = self.create_field_overview()
        self.df_overview: pd.DataFrame = self.create_df_overview()
        
        
        # Add field analysis metadata dfs
        self.df_metadata: dict[str, Any] = self.create_df_metadata()
        self.composition_df = self.field_metadata.loc[["Data type", "Distinct %", "Missing %", "Memory Usage %", "Memory Usage", "Distinct", "Count", "Missing"]]
        self.summary_stats_df = self.field_metadata.loc[["Mean", "Median", "Mode", "Standard Deviation", "Variance"]]
        self.percentiles_df = self.field_metadata.loc[["Min", "5%", "25%", "50%", "75%", "95%", "Max", "Range", "IQR"]]
        

        # Add export placeholders
        self.description = ""
        self.definitions: dict = {}
        self.notes: dict = {}
        self.lists: dict = {}
        self.source_data: pd.DataFrame = pd.DataFrame()
    

    def update_name(self, 
                    new_name: str):
        
        """
        Updates the dataframe name across the PreparedDataframe object
        
        """
        
        self.name = new_name
        self.field_overview['Dataframe'] = new_name
        self.df_overview['Dataframe'] = new_name
        
        
        
    def create_original_df(self,
                           input_df: pd.DataFrame) -> pd.DataFrame:
        
        df = input_df.copy()
        rows = len(df)
        columns = len(df)
        

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
        
        
        
        # Add index, subsample if needed, record adjusted rows/columns
        df['index'] = df.index
        df = (df.iloc[:, :columns].sample(n=rows).sort_index())
        
        return df

        

    def create_source_data(self) -> pd.DataFrame:
        
        """
        Configures source data by adding EDA workflow columns        

        """
        
        df = self.original_df.copy()
        df.insert(loc=0, column="Record List", value="False")
        df['HasBlank'] = self.source_data.isnull().any(axis=1).astype(int)
        df["Record Hash"] = pd.util.hash_pandas_object(df, index=False)
        
        return df
        
        

    def create_field_metadata(self) -> pd.DataFrame:

        """
        Produces a field_metadata dataframe

        """


        # --------------------------------------------------
        # Collect metadata

        # Omit added columns
        df = self.original_df

        # Order of output fields
        row_order = ["Data type", "Memory Usage", "Memory Usage %", "Distinct", "Distinct %", "Count", "Count %", "Missing", "Missing %", "Mean", "Median", "Mode", "Standard Deviation", "Variance", "Min", "5%", "25%", "50%", "75%", "95%", "Max", "Range", "IQR"]
        
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
                "Memory Usage": df.memory_usage(deep=True),
                "Memory Usage %": df.memory_usage(deep=True) / df.memory_usage(deep=True).sum(),
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



    def create_field_overview(self) -> pd.DataFrame:

        """
        Creates a transposed copy of the field metadata and adds placeholder columns
        
        """
        
        df = self.field_metadata.T
        
        col_order = ['_', 'Dataframe', 'Field', 'Definition', 'Field Notes', 'Data type', 'Distinct %', 'Missing %', 'Memory Usage %', 'Memory Usage', 'Distinct', 'Count', 'Count %', 'Missing', 'Mean', 'Median', 'Mode', 'Standard Deviation', 'Variance', 'Min', '5%', '25%', '50%', '75%', '95%', 'Max', 'Range', 'IQR']
        
        df['Field'] = df.index
        df['Dataframe'] = self.name
        
        df = df.assign(**dict.fromkeys(['_', 'Field Notes', 'Definition'], None))
        df = df[col_order]
        
        return df



    def create_df_metadata(self) -> dict[str, Any]:

        """
        Creates a dictonary of df-level metadata

        """
        
        df = self.original_df

        df_metadata = {'Rows': self.rows,
                       'Columns': self.columns,
                       'Memory Usage (bytes)': df.memory_usage(deep=True).sum(),
                       'Distinct Rows %': (len(df.drop_duplicates()) / len(df)),
                       'Missing %': df.isnull().mean().mean(),}
        
        return df_metadata
        
        
        
    def create_df_overview(self) -> pd.DataFrame:

        """
        Creates a dataframe of df-level metadata

        """

        df = self.original_df

        df_metadata = {'_': '',
                       'Dataframe': self.name,
                       'Dataframe Description': '',
                       'Memory Usage (bytes)': df.memory_usage(deep=True).sum(),
                       'Rows': self.rows,
                       'Columns': self.columns,
                       'Fields Defined %'
                       'Missing Values %': df.isnull().mean().mean(),
                       'Subsampled': bool(self.warning_msg)}
        
        return pd.DataFrame.from_records([df_metadata])
                