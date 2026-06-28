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
from datetime import datetime
import os
import pandas as pd
from typing import Any
import send2trash
from collections import defaultdict
import shlex

# Persistent settings imports
import json
from importlib.metadata import version
from packaging.version import parse
import urllib.request



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

# Plotting imports
import seaborn as sns
import matplotlib.style as mpl_style 
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg
mpl_style.use("dark_background")


# ---------------------------------------------
# Set variables


# Set OS variables
win = platform.system() == 'Windows'
mac = platform.system() == 'Darwin'


# Platform specific imports
if win:
    import winreg
if mac:
    from appscript import app, k # type: ignore


# print vars
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


template_objects ={"SinglePlot": [],
                   "Field Analysis": ["tbl_SourceData"],
                   "Overview": ["tbl_FieldOverview", "tbl_debug_environment", "tbl_debug_config", "tbl_DfOverview",  "tbl_debug_section"]}


help_message = (f"{separator}\n\n"
                r"Use 'xleda wb \<data file path\>' to create a workbook" + "\n\n"
                "### Supported file types:\nCSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML" + "<br><br>\n\n"
                f"### Expected extensions:{str(supported_extensions)[1:-1]}\n\n"
                "For more documentation, visit https://github.com/InfoDesigner/xleda")


class Settings():
    
    def __init__(self, 
                 env: Environment,
                 locals: dict = {},
                 version: bool = False) -> None:
        
        """
        Evaluates inputs and incorporate/updates persistent settings
        
        """
        
        # Save env, Get/create persisent settings path
        self.env = env
        self.settings_path = self.get_settings_path()
        

        # Runtime settings
        self.overwrite: bool = locals.get('overwrite', False)
        self.debug: bool = locals.get('debug', False)
        self.open_wb: bool = locals.get('open_wb', True)
        self.large_report = locals.get('large_report', False)
        self.export = locals.get('export', False)
        self.wb_path: str | Path = locals.get('wb_path', False)
        self.data: pd.DataFrame | str | Path | dict[str, pd.DataFrame] | None  = locals.get('data', None)  
        self.file_name: str = locals.get('file_name', 'xleda')
        self.input_df: pd.DataFrame | None = locals.get('input_df', None)
        self.plots: dict[str, Figure] = locals.get('plots', {})
        self.data_argument: str = ''



        # Persistent setting defaults
        self.no_vba: bool = False
        self.theme_color: str = "#262626"
        self.version_check_date: str = datetime.now().isoformat()
        self.update_msg = ''

        # Override defaults with persistent settings from disk if they exist
        self.load_settings()
        
        # If version is passed, run the update check
        if version:
            self.version_check()
            
        elif locals:
        
            # Otherwise, parse inputs 
            self.parse_inputs(input_args=locals)
            
            # Write persistent settings to disk
            self.save_settings()
            
            # TODO: Figure out when/if to incorporate this
            # # Check for updates if enough time has passed
            # if (datetime.now() - datetime.fromisoformat(self.version_check_date)).days > 3:
            #     self.version_check()
            
            

            
        


        
        
        

    def get_settings_path(self) -> str:
        
        """
        Determines persistent settings folder
        
        """
        
        # Windows: AppData/Local
        if self.env.win:
            base_dir = Path(os.environ.get("LOCALAPPDATA", "~")).expanduser()
            
        # macOS: ~/Library/Application Support
        if self.env.mac:
            base_dir = Path("~/Library/Application Support").expanduser()

        # Create/use a hidden xleda folder
        assert isinstance(base_dir, Path)
        storage_dir = base_dir / '.xleda'
        
        # Try to write settings but fail silently
        try:
            storage_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        
        # Return settings.json path
        return str(storage_dir / "settings.json")


    def parse_inputs(self,
                     input_args: dict):
        
        """
        Update persistent settings as needed
        
        """

        no_vba = input_args.get('no_vba', None)
        theme_color = input_args.get('theme_color', '')

        
        # if provided arguments are explicit, change class properties
        if no_vba is not None:        
            self.no_vba = no_vba
            
        # if provided arguments are explicit, change class properties
        if theme_color:
            self.theme_color = theme_color

    
    def version_check(self):
        
        
        # Get versions to compare
        installed_version = version('xleda')

        
        # Pull the latest version from pypi
        pypi_url = "https://pypi.org/pypi/xleda/json"
        try:
            with urllib.request.urlopen(pypi_url) as response:
                data = json.loads(response.read().decode('utf-8'))
                latest_ver = data['info']['version']
                
        except Exception:
            pass

        # Provide an update suggestion if a newer version is available
        if parse(latest_ver) > parse(installed_version):

            # Store the update version message
            
            # TODO: Figure out when/if to incorporate this into runtime
            self.update_msg += separator + "\n\n✨ A newer version of xleda is available\n\n"
            self.update_msg += f"   Installed version: {installed_version}\n"
            self.update_msg += f"   Latest PyPI version: {latest_ver}\n" + separator
            
            
        else:

            # Store the update version message
            self.update_msg += separator + "\n\n✨ You are running the latest version of xleda\n\n"
            self.update_msg += f"   Installed version: {installed_version}\n"
            self.update_msg += f"   Latest PyPI version: {latest_ver}\n" + separator
        
        # Reset the counter
        self.version_check_date = datetime.now().isoformat()


    def load_settings(self):
        
        """
        Loads persistent data from disk
        
        """
        path = Path(self.settings_path)
        
        if path.is_file():

            try:
                with open(path, "r", encoding="utf-8") as f:
                    
                    json_data = json.load(f)
                    
                    # Save persistent settings to self
                    self.no_vba = json_data.get('no_vba', False)
                    self.theme_color = json_data.get('theme_color', "#262626") 
                    self.version_check_date = json_data.get('version_check_date', datetime.now())
                            
                
            except (json.JSONDecodeError, OSError):
                
                pass


    def save_settings(self) -> None:
        
        """
        Saves persistent settings to disk
        
        """
        
        path = Path(self.settings_path)
        
        # Write to a temporary file first, then rename it
        tmp_path = path.with_suffix(".tmp")
        
        persistent_settings = {'no_vba': self.no_vba,
                               'theme_color': self.theme_color,
                               'version_check_date': self.version_check_date,
                               'update_msg': self.version_check_date}
        

        try:
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(persistent_settings, f, indent=4)
            tmp_path.replace(path)
            
        except OSError:
            if tmp_path.exists():
                tmp_path.unlink() 




class Environment():
    

    def __init__(self) -> None:
        
        """
        Gathers OS, Python, and terminal details for debugging.

        """    
        
        # Primary environment detail

        self.os = platform.system()
        self.win = win
        self.mac = mac
        self.env_type = self.get_env_type()
        self.excel_version = self.get_excel_version()        

        
        
        # Verify Compatibility
        self.validate_compatibility()
        

        # Determine file recovery tool
        if self.win:
            self.junk_drawer = 'Recycle Bin'
        elif self.mac:
            self.junk_drawer = 'Trash'


        # Additional environment details
        terminal_size = shutil.get_terminal_size()
        self.date = datetime.now().date().strftime("%#m/%#d/%Y")
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
        Uses Windows Registry/Appscript to obtain the Excel version

        Returns
        -------
        str
            The version of Excel discovered
            
        """
        
        try:
            if win:
                
                # Read the current Excel version from the Windows registry
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer") as key:
                    app_version, _ = winreg.QueryValueEx(key, "")
                    app_version = app_version.split(".")[-1]
                    
            if mac:
                
                # Use appscript to look up the Excel application version
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

        
        # Provide a debug output if an incompatible environment is detected
        if not self.compatible:

            msg = (
                separator + "\nxleda requires a full desktop version of Microsoft Excel\n"
                "\nIt has been developed and tested on Windows\n"
                "\nIt should also work on MacOS though this has not yet been tested\n\n"
                f"{'Requires MacOS/Windows':<25} | Detected {self.os:<20} | {os_compatible:<20}\n"
                f"{'Requires Excel >=16.0':<25} | Detected {self.excel_version:<20} | {excel_compatibility:<20}\n"
                + separator)
            
            raise CompatibilityError(msg)



    def warn_print(self, text: str):
        
        """
        Prints text in red bold for warning messages

        """
        
        print(f"\033[1;31m{text}\033[0m")





class Theme():

    """
    CLass that represents an xleda theme
    
    """

    def __init__(self,
                 settings: Settings) -> None:

        """
        Configures an xleda theme that includes workbook 
            theme, progress bars, and console printing

        """
    
    
        # ------------------------------------------------------------------------------
        # Configure Theme
        
        
        if settings.theme_color == 'random':
            color = "#{:06x}".format(random.randint(0, 0xFFFFFF))
        else:
            
            # Protect against too long hex colors
            color = settings.theme_color[:7]
        
        # Save updated theme color
        settings.theme_color = color
        self.theme_color = color
        
        self.black_text: bool = self.use_black_text(self.theme_color)
        self.print_theme: str = self.ensure_readable(self.theme_color[:7])
        self.env = settings.env


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





class Plotter():
    
    """
    Class that represents a xleda plotting object

    """

    def __init__(self, 
                 settings: Settings) -> None:
        
        """
        Creates theme appropriate plots and optinally writes them to a range

        """
        

        self.theme_color = settings.theme_color
        self.env = settings.env

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
        
        # Initialize a Figure object, attach it to a canvas, and add a 1 column/row subplot
        fig = Figure(figsize=(9, 9))
        canvas = FigureCanvasAgg(fig) # noqa: F841
        ax = fig.add_subplot(111)


        # Add bars to plot
        y_pos = range(len(categories))[::-1]
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
        fig.subplots_adjust(left=0.4, right=0.9)
        max_val = max(values)



        # --------------------------------------------------
        # Setup data bars

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

        
        # Initialize a Figure object, attach it to a canvas, and add a 1 column/row subplot 
        fig = Figure(figsize=(6, 6))
        canvas = FigureCanvasAgg(fig) # noqa: F841
        ax = fig.add_subplot(111)

        
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
                    warn_singular=False)

        
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


        for key in ['description', 'definitions', 'notes', 'lists', 'field_overview', 'df_overview', 'source_data']:
            
            val = eval(f"ds.{key}")

            # If the value exists or is a dataframe, add it to the dictionary and class object
            if isinstance(val, pd.DataFrame) or val:
                setattr(self, key, eval(f"ds.{key}"))
                self[key] = eval(f"ds.{key}")
    


class Template():

    """
    Class that represents an xleda template

    """

    def __init__(self,
                 wb: wb) -> None:
        
        """
        Handles creating the template and select operations

        """
        
        # Set initial file name
        self.file_name = self.sanitize_name(input_str=wb.settings.file_name, 
                                            name_type='file')
        self.env = wb.env
        self.no_vba = wb.settings.no_vba
        self.overwrite = wb.settings.overwrite
        self.logger = wb.logger
        self.theme_color = wb.theme.theme_color
        self.black_text = wb.theme.black_text
        self.wb = wb
        self.datasets = wb.datasets
        self.settings = wb.settings
        self.plots = wb.settings.plots


        # Calculate the target file path
        self.path: Path = self.calculate_full_path()

        # Ensure unique names for all worksheets/tables
        self.ensure_unique_names()
        
        
        # If not exporting, create the blank template
        if not wb.settings.export:
            self.create_blank_template()


    def calculate_full_path(self) -> Path:
    
        """    
        Calculates a full file path for an xleda workbook

        """
        
        input_path = Path(self.settings.wb_path)
                       
        
        # Construct file name
        if self.no_vba:
            wb_file_name = f"{self.file_name}.xlsx"
        else:
            wb_file_name = f"{self.file_name}.xlsm"
        
        
        # Handle correct extension is passed
        if input_path.suffix in ['.xlsx', '.xlsm']:

            # if a correct extension with a full path is provided, use it
            if input_path.is_absolute():      
                new_path = input_path
                
                # If a dataframe is provided without a name and a wb_path 
                # is provided with a full path, use that for the name
                if self.wb.datasets[0].name == 'xleda':
                    self.wb.datasets[0].update_name(new_path.stem)
                    
                
                
                            
            # if a correct extension with a partial path is provided, construct the full path
            else:
                new_path = Path().cwd() / input_path

            
            return new_path
            
        # Handle other full paths
        elif input_path.is_absolute():
            
            # Handle full path with some other extension
            if input_path.is_absolute() and bool(input_path.suffix):
                
                new_path = input_path.parent / wb_file_name
            
            # Handle full path without a file name/exension
            elif input_path.is_absolute() and not bool(input_path.suffix):
                    new_path = input_path / wb_file_name
        
        # Handle partial paths
        elif not input_path.is_absolute():
            
            # with no extension
            if not bool(input_path.suffix):
                new_path = Path().cwd() / input_path / wb_file_name

            # With an incorrect extension
            elif bool(input_path.suffix) and input_path.suffix not in ['.xlsx', '.xlsm']:
                new_path = Path().cwd() / input_path.parent / wb_file_name

        return new_path



    def ensure_unique_names(self):
        
        """
        Ensures:
            Each dataframe has a unique name for worksheets and tables
            Each plot has a unique name for worksheet
            
        """
        
        datasets = self.datasets


        # --------------------------------------------------------
        # Set up counter
        
        # Track how many times each item has been seen
        seen_counts = Counter()
                              

        # Add pre-existing worksheet/table names to counter
        pre_existing_worksheets = list(template_objects.keys())
        pre_existing_tables = [table for table_list in template_objects.values() for table in table_list]
        pre_existing_names = pre_existing_worksheets + pre_existing_tables
        
        for name in pre_existing_names:
            seen_counts[name] += 1


        # --------------------------------------------------------
        # Ensure unique names for each prepared dataframe
        
                
        for i, dataset in enumerate(datasets):
            
            # Set vars
            name = self.sanitize_name(dataset.name, name_type='name')
            table_name = f'tbl_{self.sanitize_name(dataset.name, name_type="table")}'
            
            # Update seen counts
            seen_counts[name] += 1
            seen_counts[table_name] += 1
            
            
            # ------------------------------------------------------
            # Handle name
            
            # Append the occurrence number to duplicates
            if seen_counts[name] > 1:
                
                # Append the occurrence number to the duplicate
                if len(name) > 31:
                    
                    # Truncate name if necessary to make room for the occurence number
                    past_limit = len(name) - 31 - len(str(seen_counts[table_name]))
                    name = name[:past_limit]
                
                # Append the occurence number
                name = f"{name}_{seen_counts[name]}"

            # Update the dataset name
            dataset.update_name(name)
            

            # ------------------------------------------------------
            # Handle table_name
            
            # Append the occurrence number to duplicates
            if seen_counts[table_name] > 1:
                
                # Truncate name if necessary to make room for the occurence number
                if len(table_name) > 31:
                    past_limit = len(table_name) - 31 - len(str(seen_counts[table_name]))
                    table_name = f"{table_name[:past_limit]}_{seen_counts[table_name]}"

            dataset.table_name = table_name
                

        # ------------------------------------------------------
        # Handle plots
        
        if self.plots:
            
            # Replacement dictionary
            new_plots = {}


            
            for plot_name, figure in self.plots.items():
                
                # Set var, update count
                plot_name = self.sanitize_name(plot_name, name_type='name')
                seen_counts[plot_name] += 1
                                
                
                # Append the occurrence number to duplicates
                if seen_counts[plot_name] > 1:
                    
                    # Truncate name if necessary to make room for the occurence number
                    if len(plot_name) > 31:
                        past_limit = len(plot_name) - 31 - len(str(seen_counts[plot_name]))
                        plot_name = f"{plot_name[:past_limit]}_{seen_counts[plot_name]}"

                # Add amended plot to the replacement dictionary
                new_plots[plot_name] = figure
                
            
            self.plots = new_plots
                


    def create_blank_template(self):

        """
        Creates a blank xleda template, overwriting if necessary

        """


        # Return an error if there's an existing file and no overwrite flag

        if self.path.is_file() and not self.overwrite:
            
            msg = f"Error: There is already a workbook named {self.path}!"
            msg += "Use overwrite=True or rename/remove the existing workbook"

            raise TemplateError(msg)


        # Delete the file if there's an overwrite flag, return error if it's open
    
        elif self.path.is_file() and self.overwrite:
            try:

                send2trash.send2trash(self.path)
                
                self.logger.exit_msg += f"\nThe previously existing file was sent to your {self.env.junk_drawer}\n"
                
            except OSError:
                
                
                msg = "\nError: The workbook cannot be overwritten while open!"
                raise TemplateError(msg)

                

            except Exception:
                
                
                msg = f"An unexpected error occurred when deleting {self.path.name}"
                raise TemplateError(msg)
                        

        # Create parent directories if necessary
        self.path.parent.mkdir(parents=True, exist_ok=True)

        

        # --------------------------------------------------------
        # Create a copy of the template

        # Remove MOTW from the templates before using if on MacOS
        if self.env.mac:
            self.white_list_templates()
        

        if self.path.suffix == '.xlsx':
            shutil.copy(xlsx_file, self.path) 
        else:
            shutil.copy(xlsm_file, self.path)



    def add_book(self, book: xw.Book):
        
        """
        Adds the book object to the template
        
        """
        
        self.book = book
        


    def validate(self):
        
        """
        Validates that expected objects are present in the workbook
        
        """
        
        # Use short var
        book = self.book
        
        
        try:
            
            # ---------------------------------------------
            # Validate worksheets and tables exist
            
            actual_objects = {}
            
            # Check actual worksheets/tables
            for sheet in book.sheets:
                actual_objects[sheet.name] = [tbl.name for tbl in sheet.tables]
                
                
        except Exception:
            
            pass
        
        
            
        # Collate missing worksheets
        missing_sheets = [sheet for sheet in template_objects.keys() if sheet not in actual_objects.keys()]
        
        # Collate missing tables
        expected_tables = [table for table_list in template_objects.values() for table in table_list]
        actual_tables = [table for table_list in actual_objects.values() for table in table_list]
        missing_tables = [table for table in expected_tables if table not in actual_tables]
        
        # Provide a constructive output message if something is missing
        if missing_sheets or missing_tables:
        
            # Provide an output message
            msg = "\n\nTemplate has been modifed:\n\n"
            
            if missing_sheets:
                msg += f"The following worksheets are missing\n    {missing_sheets}\n\n"
                
            if missing_tables:
                msg += f"The following tables are missing\n    {missing_tables}"
            
            raise TemplateError('msg')



    def add_worksheets(self,
                       progress_bar: tqdm):

        """
        Creates worksheets for each dataframe and any plots

        """

        
        # -----------------------------------------------------------------
        # Set vars
        
        # Set short vars
        book = self.book
        datasets = self.datasets

        # -----------------------------------------------------------------
        # Validates that the expected template objects are present
        
        self.validate()
        
        progress_bar.update(1)

       
        # Add field analysis worksheets for each dataframe except the first

        for i, dataset in enumerate(datasets):


            # Create a worksheet for all datasets except the first one
            if i:
                
                # Copy the sheet, rename the table
                ws = book.sheets('Field Analysis').copy(name=dataset.name)
                ws.tables[0].name = dataset.table_name

                # Add a color gradient to the worksheet tab to distinguish among them
                self.greyscale_tab(ws=ws, iteration=i)
        
            progress_bar.update(1)

        
        # Use the worksheet template for the first dataset
        ws = book.sheets("Field Analysis")
        ws.tables[0].name = datasets[0].table_name
        ws.name = datasets[0].name



    def add_field_analyses(self,
                            progress_bar: tqdm):

        """
        Configures all Field Analysis worksheets

        """

        
        
        
        for ds in self.datasets:

            # --------------------------------------------------
            # Set variables

            book = self.book
            ws = book.sheets(ds.name)
            ws.activate()
            source_table = ws.tables[ds.table_name]
            df = ds.source_data.copy()

            # --------------------------------------------------
            # Format metadata placeholders

            # Set worksheet theme/name
            self.expand_range(name="Theme", ws=ws, columns=ds.columns + 2)
            self.set_theme(ws.range("Theme"))
            ws.range("Name").value = ds.name


            # Use the FormatRange column to create placeholders for each source data column
            columns_to_format = ds.columns -3
            if columns_to_format > 0:
                format_from = ws.range("FormatRange")
                format_to = (ws.range("FormatRange").offset(0, 1).resize(None, columns_to_format))
                format_from.copy()
                format_to.paste()

            # Clear clipboard
            book.app.cut_copy_mode = False
            
            # Add all header values except Record List
            headers = ds.source_data.columns.to_list()[1:]
            
            # Add header values
            ws.range("Headers_Start").value = headers
            
            progress_bar.update(1)


            # --------------------------------------------------
            # Adjust Named Ranges to fit dataframe size
            

            # Expand named ranges to fit number of columns
            expand_ranges = (["FieldList" + str(i) for i in range(1, 9)] + ["FieldRange", "Notes", "Definitions", "Headers"])

            for name_range in expand_ranges:
                
                self.expand_range(name=name_range, 
                                      ws=ws, 
                                      columns=ds.columns)

            # Resize the dataset description range and merge it
            ws.range("Description").resize(None, 2).merge()
          
                        
            
            # Show/Hide Data Size Warning
            if ds.warning:
                ws.range("Warning").value = ds.warning_msg
                self.hide_rows(ws.range("Warning"), hide=not ds.warning)
                

            progress_bar.update(1)



            # --------------------------------------------------
            # Add metadata

            ws.range("Dimensions").options(transpose=True).value = list(ds.df_metadata.values())
            ws.range("Composition")[0, 0].offset(0,1).value = ds.composition_df.values
            ws.range("Summary_Stats")[0, 0].offset(0,1).value = ds.summary_stats_df.values
            ws.range("Percentiles")[0, 0].offset(0,1).value = ds.percentiles_df.values
            
            
            progress_bar.update(1)
            

            # --------------------------------------------------
            # Add Source Data
            
            
            # Convert fields with datatypes that Excel doesn't support to string
            supported_dtype_columns = df.select_dtypes(include=['number', 'bool', 'datetime64', 'str'], exclude='timedelta').columns
            unsupported_dtype_columns = [col for col in df.columns if col not in supported_dtype_columns]
            df[unsupported_dtype_columns] = df[unsupported_dtype_columns].astype(str)

            
            # Add to worksheet
            try:
                source_table.update(df, index=False)
            
            # If any unsupported columns prevent writing to Excel, convert the df to string before writing
            except Exception:
                
                source_table.update(df.astype(str), index=False)

            progress_bar.update(1)


            
            # --------------------------------------------------
            # Set formatting for tbl_SourceData and added columns

            
            self.set_cell_alignment(input_range=source_table.range,
                                    horizontal='center')
            
            record_list = source_table.range[:, :1 ]
            other_added_columns = source_table.range[:, -2: ]

            # Reduce contrast to subdue added fields
            for dimmed_range in [record_list, other_added_columns]:
                self.greyscale_range(dimmed_range)


            progress_bar.update(1)
            



    def add_overview(self, 
                     progress_bar: tqdm):

        """
        Configure Overview worksheet
        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object
        
        """

        
        # --------------------------------------------------
        # Set variables
        
        book = self.book
        datasets = self.datasets
        ws = book.sheets("Overview")
        df_overview_table = ws.tables["tbl_DfOverview"]
        field_overview_table = ws.tables["tbl_FieldOverview"]
        
        # Compile both overview dfs
        df_overview_df = pd.concat([ds.df_overview for ds in datasets], ignore_index=True)
        field_overview_df = pd.concat([ds.field_overview for ds in datasets], ignore_index=True)
        
        ws.activate()
        
        progress_bar.update(1)
        
        
        # --------------------------------------------------
        # # Add rows to the df_overview section if there is more than 2 dataframes
        
        if len(datasets) > 2:
    
            start_row = df_overview_table.range.last_cell.row + 2
            end_row = start_row + len(datasets) - 3
            row_range_string = f"{start_row}:{end_row}"
            
            ws.range(row_range_string).insert(shift="down")
        
        # Adjust named ranges to fit new data
        ws.range("Dataframes").resize(row_size=len(datasets) +2, column_size=None).name = "Dataframes"
        ws.range("Fields").resize(row_size=sum([ds.columns for ds in datasets]) +2, column_size=None).name = "Fields"
        
        # Group rows here with the new ranges
        
        self.group_rows(ws.range("Dataframes"))
        self.group_rows(ws.range("Fields"))

                        
        progress_bar.update(1)
        
        
        
        # --------------------------------------------------
        # Set theme, and write data to tables
                    
        # Set theme on target tables
        self.set_theme(df_overview_table.range[0,:])
        self.set_theme(field_overview_table.range[0,:])
        
        # Update the primary tables
        df_overview_table.update(df_overview_df, index=False)
        field_overview_table.update(field_overview_df, index=False)
        
        progress_bar.update(1)
        
        
        # --------------------------------------------------
        # Add formulas to tables 
        
        # Set formulas
        df_description_formula = r'''=INDIRECT("'"&[@Dataframe]&"'!Description")'''
        df_links_formula = r'''=HYPERLINK("#'"&[@Dataframe]&"'!Headers_Start", "Link")'''
        df_fields_defined_pct_formula = r'''=IFERROR(SUMPRODUCT((tbl_FieldOverview[Dataframe]=[@Dataframe]) * (tbl_FieldOverview[Definition]<>"Definition"))/[@Columns],"")'''
        field_links_formula = r'''=HYPERLINK("#" & CELL("address", XLOOKUP([@Field],INDIRECT("'"&[@Dataframe]&"'!Headers"),INDIRECT("'"&[@Dataframe]&"'!Headers"),"")), "Link")'''
        field_definitions_formula = r'''=XLOOKUP([@Field],INDIRECT("'"&[@Dataframe]&"'!Headers"),INDIRECT("'"&[@Dataframe]&"'!Definitions"),"")'''
        field_notes_formula = r'''=XLOOKUP([@Field],INDIRECT("'"&[@Dataframe]&"'!Headers"),INDIRECT("'"&[@Dataframe]&"'!Notes"),"")'''
        
        
        # Add df_overview formulas
        ws.range("tbl_DfOverview[_]").formula = df_links_formula
        ws.range("tbl_DfOverview[Dataframe Description]").formula = df_description_formula
        ws.range("tbl_DfOverview[Fields Defined %]").formula = df_fields_defined_pct_formula
        
        
        # Add field_overview formulas
        ws.range("tbl_FieldOverview[_]").formula = field_links_formula
        ws.range("tbl_FieldOverview[Definition]").formula = field_definitions_formula
        ws.range("tbl_FieldOverview[Field Notes]").formula = field_notes_formula
        
        
        # Set the cursor to the first link in the df_overview table
        df_overview_table.range[1,0].select()
                    
        progress_bar.update(1)



    def add_plots(self,
                  progress_bar: tqdm):
        
        """
        Adds all plots to an xleda workbook
        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object

        """

        
        # Set primary vars
        
        datasets = self.datasets
        plotter = self.wb.plotter
        book = self.book
        
        
      
        for ds in datasets:
        
            # --------------------------------------------------
            # Set vars, unhide target rows

            ws = book.sheets(ds.name)
            df = ds.source_data.copy().iloc[:, 1:-2]
            ws.activate()
                

            # Set initial ranges for added plots 
            histogram_range = ws.range("Histogram")
            composition_range = ws.range("CompositionTable")
            
            # Unhide the target ranges
            self.hide_rows(histogram_range, hide=False)
            self.hide_rows(composition_range, hide=False)


            # --------------------------------------------------
            # Add plots for all except added columns

            for col in df.columns:


                # --------------------------------------------------
                # Add Composition Table

                composition_table = plotter.create_composition_plot(df[col])

                plotter.add_small_plot(target_range=composition_range,
                                       fig=composition_table,
                                       name=f'composition_{col}')

                if pd.api.types.is_numeric_dtype(df[col]):



                    # --------------------------------------------------
                    # Add Histogram

                    histogram = plotter.create_histogram_plot(df[col])

                    plotter.add_small_plot(target_range=histogram_range,
                                           fig=histogram,
                                           name=f'histogram_{col}')


                # --------------------------------------------------
                # Increment Target Ranges/progress bar

                histogram_range = histogram_range.offset(0, 1)
                composition_range = composition_range.offset(0, 1)
                
                
                
                
                progress_bar.update(1)
            
            
            # --------------------------------------------------
            # Initialize the UI for use
            
            # Set cursor position
            ws.range('Headers_Start').select()

            # Orient toggles, and collapse subsections
            self.set_text_orientation(input_range=ws.range("Toggles"))
            self.set_text_orientation(input_range=ws.range("TopToggle"), degrees=-90)

            for excel_range in ["Data_Description", "Composition", "Summary_Stats", 
                                "Percentiles", "Field_Lists", "Compiled_Lists"]:
                            
                self.hide_rows(ws.range(excel_range), hide=True)
            
            progress_bar.update(1)



    def add_plot_sheets(self,
                        progress_bar: tqdm):
            
        """
        Adds additional plot worksheets

        """

        
        # Set vars
        book = self.book


        # --------------------------------------------------
        # Add additional plots



        for plot_name, figure in self.plots.items():
           
            
            # Plots will be added before all other sheets
            anchor_sheet = book.sheets[0]
            
            # Create a copy of the SinglePlot template, make it visible, set theme
            ws = book.sheets("SinglePlot").copy(before=anchor_sheet, name=plot_name)
            ws.visible = True
            ws.activate()
            self.set_theme(ws.range("Theme"))

            # Set target range, name, and autofit name range
            plot_range = ws.range("SinglePlot")
            ws.range("Name").value = plot_name
            ws.range("Theme")[0].select()

            
            ws.pictures.add(figure, 
                            name=plot_name, 
                            update=True,
                            left=plot_range.left, 
                            top=plot_range.top)
            
            progress_bar.update(1)



    def add_debug(self):

        """
        Configures the debug section of the Overview worksheet
        
        """
        
        
        # ------------------------------------------------------------
        # Set vars, add data, ensure section is hidden
        
        book = self.book
        ws = book.sheets('Overview')
        ws.activate()
       
        
        # Write debug tables
        ws.tables("tbl_debug_environment").update(self.logger.env, index=False)
        ws.tables("tbl_debug_config").update(self.logger.config.astype(str), index=False)
        ws.tables("tbl_debug_section").update(self.logger.section_performance, index=False)


        # Hide debug section and set toggle orientation
        self.hide_rows(ws.range("Debug"), hide=True)
        self.set_text_orientation(ws.range("DebugToggle"))

        
        # ----------------------------------------------------------
        # Since this is the last operation...


        # Scroll worksheet tabs and set focus to Overview
        book.sheets[0].activate()
        book.sheets("Overview").activate()
        


    def expand_range(self, 
                     name: str, 
                     ws: xw.Sheet, 
                     columns: int = 0):
        
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
                
            if name_type == 'range':
                sanitized_name = sanitized_name.replace(" ","_")
            
            return sanitized_name



    def white_list_templates(self):

        """
        Remove mark of the web from the templates on MacOS
        
        """

        # Ensure the templates exist
        if xlsm_file.is_file() and xlsx_file.is_file():

            # Remove the quarantine attribute if it exists
            for file in [xlsx_file, xlsm_file]:
            
                try:
                    subprocess.run(["xattr", "-d", "com.apple.quarantine", 
                                    str(file)],
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                
                except subprocess.CalledProcessError:
                    pass

        else:
            self.env.warn_print("Templates not found.")



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
            


    def group_rows(self,
                   input_range: xw.Range):
        
        """
        Groups a provided range by rows
        
        """
        
        
        if self.env.win:
            input_range.api.Rows.Group()
            
        elif self.env.mac:
            input_range.api.group()
        
        
        
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
        iteration = (iteration + 1) % 19
        multiplier = 13 * iteration


        # Set color for field analysis worksheet
        if self.env.win:
            color = (multiplier) + (multiplier*256)  + (multiplier*256*256)
            ws.api.Tab.Color = color
        
        elif self.env.mac:
            color = ((multiplier), (multiplier), (multiplier))
            ws.api.sheet_tab.color.set((color))




class Logger():

    """
    Class representing a performance logger

    """
    
    def __init__(self) -> None:

        # ------------------------------------------------------------------------------
        # Initialize Performance Logging
        
        now = time.time()
        
        self.start: float = now
        self.last: float = now
        self.exit_msg: str = separator + '\n'

        self.performance_logs: dict[str, list] = defaultdict(list)
        self.section_performance: pd.DataFrame = pd.DataFrame()
        self.config: pd.DataFrame = pd.DataFrame()
        self.env: pd.DataFrame = pd.DataFrame()
        self.total_production_time: float       


    def print_initialization_msg(self, wb: wb):
        
        data = wb.settings.data
        file_name = wb.settings.file_name
        export = wb.settings.export


        # Get data_name for initial text output
        data_name = ''
        if isinstance(data, pd.DataFrame):
            data_name = file_name
        elif isinstance(data, dict) and file_name == 'xleda':
            data_name = list(data.keys())[0]
        elif isinstance(data, (Path, str)):
            path = Path(data).expanduser().resolve()
            data_name = path.stem
        
        
        # if default of xleda is being used for data_name, omit the data name from the output message
        if data_name == 'xleda':
            data_name = ''
        else:
            data_name = f"with {data_name} data"
        
        
        
        # If not exporting, provide a meaningful initial output message
        if not export:
            msg = separator + f"\nStarted preparing an xleda workbook {data_name} at {time.strftime('%H:%M:%S')}\n"
        elif export:
            msg = separator + f"\nStarted preparing an xleda export {data_name} at {time.strftime('%H:%M:%S')}\n\n"
        
        wb.theme.print(msg)
        


    def add_variable_logs(self, wb: wb):
        
        """
        Logs xleda environment and configuration

        """
        
        
        # ---------------------------------------------------------
        # Set vars
        
        settings = wb.settings
        template = wb.template
        
        
        
        # ---------------------------------------------------------
        # Set up envirnment df

        # Relevant keys
        env_keys = ['os', 'env_type', 'excel_version', 'compatible', 'junk_drawer', 
                    'date', 'os_release', 'os_version', 'architecture', 'processor', 
                    'python_version', 'python_implementation', 'console_columns', 'console_lines']
        
        # Remove irrelevant keys
        env_dict = {k:v for k, v in vars(wb.env).copy().items() if k in env_keys}

        
        # Create a dataframe, transpose it, add column names, and save it
        env_df = pd.DataFrame.from_records([env_dict]).T
        env_df = env_df.reset_index()
        env_df.columns = ['Environment Variable', 'Value']
        
        self.env =  env_df
        
        

        # ---------------------------------------------------------
        # Set up config df
               
               

        config = {'file path': str(template.path),
                  'settings path': str(settings.settings_path),
                  'data argument': settings.data_argument,
                  'dataframes': ', '.join([ds.name for ds in wb.datasets]),
                  'plots': ', '.join(settings.plots.keys()),
                  'theme_color': settings.theme_color,
                  'large_report': settings.large_report,
                  'overwrite': settings.overwrite,
                  'open_wb': settings.open_wb,
                  'no_vba': settings.open_wb,
                  'export': settings.export,
                  'debug': settings.debug}

        
        # Convert to dataframe, transpose, set column names, and store
        config_df = pd.DataFrame.from_records([config]).T.astype(str)
        config_df = config_df.reset_index()
        config_df.columns = ['Input Argument', 'Value']
        
        self.config = config_df



    def log(self,
            section: str):

        """
        Logs production performance data

        """

        now = time.time()
        
        
        # Construct log item
        log = {'Production Section': section,
               'Production Time in Seconds': now-self.last}
        

        # Append to log storage
        self.performance_logs['section'].append(log)
            
        # Set last values
        self.last = now
      


    def close(self, wb: wb):

        """
        Closes performance logging by converting logs to dataframes

        """

        # Add variable logs
        self.add_variable_logs(wb=wb)
        
        
        # Close the production timer
        self.total_production_time = self.last - self.start


        # Create section performance df
        self.section_performance = pd.DataFrame.from_records(self.performance_logs['section'])
        
        
        # Add % of Production Time columns
        df = self.section_performance
        df['% of Production Time'] = df['Production Time in Seconds']/self.total_production_time






        
class DataSetParser():
    
    """
    A class that represents a data source parser.
    
    """
    

    def __init__(self, 
                 settings: Settings):
        

        # Save vars to class instance
        self.large_report = settings.large_report
        self.env = settings.env
        
        
        # Create a datasets placeholder
        self.datasets: list[DataSet] = []
        
        
        # Parse data inputs
        self._parse_data_inputs(settings=settings)
        
        
 
    def _parse_data_inputs(self,
                           settings: Settings) -> None:
        
        """
        Validates that a supported data source has been provided

        """
           

        # Set vars
        data = settings.data
        input_df = settings.input_df
        supported = ", ".join(sorted(supported_extensions))


        
        # TODO: Remove placeholder API on 8.18
        
        # Handle neither data argument provided
        if data is None and input_df is None:
        
            raise DataError("No Data Provided")
        
        # Handle 'input_df' argument provided without 'data'
        elif (isinstance(input_df, pd.DataFrame) and data is None):
            
            settings.env.warn_print("The 'input_df' argument has been replaced by 'data'")
            self.from_dataframes({settings.file_name: input_df})
            
        # If both are 'input_df' and  'data' are provided, ignore input_df
        elif (input_df is not None and data is not None):
            
            settings.env.warn_print("The 'input_df' argument has been replaced by 'data', ignoring 'input_df'.")
            
        # Only the data argument has been provided, validate that it is supported
        else:
            
            # if data is a dictionary of dataframes, use it
            if isinstance(data, dict) and all(
                isinstance(k, str) and isinstance(v, pd.DataFrame) for k, v in data.items()):
                
                settings.data_argument = 'Dataframe dictionary'
                self.from_dataframes(data=data)
                
                # if no file_name has been provided, use the first key as the file name
                if settings.file_name == 'xleda':
                    settings.file_name = list(data.keys())[0]
            
            # if data is a dataframe, convert it to a dataframe dictionary and use it
            elif isinstance(data, pd.DataFrame):
                
                settings.data_argument = 'Dataframe'
                self.from_dataframes(data = {settings.file_name: data})

            
            # if data is neither a dataframe or a dataframe dict, ensure it's a supported file
            elif isinstance(data, (str, Path)):
                
                # Get the path
                path = Path(data).expanduser().resolve()
                settings.data_argument = str(path)

                
                # Ensure it's a file
                if not path.is_file():
                    raise DataError(f"\n\nData file not found.\n\nProvided data argument:\n\n    {str(data)}")

                # Since it is a file, ensure that it has supported extension or is an .rdata file
                if not (path.stem.lower() == '.rdata') and (path.suffix.lower() not in supported_extensions):
                    raise DataError(f"\n\nUnsupported file type: {path.suffix}\n\nSupported file types:\n\n    {supported}")
                
                # Since it's a file with a supported extension, get details from it as needed
                else:
                    
                    # If wb_path hasn't been provided, use the data file directory
                    if not settings.wb_path:
                        settings.wb_path = path.parent
                        
                        
                    # if file_name hasn't been provided, use the data file name
                    if settings.file_name == 'xleda' or not settings.file_name:
                        settings.file_name = path.stem
                
                        # Adjust the name if the source file is an excel file to prevent collissions
                        if path.suffix in ['.xlsm', '.xlsx']:
                            settings.file_name += '_xleda'
                    
                    
                    # Since it's a file with a supported extension, use it
                    self.from_file(data=path)
            
            else:
                raise DataError(f"\n\nData file not found.\n\nProvided data argument:\n\n    {str(data)}")
 


    def from_file(self, 
                  data: Path):
        
        """
        Creates a list of DataSet objects from from supported data files

        """    
        
        # ----------------------------------------------------
        # Set vars and handle data file edge cases
        
        # Set vars      
        self.file_path = data
        self.file_name = data.stem
        
        
        # Extract extension and handle .rdata files that don't have an extension
        if self.file_path.suffix.lower() == '.rdata':
            extension = '.rdata'
        else:
            extension = self.file_path.suffix.lower()
            

        # ----------------------------------------------------
        # Create datasets from tabular files
        
        if extension == ".csv":
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
            
        
        elif extension in excel_extension:
            self.from_excel()    
        
            
        elif extension in db_file_extensions:
            self.from_db_file()
            

        elif extension in ['.pkl', '.pickle', '.pck']:
            self.from_pickle()

               


    def identify_db_type(self) -> str:
        
        """
        Identifies the correct database type from files that could be duckdb, sqlite, or unknown

        """
        
        db_type = ""
        
        
        # Open the file and inspect the header for sqlite/duckdb markers
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
        Creates a list of DataSet objects from a supported xml or json files

        """
        
        # Lazy imports
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
            raise DataError("\n\nXML/JSON file not successfully parsed") 
            
        

    def from_pickle(self):
        
        """
        Creates a list of DataSet objects from dataframes inside a pickle file

        """
        

        # Lazy import
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
                                                     
                                                     
        except EOFError:
            pass
            
        except Exception: 
            raise DataError("\n\nPickle file not successfully parsed")
        
        if len(self.datasets)==0:
            raise DataError("\n\nPickle file not successfully parsed")


        

    def from_rdata(self):
        
        """
        Creates a list of DataSet objects from an RData file
        
        """
        
        # Lazy import
        import rdata
        
        try:
            # Parse the file into pure Python/R objects.
            parsed_file = rdata.parser.parse_file(str(self.file_path))
            
            # Convert the entire workspace to its Python equivalents
            converted_data = rdata.conversion.convert(parsed_file)
            
        except Exception as e:

            raise DataError(f"\n\nRData file not successfully parsed: {e}")
        
        
        # Iterate through variables and extract valid DataFrames
        for name, value in converted_data.items():
            try:
                df = None
                
                # Objects that have already been converted to a dataframe
                if isinstance(value, pd.DataFrame):
                    df = value
                    
                # Append successfully extracted dfs
                if df is not None and not df.empty:
                    self.datasets.append(DataSet(input_df=df,
                                                 name=name,
                                                 large_report=self.large_report))
            except Exception:
                
                pass





    def from_db_file(self):
        
        """
        Creates a list of DataSet objects supported database files

        """
        
        # Lazy imports
        import duckdb

        
        db_type = self.identify_db_type()

            
        # Handle unsupported database file types
        if not db_type:
            
            # database file type undetermined
            raise DataError("\n\nDatabase file is not sqlite or duckdb")


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
            raise DataError("\n\nDatabase file not successfully parsed")
                
        finally:
                conn.close()
                
        if len(self.datasets)==0:
            raise DataError("\n\nDatabase file not successfully parsed")
            


    def from_excel(self):
        
        """
        Creates a list of DataSet objects from an Excel file

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
            raise DataError("\n\nExcel file not successfully parsed")
        
        
        if len(self.datasets)==0:
            raise DataError("\n\nExcel file not successfully parsed")


    def from_feather(self):
            
            """
            Creates a list of DataSet objects from a feather file

            """    
            
            try:
                self.datasets.append(DataSet(input_df=pd.read_feather(self.file_path),
                                             name=self.file_name,
                                             large_report=self.large_report))
                    
            except Exception:
                raise DataError("\n\nFeather file not successfully parsed")

    def from_csv(self):
            
            """
            Creates a list of DataSet objects from a feather file

            """    
            
            try:
                self.datasets.append(DataSet(input_df=pd.read_csv(self.file_path),
                                                  name=self.file_name, 
                                                  large_report=self.large_report))
                    
            except Exception:
                raise DataError("\n\nCSV file not successfully parsed")

    def from_parquet(self):
            
            """
            Creates a list of DataSet objects from a feather file

            """
            
            try:
                self.datasets.append(DataSet(input_df=pd.read_parquet(self.file_path),
                                                  name=self.file_name, 
                                                  large_report=self.large_report))
                    
            except Exception:
                raise DataError("\n\nParquet file not successfully parsed")


    def from_dataframes(self, 
                        data: dict[str, pd.DataFrame]):
            
        """
        Creates a list of DataSet objects from a dictionary of dataframes

        """    
        
        assert isinstance(data, dict)
        
        try:
            
            for name, df in data.items():

                self.datasets.append(DataSet(input_df=df,
                                             name=name,
                                             large_report=self.large_report))

                
        except Exception:
            raise DataError("\n\nDataframes not successfully parsed")
        
        
        if len(self.datasets)==0:
            raise DataError("\n\nDataframes not successfully parsed")



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
    

    def update_name(self, 
                    new_name: str):
        
        """
        Updates the dataframe name across the DataSet object
        
        """
        
        self.name = new_name
        self.field_overview['Dataframe'] = new_name
        self.df_overview['Dataframe'] = new_name
        
        
        
    def create_original_df(self,
                           input_df: pd.DataFrame) -> pd.DataFrame:
        
        """
        Creates a near original dataframe that is subsampled if needed

        """
        
        df = input_df.copy()
        rows = len(df)
        columns = len(df.columns)
        

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
        
        
        
        # Subsample if needed
        if self.warning:
            df = (df.iloc[:, :columns].sample(n=rows).sort_index())
        
        # Capture memory usage before promoting index to a column
        self.memory_usage = df.memory_usage(deep=True).sum()
        self.index_mem = df.index.memory_usage()
        
        df['index'] = df.index
        
        self.column_order = df.columns.to_list()
        
        return df

        

    def create_source_data(self) -> pd.DataFrame:
        
        """
        Configures source data by adding EDA workflow columns        

        """
        
        df = self.original_df.copy()
        df.insert(loc=0, column="Record List", value="False")
        df['HasBlank'] = df.isnull().any(axis=1).astype(int)
        df["Record Hash"] = pd.util.hash_pandas_object(df, index=False)
        
        
        return df
        
        

    def create_field_metadata(self) -> pd.DataFrame:

        """
        Produces a field_metadata dataframe

        """


        # --------------------------------------------------
        # Collect metadata

        # Omit added columns
        df = self.source_data.copy().drop(columns=['Record Hash', 'HasBlank', 'Record List'])

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
                "Memory Usage": df.memory_usage(deep=True, index=False),
                "Memory Usage %": df.memory_usage(deep=True, index=False) / self.memory_usage,
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
                "Distinct %": df.nunique() / rows_count}).T

        
        
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
        summary_df = summary_df.loc[row_order, self.column_order]
        
        # Adjust index to it's original memory footprint before becoming a column
        summary_df.at['Memory Usage', 'index'] = self.index_mem
        summary_df.at['Memory Usage %', 'index'] = self.index_mem/self.memory_usage
        

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
                       'Memory Usage (bytes)': self.memory_usage,
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
                       'Memory Usage (bytes)': self.memory_usage,
                       'Rows': self.rows,
                       'Columns': self.columns,
                       'Fields Defined %': '',
                       'Missing %': df.isnull().mean().mean(),
                       'Subsampled': self.warning}
        
        return pd.DataFrame.from_records([df_metadata])
    


class CLI():
    
    
    def __init__(self):
        
        self.windows_menu_name = "Create xleda Workbook"
        self.macos_service_name = "Create xleda Workbook.workflow"
               
        
        self.env = Environment()


        
    def windows_command(self) -> str:

        """
        Constructs the context menu command for Windows
        
        """
        
        python = str(Path(sys.executable).resolve())
        module_command = f"& {shlex.quote(python)} -m xleda wb '%1'"
        pause_command = "Write-Host ''; Read-Host 'Operation Completed, you can now close this window'"
        return f'powershell.exe -NoExit -ExecutionPolicy Bypass -Command "{module_command}; {pause_command}"'


    def install_windows_context_menu(self) -> bool:
        
        """
        Installs the context menu on Windows
        
        Returns
        -------
        bool
            A boolean indicating success

        """
        
        import winreg
        command = self.windows_command()
        
        try:

            icon_path = Path(__file__).parent / 'rectangle_icon.ico'
            
            for extension in supported_extensions:
                base_key = rf"Software\Classes\SystemFileAssociations\{extension}\shell\xleda"
                
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base_key) as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, self.windows_menu_name)
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, str(icon_path))

                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base_key}\command") as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)
                
            return True
        
        except Exception:
            
            return False
        


    def uninstall_windows_context_menu(self) -> bool:
        
        """
        Uninstalls the context menu on Windows
        
        Returns
        -------
        bool
            A boolean indicating success

        """

        
        try:
            import winreg

            for extension in supported_extensions:
                base_key = rf"Software\Classes\SystemFileAssociations\{extension}\shell\xleda"
                try:
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, rf"{base_key}\command")
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, base_key)
                except FileNotFoundError:
                    pass
        
            return True
        
        except Exception:
            return False



    def macos_workflow_shell_script(self) -> str:
        
        """
        Constructs the context menu command for MacOS
        
        """

        python = shlex.quote(str(Path(sys.executable).resolve()))
        
        cmd = ('for data_file_path in "$@"'
               'do'
             """  /usr/bin/osascript - "$filePath" <<'APPLESCRIPT'"""
               'on run argv'
               '  set filePath to item 1 of argv'
              f'  set commandText to "{python} -m xleda wb " & quoted form of filePath & "; echo; read -n 1 -s -r -p " & quoted form of "Press any key to close this window..."'
               '  tell application "Terminal"'
               '    activate'
               '    do script commandText'
               '  end tell'
               'end run'
               'APPLESCRIPT'
               'done')
        
        return cmd


    def macos_workflow_document(self) -> dict:
        
        """
        Constructs the context menu workflow for MacOS
        
        """
        
        return {
            "AMApplicationBuild": "523",
            "AMApplicationVersion": "2.10",
            "AMDocumentVersion": "2",
            "actions": [
                {
                    "action": {
                        "AMAccepts": {
                            "Container": "List",
                            "Optional": False,
                            "Types": ["com.apple.cocoa.path"],
                        },
                        "AMActionVersion": "2.0.3",
                        "AMApplication": ["Automator"],
                        "AMParameterProperties": {},
                        "AMProvides": {
                            "Container": "List",
                            "Types": ["com.apple.cocoa.path"],
                        },
                        "ActionBundlePath": "/System/Library/Automator/Run Shell Script.action",
                        "ActionName": "Run Shell Script",
                        "ActionParameters": {
                            "COMMAND_STRING": self.macos_workflow_shell_script(),
                            "CheckedForUserDefaultShell": True,
                            "inputMethod": 1,
                            "shell": "/bin/zsh",
                            "source": "",
                        },
                        "BundleIdentifier": "com.apple.RunShellScript",
                        "CFBundleVersion": "2.0.3",
                    },
                    "isViewVisible": True,
                }
            ],
            "connectors": {},
            "workflowMetaData": {
                "applicationBundleIDsByPath": {"/System/Library/CoreServices/Finder.app": "com.apple.finder"},
                "applicationPaths": ["/System/Library/CoreServices/Finder.app"],
                "inputTypeIdentifier": "com.apple.Automator.fileSystemObject",
                "outputTypeIdentifier": "com.apple.Automator.nothing",
                "presentationMode": 15,
                "processesInput": True,
                "serviceApplicationBundleID": "com.apple.finder",
                "serviceApplicationPath": "/System/Library/CoreServices/Finder.app",
                "serviceInputTypeIdentifier": "com.apple.Automator.fileSystemObject",
                "serviceOutputTypeIdentifier": "com.apple.Automator.nothing",
                "serviceProcessesInput": True,
            },
        }




    def install_macos_context_menu(self) -> bool:

        """
        Installs the context menu on MacOS
        
        Returns
        -------
        bool
            A boolean indicating success

        """
        
        import plistlib
        
        service_path = Path.home() / "Library" / "Services" / self.macos_service_name
        contents_path = service_path / "Contents"
        
        
        try:
            contents_path.mkdir(parents=True, exist_ok=True)

            info = {
                "CFBundleIdentifier": "com.infodesigner.xleda.create-workbook",
                "CFBundleName": "Create xleda Workbook",
                "CFBundlePackageType": "FMWK",
                "NSServices": [
                    {
                        "NSMenuItem": {"default": "Create xleda Workbook"},
                        "NSMessage": "runWorkflowAsService",
                        "NSRequiredContext": {"NSApplicationIdentifier": "com.apple.finder"},
                        "NSSendFileTypes": [
                            
                            # Text Data Formats
                            "public.comma-separated-values-text",  # .csv
                            "public.json",                        # .json
                            "public.xml",                         # .xml
                            
                            # Excel Formats
                            "org.openxmlformats.spreadsheetml.sheet",               # .xlsx
                            "com.microsoft.excel.xls",                              # .xls
                            "org.openxmlformats.spreadsheetml.sheet.macroenabled",  # .xlsm
                            "com.microsoft.excel.sheet.binary.macroenabled",        # .xlsb
                            
                            # Databases
                            "org.sqlite.sqlite3",                 # .sqlite, .sqlite3
                            "public.database",                    # .db, .db3, .s3db, .sl3
                            
                            # Big Data & Custom Catch-all
                            "public.data"                         # .parquet, .feather, .duckdb, .ddb, .rdata, '.pkl', '.pickle', '.pck'
                        ],
                        "NSSendTypes": ["NSFilenamesPboardType"],
                    }
                ],
            }
            with (contents_path / "Info.plist").open("wb") as f:
                plistlib.dump(info, f)

            with (contents_path / "document.wflow").open("wb") as f:
                plistlib.dump(self.macos_workflow_document(), f)

            subprocess.run(["/System/Library/CoreServices/pbs"], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            return True
        
        except Exception:
            return False


    def uninstall_macos_context_menu(self) -> bool:

        """
        Uninstalls the context menu on MacOS
        
        Returns
        -------
        bool
            A boolean indicating success

        """

        try:
            service_path = Path.home() / "Library" / "Services" / self.macos_service_name
            if service_path.exists():
                import shutil

                shutil.rmtree(service_path)

            subprocess.run(["/System/Library/CoreServices/pbs"], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            return True
        
        except Exception:
            
            return False


    def install(self) -> None:
        
        """
        Installs the right-click context menu in either MacOS/Windows
        
        """

        if self.env.win:
            success = self.install_windows_context_menu()       
            
        elif self.env.mac:
            success = self.install_macos_context_menu()
        
        if success:

            
            success_message = (f"{separator}\033[1m\n\nInstalled the xleda right-click menu\033[0m\n\n"
                               "Supported file types:\nCSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML\n\n"
                               f"Expected extensions:\n{supported_extensions}\n\n"
                               "For more documentation, visit https://github.com/InfoDesigner/xleda")
                    
            print(success_message)

    def uninstall(self) -> None:
        
        """
        Uninstalls the right-click context menu in either MacOS/Windows
        
        """

        if self.env.win:
            success = self.uninstall_windows_context_menu()
        elif self.env.mac:
            success = self.uninstall_macos_context_menu()
        
        if success:
            print("Uninstalled the xleda right-click menu.")
            
    
    def version(self):
        
        """
        Checks for an updated version on PyPi
        
        """
        
        # Validate compatibility, get persistent settings
        env = Environment()
        settings = Settings(env=env, version=True)
        
        # Use theme color setting to return version information
        theme = Theme(settings=settings)
        theme.print(settings.update_msg)



class TemplateError(Exception):
    
    """
    An exception class for capturing template parsing errors

    """
    
    def __init__(self, 
                 message: str):
        
        super().__init__(message)


class CompatibilityError(Exception):
    
    """
    An exception class for capturing compatibility errors

    """
    
    def __init__(self, 
                 message: str):
        
        super().__init__(message)


class DataError(Exception):
    
    """
    An exception class for capturing file parsing errors

    """
    
    def __init__(self, 
                 message: str):
        
        super().__init__(message)