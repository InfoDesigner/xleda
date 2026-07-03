from __future__ import annotations

import json
import shlex
import subprocess
import sys
import urllib.request
from importlib.metadata import version as package_version
from pathlib import Path
import shutil


from packaging.version import parse

from .utilities import Logger, Settings, separator, supported, Environment, supported_extensions
from .main import wb

# -------------------------------------------------
# Configure CLI

import typer
from typer import rich_utils

logger = Logger()
settings = Settings(logger=logger)
color = settings.theme.color
    

# Apply Hex colors for typer app
rich_utils.STYLE_OPTION = color
rich_utils.STYLE_HELPTEXT = color
rich_utils.STYLE_OPTION_HELP = color
rich_utils.STYLE_OPTION_DEFAULT = color
rich_utils.STYLE_OPTION_ENVVAR = color
rich_utils.STYLE_REQUIRED_SHORT = color
rich_utils.STYLE_REQUIRED_LONG = color
rich_utils.STYLE_OPTIONS_PANEL_BORDER = color
rich_utils.STYLE_OPTIONS_TABLE_BOX = color
rich_utils.STYLE_OPTIONS_TABLE_BORDER_STYLE = color
rich_utils.STYLE_COMMANDS_PANEL_BORDER = color
rich_utils.STYLE_COMMANDS_TABLE_FIRST_COLUMN = color
rich_utils.STYLE_METAVAR_SEPARATOR = color
rich_utils.STYLE_USAGE = color
rich_utils.STYLE_USAGE_COMMAND = color




        
# -------------------------------------------------
# CLI

class CLI():
    
    # A class representing a CLI application
    
    @staticmethod
    def help_message() -> str:

        return (
            f"[{color}]{separator}[/]\n\n\n\n"
            f"[{color}]Use 'xleda wb <data file path>' to create a workbook[/]\n\n"
            f"[{color}] [/]\n\n"
            f"[{color}]Supported file types:[/]\n\n"
            f"[{color}]    CSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML[/]\n\n\n"
            f"[{color}] [/]\n\n"
            f"[{color}]Expected extensions:[/]\n\n"
            f"[{color}]    {supported}[/]\n\n\n"
            f"[{color}] [/]\n\n"
            f"[{color}]Create an example workbook:[/]\n\n"
            f"[{color}]    xleda wb 'https://github.com/InfoDesigner/xleda/blob/main/examples/data/penguins_raw.csv'[/]\n\n"
            f"[{color}] [/]\n\n"
            f"[{color}]Set your theme without creating a workbook:[/]\n\n"
            f"[{color}]    xleda theme '#262626'[/]\n\n"
            f"[{color}] [/]\n\n"
            f"[{color}]Toggle your workbook style to use or not use VBA:[/]\n\n"
            f"[{color}]    xleda vba[/]\n\n"
            f"[{color}] [/]\n\n"
            f"[{color}]wb help command:[/]\n\n"
            f"[{color}]    xleda wb --help[/]\n\n"
            f"[{color}] [/]\n\n\n"
            f"[{color}]For more documentation, visit https://github.com/InfoDesigner/xleda[/]"
        )
    
    
    

    def __init__(self):

        self.windows_menu_name = "Create xleda Workbook"
        self.macos_service_name = "Create xleda Workbook.workflow"

        self.env = Environment()
        

    def wb(self, **kwargs):
        
        """
        Create a workbook via the underlying workbook class
        
        """
        return wb(**kwargs)
    
    
    

    def create_windows_context_menu_command(self) -> str:

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
        command = self.create_windows_context_menu_command()
        
        try:

            icon_path = Path(__file__).parent / 'rectangle_icon.ico'
            
            for extension in supported_extensions:
                base_key = rf"Software\Classes\SystemFileAssociations\{extension}\shell\xleda"
                
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base_key) as key: # type: ignore
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, self.windows_menu_name) # type: ignore
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, str(icon_path)) # type: ignore

                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base_key}\command") as key: # type: ignore
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command) # type: ignore
                
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
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, rf"{base_key}\command") # type: ignore
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, base_key) # type: ignore
                except FileNotFoundError:
                    pass
        
            return True
        
        except Exception:
            return False

    def install_macos_context_menu(self) -> bool:
        """
        Installs a context menu Quick Service into MacOS
        
        """
        try:
            workflow_name = "Create xleda Workbook.workflow"
            source_workflow_path = Path(__file__).parent / workflow_name
            
            if not source_workflow_path.exists():
                print(f"Error: Internal asset framework '{workflow_name}' missing.")
                return False
                
            destination_path = Path.home() / "Library" / "Services" / workflow_name
            
            if destination_path.exists():
                shutil.rmtree(destination_path)
                
            # Copies the folder structure, plist properties, and your custom png image natively!
            shutil.copytree(source_workflow_path, destination_path)
            
            # Flush the system Pasteboard background registries to activate the menu item row instantly
            subprocess.run(["/System/Library/CoreServices/pbs", "-flush"], capture_output=True, check=False)
            return True
        except Exception as e:
            print(f"Installation failed: {e}")
            return False

    def uninstall_macos_context_menu(self) -> bool:
        """
        Uninstalls a context menu Quick Service into MacOS
        
        """

        try:
            service_path = Path.home() / "Library" / "Services" / "Create xleda Workbook.workflow"
            
            # Physically erase the folder layout from disk if it exists
            if service_path.exists():
                shutil.rmtree(service_path)
                
            # Re-flush the background system services daemon to sync the menu instantly
            subprocess.run(["/System/Library/CoreServices/pbs", "-flush"])
            return True
        except Exception as e:
            print(f"Uninstallation failed: {e}")
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
                               f"Expected extensions:\n{supported}\n\n"
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
        
        # Set Var
        msg = ""
        
        # Get versions to compare
        installed_version = package_version('xleda')
        
        
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

            msg += separator + "\n\n✨ A newer version of xleda is available\n\n"
            msg += f"   Installed version: {installed_version}\n"
            msg += f"   Latest PyPI version: {latest_ver}\n" + separator
            
            
        else:

            # Store the update version message
            msg += separator + "\n\n✨ You are running the latest version of xleda\n\n"
            msg += f"   Installed version: {installed_version}\n"
            msg += f"   Latest PyPI version: {latest_ver}\n" + separator
        
        
        # Print version output in theme color
        logger = Logger()
        settings = Settings(logger=logger)  # noqa: F841
        logger.print(msg)


help_message = str(CLI.help_message())

cli = typer.Typer(epilog=help_message, rich_markup_mode="rich", no_args_is_help=True)


@cli.command()
def install():
    """Installs the right-click context menu."""
    CLI().install()



@cli.command()
def uninstall():
    """Uninstalls the right-click context menu."""
    CLI().uninstall()



@cli.command()
def version():
    """Checks for an updated version on PyPI."""
    CLI().version()
    


@cli.command()
def theme(color: str = typer.Argument(None, help="Change your theme preference for future workbooks")):
    
    """
    Change your theme preference without creating a workbook
    
    """


    # If no color was provided, show a helpful message and exit
    if not color:
        logger.print("No color provided. Usage: xleda theme '#305CDE' or: xleda theme 305CDE (quote the value on PowerShell).")
        return

    # Ensure color starts with a '#'
    if not color.startswith('#'):
        color = f"#{color}"

    # Basic validation for a hex color (#RGB or #RRGGBB)
    import re
    if not re.match(r"^#(?:[0-9a-fA-F]{3}){1,2}$", color):
        logger.print(f"Invalid color '{color}'. Provide a hex color like '#305CDE'.")
        return

    # Pass the theme value correctly to Settings (use the 'theme' key)
    settings = Settings(logger=logger, locals={'theme': color})

    logger.print(f"Theme preference has been changed to {settings.theme.color}")



@cli.command()
def vba():
    
    """
    Toggles your VBA preference without creating a workbook
        
    """
    
    logger = Logger()
    settings = Settings(locals={'vba_toggle': True},
                        logger=logger)
    
    logger.print(f"Future workbooks will be created {'without' if settings.no_vba else 'with'} VBA")
    


@cli.command(name="wb", epilog=help_message)
def cli_wb(data: str = typer.Argument(..., help="Path to a supported data file"),
           file_name: str = typer.Option(None, "--file_name", show_default=False, help="Name of the created workbook. Defaults to the same name as the data file"),
           theme: str = typer.Option(None, "--theme", help="Hex color used for theme in workbook. This setting will persist after using.  Defaults to a neutral color"),
           export: bool = typer.Option(False, help="Export from an xleda workbook"),
           large_report: bool = typer.Option(False, "--large_report", help="Only subsample when required to fit within Excel's worksheet limits"),
           overwrite: bool = typer.Option(False, help="Overwrite existing workbook"),
           wb_path: str = typer.Option('', "--wb_path", show_default=False, help="Workbook directory with/without filename"),
           open_wb: bool | None = typer.Option(True, "--open_wb/--no_open_wb", help="Don't automatically open the workbook on finish"),
           no_vba: bool | None = typer.Option(None, "--vba/--no_vba", help="Create a VBA-free xlsx file.  This setting will persist after using.  Defaults to False"),
           debug: bool = typer.Option(False, "--debug", help="View the workbook while it's being created")):
    
    """
    Create an xleda workbook from a supported data file.
    
    """
    cli_args = locals()

    if not wb_path:
        cli_args["file_name"] = file_name

    return CLI().wb(**cli_args)
