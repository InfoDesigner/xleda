from __future__ import annotations

import json
import subprocess
import urllib.request
from importlib.metadata import version as package_version
from pathlib import Path
import shutil
from importlib.resources import files, as_file
import zipfile


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
    


if settings.env.win:
    import winreg


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
        Constructs a context menu command for Windows to run xleda on Windows
        
        """
        settings_json = r"$env:USERPROFILE\AppData\Roaming\.xleda\settings.json"
        fallback_py = "python"
        
        # Constructs a powershell command to pull the last python interpreter that xleda 
        # was run with from persistent settings
        ps_cmd = (
            f"$s = '{settings_json}'; "
            f"$p = '{fallback_py}'; "
            "if (Test-Path $s) { "
                "$j = Get-Content $s | ConvertFrom-Json; "
                "if ($j.python_executable) { $p = $j.python_executable } "
            "}; "
            
            "Start-Process $p -ArgumentList '-m', 'xleda', 'wb', $args[0] -NoNewWindow -Wait"
        )
        
        # '%1' is passed outside of the -Command string block to care for files with spaces or special characters
        return f'powershell.exe -NoExit -ExecutionPolicy Bypass -Command "{ps_cmd}" -- "%1"'



    def install_windows_context_menu(self) -> bool:
        
        """
        Installs the icon and context menu on Windows
        
        """
        command = self.create_windows_context_menu_command()
        
        try:
            
            # Contruct pesistent settings directory if needed
            appdata_dir = Path.home() / 'AppData' / 'Roaming' / '.xleda'
            appdata_dir.mkdir(parents=True, exist_ok=True)
            
            source_icon = files('xleda').joinpath('assets', 'rectangle_icon.ico')
            dest_icon = appdata_dir / 'rectangle_icon.ico'
            
            
            # Copy icon to the pesistent settings folder
            with as_file(source_icon) as src_icon:
                shutil.copy2(src_icon, dest_icon)
            
            for extension in supported_extensions:
                base_key = rf"Software\Classes\SystemFileAssociations\{extension}\shell\xleda"
                
                # Creates a key for each extension
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base_key) as key:   # type: ignore
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, self.windows_menu_name)  # type: ignore
                    
                    # Add icon from persistent settings to the key
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, str(dest_icon))  # type: ignore
                
                # Creates a key for the xleda command
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base_key}\command") as key:  # type: ignore
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)  # type: ignore
                    
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
            
            # Adds xleda for supported file types
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
        Installs a context menu Quick Service into macOS
        
        """
        
        try:
        
            # Define paths and names
            zip_name = "Create xleda Workbook.workflow.zip"
            destination_dir = Path.home() / "Library" / "Services"
            destination_workflow = destination_dir / "Create xleda Workbook.workflow"

            # Locate the workflow archive using importlib
            source_zip_resource = files("xleda").joinpath("assets", zip_name)
            
            # If the workflow archive doesn't exist, raise an error
            if not source_zip_resource.exists(): # type: ignore
                print(f"Error: Internal asset archive '{zip_name}' is missing.")
                return False

            # Perform a clean reinstall
            if destination_workflow.exists():
                shutil.rmtree(destination_workflow)

            # Install the .workflow
            with as_file(source_zip_resource) as zip_disk_path:
                with zipfile.ZipFile(zip_disk_path, 'r') as zip_ref:
                    
                    # Extracts the contents directly into ~/Library/Services/
                    zip_ref.extractall(destination_dir)

            # Flush the system Pasteboard registries to activate the quick action entry
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
            
            success_message = (f"{separator}\n\nInstalled the xleda right-click menu\n\n"
                               "Supported file types:\nCSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML\n\n"
                               f"Expected extensions:\n{supported}\n\n"
                               "For more documentation, visit https://github.com/InfoDesigner/xleda\n")
                    
            settings.logger.print(success_message)


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
