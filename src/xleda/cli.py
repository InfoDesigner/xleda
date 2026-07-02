from __future__ import annotations

import json
import shlex
import subprocess
import sys
import urllib.request
from importlib.metadata import version as package_version
from pathlib import Path

from packaging.version import parse

from .utilities import Logger, Settings, separator, supported, Environment, supported_extensions
from .main import wb

# -------------------------------------------------
# Configure CLI

import typer
from typer import rich_utils

color = Settings(Logger()).theme.color

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

    @staticmethod
    def help_message() -> str:
        color = Settings(Logger()).theme.color
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
            f"[{color}]Example Command:[/]\n\n"
            f"[{color}]    xleda wb 'https://github.com/InfoDesigner/xleda/blob/main/examples/data/penguins_raw.csv'[/]\n\n"
            f"[{color}] [/]\n\n"
            f"[{color}]wb Help Command:[/]\n\n"
            f"[{color}]    xleda wb --help[/]\n\n"
            f"[{color}] [/]\n\n\n"
            f"[{color}]For more documentation, visit https://github.com/InfoDesigner/xleda[/]"
        )

    def __init__(self):

        self.windows_menu_name = "Create xleda Workbook"
        self.macos_service_name = "Create xleda Workbook.workflow"

        self.env = Environment()

    def wb(self, **kwargs):
        """Create a workbook via the underlying workbook class."""
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

cli = typer.Typer(epilog=CLI.help_message(), rich_markup_mode="rich", no_args_is_help=True)


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


@cli.command(name="wb", epilog=CLI.help_message())
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
