from __future__ import annotations


import shlex
import subprocess
import sys
from pathlib import Path



# from xleda import wb
from xleda.utilities import supported_extensions, Environment

# Validate compatibility
env = Environment()


windows_menu_name = "Create xleda Workbook"
macos_service_name = "Create xleda Workbook.workflow"

separator = "\n" + ("-" * 100)


    
def windows_command() -> str:


    """
    Constructs the context menu command for Windows
    
    """
    
    python = str(Path(sys.executable).resolve())
    module_command = f"& {shlex.quote(python)} -m xleda wb '%1'"
    pause_command = "Write-Host ''; Read-Host 'Operation Completed, you can now close this window'"
    return f'powershell.exe -NoExit -ExecutionPolicy Bypass -Command "{module_command}; {pause_command}"'


def install_windows_context_menu() -> bool:
    
    """
    Installs the context menu on Windows
    
    Returns
    -------
    bool
        A boolean indicating success

    """
    
    import winreg
    command = windows_command()
    
    try:

        icon_path = Path(__file__).parent / 'rectangle_icon.ico'
        
        for extension in supported_extensions:
            base_key = rf"Software\Classes\SystemFileAssociations\{extension}\shell\xleda"
            
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base_key) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, windows_menu_name)
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, str(icon_path))

            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base_key}\command") as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)
               
        return True
    
    except Exception:
        
        return False
    


def uninstall_windows_context_menu() -> bool:
    
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



def macos_workflow_shell_script() -> str:
    
    """
    Constructs the context menu command for MacOS
    
    """
    
    
    python = shlex.quote(str(Path(sys.executable).resolve()))
    return f"""for data_file_path in "$@"
do
  /usr/bin/osascript - "$filePath" <<'APPLESCRIPT'
on run argv
  set filePath to item 1 of argv
  set commandText to "{python} -m xleda wb " & quoted form of filePath & "; echo; read -n 1 -s -r -p " & quoted form of "Press any key to close this window..."
  tell application "Terminal"
    activate
    do script commandText
  end tell
end run
APPLESCRIPT
done
"""


def macos_workflow_document() -> dict:
    
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
                        "COMMAND_STRING": macos_workflow_shell_script(),
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




def install_macos_context_menu() -> bool:

    """
    Installs the context menu on MacOS
    
    Returns
    -------
    bool
        A boolean indicating success

    """
    
    import plistlib
    
    service_path = Path.home() / "Library" / "Services" / macos_service_name
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
            plistlib.dump(macos_workflow_document(), f)

        subprocess.run(["/System/Library/CoreServices/pbs"], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        return True
    
    except Exception:
        return False


def uninstall_macos_context_menu() -> bool:

    """
    Uninstalls the context menu on MacOS
    
    Returns
    -------
    bool
        A boolean indicating success

    """

    try:
        service_path = Path.home() / "Library" / "Services" / macos_service_name
        if service_path.exists():
            import shutil

            shutil.rmtree(service_path)

        subprocess.run(["/System/Library/CoreServices/pbs"], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        return True
    
    except Exception:
        
        return False


def install() -> None:
    
    """
    Installs the right-click context menu in either MacOS/Windows
    
    """

    if env.win:
        success = install_windows_context_menu()       
        
    elif env.mac:
        success = install_macos_context_menu()
    
    if success:
        
        
        print(separator + "\033[1m" + "\n\nInstalled the xleda right-click menu" + "\033[0m" + "\n\n"
              "Create xleda worbooks by right-clicking supported data file types\n\n"
              "Use 'xleda --help'"
              f"Supported file types:\nCSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML\n\n"
              f"Expected extensions:\n{supported_extensions}\n" + separator)
        

def uninstall() -> None:
    
    """
    Uninstalls the right-click context menu in either MacOS/Windows
    
    """

    if env.win:
        success = uninstall_windows_context_menu()
    elif env.mac:
        success = uninstall_macos_context_menu()
    
    if not success:
        print("Uninstalled the xleda right-click menu.")