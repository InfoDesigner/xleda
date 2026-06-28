import pandas as pd
from pathlib import Path
import ast
import time
import typer
from tqdm.auto import tqdm

from matplotlib.figure import Figure
import xlwings as xw

from .utilities import (Environment, Config, Theme, Plotter,
                        Blueprint, ExportDict, PerformanceLogger)

from .utilities import (Environment, Template, Theme, Plotter, DataSet, DataSetParser, ExportDict, 
                        Logger, CLI, Settings, help_message, TemplateError)


# -------------------------------------------------
# Set primary vars/config


separator = "\n" + ("-" * 100)



class wb():

    """
        A class that represents an xleda workbook.

    """

    def __init__(self, 
                 input_df: pd.DataFrame, 
                 name: str = 'xleda', 
                 theme_color: str = "#262626", 
                 add_plots: dict[str, Figure] = {},
                 add_dfs: dict[str, pd.DataFrame] = {},
                 large_report: bool = False, 
                 overwrite: bool = False, 
                 wb_path: Path | str = Path().cwd(),
                 open_wb: bool = True,
                 no_vba: bool = False,
                 export: bool = False,
                 debug: bool = False,
                 ) -> None:

        """
        Creates an xleda workbook

        Parameters
        ----------

        input_df : pd.DataFrame
            * A pandas dataframe of any size.  
            * Will create an xleda workbook that is 25,000 rows/50 columns by default.  
            * Use large_report=True to expand this to Excel's limits.

        name : str
            * Name of the workbook to be created.

        theme_color : str, optional
            * A hexidecimal color used for charts/accent color.  
            * Use theme_color='random' for random colors
            * Defaults to "#262626"

        add_plots : dict[str, Figure], optional
            * Additional plots to be included 
            * Uses "{'plot1_name': Plot1Figure, 'plot2_name': Plot2Figure, ...}" format.  
            * Each entry will get it's own worksheet.  
            * No resizing or syling is done for plots added this way
            * Defaults to None

        add_dfs : dict[str, pd.Dataframe], optional
            * Additional dataframes to be included
            * Creates Field Analyis/Overview worksheets for each additional df
            * Uses "{'df1_name': df1, 'df2_name': df2, ...}" format.  
            * Each df will get it's own entry in export_dicts.
            * Defaults to None

        large_report : bool, optional
            * Used to override default limits of 25,000 rows/50 columns. 
            * Sets limits to 1,000,000 rows/16,000 columns
            * Defaults to False

        overwrite : bool, optional
            * Whether to overwrite existing reports with the same name
            * Defaults to False

        wb_path : Path | string, optional
            * String or Pathlib path of a directory or file
            * If a directory is provided, the workbook will be created in that directory.
            * If a file name ending in ".xlsm" or ".xlsx" is provided, 
                it will either create that file or export from that file, 
                depending on whether export=True is also selected.
            * Defaults to current working directory

        no_vba : bool, optional
            * Will create the workbook as an xlsx file that has no VBA.
            * Defaults to False
            
        open_wb: bool, optional
            * Whether to open the workbook after creating.  
            * Set to False if creating multiple workbooks.
            * Defaults to True

        export: bool, optional
            * Exports data from an xleda workbook instead of creating one.
            * Exported data is available through wb().export_dict
            * Defaults to False
            * Export includes the follwing fields for each provided dataframe

                * `description`: Dataframe description if you've added one
                * `definitions`: Any field definitions you've added.
                * `notes`: Any field notes you've added
                * `lists`: Any lists showing in the compiled lists section
                * 'field_metadata': A basic metadata dataframe, combining information from 
                                    pandas info/describe/quantile.
                * 'overview_metadata': A transposed copy of the field_metadata.
                * `source_data`: A copy of your unaltered source data that includes 
                                `Record Hash`/`Record List`/`HasBlank` columns.
                * `altered_source_data`: source data from the workbook that includes 
                                         any manual edits you've made such as removing 
                                         records, renaming fields, etc. 
                                         
                                         ** Note that data types will likely change in the round-trip translation. **

        debug: bool, optional
            * Shows the workbook being created                          

        """

        # ------------------------------------------------------------------------------
        # Create Base Components


        # Initialize logger
        self.logger = Logger()
        
        
        # Initialize/Check Environment
        self.env = Environment()
        

        # Create/update settings
        self.settings = Settings(locals=locals(),
                                 env=self.env)


        # Intialize theme
        self.theme = Theme(settings=self.settings)
        
        
        # Intialize plotter
        self.plotter = Plotter(settings=self.settings)

        
        # ------------------------------------------------------------------------------
        # Initial Output, Prepare Datasets/Template
        
        
        # Print initial text/progress bar
        self.logger.print_initialization_msg(wb=self)
        pbar = self.theme.create_progress_bar(desc="Initializing wb Components...", total=10)
        
        # Prepare datasets
        self.datasets: list[DataSet] = DataSetParser(settings=self.settings).datasets
        pbar.update(2) # 2
        

        # Add initial export dicts
        self.export_dicts: list[ExportDict] = [ExportDict(ds) for ds in self.datasets]
        pbar.update(2) # 4


        # Create Template
        self.template = Template(wb=self)
        pbar.update(2) # 6
      
        

        # ------------------------------------------------------------------------------
        # Either create a workbook or export

        if export:
            self._export_analysis(progress_bar=pbar)
        else:            
            self._assemble_wb(progress_bar=pbar)



    def _assemble_wb(self, progress_bar: tqdm):
        
        """
        Assembles an xleda workbook

        """
        
        # -------------------------------------------------
        # Set vars
        

        template = self.template
        datasets = self.datasets
        plots = self.settings.plots
        debug = self.settings.debug
        logger = self.logger
        theme = self.theme

        num_datasets = len(datasets)
        num_columns = sum([ds.columns for ds in datasets])
        num_plots = len(plots)
        
        progress_bar.update(2) # 8



        # --------------------------------------------------
        # Open/Prepare the Template with a Context Manager

        with xw.App(visible=debug, add_book=False) as app:


            # Set vars, open workbook
            book = app.books.open(template.path, read_only=False)
            template.add_book(book=book)
            
            # Close progress bar/Add Init Log
            progress_bar.update(2) # 10
            progress_bar.close()
            self.logger.log(section='Initializing wb Components')
            
            

            app.display_alerts = debug
            app.screen_updating = debug
            
            # --------------------------------------------------
            # Adding Data
            
            with theme.create_progress_bar(desc="Adding Data...",
                                           total=5 + (6 * num_datasets)) as pbar:
            

                # Create placeholder worksheets
                self.template.add_worksheets(progress_bar=pbar) # length of datasets + 1


                # Add field analysis worksheets
                template.add_field_analyses(progress_bar=pbar) # length of datasets * 5
                
                
                # Add overview
                template.add_overview(progress_bar=pbar) # 4
                
                # Close section log
                logger.log(section='Adding Data')



            # --------------------------------------------------
            # Adding Plots
            
            # Once for each dataset/plot/column            
            with theme.create_progress_bar(desc="Adding Plots...",
                                           total=num_datasets + num_plots + num_columns) as pbar:
            
                # Adds additional plots
                if plots:
                    template.add_plot_sheets(progress_bar=pbar)

                # Add plots for all fields
                template.add_plots(progress_bar=pbar)
                
                
                logger.log(section ='Adding Plots')



            # --------------------------------------------------
            # Clean up and close logs

            
            # Close logs
            logger.close(wb=self)

            # Add logging to workbook
            template.add_debug()

            # Restore app settings
            app.display_alerts = True
            app.screen_updating = True
                        
            # Save/exit context manager
            book.save(template.path)
            
            # Wait 2 seconds to ensure it saves to disk fully before exiting context manager
            time.sleep(2)
            
            # Provide an output message
            theme.print(f"\nProcess completed in {int(self.logger.total_production_time)} seconds")
            

        # --------------------------------------------------
        # Reopen the workbook in a new context manager if needed
        
        
        if self.settings.open_wb:
            
            app = xw.App(visible=True, add_book=False) 
            book = app.books.open(template.path)



        # ----------------------------------------------------------------------------------
        # Print closing messsage
        

        exit_msg = self.logger.exit_msg
        

        # Note dataframes          
        dataframes = ", ".join([ds.name for ds in datasets])
        exit_msg += f"\nDataframes included:\n    {dataframes}"

        # Note plots
        if plots:
            
            plots = (", ").join(plots.keys())
            exit_msg += f"\n\nAdditional plots included:\n    {plots}"

        exit_msg += f"\n\nWorkboook is located at:\n {self.template.path}"

        # Print closing message
        self.theme.print(exit_msg + '\n' + separator)
    



    def _export_analysis(self, progress_bar: tqdm):

        """
        Exports data from an xleda workbook into self.export_dicts

        """
        
        # -------------------------------------------------
        # Set vars
        
        template = self.template
        settings = self.settings
        theme = self.theme
        datasets = self.datasets
        env = self.env
        logger = self.logger

        missing_dfs = []



        # --------------------------------------------------
        # If the file is not found return messaging

        if not template.path.is_file():

            msg = f"File not found at {template.path}\n"
            msg +=  "wb().export_dicts will be limited" + separator
            
            raise TemplateError(msg)
        
        progress_bar.update(2) # 8


        # --------------------------------------------------
        # Print starting output


        with xw.App(visible=settings.debug, add_book=False) as app:
            
            app.display_alerts = False
            app.screen_updating = settings.debug
                        
            book = app.books.open(template.path, read_only=False)
            
            # Close the initial progress bar
            progress_bar.update(2) # 10
            progress_bar.close()
            


            with theme.create_progress_bar(desc="Reading workbook...",
                                           total=4 + (2 * len(datasets))) as pbar:
                              
                pbar.update(2)
                wb = app.books.open(template.path)
                pbar.update(2)

                
                for ds in datasets:

                    definitions = {}
                    notes = {}
                    lists = {}



                    
                    if bp.field_analysis in wb.sheet_names:
                        ws = wb.sheets(bp.field_analysis)
                    else:

                        missing_dfs.append(bp.name)
                        
                        continue
                   


                    # --------------------------------------------------
                    # Export lists/notes/definitions from workbook


                    # Definitions/Notes/Fields
                    field_notes = ws.range("Notes").value
                    field_definitions = ws.range("Definitions").value
                    fields = ws.range("FieldRange").value

                    # Compiled Lists
                    compiled_lists_names = ws.range("Compiled_Lists").value
                    compiled_lists = ws.range("Compiled_Lists").offset(0, 1).value
                    
                    pbar.update(1)



                    # --------------------------------------------------
                    # Extract description

                    description = ws.range("Description").value

                    

                    # --------------------------------------------------
                    # Compile definitions

                    for (i, definition) in enumerate(field_definitions):
                        if definition != "Definition":
                            definitions[fields[i]] = definition



                    # --------------------------------------------------
                    # Compile notes

                    for (i, note) in enumerate(field_notes):
                        if note != "Notes":
                            notes[fields[i]] = note



                    # --------------------------------------------------
                    # Compile lists

                    for (i, list_name) in enumerate(compiled_lists_names):
                        
                        if list_name:
                            lists[list_name[:-3]] = ast.literal_eval(compiled_lists[i])



                    # --------------------------------------------------
                    # Extract altered_source_data
                    
                    ds.source_data = ws.tables[ds.table_name].range.options(pd.DataFrame, index=False).value     
                    ds.df_overview = book.sheets("Overview").tables["tbl_DfOverview"].range.options(pd.DataFrame, index=False).value
                    ds.field_overview = book.sheets("Overview").tables["tbl_FieldOverview"].range.options(pd.DataFrame, index=False).value
                    


                    # --------------------------------------------------
                    # Prepare exports

                    # Add export properties
                    bp.description = description
                    bp.definitions = definitions
                    bp.notes = notes
                    bp.lists = lists
        
                    pbar.update(1)
            
            

            # Restore app settings before exiting the context manager
            app.display_alerts = True
            app.screen_updating = True



        # Add an ExportDict object for each dataset into self.export_dicts
        self.export_dicts = [ExportDict(ds) for ds in datasets]

        

        # --------------------------------------------------
        # Print closing output
        
        
        if missing_dfs:
            
            env.warn_print("\nThe following worksheets were not found and are using default metadata:\n")
            for sht in missing_dfs:
                env.warn_print(f"    {sht}")
                        

        duration = time.time() - logger.start
        theme.print(f"\nExport completed after {int(duration)} seconds" + separator)


        
        
# -------------------------------------------------
# CLI


# Construct typer app and layout primary commands
cli = typer.Typer(epilog=help_message, rich_markup_mode="markdown")



@cli.command()
def install():
    """Installs the right-click context menu."""
    app = CLI() 
    app.install()



@cli.command()
def uninstall():
    """Uninstalls the right-click context menu."""
    app = CLI() 
    app.uninstall()



@cli.command()
def version():
    """Checks for an updated version on PyPI."""
    app = CLI() 
    app.version()



@cli.command(name="wb", epilog=help_message)
def cli_wb(data: str = typer.Argument(..., help="Path to a supported data file"),
           file_name: str = typer.Option(None, "--file_name", show_default=False, help="Name of the created workbook. Defaults to the same name as the data file"),
           theme_color: str = typer.Option(None, "--theme_color", help="Hex color used for theme in workbook. Using this setting will change the default.  Defaults to a neutral color"),
           export: bool = typer.Option(False, help="Export from an xleda workbook"),
           large_report: bool = typer.Option(False, "--large_report", help="Only subsample when required to fit within Excel's worksheet limits"),
           overwrite: bool = typer.Option(False, help="Overwrite existing workbook"),
           wb_path: str = typer.Option('', "--wb_path", show_default=False, help="Workbook directory with/without filename"),
           open_wb: bool | None = typer.Option(True, "--open_wb/--no_open_wb", help="Don't automatically open the workbook on finish"),
           no_vba: bool | None = typer.Option(None, "--vba/--no_vba", help="Create a VBA-free xlsx file.  Using this setting will change the default.  Defaults to False"),
           debug: bool = typer.Option(False, "--debug", help="View the workbook while it's being created")):

        """
        Runs export from the created wb object

        Returns
        -------
        list[ExportDict]
            A list of ExportDict objects

        """

        self.theme.print(separator + "FieldAnalysis() has been replaced by wb()")
        self.theme.print("\nUse xleda.wb(export=True) to export from an xleda workbook")

        self._export_analysis()

        return self.export_dicts


