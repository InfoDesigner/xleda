from __future__ import annotations

# Base imports
import pandas as pd
from pathlib import Path
import ast
import time
from tqdm.auto import tqdm
from matplotlib.figure import Figure
import xlwings as xw


from .utilities import (Template, DataSet, DataSetParser, ExportDict, 
                        Logger, Settings, separator, TemplateError)





# -------------------------------------------------------
# Primary class


class wb():

    """
    A class that represents an xleda workbook.

    """

    def __init__(self,
                 data: pd.DataFrame | str | Path | dict[str, pd.DataFrame] | None = None,       
                 # Every other argument is keyword only
                 *,
                 file_name: str = 'xleda',
                 theme: str = "",
                 plots: dict[str, Figure] = {},
                 
                 export: bool = False,
                 debug: bool = False,
                 large_report: bool = False,
                 overwrite: bool = False,
                 wb_path: str | Path = '',
                 open_wb: bool = True,
                 no_vba: bool | None = None,
                 
                 # TODO: Remove this on 8.18
                 input_df: pd.DataFrame | None = None) -> None:
        """
        Creates an xleda workbook

        Parameters
        ----------

        data : pd.DataFrame | str | Path | dict[str, pd.DataFrame]
            * Options:
                * A pandas dataframe
                * A str or Path to a supported data file
                * Dictonary of pandas dataframes in '{'df1_name': df1, ...}' format
                * Defaults to limit each provided dataframe to a sample of 25,000 rows/first 50 columns

        file_name : str
            * File name of the workbook to be created.
            * Defaults to:
                * The name of the file provided for 'wb_path' or 'data'
                * The first key if a dict[str, pd.DataFrame] argument is provided for 'data'
                * 'xleda' if neither option above is valid
                
        theme : str, optional
            * A hexidecimal color used for charts/accent color.  
            * Use theme='random' for random colors.
            * Changing this setting changes the default color.
            * Defaults to "#262626"

        plots : dict[str, Figure], optional
            * Additional plots to be included in "{'plot1_name': Plot1Figure, ...}" format.  
            * Each entry will get it's own worksheet.  
            * No resizing or syling is done for plots added this way
            * Defaults to None

        large_report : bool, optional
            * This flag overrides the default limits of 25,000 rows/50 columns
            * Sets limits to 1,000,000 rows/16,000 columns
            * Defaults to False

        overwrite : bool, optional
            * Whether to overwrite existing reports with the same name
            * Defaults to False

        wb_path : Path | string, optional
            * String or Pathlib path
            * If a directory is provided, the workbook will be created in that directory.
            * If a file name ending in ".xlsm" or ".xlsx" is provided, 
                it will either create that file or export from that file, 
                depending on whether export=True is also selected.
            * Defaults to current working directory

        no_vba : bool, optional
            * Will create the workbook as a '.xlsx' file that has no VBA.
            * Providing a 'wb_path' argument with a '.xlsx' extension has the same effect.
            * Changing this setting changes the default.
            * Defaults to False
            
        open_wb: bool, optional
            * Whether to open the workbook after creating.  
            * Useful if creating multiple workbooks.
            * Defaults to True

        export: bool, optional
            * Exports data from an xleda workbook instead of creating one.
            * Exported data is available through wb().export_dicts
            * Defaults to False
            * Export includes the follwing fields for each provided dataframe
                * `description`: Dataframe description if you've added one
                * `definitions`: Any field definitions you've added.
                * `notes`: Any field notes you've added
                * `lists`: Any lists showing in the compiled lists section
                * 'field_metadata': A basic metadata dataframe, combining information from 
                                    pandas info/describe/quantile.
                * 'df_overview': df metadata exported from the Overview worksheet.
                * 'field_overview': Field overview metadata exported from the Overview worksheet.
                * `source_data`: Source data exported from the workbook that includes any manual edits you've 
                                 made such as removing records, renaming fields, etc. 
                ** Note that data types will likely change in the round-trip translation. **

        debug: bool, optional
            * Shows the workbook being created
            
            
        input_df: pd.DataFrame, optional
            * A placeholder to retain backwards compatibility for the new 'data' argument

        """

        # ------------------------------------------------------------------------------
        # Create Base Components


        # Initialize logger
        self.logger = Logger()
        
    
        # Create/update settings
        self.settings = Settings(locals=locals(), 
                                 logger=self.logger)
        
        

        # ------------------------------------------------------------------------------
        # Initial Output, Prepare Datasets/Template
        
        
        # Print initial text/progress bar
        pbar = self.logger.start_output(wb=self)
        
        
        # Prepare datasets
        self.datasets: list[DataSet] = DataSetParser(settings=self.settings, 
                                                     logger=self.logger).datasets
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
            logger.log(section='Initializing wb Components')
            

            app.display_alerts = debug
            app.screen_updating = debug
            
            # --------------------------------------------------
            # Adding Data
            
            with logger.create_progress_bar(desc="Adding Data...",
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
            with logger.create_progress_bar(desc="Adding Plots...",
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
            

            

        # --------------------------------------------------
        # Reopen the workbook in a new context manager if needed
        
        
        if self.settings.open_wb:
            
            app = xw.App(visible=True, add_book=False) 
            book = app.books.open(template.path)




    def _export_analysis(self, progress_bar: tqdm):

        """
        Exports data from an xleda workbook into self.export_dicts

        """
        
        # -------------------------------------------------
        # Set vars
        
        template = self.template
        settings = self.settings
        datasets = self.datasets
        logger = self.logger

        missing_dfs = []



        # --------------------------------------------------
        # If the file is not found return messaging
        
        path = template.path

        if not path.is_file():

            msg = f"File not found at {template.path}\n"
            msg +=  "wb().export_dicts will be limited" + separator
            
            raise TemplateError(msg)
        
        progress_bar.update(2) # 8


        # --------------------------------------------------
        # Open Excel/export data using a context manager


        with xw.App(visible=settings.debug, add_book=False) as app:
            
            app.display_alerts = False
            app.screen_updating = settings.debug
                        
            
            # Close the initial progress bar
            progress_bar.update(2) # 10
            progress_bar.close()
            


            with logger.create_progress_bar(desc="Reading workbook...",
                                           total=4 + (2 * len(datasets))) as pbar:
                              
                pbar.update(2)
                book = app.books.open(path, read_only=True)
                pbar.update(2)

                
                for ds in datasets:

                    definitions = {}
                    notes = {}
                    lists = {}

                    if ds.name in book.sheet_names:
                        ws = book.sheets(ds.name)
                    else:

                        missing_dfs.append(ds.name)
                        pbar.update(2)
                        
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
                    # Extract source_data, df_overview, field_overview
                    
                    ds.source_data = ws.tables[ds.table_name].range.options(pd.DataFrame, index=False).value     
                    ds.df_overview = book.sheets("Overview").tables["tbl_DfOverview"].range.options(pd.DataFrame, index=False).value
                    ds.field_overview = book.sheets("Overview").tables["tbl_FieldOverview"].range.options(pd.DataFrame, index=False).value
                    


                    # --------------------------------------------------
                    # Prepare exports

                    # Add export properties
                    ds.description = description
                    ds.definitions = definitions
                    ds.notes = notes
                    ds.lists = lists
        
                    pbar.update(1)
            
            

            # Restore app settings before exiting the context manager
            app.display_alerts = True
            app.screen_updating = True



        # Add an ExportDict object for each dataset into self.export_dicts
        self.export_dicts = [ExportDict(ds) for ds in datasets]



        # --------------------------------------------------
        # Print closing output
        
        
        # Note any datframes that weren't exported
        if missing_dfs:
            
            logger.warn_print("\nThe following worksheets were not found and are using default metadata:\n")
            for sht in missing_dfs:
                logger.warn_print(f"    {sht}")


        duration = time.time() - logger.start
        logger.print(f"\nExport completed after {int(duration)} seconds" + separator)


        
