import pandas as pd
from pathlib import Path
import ast
import sys
import time
import typer

from tqdm.auto import tqdm
from matplotlib.figure import Figure
import xlwings as xw

from .utilities import (Environment, Config, Theme, Plotter, DataSet, DataSetParser, ExportDict, 
                        PerformanceLogger, DataError, template_objects)

from .os_interface import install, uninstall, supported_extensions

separator = "\n" + ("-" * 100)






# -------------------------------------------------
# Construct CLI


def construct_help_msg() -> str:
    separator = "-" * 80

    
    message = (
        f"{separator}\n\n"
        r"Use 'xleda wb \<data file path\>' to create a workbook" + "\n\n"
        "### Supported file types:\nCSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML" + "<br><br>\n\n"
        f"### Expected extensions:{str(supported_extensions)[1:-1]}\n\n"
        "For more documentation, visit https://github.com/InfoDesigner/xleda"
        )
    return message


app = typer.Typer(epilog=construct_help_msg(), rich_markup_mode="markdown")

app.command()(install)
app.command()(uninstall)


@app.command(name="wb", epilog=construct_help_msg())
def wb_cli(data: str = typer.Argument(..., help="Path to a supported data file"),
           name: str = typer.Option(None, show_default=False, help="Name of the created workbook. Defaults to the same name as the data file"),
           theme_color: str = typer.Option("#262626", "--theme_color", help="Hex color used for theme in workbook and default plots"),
           export: bool = typer.Option(False, help="Export from an xleda workbook"),
           large_report: bool = typer.Option(False, "--large_report", help="Only subsample when required to fit within Excel's worksheet limits"),
           overwrite: bool = typer.Option(False, help="Overwrite existing workbook"),
           wb_path: Path = typer.Option(Path().cwd(), "--wb_path", show_default=False, help="Workbook directory with/without filename"),
           open_wb: bool = typer.Option(True, "--open_wb", help="Automatically open the workbook on finish"),
           no_vba: bool = typer.Option(False, "--no_vba", help="Create an xlsx file without VBA")):

    """
    Creates an xleda workbook from a supported data file

    """


    return wb(**locals())



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
                 theme_color: str = "#262626",
                 plots: dict[str, Figure] = {},
                 export: bool = False,
                 debug: bool = False,
                 large_report: bool = False,
                 overwrite: bool = False,
                 wb_path: str | Path = Path().cwd(),
                 open_wb: bool = True,
                 no_vba: bool = False,
                 
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
                * Dictonary of pandas dataframes in '{'df1_name': df1, 'df2_name': df2, ...}' format
            * Will create an xleda workbook that is 25,000 rows/50 columns by default.  

        file_name : str
            * Name of the workbook to be created.
            * Defaults to:
                * The name of the file provided for 'wb_path' or 'data'
                * The first key if a dict[str, pd.DataFrame] argument is provided for 'data'
                * 'xleda' if neither option above is valid
                
        theme_color : str, optional
            * A hexidecimal color used for charts/accent color.  
            * Use theme_color='random' for random colors
            * Defaults to "#262626"

        plots : dict[str, Figure], optional
            * Additional plots to be included 
            * Uses "{'plot1_name': Plot1Figure, 'plot2_name': Plot2Figure, ...}" format.  
            * Each entry will get it's own worksheet.  
            * No resizing or syling is done for plots added this way
            * Defaults to None

        large_report : bool, optional
            * Used to override default limits of 25,000 rows/50 columns. 
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
        # Check Environment/Create Dataset


        # Initialize/Check Environment
        self.env = Environment()
        
    
        
        # TODO: Progress bar should start here 
        
        # Add source datasets and plots
        self.datasets: list[DataSet] = self._configure_datasets(data=data,
                                                                input_df=input_df)
        self.plots: dict[str, Figure] = plots



        # ------------------------------------------------------------------------------
        # Initialize other components
            

        # Initialize logger/log input variables
        self.logger = PerformanceLogger(wb=self)

        # Intialize theme
        self.theme = Theme(theme_color,
                           env=self.env)

        # Intialize plotter
        self.plotter = Plotter(theme=self.theme,
                               env=self.env)

        # Intialize Config
        self.cfg = Config(wb=self, **locals())

        # Add initial export dicts
        self.export_dicts: list[ExportDict] = [ExportDict(ds) for ds in self.datasets]
    

        # ------------------------------------------------------------------------------
        # Either create a workbook or export

        if export:
            self._export_analysis()
        else:
            self._assemble_wb()


    def _configure_datasets(self,input_df: pd.DataFrame | None,
                            data: pd.DataFrame | str | Path | dict[str, pd.DataFrame] | None) -> list[DataSet]:
        
        """
        Create a DataSet/adjust properties as needed
        
        """
        
        
        # ----------------------------------------------------------
        # Parse data from input arguments
        
        
        # TODO: Delete me on 8.18
         # Handle neither data argument provided
        if data is None and input_df is None:
            self.env.warn_print("No Data Provided")
            sys.exit()
        
        # Handle the old placeholder argument provided
        elif (input_df is not None and data is None):
            self.env.warn_print("The 'input_df' argument has been changed to 'data'")
            data = input_df
            
        # Handle both the old and new arguments provided
        elif (input_df is not None and data is not None):
            self.env.warn_print("The 'input_df' argument has been changed to 'data'")

        
        
        # ----------------------------------------------------------
        # Create dataset from data source
        
        # If a dataframe is provided, create a dataset from it
        if isinstance(data, pd.DataFrame):
            source_data = {self.file_name: data}
            return DataSetParser(data=source_data,
                                 large_report=self.cfg.large_report).datasets
            
        # If a dictionary of dataframes is provided, create a dataset from it
        elif isinstance(data, dict) and all(isinstance(v, pd.DataFrame) for v in data.values()):
            return DataSetParser(data=data,
                                 large_report=self.cfg.large_report).datasets
        
        # If a path has been provided, parse it for other details and create a dataset from it
        elif isinstance(data, Path) or isinstance(data, str):
        
            # Resolve the path, including any provided user profile components
            resolved_path = Path(data).expanduser().resolve()
        
            # Use the datafile name for name if it hasn't been explicitly provided
            if not self.file_name or self.file_name == 'xleda':
                self.file_name = resolved_path.stem
                
                # Adjust the name if the source file is an xlsm file
                if resolved_path.suffix in ['.xlsm', '.xlsx']:
                    self.file_name += '_xleda'
        
            # Use the datafile directory for wb_path if it hasn't been explicitly provided
            if not self.wb_path:
                self.wb_path = resolved_path.parent
                
            return DataSetParser(data=resolved_path,
                                 large_report=self.cfg.large_report).datasets
        else:
            
            msg = f"Unsupported data argument was provided: {type(data)!r}"
            raise DataError(msg)



    def _assemble_wb(self):
        
        """
        Assembles an xleda workbook

        """
        
        # --------------------------------------------------   
        # Construct/print entry message
        

        entry_message = separator

        entry_message += f"\nPreparing an xleda workbook located at:\n    {self.cfg.path}\n"
        entry_message += (f"\n\nProcess started at {time.strftime('%H:%M:%S')}\n")

        # Print entry message
        self.theme.print(entry_message)
        

        # --------------------------------------------------
        # Create the template and initial progress bar
        
        total_iteratons = 7 + (len(self.datasets[1:])) + (2*len(self.datasets))
        pbar = self.theme.create_progress_bar(desc="Creating Workbook...",
                                              total=total_iteratons)
        self.cfg.create_blank_template(progress_bar=pbar)
        


        # --------------------------------------------------
        # Open Excel using a context manager to while creating the workbook

        with xw.App(visible=self.cfg.debug, add_book=False) as app:

            
            
            # Set vars, open workbook
            book = app.books.open(self.cfg.path, read_only=False)
            
            app.display_alerts = self.cfg.debug
            app.screen_updating = self.cfg.debug
            self.book = book
            pbar.update(2)

            
            

            # Create placeholder worksheets
            self._create_worksheets(progress_bar=pbar)
            self.logger.log(log_type='section',
                            details={'Production Section': 'Creating Workbook',
                                     'Production Time in Seconds': time.time() - self.logger.start})
            
            
            # Configure field analyses worksheets
            self._configure_field_analyses(progress_bar=pbar)
            self.logger.log(log_type='section',
                            details={'Production Section': 'Configure Template',
                                     'Production Time in Seconds': ''})          
            pbar.close()



            # --------------------------------------------------
            # Adding Data

            with self.theme.create_progress_bar(desc="Adding Data...", 
                                                total=8*len(self.datasets)) as pbar:
                
                
                # Add field metadata
                self._add_field_metadata(progress_bar=pbar)
                self.logger.log(log_type='section',
                                details={'Production Section': 'Adding Field Metadata',
                                         'Production Time in Seconds': ''})
                

                # Add overview
                self._add_overview(progress_bar=pbar)
                self.logger.log(log_type='section',
                                details={'Production Section': 'Adding Overview',
                                         'Production Time in Seconds': ''})
                

                # Add source data
                self._add_source_data(progress_bar=pbar)
                self.logger.log(log_type='section',
                                details={'Production Section': 'Adding Source Data',
                                         'Production Time in Seconds': ''})



            # --------------------------------------------------
            # Adding Plots
            
            total_iteratons = sum([ds.columns for ds in self.datasets]) + len(self.cfg.plots or {})
            with self.theme.create_progress_bar(desc="Adding Plots...",
                                                total=total_iteratons) as pbar:
            
                # Adds additional plots
                if self.cfg.plots:
                    self._add_additional_plots(progress_bar=pbar)


                self._add_plots(progress_bar=pbar)
                
                
                self.logger.log(log_type='section',
                                details={'Production Section': 'Adding Plots',
                                         'Production Time in Seconds': ''})


            # --------------------------------------------------
            # Configuring Workbook

            total_iteratons = 8 + len(self.datasets)

            with self.theme.create_progress_bar(desc="Configuring Workbook...", 
                                                total=total_iteratons) as pbar:

                self._configure_pivot(progress_bar=pbar)
                self.logger.log(log_type='section',
                                details={'Production Section': 'Configure Pivots',
                                         'Production Time in Seconds': ''})

                self._initialize_ui(progress_bar=pbar)
                self.logger.log(log_type='section',
                                details={'Production Section': 'Configuring Workbook',
                                         'Production Time in Seconds': ''})
                

                # Add production time to log
                self.logger.total_production_time = time.time() - self.logger.start
                
                # Add performance_metadata log
                self.logger.log(log_type='performance_metadata',
                                details={'Dataframes Included': len(self.datasets),
                                         'Plots Included': len(self.cfg.plots or {}) + len(self.logger.performance_logs['plots']),
                                         'Columns Included': sum([bp.columns for bp in self.datasets]),
                                         'Rows Included': sum([bp.rows for bp in self.datasets]),
                                         'Production Time': self.logger.total_production_time})
                
                pbar.update(1)
                
                
                # Close logs and write to workbook
                self.logger.close()
                
                self._add_debug()

                pbar.update(1)


                # --------------------------------------------------
                # Restore app configuration, save workbook, close Excel

                        
                # Set focus to primary field analysis worksheet
                book.sheets(self.datasets[0].name).activate()

                # Restore app settings
                app.display_alerts = True
                app.screen_updating = True
                
                
                # Save/exit context manager
                book.save(self.cfg.path)
               
                pbar.update(1)



        # --------------------------------------------------
        # Compile closing messsage

        exit_msg = self.cfg.exit_msg
        
        exit_msg = f"\n\nxleda workbook created in {int(self.logger.total_production_time)} seconds.\n"

        
        # Create a new Excel instance for usage
        if self.cfg.open_wb:


            app = xw.App(visible=True, add_book=False) 
            book = app.books.open(self.cfg.path)




        # ----------------------------------------------------------------------------------
        # Construct output messaging


        # Note additional plots
        if self.cfg.plots:
            
            plots = (",").join(self.cfg.plots.keys())
            exit_msg += f"Additional plots included:    {plots}\n"

   
        # Note additional DFs
        if len(self.datasets) > 1:
            
            dataframes = ", ".join([ds.name for ds in self.datasets])
            exit_msg += f"\nDataframes included:\n    {dataframes}\n"

        # Print closing message
        self.theme.print(exit_msg + '\n' + separator)
    
        
    
    def _validate_template(self):
        
        """
        Validates that expected objects are present in the workbook
        
        """

        valid = True
        
        try:
            
            # ---------------------------------------------
            # Validate worksheets and tables exist
            
            actual_objects = {}

            for sheet in self.book.sheets:

                # Omit Pivot which will be handled separately
                if sheet.name != "Pivot":
                    actual_objects[sheet.name] = [tbl.name for tbl in sheet.tables]
                
                else:

                    # ---------------------------------------------
                    # Validate pivot table exists
                    
                    sheet_api = self.book.sheets["Pivot"].api
                    
                    if self.env.win:
                        
                        # Windows uses name string
                        sheet_api.PivotTables("pvt_Pivot")

                    elif self.env.mac:
                        
                        # Appscript uses lowercase dictionary-style bracket
                        sheet_api.pivot_tables["pvt_Pivot"]()
                        
                    actual_objects['Pivot'] = "pvt_Pivot"
            
        except Exception:
            
            valid = False
        
        if not (valid and actual_objects != template_objects):
            
            # Collate missing worksheets
            missing_sheets = [sheet for sheet in template_objects.keys() if sheet not in actual_objects.keys()]
            
            # Collate missing tables
            expected_tables = [table for table_list in template_objects.values() for table in table_list]
            actual_tables = [table for table_list in actual_objects.values() for table in table_list]
            missing_tables = [table for table in expected_tables if table not in actual_tables]
            

            self.env.warn_print(f"Template has been modifed:\n\nThe following worksheets are missing\n    {missing_sheets}\n\nThe following tables are missing\n    {missing_tables}")

            sys.exit()

        
    
    def _create_worksheets(self, progress_bar: tqdm):

        """
        Creates worksheets for any additional dataframes

        """

        book = self.book
        
        # -----------------------------------------------------------------
        # Validates that the expected template objects are present
        
        self._validate_template()


        # Add field analysis worksheets for each dataframe

        for i, dataset in enumerate(self.datasets):
                    
            # Copy the sheet, rename the table
            ws = book.sheets('Field Analysis').copy(name=dataset.name)
            ws.tables[0].name = dataset.table_name

            # Add a color gradient to the worksheet tab to distinguish among them
            self.theme.greyscale_tab(ws=ws,
                                     iteration=i)
    
            progress_bar.update(1)

        
        # Delete the default worksheet now that the rest have bene cared for
        book.sheets("Field Analysis").delete

        progress_bar.update(1)



    def _configure_field_analyses(self, progress_bar: tqdm):

        """
        Configures all Field Analysis worksheets

        """

        
        for ds in self.datasets:

            # --------------------------------------------------
            # Set variables

            book = self.book
            ws = book.sheets(ds.name)
            ws.activate()


            # --------------------------------------------------
            # Format metadata placeholders

            # Set worksheet theme/name
            self.cfg.expand_range(name="Theme", ws=ws, columns=ds.columns + 2)
            self.theme.set_theme(ws.range("Theme"))
            ws.range("Name").value = ds.name



            columns_to_format = ds.columns -3
            if columns_to_format > 0:
                format_from = ws.range("FormatRange")
                format_to = (ws.range("FormatRange").offset(0, 1).resize(None, columns_to_format))
                format_from.copy()
                format_to.paste()

            # Clear clipboard and move selection back to upper left
            book.app.cut_copy_mode = False
            
            # Add all header values except Record List
            headers = ds.source_data.columns.to_list()[1:]
            ws.range("Headers_Start").value = headers
            
            
            
            progress_bar.update(1)



            # --------------------------------------------------
            # Adjust Named Ranges to fit dataframe size
            

            # Expand named ranges to fit number of columns and set input prompt for notes/definitions
            expand_ranges = (["FieldList" + str(i) for i in range(1, 9)] + ["FieldRange", "Notes", "Definitions", "Headers"])

            for name_range in expand_ranges:
                
                self.cfg.expand_range(name=name_range, 
                                      ws=ws, 
                                      columns=ds.columns)

                # Add placeholder values if necessary.
                if name_range in ["Notes", "Definitions"]:    
                    ws.range(name_range).value = name_range[:-1]

                elif name_range.startswith("FieldList"):
                    ws.range(name_range).value = "FALSE"
                    
            
            # Add worksheet level named ranges for each field
            for cell in ws.range("Headers"):
                ws.names.add(name=cell.value, refers_to=cell.address)


            # Set FieldLists range so that it also includes the names on the left
            # Used to recreate completed examples
            self.cfg.expand_range(name="FieldLists", 
                                  ws=ws, 
                                  columns=ds.columns + 1)


            # Show/Hide Data Size Warning
            if ds.warning:
                
                ws.range("Warning").value = ds.warning_msg
                self.cfg.hide_rows(ws.range("Warning"), hide=not ds.warning)

            progress_bar.update(1)



    def _add_field_metadata(self, progress_bar: tqdm):

        """
        Adds field metadata to all Field Analysis worksheets

        Parameters
        ----------

        progress_bar
            A tqdm progress bar object
        
        """



        for ds in self.datasets:

            # --------------------------------------------------
            # Set variables

            book = self.book
            ws = book.sheets(ds.name)
            ws.activate()

            # --------------------------------------------------
            # Add Field Analysis sections to workbook

            ws.range("Dimensions").options(transpose=True).value = list(ds.df_metadata.values())
            ws.range("Composition")[0, 0].offset(0,1).value = ds.composition_df.values
            ws.range("Summary_Stats")[0, 0].offset(0,1).value = ds.summary_stats_df.values
            ws.range("Percentiles")[0, 0].offset(0,1).value = ds.percentiles_df.values

            progress_bar.update(1)



    def _add_overview(self, progress_bar: tqdm):

        """
        Adds overview_metadata to all Overview worksheets
        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object
        
        """

        
        # --------------------------------------------------
        # Set variables
        
        book = self.book
        ws = book.sheets("|")
        df_overview_table = ws.tables["tbl_DfOverview"]
        field_overview_table = ws.tables["tbl_FieldOverview"]
        
        # Set aside placeholder variables for all metadata
        df_overview_df = pd.concat([ds.df_overview for ds in self.datasets], ignore_index=True)
        field_overview_df = pd.concat([ds.field_overview for ds in self.datasets], ignore_index=True)
        
        progress_bar.update(1)
        
        
        # --------------------------------------------------
        # Add rows to make room for df_overview table
        
        start_row = df_overview_table.range.last_cell.row + 1
        end_row = start_row + len(self.datasets) - 1
        row_range_string = f"{start_row}:{end_row}"
        ws.range(row_range_string).insert(shift="down")
        
        # Adjust named ranges to fit new data
        ws.range("Dataframes").resize(row_size=len(df_overview_df) +3, column_size=None).name = "Dataframes"
        ws.range("Fields").resize(row_size=len(field_overview_df) +3, column_size=None).name = "Fields"
                        
        progress_bar.update(1)
        
        
        
        # --------------------------------------------------
        # Set theme, and write data to tables
                    
        # Set theme on target tables
        self.theme.set_theme(df_overview_table.range[0,:])
        self.theme.set_theme(field_overview_table.range[0,:])
        
        # Update the primary tables
        df_overview_table.update(df_overview_df, index=False)
        field_overview_table.update(field_overview_df, index=False)
        
        progress_bar.update(1)
        
        
        # --------------------------------------------------
        # Add formulas to tables 
        
        # Set formulas
        df_description_formula = '''=INDIRECT("'"&[@Dataframe]&"'!Description")'''
        df_links_formula = '''=HYPERLINK("#'"&[@Dataframe]&"'!Headers_Start", "Link")'''
        df_fields_defined_pct_formula = '''=IFERROR(SUMPRODUCT((tbl_FieldOverview[Dataframe]=[@Dataframe]) * (tbl_FieldOverview[Definition]<>"Definition"))/[@Columns],"")'''
        field_links_formula = '''=HYPERLINK("#'"&[@Dataframe]&"'!"&SUBSTITUTE([@Field]," ","_"), "Link")'''
        field_definitions_formula = '''=XLOOKUP([@Field],INDIRECT("'"&[@Dataframe]&"'!Headers"),INDIRECT("'"&[@Dataframe]&"'!Definitions"),"")'''
        field_notes_formula = '''=XLOOKUP([@Field],INDIRECT("'"&[@Dataframe]&"'!Headers"),INDIRECT("'"&[@Dataframe]&"'!Notes"),"")'''
        
        
        # Add df_overview formulas
        ws.range("tbl_DfOverview[_]").formula = df_links_formula
        ws.range("tbl_DfOverview[Dataframe description").formula = df_description_formula
        ws.range("tbl_DfOverview[Fields Defined %").formula = df_fields_defined_pct_formula
        
        
        # Add field_overview formulas
        ws.range("tbl_FieldOverview[_]").formula = field_links_formula
        ws.range("tbl_FieldOverview[Definition").formula = field_definitions_formula
        ws.range("tbl_FieldOverview[Field Notes").formula = field_notes_formula
                    
        progress_bar.update(1)



    def _add_source_data(self, progress_bar: tqdm):

        """
        Adds source_data to all Field Analysis worksheets

        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object

        """


        for ds in self.datasets:

            # --------------------------------------------------
            # Set variables, dtype conversion if necessary

            book = self.book
            ws = book.sheets(ds.name)
            df = ds.source_data


            # Convert fields with datatypes that Excel doesn't support to string
            supported_dtype_columns = df.select_dtypes(include=['number', 'bool', 'datetime64', 'str'], exclude='timedelta').columns
            unsupported_dtype_columns = [col for col in df.columns if col not in supported_dtype_columns]
            df[unsupported_dtype_columns] = df[unsupported_dtype_columns].astype(str)
            
            progress_bar.update(1)
            


            # --------------------------------------------------
            # Add source data, format cells in Record List/Record Hash/HasBlank columns

            
            source_table = ws.tables[0]


            # If any unsupported columns prevent writing to Excel, convert to string before writing
            try:
                source_table.update(df, index=False)
            except Exception:
                self.logger.log(log_type="error", 
                                details={'Error Type': 'Error writing source data that required string conversion',
                                         'Detail': ds.name})
                source_table.update(df.astype(str), index=False)

            progress_bar.update(1)



            
            # --------------------------------------------------
            # Set formatting for tbl_SourceData and added columns

            
            self.cfg.set_cell_alignment(input_range=source_table.range,
                                        horizontal='center')
            record_list = source_table.range[:, :1 ]
            other_added_columns = source_table.range[:, -3: ]

            # Reduce contrast to subdue added fields
            for dimmed_range in [record_list, other_added_columns]:
                self.theme.greyscale_range(dimmed_range)


            progress_bar.update(1)

    

    def _add_plots(self, progress_bar: tqdm):
        
        """
        Adds plots to all Field Analysis worksheets
        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object

        """

        

      
        for ds in self.datasets:
        
            # --------------------------------------------------
            # Set vars, start performance logging

            book = self.book
            ws = book.sheets(ds.name)
            df = ds.source_data.iloc[:, 1:-3]
                

            # Set initial ranges for added plots and ensure they aren't hidden
            histogram_range = ws.range("Histogram")
            composition_range = ws.range("CompositionTable")
            

            self.cfg.hide_rows(histogram_range, hide=False)
            self.cfg.hide_rows(composition_range, hide=False)


            # --------------------------------------------------
            # Add plots for all except added columns

            for col in df.columns:


                # --------------------------------------------------
                # Add Composition Table Plot

                composition_table = self.plotter.create_composition_plot(df[col])

                self.plotter.add_small_plot(target_range=composition_range,
                                            fig=composition_table,
                                            name=f'composition_{col}')

                if pd.api.types.is_numeric_dtype(df[col]):



                    # --------------------------------------------------
                    # Add Histogram Multiple

                    histogram = self.plotter.create_histogram_plot(df[col])

                    self.plotter.add_small_plot(target_range=histogram_range,
                                                fig=histogram,
                                                name=f'histogram_{col}')


                # --------------------------------------------------
                # Increment Target Ranges/progress bar

                histogram_range = histogram_range.offset(0, 1)
                composition_range = composition_range.offset(0, 1)
                progress_bar.update(1)



    def _add_additional_plots(self, progress_bar: tqdm):
            
        """
        Adds additional plot worksheets

        """

        
        # Set vars
        book = self.book


        # --------------------------------------------------
        # Add additional plots



        for plot_name, figure in self.plots:
           
            
            # Plots will be added before all other sheets
            anchor_sheet = self.book.sheets[0]
            
            # Create a copy of the SinglePlot template, make it visible, set theme
            ws = book.sheets("SinglePlot").copy(before=anchor_sheet, name=plot_name)
            ws.visible = True
            self.theme.set_theme(ws.range("Theme"))

            # Set target range, title, and autofit title range
            plot_range = ws.range("SinglePlot")
            ws.range("Name").value = plot_name

            
            ws.pictures.add(figure, 
                            name=plot_name, 
                            update=True,
                            left=plot_range.left, 
                            top=plot_range.top)
            
            progress_bar.update(1)



    def _configure_pivot(self, progress_bar: tqdm):
            
        """
        Configures the pivot table
        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object

        """


        # --------------------------------------------------
        # Set vars/intialize logging

        ds = self.datasets[0]
        df = ds.original_df.copy()
        book = self.book

        # Fields to be added to the pivot tables
        pivot_fields = df.columns.to_list()[:min(10, len(df.columns))]
        
        ws = book.sheets("Pivot")
        ws.activate()
        
        progress_bar.update(1)



        # --------------------------------------------------
        # Set theme/show warning if necessary


        self.theme.set_theme(ws.range("Theme"))

        if ds.warning:
            ws.range("Warning").value = ds.warning_msg
            self.cfg.hide_rows(ws.range("Warning"), hide=not ds.warning)

        
        # Show field limit warning for datasets with >10 columns
        if len(df.columns) <= 10:
            ws.range("DefaultWarn").value = ''
        else:
            ws.range("DefaultWarn").value = 'Pivots only show first 10 fields by default'

        progress_bar.update(1)



        # --------------------------------------------------
        # configure pivot tables and get ranges together
        
        # Range of pivot table
        pivot_range = self.cfg.get_configured_pivot_range(ws=ws,
                                                          pivot_fields=pivot_fields)
        # Top row/first column
        pivot_headers = pivot_range[0,:]            
        header_column = pivot_range[: , 0]
        
        progress_bar.update(1)
        
        
        
        
        # --------------------------------------------------
        # Update the metadata table and refresh it's slicers and pivot table
        
        tables_table = book.sheets("meta").tables("tbl_Tables")
        
        tables_df = pd.DataFrame({"TableNames": [ds.table_name for ds in self.datasets],
                                  "WorksheetNames": [ds.name for ds in self.datasets]})
        
        tables_table.update(tables_df, index=False)
        
        
        self.cfg.get_updated_pivot(ws=book.sheets("meta"), 
                                   pt_name = 'pvt_TableSelector')
                  
            
        # --------------------------------------------------
        # Format headers, first column, and column width

        pivot_headers.wrap_text = True

        self.cfg.set_cell_alignment(input_range=pivot_headers,
                                    horizontal='center',
                                    vertical='center')
        
        self.cfg.set_cell_alignment(input_range=header_column,
                                    horizontal='left')
        
        pivot_range.columns.autofit()

        progress_bar.update(1)

            

        # --------------------------------------------------
        # Set last three pivot columns width to 13 and center


        pivot_range[:, -3:].column_width = 13
        self.theme.greyscale_range(pivot_range[:, -3:])

        self.cfg.set_cell_alignment(input_range=ws.range(pivot_range[:, -2:].address),
                                    horizontal='center')
        
        progress_bar.update(1)
            


    def _initialize_ui(self, progress_bar: tqdm):
                
        """
        Prepares all field analyis worksheet UIs

        """

        book = self.book


        # Loop through all datasets and configure Field Analysis worksheets

        for ds in self.datasets:
            ws = book.sheets(ds.name)
            ws.activate()
            ws.range('Headers_Start').select()

            # Orient toggles, and collapse subsections
            self.cfg.set_text_orientation(input_range=ws.range("Toggles"))
            self.cfg.set_text_orientation(input_range=ws.range("TopToggle"), degrees=-90)

            for excel_range in ["Data_Description", "Composition", "Summary_Stats", 
                                "Percentiles", "Field_Lists", "Compiled_Lists"]:
                            
                self.cfg.hide_rows(ws.range(excel_range), hide=True)

            progress_bar.update(1)

            ws.range("Headers_Start").select()



    def _add_debug(self):

        """
        Configures the debug worksheet
        
        """
        
        book = self.book
        ws = book.sheets('Overview')
        ws.activate()
       
        
        # Write debug range
        ws.range("debug_metadata").options(transpose=True).value = self.logger.performance_metadata.values


        # Write debug tables
        ws.tables("tbl_debug_environment").update(self.logger.env, index=False)
        ws.tables("tbl_debug_config").update(self.logger.config, index=False)
        ws.tables("tbl_debug_errors").update(self.logger.errors, index=False)
        ws.tables("tbl_debug_section").update(self.logger.section_performance, index=False)


        # Hide config section and set toggle orientation
        self.cfg.set_text_orientation(input_range=ws.range("Toggles"))
        self.cfg.hide_rows(ws.range("Debug"), hide=True)
        


    def _export_analysis(self):

        """
        Exports data from an xleda workbook into self.export_dicts

        """
       

        missing_dfs = []

        # --------------------------------------------------
        # If the file is not found return messaging

        if not self.cfg.path.is_file():

            self.theme.print(f"File not found at {self.cfg.path}\n")
            self.theme.print("wb().export_dicts will be limited" + separator)

            sys.exit()



        # --------------------------------------------------
        # Print starting output

        self.theme.print(separator + f"\nExport started at {time.strftime('%H:%M:%S')}\n")
        


        # --------------------------------------------------
        # Open Excel/export data using a context manager


        with xw.App(visible=self.cfg.debug, add_book=False) as app:
            
            app.display_alerts = self.cfg.debug
            app.screen_updating = self.cfg.debug
           

            with self.theme.create_progress_bar(desc="Reading workbook...",
                                                total=4 + (2 * len(self.datasets))) as pbar:
                              
                pbar.update(2)

                try:
                    wb = app.books.open(self.cfg.path)
                    pbar.update(2)

                except FileNotFoundError:
                    self.theme.print(f"File not found at {self.cfg.path}")
                    sys.exit()

                
                for ds in self.datasets:

                    definitions = {}
                    notes = {}
                    lists = {}

                    if ds.name in wb.sheet_names:
                        ws = wb.sheets(ds.name)
                    else:

                        missing_dfs.append(ds.name)
                        
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
                    # Extract altered_source_data, field_overview, and df_overview
                    
                    ds.source_data = ws.tables[ds.table_name].range.options(pd.DataFrame, index=False).value     
                                   
                    ds.df_overview = self.book.sheets("|").tables["tbl_DfOverview"].range.options(pd.DataFrame, index=False).value
                    ds.field_overview = self.book.sheets("|").tables["tbl_FieldOverview"].range.options(pd.DataFrame, index=False).value
                    

                    # --------------------------------------------------
                    # Prepare exports

                    # Add export properties
                    ds.description = description
                    ds.definitions = definitions
                    ds.notes = notes
                    ds.lists = lists
        
                    pbar.update(1)
            
            app.display_alerts = True
            app.screen_updating = True


        # Add an ExportDict object for each dataframe into self.export_dicts
        self.export_dicts = [ExportDict(ds) for ds in self.datasets]


        # --------------------------------------------------
        # Print closing output
        
        
        if missing_dfs:
            
            self.env.warn_print("\nExports for the following dataframes were not found:\n")
            for sht in missing_dfs:
                self.env.warn_print(f"    {sht}")
                        

        duration = time.time() - self.logger.start
        self.theme.print(f"\nExport completed after {int(duration)} seconds" + separator)





class FieldAnalysis(wb):

    """
    Class representing a FieldAnalysis object
    
    """
    def __init__(self, input_df: pd.DataFrame, export: bool = False, **kwargs) -> None:

        """ 
        A placeholder for legacy FieldAnalysis support.

        Creating a separate Field Analysis configuration is no longer required.
        
        Use xleda.wb to both configure and create a workbook
             
        """
        self.env.warn_print(separator + 
              "\nCreating a separate Field Analysis configuration is no longer required.\n"
              "Use xleda.wb to both configure and create a workbook\n" + separator)

        super().__init__(data=input_df,
                         **kwargs)


    def create_workbook(self):

        """
        FieldAnalysis() has been replaced by wb()
        
        Use xleda.wb() to both configure and create a workbook

        """ 

        self.theme.print(separator + "FieldAnalysis() has been replaced by wb()")
        self.theme.print("\nUse xleda.wb() to both configure and create a workbook")


    
    def export_analysis(self) -> list[ExportDict]:

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
