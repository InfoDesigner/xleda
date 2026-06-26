from __future__ import annotations

# Base imports
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
                        PerformanceLogger, DataError, CLI, template_objects, supported_extensions, help_message)


separator = "\n" + ("-" * 100)






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
                 theme_color: str = "",
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
                
        theme_color : str, optional
            * A hexidecimal color used for charts/accent color.  
            * Use theme_color='random' for random colors.
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
        # Check Environment/Create wb Components


        # Add base properties from input
        self.plots: dict[str, Figure] = plots
        
        
        # Create an input variable dict without self
        input_vars = locals()
        input_vars.pop("self", None)

        # Initialize/Check Environment
        self.env = Environment(input_vars=input_vars)

        
        # Intialize theme
        self.theme = Theme(env=self.env)
        
        # Intialize plotter
        self.plotter = Plotter(theme=self.theme,
                               env=self.env)

        # Initialize logger
        self.logger = PerformanceLogger(wb=self)
        
        # Ensure that a valid data argument is provided and format it
        self._validate_data(input_vars=input_vars)
        data = input_vars.get('data')
        self.env.data_argument = str(input_vars.get('data_argument'))
        
        
        assert isinstance(data, (dict, Path))
        
        
        # --------------------------------------------------   
        # Create DataSets, print initial outputs
        
        
        if not export:
            self.theme.print(separator + f"\nStarted preparing an xleda workbook at {time.strftime('%H:%M:%S')}\n\n")
        else:
            self.theme.print(separator + f"\nExport started at {time.strftime('%H:%M:%S')}\n\n")


        pbar = self.theme.create_progress_bar(desc="Preparing Data...",
                                              total=10)
       
        # Prepare datasets
        self.datasets: list[DataSet] = DataSetParser(data=data, 
                                                     large_report=large_report).datasets
        pbar.update(6) # 6


        # ------------------------------------------------------------------------------
        # Initialize Config, prepare initial logs
            

        # Intialize Config
        self.cfg = Config(wb=self, input_vars=input_vars)
        pbar.update(2) # 8

        # Add initial export dicts
        self.export_dicts: list[ExportDict] = [ExportDict(ds) for ds in self.datasets]
        
        # Close the "Preparing Data" progress bar
        pbar.update(2) # 10
        pbar.close()
        
        # Add config/section log
        self.logger.add_config_log(wb=self)
        self.logger.log(log_type='section',
                        details={'Production Section': 'Preparing Data',
                                 'Production Time in Seconds': ''})  
    

        # ------------------------------------------------------------------------------
        # Either create a workbook or export

        if export:
            self._export_analysis()
        else:
            self._assemble_wb()


    def _validate_data(self, input_vars: dict) -> None:
        
        """
        Validates that a supported data source has been provided

        """
           

        # Set vars
        data = input_vars.get('data', None)
        input_df = input_vars.get('input_df', None)
        supported = ", ".join(sorted(supported_extensions))

        
        
        # TODO: Remove placeholder API on 8.18
        
        
        # Handle neither data argument provided
        if data is None and input_df is None:
            
            raise DataError("No Data Provided")
        
        # Handle 'input_df' argument provided without 'data'
        elif (isinstance(input_df, pd.DataFrame) and data is None):
            
            self.env.warn_print("The 'input_df' argument has been replaced by 'data'")
            input_vars['data'] = {input_vars['name']: input_df}
            
        # Handle both 'input_df' and  'data' arguments provided
        elif (input_df is not None and data is not None):
            
            self.env.warn_print("The 'input_df' argument has been replaced by 'data', ignoring 'input_df'.")
            
        # Only the data argument has been provided, validate that it is supported
        else:
            
            # if data is a dictionary of dataframes, use it
            if isinstance(data, dict) and all(
                isinstance(k, str) and isinstance(v, pd.DataFrame) for k, v in data.items()):
                
                input_vars['data_argument'] = 'Dataframe dictionary'
                input_vars['data'] = data
                
                # if no file_name has been provided, use the first key as the file name
                if input_vars.get('file_name', 'xleda') == 'xleda':
                    input_vars['file_name'] = list(data.keys())[0]
                

            
            # if data is a dataframe, convert it to a dataframe dictionary
            elif isinstance(data, pd.DataFrame):
                
                input_vars['data_argument'] = 'Dataframe'
                input_vars['data'] = {input_vars['name']: data}

            
            
            # if data is neither a dataframe or a dataframe dict, ensure it's a supported file
            elif isinstance(data, (str, Path)):
                
                # Get the path
                path = Path(data).expanduser().resolve()
                input_vars['data_argument'] = str(path)
                
                # Ensure it's a file
                if not path.is_file():
                    raise DataError("Data file not found")

                # Since it is a file, ensure that it has supported extension or is an .rdata file
                if not (path.stem.lower() == '.rdata') and (path.suffix.lower() not in supported_extensions):
                    raise DataError(f"Unsupported file type '{path.suffix}'. Supported files: {supported}")
                
                # Since it's a file with a supported extension,
                else:
                    
                    # Use it for data
                    input_vars['data'] = path
                    
                    # If wb_path hasn't been provided, use the data file directory
                    if not input_vars.get('wb_path', None):
                        input_vars['wb_path'] = path.parent
                        
                        
                    # if file_name hasn't been provided, use the data file name
                    if input_vars.get('file_name', 'xleda') == 'xleda':
                        input_vars['file_name'] = path.stem
                
                        # Adjust the name if the source file is an excel file to prevent collissions
                        if path.suffix in ['.xlsm', '.xlsx']:
                            input_vars['file_name'] += '_xleda'
        
                    
            
            else:
                raise DataError("Data file not found") 


    def _assemble_wb(self):
        
        """
        Assembles an xleda workbook

        """
    
        

        # --------------------------------------------------
        # Create the template and initial progress bar
        
        total_iterations = 7 + (3 * len(self.datasets))
        
        pbar = self.theme.create_progress_bar(desc="Preparing Template...",
                                              total=total_iterations)
        
        self.cfg.create_blank_template(progress_bar=pbar) # 4
        


        # --------------------------------------------------
        # Open Excel using a context manager to while creating the workbook

        with xw.App(visible=self.env.debug, add_book=False) as app:

            
            # Set vars, open workbook
            book = app.books.open(self.cfg.path, read_only=False)
            
            app.display_alerts = self.env.debug
            app.screen_updating = self.env.debug
            self.book = book
            
            pbar.update(2) # 6


            # Create placeholder worksheets
            self._add_worksheets(progress_bar=pbar) # length of datasets + 1

            
            # Configure field analyses worksheets
            self._prep_field_analyses(progress_bar=pbar) # 2 * length of datasets
            
            # Close progress bar/section log
            pbar.close()
            self.logger.log(log_type='section',
                            details={'Production Section': 'Preparing Workbook',
                                     'Production Time in Seconds': ''})          



            # --------------------------------------------------
            # Adding Data


            with self.theme.create_progress_bar(desc="Adding Data...", 
                                                total=4 + 3*len(self.datasets)) as pbar:
    
                # Add source data
                self._add_field_analyses(progress_bar=pbar) # 3 * length of datasets
                
                # Add overview
                self._add_overview(progress_bar=pbar) # 4
                
                
                # Close section log
                self.logger.log(log_type='section',
                                details={'Production Section': 'Creating Workbook',
                                         'Production Time in Seconds': ''})
                

            # --------------------------------------------------
            # Adding Plots
            
            # Once for each dataset/plot/column
            total_iterations = len(self.datasets) + sum([ds.columns for ds in self.datasets]) + len(self.plots or {})
            
            with self.theme.create_progress_bar(desc="Adding Plots...",
                                                total=total_iterations) as pbar:
            
                # Adds additional plots
                if self.plots:
                    self._add_additional_plots(progress_bar=pbar)

                # Add plots for all fields
                self._add_plots(progress_bar=pbar)
                
                
                self.logger.log(log_type='section',
                                details={'Production Section': 'Adding Plots',
                                         'Production Time in Seconds': ''})



            # --------------------------------------------------
            # Close logging


            # Add production time to log
            self.logger.total_production_time = time.time() - self.logger.start
            
            # Close logs and 
            self.logger.close()

            # Add logging to workbook
            self._add_debug()


            # --------------------------------------------------
            # Restore app configuration, save workbook, close Excel

                    
            # Set focus to first worksheet temporarily to scroll worksheet tabs then Overview
            book.sheets[0].activate()
            book.sheets("Overview").activate()

            # Restore app settings
            app.display_alerts = True
            app.screen_updating = True
                        
            # Save/exit context manager
            book.save(self.cfg.path)
            

        # --------------------------------------------------
        # Compile closing messsage
        
        self.theme.print(f"\nProcess completed in in {int(self.logger.total_production_time)} seconds")

        exit_msg = self.cfg.exit_msg
        
        
        # Create a new Excel instance for usage
        if self.env.open_wb:

            app = xw.App(visible=True, add_book=False) 
            book = app.books.open(self.cfg.path)



        # ----------------------------------------------------------------------------------
        # Construct output messaging

        # Note dataframes          
        dataframes = ", ".join([ds.name for ds in self.datasets])
        exit_msg += f"\n\nDataframes included:\n    {dataframes}"

        # Note plots
        if self.plots:
            
            plots = (", ").join(self.plots.keys())
            exit_msg += f"\n\nAdditional plots included:\n    {plots}"

        exit_msg += f"\n\nWorkboook is located at:\n {self.cfg.path}"

        # Print closing message
        self.theme.print(exit_msg + '\n' + separator)
    
        
    
    def _validate_template(self):
        
        """
        Validates that expected objects are present in the workbook
        
        """

        
        try:
            
            # ---------------------------------------------
            # Validate worksheets and tables exist
            
            actual_objects = {}
            
            # Check actual worksheets/tables
            for sheet in self.book.sheets:
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
            self.env.warn_print(f"Template has been modifed:\n\nThe following worksheets are missing\n"
                                f"    {missing_sheets}\n\nThe following tables are missing\n    {missing_tables}")

            sys.exit()

        
    
    def _add_worksheets(self, progress_bar: tqdm):

        """
        Creates worksheets for each dataframe and any plots

        """

        book = self.book
        
        # -----------------------------------------------------------------
        # Validates that the expected template objects are present
        
        self._validate_template()
        
        progress_bar.update(1)

       
        # Add field analysis worksheets for each dataframe

        for i, dataset in enumerate(self.datasets):


            # Create a worksheet for all datasets except the first one
            if i:
                
                # Copy the sheet, rename the table
                ws = book.sheets('Field Analysis').copy(name=dataset.name)
                ws.tables[0].name = dataset.table_name

                # Add a color gradient to the worksheet tab to distinguish among them
                self.theme.greyscale_tab(ws=ws, iteration=i)
        
            progress_bar.update(1)

        
        # Use the worksheet template for the first one
        ws = book.sheets("Field Analysis")
        ws.tables[0].name = self.datasets[0].table_name
        ws.name = self.datasets[0].name



    def _prep_field_analyses(self, progress_bar: tqdm):

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
            
            # Set cursor to the first header and add header values
            ws.range("Headers_Start").select()
            ws.range("Headers_Start").value = headers
            
            
            progress_bar.update(1)



            # --------------------------------------------------
            # Adjust Named Ranges to fit dataframe size
            

            # Expand named ranges to fit number of columns
            expand_ranges = (["FieldList" + str(i) for i in range(1, 9)] + ["FieldRange", "Notes", "Definitions", "Headers"])

            for name_range in expand_ranges:
                
                self.cfg.expand_range(name=name_range, 
                                      ws=ws, 
                                      columns=ds.columns)


            
            # Resize the dataset description range and merge it
            ws.range("Description").resize(None, 2).merge()
          
                        
            
            # Show/Hide Data Size Warning
            if ds.warning:
                ws.range("Warning").value = ds.warning_msg
                self.cfg.hide_rows(ws.range("Warning"), hide=not ds.warning)

            progress_bar.update(1)
                        
            



    def _add_overview(self, progress_bar: tqdm):

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
        ws = book.sheets("Overview")
        df_overview_table = ws.tables["tbl_DfOverview"]
        field_overview_table = ws.tables["tbl_FieldOverview"]
        
        # Compile both overview dfs
        df_overview_df = pd.concat([ds.df_overview for ds in self.datasets], ignore_index=True)
        field_overview_df = pd.concat([ds.field_overview for ds in self.datasets], ignore_index=True)
        
        ws.activate()
        
        progress_bar.update(1)
        
        
        # --------------------------------------------------
        # # Add rows to the df_overview section if there is more than 2 dataframes
        
        if len(self.datasets) > 2:
    
            start_row = df_overview_table.range.last_cell.row + 2
            end_row = start_row + len(self.datasets) - 3
            row_range_string = f"{start_row}:{end_row}"
            
            ws.range(row_range_string).insert(shift="down")
        
        # Adjust named ranges to fit new data
        ws.range("Dataframes").resize(row_size=len(self.datasets) +2, column_size=None).name = "Dataframes"
        ws.range("Fields").resize(row_size=sum([ds.columns for ds in self.datasets]) +2, column_size=None).name = "Fields"
        
        # Group rows here with the new ranges
        
        self.cfg.group_rows(ws.range("Dataframes"))
        self.cfg.group_rows(ws.range("Fields"))

                        
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



    def _add_field_analyses(self, progress_bar: tqdm):

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
            source_table = ws.tables[ds.table_name]
            
            
                        
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
            except Exception:
                
                # If any unsupported columns prevent writing to Excel, convert the df to string before writing
                self.logger.log(log_type="error", 
                                details={'Detail': ds.name,
                                         'Error': 'Error writing source data that required string conversion',})
                source_table.update(df.astype(str), index=False)

            progress_bar.update(1)


            
            # --------------------------------------------------
            # Set formatting for tbl_SourceData and added columns

            
            self.cfg.set_cell_alignment(input_range=source_table.range,
                                        horizontal='center')
            record_list = source_table.range[:, :1 ]
            other_added_columns = source_table.range[:, -2: ]

            # Reduce contrast to subdue added fields
            for dimmed_range in [record_list, other_added_columns]:
                self.theme.greyscale_range(dimmed_range)


            progress_bar.update(1)

    

    def _add_plots(self, progress_bar: tqdm):
        
        """
        Adds all plots to an xleda workbook
        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object

        """

        

      
        for ds in self.datasets:
        
            # --------------------------------------------------
            # Set vars, unhide target rows

            book = self.book
            ws = book.sheets(ds.name)
            df = ds.source_data.iloc[:, 1:-2]
            ws.activate()
                

            # Set initial ranges for added plots 
            histogram_range = ws.range("Histogram")
            composition_range = ws.range("CompositionTable")
            
            # Unhide the target ranges
            self.cfg.hide_rows(histogram_range, hide=False)
            self.cfg.hide_rows(composition_range, hide=False)


            # --------------------------------------------------
            # Add plots for all except added columns

            for col in df.columns:


                # --------------------------------------------------
                # Add Composition Table

                composition_table = self.plotter.create_composition_plot(df[col])

                self.plotter.add_small_plot(target_range=composition_range,
                                            fig=composition_table,
                                            name=f'composition_{col}')

                if pd.api.types.is_numeric_dtype(df[col]):



                    # --------------------------------------------------
                    # Add Histogram

                    histogram = self.plotter.create_histogram_plot(df[col])

                    self.plotter.add_small_plot(target_range=histogram_range,
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
            self.cfg.set_text_orientation(input_range=ws.range("Toggles"))
            self.cfg.set_text_orientation(input_range=ws.range("TopToggle"), degrees=-90)

            for excel_range in ["Data_Description", "Composition", "Summary_Stats", 
                                "Percentiles", "Field_Lists", "Compiled_Lists"]:
                            
                self.cfg.hide_rows(ws.range(excel_range), hide=True)
            
            progress_bar.update(1)
            



    def _add_additional_plots(self, progress_bar: tqdm):
            
        """
        Adds additional plot worksheets

        """

        
        # Set vars
        book = self.book


        # --------------------------------------------------
        # Add additional plots



        for plot_name, figure in self.plots.items():
           
            
            # Plots will be added before all other sheets
            anchor_sheet = self.book.sheets[0]
            
            # Create a copy of the SinglePlot template, make it visible, set theme
            ws = book.sheets("SinglePlot").copy(before=anchor_sheet, name=plot_name)
            ws.visible = True
            ws.activate()
            self.theme.set_theme(ws.range("Theme"))

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



    def _add_debug(self):

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
        ws.tables("tbl_debug_errors").update(self.logger.errors, index=False)


        # Hide debug section and set toggle orientation
        self.cfg.hide_rows(ws.range("Debug"), hide=True)
        self.cfg.set_text_orientation(ws.range("DebugToggle"))
        


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
        # Open Excel/export data using a context manager


        with xw.App(visible=self.env.debug, add_book=False) as app:
            
            app.display_alerts = False
            app.screen_updating = self.env.debug
            
            
            book = app.books.open(self.cfg.path, read_only=False)
            self.book = book
            
           

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
                    # Extract source_data, df_overview, field_overview
                    
                    ds.source_data = ws.tables[ds.table_name].range.options(pd.DataFrame, index=False).value     
                    ds.df_overview = self.book.sheets("Overview").tables["tbl_DfOverview"].range.options(pd.DataFrame, index=False).value
                    ds.field_overview = self.book.sheets("Overview").tables["tbl_FieldOverview"].range.options(pd.DataFrame, index=False).value
                    

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


        # Add an ExportDict object for each dataset into self.export_dicts
        self.export_dicts = [ExportDict(ds) for ds in self.datasets]



        # --------------------------------------------------
        # Print closing output
        
        
        # Note any datframes that weren't exported
        if missing_dfs:
            
            self.env.warn_print("\nThe following worksheets were not found and are using default metadata:\n")
            for sht in missing_dfs:
                self.env.warn_print(f"    {sht}")
                        

        duration = time.time() - self.logger.start
        self.theme.print(f"\nExport completed after {int(duration)} seconds" + separator)
        
        
        
        
        
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
def cli_wb(self,
           data: str = typer.Argument(..., help="Path to a supported data file"),
           file_name: str = typer.Option(None, show_default=False, help="Name of the created workbook. Defaults to the same name as the data file"),
           theme_color: str = typer.Option(None, "--theme_color", help="Hex color used for theme in workbook. Using this setting will change the default.  Defaults to a neutral color"),
           export: bool = typer.Option(False, help="Export from an xleda workbook"),
           large_report: bool = typer.Option(False, "--large_report", help="Only subsample when required to fit within Excel's worksheet limits"),
           overwrite: bool = typer.Option(False, help="Overwrite existing workbook"),
           wb_path: Path = typer.Option('', "--wb_path", show_default=False, help="Workbook directory with/without filename"),
           open_wb: bool = typer.Option(True, "--open_wb", help="Automatically open the workbook on finish"),
           no_vba: bool = typer.Option(None, "--no_vba", help="Create an xlsx file without VBA.  Using this setting will change the default.  Defaults to False")):

        """
        Creates an xleda workbook from a supported data file

        """
        
        # Create a dictionary of input arguments
        cli_args = locals()
        
        # if file name isn't provided, extract it from the data argument 
        if not wb_path:
            cli_args['file_name'] = file_name
        

        return wb(**cli_args)