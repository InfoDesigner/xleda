import pandas as pd
from pathlib import Path
import ast
import sys
import time

from tqdm.auto import tqdm
from matplotlib.figure import Figure
import xlwings as xw

from .utilities import (Environment, Config, Theme, Plotter,
                        Blueprint, ExportDict, PerformanceLogger)



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
        # Configure xleda


        # Initialize/Check Environment
        self.env = Environment()



        # Initialize logger/log input variables
        self.logger = PerformanceLogger(input_args=locals().copy(),
                                        env=self.env)


        # Intialize theme
        self.theme = Theme(theme_color,
                           env=self.env)

        # Intialize plotter
        self.plotter = Plotter(theme=self.theme,
                               env=self.env)


        # Intialize Config
        self.cfg = Config(wb_path=wb_path,
                          input_df=input_df,
                          theme=self.theme,
                          env=self.env,
                          debug=debug,
                          large_report=large_report,
                          overwrite=overwrite,
                          no_vba=no_vba,
                          name=name,
                          add_plots=add_plots,
                          add_dfs=add_dfs,
                          open_wb=open_wb)
        

        # Set placeholder vars
        self.blueprints: list[Blueprint] = self.cfg.blueprints
        self.export_dicts: list[ExportDict] = [ExportDict(bp) for bp in self.blueprints]
    
       

        # ------------------------------------------------------------------------------
        # Either create a workbook or export

        if export:
            self._export_analysis()
        else:
            self._assemble_wb()




    def _assemble_wb(self):
        
        """
        Assembles an xleda workbook

        """
        
        # --------------------------------------------------   
        # Construct/print entry message
        

        entry_message = separator

        if self.cfg.name == 'xleda':
            entry_message += f"\nPreparing an xleda workbook located at:\n    {self.cfg.path}\n"
        else:
            entry_message += f"\nPreparing an xleda workbook with {self.blueprints[0].title}"
            entry_message += f" data located at:\n    {self.cfg.path}"

        entry_message += (f"\n\nProcess started at {time.strftime('%H:%M:%S')}\n")

        # Print entry message
        self.theme.print(entry_message)
        

        # --------------------------------------------------
        # Create the template and initial progress bar
        
        total_iteratons = 7 + (len(self.blueprints[1:])) + (2*len(self.blueprints))
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
                                                total=8*len(self.blueprints)) as pbar:
                
                
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
            
            total_iteratons = sum([bp.columns for bp in self.blueprints]) + len(self.cfg.additional_plots)
            with self.theme.create_progress_bar(desc="Adding Plots...",
                                                total=total_iteratons) as pbar:
            
                # Adds additional plots
                if self.cfg.additional_plots:
                    self._add_additional_plots(progress_bar=pbar)


                self._add_plots(progress_bar=pbar)
                
                
                self.logger.log(log_type='section',
                                details={'Production Section': 'Adding Plots',
                                         'Production Time in Seconds': ''})


            # --------------------------------------------------
            # Configuring Workbook

            total_iteratons = (2* min(10, self.blueprints[0].columns)) + 2 + len(self.blueprints)

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
                                details={'Dataframes Included': len(self.blueprints),
                                         'Plots Included': len(self.cfg.additional_plots) + len(self.logger.performance_logs['plots']),
                                         'Columns Included': sum([bp.columns for bp in self.blueprints]),
                                         'Rows Included': sum([bp.rows for bp in self.blueprints]),
                                         'Production Time': self.logger.total_production_time})
                
                
                # Close logs and write to workbook
                self.logger.close(blueprints=self.blueprints,
                                  additional_plots=len(self.cfg.additional_plots))
                
                self._add_debug()


                # --------------------------------------------------
                # Restore app configuration, save workbook, close Excel

                        
                # Set focus to primary field analysis worksheet
                book.sheets(self.blueprints[0].field_analysis).activate()

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

        
        if self.cfg.open_wb:
            

            # Create a new Excel instance
            app = xw.App(visible=True, add_book=False) 

            # Open the workbook
            book = app.books.open(self.cfg.path)


        # Note additional plots
        if self.cfg.additional_plots:
            
            exit_msg += "Additional plots included:"
            
            for plot in self.cfg.additional_plots:
                exit_msg += f"    {plot['title']}\n"
        
        # Note additional DFs
        if len(self.blueprints) > 1:
            
            exit_msg += "\nAdditional dataframes included:\n"
            
            for df_title in [bp.title for bp in self.blueprints[1:]]:
                exit_msg += f"    {df_title}\n"


        
       
        # Print closing message
        self.theme.print(exit_msg + '\n' + separator)
    


    def _create_worksheets(self, progress_bar: tqdm):

        """
        Creates worksheets for any additional dataframes

        """

        book = self.book
        
        # -----------------------------------------------------------------
        # Validate template worksheets are present

        expected_sheets = ['Field Analysis', 'Overview', 'Pivot', 'debug', 'SinglePlot', 'MultiplePlots']
        actual_sheets = [sht for sht in expected_sheets if sht in book.sheet_names]
        missing_sheets = [sht for sht in expected_sheets if sht not in book.sheet_names]

        if expected_sheets != actual_sheets:
            self.theme.warn_print(f"Template has been modifed, the following worksheets are missing\n    {missing_sheets}")
            sys.exit()


        # If there are additonal dataframes, add worksheets for them
        if len(self.blueprints) > 1:

            for i, bp in enumerate(self.blueprints[1:]):
                        
                # Copy the sheet, rename the table
                book.sheets('Field Analysis').copy(name=bp.field_analysis)
                book.sheets(bp.field_analysis).tables[0].name = f'tbl_SourceData_{bp.name.replace(" ","")}'

                # Copy the sheet, rename the table
                book.sheets('Overview').copy(name=bp.overview)
                book.sheets(bp.overview).tables[0].name = f'tbl_Overview_{bp.name.replace(" ","")}'
                
                # Add a color gradient to the tabs to distinguish among them
                self.theme.greyscale_tabs(bp=bp,
                                          book=self.book,
                                          iteration=i)
        
                progress_bar.update(1)

        
        # Care for the default worksheets
        bp = self.blueprints[0]


        book.sheets("Field Analysis").name = bp.field_analysis
        book.sheets("Overview").name = bp.overview
        book.sheets("Pivot") .name = bp.pivot

        progress_bar.update(1)


    def _configure_field_analyses(self, progress_bar: tqdm):

        """
        Configures named ranges on all Field Analysis worksheets

        """

        
        for bp in self.blueprints:

            # --------------------------------------------------
            # Set variables

            book = self.book
            ws = book.sheets(bp.field_analysis)
            ws.activate()


            # --------------------------------------------------
            # Format metadata placeholders

            # Set worksheet theme/name
            self.cfg.expand_range(name="Theme", ws=ws, columns=bp.columns + 2)
            self.theme.set_theme(ws.range("Theme"))
            ws.range("Name").value = bp.title

            # Add all header values except Record List
            headers = bp.source_data.columns.to_list()[1:]
            ws.range("Headers_Start").value = headers


            columns_to_format = bp.columns -3
            if columns_to_format > 0:
                format_from = ws.range("FormatRange")
                format_to = (ws.range("FormatRange").offset(0, 1).resize(None, columns_to_format))
                format_from.copy()
                format_to.paste(paste='formats')

            # Clear clipboard and move selection back to upper left
            book.app.cut_copy_mode = False
            
            progress_bar.update(1)



            # --------------------------------------------------
            # Adjust Named Ranges to fit dataframe size
            

            # Expand named ranges to fit number of columns and set input prompt for notes/definitions
            expand_ranges = (["FieldList" + str(i) for i in range(1, 9)] + ["FieldRange", "Notes", "Definitions", "Headers"])

            for name_range in expand_ranges:
                
                self.cfg.expand_range(name=name_range, 
                                      ws=ws, 
                                      columns=bp.columns)

                # Add placeholder values if necessary.
                if name_range in ["Notes", "Definitions"]:    
                    ws.range(name_range).value = name_range[:-1]

                elif name_range.startswith("FieldList"):
                    ws.range(name_range).value = "FALSE"


            # Set FieldLists range so that it also includes the names on the left
            # Used to recreate completed examples
            self.cfg.expand_range(name="FieldLists", 
                                  ws=ws, 
                                  columns=bp.columns + 1)


            # Show/Hide Data Size Warning
            if bp.warning:
                
                ws.range("Warning").value = bp.warning_msg
                self.cfg.hide_rows(ws.range("Warning"), hide=not bp.warning)

            progress_bar.update(1)



    def _add_field_metadata(self, progress_bar: tqdm):

        """
        Adds field metadata to Field Analysis workbook

        Parameters
        ----------

        progress_bar
            A tqdm progress bar object
        
        """



        for bp in self.blueprints:

            # --------------------------------------------------
            # Set variables

            df = bp.field_metadata
            book = self.book
            ws = book.sheets(bp.field_analysis)
            ws.activate()



            # --------------------------------------------------
            # Add basic metadata

            # Add Data Description metadata
            ws.range("Dimensions").options(transpose=True).value = list(bp.df_metadata.values())
            
            progress_bar.update(1)




            # --------------------------------------------------
            # Split base analysis into Field Analysis sections

            composition_df = df.loc[["Data type", "Distinct %", "Missing %", "Memory Usage %", "Memory Usage", "Distinct", "Count", "Missing"]]
            summary_stats_df = df.loc[["Mean", "Median", "Mode", "Standard Deviation", "Variance"]]
            percentiles_df = df.loc[["Min", "5%", "25%", "50%", "75%", "95%", "Max", "Range", "IQR"]]

            progress_bar.update(1)



            # --------------------------------------------------
            # Add Field Analysis sections to workbook

            ws.range("Composition")[0, 0].offset(0,1).value = composition_df.values
            ws.range("Summary_Stats")[0, 0].offset(0,1).value = summary_stats_df.values
            ws.range("Percentiles")[0, 0].offset(0,1).value = percentiles_df.values

            progress_bar.update(1)







    def _add_overview(self, progress_bar: tqdm):

        """
        Adds overview_metadata to all Overview worksheets
        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object
        
        """


        for bp in self.blueprints:


            # --------------------------------------------------
            # Set variables and theme

            book = self.book
            df = bp.overview_metadata
            ws = book.sheets(bp.overview)
            overview_table = ws.tables[0]

            col_order = ['Field', 'Definition', 'Field Notes', 'Data type', 'Distinct %', 'Missing %', 'Memory Usage %', 'Memory Usage', 'Distinct', 'Count', 'Count %', 'Missing', 'Mean', 'Median', 'Mode', 'Standard Deviation', 'Variance', 'Min', '5%', '25%', '50%', '75%', '95%', 'Max', 'Range', 'IQR']


            # Set name/theme/warning if necessary
            ws.range("Name").value = bp.title
            self.theme.set_theme(overview_table.range[0,:])

            # Show/Hide Data Size Warning
            if bp.warning:
                ws.range("Warning").value = bp.warning_msg
                self.cfg.hide_rows(ws.range("Warning"), hide= not bp.warning)



            # --------------------------------------------------
            # Configure overview df, reorder columns

            df['Field Notes'], df['Definition'], df['Field'] = None, None, df.index
            df = df[col_order]


            progress_bar.update(1)



            # --------------------------------------------------
            # Add metadata/overview table, configure the Field Notes/Definitions columns

            ws.range("Metadata").options(transpose=True).value = list(bp.df_metadata.values())
            overview_table.update(df, index=False)
            
            # -------------------------------------------------------------------------
            # Add Formulas to pull field definitions, notes into overview table
            
            source_ws_name = bp.field_analysis
            definitions_formula = f'''=IF(INDEX('{source_ws_name}'!Definitions,1,MATCH([@Field],'{source_ws_name}'!Headers,0))="Definition","",INDEX('{source_ws_name}'!Definitions,1,MATCH([@Field],'{source_ws_name}'!Headers,0)))'''
            notes_formula = f'''=IF(INDEX('{source_ws_name}'!Notes,1,MATCH([@Field],'{source_ws_name}'!Headers,0))="Note","",INDEX('{source_ws_name}'!Notes,1,MATCH([@Field],'{source_ws_name}'!Headers,0)))'''

            overview_table.range[1:, 1][0].formula = definitions_formula
            overview_table.range[1:, 2][0].formula = notes_formula
            


            progress_bar.update(1)




    def _add_source_data(self, progress_bar: tqdm):

        """
        Adds source_data to all Field Analysis worksheets

        
        Parameters
        ----------

        progress_bar
            A tqdm progress bar object

        """


        for bp in self.blueprints:

            # --------------------------------------------------
            # Set variables, dtype conversion if necessary

            book = self.book
            ws = book.sheets(bp.field_analysis)
            df = bp.source_data


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
                                         'Detail': bp.name})
                source_table.update(df.astype(str), index=False)

            progress_bar.update(1)



            
            # --------------------------------------------------
            # Set formatting for tbl_SourceData and added columns

            
            self.cfg.set_cell_alignment(input_range=source_table.range.api.HorizontalAlignment,
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

        

      
        for bp in self.blueprints:
        
            # --------------------------------------------------
            # Set vars, start performance logging

            book = self.book
            ws = book.sheets(bp.field_analysis)
            df = bp.source_data.iloc[:, 1:-3]
                

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
                
                self.logger.log(log_type='plots',
                                details={'Dataset': bp.title,
                                         'Field': col,
                                         'Activity': 'Add Composition Table',
                                         'Rows': bp.rows,
                                         'Columns': bp.columns,
                                         'Datatype': str(df[col].dtype),
                                         'Production Time in Seconds': ''})




                if pd.api.types.is_numeric_dtype(df[col]):



                    # --------------------------------------------------
                    # Add Histogram Multiple

                    histogram = self.plotter.create_histogram_plot(df[col])

                    self.plotter.add_small_plot(target_range=histogram_range,
                                                fig=histogram,
                                                name=f'histogram_{col}')

                    self.logger.log(log_type='plots',
                                    details={'Dataset': bp.title,
                                             'Field': col,
                                             'Activity': 'Add Histogram',
                                             'Rows': bp.rows,
                                             'Columns': bp.columns,
                                             'Datatype': str(df[col].dtype),
                                             'Production Time in Seconds': ''})



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



        for plot in self.cfg.additional_plots:
           
            
            # Plots will be added after last primary dataframe sheet
            after_sheet = self.book.sheets(self.blueprints[0].pivot)
            
            # Create a copy of the SinglePlot template, make it visible, set theme
            ws = book.sheets("SinglePlot").copy(after=after_sheet, name=plot['name'])
            ws.visible = True
            self.theme.set_theme(ws.range("Theme"))

            # Set target range, title, and autofit title range
            plot_range = ws.range("SinglePlot")
            ws.range("SinglePlotTitle").value = plot['title']
            ws.range("SinglePlotTitle").columns.autofit()
            
            ws.pictures.add(plot['fig'], 
                            name=plot['name'], 
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

        bp = self.blueprints[0]
        df = bp.input_df.copy()
        book = self.book

        # Fields to be added to the pivot tables
        pivot_fields = df.columns.to_list()[:min(10, bp.columns)]
        
        if bp.pivot is not None:
            ws = book.sheets(bp.pivot)

        pt = self.cfg.get_updated_pivot_table(ws=ws)
        


        # --------------------------------------------------
        # Set theme/show warning if necessary

        ws.activate()
        self.theme.set_theme(ws.range("Theme"))

        if bp.warning:
            ws.range("Warning").value = bp.warning_msg
            self.cfg.hide_rows(ws.range("Warning"), hide=not bp.warning)

        

        # Show field limit warning for datasets with >10 columns
        if len(df.columns) <= 10:
            ws.range("DefaultWarn").value = ''
        else:
            ws.range("DefaultWarn").value = 'Pivots only show first 10 fields by default'


        progress_bar.update(1)



        # --------------------------------------------------
        # Add fields to pivot table


        for field in pivot_fields:

            # Add field to row section of pivot table
            pt.PivotFields(field).Orientation = xw.constants.PivotFieldOrientation.xlRowField
            

            # Remove subtotals if possible
            try:
                pt.PivotFields(field).Subtotals = tuple([False] * 12)
            except Exception:
                pass
            
            # Log performance, update progress bar
            self.logger.log(log_type='pivot',
                            details={'Dataset': bp.title,
                                     'Field': field,
                                     'Activity': 'Add Pivot Field',
                                     'Rows': bp.rows,
                                     'Columns': bp.columns,
                                     'Datatype': str(df[field].dtype),
                                     'Production Time in Seconds': ''})

                                
            progress_bar.update(1)
                


            # --------------------------------------------------
            # Format pivot tables
            
            # Collapse fields if possible, starting from the end
            for field in pivot_fields[::-1]:
                try:
                    
                    if self.env.win:
                        pt.PivotFields(field).ShowDetail = False

                    elif self.env.mac:
                         pt.PivotFields(field).ShowDetail.set(False)

                except Exception:
                    pass

            
            # --------------------------------------------------
            # Set ranges for formatting
            
            if self.env.win:
                pivot_range = ws.range(pt.TableRange1.Address)

            elif self.env.mac:
                pivot_range = ws.range(pt.table_range1.get_address())

            pivot_headers = pivot_range[0,:]            
            header_column = pivot_range[: , 0]

            
            
            # --------------------------------------------------
            # Format headers, first column, and column width

            pivot_headers.wrap_text = True

            self.cfg.set_cell_alignment(input_range=pivot_headers,
                                        horizontal='center',
                                        vertical='center')
            
            self.cfg.set_cell_alignment(input_range=header_column,
                                        horizontal='left')
            
            pivot_range.columns.autofit()

            

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


        # Loop through all blueprints and configure Field Analysis worksheets

        for bp in self.blueprints:
            ws = book.sheets(bp.field_analysis)
            ws.activate()
            ws.range('Headers_Start').select()

            # Orient toggles, and collapse subsections

            self.cfg.set_text_orientation(input_range=ws.range("Toggles"))
            self.cfg.set_text_orientation(input_range=ws.range("TopToggle"), degrees=-90)

            for excel_range in ["Data_Description", "Composition", "Summary_Stats", 
                                "Percentiles", "Field_Lists", "Compiled_Lists"]:
                            
                self.cfg.hide_rows(ws.range(excel_range), hide=True)

            progress_bar.update(1)



    def _add_debug(self):

        """
        Adds debug data to the debug worksheet
        
        """
        
        book = self.book
        
        ws = book.sheets('debug')
        ws.activate()
       
        
        # Write debug range
        ws.range("debug_metadata").options(transpose=True).value = self.logger.performance_metadata.values


        # Write debug tables
        ws.tables("tbl_debug_environment").update(self.logger.env, index=False)
        ws.tables("tbl_debug_config").update(self.logger.config, index=False)
        ws.tables("tbl_debug_errors").update(self.logger.errors, index=False)
        ws.tables("tbl_debug_section").update(self.logger.section_performance, index=False)
        ws.tables("tbl_debug_field").update(self.logger.field_performance, index=False)


        # Hide config section and set toggle orientation
        self.cfg.set_text_orientation(input_range=ws.range("toggle"))
        self.cfg.hide_rows(ws.range("Config"), hide=True)
        
        
        # Move the debug worksheet to the end
        if self.env.win:
            ws.api.Move(Before=None, After=book.sheets[-1].api)
            
        elif self.env.mac:
            ws.api.move(after=book.sheets[-1].api)



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
                                                total=4 + (2 * len(self.blueprints))) as pbar:
                              
                pbar.update(2)

                try:
                    wb = app.books.open(self.cfg.path)
                    pbar.update(2)

                except FileNotFoundError:
                    self.theme.print(f"File not found at {self.cfg.path}")
                    sys.exit()

                
                for bp in self.blueprints:

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
                    
                    bp.altered_source_data = ws.tables(1).range.options(pd.DataFrame, index=False).value

                    

                    # --------------------------------------------------
                    # Prepare exports

                    # Add export properties
                    bp.description = description
                    bp.definitions = definitions
                    bp.notes = notes
                    bp.lists = lists
        
                    pbar.update(1)
            
            app.display_alerts = True
            app.screen_updating = True


        # Add an ExportDict object for each dataframe into self.export_dicts
        self.export_dicts = [ExportDict(bp) for bp in self.blueprints]

        

        # --------------------------------------------------
        # Print closing output
        
        
        if missing_dfs:
            
            self.theme.warn_print("\nExports for the following dataframes were not found:\n")
            for sht in missing_dfs:
                self.theme.warn_print(f"    {sht}")
                        

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
        self.theme.warn_print(separator + 
              "\nCreating a separate Field Analysis configuration is no longer required.\n"
              "Use xleda.wb to both configure and create a workbook\n" + separator)

        super().__init__(input_df=input_df,
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


