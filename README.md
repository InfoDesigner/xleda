<div align="center">

[![License: Apache](https://img.shields.io/badge/license-Apache-**blue**)](https://www.apache.org/licenses/LICENSE-2.0.txt)
[![PyPI - Version](https://img.shields.io/pypi/v/xleda)](https://pypi.org/project/xleda)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/xleda.svg)](https://pypi.org/project/xleda)
[![Downloads](https://static.pepy.tech/badge/xleda)](https://pepy.tech/project/xleda)
[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ffdd00?&logo=buy-me-a-coffee&logoColor=black)](https://buymeacoffee.com/informationdesigner)

</div>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/logo.webp?raw=true"  width="250" alt="Example Top View"> 
	<br>
</p>
<p align="center" style="font-size: 26px; font-weight: bold;">A Microsoft Excel powered EDA tool for Python data.</p>

* Produces Microsoft Excel workbooks from pandas dataframes that are highly optimized to both perform and document [the activity of Exploratory Data Analysis](https://www.geeksforgeeks.org/data-analysis/what-is-exploratory-data-analysis/) .

* Visually explore your data, navigate with your keyboard, take field or record notes, create lists of fields/records for editing, round-trip your edits/analysis back into python, share your workbook with other contributors.

* There are some amazing EDA tools for Python.  You shouldn't have to start from scratch to include Microsoft Excel among them.

* xleda provides a good start to a robust EDA.  

* See [some example xleda workbooks](https://github.com/InfoDesigner/xleda/tree/main/examples).<br><br>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/top_view.webp?raw=true"  width="600" alt="Example Top View"> 
	<br>
	<em>An xleda workbook made with diamond data.</em>
</p>
<br><br>

# **Requirements/Compatibility**

* Requires the full version of Microsoft Excel to create workbooks. 

* Once created, xleda workbooks should work in anything that reads Microsoft Excel workbooks.

* It has been developed and tested on Windows.  

* It should also work on MacOS though this has not yet been tested.<br>

<br><br>


# **Installation**

* Install with 

	```bash
	uv add xleda
	```
	or 
	```bash
	pip install xleda
	```
<br><br>


# **Quick Start**

* Use wb() to quickly create an xleda workbook of a dataframe. 

* See the configuration section below for how to name the workbook, set a theme, add plots etc.  <br><br>

	```python
	from xleda import wb
	import seaborn as sns

	# < your dataframe goes here >
	df = sns.load_dataset("titanic")

	# Creates xleda.xlsm in the current directory
	wb(df)
	```
<br><br>

# **Basic xleda Components**
### **Field Metadata**

* Most of the field metadata comes from the built-in pandas features [describe](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.describe.html), [info](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.info.html), and **[quantile](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.quantile.html)**<br><br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/basic_metadata.webp?raw=true"  width="600" alt="Basic Metadata"> 
	<br>
	<em>Included Field Metadata</em>
</p><br>


### **Overview**

* The `Overview` worksheet rotates the field metadata 90 degrees so that you can sort/filter fields by their name, metadata, or your notes/definitions/etc. <br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/overview.webp?raw=true"  width="600" alt="Composition Table"> 
	<br>
	<em>Sorting fields from MLB data by memory usage.</em>
</p><br>

### **Per Field Charts**

Two charts are produced for each column in your dataframe.
1. A composition table showing the top 5 values per column and their percentages.
2. A histogram/KDE showing min/max, distribution, and mean<br>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/default_charts.webp?raw=true"  width="300" alt="Default Charts"> 
	<br>
	<em>Default charts for MLB player height.</em>
</p><br>

### **Source Data Table**

* A copy of your source data is included as an Excel table so you can visually inspect it, sort/filter it, etc. 

* Includes a way to make lists of individual records. 

* Includes a `HasBlank` field  for filtering records with that have blank values.

* Includes a way to round-trip your source data back into Python so that you can use Excel to replace values, delete records/columns/etc.  

* More on how to use these features below.<br><br>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/source_data.webp?raw=true"  width="600" alt="Source Data"> 
	<br>
	<em>Source Data Table for MLB player data.</em>
</p><br>

### **Pivot**

* A bare-bones pivot table configured to drill down into where blank values are.
* The first 10 fields of the source data are added by default. <br><br>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/blanks.webp?raw=true"  width="600" alt="Blanks"> 
	<br>
	<em>Bare-bones pivot table for Titanic survivor data.</em>
</p><br><br>


# **xleda.wb() Configuration**

### **input_df** | Dataframe | Mandatory

* A pandas dataframe of any size<br>

### **name** | str | Optional

* Name of the dataset/file name of the created workbook
* Defaults to `xleda`<br>

### **theme_color** | str | Optional

* Sets the primary color of the charts and the color of the headings in the workbook to a hex color of your choice.  `theme_color="random"` sets a random theme<br>
* Defaults to a neutral color <br>
<br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/theme_colors.webp?raw=true" width="600" alt="Theme Colors">
	<br>
	<em>theme_color affects the workbooks and default charts.</em>
</p><br>

### **add_plots** | dict | Optional

* `add_plots={'plotname': Figure, ...}` will add additional worksheets with plots of your choosing. 

* No styling/sizing of additional plots is performed by xleda. 

* The example below adds two additional plot worksheets, one from seaborn and another from missingno.  The workbook can be found [here](https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/Penguins.xlsm).


	```python
	from xleda import wb

	import matplotlib.pyplot as plt
	import seaborn as sns
	import missingno as msno
	

	# < your dataframe goes here >
	df = penguins = sns.load_dataset("penguins")
	

	# Style the additional plots | optional
	plt.style.use("dark_background")


	# Create additional plots
	pair_plots = sns.pairplot(df, hue="species").figure
	null_matrix = msno.matrix(df).get_figure()


	# Resize the null matrix  | optional
	null_matrix.set_size_inches(9.35, 4.5)


	# Creates an xleda workbook named Penguins.xlsm with two extra plot sheets
	wb(input_df=df,
	   name="Penguins",
	   theme_color="#4C4C4C",
	   add_plots={'Pair Plots': pair_plots,
	              'Null Matrix': null_matrix})
	```
<br>

### **wb_path** | Path | Optional

* Sets the target folder for the xleda workbook. 
* Uses a pathlib Path object.
* Defaults to current working directory<br>

	```python
	from xleda import wb
	from pathlib import Path


	# Creates "c:\my_target_folder\Penguins.xlsm"
	wb(input_df=df,
	   name="Penguins",
	   wb_path=Path(r"c:\my_target_folder"))

	```
	<br>

### **overwrite** | bool | Optional

* Whether to overwrite existing workbooks of the same name. 
* Existing workbooks are sent to Trash/Recycle Bin
* Defaults to False<br>


### **large_report** | bool | Optional

* Raises the default dataframe size limits of 25,000 rows/50 columns to Excel's limits of 1,000,000 rows and 16,000 columns.  
* The closer your are to this limit, the more RAM and patience you'll need to produce a workbook.
* See additional details in the performance section below.
* Defaults to False<br>

### **no_vba** | bool | Optional

* Creates the workbook as an xlsx file without VBA. 
* Defaults to False<br>


### **open_wb** | bool | Optional

* Opens the workbook after creating.
* Setting this to `False` is useful when creating multiple workbooks
* Defaults to True.<br>

### **export** | bool | Optional

* Performs an export from an xleda workbook instead of creating one. 
* Replaces the `export_analysis` method.
* See details below.
* Defaults to False.<br><br>


# **Usage Notes**

## **Performance**

* On an average machine, xleda creates workbooks for most data sets less than 20 seconds. 

* Performance is largely dependent on how powerful of a machine you have and how large your dataframe is.  

* There is a detailed output provided when creating an xleda workbook that does a pretty good job of letting you know what it's doing. 

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/create_example.webp?raw=true" width="400" alt="Create Example">
	<br>
	<em>Console output of a Planets workbook </em>
</p><br><br>

### **Limits with Large Data Sets**

* To ensure workbooks are created quickly, defaults limit data to the first 50 columns and a random sample of 25,000 records. 

* You can optionally override this to Excel's limits (see `large_report=True` above)

* You'll see a warning banner if you hit a limit.

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/warning.webp?raw=true" width="600" alt="Create Example">
</p><br>

* One of the more complex data sets tested was a 600 column/1,200 row dataframe.
  
	* It took ~5 minutes to create, in part because nearly all values are large unique numbers in all 600 columns.

	* It is still snappy to use even though it has 1,200 charts on a single worksheet.

	* That example is [here](https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/African%20Soil.xlsm).<br><br>

## **Field/Record Lists**


* The `Field Lists` section helps you create lists of the fields in your data.
	* e.g. lists of fields to rename, delete, standard scale, encode, impute, investigate, etc. 

* Anything not marked as `False` will be included in each list.   

* You can rename any list to `Anything You Want` and the list will be renamed to `anything_you_want`.

* The `Record List` field added to your source data works the same way except it creates a list of records instead of a list of fields.  More on that below.

* The Compiled Lists section formats your lists as python lists for easy copy/pasting.<br>
<br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/field_lists.webp?raw=true" width="500" alt="Field Lists">
	<br>
	<em>Easily create lists of fields in your data.</em>
</p>
<br><br>

## **Columns Added to Source Data**

* Although the source data is unchanged before it goes into Excel,  there are some columns added to support an EDA workflow. 

	* `HasBlank`: If any field in a record has a missing value, this will show 1 otherwise 0

	* `Record Hash`:  Uses a built-in pandas feature [hash_pandas_object](https://pandas.pydata.org/docs/reference/api/pandas.util.hash_pandas_object.html) to uniquely identify records.  If two records share all column values they also share a `Record Hash`. 

	* `Record List`:  Used to create a list of `Record Hash` values.  Like `Field Lists` above, anything not marked false gets added to a list.

	* `index`:  This is a copy of the index from the provided dataframe as a column.<br><br>

	```python
	# .....After editing your workbook and assuming you marked 
	# records to delete in the record list column
	
	# exports your source dataframe with the added record list column
	export_dict = wb(df, export=True).export_dict
	df = export_dict['source_data']
	
	# Uses the record list to delete records
	delete_records = df = export_dict['lists']['record_list']
	df = df[not df['Record Hash'].isin(delete_records)]
	```
<br><br>


## **Accessing Metadata in Python**


##### **Default Metadata**

* Metadata from  all `xleda.wb()` objects is collected into a dictionary accessible through `xleda.wb().export_dict` <br><br>
	```python
	# Creates "Titanic.xlsm" and exports the metadata dictionary
	export_dict = wb(input_df=df,
	           		 name="Titanic").export_dict

	# Returns the field metadata df used in the Field Analysis worksheet
	export_dict['field_metadata'] 
	```
<br>

* The following metadata is available without using `export=True`

	* `field_metadata`: A basic metadata dataframe, combining information from pandas info/describe/quantile.
	* `overview_metadata`: A transposed copy of the field_metadata.
	* `source_data`: A copy of your unaltered source data that includes `Record Hash`/`Record List`/`HasBlank`/`index` columns.<br><br>

##### **Expanded Metadata**

* Using `xleda.wb(export=True)` reads an xleda workbook instead of creating one.  

* It expands the available metadata within `xleda.wb().export_dict` to include:

	* `description`: Dataframe description if you've added one
	* `definitions`: Any field definitions you've added.
	* `notes`: Any field notes you've added
	* `lists`: Any lists showing in the compiled lists section
	* `altered_source_data`: Reads the Excel table named `tbl_SourceData` from the workbook and will include any manual edits you've made such as removing records, renaming fields, replacing values, etc. **

		 ** *Note that data types will likely change in the round-trip translation.* <br><br>


### **Completed Example**

* The xleda workbook pictured here is used in for the export code example below .  

* It can be found [here.](https://github.com/InfoDesigner/xleda/blob/main/examples/Titanic%20(Completed).xlsx).

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/completed_field_analysis.webp?raw=true" width="500" alt="Completed Field Analysis">
	<br>
	<em>A completed xleda workbook of Titanic passenger showing definitions, notes, lists, etc.</em>
</p>
<br>

### **Example Export Dictionary**

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/completed_analysis_export.webp?raw=true" width="400" alt="Export Dict">
	<br>
	<em>An example export dictionary from a completed field analysis on Titanic passenger data.</em>
</p>
<br><br>

### **Example Export Code**

* This example exports everything from an xleda workbook named  "Titanic (Completed).xlsx" in the current directory.  

* Either download [this one](https://github.com/InfoDesigner/xleda/blob/main/examples/Titanic%20(Completed).xlsx) or create your own.

	```python
	from xleda import wb

	# Performs a full export from "Titanic (Completed).xlsx"
	export_dict = wb(input_df=df,
					 name="Titanic (Completed)", 
					 no_vba=True,
					 export=True).export_dict

	
	# Returns dict_keys(['description', 'definitions', 'notes', 
	# 'lists', 'field_metadata', 'overview_metadata', 'source_data', 
	# 'altered_source_data'])
	print(export_dict.keys()) 
	```
<br>

## **VBA Code**

* If you can't or don't want to enable VBA, you may want to use `no_vba=True` which creates an  xlsx file that contains no VBA.<br>

* The [small amount of VBA code](https://github.com/InfoDesigner/xleda/blob/main/src/xleda/mdl_xleda.bas) in the template does two optional things. 

	1. Makes the sections expand/collapse when you select them as pictured above.  

	2. Adds two UDFs, PythonList and PythonDict, that format cell values as lists/dicts. <br><br>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/row_groupings.webp?raw=true" width="462" alt="Row Groupings">
	<br>
	<em>Use row groupings to navigate if you can't use VBA.</em>
</p>
<br>

## **Extensible**

* xleda is only meant to give a good start to EDA.  

* If it accomplishes one thing it will be to give you a way to quickly get Python data into Excel so that you can make sense of it...without making you do everything from scratch.  

* Where you go from there is up to you.

* Because it's an ordinary workbook, you can use any tool that works with Microsoft Excel workbooks to do more.  

* [xlwings](https://www.xlwings.org/) is recommended if you do. <br><br>

## **Built With**

* This was primarily built with Python, xlwings, Pandas, and of course, Microsoft Excel.

## **Roadmap**

* [x] Add a barebones pivot that is ready to configure
* [x] Make xleda even more accessible by simplifying the API and making it easier to remember.
* [x] Create a way to quickly view dataframe data that is editable, shareable, and presentation ready.
* [x] Add a way to include extra plots for a dataset.
* [ ] Add a way to include extra supporting table worksheets.
* [ ] Add a way to use on desktop files e.g. by right-clicking csv/parquet files/other tabular data files. 
* [ ] Add a basic version for even quicker dataframe inspection.
* [ ] Add a way to include multiple xleda analyses in a single workbook.
* [ ] Test on MacOS
* [ ] Investigate starting from Excel data.
* [ ] Your idea here.



## Changelog


> [!NOTE]Version 0.8.185 - New simplified API, simplified export, general polish
> **Simplified basic usage to make it quicker to use and easier to memorize.** 
> * Changed the default entry point to `xleda.wb()` from `xleda.FieldAnalysis()`.  
> * `xleda.wb()` now creates and automatically opens workbooks.
> * The only argument needed to create a workbook is now a dataframe. `wb(df)`
> * Workbook name now defaults to `xleda` if no name is given.
> * Protected backwards compatibility while providing guidance to use the new API.
> 	* Subclassed the new API to create plugs for the old one.
> 
> **Simplified export functionality**
> * Changed `export_analysis` functionality from a class method to a class argument `wb(df, export=True)`.  
> * All wb() objects now include a export_dict metadata collection that is accessible using dot notation.
> * Added field_metadata, and overview_metadata to export_dict.
> * using wb(export=True) adds the metadata exported from workbooks to export_dict.
> * Added file exists checks for `export=True` with messaging that the export will be limited if the file isn't found. 
> 
> **Template updates**
> * Recreated the template, moved formatting to cell styles for simplicity/consistency in maintenance where appropriate. 
> *  Pivot was removed and Blanks was renamed Pivot.
> * `% of Records` field was added to the new Pivot
> * Added dataframe index to source data by default.
> * Added dataframe level metadata to the Data Description section.
> * Added two UDFs to the template, PythonList/PythonDict to create Python formatted strings from cell values
> * Adjusted the named range to support being able to delete almost any column without affecting lists or navigation.
> * General polish.
> 
> **Other updates**
> * Default limits were reduced to 25,000 rows/50 columns
> * Good deal of refactoring to support new entry point, minimize errors, reduce redundancy. 
> 	* Removed clipboard usage in all except one place which is formatting instead of data. 
> 	* Added `open_wb` argument to prevent automatically opening the workbook.  Useful if creating many workbooks.
> * Replaced rich progress bars with TQDM for better support in notebooks/vs code notebook/console environments.
> * When using overwrite=True, overwritten files now go to the recycle bin/trash.  Console output includes messaging about these files.
> * Clarified/organized readme to support the new API/template.
> * Added production logging metrics so you can see how the time required to create a workbook was utilized.  
> 	* This is useful if you're trying to find a good size to subsample to. 
> 	* You can find it at `wb().performance` for now.  
