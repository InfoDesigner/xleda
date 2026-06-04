<div align="center">

[![License: Apache](https://img.shields.io/badge/license-Apache-**blue**)](https://www.apache.org/licenses/LICENSE-2.0.txt)
[![PyPI - Version](https://img.shields.io/pypi/v/xleda)](https://pypi.org/project/xleda)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/xleda.svg)](https://pypi.org/project/xleda)
[![Downloads](https://static.pepy.tech/badge/xleda)](https://pepy.tech/project/xleda)
[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ffdd00?&logo=buy-me-a-coffee&logoColor=black)](https://buymeacoffee.com/informationdesigner)

</div>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/logo.webp?raw=true"  width="250" alt="Example Top View"    
	<br>
</p>
<p align="center" style="font-size: 26px; font-weight: bold;">A Microsoft Excel powered EDA tool for Python data.</p><br><br>

* Produces Microsoft Excel workbooks from pandas dataframes that are highly optimized to both perform and document [the activity of Exploratory Data Analysis](https://www.geeksforgeeks.org/data-analysis/what-is-exploratory-data-analysis/).<br><br>

* Visually explore your data, navigate with your keyboard, take field or record notes, create lists of fields/records for editing, round-trip your edits/analysis back into python, share your workbook with other contributors.<br><br>

* There are some amazing EDA tools for Python.  You shouldn't have to start from scratch to include Microsoft Excel among them.<br><br>

* xleda provides a good start to a robust EDA.<br><br>

* See [some example xleda workbooks](https://github.com/InfoDesigner/xleda/tree/main/examples).<br><br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/top_view.webp?raw=true"  width="600" alt="Example Top View"> 
	<br>
	<em>An xleda workbook made with diamond data.</em>
</p>
<br>

# **Requirements/Compatibility**

* Requires the full version of Microsoft Excel to create workbooks. 

* Once created, xleda workbooks should work in anything that reads Microsoft Excel workbooks.

* It has been developed and tested on Windows.  

* It should also work on MacOS though this has not yet been tested and is currently unsupported.<br><br>
# **Installation**

* Install with 

	```bash
	uv add xleda
	```
	or 
	```bash
	pip install xleda
	```

# **Quick Start**

* Use wb() to quickly create an xleda workbook of a dataframe. 

* See the configuration section below for how to name the workbook, set a theme, add additional dataframes/plots etc.  <br>

	```python
	from xleda import wb
	import seaborn as sns

	# < your dataframe goes here >
	df = sns.load_dataset("titanic")

	# Creates xleda.xlsm in the current directory
	wb(df)
	```
<br>

# **xleda Components**
### **Field Metadata**

* Most of the field metadata comes from the built-in pandas features [describe](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.describe.html), [info](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.info.html), and **[quantile](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.quantile.html)**<br><br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/basic_metadata.webp?raw=true"  width="600" alt="Basic Metadata"> 
	<br>
	<em>Included Field Metadata</em>
</p>


### **Overview**

* The `Overview` worksheet rotates the field metadata 90 degrees so that you can sort/filter fields by their name, metadata, or your notes/definitions/etc. <br><br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/overview.webp?raw=true"  width="600" alt="Composition Table"> 
	<br>
	<em>Sorting fields from MLB data by memory usage.</em>
</p>

### **Per Field Charts**

Two charts are produced for each column in your dataframe.

1. A composition table showing the top 5 values per column and their percentages.<br><br>
2. A histogram/KDE showing min/max, distribution, and mean<br><br>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/default_charts.webp?raw=true"  width="300" alt="Default Charts"> 
	<br>
	<em>Default charts for MLB player height.</em>
</p>

### **Source Data Table**

* A copy of your source data is included as an Excel table so you can visually inspect it, sort/filter it, etc. 

* Includes a way to make lists of individual records. 

* Includes a `HasBlank` field  for isolating incomplete records.

* Includes a way to round-trip your source data back into Python so that you can use Excel to replace values, delete records/columns/etc.  

* More on how to use these features below.


<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/source_data.webp?raw=true"  width="600" alt="Source Data"> 
	<br>
	<em>Source Data Table for MLB player data.</em>
</p>

### **Pivot**

* A bare-bones pivot table ready to be configured. <br><br>
* Defaults to include:<br>
  * The first 10 fields of the source data 
  * Measures to to identify blanks/dataset composition.<br><br>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/blanks.webp?raw=true"  width="600" alt="Blanks"> 
	<br>
	<em>Bare-bones pivot table for Titanic survivor data.</em>
</p>

### **Debug**

* A worksheet that includes details on configuration, environment, and how the time spent to produce an xleda workbook was allocated.<br><br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/debug.webp?raw=true"  width="600" alt="Blanks"> 
	<br>
	<em>Debug worksheet for troubleshooting.</em>
</p>
<br>

# **xleda.wb() Configuration**

### **input_df** | Dataframe | Mandatory

* A pandas dataframe of any size

### **name** | str | Optional

* Name of the dataset/file name of the created workbook.

* Punctuation will be removed to prevent issues with file name/workbook object names.

* Defaults to `xleda`<br>

### **theme_color** | str | Optional

* Sets the primary color of the charts and the color of the headings in the workbook to a hex color of your choice.  

* `theme_color="random"` sets a random theme<br>
* Defaults to a neutral color <br>
<br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/theme_colors.webp?raw=true" width="600" alt="Theme Colors">
	<br>
	<em>theme_color affects the workbooks and default charts.</em>
</p>
<br>

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



### **add_df** | dict | Optional

* `add_df={'dataset_name': dataset_df, ...}` will add Field Analysis/Overview worksheets for each additional dataframe provided into the same workbook.<br>
* Useful for grouping related data together.<br>

	```python
	import seaborn as sns
	import pandas as pd
	from xleda import wb
	
	  
	og_penguins = pd.read_csv("https://raw.githubusercontent.com/allisonhorst/palmerpenguins/refs/heads/main/inst/extdata/penguins_raw.csv")
	penguins = sns.load_dataset('penguins')
	seaice = sns.load_dataset('seaice')
	
	
	# Creates "OG Penguins.xlsm" in the current directory with worksheets included for each of the OG Penguins/Sea Ice/Seaborn Penguins dataframes.
	wb(input_df=og_penguins,
	   name="OG Penguins",
	   add_df={'Sea Ice': seaice,
	           'Seaborn Penguins': penguins})
	```

<br>

### **wb_path** | Path or string | Optional

* Uses a string or Pathlib path of a directory or file<br>
* Sets the target folder or workbook name of an xleda workbook.<br>
* If a directory is provided, the workbook will be created in that directory.<br>
* If a file name ending in ".xlsm" or ".xlsx" is provided:<br>
	* It will either create that file or export from that file depending on whether export=True is also selected. <br>
* Defaults to current working directory<br>

	```python
	from xleda import wb
	from pathlib import Path

	# Creates "c:\my_target_folder\Penguins.xlsm"
	wb(input_df=df,
	   name="Penguins",
	   wb_path=Path(r"c:\my_target_folder"))
	
	# Creates "c:\my_awesome_workbook.xlsx"
	wb(input_df=df,
	   name="Penguins",
	   wb_path=r"c:\my_awesome_workbook.xlsx")
	   
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

	```python
	from xleda import wb
	import seaborn as sns
	
	df = sns.load_dataset('penguins')
	
	# Creates "Penguins.xlsx" in the current directory
	wb(input_df=df,
	   name="Penguins",
	   no_vba=True)

	```


### **open_wb** | bool | Optional

* Opens the workbook after creating.

* Setting this to `False` is useful when creating multiple workbooks

* Defaults to True.<br>

### **export** | bool | Optional

* Performs an export from an xleda workbook instead of creating one. 

* Replaces the `export_analysis` method.

* See details below.

* Defaults to False.<br><br>

<br>

# **Usage Notes**

## **Performance**

* On an average machine, xleda creates workbooks for most data sets less than 20 seconds. 

* Performance is largely dependent on how powerful of a machine you have and how large/complex your dataframes are.  

* The `debug` worksheet will show you how the time spent to produce your workbook was allocated.

* There is a detailed output provided when creating an xleda workbook that does a pretty good job of letting you know what it's doing. 

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/create_example.webp?raw=true" width="400" alt="Create Example">
	<br>
	<em>Console output of a Planets workbook </em>
</p>
<br>

### **Limits with Large Data Sets**

* To ensure workbooks are created quickly, defaults limit data to the first 50 columns and a random sample of 25,000 records. 

* You can optionally override this to Excel's limits (see `large_report=True` above)

* You'll see a warning banner if you hit a limit.

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/warning.webp?raw=true" width="600" alt="Create Example">
</p><br>

* One of the more complex data sets tested was a 600 column/1,200 row dataframe.
  
	* It took ~5 minutes to create, in part because nearly all values are unique for all 600 columns.

	* It is still snappy to use even though it has 1,200 charts on a single worksheet.

	* That example is [here](https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/African%20Soil.xlsm).<br><br>

## **Field/Record Lists**


* The `Field Lists` section helps you create lists of the fields in your data.
	* e.g. lists of fields to rename, delete, standard scale, encode, impute, investigate, etc. 

* Anything not marked as `False` will be included in each list.   

* You can rename any list to `Anything You Want` and the list will be renamed to `anything_you_want`.

* The `Record List` field added to your source data works the same way except it creates a list of records instead of a list of fields.  More on that below.

* The Compiled Lists section formats your lists as python lists for easy copy/pasting.<br><br>

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

	* `index`:  This is a copy of the index from the provided dataframe as a column.<br>

## **Accessing Metadata in Python**


##### **Default Metadata**

* Metadata from  all `xleda.wb()` objects is collected into a list of dictionary objects, one for each dataframe, accessible through `xleda.wb().export_dicts`  <br>
	```python
	# Creates "Titanic.xlsm" and exports the metadata dictionaries
	export_dicts = wb(input_df=df,
	           	   name="Titanic").export_dicts

	# Returns the field metadata df from the primary dataframe
	export_dicts[0]['field_metadata'] 
	```


* The following metadata is available without using `export=True`

	* `field_metadata`: A basic metadata dataframe, combining information from pandas info/describe/quantile.

	* `overview_metadata`: A transposed copy of the field_metadata.

	* `source_data`: A copy of your unaltered source data that includes `Record Hash`/`Record List`/`HasBlank`/`index` columns.<br><br>

##### **Expanded Metadata**

* Using `xleda.wb(export=True)` reads an xleda workbook instead of creating one.  <br>
* It expands the available metadata within `xleda.wb().export_dicts` to include the following for each provided dataframe:

	* `description`: Dataframe description if you've added one

	* `definitions`: Any field definitions you've added.

	* `notes`: Any field notes you've added

	* `lists`: Any lists showing in the compiled lists section

	* `altered_source_data`: Reads the Source Data table named from the workbook and will include any manual edits you've made such as removing records, renaming fields, replacing values, etc. **

		 ** *Note that data types will likely change in the round-trip translation.* <br>
<br>


### **Completed Example**

* The xleda workbook pictured here is used in for the export code example below .  

* It can be found [here.](https://github.com/InfoDesigner/xleda/blob/main/examples/Titanic%20(Completed).xlsm).

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

* This example exports everything from an xleda workbook named  "Titanic Completed.xlsm" in the current directory.<br><br>

* Either download [this one](https://github.com/InfoDesigner/xleda/blob/main/examples/Titanic%20Completed.xlsm) or create your own.<br><br>

	```python
	from xleda import wb

	# Performs a full export from "Titanic Completed.xlsm"
	export_dicts = wb(input_df=df,
	                  path="Titanic Completed.xlsm",
	                  export=True).export_dicts

	
	# Returns dict_keys(['description', 'definitions', 'notes', 
	# 'lists', 'field_metadata', 'overview_metadata', 'source_data', 
	# 'altered_source_data'])
	print(export_dicts[0].keys()) 
	```
<br>

## **VBA Code**

* If you can't or don't want to enable VBA, you may want to use `no_vba=True` which creates an  xlsx file that contains no VBA.<br>
* The code that is there does two optional things. <br>

	1. Makes the sections expand/collapse when you select them as pictured above.<br>
		* You can use row groupings to navigate without VBA as pictured below.<br>
	2. Adds two UDFs, PythonList and PythonDict, that format cell values as lists/dicts. <br><br>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/row_groupings.webp?raw=true" width="462" alt="Row Groupings">
	<br>
	<em>Use row groupings to navigate if you can't use VBA.</em>
</p>
<br><br>

## **Extensibility**

* xleda is only meant to give a good start to EDA.<br><br>

* If it accomplishes one thing it will be to give you a way to quickly get data into Excel so that you can see and make sense of it...without making you do everything from scratch.  <br><br>

* Where you go from there is up to you.<br><br>

* Because it's an ordinary workbook, you can use any tool that works with Microsoft Excel workbooks to do more.  <br><br>

* [xlwings](https://www.xlwings.org/), is recommended if you do. <br><br>


## Troubleshooting

* if xleda is slow:
	* Try reducing the amount of data you're sending to it, and let it finish.
	* After production, refer to the debug worksheet for how the time to produce your xleda workbook is being spent so that you can reduce your dataframe size more strategically.<br><br>

* If you receive the "Error: The workbook cannot be overwritten while open!" and don't see any open workbooks:
	* You may have a hidden Excel instance that needs to be closed. 
	* Guidance on closing hidden Excel windows for [MacOS](https://www.google.com/search?q=hidden+excel+instance+in+macos)/[Windows](https://www.google.com/search?q=hidden+excel+instance+in+windows)<br><br>
* If you can't get xleda to run at all and are using Windows with a full Office Installation:
	* Try getting the following script to run using xlwings (not xlwings-lite).
	* All it does is open Excel and create a new workbook.
	* You should be able to pip install xlwings and run the script successfully. 
	* If that doesn't work, see their [installation instructions](https://docs.xlwings.org/en/latest/installation.html) for details on how to get it set up.
	* Be aware that xlwings has a ton of functionality and that for xleda to work, it only requires communication with Excel and not the addin, xlwings lite, udfs, or many of the other things xlwings can potentially do.
	* If you can get the script to run successfully, xleda has a good chance of working reliably.<br><br>

	```python
	import xlwings as xw
	
	app = xw.App()
	
	```


## Built With

* This was primarily built with [Python](https://www.python.org/), [xlwings](https://www.xlwings.org), [Pandas](https://pandas.pydata.org/), and of course, [Microsoft Excel](https://developer.microsoft.com/en-us/excel)<br>

## **Roadmap**

* [x] Add a barebones pivot that is ready to configure
* [x] Make xleda even more accessible by simplifying the API and making it easier to remember.
* [x] Create a way to quickly view dataframe data that is editable, shareable, and presentation ready.
* [x] Add a way to include extra plots for a dataset.
* [x] Add a way to include extra supporting dataframe worksheets.
* [ ] Add a way to use on desktop files e.g. by right-clicking csv/parquet files/other tabular data files. 
* [x] Add a way to include multiple xleda analyses in a single workbook.
* [ ] Test on MacOS
* [ ] Your idea here.






## Changelog

<details> 
<summary>Version 0.8.185 - New simplified API, simplified export, general polish
</summary>

<br>

&emsp;**Simplified basic usage to make it quicker to use and easier to memorize.** 
* Changed the default entry point to `xleda.wb()` from `xleda.FieldAnalysis()`.  
* `xleda.wb()` now creates and automatically opens workbooks.
* The only argument needed to create a workbook is now a dataframe. `wb(df)`
* Workbook name now defaults to `xleda` if no name is given.
* Protected backwards compatibility while providing guidance to use the new API.
	* Subclassed the new API to create plugs for the old one.<br><br>

&emsp;**Simplified export functionality**
* Changed `export_analysis` functionality from a class method to a class argument `wb(df, export=True)`.  
* All wb() objects now include a export_dict metadata collection that is accessible using dot notation.
* Added field_metadata, and overview_metadata to export_dict.
* using wb(export=True) reads a workbook instead of creating one and adds the metadata from the workbooks to export_dict.
* Added file exists checks for `export=True` with messaging that the export will be limited if the file isn't found. <br><br>

&emsp;**Template updates**
* Recreated the template, moved formatting to cell styles for simplicity/consistency in maintenance where appropriate. 
*  Pivot was removed and Blanks was renamed Pivot.
* `% of Records` field was added to the new Pivot
* Added dataframe index to source data by default.
* Added dataframe level metadata to the Data Description section.
* Added two UDFs to the template, PythonList/PythonDict to create Python formatted strings from cell values
* Adjusted the named range to support being able to delete almost any column without affecting lists or navigation.
* General polish.<br><br>

&emsp;**Other updates**
* Default limits were reduced to 25,000 rows/50 columns
* Good deal of refactoring to support new entry point, minimize errors, reduce redundancy. 
	* Removed clipboard usage in all except one place which is formatting instead of data. 
	* Added `open_wb` argument to prevent automatically opening the workbook.  Useful if creating many workbooks.
* Replaced rich progress bars with TQDM for better support in notebooks/vs code notebook/console environments.
* When using overwrite=True, overwritten files now go to the recycle bin/trash.  Console output includes messaging about these files.
* Clarified/organized readme to support the new API/template.
* Added production logging metrics so you can see how the time required to create a workbook was utilized.  
	* This is useful if you're trying to find a good size to subsample to. 
	* You can find it at `wb().performance` for now.<br>

<br>
</details>


<details> 
<summary>Version 0.8.186 - Create xleda workbooks of multiple dataframes, module refactoring into classes, logging, general polish
</summary> 

<br>

&emsp;**Implemented  `add_df`**

* Adds Field Analysis/Overview reports for each additional dataframe.
* Pivot is only provided for primary dataframe.
* Useful for supporting or related data.
* Worksheet names now include the dataframe name.
* Each dataframe's worksheet set gets a greyscale gradient so they can be visually distinguished among worksheet tabs.<br><br>

&emsp;**Export adjustments**
* Implemented an ExportDict class to add structure to export functionalities
* To support the additional dataframes from `add_df` functionality, `export_dict` has been renamed to `export_dicts` and now provides a list of ExportDict objects, one for each provided dataframe.
* ExportDict allows access to metadata through both dot notation and `dict[key]`.
* Reinforced handling of modified export workbook. 
	* If a workbook is found but that the expected worksheets aren't found, i.e. they've been deleted or renamed, it will export what it can and return a list of what wasn't found.<br><br>

&emsp;**Reinforced `wb_path`/`name` handling**
* `wb_path` now accepts strings, or pathlib Path objects.
* Also accepts full/partial paths with/without correct extensions
* Providing a path ending in `xlsx` or `.xlsm` will set `no_vba` to `True`/`False` respectively
* Illegal characters are now properly stripped from provided names before using in file/object names<br><br>


&emsp;**Added production logging/debug worksheet**
* The `debug` worksheet details how the time it took to produce the workbook was allocated on both field and workbook levels.
* Also includes configuration and system details<br><br>

&emsp;**Other Updates**
* Tests, examples, readme updated to reflect new functionality
* In the template, the `Field Notes` section of the `Field Analysis`worksheet was merged into `Data Description` section.
* Refactored the primary module into more specialized classes.
	* Configuration/environment/plotting/logging/theme all have their own classes
	* Also implemented new Blueprint class
		* Workbooks are now constructed from config object that includes a list of  Blueprints
		* Each provided dataframe gets it's own Blueprint
* Improved handling of datatypes that are unsupported in Excel/xlwings such as TimeDelta
* Reinforced system configuration checks with more informative offramps for:
	* Unsupported system configurations
	* Situations where necessary template components have been removed/renamed.
* Adjusted Github Action script to remove all but last changelog and convert the details/summary to standard markdown.

<br>
</details>