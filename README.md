<div align="center">


<a href="https://www.apache.org/licenses/LICENSE-2.0.txt"><img src="https://img.shields.io/badge/license-Apache-**blue**"></a> <a href="https://pypi.org/project/xleda"><img src="https://img.shields.io/pypi/v/xleda"></a> <a href="https://pypi.org/project/xleda"><img src="https://img.shields.io/pypi/pyversions/xleda.svg"></a> <a href="https://pepy.tech/project/xleda"><img src="https://static.pepy.tech/badge/xleda"></a> <a href="https://github.com/InfoDesigner/xleda"><img src="https://img.shields.io/badge/Made%20By%20A%20Human-99%25-blue)"></a> <a href="https://buymeacoffee.com/informationdesigner"><img src="https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ffdd00?&logo=buy-me-a-coffee&logoColor=black"></a>


</div>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/logo.webp?raw=true" width="250" alt="Logo">
	<br>
</p>
<p align="center" style="font-size: 26px; font-weight: bold;">A Microsoft Excel powered EDA tool</p>

<br>

* Produces Microsoft Excel workbooks from dataframes or data files that are highly optimized to explore, define, and document data sets. <br><br>

* Works on both MacOS/Windows as a Python package, a CLI, or as a service that lets you create workbooks by right-clicking supported files.<br><br>

* There are some amazing EDA tools available to data professionals. You shouldn't have to start from scratch to include Microsoft Excel among them.<br><br>

* See [some example xleda workbooks](https://github.com/InfoDesigner/xleda/tree/main/examples).<br><br>
<p align="center">
	<img src="assets/images/top_view.webp"  width="600" alt="Example Top View"> 
	<br>
	<em>Top view of a Field Analysis worksheet.</em>
</p><br>

## **Requirements/Compatibility**
<br>
<table>
  <tbody>
    <tr>
      <td width="30%" valign="top"><strong>Desktop Excel</strong></td>
      <td valign="top">
        <ul>
          <li>Requires the full version of Microsoft Excel (2016+) on either MacOS or Windows to create workbooks.</li>
          <li>See MacOS Support section below for details on MacOS usage.</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td valign="top"><strong>Supported Data</strong></td>
      <td valign="top">
        <ul>
          <li>Supports pandas dataframes, CSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML.
          </li>
        </ul>
      </td>
    </tr>
  </tbody>
</table><br>

## **Installation**


| **Package/CLI**      | `uv add xleda` or `pip install xleda`                                                                                                                                                                                                                                                                                 |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Right-Click Menu** | - xleda can optionally be installed into the OS such that it will create workbooks from a right-click context menu action on supported file types.<br>  <br>- Works on both Windows and MacOS<br>  <br>- After installing as a package, use `xleda install` or `xleda uninstall` to add/remove right-click menus.<br> |


<details markdown="1">
  <summary><strong>Tips:</strong> Managing the Install</summary>
  
- `xleda install` adds right-click functionality to your OS but it does not modify your path to make the CLI globally available.
  
- If the Python environment that xleda was installed into is deleted after running `xleda install`, the right-click functionality will need to be either repaired or uninstalled by running `xleda install/xleda uninstall` again from a new Python environment.
  
- If you have UV installed, you can install/uninstall the package, CLI, and right-click menus systemwide without having to maintain a venv with these one-liners.

**Windows PowerShell**

```bash
# Installs the package and right-click menus
uv tool install xleda; if ($?) { xleda install }

# Uninstalls the package and right-click menus
xleda uninstall; if ($?) { uv tool uninstall xleda }
```

**MacOS**

```bash
# Installs the package and right-click menus
uv tool install xleda &amp;&amp; xleda install

# Uninstalls the package and right-click menus
xleda uninstall &amp;&amp; uv tool uninstall xleda
```

</details><br>

## **Quick Start**

Use <code>wb()</code> to quickly create an xleda workbook from a dataframe or a supported data file.

### From a Dataframe

```python
from xleda import wb
import seaborn as sns

# < your dataframe goes here >
df = sns.load_dataset("titanic")

# Creates xleda.xlsm in the current directory
wb(df)
```

### From a File

```python
from xleda import wb
from pathlib import Path

# < your data file goes here >
duckdb_file = Path("my_fancy_db.duckdb")

# Creates my_fancy_db.xlsm in the current directory
# Includes data from all tables in the db file
wb(duckdb_file)
```

### From the CLI

```bash
# Creates 'my_parquet_file.xlsm' in the current directory
xleda wb 'my_parquet_file.parquet'
```

### From Right-Clicking

<p align="left">
  <img src="assets/images/from_right_click.webp" width="600" alt="From right-click">
</p><br>

## **xleda Components**

All xleda workbooks include an **Overview** worksheet and a **Field Analysis** worksheet for each provided dataframe.

| **Component**      | **Description**                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| ------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Field Analysis** | Each **Field Analysis** worksheet Includes:    <br><br>- Placeholders for inputting definitions/notes<br>- Metadata for each field and dataframe<br>- Per Field Charts<br>- A way to create lists of fields/records<br>- A Source Data Table.<br><br><details><br>  <summary><strong>Anatomy of a Field Analysis Worksheet:</strong></summary><br><br>  <p align="center"><br>	<img src="assets/images/field_analysis_anatomy.webp" width="600" alt="Field Analysis Anatomy"> <br>	<br><br>	<em>Field Analysis Anatomy</em><br>  </p><br></details><br>                                                                                                                                                                           |
| **Overview**       | The **Overview** worksheet includes:  <br> <br>- A dataframe-level metadata table with links to each Field Analysis worksheet and a calculated field for tracking how many fields have definitions.<br><br>- A field-level metadata table that includes all fields from all provided dataframes in one table so that you can sort/filter fields by their name, memory usage, missing %, your notes/definitions/etc.<br><br><details><br>  <summary><strong>Anatomy of an Overview Worksheet:</strong></summary><br><br>  <p align="center"><br>	<img src="assets/images/overview_anatomy.webp" width="800" alt="Overview Anatomy"> <br>	<br><br>	<em>Overview Worksheet with Multiple Dataframes</em><br>  </p><br></details><br> |

<br><br>


## **xleda.wb() Configuration**



| `data`         | **Dataframe or Path or string \| Mandatory**<br><br>- Accepts a pandas dataframe, a dictionary of dataframes, or a supported data file.<br>- For files, xleda will create a workbook from all tabular objects in the file.<br>- Supported types include CSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML.<br><br><details><br>  <summary>Data File Limitations</summary><br>If the provided data file doesn't parse correctly, try creating a dataframe first and use that with xleda instead of the file.  <br>  <br>**Expect problems with**:  <br>- Deeply nested `JSON/XML`<br>- A `.CSV` file with tabs instead of commas<br>- `.db` files that are neither SQLite nor DuckDB files<br>- DuckDB files with a `.txt` extension<br>    <br>  <br>**Don't expect to see**:  <br>- R Functions from an RData file<br>- matplotlib Figures from a pickle file<br>- Anything sourced from an Excel file that isn't a proper Excel Table<br><br></details><br> |
| -------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `file_name`    | **str \| Optional**<br><br>- The workbook file name to create.<br>- Defaults to the source file name, the first key in a dataframe dict, or `xleda`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| `wb_path`      | **Path or string \| Optional**<br><br>- Use a directory or file path.<br>- If a directory is provided, the workbook is created there.<br>- If a filename ends with `.xlsm` or `.xlsx`, xleda will create or export that file.<br>- Defaults to the current working directory or the source file directory.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| `theme_color`  | **str \| Optional**<br><br>- Sets the primary workbook and chart color.<br>- Supports hex colors or `"random"`.<br>- Defaults to a neutral theme.<br><p align="center"><br>          <img src="assets/images/theme_colors.webp" width="800" alt="Theme Colors"><br>          <em>theme_color affects the workbooks and default charts.</em><br>        </p>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| `plots`        | **dict \| Optional**<br><br>- Adds extra worksheets using a dict of matplotlib Figure objects.<br>- Accepts format `{'plotname': Figure, ...}`.<br>- No automatic styling or sizing is applied.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| `overwrite`    | **bool \| Optional**<br><br>- Overwrites existing workbooks of the same name.<br>- Existing files are moved to Trash/Recycle Bin.<br>- Defaults to `False`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| `large_report` | **bool \| Optional**<br><br>- Raises limits to Excel's maximum: 1,000,000 rows and 16,000 columns.<br>- Requires more memory and time for large datasets.<br>- Defaults to `False`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| `no_vba`       | **bool \| Optional**<br><br>- Creates a `.xlsx` workbook without VBA.<br>- Setting this flag persists the preference.<br>- Use a `.xlsx` `wb_path` as an alternative.<br>- Defaults to `False`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| `open_wb`      | **bool \| Optional**<br><br>- Opens the workbook after creation.<br>- Set to `False` when creating multiple workbooks.<br>- Defaults to `True`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| `export`       | **bool \| Optional**<br><br>- Exports data from an xleda workbook instead of creating one.<br>- See the _Exporting Metadata_ section below for examples.<br>- Defaults to `False`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |


	
## **Examples**

<details>
  <summary>Example: Using a dictionary of dataframes</summary><br>

```python
import seaborn as sns
from xleda import wb

seaborn_datasets = ['diamonds', 'dots', 'dowjones']
dataframe_dict = {df_name: sns.load_dataset(df_name) for df_name in seaborn_datasets}

# Creates Diamonds.xlsm in the current directory
# Also includes dots and dow jones data
wb(data=dataframe_dict)
```

</details><br>

<details>
  <summary>Example: Using wb_path</summary><br>

```python
from xleda import wb
from pathlib import Path

# Creates "c:\my_target_folder\Penguins.xlsm"
wb(data={"Penguins": df},
   wb_path=Path(r"c:\my_target_folder"))

# Creates "c:\my_awesome_workbook.xlsx"
wb(data={"Penguins": df},
   wb_path=r"c:\my_awesome_workbook.xlsx")
    
```

</details><br>
<details>
  <summary>Example: Including additional plots</summary><br>

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

# Creates Penguins.xlsm with two extra plot sheets
wb(data={"Penguins": df},
   theme_color="#4C4C4C",
   plots={'Pair Plots': pair_plots,
          'Null Matrix': null_matrix})
```
</details><br>


<details>
  <summary>Example: Creating workbooks without VBA</summary><br>

```python
from xleda import wb
import seaborn as sns

df = sns.load_dataset('penguins')

# Creates "Penguins.xlsx" in the current directory and changes the default workbook style to .xlsx
wb(data={"Penguins": df},
   no_vba=True)

# Also creates "Penguins.xlsx" but doesn't change the default workbook style 
wb(data=df,
   wb_path="Penguins.xlsx")
```

</details><br>


## **Usage Notes**


### Field and Record Lists

* The `Field Lists` section includes placeholders to create 8 custom lists of fields and one somewhat more robust list of records. 

* Use these to organize fields into groups such as "fields_to_delete", "fields_from_system_a", "fields_to_fix", or whatever your workflow needs.<br><br>

<details markdown="1">
<summary><strong>Field/Record List Details</strong></summary><br>

* Anything not marked as `False` will be included in each list.   

* You can rename any list to `Anything You Want` and the list will be renamed to `anything_you_want`.

* The `Record List` field added to your source data works the same way except it creates a list of all tagged records instead of a list of fields.

* The Compiled Lists section formats your lists as python lists.


<p align="center">
	<img src="assets/images/field_lists.webp" width="800" alt="Field Lists">
	<br>
	<em>Easily create lists of fields in your data.</em>
</p>
<br>
</details><br>

### Large Data Sets

On an average machine, xleda creates workbooks for most data sets less than 20 seconds on Windows/1-2 minutes on MacOS

* To ensure workbooks are created quickly, each dataframe is by default subsampled to only include the first 50 columns and a random sample of 25,000 records.<br><br>

<details markdown="1">
<summary><strong>More Details</strong></summary>


* Performance is largely dependent on how powerful of a machine you have and how many/how large/how complex your dataframes are.  

* You can optionally override default limits to use Excel's limits of 16,000 columns, 1,000,000 rows by using `large_report=True`.

* You'll see a warning banner on Field Analysis worksheets of affected dataframes if they've exceeded a limit.

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/warning.webp?raw=true" width="900" alt="Warning Banner">
</p>

</details>

### Source Data Table

There are some columns added to the **Source Data** table of each dataframe's **Field Analysis** worksheet.

* Use them to create lists of individual records and to isolate records with missing values.  

* Use `Record List` column to tag records for `deletion`, `curing defects`, `investigation`, or whatever your workflow demands<br><br>

<details markdown="1">
<summary>More Details</summary>

<p align="center">
	<img src="assets/images/souce_data_table.webp" width="900" alt="Warning Banner">
</p>


*  They are colored grey on grey to easily distinguish them from your source data.

* `HasBlank`: If any field in a record has a missing value, this will show 1 otherwise 0

* `Record Hash`:  Uses a built-in pandas feature [hash_pandas_object](https://pandas.pydata.org/docs/reference/api/pandas.util.hash_pandas_object.html) to uniquely identify records.  If two records share all column values they also share a `Record Hash`. 

* `Record List`:  Used to tag individual records.  Like `Field Lists` above, anything not marked false gets added to 'Record List'.

</details><br>

### Exporting Metadata

* Metadata from  all `xleda.wb()` objects is collected into a list of dictionary objects, one for each dataframe, accessible through `xleda.wb().export_dicts`. 

* You can access expanded metadata, sourced from the workbook by using `export=True`<br><br>

**Default Metadata**<br>

The default metadata is the same field and dataframe metadata that is added to the workbooks.


<details markdown="1">
<summary>Metadata Details</summary>

The following metadata is available without using `export=True`

* `field_metadata`: A basic metadata dataframe, combining information from pandas info/describe/quantile.

* `field_overview`:  Field-level metadata as seen in the field section of the **Overview** worksheet.

* `df_overview`: Dataframe level metadata as seen in the dataframe section of the **Overview** worksheet.

* `source_data`: A copy of the source data that also includes  `Record Hash`/`Record List`/`HasBlank`/`index` columns.<br><br>
</details><br>

<details markdown="1">
<summary>Example: Accessing basic metadata
</summary>

```python
# Creates "Titanic.xlsm" and exports default metadata
export_dicts = wb(data={"Titanic", df}).export_dicts

# Returns ['field_metadata', 'field_overview', 'df_overview', 'source_data']
export_dicts[0].keys()
```

</details><br><br>

**Expanded Metadata**

Using `export=True` reads all metadata from the workbook instead of creating it and includes your notes, lists, definitions, etc.

* It will reflect any changes you've made in Excel such as renaming fields/deleting values/etc.

* There are a lot of ways to read Excel data in Python, this just makes it easy.<br><br>

<details markdown="1">
<summary>Expanded Metadata:</summary>


* Also includes the following for each provided dataframe:

	* `description`: Dataframe description if you've added one

	* `definitions`: Any field definitions you've added.

	* `notes`: Any field notes you've added

	* `lists`: Any lists showing in the compiled lists section

	 Note that data types will likely change in the round-trip translation. <br>
   
</details><br>

<details markdown="1">
<summary>Example: Accessing Expanded Metadata From a Completed Workbook</summary>

* The xleda workbook pictured here is used in for the export code example below .  

* It can be found [here.](https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/Titanic%20Completed.xlsm).

<p align="center">
	<img src="assets/images/completed_field_analysis.webp" width="800" alt="Completed Field Analysis">
	<br>
	<em>A completed xleda workbook showing definitions, notes, lists, etc.</em>
</p>
<br>


```python
from xleda import wb

# Performs a full export from "Titanic Completed.xlsm"
export_dicts = wb(data={"Titanic Completed", df},
				  export=True).export_dicts


# Returns ['description', 'definitions', 'notes', 'lists', 'field_metadata', 'field_overview', 'df_overview', 'source_data']
print(export_dicts[0].keys())
```

<br>
</details><br>

### MacOS Support

xleda will create the same workbooks in MacOS.

* Creating them is significantly slower and you may get two different types of prompts that require your attention.  

* Look for the bouncing Excel icon.

<details markdown="1">
<summary><strong>MacOS Details</strong></summary>



<table>
  <thead>
    <tr>
      <th width="120"></th>
      <th width="330">To Access Files</th>
      <th width="330">To Enable Macros</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>Source</strong></td>
      <td>MacOS</td>
      <td>Excel</td>
    </tr>
    <tr>
      <td><strong>Details</strong></td>
      <td>Prompts to Allow Excel to access the file it's creating.<br><br>If you get these prompts, you'll potentially get one for each unique file you create.</td>
      <td>Prompts to "Enable Macros".  <br><br>If you get these prompts, you'll get two when creating a workbook:<br><br>1. When opening the blank template<br>2. When opening your created workbook.</td>
    </tr>
    <tr>
      <td><strong>Example</strong></td>
      <td><img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/grant_file_access.webp?raw=true" width="350" alt="Grant file access prompt"></td>
      <td><img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/enable_macros.webp?raw=true" width="350" alt="Enable macros prompt"></td>
    </tr>
    <tr>
      <td><strong>Remedy</strong></td>
      <td>There's not a reliable remedy to this.<br><br>MacOS doesn't permit applications like Microsoft Excel real access to the file system, even after explicitly granting Excel Full Disk Access under <code>Settings &gt; Privacy & Security &gt; Full Disk Access</code>.</td>
      <td>You can either:<br><br>1. Create a VBA free workbook (see the next section for details).<br><br>2. Change Excel's default macro settings (shown) to one of the other two options.<br><br><img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/excel_macos_macro_options.webp?raw=true" width="350" alt="Excel MacOS macro settings"></td>
    </tr>
  </tbody>
</table>



</details><br>

### VBA Code

* The included [VBA code](https://github.com/InfoDesigner/xleda/blob/main/src/xleda/vba.bas) is fairly short and easy to understand.  

* You can create a VBA-free, xlsx workbook by either setting `no_vba=True` or providing a `wb_path` ending in `.xlsx`. 

* Providing the `no_vba` flag will change the default so that the setting will persist.  Set it once and forget it.  Using `wb_path` doesn't work this way.



<details markdown="1">
<summary><strong>VBA Code Details</strong></summary><br>

**What it does**

1. Makes the sections expand/collapse when you select them as pictured on the left which can also be performed by using row groupings as pictured on the right.<br>
2. Adds a PythonList UDF that creates Python lists from cell values. <br><br>

**The Annoyance Cost**

Besides Enable Macros prompts, using VBA also includes one annoying side effect:

* Every time a macro is used, it clears your undo history.
* The PythonList UDF is immune to this but expanding/collapsing headings is not.

<table>
  <tr>
    <td align="center">
      <img src="assets/images/top_view.webp" width="462" alt="Row Groupings">
      <br>
      <em>Use headings like web pages to navigate with VBA.</em>
    </td>
    <td align="center">
      <img src="assets/images/row_groupings.webp" width="462" alt="Row Groupings">
      <br>
      <em>Use row groupings to navigate without VBA.</em>
    </td>
  </tr>
</table>
<br>
</details><br>

## Troubleshooting

<details markdown="1"> 
<summary>xleda is slow
</summary><br>

* Try reducing the amount of data you're sending to it, and let it finish.
* After production, refer to the `debug` section of the `Overview` worksheet for how the time to produce your workbook is being spent.
* Note that on MacOS, `xleda` is much slower by default and the timings in the debug section may be inflated from missed permission prompts during production.

</details><br>

<details markdown="1"> 
<summary>If you receive the "Error: The workbook cannot be overwritten while open!" and don't see any open workbooks:
</summary><br>

* You may have a hidden Excel instance that needs to be closed. 
* Guidance on closing hidden Excel windows for [MacOS](https://www.google.com/search?q=hidden+excel+instance+in+macos)/[Windows](https://www.google.com/search?q=hidden+excel+instance+in+windows)

</details><br>

<details markdown="1"> 
<summary>If you receive the "Exception: Could not activate App!" or "The RPC server is unavailable". errors:
</summary><br>

* The Excel app may have crashed or is otherwise disconnected from Python.
* Close all Excel windows and try running the command again.

</details><br>


<details markdown="1"> 
<summary>If you can't get xleda to run at all and are using Windows/MacOS with a full Office Installation:
</summary><br>

* Try getting the following script to run using xlwings (not xlwings-lite).
* All it does is open Excel and create a new workbook.
* You should be able to `pip install xlwings` and run the script successfully. 
* If that doesn't work, see their [installation instructions](https://docs.xlwings.org/en/latest/installation.html) for details on how to get it set up.
* Be aware that xlwings has a ton of functionality and that for xleda to work, it only requires communication with Excel and not the addin, xlwings lite, udfs, or many of the other things xlwings can potentially do.
* If you can get the script below to run successfully, xleda has a good chance of working reliably.
* If you can't get it to work and you're on Windows, [this may help](https://www.google.com/search?q=win32+com+corruption+xlwings).

 <br><br>

```python
import xlwings as xw

app = xw.App()

```

</details><br>


## Changelog



| **Version 0.8.185** | **New simplified API, simplified export, general polish**<br><br><br><details markdown="1"> <br><summary>Details:</summary><br>**Simplified basic usage to make it quicker to use and easier to memorize.**  <br>  <br>- Changed the default entry point to `xleda.wb()` from `xleda.FieldAnalysis()`.  <br>- `xleda.wb()` now creates and automatically opens workbooks.  <br>- The only argument needed to create a workbook is now a dataframe: `wb(df)`.  <br>- Workbook name now defaults to `xleda` if no name is given.  <br>- Protected backwards compatibility while providing guidance to use the new API.  <br>- Subclassed the new API to create plugs for the old one.  <br>  <br>**Simplified export functionality**  <br>  <br>- Changed `export_analysis` functionality from a class method to a class argument `wb(df, export=True)`.  <br>- All `wb()` objects now include a `export_dict` metadata collection that is accessible using dot notation.  <br>- Added `field_metadata` and `overview_metadata` to `export_dict`.  <br>- Using `wb(export=True)` reads a workbook instead of creating one and adds the metadata from the workbooks to `export_dict`.  <br>- Added file exists checks for `export=True` with messaging that the export will be limited if the file isn't found.  <br>  <br>**Template updates**  <br>  <br>- Recreated the template, moved formatting to cell styles for simplicity/consistency in maintenance where appropriate.  <br>- Pivot was removed and Blanks was renamed Pivot.  <br>- `% of Records` field was added to the new Pivot.  <br>- Added dataframe index to source data by default.  <br>- Added dataframe level metadata to the Data Description section.  <br>- Added two UDFs to the template, PythonList/PythonDict, to create Python formatted strings from cell values.  <br>- Adjusted the named range to support deleting almost any column without affecting lists or navigation.  <br>- General polish.  <br>  <br>**Other updates**  <br>  <br>- Default limits were reduced to 25,000 rows/50 columns.  <br>- Good deal of refactoring to support the new entry point, minimize errors, reduce redundancy. <br>- Removed clipboard usage in all except one place, where it is used for formatting instead of data. <br>* Added `open_wb` argument to prevent automatically opening the workbook, useful when creating many workbooks.<br>* Replaced rich progress bars with TQDM for better support in notebooks/vs code notebook/console environments.  <br>- When using `overwrite=True`, overwritten files now go to the recycle bin/trash. Console output includes messaging about these files.  <br>- Clarified/organized readme to support the new API/template.  <br>- Added production logging metrics so you can see how the time required to create a workbook was utilized.<br>* This is useful if you're trying to find a good size to subsample to.  <br>- You can find it at `wb().performance` for now.<br><br></details><br> |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Version 0.8.186** | **Add multiple dataframes, module refactoring into classes, added logging**<br><br><br><details markdown="1"> <br><summary>Details:</summary><br>**Implemented `add_dfs`**  <br>  <br>- Adds Field Analysis/Overview reports for each additional dataframe.  <br>- Pivot is only provided for the primary dataframe.  <br>- Useful for supporting or related data.  <br>- Worksheet names now include the dataframe name.  <br>- Each dataframe's worksheet set gets a greyscale gradient so they can be visually distinguished among worksheet tabs.  <br>  <br>**Export adjustments**  <br>  <br>- Implemented an ExportDict class to add structure to export functionalities.  <br>- To support the additional dataframes from `add_dfs` functionality, `export_dict` has been renamed to `export_dicts` and now provides a list of ExportDict objects, one for each provided dataframe.  <br>- ExportDict allows access to metadata through both dot notation and `dict[key]`.  <br>- Reinforced handling of modified export workbooks.  <br>- If a workbook is found but the expected worksheets aren't found (for example, if they've been deleted or renamed), it will export what it can and return a list of what wasn't found.  <br>  <br>**Reinforced `wb_path`/`name` handling**  <br>  <br>- `wb_path` now accepts strings or pathlib Path objects.  <br>- Also accepts full/partial paths with/without correct extensions.  <br>- Providing a path ending in `.xlsx` or `.xlsm` will set `no_vba` to `True`/`False` respectively.  <br>- Illegal characters are now properly stripped from provided names before use.  <br>  <br>**Added production logging/debug worksheet**  <br>  <br>- The `debug` worksheet details how the time it took to produce the workbook was allocated on both field and workbook levels.  <br>- Also includes configuration and system details.  <br>  <br>**Other Updates**  <br>  <br>- Tests, examples, readme updated to reflect new functionality.  <br>- In the template, the `Field Notes` section of the `Field Analysis` worksheet was merged into the `Data Description` section.  <br>- Refactored the primary module into more specialized classes.  <br>- Configuration/environment/plotting/logging/theme all have their own classes.  <br>- Also implemented a new Blueprint class.  <br>- Workbooks are now constructed from a config object that includes a list of Blueprints.  <br>- Each provided dataframe gets its own Blueprint.  <br>- Improved handling of datatypes that are unsupported in Excel/xlwings such as TimeDelta.  <br>- Reinforced system configuration checks with more informative offramps for:  <br>- Unsupported system configurations.  <br>- Situations where necessary template components have been removed or renamed.  <br>- Adjusted Github Action script to remove all but last changelog and convert the details/summary to standard markdown.<br><br></details><br>                                                       |
| **Version 0.8.193** | **Added MacOS support**<br><br><br><details markdown="1"> <br><summary>Details:</summary><br><br>**Added MacOS support**  <br>  <br>- Used xlwings when possible, appscript/AppleScript/subprocess otherwise.  <br>- Reduced OS branching when possible.  <br>- Documentation/test/tools/examples updated to be cross platform.  <br>  <br>**Other Updates**  <br>  <br>- Removed field logging from logging/template.  <br>- Simplified some of the pivot configuration where possible.  <br>- Added multi-threading for the progress bar which keeps the time elapsed ticking during longer iterations.  <br>  <br>**Template Adjustments**  <br>  <br>- Moved the xlsx conversion to a pre-commit hook instead of an on-demand end-user task.  <br>- Adjusted expand/collapse icons to use a more reliable cross-platform character.  <br>- Removed navigation shapes from the xlsx template.<br><br></details><br>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| **Version 0.8.197** | **Readme/pyproject.toml polish/minor fixes**<br><br><details markdown="1"> <br><summary>Details:</summary><br>- Moved code examples/troubleshooting/usage notes into details/summary blocks to reduce clutter in README.  <br>- Fixed a cross-platform formatting issue with the debug worksheet.  <br>- Updated a few older screenshots to use the current template.  <br>- Adjusted `type: ignore` lines where possible.  <br>- Organized pyproject.toml, added `required-environments` section.<br><br></details><br>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| **Version 0.9.001** | **Expanded Input Options and Ways to Use, Simplified API, Improved Experience with Multiple Dataframes, Added Persistent Settings.**<br><br><details markdown="1"> <br><summary>Details:</summary><br><br>**Expanded Input Options**  <br>  <br>- Added a way to use dataframe dictionaries and data files as sources.  <br>- A file that represents a tabular object, such as parquet/csv files, will create a dataframe and then create a workbook from that dataframe.  <br>- Complex files, such as duckdb, RData, or sqlite files, will create dataframes from all tabular objects within each file and create a workbook with those dataframes.  <br>  <br>**Expanded Ways to Use**  <br>  <br>- Added a **CLI** interface which replicates most of the Python API.  <br>- Includes structured help and is consistent with the Python API in almost every way.  <br>- CLI also includes `install/uninstall` commands which install right-click on supported files functionality on MacOS/Windows.  <br>  <br>**Simplified API**  <br>  <br>- Changed the `input_df` argument to `data` and opened it up to accept dataframes, dictionaries of dataframes, and data files.  <br>- Enabled support for csv, feather, parquet, excel, duckdb, sqlite, rdata, xml, json, and pickle files.  <br>- Created an `input_df` placeholder API for backwards compatibility.  <br>- Changed the `add_plots` argument to `plots`.  <br>  <br>**Significantly overhauled the experience when using multiple dataframes**  <br>  <br>- Converted the **Overview** worksheet to a landing page which:  <br>- Includes a dataframe-level table that tracks how many of each dataframe's fields have definitions.  <br>- Includes a field-level table that collates metadata, notes, and definitions from all fields across all dataframes in one table.  <br>- Has links to each dataframe's worksheet and to each field within each worksheet. <br>- Each worksheet includes links back to the **Overview**.  <br>  <br>**Persistent Settings**  <br>  <br>- Most xleda settings depend on the provided data except for two: `no_vba` and `theme_color`.  <br>- Changing these settings will change the default so that you can set them and forget about it.  <br>- Set favorite/company color or decide whether to use vba or not once and for all.  <br>  <br>**Other Updates**  <br>  <br>- Simplified/updated documentation where possible/necessary.  <br>- Removed the pivot worksheet/functionality.  <br>- Corrected the matplotlib headless backend again.  <br>- Amended the PythonList UDF to also accept arrays which lets it create Python lists using results from other Excel functions that return arrays like Filter, SumProduct, etc.  <br>- Removed the PythonDict UDF.<br><br></details><br>                                                                                                                                                                                                                    |



