



<div align="center">
    <a href="https://www.apache.org/licenses/LICENSE-2.0.txt"><img src="https://img.shields.io/badge/license-Apache-**blue**"></a> <a href="https://pypi.org/project/xleda"><img src="https://img.shields.io/pypi/v/xleda"></a> <a href="https://pypi.org/project/xleda"><img src="https://img.shields.io/pypi/pyversions/xleda.svg"></a> <a href="https://pepy.tech/project/xleda"><img src="https://static.pepy.tech/badge/xleda"></a> <a href="https://github.com/InfoDesigner/xleda"><img src="https://img.shields.io/badge/Made%20By%20A%20Human-99%25-blue)"></a> <a href="https://buymeacoffee.com/informationdesigner"><img src="https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ffdd00?&logo=buy-me-a-coffee&logoColor=black"></a>
</div><br>

<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/logo.webp?raw=true" width="250" alt="Logo">
	<br>
</p>

xleda is a Python/Excel powered EDA tool that creates workbooks from dataframes or data files that are highly optimized to explore, define, and document data sets.<br><br>

* Works on Windows or MacOS as a Python package, a CLI, or as a service that lets you create workbooks by right-clicking supported files.<br><br>

* There are some amazing EDA tools available to data professionals. You shouldn't have to start from scratch to include Microsoft Excel among them.<br><br>

* See some example xleda workbooks <a href="https://github.com/InfoDesigner/xleda/tree/main/examples">here</a>, docs <a href="https://infodesigner.github.io/xleda/">here</a>, and a  quick-start guide for non-developers <a href="https://infodesigner.github.io/xleda/#xleda-for-non-developers">here</a>.<br><br>





<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/top_view.webp?raw=true"  width="800" alt="Example Top View"> 
	<br>
	<em>Top view of a Field Analysis worksheet.</em>
</p><br><br>

<hr><br>

## xleda Components<br><br>


All xleda workbooks include an **Overview** worksheet and a **Field Analysis** worksheet for each provided dataframe.<br><br>


<table width="100%">
  <tbody>
    <tr>
      <!-- width: 1% forces this column to stay as narrow as its text -->
      <td valign="top" style="width: 1%; white-space: nowrap; padding-right: 20px;">
        <strong>Field&nbsp;Analysis</strong>
      </td>
      <td valign="top">
        <details>
          <summary>Anatomy of a Field Analysis Worksheet<br><br></summary>
          <p align="left">
            <img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/field_analysis_anatomy.webp?raw=true" style="width: 100%; max-width: 100%;" alt="Field Analysis Anatomy">
          </p>
          <em>Field Analysis Anatomy</em>
        </details>
        <br>
      </td>
    </tr>
    <tr>
      <td valign="top" style="width: 1%; white-space: nowrap; padding-right: 20px;">
        Overview
      </td>
      <td valign="top">
        <details>
          <summary>Anatomy of an Overview Worksheet<br><br></summary>
          <p align="left">
            <img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/overview_anatomy.webp?raw=true" style="width: 100%; max-width: 100%;" alt="Overview Anatomy">
          </p>
          <em>Overview Worksheet with Multiple Dataframes</em>
        </details>
        <br>
      </td>
    </tr>
  </tbody>
</table>




<hr><br>

## Requirements/Compatibility<br><br>


<table>
  <tbody>
    <tr>
      <td width="20%" valign="top"><strong>Desktop&nbsp;Excel</strong></td>
      <td valign="top">
      <p>Requires the full version of Microsoft Excel (2016+) on either MacOS or Windows to create workbooks</p>
        <ul>
          <li>See MacOS Support section below for details on MacOS usage.</li><p></p>
        </ul>
      </td>
    </tr>
    <tr>
      <td valign="top"><strong>Supported&nbsp;Data</strong></td>
      <td valign="top">
      <p>Supports pandas dataframes, CSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML</p><p></p>
      </td>
    </tr>
  </tbody>
</table><br>



          
<details markdown="1">
  <summary><strong>Tips:</strong> Managing the Install<br><br></summary>
  
* `xleda install` adds right-click functionality to your OS but it does not modify your path to make the CLI available systemwide<p></p>
  
* If the Python environment that xleda was installed into is deleted after running `xleda install`, the right-click functionality will need to be either repaired or uninstalled by running `xleda install/xleda uninstall` from a new Python environment.<p></p>

* If you have UV installed, you can install the package, CLI, and right-click menus systemwide without having to maintain a venv with these two lines.<p></p><br>

**Windows or MacOS**

```bash
# Installs the package and makes the xleda command available
uv tool install xleda

# Note you may need a new terminal window to see the newly installed xleda command
xleda install
```

<br>

</details><br><br>

<hr><br>


## xleda for Non-Developers<br><br>

<p>If you're working with data professionally in any role and find yourself looking at foreign data, one of the most important things you can do is document and define your data so that you can ensure everyone is working with the same data and definitions.</p>

<p>xleda can help you perform this task easily, quickly, and without having to write a single line of Python code.</p><br>

<details id="non_developer_guide">
<summary>Non-Developer Quick Start Guide<br><br></summary><br><br>
  <p><strong>Following the steps below will provide you with:</strong></p>
  <ol>
    <li>A comprehensive document for your data with worksheets for each related data source and placeholders for field definitions and notes that you can share with other contributors</li><br><br>
    <li>The ability to create the same workbooks in the future by right-clicking on your source data files and choosing <strong>Create xleda Workbook<br><br><br></strong></li>
  </ol>
<table>
  <tbody>
    <tr>
      <th align="left" width="30%" valign="top"><strong>1. Prepare Your Source Data</strong></th>
      <td>
        <p>We'll start by gathering your source data into one place that we can provide to xleda</p><br>
        <ul>
          <li>If you're only working with one data source and you already have it in a single file such as a CSV/DuckDB/Parquet/SQLite/JSON/XML file, you can skip this step.</li><br>
		      <li>Using an Excel workbook as a data souce for xleda, will let you create an xleda workbook that includes data from multiple sources</li><br>
          <li>Get an export or dump from each data source.  You can include as many records for each source as you want though xleda will use a sample of 25,000 rows from each by default.</li><br>
		      <li>Create an Excel workbook that has a <a href="https://www.google.com/search?q=creating+named+excel+tables&oq=creating+named+excel+tables">proper Excel Table for each data source</a>.</li><br>
		      <li>Be sure to name each table clearly.  Worksheet names in this workbook will be ignored by xleda but the table names will be used as worksheet names in your xleda workbook.</li><br>
        </ul>
      </td>
    </tr>
    <tr>
      <th align="left" valign="top"><strong>2. Install UV</strong></th>
      <td>
        <p>UV will be installed to manage Python</p><br>
          <ul>
            <li>It can be installed safely and easily without affecting other software on your system.</li><br>
		        <li>They have <a href="https://docs.astral.sh/uv/getting-started/installation/">one-line installs for all operating systems</a></li><br>
		        <li>See their site for a variety of install options or use one of the TLDR versions below in your operating system's terminal</li><br><br>
          </ul>
		  <details>
			<summary><strong>Windows PowerShell</strong><br><br></summary>
      
```bash
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```
</details><br>
		  <details>
			<summary><strong>MacOS</strong><br><br></summary>
      
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```
</details><br>
      </td>
    </tr>
  </tbody>
  <tbody>
    <tr>
      <th align="left" valign="top"><strong>3. Install xleda</strong></th>
      <td>
        <p>UV will be used to install xleda</p><br>
          <ul>
            <li>If step 2 was the first time you've installed UV, you'll need to open a new terminal window to use the newly installed <code>uv</code> command in this step.  See <a href="https://www.google.com/search?q=how+to+get+the+uv+command+to+show+after+install%3F">this</a> if you run into issues</li><br>
            <li>Now that UV is installed and the <code>uv</code> command is available, you can install xleda on either MacOS/Windows with</li><br>
          </ul>

```bash
uv tool install xleda
```
</td>
    </tr>
    <tr>
      <th align="left" valign="top"><strong>4. Install right-click functionality</strong></th>
      <td>
        <p>xleda will be used to install right-click on supported files functionality</p>
          <ul>
            <li>This step lets you create xleda workbooks without any terminal commands in the future</li><br>
            <li>If step 3 was the first time you've installed xleda, you'll need to open a new terminal window to use the newly installed <code>xleda</code> command in this step.  See <a href="https://www.google.com/search?q=how+to+ensure+uv+tool+is+installed+to+path%3F&oq=how+to+ensure+uv+tool+is+installed+to+path%3F">this</a> if you run into issues</li><br>
            <li>Now that xleda is installed and the <code>xleda</code> command is available, we'll install right-click functionality by running:</li><br>

```bash
xleda install
```
</ul>
      </td>
    </tr>     
    <tr>
      <th align="left" valign="top"><strong>5. Create an xleda workbook</strong></th>
      <td>
        <p>This step lets you create your workbook and choose your theme for future xleda workbooks at the same time</p><br>
          <ul>  
            <li>If you're fine using the default neutral theme, you can skip this step and instead right-click on the workbook you compiled in step 1 and choose <strong>Create xleda Workbook</strong></li><br>  
            <li>Create your workboook on either MacOS or Windows with the same command.</li><br>
            <li>Be sure to Replace <code>YourSourceData</code> and <code>305CDE</code> with your source data and favorite color</li><br><br>
            <code>xleda wb YourSourceData.xlsx --theme '#305CDE'</code><br><br>
        </ul>
      </td>
    </tr>
  </tbody>
</table><br>



<strong>What's the Catch?</strong><br>

Now that everything is installed, this guide is no longer necessary.  You can now create workbooks in the future without any terminal commands by right-clicking on supported files.<br><br>

If you want to change your theme, use that very last line to create a new workbook once and it will remember your preference.  VBA preference persists in the same way after it's set.  Run <code>xleda wb --help</code> in your terminal for guidance on how to set either with/without creating a workbook and other settings.<br><br>

What's the Catch?  There's not one really.  The only thing that's close to a catch is that using this method doesn't include automatic updates.   If you've followed these steps, be sure to periodically update xleda using `uv tool upgrade xleda`.<br><br>



</details><br><br>

<hr><br>


## Basic Usage<br><br>

Use <code>wb()</code> to quickly create an xleda workbook from a dataframe, a dictionary of dataframes, or a supported data file.<br><br>

#### From a Dataframe

```python
from xleda import wb
import seaborn as sns

# < your dataframe goes here >
df = sns.load_dataset("titanic")

# Creates xleda.xlsm in the current directory
wb(df)
```

<br><br>

#### From a Dictionary of Dataframes

```python
from xleda import wb
import seaborn as sns

# < your dataframes go here >
df1 = sns.load_dataset("titanic")
df2 = sns.load_dataset("penguins")

# Creates Titanic.xlsm in the current directory
wb({"Titanic": df1,
    "Penguins": df2})
```

<br><br>

#### From a File

```python
from xleda import wb
from pathlib import Path

# < your data file goes here >
duckdb_file = "https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/data/duckdb.duckdb"

# Creates duckdb.xlsm in the current directory
# Includes data from all tables in the db file
wb(duckdb_file)
```

<br><br>

#### From the CLI

```bash
# Creates 'userdata.xlsm' in the current directory
xleda wb 'https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/data/userdata.parquet'

# Shows the help command
xleda --help

# Shows the wb help command
xleda wb --help
```

<br><br>

#### From Right-Clicking


<table>
  <tbody>
    <tr>
      <td width="50%" valign="top"><br><strong>Windows</strong><br><br></td>
      <td width="50%" valign="top"><br><strong>MacOS</strong><br><br></td>
    </tr>
    <tr>
      <td valign="top"><br>
        <p align="left">
          <img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/right_click_win.webp?raw=true" width="400" alt="From right-click Win">
        </p><br>
      </td>
      <td valign="top"><br>
        <p align="left">
          <img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/right_click_mac.webp?raw=true" width="400" alt="From right-click MacOS">
        </p><br>
      </td>
    </tr>
  </tbody>
</table><br><br>

<hr><br>

## xleda.wb() Configuration<br><br>


<table>
  <tbody>
    <tr>
      <td width="20%" valign="top"><strong><br><code>data</code></strong>
      </td>
      <td valign="top"><br>
        <strong>Dataframe or Path or string | Mandatory</strong><br><br>
        <ul>
          <li>Accepts a pandas dataframe, a dictionary of dataframes, or a supported data file</li><br>
          <li>For files, xleda will create a workbook from all tabular objects in the file</li><br>
          <li>Supported types include CSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML</li><br>
        </ul>
          <details>
            <summary>Data File Limitations<br><br></summary><br>
            <p>If the provided data file doesn't parse correctly, try creating a 
            dataframe first and use that with xleda instead of the file</p><br>
            <strong>Expect problems with</strong>:<br><br>
            <ul>
              <li>Deeply nested JSON/XML</li><br>
              <li>A <code>.CSV</code> file with tabs instead of commas</li><br>
              <li><code>.db</code> files that are neither SQLite nor DuckDB files</li><br>
              <li>DuckDB files with a <code>.txt</code> extension</li><br>
            </ul><br>
          <strong>Don't expect to see</strong>:<br><br>
            <ul>
              <li>Anything from an RData file that isn't a dataframe</li><br>
              <li>Dataframes that are nested somewhere inside a pickle file</li><br>
              <li>Anything sourced from an Excel file that isn't a proper Excel Table</li><br>
            </ul>
          </details><br>
      </td>
    </tr>
    <!-- file name -->
    <tr>
      <td width="30%" valign="top"><strong><br><code>file_name</code></strong>
      </td>
      <td valign="top"><br>
        <strong>str | Optional</strong><br><br>
        <ul>
          <li>The workbook file name to create.</li><br>
          <li>Defaults to same name as data files provided for <code>data</code>, the first key in a dataframe dict provided for <code>data</code>, or <code>xleda</code></li><br>
        </ul>
      </td>
    </tr>
    <!-- wb path -->
    <tr>
      <td width="30%" valign="top"><strong><br><code>wb_path</code></strong>
      </td>
      <td valign="top"><br>
        <strong>Path or string | Optional</strong><br><br>
        <ul>
          <li>Use a directory or file path</li><br>
          <li>If a directory is provided, the workbook is created there</li><br>
          <li>If a filename ends with <code>.xlsm</code> or <code>.xlsx</code>, xleda will create or export from that file</li><br>
          <li>Defaults to the current working directory or the source file directory</li><br>
        </ul>
      </td>
    </tr>
    <!-- theme -->
    <tr>
      <td width="30%" valign="top"><strong><br><code>theme</code></strong>
      </td>
      <td valign="top"><br>
        <strong>str | Optional</strong><br><br>
        <ul>
          <li>Sets the primary workbook theme</li><br>
          <li>Accepts a hex color or <code>random</code></li><br>
          <li>Defaults to a neutral color.</li><br>
        </ul>
        <p align="center"><br>
          <img
            src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/theme.webp?raw=true"
            width="800"
            alt="Theme"
          /><br>
          <em>theme affects the workbooks and default charts.</em><br>
        </p><br>
      </td>
    </tr>
    <!-- plots -->
    <tr>
      <td width="30%" valign="top"><strong><br><code>plots</code></strong>
      </td>
      <td valign="top"><br>
        <strong>dict | Optional</strong><br><br>
        <ul>
          <li>Adds extra plot worksheets using a dict of matplotlib Figure objects</li><br>
          <li>Accepts <code>{'plotname': Figure, ...}</code> format</li><br>
          <li>No automatic styling or sizing is applied</li><br>
        </ul>
      </td>
    </tr>
    <!-- overwrite -->
    <tr>
      <td width="30%" valign="top"><strong><br><code>overwrite</code></strong>
      </td>
      <td valign="top"><br>
        <strong>bool | Optional</strong><br><br>
        <ul>
          <li>Overwrites existing workbooks of the same name</li><br>
          <li>Existing files are moved to Trash/Recycle Bin</li><br>
          <li>Defaults to <code>False</code></li><br>
        </ul>
      </td>
    </tr>
    <!-- large_report -->
    <tr>
      <td width="30%" valign="top"><strong><br><code>large_report</code></strong>
      </td>
      <td valign="top"><br>
        <strong>bool | Optional</strong><br><br>
        <ul>
          <li>Raises <code>data</code> limits to Excel's maximum: 1,000,000 rows and 16,000 columns</li><br>
          <li>Requires more memory and time for large datasets</li><br>
          <li>Defaults to <code>False</code></li><br>
        </ul>
      </td>
    </tr>
    <!-- no_vba -->
    <tr>
      <td width="30%" valign="top"><br><strong><code>no_vba</code></strong>
      </td>
      <td valign="top"><br>
        <strong>bool | Optional</strong><br><br>
        <ul>
          <li>Creates a <code>.xlsx</code> workbook without VBA</li><br>
          <li>Setting this flag persists the preference so that you can set it and forget it</li><br>
          <li>Use an <code>.xlsx</code> file for <code>wb_path</code> as an alternative though this won't persist</li><br>
          <li>Defaults to <code>False</code></li><br>
        </ul>
      </td>
    </tr>
    <!-- open_wb -->
    <tr>
      <td width="30%" valign="top"><br><strong><code>open_wb</code></strong>
      </td>
      <td valign="top"><br>
        <strong>bool | Optional</strong><br><br>
        <ul>
          <li>Opens the workbook after creation</li><br>
          <li>Set to False when creating multiple workbooks</li><br>
          <li>Defaults to <code>True</code></li><br>
        </ul>
      </td>
    </tr>
    <!-- export -->
    <tr>
      <td width="30%" valign="top"><br><strong><code>export</code></strong>
      </td>
      <td valign="top"><br>
        <strong>bool | Optional</strong><br><br>
        <ul>
          <li>Exports data from an xleda workbook instead of creating one</li><br>
          <li>See the Examples/Exporting Metadata sections below for details</li><br>
          <li>Defaults to <code>False</code></li><br>
        </ul>
      </td>
    </tr>
  </tbody>
</table><br><br>

<hr><br>

## CLI Basics<br><br>

<p>Installing the Python package also installs the <code>xleda</code> CLI command<br><br>

It works almost the same way as the Python API except that it only accepts files for <code>data</code> and doesn't accept the <code>plots</code> argument</p><br>


<details>
  <summary>CLI Commands<br><br></summary>

<table>
  <tbody>
    <tr>
      <td width="20%" valign="top"><br><code>xleda --help</code><br><br></td>
      <td>Shows the xleda help menu</td>
    </tr>
    <!-- xleda wb help -->
    <tr>
      <td valign="top"><br><code>xleda wb --help</code><br><br></td>
      <td>Shows help for the wb command and it's flags</td>
    </tr>  
    <!-- xleda install -->
    <tr>
      <td valign="top"><br><code>xleda install</code><br><br></td><br>
      <td>Installs right-click on supported files to create workbooks functionality</td>
    </tr>  
    <!-- `xleda uninstall -->
    <tr>
      <td valign="top"><br><code>xleda uninstall</code><br><br></td>
      <td>Uninstalls right-click on supported files to create workbooks functionality</td>
    </tr>  
    <!-- xleda version -->
    <tr>
      <td valign="top"><br><code>xleda version</code><br><br></td>
      <td>Compares your installed version with the latest available version on PyPi</td>
    </tr>
    <!-- xleda vba -->
    <tr>
      <td valign="top"><br><code>xleda vba</code><br><br></td>
      <td>This toggles your preference for creating workbooks with/without VBA and persists once set.</td>
    </tr> 
    <!-- xleda theme -->
    <tr>
      <td valign="top"><br><code>xleda theme</code><br><br></td>
      <td><br>
        This changes your theme preference without creating a workbook and persists once set<br><br>
        
```bash
# Sets theme to a dark grey
xleda theme '#262626'

# Also sets theme to a dark grey
xleda theme 262626
```
<br>
        </td>
    </tr>
  </tbody>
</table><br><br>

</details><br>


<hr><br>


## Examples<br><br>


<details>
  <summary>Example: Creating a workbook from multiple dataframes<br><br></summary><br>

```python
import seaborn as sns
from xleda import wb

seaborn_datasets = ['diamonds', 'dots', 'dowjones']
dataframe_dict = {df_name: sns.load_dataset(df_name) for df_name in seaborn_datasets}

# Creates diamonds.xlsm in the current directory
# Also includes dots and dow jones data
wb(data=dataframe_dict)
```

<hr><br>

</details>

<details>
  <summary>Example: Using wb_path as a directory or a file<br><br></summary><br>

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

<hr><br>

</details>
<details>
  <summary>Example: Adding custom plots to a workbook<br><br></summary><br>

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
   theme="#4C4C4C",
   plots={'Pair Plots': pair_plots,
          'Null Matrix': null_matrix})
```

<hr><br>

</details>


<details>
<summary>Example: Creating workbooks without VBA<br><br></summary><br>

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

<hr><br>

</details>

<details>
<summary>Example: Creating a workbook from a database<br><br></summary><br>

<strong>From Python</strong>
```python
from xleda import wb

# < your database goes here>
sqlite_db = "https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/data/chinook.db"

# Creates "Chinook.xlsm" in the current directory with 11 dataframes
wb(data=sqlite_db,
   file_name="Chinook")
```

<strong>From the CLI</strong>
```bash
# Creates "Chinook.xlsm" in the current directory with 11 dataframes
xleda wb chinook.db --name "Chinook"
```

<hr><br><br>

</details>



<details>
<summary>Example: Basic Metadata Export<br><br></summary><br>

Basic metadata export sources data from Python

```python
from xleda import wb
import seaborn as sns

# < your dataframe goes here >
df = sns.load_dataset("titanic")
  
# Creates "Titanic.xlsm" and returns basic metadata
export_dicts = wb(data={"Titanic": df},
                  file_name="Titanic").export_dicts

# returns ['field_overview', 'df_overview', 'source_data']
print(export_dicts[0].keys())
```

<hr><br><br>

</details>

<details markdown="1">
<summary>Example: Full Metadata Export<br><br></summary><br>

Full export sources data from the workbook when possible<br><br>

* The xleda workbook pictured here is used in for the export code example below .  

* It can be found [here.](https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/Titanic%20Completed.xlsm).<br><br>

<p align="center">
  <img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/completed_field_analysis.webp?raw=true" width="800" alt="Completed Field Analysis">
	<br>
	<em>A completed xleda workbook showing definitions, notes, lists, etc.</em>
</p>
<br>


```python
from xleda import wb
import seaborn as sns

# < your dataframe goes here >
df = sns.load_dataset("titanic")

# < your completed workbook goes here >
edited_workbook_path = "https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/Titanic%20Completed.xlsm"

# Performs a full export from "Titanic Completed.xlsm"
export_dicts = wb(data={"Titanic": df},
                  wb_path=edited_workbook_path,
                  export=True).export_dicts

# Returns ['description', 'definitions', 'notes', 'lists', 'field_overview', 'df_overview', 'source_data']
print(export_dicts[0].keys())
```

---

</details><br><br>


<hr><br>

## Usage Notes<br><br>


<table>
  <!-- Field and Record Lists -->
  <tbody>
    <tr>
      <td width="30%" valign="top"><br>Field&nbsp;and&nbsp;Record&nbsp;Lists</td>
      <td valign="top"><br>
        The <code>Field Lists</code> section includes placeholders to create 8 custom lists of fields<br><br>
        <ul>
          <li>Use these to organize fields into groups such as "fields_to_delete", "fields_from_system_a", "fields_to_fix", or whatever your workflow needs</li><br>
          <li>The Record List works similarly though it tags individual records instead of lists</li><br>
        </ul>
        <details>
          <summary>List Details<br><br></summary>
          <ul><br>
            <li>Anything not marked as False will be included in each list</li><br>
            <li>You can rename any list to <code>Anything You Want</code> and the list will be renamed to <code>anything_you_want</code></li><br>
            <li>The <code>Record List</code> field added to your source data works the same way except it creates a list of all tagged records instead of a list of fields</li><br>
            <li>The <code>Compiled Lists</code> section formats your lists as python lists</li><br>
            <li>VBA workbooks also include an Excel function, <code>PythonList</code>, that creates Python formatted lists out of cell values</li><br>
          </ul>
          <p align="center">
            <img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/field_lists.webp?raw=true" width="800" alt="Field Lists">
            <br>
            <em>Easily create lists of fields in your data.</em>
          </p>
        </details><br>
      </td>
    </tr>
    <!-- Large Data Sets -->
    <tr>
      <td width="30%" valign="top"><br>Large Data Sets</td>
      <td valign="top"><br>
        On an average machine, xleda creates workbooks for most data sets less than 20 seconds on Windows/1-2 minutes on MacOS<br><br>
        <ul>
          <li>To ensure workbooks are created quickly, each dataframe is by default subsampled to only include the first 50 columns and a random sample of 25,000 records.</li><br>
          <li>You can optionally override default limits to use Excel's limits of 16,000 columns, 1,000,000 rows by using <code>large_report=True</code></li><br>
        </ul>
        <details>
        <summary>Performance Details<br><br></summary><br>
          <ul>
            <li>Performance is largely dependent on how powerful of a machine you have and how many/how large/how complex your dataframes are</li><br>
            <li>You'll see a warning banner on Field Analysis worksheets of affected dataframes if they've exceeded a limit</li><br>
            <p align="center">
            <img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/warning.webp?raw=true" width="800" alt="Warning Banner">
            </p>
            <li>The <code>debug</code> section of the <code>Overview</code> worksheet has a breakdown of how the time spent to produce your workbook was allocated.</li>
          </ul>
        </details><br>
      </td>
    </tr>
    <!-- Exporting Metadata -->
    <tr>
      <td width="30%" valign="top">Exporting Metadata</td>
      <td valign="top"><br>
        Accessing your notes/lists/defintions from Python is easy<br><br>
        <ul>
          <li>Metadata from all <code>xleda.wb()</code> objects is collected into a list of dictionary objects, one for each dataframe, accessible through <code>xleda.wb().export_dicts</code></li><br>
          <li>You can also access expanded metadata, sourced from the workbook by using <code>export=True</code></li><br>
        </ul>
        <details>
        <summary>Default Metadata<br><br></summary><br>
          The default metadata is the same field and dataframe metadata that is added to the workbooks and is available without using <code>export=True</code><br><br>
          <ul>
            <li><code>df_overview</code>: Dataframe level metadata</li><br>
            <li><code>field_overview</code>: Field-level metadata</li><br>
            <li><code>field_metadata</code>: A basic metadata dataframe, combining information from pandas info/describe/quantile</li><br>
            <li><code>source_data</code>: A copy of the source data that also includes Record Hash/Record List/HasBlank/index columns</li><br>
          </ul>
        </details><br>
        <details>
          <summary>Expanded Metadata:<br><br></summary><br>
          Using <code>export=True</code> also provides the default metadata though it is sourced from the workbook instead.<br><br>
          This includes your notes, lists, definitions, etc. and will reflect any changes you've made in Excel such as renaming fields/deleting values/etc.<br><br>
          Includes the following for each provided dataframe:<br><br>
          <ul>
            <li><code>description</code>: Dataframe description if you've added one</li><br>
            <li><code>definitions</code>: Any field definitions you've added</li><br>
            <li><code>notes</code>: Any field notes you've added</li><br>
            <li><code>lists</code>: Any lists showing in the compiled lists section</li><br>
            <li><code>lists</code>: Note that data types will likely change in the round-trip translation</li><br>
          </ul>
        </details><br>
      </td>
    </tr>
    <!-- MacOS Support -->
    <tr>
      <td width="30%" valign="top">MacOS Support</td>
      <td valign="top"><br>
       xleda will create the same workbooks in MacOS<br><br>
        <ul>
          <li>Creating them is significantly slower and you may get two different types of prompts that require your attention</li><br>
          <li>Look for the bouncing Excel icon</li><br>
        </ul>
      <details>
        <summary>MacOS Details<br><br></summary>
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
              <td>MacOS</td><br>
              <td>Excel</td><br>
            </tr>
            <tr>
              <td><strong>Details</strong></td>
              <td>Prompts to Allow Excel to access the file it's creating.<br><br>If you get these prompts, you'll potentially get one for each unique file you create.</td>
              <td>Prompts to "Enable Macros".  <br><br>If you get these prompts, you'll get two when creating a workbook:<br><br>1. When opening the blank template<br><br>2. When opening your created workbook.<br><br></td>
            </tr>
            <tr>
              <td><strong>Example</strong></td>
              <td><img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/grant_file_access.webp?raw=true" width="350" alt="Grant file access prompt"></td>
              <td><img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/enable_macros.webp?raw=true" width="350" alt="Enable macros prompt"></td>
            </tr>
            <tr>
              <td><strong>Remedy</strong></td>
              <td>There's not a reliable remedy to this.<br><br>MacOS doesn't permit applications like Microsoft Excel real access to the file system, even after explicitly granting Excel Full Disk Access under <br><br><code>Settings &gt; Privacy & Security &gt; Full Disk Access</code>.</td>
              <td>You can either:<br><br>1. Create a VBA free workbook (see the next section for details).<br><br>2. Change Excel's default macro settings (shown) to one of the other two options.<br><br><img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/excel_macos_macro_options.webp?raw=true" width="350" alt="Excel MacOS macro settings"></td>
            </tr>
          </tbody>
        </table>
      </details><br>
      </td>
    </tr>
    <!-- VBA Code -->
    <tr>
      <td width="30%" valign="top"><br><strong>VBA Code</strong></td>
      <td valign="top"><br>
        The included <a href="https://github.com/InfoDesigner/xleda/blob/main/src/xleda/assets/mdl_xleda.bas">VBA code</a> is short and easy to understand<br><br>
        <ul>
          <li>You can create a VBA-free, xlsx workbook by either setting <code>no_vba=True</code> or providing a <code>wb_path</code> ending in <code>.xlsx</code></li><br>
          <li>Providing the <code>no_vba flag</code> will change the default so that the setting will persist. Using <code>wb_path</code> doesn't work this way.</li><br>  
        </ul>
      <details>
        <summary>What the VBA Code Does<br><br></summary><br>
        <ol>
          <li>Makes the sections expand/collapse when you select them as pictured on the left which can also be performed by using row groupings as pictured on the right</li><br>
        </ol>
        <table>
          <tr>
            <td align="center">
              <img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/top_view.webp?raw=true" width="400" alt="Row Groupings">
              <br>
              <em>Use headings like web pages to navigate with VBA.</em>
            </td>
            <td align="center">
              <img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/row_groupings.webp?raw=true" width="400" alt="Row Groupings">
              <br>
              <em>Use row groupings to navigate without VBA.</em>
            </td>
          </tr>
        </table><br>
        <strong>The Annoyance Cost</strong><br><br>
        Besides <strong>Enable Macros</strong> prompts, using VBA also includes one annoying side effect:<br><br>
        <ul>
          <li>Every time a macro is used, it clears your undo history</li><br>
        </ul>
      </details><br>
      </td>
    </tr>
    </tbody>
</table><br><br>

<hr><br><br>

## Troubleshooting<br><br>

<details>
<summary>xleda is slow<br><br></summary><br>
  <ul>
    <li>Try reducing the amount of data you're sending to it, and let it finish.</li><br>
    <li>After production, refer to the `debug` section of the `Overview` worksheet for how the time to produce your workbook is being spent.</li><br>
    <li>Note that on MacOS, `xleda` is much slower by default and the timings in the debug section may be inflated from missed permission prompts during production.</li><br>
  </ul>
</details><br>

<details>
  <summary>"Error: The workbook cannot be overwritten while open!" and you don't see any open workbooks<br><br></summary><br>
    <ul>
      <li>You may have a hidden Excel instance that needs to be closed.</li><br>
      <li>Guidance on closing hidden Excel windows for <a href="https://www.google.com/search?q=hidden+excel+instance+in+macos">MacOS</a> and <a href="https://www.google.com/search?q=hidden+excel+instance+in+windows">Windows</a></li><br>
    </ul>
</details><br>

<details> 
  <summary>If you receive the "Exception: Could not activate App!" or "The RPC server is unavailable". errors:<br><br></summary><br>
  <ul>
    <li>The Excel app may have crashed or is otherwise disconnected from Python.</li><br>
    <li>Close all Excel windows and try running the command again.</li><br>
  </ul>
</details><br>


<details markdown="1">
  <summary>xleda won't run at all and are using Windows/MacOS with a full Office Installation<br><br></summary><br>
    <ul>
      <li>If you can get the script below to run successfully using xlwings (not xlwings-lite), xleda has a good chance of working reliably.</li><br>
      <li>All it does is open Excel and create a new workbook.</li><br>
      <li>You should be able to `pip install xlwings` and run the script successfully.</li><br>
      <li>If that doesn't work, see their <a href="https://docs.xlwings.org/en/latest/installation.html">installation instructions</a> for details on how to get it set up.</li><br>
      <li>Be aware that xlwings has a ton of functionality and that for xleda to work, it only requires communication with Excel and not the addin, xlwings lite, udfs, or many of the other things xlwings can potentially do.</li><br>
      <li>If you can't get it to work and you're on Windows, <a href="https://www.google.com/search?q=win32+com+corruption+xlwings">this may help</a>.</li><br><br>
    </ul>
          
```python
import xlwings as xw
app = xw.App()
```
</details><br>

<hr><br>

## Changelog<br><br>


<table>
<!-- Version 0.8.185 -->
  <tbody>
    <tr>
      <td width="30%" valign="top"><br><strong>Version&nbsp;0.8.185</strong></td>
      <td valign="top"><br>
        <details>
          <summary>New simplified API, simplified export, general polish<br><br></summary><br>
            <strong>Simplified basic usage to make it quicker to use and easier to memorize.</strong><br><br>
              <ul>
                <li>Changed the default entry point to <code>xleda.wb()</code> from <code>xleda.FieldAnalysis()</code></li>
                <li><code>xleda.wb()</code> now creates and automatically opens workbooks</li>
                <li>The only argument needed to create a workbook is now a dataframe: <code>wb(df)</code></li>
                <li>Workbook name now defaults to <code>xleda</code> if no name is given</li>
                <li>Protected backwards compatibility while providing guidance to use the new API</li>
                <li>Subclassed the new API to create plugs for the old one</li>
              </ul><br>
            <strong>Simplified export functionality</strong><br>
              <ul>
                <li>Changed <code>export_analysis</code> functionality from a class method to a class argument <code>wb(df, export=True)</code></li>
                <li>All <code>wb()</code> objects now include a <code>export_dict</code> metadata collection that is accessible using dot notation</li>
                <li>Added <code>field_metadata</code> and <code>overview_metadata</code> to <code>export_dict</code></li>
                <li>Using <code>wb(export=True)</code> reads a workbook instead of creating one and adds the metadata from the workbooks to <code>export_dict</code></li>
                <li>Added file exists checks for <code>export=True</code> with messaging that the export will be limited if the file isn't found</li>
              </ul><br>
            <strong>Template updates</strong><br>
              <ul>
                <li>Recreated the template, moved formatting to cell styles for simplicity/consistency in maintenance where appropriate</li>
                <li>Pivot was removed and Blanks was renamed Pivot</li>
                <li><code>% of Records</code> field was added to the new Pivot</li>
                <li>Added dataframe index to source data by default</li>
                <li>Added dataframe level metadata to the Data Description section</li>
                <li>Added two UDFs to the template, PythonList/PythonDict, to create Python formatted strings from cell values</li>
                <li>Adjusted the named range to support deleting almost any column without affecting lists or navigation</li>
                <li>General polish</li>
              </ul><br>
            <strong>Other updates</strong><br>
              <ul>
                <li>Default limits were reduced to 25,000 rows/50 columns</li>
                <li>Good deal of refactoring to support the new entry point, minimize errors, reduce redundancy</li>
                <li>Removed clipboard usage in all except one place, where it is used for formatting instead of data</li>
                <li>Added <code>open_wb</code> argument to prevent automatically opening the workbook, useful when creating many workbooks</li>
                <li>Replaced rich progress bars with TQDM for better support in notebooks/vs code notebook/console environments</li>
                <li>When using <code>overwrite=True</code>, overwritten files now go to the recycle bin/trash.</li>
                <li>Console output includes messaging about these files</li>
                <li>Clarified/organized readme to support the new API/template</li>
                <li>Added production logging metrics so you can see how the time required to create a workbook was utilized.</li>
              </ul>
        </details><br>
      </td>
    </tr>
<!-- Version 0.8.186 -->
    <tr>
      <td valign="top"><br><strong>Version 0.8.186</strong></td>
      <td valign="top"><br>
        <details>
          <summary><strong>Add multiple dataframes, module refactoring into classes, added logging</strong><br><br></summary><br><br>
          <strong>Implemented add_dfs</strong><br>
          <ul>
            <li>Adds Field Analysis/Overview reports for each additional dataframe</li>
            <li>Pivot is only provided for the primary dataframe</li>
            <li>Useful for supporting or related data</li>
            <li>Worksheet names now include the dataframe name</li>
            <li>Each dataframe's worksheet set gets a greyscale gradient so they can be visually distinguished among worksheet tabs</li>
          </ul>
          <strong>Export adjustments</strong><br>
          <ul>
            <li>Implemented an ExportDict class to add structure to export functionalities</li>
            <li>To support the additional dataframes from add_dfs functionality, export_dict has been renamed to export_dicts and now provides a list of ExportDict objects, one for each provided dataframe</li>
            <li>ExportDict allows access to metadata through both dot notation and dict[key]</li>
            <li>Reinforced handling of modified export workbooks</li>
            <li>If a workbook is found but the expected worksheets aren't found (for example, if they've been deleted or renamed), it will export what it can and return a list of what wasn't found</li>
          </ul>
        <strong>Reinforced wb_path/name handling</strong><br>
          <ul>
            <li>wb_path now accepts strings or pathlib Path objects</li>
            <li>Also accepts full/partial paths with/without correct extensions.</li>
            <li>Providing a path ending in .xlsx or .xlsm will set no_vba to True/False respectively.</li>
            <li>Illegal characters are now properly stripped from provided names before use.</li>
          </ul>
        <strong>Added production logging/debug worksheet</strong><br>
          <ul>
            <li>The debug worksheet details how the time it took to produce the workbook was allocated on both field and workbook levels</li>
            <li>Also includes configuration and system details</li>
          </ul>
        <strong>Other Updates</strong><br>
          <ul>
            <li>Tests, examples, readme updated to reflect new functionality</li>
            <li>In the template, the Field Notes section of the Field Analysis worksheet was merged into the Data Description section</li>
            <li>Refactored the primary module into more specialized classes</li>
            <li>Configuration/environment/plotting/logging/theme all have their own classes</li>
            <li>Also implemented a new Blueprint class</li>
            <li>Workbooks are now constructed from a config object that includes a list of Blueprints</li>
            <li>Each provided dataframe gets its own Blueprint</li>
            <li>Improved handling of datatypes that are unsupported in Excel/xlwings such as TimeDelta</li>
            <li>Reinforced system configuration checks with more informative offramps for:
              <ul>
                <li>Unsupported system configurations</li>
                <li>Situations where necessary template components have been removed or renamed</li>
              </ul>
            </li>
            <li>Adjusted Github Action script to remove all but last changelog and convert the details/summary to standard markdown</li>
          </ul>
        </details><br>
      </td>
    </tr>
<!-- Version 0.8.193 -->
    <tr>
      <td width="30%" valign="top"><br><strong>Version 0.8.193</strong></td>
      <td valign="top"><br>
      <details>
        <summary><strong>Added MacOS support</strong><br><br></summary><br>
        <ul>
          <li>Used xlwings when possible, appscript/AppleScript/subprocess otherwise.</li>
          <li>Reduced OS branching when possible.</li>
          <li>Documentation/test/tools/examples updated to be cross platform.</li>
        </ul><br>
        <strong>Other Updates</strong><br>
        <ul>
          <li>Removed field logging from logging/template.</li>
          <li>Simplified some of the pivot configuration where possible.</li>
          <li>Added multi-threading for the progress bar which keeps the time elapsed ticking during longer iterations.</li>
        </ul><br>
        <strong>Template Adjustments</strong><br>
        <ul>
          <li>Moved the xlsx conversion to a pre-commit hook instead of an on-demand end-user task.</li>
          <li>Adjusted expand/collapse icons to use a more reliable cross-platform character.</li>
          <li>Removed navigation shapes from the xlsx template.</li>
        </ul>
      </details><br>
      </td>
    </tr>
<!-- Version 0.8.197 -->
    <tr>
      <td valign="top"><br><strong>Version 0.8.197</strong></td>
      <td valign="top"><br>
      <details>
        <summary><strong>Readme/pyproject.toml polish/minor fixes</strong><br><br></summary><br>
        <ul>
          <li>Moved code examples/troubleshooting/usage notes into details/summary blocks to reduce clutter in README</li>
          <li>Fixed a cross-platform formatting issue with the debug worksheet</li>
          <li>Updated a few older screenshots to use the current template</li>
          <li>Adjusted type: ignore lines where possible</li>
          <li>Organized pyproject.toml, added required-environments section</li>
        </ul>
      </details><br>
      </td>
    </tr>
    <!-- Version 0.9.010 -->
    <tr>
      <td valign="top"><br><strong>Version 0.9.010</strong></td>
      <td valign="top"><br>
      <details>
        <summary><strong>Simplified API, Expanded Input Options/Interfaces, Significantly improved experience with multiple dataframes</strong><br><br></summary><br>
        <strong>Simplified API</strong><br>
        <ul>
          <li>'input_df' was changed to 'data'</li>
          <li>'add_plots' was changed to 'plots'</li>
          <li>'name' was changed to 'file_name'</li>
          <li>'theme_color' was changed to 'theme'</li>
        </ul>
        <strong>Expanded Input Options</strong><br>
        <ul>
          <li>Changed the 'input_df' argument to 'data' and opened it up to accept dataframes, dictionaries of dataframes, and data files that are either local or http/https</li>
          <li>Enabled support for csv, feather, parquet, excel, duckdb, sqlite, rdata, xml, json, and pickle files</li>
          <li>A file that represents a tabular object, such as parquet/csv files, will create a dataframe and then create a workbook from that dataframe</li>
          <li>Complex files, such as duckdb, RData, or sqlite files, will create dataframes from all tabular objects within each file and create a workbook with those dataframes</li>
          <li>Created an input_df placeholder API for backwards compatibility.</li>
        </ul>
        <strong>Expanded Interfaces</strong><br>
        <ul>
          <li>Added a CLI interface which replicates most of the Python API</li>
          <li>Includes structured help and is consistent with the Python API in almost every way</li>
          <li>CLI also includes install/uninstall commands which install right-click on supported files functionality on MacOS/Windows</li>
          <li>Added vba toggle and theme setting to cli</li>
        </ul>
        <strong>Significantly overhauled the experience when using multiple dataframes</strong><br>
        <ul>
          <li>Converted the Overview worksheet to a landing page which:
            <ul>
              <li>Includes a dataframe-level table that tracks how many of each dataframe's fields have definitions</li>
              <li>Includes a field-level table that collates metadata, notes, and definitions from all fields across all dataframes in one table</li>
              <li>Has links to each dataframe's worksheet and to each field within each worksheet</li>
              <li>Each worksheet includes links back to the Overview.</li>
            </ul>
          </li>
        </ul>
        <strong>Persistent Settings</strong><br>
        <ul>
          <li>Most xleda settings depend on the provided data except for two: no_vba and theme.</li>
          <li>Changing these settings will change the default so that you can set them and forget about it.</li>
          <li>Set favorite/company color or decide whether to use vba or not once and for all.</li>
          <li>Added vba toggle and theme setting to cli to affect these settings without creating workbooks</li>
        </ul>
        <strong>Other Updates</strong><br>
        <ul>
          <li>Changed the add_plots argument to plots.</li>
          <li>Simplified/updated documentation where possible/necessary.</li>
          <li>Removed the pivot worksheet/functionality.</li>
          <li>Corrected the matplotlib headless backend again.</li>
          <li>Moved the PythonList UDF to a named lambda function</li>
          <li>Removed the PythonDict UDF</li>
          <li>Updated, reorganized tests to be more robust and concise. They now create/test the examples included in the documentation and were expanded to include the new functionality and more of the potential error paths</li>
          <li>Added CLI/xleda for Non-Developers sections on readme</li>
        </ul>
      </details><br>
      </td>
    </tr>
  </tbody>
</table><br>
