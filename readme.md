



<div align="center">
  <a href="https://www.apache.org/licenses/LICENSE-2.0.txt"><img src="https://img.shields.io/badge/license-Apache-**blue**"></a>
  <a href="https://pypi.org/project/xleda"><img src="https://img.shields.io/pypi/v/xleda"></a>
  <a href="https://pypi.org/project/xleda"><img src="https://img.shields.io/pypi/pyversions/xleda.svg"></a>
  <a href="https://pepy.tech/project/xleda"><img src="https://static.pepy.tech/badge/xleda"></a>
  <a href="https://github.com/InfoDesigner/xleda"><img src="https://img.shields.io/badge/Made%20By%20A%20Human-99%25-blue"></a>
  <a href="https://buymeacoffee.com/informationdesigner"><img src="https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ffdd00?&logo=buy-me-a-coffee&logoColor=black"></a>
</div>
<br>


<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/docs/assets/images/logo.webp?raw=true" width="250" alt="Logo">
	<br>
</p>

xleda is a Python/Excel powered EDA tool that creates workbooks from dataframes or data files that are highly optimized to explore, define, and document data sets.<br><br>

* Works on Windows or MacOS as a Python package, a CLI, or as a service that lets you create workbooks by right-clicking supported files.<br><br>

* There are some amazing EDA tools available to data professionals. You shouldn't have to start from scratch to include Microsoft Excel among them.<br><br><br>

<p align="center"><a href="https://infodesigner.github.io/xleda">Docs</a> | 
<a href="https://github.com/InfoDesigner/xleda/tree/main/examples">Sample Workbooks</a> | 
<a href="https://infodesigner.github.io/xleda/#basic-usage">Basic Usage</a> | 
<a href="https://infodesigner.github.io/xleda/#xledawb-configuration">Configuration</a> | 
<a href="https://infodesigner.github.io/xleda/#usage-notes">Usage Notes</a> | 
<a href="https://infodesigner.github.io/xleda/#examples">Examples</a> | 
<a href="https://infodesigner.github.io/xleda/#xleda-for-non-developers">Quick-Start Guide for Non-Developers</a><br></p><br><br>



<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/docs/assets/images/top_view.webp?raw=true"  width="800" alt="Example Top View"> 
	<br>
	<em>Top view of a Field Analysis worksheet.</em>
</p><br><br>

<hr>

## xleda Components

<br>

<div class="doc-grid-container" markdown="1">

#### Field Analysis

  <div class="section-content" markdown="1">

<details>
  <summary>Anatomy of a Field Analysis Worksheet<br><br></summary>
  <p align="left">
    <img src="https://github.com/InfoDesigner/xleda/blob/main/docs/assets/images/field_analysis_anatomy.webp?raw=true" style="width: 100%; max-width: 100%;" alt="Field Analysis Anatomy">
  </p>
  <em>Field Analysis Anatomy</em>
</details>

  </div>
</div>


<div class="doc-grid-container" markdown="1">

#### Overview

  <div class="section-content" markdown="1">
  
<details>
  <summary>Anatomy of an Overview Worksheet<br><br></summary>
  <p align="left">
    <img src="https://github.com/InfoDesigner/xleda/blob/main/docs/assets/images/overview_anatomy.webp?raw=true" style="width: 100%; max-width: 100%;" alt="Overview Anatomy">
  </p>
  <em>Overview Worksheet with Multiple Dataframes</em>
</details>

  </div>
</div><br><br>


<hr>

## Basic Usage

<br>

Use <code>wb()</code> to quickly create an xleda workbook from a dataframe, a dictionary of dataframes, or a supported data file.<br><br>



**From a Dataframe**

  ``` python
  from xleda import wb
  import seaborn as sns

  # < your dataframe goes here >
  df = sns.load_dataset("titanic")

  # Creates xleda.xlsm in the current directory
  wb(df)
  ```

<hr><br>

**From a Dictionary of Dataframes**

  ``` python
  from xleda import wb
  import seaborn as sns

  # < your dataframes go here >
  df1 = sns.load_dataset("titanic")
  df2 = sns.load_dataset("penguins")

  # Creates Titanic.xlsm in the current directory
  wb({"Titanic": df1,
      "Penguins": df2})
  ```

<hr><br>

**From a File**

  ```python
  from xleda import wb
  from pathlib import Path

  # < your data file goes here >
  duckdb_file = "https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/data/duckdb.duckdb"

  # Creates duckdb.xlsm in the current directory
  # Includes data from all tables in the db file
  wb(duckdb_file)
  ```

<hr><br>

**From the CLI**

  ```bash
  # Creates 'userdata.xlsm' in the current directory
  xleda wb 'https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/data/userdata.parquet'

  # Shows the help command
  xleda --help

  # Shows the wb help command
  xleda wb --help
  ```

<hr><br>

**From Right-Clicking | After running `xleda install`**

  <table>
    <tbody>
      <tr>
        <td width="50%" valign="top"><br><strong>Windows</strong><br><br></td>
        <td width="50%" valign="top"><br><strong>MacOS</strong><br><br></td>
      </tr>
      <tr>
        <td valign="top"><br>
          <p align="left">
            <img src="https://github.com/InfoDesigner/xleda/blob/main/docs/assets/images/right_click_win.webp?raw=true" width="400" alt="From right-click Win">
          </p><br>
        </td>
        <td valign="top"><br>
          <p align="left">
            <img src="https://github.com/InfoDesigner/xleda/blob/main/docs/assets/images/right_click_mac.webp?raw=true" width="400" alt="From right-click MacOS">
          </p><br>
        </td>
      </tr>
    </tbody>
  </table><br><br>

<hr>

## Compatibility

<br>


<table>
  <tbody>
    <tr>
      <td width="20%" valign="top"><strong>Desktop&nbsp;Excel</strong></td>
      <td valign="top">
      <p>Requires the full version of Microsoft Excel (2016+) on either MacOS or Windows to create workbooks</p>
        <ul>
          <li>See the <a href="https://infodesigner.github.io/xleda/#macos-support">MacOS Support</a> section of the docs for details on MacOS usage.</li><p></p>
        </ul>
      </td>
    </tr>
    <tr>
      <td valign="top"><strong>Supported&nbsp;Data</strong></td>
      <td valign="top">
      <p>Supports pandas dataframes, CSV, DuckDB, SQLite, Feather, Parquet, Pickle, Excel, RData, JSON, and XML</p><p></p><br>
      </td>
    </tr>
  </tbody>
</table><br>


