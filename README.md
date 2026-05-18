# **xleda is a Microsoft Excel powered EDA tool for Python data.**

* Produces a Microsoft Excel workbook from a pandas dataframe that is highly optimized to both perform and document [the activity of Exploratory Data Analysis](https://www.geeksforgeeks.org/data-analysis/what-is-exploratory-data-analysis/) .

* Visually explore your data, navigate with your keyboard, take field or record notes, create lists of fields/records for editing, round-trip your edits/analysis back into python, share your workbook with other contributors.

* There are some amazing EDA tools for Python.  You shouldn't have to start from scratch to include Microsoft Excel among them.

* xleda provides a good start to a robust EDA.  

* See [some example xleda workbooks](https://github.com/InfoDesigner/xleda/tree/main/examples).<br>

<br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/top_view.gif?raw=true"  width="700" alt="Example Top View"> 
	<br>
	<em>An xleda workbook made with Titanic passenger data.</em>
</p>
<br><br>


# **Requirements/Compatibility**

* Requires Microsoft Excel to create the workbook. 

* It has been tested on Windows though it should also work on MacOS.  

* xleda workbooks should work in anything that reads Microsoft Excel workbooks.<br><br>


# **Quick Start**


```python
from xleda import FieldAnalysis
import seaborn as sns

# < your dataframe goes here >
df = sns.load_dataset("titanic")
 

# Configure xleda
xleda = FieldAnalysis(input_df=df,
                      name="Titanic")

# Create workbook
wb = xleda.create_workbook()

```

<br><br>

# **Usage Notes**

<br>

## **Included Metadata**

* Most of the included field metadata is from the built-in pandas features [describe](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.describe.html), [info](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.info.html), and **[quantile](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.quantile.html)**. <br><br>

## **Theme Color**

* `theme_color` sets the primary color of the charts and the color of the headings in the workbook to a hex color of your choice.  `theme_color=random` sets a random theme<br>

<br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/theme_colors.webp?raw=true" width="800" alt="Theme Colors">
	<br>
	<em>theme_color affects the workbooks and default charts.</em>
</p>
<br><br>

## **Add Additional Plots**

* `xleda.add_plot()` will add additional worksheets with a plot of your choosing. 

	* No styling/sizing of additional plots is performed.
	
	* The example below adds two additional plot worksheets, one from seaborn and another from missingno.  The workbook can be found [here](https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/Penguins.xlsm).


<br>


```python
from xleda import FieldAnalysis

import matplotlib.pyplot as plt
import seaborn as sns
import missingno as msno
  

# Import example data
df = penguins = sns.load_dataset("penguins")
  

# Style the additional plots (optional)
plt.style.use("dark_background")


# Create additional plots
pair_plots = sns.pairplot(df, hue="species").figure
null_matrix = msno.matrix(df).get_figure()


# Resize the null matrix for good measure
null_matrix.set_size_inches(9.35, 4.5)


# Configures an xleda workbook
xleda = FieldAnalysis(input_df=df,
                      name="Penguins",
                      theme_color="#4C4C4C",
                      add_plots={'Pair Plots': pair_plots,
                                 'Null Matrix': null_matrix})

# Creates the workbook
xleda.create_workbook()

```

<br>



## **Field/Record Lists**


* The `Field Lists` section helps you create lists of the fields in your data.  

	* Anything not marked as `False` will be included in each list.   

	* The `Record List` field added to your source data works the same way except it creates a list of records instead of a list of fields.  More on that [below](#**Record%20List%20Details**)

	* You can see your lists in the `Compiled Lists` section.  

	* You can rename `Record List` or any `Field List` to `Anything You Want` and the list will be renamed to `anything_you_want`.

	* The Compiled Lists section formats your lists as python lists for easy copy/pasting.

	* You can use `export_analysis()` to get your lists, and other things, back into Python.  See details [below](##%20**Exporting%20back%20into%20Python**).

	* Altering the Field Analysis worksheet may offset the formulas which compile your lists.   Spot check them before using them if you have.<br>

<br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/field_lists.gif?raw=true" width="606" alt="Field Lists">
	<br>
	<em>Easily create lists of fields in your data.</em>
</p>
<br><br>



## **Record List Details**

* Two additional columns are added to your data to support being able to create a list of records for further processing.

	* `Record Hash`:  Uses a built-in pandas feature [hash_pandas_object](https://pandas.pydata.org/docs/reference/api/pandas.util.hash_pandas_object.html) to uniquely identify records.  If two records share all column values they also share a `Record Hash`. 

	* `Record List`:  Used to create a list of `Record Hash` values.<br><br>

## **Exporting back into Python**

* At some point, you'll likely need to get some of your analysis back into Python.  `export_analysis()` exports your notes, definitions, lists (all pictured below) and more back into Python.<br>

<br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/completed_field_analysis.webp?raw=true" width="606" alt="Completed Field Analysis">
	<br>
	<em>A completed xleda workbook of Titanic passenger showing definitions, notes, lists, etc.</em>
</p>
<br><br>


* All exported data comes from the Field Analysis worksheet.

* It is assumed you haven't altered the structure of your workbook such as adding rows/columns. 

* The dictionary includes:

	* `description`: Dataframe description if you've added one
	* `definitions`: Any field definitions you've added.
	* `notes`: Any field notes you've added
	* `lists`: Any lists showing in the compiled lists section
	* `source_data`: A copy of your unaltered source data that includes `Record Hash`/`Record List` columns.
	* `altered_source_data`: source data from the workbook that includes any manual edits you've made such as removing records, renaming fields, etc. 

		 ** *Note that data types will likely change in the round-trip translation.* **<br>

<br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/completed_analysis_export.webp?raw=true" width="500" alt="Export Dict">
	<br>
	<em>An example export from a completed field analysis on Titanic passenger data.</em>
</p>
<br><br>

## **Limits with Large Data Sets**

* xleda creates workbooks for most data sets less than 20 seconds.  To ensure that they are created quickly, defaults limit data to the first 100 columns and a random sample of 100,000 records.  You'll see a warning if you hit a limit.

<br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/create_example.webp?raw=true" width="400" alt="Create Example">
	<br>
	<em>The Penguin example from above that includes extra plots took only 7 seconds to create.</em>
</p>
<br><br>

* `large_report=True` raises the limits to Excel's limits of 1,000,000 rows and 16,000 columns.  The closer your are to this limit, the longer it will take to produce. 

* One of the larger/more complex data sets tested was a 600 column/1,200 row dataframe.  
	* It took ~12 minutes to create, in part because most values are unique for all 600 columns and xleda give you a top 5 members composition chart per column.
	* It is still snappy to use even though it has 1,200 charts on a single worksheet.  That example is [here](https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/African%20Soil.xlsm).<br><br>


## **VBA Code**

* There is a small amount of VBA code in the template that makes the sections expand/collapse when you select them as pictured above.  If you can't or don't want to enable VBA, use row groupings as pictured below. <br>

<br>
<p align="center">
	<img src="https://github.com/InfoDesigner/xleda/blob/main/assets/images/row_groupings.gif?raw=true" width="462" alt="Row Groupings">
	<br>
	<em>Use row groupings to navigate if you can't use VBA.</em>
</p>
<br><br>

## **Extensible**

* Because it's an ordinary workbook, you can use any tool that works with Microsoft Excel workbooks to do more.  [xlwings](https://www.xlwings.org/) is recommended if you need more. 

