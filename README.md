# **xleda is a Microsoft Excel powered EDA tool for Python data.**

* Easily produces a Microsoft Excel workbook from a pandas dataframe that is highly optimized to both perform and document [the activity of Exploratory Data Analysis](https://www.geeksforgeeks.org/data-analysis/what-is-exploratory-data-analysis/) .

* Microsoft Excel is great for EDA:

	* Visually explore your data, navigate with your keyboard, take notes, mark fields/records for editing, share your workbook with other contributors.

* See [some example xleda workbooks](examples).


<center>
	<figure> 
		<img src="docs/img/top_view.gif" width="700" alt="Example Top View"> 
		<figcaption>An xleda workbook made with with Titanic passenger data.</figcaption> 
	</figure>
</center>

# **xleda is Not a Data Transformation Tool**

* This is not intended to edit your data.  It is simply a way to quickly get Python data into Microsoft Excel so that you can see, document, and understand it.

# **Requirements/Compatibility**

* Requires Microsoft Excel.  Should work for most MS Office SKUs.

* It has been tested on Windows though it should also work on MacOS.


# **Quick Start**

```python
import seaborn as sns
from xleda import FieldAnalysis
  

# < Insert your data here >
df = sns.load_dataset("titanic")  

# Configures an xleda workbook
xleda_wb = FieldAnalysis(input_df=df,
                         name="Titanic",
                         theme_color="#053476",
                         overwrite=True)

# Creates the workbook
xleda_wb.create_workbook()
  

# Imports your analysis back into Python
export_dict = xleda_wb.export_analysis()

```


# **Usage Notes**

## **Included Metadata**

* Most of the included field metadata is from the built-in pandas features [describe](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.describe.html), [info](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.info.html), and **[quantile](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.quantile.html)**. 

## **Theme Color**

* `theme_color` sets the primary color of the charts and the color of the headings in the workbook.  

* Because this tool creates workbooks in dark mode, choose a dark color though any hex color will technically work.

## **Field/Record Lists**

* The `Field Lists` helps you create lists of the fields in your data.  

	* Anything not marked as `False` will be included in each list.   

	* The `Record List` field added to your source data works the same way except it creates a list of records instead of a list of fields.  More on that field [below](#**Record%20List%20Details**)

	* You can see your lists in the `Field Actions Lists` section or you can use `export_analysis()` to get them into Python.

	* You can rename `Record List` or any `FieldList` to `Anything You Want` and the list will be renamed to `anything_you_want`.

	* The Compiled Lists section formats your lists as python lists for easy copy/pasting.

	* Caution: Altering the Field Analysis worksheet may offset some of the few formulas in this workbook which compile your lists.   Spot check them before using them if you have.


<center>
	<figure> 
		<img src="docs/img/field_lists.gif" width="606" alt="Field Actions"> 
		<figcaption>Easily create lists of fields in your data.</figcaption> 
	</figure>
</center>

<center>
	<figure> 
		<img src="docs/img/completed_field_analysis.webp" width="606" alt="Field Actions"> 
		<figcaption>Example of a completed field analysis of Titanic passenger data.</figcaption> 
	</figure>
</center>

## **Record List Details**

* Two additional columns are added to your data to support being able to create a list of records for further processing.

	* `Record Hash`:  Uses a built-in pandas feature [hash_pandas_object](https://pandas.pydata.org/docs/reference/api/pandas.util.hash_pandas_object.html) to uniquely identify records.  If two records share all column values they also share a `Record Hash`. 

	* `Record List`:  Used to create a list of `Record Hash` values.

## **Exporting back into Python**

*  `xleda_wb.export_analysis()` exports the notes/lists/data from your workbook into Python.  
 
* This function assumes you haven't altered the structure of your workbook such as adding rows/columns.  It is provided as a convenience and is otherwise unsupported.  

* The dictionary includes:

	* `lists`: Any lists showing in the compiled lists section
	* `notes`: Any field notes you've added
	* `source_data`: A copy of your unaltered source data that includes `Record Hash`/`Record List` columns.
	* `altered_source_data`: source data from the workbook that includes any manual edits you've made such as removing records, renaming fields, etc.  Note that data types will likely change in the round-trip translation.


<center>
	<figure> 
		<img src="docs/img/completed_analysis_export.webp" width="500" alt="Export Dict"> 
		<figcaption>An example export from a completed field analysis on Titanic passenger data. 
		</figcaption> 
	</figure>
</center>


## **Limits with Large Data Sets**

* This creates EDA workbooks for most data sets in 10-20 seconds.

<center>
	<figure> 
		<img src="docs/img/create_example.webp" width="400" alt="Export Dict"> 
		<figcaption>The Air Quality example took only 7 seconds to create. 
		</figcaption> 
	</figure>
</center>

* To ensure that they are created quickly, defaults limit data to the first 100 columns and a random sample of 100,000 records.  You'll see a warning if you hit a limit.

* `large_report=True` raises the limits to Excel's limits of 1,000,000 rows and 16,000 columns.  The closer your are to this limit, the longer it will take to produce. 

* One of the largest data sets tested was a 600 column/1,200 row dataframe.  It took ~12 minutes to create but is still snappy to use even though it has 1,200 charts on a single worksheet.  That example is [here]("examples\African Soil.xlsm").


## **VBA Code**

* There is a small amount of VBA code in the template that makes the sections expand/collapse when you select them as pictured above.  If you can't or don't want to enable VBA, use row groupings as pictured below. 

<center>
	<figure> 
		<img src="docs/img/row_groupings.gif" width="462" alt="Row Groupings"> 
		<figcaption>Use row groupings to navigate if you can't use VBA.</figcaption> 
	</figure>
</center>

