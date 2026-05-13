# **XLEDA is a Microsoft Excel powered EDA tool for Python data.**

* Easily produces a Microsoft Excel workbook from a pandas dataframe that is highly optimized to both perform and document [the activity of Exploratory Data Analysis](https://www.geeksforgeeks.org/data-analysis/what-is-exploratory-data-analysis/) .

* Microsoft Excel is great for EDA:
	* Visually explore your data, navigate with your keyboard, take notes, mark fields/records for editing, share your workbook with other contributors.


# **XLEDA is Not a Data Transformation Tool**

* This will not edit your data.  

* It will help you quickly get your data into Excel so that you can understand and document your data.

# **Requirements/Compatibility**

* This requires Microsoft Excel though it should work for most MS Office SKUs.

* This is powered by the amazing duo of [pandas](https://pandas.pydata.org/) and [xlwings](https://www.xlwings.org/) so it should work for most on MacOS and Windows. 

* There is a small amount of VBA code in the template that isn't actually required.  It makes the sections expand/collapse when you select them, similar to web-based tools.  Use row groupings if you can't or don't want to enable VBA.  


# **Quick Start**

```python

from xleda import create_workbook
import seaborn as sns

# <Insert your df here>
df = sns.load_dataset('titanic')

# Creates a Field Analysis workbook named Titanic
create_workbook(input_df=df, name="Titanic", color='#339803')

```


# **Usage Notes**

## **Included Metadata**

* Most of the included field metadata is from the built-in pandas features [describe](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.describe.html), [info](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.info.html), and **[quantile](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.quantile.html)**. 

## **Theme Color**

* `theme_color` sets the primary color of the charts and the color of the headings in the workbook.  

* Because this tool creates workbooks in dark mode, choose a dark color though any hex color will technically work.

## **Mark for Removal/Record Hashes**

* Although your data isn't edited by this tool, there are two additional columns added to support being able to create a list of records for further processing 

	* Record Hash:  Used to uniquely identify records.  This column is using a [built-in pandas feature](https://pandas.pydata.org/docs/reference/api/pandas.util.hash_pandas_object.html).  If two records share all column values they also share a Record Hash. 

	* Mark For Removal:  Used to create a list of Record Hash values.


## **Field Actions/Lists**

* The Field Actions section is really a way to easily create lists of the fields/records in your data.  The action part is something you're responsible for. 

	* Anything not marked as False will be included in each list.   

	* The Mark For Removal field in the Source Data table works the same way as Field Actions except it creates a list of records instead of a list of fields.

	* You can rename ***Mark For Removal*** or any ***Field Action*** to **Anything You Want** and the list will be renamed to ***anything_you_want***. 

	* You can see your lists in the ***Field Actions Lists*** section.


## **Large Data Sets**

* To ensure that reports are created quickly, defaults limit data to 100 first columns and a random sample of 100,000 records.  

* `large_report=True` raises the limits to Excel's limits of 1,000,000 rows and 16,000 columns.  The closer your are to this limit, the longer it will take to produce. 

* This was tested with a lot of different datasets on an average machine.  For example, a 600 column, 1,200 row df took 10+ minutes to create but was snappy to use once it was created even though it had 1,200 charts on a single worksheet.

* You'll see a warning if you hit a limit.