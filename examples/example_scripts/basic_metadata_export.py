from xleda import wb
import seaborn as sns


# < your dataframe goes here >
df = sns.load_dataset("titanic")
  
# -----------------------------------------
# Basic Export

# Creates "Titanic.xlsm" and returns basic metadata
export_dicts = wb(data={"Titanic": df},
                  file_name="Titanic").export_dicts


# returns ['field_overview', 'df_overview', 'source_data']
print(export_dicts[0].keys())