from xleda import wb
import seaborn as sns

# < your dataframe goes here >
df = sns.load_dataset("titanic")
  

# Creates "Titanic.xlsx" and returns basic metadata
export_dicts = wb(data=df,name="Titanic", overwrite=True, no_vba=True).export_dicts


# .....After editing your workbook....

# # Performs a full export from a completed xleda workbook named "Titanic (Completed).xlsm" 
export_dict = wb(data=df,
                 name="Titanic Completed",
                 export=True).export_dicts

print(export_dict[0].keys()) # Returns dict_keys(['description', 'definitions', 'notes', 'lists', 'field_metadata', 'overview_metadata', 'source_data', 'altered_source_data'])