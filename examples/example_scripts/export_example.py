from xleda import wb
import seaborn as sns

# < your dataframe goes here >
df = sns.load_dataset("titanic")
  

# Creates "Titanic.xlsx" and returns basic metadata
export_dict = wb(input_df=df,name="Titanic", overwrite=True, no_vba=True).export_dict


# .....After editing your workbook....

# # Performs a full export from a completed xleda workbook named "Titanic (Completed).xlsx" 
export_dict = wb(input_df=df,
                 name="Titanic (Completed)",
                 no_vba=True,
                 export=True).export_dict

print(export_dict.keys()) # Returns dict_keys(['description', 'definitions', 'notes', 'lists', 'field_metadata', 'overview_metadata', 'source_data', 'altered_source_data'])