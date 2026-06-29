from xleda import wb
import seaborn as sns

# -----------------------------------------
# Full export

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
