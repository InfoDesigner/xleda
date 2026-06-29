import seaborn as sns
from xleda import wb

seaborn_datasets = ['diamonds', 'dots', 'dowjones']
dataframe_dict = {df_name: sns.load_dataset(df_name) for df_name in seaborn_datasets}

# Creates diamonds.xlsm in the current directory
# Also includes dots and dow jones data
wb(data=dataframe_dict, overwrite=True)