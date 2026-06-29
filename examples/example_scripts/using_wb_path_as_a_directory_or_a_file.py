from xleda import wb
from pathlib import Path
import seaborn as sns


df = sns.load_dataset('penguins')


# # Creates "c:\my_target_folder\Penguins.xlsm"
wb(data={"Penguins": df},
   wb_path=Path(r"c:\my_target_folder"))


# Creates "c:\my_target_folder\my_awesome_workbook.xlsx"
wb(data={"Penguins": df},
   wb_path=r"c:\my_target_folder\my_awesome_workbook.xlsx")