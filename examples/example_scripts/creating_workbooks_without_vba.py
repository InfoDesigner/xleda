from xleda import wb
import seaborn as sns

df = sns.load_dataset('penguins')

# Creates "Penguins.xlsx" in the current directory and changes the default workbook style to .xlsx
wb(data={"Penguins": df},
   vba=False)

# Also creates "Penguins.xlsx" but doesn't change the default workbook style 
wb(data={"Penguins": df},
   wb_path="Penguins.xlsx")