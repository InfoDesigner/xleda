from xleda import wb
import seaborn as sns



# < your dataframe goes here > 
df = sns.load_dataset("titanic")


# Creates an xleda workbook
wb(df, theme_color='random', overwrite=True)