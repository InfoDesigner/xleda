from xleda import FieldAnalysis
import seaborn as sns


# < your dataframe goes here > 
df = sns.load_dataset("titanic")


# Configure xleda
xleda = FieldAnalysis(input_df=df, 
                      name="Titanic",
                      overwrite=True)

# Create workbook
wb = xleda.create_workbook()


# Export your analysis back into Python
export_dict = xleda.export_analysis()