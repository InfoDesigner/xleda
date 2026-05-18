from xleda import FieldAnalysis

import matplotlib.pyplot as plt
import seaborn as sns
import missingno as msno


# Import example data 
df = penguins = sns.load_dataset("penguins")


# Style the additional plots (optional)
plt.style.use("dark_background")


# Create additional plots
pair_plots = sns.pairplot(df, hue="species").figure
null_matrix = msno.matrix(df).get_figure() # type: ignore


# Resize the null matrix for good measure
null_matrix.set_size_inches(9.35, 4.5) # type: ignore


# Configures an xleda workbook
xleda = FieldAnalysis(input_df=df, 
                      name="Penguins", 
                      theme_color="#4C4C4C",
                      overwrite=True,
                      add_plots={'Pair Plots': pair_plots,
                                #  'Null Matrix': null_matrix,   # type: ignore 
                                 }) 


# Creates the workbook
xleda.create_workbook()