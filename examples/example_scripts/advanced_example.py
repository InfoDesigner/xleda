import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import missingno as msno

from xleda import wb

# Import example data 
og_penguins = pd.read_csv("https://raw.githubusercontent.com/allisonhorst/palmerpenguins/refs/heads/main/inst/extdata/penguins_raw.csv")
penguins = sns.load_dataset('penguins')
seaice = sns.load_dataset('seaice')

# Style the additional plots (optional)
plt.style.use("dark_background")


# Create additional plots
pair_plots = sns.pairplot(penguins, hue="species").figure
null_matrix = msno.matrix(penguins).get_figure()


# Resize the null matrix for good measure
null_matrix.set_size_inches(9.35, 4.5) # type: ignore


# Creates "OG Penguins.xlsm" in the current directory that includes:
#    * Two additonal plot worksheets
#    * xleda analyses for two additonal dataframes.

wb(input_df=penguins, 
   name="Penguins", 
   theme_color="#817CA2",
   overwrite=True,
   add_plots={'Pair Plots': pair_plots,
              'Null Matrix': null_matrix}, # type: ignore
   add_dfs={'Sea Ice': seaice,
            'OG Penguins': og_penguins})