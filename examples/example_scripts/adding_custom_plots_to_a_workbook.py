from xleda import wb
import matplotlib.pyplot as plt
import seaborn as sns
import missingno as msno

# < your dataframe goes here >
df = penguins = sns.load_dataset("penguins")

# Style the additional plots | optional
plt.style.use("dark_background")

# Create additional plots
pair_plots = sns.pairplot(df, hue="species").figure
null_matrix = msno.matrix(df).get_figure()

# Resize the null matrix  | optional
null_matrix.set_size_inches(9.35, 4.5)

# Creates Penguins.xlsm with two extra plot sheets
wb(data={"Penguins": df},
   theme="#4335A0",
   plots={'Pair Plots': pair_plots,
          'Null Matrix': null_matrix})