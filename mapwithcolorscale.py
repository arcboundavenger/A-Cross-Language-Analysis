import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import geopandas as gpd
import seaborn as sns
from sklearn.preprocessing import MinMaxScaler

# 读取xlsx文件
file_path = 'pvalue_results_2.xlsx'  # 替换为你的文件路径
data = pd.read_excel(file_path, sheet_name='Sheet2')

# Create a DataFrame
table_data = pd.DataFrame(data)

# Set up the colormap
cmap = sns.color_palette("RdYlBu", as_cmap=True)

# Load the low resolution world map
world = gpd.read_file(gpd.datasets.get_path('naturalearth_lowres'))

# Merge the table data with the world GeoDataFrame based on region names
merged_data = world.merge(table_data, left_on='name', right_on='Region', how='left')

# Plot the map with colors and colormap
fig, ax = plt.subplots(figsize=(10, 6))
# Fill all regions with a default color
world.plot(ax=ax, color='white', edgecolor='0.8', linewidth=0.5)

# Now plot the merged data to overlay your data
merged_data.plot(column='Ratings', cmap=cmap, linewidth=0.8, edgecolor='0.8', legend=True, ax=ax)

# Remove axis ticks and labels
ax.set_axis_off()

# Show the plot
plt.show()