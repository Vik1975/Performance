import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Prompt user to enter file names
file_name1 = input("Enter the name of the first Excel file: ")
file_name2 = input("Enter the name of the second Excel file: ")

# Define the worksheet name
worksheet_name = "Summary for avg"

# Define the cell values to extract
cell_values = [
    "CPU Total by Process (Avg)",
    "Android Process",
    "Debug Tools",
    "Mixed Reality",
    "VR Shell",
]

# Read the Excel files
df1 = pd.read_excel(file_name1, sheet_name=worksheet_name, header=None)
df2 = pd.read_excel(file_name2, sheet_name=worksheet_name, header=None)

# Extract the cell values
values1 = [
    df1.iloc[1, 6],
    df1.iloc[2, 6],
    df1.iloc[6, 6],
    df1.iloc[13, 6],
    df1.iloc[27, 6],
]
values2 = [
    df2.iloc[1, 6],
    df2.iloc[2, 6],
    df2.iloc[6, 6],
    df2.iloc[13, 6],
    df2.iloc[27, 6],
]

# Set the bar width
bar_width = 0.4

# Set the x-coordinates of the bars
x = np.arange(len(cell_values))
x1 = x - bar_width / 2
x2 = x + bar_width / 2

# Create a figure with two subplots
fig, axs = plt.subplots(1, 2, figsize=(16, 6))

# Plot the first graph
axs[0].bar(x1, values1, width=bar_width, color="red", label="Spatial IG app results")
axs[0].bar(x2, values2, width=bar_width, color="blue", label="2D IG app results")
axs[0].set_xlabel("Processes")
axs[0].set_ylabel("CPU")
axs[0].set_title("Spatial IG vs 2D IG CPU usage")
axs[0].set_xticks(x)
axs[0].set_xticklabels(cell_values, rotation=45)
axs[0].legend()

# Plot the second graph
g46_values1 = df1.iloc[45, 6]
g46_values2 = df2.iloc[45, 6]
axs[1].bar([1, 2], [g46_values1, g46_values2], color=["red", "blue"])
axs[1].set_xlabel("GPU Utilization")
axs[1].set_ylabel("%")
axs[1].set_title("GPU Utilization Spatial IG vs 2D IG")
axs[1].set_xticks([1, 2])
axs[1].set_xticklabels(["Spatial IG", "2D IG"])

# Layout so plots do not overlap
fig.tight_layout()

# Show the plot
plt.show()
