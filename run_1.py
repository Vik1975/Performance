import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Prompt user to enter file names
file_name1 = input("Enter the name of the Excel file with 2D results: ")
file_name2 = input("Enter the name of the Excel file with Spatial results: ")

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

# Create a figure with three subplots
fig, axs = plt.subplots(1, 3, figsize=(20, 6), gridspec_kw={"width_ratios": [3, 1, 1]})

# Plot the first graph
axs[0].bar(x1, values1, width=bar_width, color="blue", label="2D app results")
axs[0].bar(x2, values2, width=bar_width, color="red", label="Spatial app results")
axs[0].set_xlabel("Processes")
axs[0].set_ylabel("CPU")
axs[0].set_title("2D vs Spatial CPU usage")
axs[0].set_xticks(x)
axs[0].set_xticklabels(cell_values, rotation=45)
axs[0].legend()
for i, rect in enumerate(axs[0].patches):
    height = rect.get_height()
    axs[0].text(
        rect.get_x() + rect.get_width() / 2,
        height,
        f"{height:.2f}",
        ha="center",
        va="bottom",
    )

# Plot the second graph
g46_values1 = df1.iloc[45, 2]
g46_values2 = df2.iloc[45, 2]
axs[1].bar([1, 2], [g46_values1 * 100, g46_values2 * 100], color=["blue", "red"])
axs[1].set_xlabel("GPU Utilization")
axs[1].set_ylabel("Value (%)")
axs[1].set_title("GPU Utilization 2D vs Spatial")
axs[1].set_xticks([1, 2])
axs[1].set_xticklabels(["2D", "Spatial"])
axs[1].yaxis.set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:.0f}%".format(x)))
for i, rect in enumerate(axs[1].patches):
    height = rect.get_height()
    axs[1].text(
        rect.get_x() + rect.get_width() / 2,
        height,
        f"{height:.0f}%",
        ha="center",
        va="bottom",
    )

# Plot the third graph
g48_values1 = df1.iloc[47, 6]
g48_values2 = df2.iloc[47, 6]
axs[2].bar([1, 2], [g48_values1, g48_values2], color=["blue", "red"])
axs[2].set_xlabel("FPS")
axs[2].set_ylabel("Value")
axs[2].set_title("FPS 2D vs Spatial")
axs[2].set_xticks([1, 2])
axs[2].set_xticklabels(["2D", "Spatial"])
for i, rect in enumerate(axs[2].patches):
    height = rect.get_height()
    axs[2].text(
        rect.get_x() + rect.get_width() / 2,
        height,
        f"{height:.2f}",
        ha="center",
        va="bottom",
    )

# Layout so plots do not overlap
fig.tight_layout()

# Show the plot
plt.show()
