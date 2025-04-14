import matplotlib.pyplot as plt
import pandas as pd

# Define the file names and worksheet name
file_names = ["Stinson - Scenario 207.xlsx", "Stinson - Scenario 208.xlsx"]
worksheet_name = "Summary for avg"

# Define the cell values to extract
cell_values = [
    "CPU Total by Process (Avg)",
    "Android Process",
    "Debug Tools",
    "Mixed Reality",
    "VR Shell",
]

# Create a figure and axis object
fig, ax = plt.subplots(1, 2, figsize=(14, 6))

# Loop through each file
for i, file_name in enumerate(file_names):
    # Read the Excel file
    df = pd.read_excel(file_name, sheet_name=worksheet_name, header=None)

    # Extract the cell values
    values = [
        df.iloc[1, 6],
        df.iloc[2, 6],
        df.iloc[6, 6],
        df.iloc[13, 6],
        df.iloc[27, 6],
    ]

    print(f"Scenario {i+1} values: {values}")

    # Create a bar graph
    bars = ax[i].bar(cell_values, values)
    if i == 0:
        ax[i].set_title("Spatial IG app results")
    elif i == 1:
        ax[i].set_title("2D IG app results")

    ax[i].set_xlabel("Processes")
    ax[i].set_ylabel("CPU")

    # Add value labels to the top of each bar
    for bar in bars:
        height = bar.get_height()
        ax[i].text(
            bar.get_x() + bar.get_width() / 2,
            height,
            f"{height:.2f}",
            ha="center",
            va="bottom",
        )

    # Set y-axis limits manually
    ax[i].set_ylim(0, max(values) * 1.1)

    # Rotate x-axis labels by 45 degrees
    ax[i].tick_params(axis="x", rotation=45)

# Layout so plots do not overlap
fig.tight_layout()

# Show the plot
plt.show()
