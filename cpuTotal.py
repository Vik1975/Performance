import matplotlib.pyplot as plt

# Data
categories = ["IGVR 2D", "IGVR Spatial"]
cpu_total_max_sum = [1118.870835, 2198.383981]

# Create a figure and axis object
fig, ax = plt.subplots(figsize=(10, 8))

# Bar graph
ax.bar(categories[0], cpu_total_max_sum[0], color="blue")
ax.bar(categories[1], cpu_total_max_sum[1], color="red")
ax.set_title("CPU Total Mean for oculus.igvr")
ax.set_xlabel("Category")
ax.set_ylabel("CPU Total Mean")

# Add value on top of each bar
for i, v in enumerate(cpu_total_max_sum):
    ax.text(i, v + 20, str(round(v)), color="black", ha="center")

# Set y-axis limits to ensure text labels fit within plot
ax.set_ylim(0, max(cpu_total_max_sum) + 50)

# Show the plot
plt.tight_layout()
plt.show()
