import matplotlib.pyplot as plt

# Data for Spatial Instagram APK
spatial_android = [
    1383.03,
    1570,
    1900,
]
spatial_vr_shell = [2092.55, 2191, 1926]
spatial_debug_tools = [837.77, 897, 847]
spatial_CPU_total = [8770, 9161, 9566]
spatial_mixed_reality = [469, 461, 433]

# Data for 2D Instagram APK
_2d_android = [1330.72, 1415, 1395]
_2d_vr_shell = [2173.57, 2236, 2268]
_2d_debug_tools = [818.41, 808, 857]
_2d_CPU_total = [8338, 8965, 8612]
_2d_mixed_reality = [483, 468, 486]

# Create a figure and axis object
fig, ax = plt.subplots(1, 2, figsize=(12, 6))

# Bar graph for Spatial Instagram APK
ax[0].bar(
    ["Android", "VR Shell", "Debug Tools", "CPU Average", "Mixed Reality"],
    [
        sum(spatial_android) / len(spatial_android),
        sum(spatial_vr_shell) / len(spatial_vr_shell),
        sum(spatial_debug_tools) / len(spatial_debug_tools),
        sum(spatial_CPU_total) / len(spatial_CPU_total),
        sum(spatial_mixed_reality) / len(spatial_mixed_reality),
    ],
)
for i, v in enumerate(
    [
        sum(spatial_android) / len(spatial_android),
        sum(spatial_vr_shell) / len(spatial_vr_shell),
        sum(spatial_debug_tools) / len(spatial_debug_tools),
        sum(spatial_CPU_total) / len(spatial_CPU_total),
        sum(spatial_mixed_reality) / len(spatial_mixed_reality),
    ]
):
    ax[0].text(i, v + 10, str(round(v, 2)), color="black", ha="center")
ax[0].set_title("Spatial Instagram APK")
ax[0].set_xlabel("Process")
ax[0].set_ylabel("CPU Usage")
ax[0].tick_params(axis="x", rotation=45)

# Bar graph for 2D Instagram APK
ax[1].bar(
    ["Android", "VR Shell", "Debug Tools", "CPU Average", "Mixed Reality"],
    [
        sum(_2d_android) / len(_2d_android),
        sum(_2d_vr_shell) / len(_2d_vr_shell),
        sum(_2d_debug_tools) / len(_2d_debug_tools),
        sum(_2d_CPU_total) / len(_2d_CPU_total),
        sum(_2d_mixed_reality) / len(_2d_mixed_reality),
    ],
)
for i, v in enumerate(
    [
        sum(_2d_android) / len(_2d_android),
        sum(_2d_vr_shell) / len(_2d_vr_shell),
        sum(_2d_debug_tools) / len(_2d_debug_tools),
        sum(_2d_CPU_total) / len(_2d_CPU_total),
        sum(_2d_mixed_reality) / len(_2d_mixed_reality),
    ]
):
    ax[1].text(i, v + 10, str(round(v, 2)), color="black", ha="center")
ax[1].set_title("2D Instagram APK")
ax[1].set_xlabel("Process")
ax[1].set_ylabel("CPU Usage")
ax[1].tick_params(axis="x", rotation=45)

# Show the plot
plt.tight_layout()
plt.show()
