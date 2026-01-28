import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# 1. File Path
# Use raw string (r"...") to handle Windows backslashes
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"

try:
    print("Loading and analyzing data...")
    df = pd.read_excel(file_path, sheet_name=0)

    # Clean column names (remove leading/trailing spaces)
    df.columns = df.columns.str.strip()

    # --- COLUMN DEFINITIONS ---
    tools = [
        "Optoelectronic", "Force-plate", "EMG", "Heart-rate-monitor",
        "Metabolic-cart","IMU",  "Wii-fit", "Other-tools"
    ]

    activities = [
        "Sit-to-stand", "Running", "Cycling", "Stair-negotiation", "Obstacle-clearance", "Game",
        "Jumping","Time-Up-and-Go", "One-leg-standing",
        "Stepping-target", "Hopping", "Squat","Kicking-a-ball"
    ]

    # Mapping GMFCS columns to simple labels
    gmfcs_map = {
        "GMFCS-I": "I",
        "GMFCS-II": "II",
        "GMFCS-III": "III",
        "GMFCS-IV": "IV"
    }

    # Verify which columns actually exist in the file
    existing_tools = [t for t in tools if t in df.columns]
    existing_activities = [a for a in activities if a in df.columns]
    existing_gmfcs = [g for g in gmfcs_map.keys() if g in df.columns]


    # --- VALIDATION FUNCTION ---
    # Checks if a cell contains 'X' or a number > 0
    def is_valid_gmfcs(val):
        s_val = str(val).strip().upper()
        if s_val in ['NAN', 'NONE', '', '???', '0']:
            return False
        if s_val == 'X':
            return True
        try:
            return float(val) > 0
        except ValueError:
            return False


    # --- DATA AGGREGATION ---
    data_list = []

    # Iterate through every Tool / Activity combination
    for tool in existing_tools:
        for activity in existing_activities:

            # Check distribution for each GMFCS level within this combination
            for gmfcs_col, gmfcs_label in gmfcs_map.items():
                if gmfcs_col in existing_gmfcs:
                    # Filter logic: Tool=1 AND Activity=1 AND GMFCS is Valid
                    mask = (
                            (df[tool] == 1) &
                            (df[activity] == 1) &
                            (df[gmfcs_col].apply(is_valid_gmfcs))
                    )
                    count = len(df[mask])

                    if count > 0:
                        data_list.append({
                            'Tool': tool,
                            'Activity': activity,
                            'GMFCS': gmfcs_label,
                            'Count': count
                        })

    df_plot = pd.DataFrame(data_list)

    if df_plot.empty:
        print("No matching data found (Check if columns contain 1s or Xs).")
    else:
        # --- GRAPH GENERATION ---
        plt.figure(figsize=(18, 10))
        sns.set_style("whitegrid")

        # Create dictionaries to map names to coordinate numbers (0, 1, 2...)
        tool_to_y = {t: i for i, t in enumerate(existing_tools)}
        act_to_x = {a: i for i, a in enumerate(existing_activities)}

        # Define Offsets (shifts) to separate the GMFCS bubbles horizontally
        # Level I on the left, Level IV on the right
        gmfcs_offsets = {
            "I": -0.25,
            "II": -0.08,
            "III": 0.08,
            "IV": 0.25
        }

        # Calculate custom X and Y positions
        # X position = Activity Index + GMFCS Offset
        df_plot['X_pos'] = df_plot['Activity'].map(act_to_x) + df_plot['GMFCS'].map(gmfcs_offsets)
        df_plot['Y_pos'] = df_plot['Tool'].map(tool_to_y)

        # --- PLOTTING ---
        scatter = sns.scatterplot(
            data=df_plot,
            x='X_pos',
            y='Y_pos',
            size='Count',
            hue='GMFCS',
            sizes=(50, 1000),  # Adjust min/max bubble size here
            hue_order=["I", "II", "III", "IV"],  # Force specific order
            palette="viridis",  # Colors: viridis, deep, or Set1
            alpha=0.8,  # Transparency
            edgecolor='black'
        )

        # --- FORMATTING ---
        # Set proper labels on the integer ticks
        plt.xticks(
            ticks=range(len(existing_activities)),
            labels=existing_activities,
            rotation=45,
            ha='right',
            fontsize=11
        )
        plt.yticks(
            ticks=range(len(existing_tools)),
            labels=existing_tools,
            fontsize=12
        )

        plt.title("Research Distribution by Tool, Activity, and Population (GMFCS)", fontsize=18, pad=20)
        plt.xlabel("Physical Activities", fontsize=14)
        plt.ylabel("Measurement Tools", fontsize=14)

        # Legend adjustment
        plt.legend(bbox_to_anchor=(1.01, 1), loc='upper left', borderaxespad=0, title="GMFCS Level & Count")

        # Grid lines to separate the main categories clearly
        plt.grid(True, which='major', linestyle='--', alpha=0.4)

        # Add margins so bubbles aren't cut off at the edges
        plt.margins(x=0.05, y=0.1)
        plt.tight_layout()
        plt.savefig(
            r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Plot\CO_occurence_combine_GMFCS_split.png")

        print("Displaying graph...")
        plt.show()

except Exception as e:
    print(f"An error occurred: {e}")