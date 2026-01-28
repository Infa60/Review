import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# 1. File Path
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"

try:
    # Load the file
    print("Loading file...")
    df = pd.read_excel(file_path, sheet_name=0)

    # Clean column names
    df.columns = df.columns.str.strip()

    # 2. Define Columns

    # GMFCS Columns (Rows of the heatmap)
    gmfcs_columns = ["Hemiplegic", "Diplegic", "Quadriplegic"]


    # Activity Columns (Columns of the heatmap)
    activity_columns = [
        "Sit-to-stand", "Running", "Cycling", "Stair-negotiation", "Obstacle-clearance", "Game",
        "Jumping","Time-Up-and-Go", "One-leg-standing",
        "Stepping-target", "Hopping", "Squat","Kicking-a-ball"
    ]

    # Verify columns exist
    existing_gmfcs = [col for col in gmfcs_columns if col in df.columns]
    existing_activities = [col for col in activity_columns if col in df.columns]


    # 3. Define the Logic Function
    # This function returns True if the value is 'X' or a number > 0
    # It returns False if the value is '???', 0, or empty
    def is_valid_entry(val):
        # Convert to string to handle text and numbers uniformly
        s_val = str(val).strip().upper()

        # Condition 1: If it is empty or 'NAN'
        if s_val == 'NAN' or s_val == 'NONE' or s_val == '':
            return False

        # Condition 2: If it is explicitly '???' or '0'
        if s_val == '???' or s_val == '0':
            return False

        # Condition 3: If it is 'X' -> Valid
        if s_val == 'X':
            return True

        # Condition 4: If it is a number > 0
        try:
            float_val = float(val)
            if float_val > 0:
                return True
        except ValueError:
            pass  # Not a number

        return False


    # 4. Create the Matrix
    matrix = pd.DataFrame(index=existing_gmfcs, columns=existing_activities)

    print("Calculating Co-occurrences with new logic (X or >0)...")

    for gmfcs in existing_gmfcs:
        # We look at each row in the Excel file
        for activity in existing_activities:
            count = 0

            # Iterate through the DataFrame rows
            for index, row in df.iterrows():
                val_gmfcs = row[gmfcs]
                val_activity = row[activity]

                # Logic:
                # The GMFCS column must be Valid (X or >0)
                # AND the Activity column must be 1 (or valid if it uses the same coding)
                if is_valid_entry(val_gmfcs) and (val_activity == 1):
                    count += 1

            matrix.loc[gmfcs, activity] = count

    # Ensure matrix is integer type
    matrix = matrix.fillna(0).astype(int)

    # 5. Generate the Heatmap
    plt.figure(figsize=(16, 6))

    sns.heatmap(matrix, annot=True, fmt="d", cmap="Greens", linewidths=.5)
    # Note: "OrRd" is Orange-Red palette, distinct from the previous Blue-Green

    plt.title("Co-occurrence of topography and functional tasks", fontsize=16)
    plt.ylabel("GMFCS level", fontsize=14)
    plt.xlabel("Functional tasks", fontsize=14)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Plot\CO_occurence_topography_task.png")

    print("Displaying graph...")
    plt.show()

except Exception as e:
    print(f"An error occurred: {e}")