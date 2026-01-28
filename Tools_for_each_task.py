import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 1. File Path definition
# Using raw string (r"...") to handle Windows backslashes correctly
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"

try:
    # 2. Load the Excel file
    print("Loading file...")
    df = pd.read_excel(file_path, sheet_name=0)

    # Cleaning column names (removing potential leading/trailing spaces)
    df.columns = df.columns.str.strip()

    # 3. Define the lists of columns based on your images

    # List of TOOLS (from your second image)
    tool_columns = [
        "Optoelectronic", "Force-plate", "EMG", "Heart-rate-monitor",
        "Metabolic-cart","IMU",  "Wii-fit", "Other-tools"
    ]

    # List of ACTIVITIES (from your first image)
    activity_columns = [
        "Sit-to-stand", "Running", "Cycling", "Stair-negotiation", "Obstacle-clearance", "Game",
        "Jumping","Time-Up-and-Go", "One-leg-standing",
        "Stepping-target", "Hopping", "Squat","Kicking-a-ball"
    ]

    # Check which columns actually exist in the file to avoid errors
    existing_tools = [col for col in tool_columns if col in df.columns]
    existing_activities = [col for col in activity_columns if col in df.columns]

    # 4. Create the Co-occurrence Matrix
    # Rows = Tools, Columns = Activities
    matrix = pd.DataFrame(index=existing_tools, columns=existing_activities)

    print("Calculating co-occurrences...")

    for tool in existing_tools:
        # Filter rows where the specific tool is used (value == 1)
        tool_data = df[df[tool] == 1]

        # Sum the occurrences of each activity for this tool
        activity_counts = tool_data[existing_activities].sum()

        # Add to the matrix
        matrix.loc[tool] = activity_counts

    # Convert to integers (filling NaNs with 0 if any)
    matrix = matrix.fillna(0).astype(int)

    # 5. Generate the Single Graph (Heatmap)
    plt.figure(figsize=(16, 9))  # Set the size of the window

    # Create the Heatmap
    # annot=True: shows the numbers inside the squares
    # fmt="d": formats the numbers as integers
    # cmap="YlGnBu": Color palette (Yellow -> Green -> Blue)
    sns.heatmap(matrix, annot=True, fmt="d", cmap="YlGnBu", linewidths=.5, square=False)

    plt.title("Co-occurrence of measurement tools and functional tasks", fontsize=18)
    plt.ylabel("Measurement tools", fontsize=14)
    plt.xlabel("Functional tasks", fontsize=14)

    # Rotate x-axis labels for better readability
    plt.xticks(rotation=45, ha='right')

    # Adjust layout to prevent cutting off labels
    plt.tight_layout()
    plt.savefig(r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Plot\CO_occurence_tools_task.png")
    print("Displaying graph...")
    plt.show()

except FileNotFoundError:
    print(f"ERROR: File not found at path: {file_path}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")

except Exception as e:
    print(f"Une erreur est survenue : {e}")