import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import os

# ==========================================
# 1. CONFIGURATION
# ==========================================
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
sheet_name = "Global_overview"
save_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Type_of_article_plot.png"

column_name = "Study_type"

# ==========================================
# 2. DATA LOADING
# ==========================================

if not os.path.exists(file_path):
    print(f"Error: File not found at path: {file_path}")
else:
    try:
        # Load Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Check if the column exists
        if column_name not in df.columns:
            print(f"Error: The column '{column_name}' was not found.")
            print("Available columns:", list(df.columns))
            print("Please update the 'column_name' variable in the script.")
        else:
            # Count occurrences and sort (descending)
            counts = df[column_name].value_counts().reset_index()
            counts.columns = ['Study Type', 'Count']

            # ==========================================
            # 3. PLOTTING (SOBER STYLE)
            # ==========================================

            # Set a minimalist theme
            sns.set_theme(style="white", context="talk")

            plt.figure(figsize=(9, 3))

            # Create Horizontal Bar Plot with a single professional color
            ax = sns.barplot(
                x="Count",
                y="Study Type",
                data=counts,
                color="#4c72b0",
                edgecolor=None
            )

            # Place the numbers INSIDE the bars
            for container in ax.containers:
                ax.bar_label(
                    container,
                    label_type='center',  # Centers text inside the bar
                    color='white',  # White text for contrast
                    fontsize=10,
                    fontweight='bold'
                )

            # Aesthetic cleanup (Removing noise)
            ax.tick_params(axis='y', labelsize=10)  # Ajustez 'labelsize' à la valeur souhaitée
            sns.despine(left=True, bottom=True)  # Remove frame borders
            ax.set_xlabel("Number of articles", fontsize=10)  # Remove X label (redundant)
            ax.set_ylabel("")  # Remove Y label
            ax.set_xticks([])  # Remove X axis ticks (0, 5, 10...) for a cleaner look

            # Title
            #plt.title("Distribution of study designs", fontsize=22, fontweight='bold', pad=20, loc='left')

            plt.tight_layout()
            plt.savefig(save_path)
            plt.show()

    except Exception as e:
        print(f"An error occurred: {e}")