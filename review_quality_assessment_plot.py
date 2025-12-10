import pandas as pd
import matplotlib.pyplot as plt
import textwrap
import os

# --- CONFIGURATION ---
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Quality_assessment_kappa.xlsx"
sheet_name = "Quality_assessment_results"
save_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Quality_assessment_plot.png"

# Color mapping
color_map = {
    0: 'red',  # Inadequate
    1: 'orange',  # Partial
    2: 'green'  # Adequate
}


def create_horizontal_quality_chart():
    if not os.path.exists(file_path):
        print(f"Error: File not found at path: {file_path}")
        return

    try:
        print("Loading data...")
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Select numeric columns only
        df_scores = df.select_dtypes(include=['number'])

        # Calculate counts
        counts = df_scores.apply(pd.Series.value_counts).fillna(0)
        counts = counts.reindex([0, 1, 2])
        counts_t = counts.T

        # Process Labels (Wrap text with more width allowed)
        labels_raw = counts_t.index.tolist()
        # Increased wrap width to 40 because horizontal space is generous
        labels_wrapped = [textwrap.fill(str(label), 23) for label in labels_raw]

        # --- PLOTTING ---
        print("Generating chart...")

        # Adjusted size: Taller (8) to accommodate the list of questions
        fig, ax = plt.subplots(figsize=(10, 8))

        colors = [color_map[0], color_map[1], color_map[2]]

        # CHANGED HERE: 'barh' for Horizontal Bar Chart
        counts_t.plot(kind='barh', stacked=True, color=colors, ax=ax, width=0.8, edgecolor='black')

        # --- OPTIMIZATIONS ---
        # 1. Invert Y-axis so Question 1 is at the TOP, not bottom
        ax.invert_yaxis()

        # 2. Force X-axis limit (Max number of articles)
        max_width = counts_t.sum(axis=1).max()
        ax.set_xlim(0, max_width + 0.5)

        # Titles and Labels
        plt.title('Distribution of quality scores per item', fontsize=14, pad=15)
        plt.xlabel("Number of articles", fontsize=10)  # Label is now on X-axis

        # Y-Axis settings (The Questions)
        ax.set_yticklabels(labels_wrapped, fontsize=9)  # Slightly larger font is readable now

        # Legend (Moved to the bottom or right)
        ax.legend(["0 - Inadequate", "1 - Partial", "2 - Adequate"],
                  title="Score", bbox_to_anchor=(1.0, 1.0), loc='upper left')

        # Bar Labels (Values inside bars)
        for c in ax.containers:
            # Logic to hide 0s remains the same
            labels = [int(v.get_width()) if v.get_width() > 0 else '' for v in c]
            ax.bar_label(c, labels=labels, label_type='center', fontsize=7, color='white', weight='bold')

        # Layout adjustment
        # 'left=0.3' reserves 30% of the image for the long question text
        plt.subplots_adjust(left=0.17, right=0.83, top=0.9)
        plt.savefig(save_path)
        plt.show()
        print("Chart generated successfully.")

    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    create_horizontal_quality_chart()