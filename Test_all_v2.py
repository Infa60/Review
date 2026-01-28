import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# 1. Chemin du fichier
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"

try:
    print("Chargement des données...")
    df = pd.read_excel(file_path, sheet_name=0)
    df.columns = df.columns.str.strip()

    # --- DEFINITIONS ---
    tools = ["Optoelectronic", "Force-plate", "IMU", "EMG", "Wii-fit", "Heart-rate-monitor", "Metabolic-cart",
             "Other-tools"]
    activities = ["Sit-to-stand", "Running", "Cycling", "Stair-negotiation", "Time-Up-and-Go", "Obstacle-clearance",
                  "Game", "One-leg-standing", "Jumping", "Squat", "Stepping-target", "Hopping", "GMFM-E",
                  "Kicking-a-ball"]
    gmfcs_map = {"GMFCS-I": "I", "GMFCS-II": "II", "GMFCS-III": "III", "GMFCS-IV": "IV"}

    existing_tools = [t for t in tools if t in df.columns]
    existing_activities = [a for a in activities if a in df.columns]
    existing_gmfcs_cols = [g for g in gmfcs_map.keys() if g in df.columns]


    def is_valid_gmfcs(val):
        s = str(val).strip().upper()
        if s == 'X': return True
        try:
            return float(val) > 0
        except:
            return False


    # --- PREPARATION DES DONNEES ---
    data_list = []

    for tool in existing_tools:
        for activity in existing_activities:
            for gmfcs_col, gmfcs_label in gmfcs_map.items():
                if gmfcs_col in existing_gmfcs_cols:
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
        print("Aucune donnée trouvée.")
    else:
        # Tri pour dessiner les grosses bulles d'abord (en arrière-plan)
        df_plot = df_plot.sort_values(by='Count', ascending=False)

        # --- CONFIGURATION GRAPHIQUE XL ---
        plt.figure(figsize=(20, 12))  # Fenêtre beaucoup plus grande
        sns.set_style("whitegrid")

        # Mapping positions
        tool_to_y = {t: i for i, t in enumerate(existing_tools)}
        act_to_x = {a: i for i, a in enumerate(existing_activities)}

        # Jitter (Décalage)
        np.random.seed(42)
        # On garde un jitter petit pour que les bulles restent groupées,
        # mais vu qu'elles sont grosses, elles vont bien se superposer.
        df_plot['X_pos'] = df_plot['Activity'].map(act_to_x) + np.random.uniform(-0.06, 0.06, size=len(df_plot))
        df_plot['Y_pos'] = df_plot['Tool'].map(tool_to_y) + np.random.uniform(-0.06, 0.06, size=len(df_plot))

        # --- DESSIN DES BULLES GEANTES ---
        scatter = sns.scatterplot(
            data=df_plot,
            x='X_pos',
            y='Y_pos',
            size='Count',
            hue='GMFCS',
            # ICI : On augmente massivement la taille min et max (ex: de 500 à 5000 pixels)
            sizes=(100, 3000),
            hue_order=["I", "II", "III", "IV"],
            palette=["blue", "darkorange", "green", "red"],
            alpha=0.7,  # Plus transparent car les bulles sont très grosses
            edgecolor='none',
            #linewidth=1
        )

        # --- TEXTES PLUS GROS ---
        plt.xticks(
            ticks=range(len(existing_activities)),
            labels=existing_activities,
            rotation=45,
            ha='right',
            fontsize=9,  # Texte axe X plus gros
            fontweight='bold'
        )
        plt.yticks(
            ticks=range(len(existing_tools)),
            labels=existing_tools,
            fontsize=9,  # Texte axe Y plus gros
            fontweight='bold'
        )

        plt.title("Research Overview: Tools, Activities & Population (Overlapping View)", fontsize=15, pad=30,
                  fontweight='bold')
        plt.xlabel("Physical Activities", fontsize=18, labelpad=15)
        plt.ylabel("Measurement Tools", fontsize=18, labelpad=15)

        # Légende ajustée
        plt.legend(bbox_to_anchor=(1.01, 1), loc='upper left', title="GMFCS Level & Volume", fontsize=9,
                   title_fontsize=9)

        # Marges pour ne pas couper les énormes bulles
        plt.margins(x=0.08, y=0.08)

        plt.tight_layout()
        plt.savefig(
            r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Plot\CO_occurence_combine.png")
        print("Affichage du graphique...")
        plt.show()

except Exception as e:
    print(f"Erreur : {e}")