import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 1. Configuration des chemins et colonnes
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
output_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Plot\Combined_Full_Analysis.png"

# Liste des activités (Axe X)
activity_columns = [
    "Sit-to-stand", "Running", "Cycling", "Stair-negotiation", "Obstacle-clearance", "Game",
    "Jumping", "Time-Up-and-Go", "One-leg-standing",
    "Stepping-target", "Hopping", "Squat", "Kicking-a-ball"
]

# Groupes de lignes (Axe Y)
tool_columns = ["Optoelectronic", "Force-plate", "EMG", "Heart-rate-monitor", "Metabolic-cart", "IMU", "Wii-fit",
                "Other-tools"]
gmfcs_columns = ["GMFCS-I", "GMFCS-II", "GMFCS-III", "GMFCS-IV"]
topo_columns = ["Hemiplegic", "Diplegic", "Quadriplegic"]
motor_columns = ["Spastic", "Ataxic", "Dyskinetic", "Mixed"]


def is_valid_entry(val):
    s_val = str(val).strip().upper()
    if s_val in ['NAN', 'NONE', '', '???', '0']: return False
    if s_val == 'X': return True
    try:
        return float(val) > 0
    except:
        return False


def is_unknown_row(row, columns_group):
    """Vérifie si une étude n'a aucune donnée valide mais contient des '???'"""
    for col in columns_group:
        if is_valid_entry(row[col]): return False
    for col in columns_group:
        if '???' in str(row[col]): return True
    return False


try:
    print("Chargement du fichier...")
    df = pd.read_excel(file_path, sheet_name=0)
    df.columns = df.columns.str.strip()


    def get_matrix(rows_list, cols_list, add_unknown=False):
        existing_rows = [c for c in rows_list if c in df.columns]
        display_rows = existing_rows.copy()
        if add_unknown: display_rows.append("Unknown")

        mat = pd.DataFrame(0, index=display_rows, columns=cols_list)

        # Co-occurrences standards
        for r_col in existing_rows:
            for c_col in cols_list:
                count = df.apply(lambda row: is_valid_entry(row[r_col]) and row[c_col] == 1, axis=1).sum()
                mat.loc[r_col, c_col] = count

        # Ligne Unknown
        if add_unknown:
            for c_col in cols_list:
                count_unk = df.apply(lambda row: is_unknown_row(row, existing_rows) and row[c_col] == 1, axis=1).sum()
                mat.loc["Unknown", c_col] = count_unk
        return mat


    print("Calcul des matrices...")
    m_tools = get_matrix(tool_columns, activity_columns, False)
    m_gmfcs = get_matrix(gmfcs_columns, activity_columns, True)
    m_topo = get_matrix(topo_columns, activity_columns, True)
    m_motor = get_matrix(motor_columns, activity_columns, True)

    # 2. Création de la figure avec 4 subplots
    # Calcul des ratios de hauteur pour l'uniformité des cases (8, 5, 4, 5)
    height_ratios = [len(m_tools), len(m_gmfcs), len(m_topo), len(m_motor)]

    fig, axes = plt.subplots(4, 1, figsize=(16, 18), sharex=True,
                             gridspec_kw={'height_ratios': height_ratios})

    # Configuration des plots
    plots_config = [
        (m_tools, "Measurement Tools", "Greens"),  # Vert émeraude doux
        (m_gmfcs, "GMFCS Levels", "Reds"),  # Rouge brique/classique
        (m_topo, "Topography", "YlOrBr"),  # Jaune ambré / Bronze (moins flashy que Wistia)
        (m_motor, "Motor Type", "Purples")  # Violet profond (pour bien différencier de l'orange/rouge)
    ]

    for ax, (data, label, cmap) in zip(axes, plots_config):
        sns.heatmap(data, ax=ax, annot=True, fmt="d", cmap=cmap, linewidths=.5)
        #ax.set_ylabel(label, fontsize=12, fontweight='bold')
        ax.tick_params(axis='y', rotation=0)

    # Axe X final
    axes[-1].set_xlabel("Functional Tasks", fontsize=14, labelpad=15)
    plt.xticks(rotation=45, ha='right')

    plt.tight_layout(rect=[0, 0, 1, 0.97])
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    print(f"Graphique complet généré avec succès : {output_path}")
    plt.show()

except Exception as e:
    print(f"Erreur lors de l'exécution : {e}")