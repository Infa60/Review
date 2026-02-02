import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# 1. Configuration des chemins
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
output_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Plot\Combined_Full_Analysis.png"

# --- Listes des colonnes ---
activity_columns = ["Sit-to-stand", "Running", "Cycling", "Stair-negotiation", "Obstacle-clearance", "Game",
                    "Jumping", "Time-Up-and-Go", "One-leg-standing", "Stepping-target", "Hopping", "Squat",
                    "Kicking-a-ball"]

tool_columns = ["Optoelectronic", "Force-plate", "EMG", "Heart-rate-monitor", "Metabolic-cart", "IMU", "Wii-fit", "Other-tools"]
outcome_columns = ["Spatiotemporal", "Kinematic", "Kinetic", "Electromyographic", "Metabolic", "Stability", "Score"]
gmfcs_columns = ["GMFCS-I", "GMFCS-II", "GMFCS-III", "GMFCS-IV"]
topo_columns = ["Hemiplegic", "Diplegic", "Quadriplegic"]
motor_columns = ["Spastic", "Ataxic", "Dyskinetic", "Mixed"]

# --- Couleurs RGB Normalisées ---
def norm_rgb(r, g, b): return (r / 255, g / 255, b / 255)

colors_rgb = {
    'TOOLS': [norm_rgb(30, 80, 40)] * 8,
    'OUTCOME': [norm_rgb(40, 90, 190)] * 7,
    'GMFCS': [norm_rgb(120, 0, 0), norm_rgb(160, 20, 20), norm_rgb(200, 40, 40), norm_rgb(230, 80, 80), norm_rgb(255, 150, 150)],
    'TOPOGRAPHY': [norm_rgb(150, 120, 0), norm_rgb(200, 160, 0), norm_rgb(240, 200, 20), norm_rgb(255, 230, 80)],
    'CP SUBTYPE': [norm_rgb(120, 45, 0), norm_rgb(165, 60, 0), norm_rgb(210, 85, 0), norm_rgb(240, 120, 30), norm_rgb(255, 160, 80)]
}

# --- Logique de calcul ---
def is_valid_entry(val):
    s_val = str(val).strip().upper()
    if s_val in ['NAN', 'NONE', '', '???', '0']: return False
    try: return s_val == 'X' or float(val) > 0
    except: return False

def is_unknown_row(row, columns_group):
    for col in columns_group:
        if is_valid_entry(row[col]): return False
    return any('???' in str(row[col]) for col in columns_group)

try:
    df = pd.read_excel(file_path, sheet_name=0)
    df.columns = df.columns.str.strip()

    def get_matrix(rows_list, cols_list, add_unknown=False):
        existing_rows = [c for c in rows_list if c in df.columns]
        display_rows = existing_rows.copy()
        if add_unknown: display_rows.append("Unknown")
        mat = pd.DataFrame(0, index=display_rows, columns=cols_list)
        for r_col in existing_rows:
            for c_col in cols_list:
                mat.loc[r_col, c_col] = df.apply(lambda row: is_valid_entry(row[r_col]) and row[c_col] == 1, axis=1).sum()
        if add_unknown:
            for c_col in cols_list:
                mat.loc["Unknown", c_col] = df.apply(lambda row: is_unknown_row(row, existing_rows) and row[c_col] == 1, axis=1).sum()
        return mat

    m_tools = get_matrix(tool_columns, activity_columns, False)
    m_outcome = get_matrix(outcome_columns, activity_columns, False)

    m_gmfcs = get_matrix(gmfcs_columns, activity_columns, True)
    m_topo = get_matrix(topo_columns, activity_columns, True)
    m_motor = get_matrix(motor_columns, activity_columns, True)

    # --- MODIFICATIONS POUR CASES ÉTROITES ET LISIBILITÉ ---
    # Largeur réduite à 10 pour des cases étroites. Hauteur à 14 pour garder de l'espace vertical.
    height_ratios = [len(m_tools), len(m_outcome), len(m_gmfcs), len(m_topo), len(m_motor)]
    fig, axes = plt.subplots(5, 1, figsize=(12, 14), sharex=True, gridspec_kw={'height_ratios': height_ratios})

    blocks = [(m_tools, 'TOOLS', axes[0]),(m_outcome, 'OUTCOME', axes[1]), (m_gmfcs, 'GMFCS', axes[2]),
              (m_topo, 'TOPOGRAPHY', axes[3]), (m_motor, 'CP SUBTYPE', axes[4])]

    for mat, group_name, ax in blocks:
        data = mat.values
        max_val = data.max() if data.max() > 0 else 1
        colors = colors_rgb[group_name]

        rgba_img = np.ones((data.shape[0], data.shape[1], 4))
        for i in range(data.shape[0]):
            for j in range(data.shape[1]):
                val = data[i, j]
                base_color = colors[i]
                alpha = val / max_val if val > 0 else 0
                rgba_img[i, j, 0:3] = [1 - alpha * (1 - c) for c in base_color]

        ax.imshow(rgba_img, aspect='auto')

        # Annotations plus GROSSES (fontsize=12)
        for (i, j), val in np.ndenumerate(data):
            if val > 0:
                ax.text(j, i, int(val), ha='center', va='center', fontsize=12, fontweight='bold',
                        color="white" if (val / max_val > 0.5) else "black")

        # Labels d'axes plus GRANDS (fontsize=11)
        ax.set_yticks(np.arange(len(mat.index)))
        ax.set_yticklabels(mat.index, fontsize=11)
        ax.set_xticks(np.arange(len(activity_columns)))
        ax.set_xticklabels(activity_columns, fontsize=11)

        ax.set_xticks(np.arange(-.5, len(activity_columns), 1), minor=True)
        ax.set_yticks(np.arange(-.5, len(mat.index), 1), minor=True)
        ax.grid(which="minor", color="white", linestyle='-', linewidth=2)

        for spine in ax.spines.values():
            spine.set_visible(False)

        ax.tick_params(which="both", length=0)
        # Label de bloc à gauche
        #ax.set_ylabel(group_name, rotation=0, ha='right', va='center', labelpad=80, fontweight='bold', fontsize=13)

    plt.xticks(rotation=45, ha='right')
    plt.tight_layout(rect=[0.2, 0, 1, 0.98]) # Augmentation de la marge gauche pour les labels plus grands
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.show()

except Exception as e:
    print(f"Erreur : {e}")