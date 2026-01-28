import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 1. Configuration des chemins
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
output_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Plot\Outside_Lab_Focus.png"

# --- Listes des catégories ---
tool_columns = ["Optoelectronic", "Force-plate", "EMG", "Heart-rate-monitor", "Metabolic-cart", "IMU", "Wii-fit",
                "Other-tools"]
gmfcs_columns = ["GMFCS-I", "GMFCS-II", "GMFCS-III", "GMFCS-IV"]
topo_columns = ["Hemiplegic", "Diplegic", "Quadriplegic"]
motor_columns = ["Spastic", "Ataxic", "Dyskinetic", "Mixed"]

all_possible_activities = [
    "Sit-to-stand", "Running", "Cycling", "Stair-negotiation", "Obstacle-clearance", "Game",
    "Jumping", "Time-Up-and-Go", "One-leg-standing", "Stepping-target", "Hopping", "Squat", "Kicking-a-ball"
]


def is_valid_entry(val):
    s_val = str(val).strip().upper()
    if s_val in ['NAN', 'NONE', '', '???', '0']: return False
    try:
        return s_val == 'X' or float(val) > 0
    except:
        return False


def is_unknown_row(row, columns_group):
    for col in columns_group:
        if is_valid_entry(row[col]): return False
    return any('???' in str(row[col]) for col in columns_group)


try:
    df_full = pd.read_excel(file_path, sheet_name=0)
    df_full.columns = df_full.columns.str.strip()

    # --- FILTRAGE : Exclure "Laboratory" pur, garder tout le reste ---
    # On garde si ce n'est pas "Laboratory" OU si ça contient "+" (ex: Laboratory + Community)
    # On s'assure de traiter les minuscules/majuscules et les espaces
    mask_outside = (df_full['Env.'].str.strip() != 'Laboratory') | (df_full['Env.'].str.contains(r'\+', na=False))
    df = df_full[mask_outside].copy()

    # Détection dynamique des tâches réellement présentes dans ce sous-groupe
    active_activities = [col for col in all_possible_activities if col in df.columns and df[col].sum() > 0]

    if not active_activities:
        print("Aucun article trouvé avec les critères de l'environnement spécifié.")
    else:
        def get_matrix(rows_list, cols_list, add_unknown=False):
            existing_rows = [c for c in rows_list if c in df.columns]
            display_rows = existing_rows.copy()
            if add_unknown: display_rows.append("Unknown")
            mat = pd.DataFrame(0, index=display_rows, columns=active_activities)
            for r_col in existing_rows:
                for c_col in active_activities:
                    mat.loc[r_col, c_col] = df.apply(lambda row: is_valid_entry(row[r_col]) and row[c_col] == 1,
                                                     axis=1).sum()
            if add_unknown:
                for c_col in active_activities:
                    mat.loc["Unknown", c_col] = df.apply(
                        lambda row: is_unknown_row(row, existing_rows) and row[c_col] == 1, axis=1).sum()
            return mat


        m_tools = get_matrix(tool_columns, active_activities, False)
        m_gmfcs = get_matrix(gmfcs_columns, active_activities, True)
        m_topo = get_matrix(topo_columns, active_activities, True)
        m_motor = get_matrix(motor_columns, active_activities, True)

        # --- Plotting Compact (Largeur ajustée au nombre de tâches trouvées) ---
        width = max(6, len(active_activities) * 1.2)
        height_ratios = [len(m_tools), len(m_gmfcs), len(m_topo), len(m_motor)]
        fig, axes = plt.subplots(4, 1, figsize=(width, 14), sharex=True, gridspec_kw={'height_ratios': height_ratios})


        # Couleurs RGB
        def norm_rgb(r, g, b):
            return (r / 255, g / 255, b / 255)


        colors_rgb = {
            'TOOLS': [norm_rgb(30, 80, 40)] * 8,
            'GMFCS': [norm_rgb(120, 0, 0), norm_rgb(160, 20, 20), norm_rgb(200, 40, 40), norm_rgb(230, 80, 80),
                      norm_rgb(255, 150, 150)],
            'TOPOGRAPHY': [norm_rgb(150, 120, 0), norm_rgb(200, 160, 0), norm_rgb(240, 200, 20),
                           norm_rgb(255, 230, 80)],
            'CP SUBTYPE': [norm_rgb(120, 45, 0), norm_rgb(165, 60, 0), norm_rgb(210, 85, 0), norm_rgb(240, 120, 30),
                           norm_rgb(255, 160, 80)]
        }

        blocks = [(m_tools, 'TOOLS', axes[0]), (m_gmfcs, 'GMFCS', axes[1]),
                  (m_topo, 'TOPOGRAPHY', axes[2]), (m_motor, 'CP SUBTYPE', axes[3])]

        for mat, group_name, ax in blocks:
            data = mat.values
            max_val = data.max() if data.max() > 0 else 1
            colors = colors_rgb[group_name]
            rgba_img = np.ones((data.shape[0], data.shape[1], 4))
            for i in range(data.shape[0]):
                for j in range(data.shape[1]):
                    val = data[i, j]
                    alpha = val / max_val if val > 0 else 0
                    rgba_img[i, j, 0:3] = [1 - alpha * (1 - c) for c in colors[i]]

            ax.imshow(rgba_img, aspect='auto')

            for (i, j), val in np.ndenumerate(data):
                if val > 0:
                    ax.text(j, i, int(val), ha='center', va='center', fontsize=13, fontweight='bold',
                            color="white" if (val / max_val > 0.5) else "black")

            ax.set_yticks(np.arange(len(mat.index)))
            ax.set_yticklabels(mat.index, fontsize=11)
            ax.set_xticks(np.arange(len(active_activities)))
            ax.set_xticklabels(active_activities, fontsize=11)

            # Grille blanche et suppression cadre
            ax.set_xticks(np.arange(-.5, len(active_activities), 1), minor=True)
            ax.set_yticks(np.arange(-.5, len(mat.index), 1), minor=True)
            ax.grid(which="minor", color="white", linestyle='-', linewidth=2)
            for spine in ax.spines.values(): spine.set_visible(False)
            ax.tick_params(which="both", length=0)
            ax.set_ylabel(group_name, rotation=0, ha='right', va='center', labelpad=70, fontweight='bold')

        plt.xticks(rotation=45, ha='right')
        plt.tight_layout(rect=[0.15, 0, 1, 0.98])
        plt.show()

except Exception as e:
    print(f"Erreur : {e}")