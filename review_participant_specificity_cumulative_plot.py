import pandas as pd
import matplotlib.pyplot as plt
import numpy as np


# ==========================================
# 0. FONCTION UTILITAIRE : RGB (0-255)
# ==========================================
def rgb(r, g, b):
    return (r / 255, g / 255, b / 255)


# ==========================================
# 1. FONCTION : TEXTE COURBE
# ==========================================
def draw_curved_text(text, radius, center_angle, ax, fontsize=12, facteur_espacement=1.2):
    if not text: return
    center_angle = center_angle % 360
    center_angle_rad = np.deg2rad(center_angle)
    angle_per_char = (fontsize / radius / 100) * facteur_espacement * 3.5
    total_angle = len(text) * angle_per_char
    is_bottom = 180 < center_angle < 360

    if is_bottom:
        angles = np.linspace(center_angle_rad - total_angle / 2,
                             center_angle_rad + total_angle / 2, len(text))
    else:
        angles = np.linspace(center_angle_rad + total_angle / 2,
                             center_angle_rad - total_angle / 2, len(text))

    for char, angle in zip(text, angles):
        x = radius * np.cos(angle)
        y = radius * np.sin(angle)
        degrees = np.degrees(angle)
        if is_bottom:
            rotation = degrees - 90 + 180
        else:
            rotation = degrees - 90
        ax.text(x, y, char, rotation=rotation,
                ha='center', va='center', fontsize=fontsize, fontweight='bold', color='black')


# ==========================================
# 2. CONFIGURATION & COULEURS PERSONNALISÉES
# ==========================================

fichier_excel = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
col_total_article = 'N_CP'

config = {
    'SEX': {'Boys': 'Boy_with_CP', 'Girls': 'Girl_with_CP'},
    'TOPOGRAPHY': {'Hemiplegia': 'Hemiplegic', 'Diplegia': 'Diplegic', 'Quadriplegia': 'Quadriplegic'},
    'CP SUBTYPE': {'Spastic': 'Spastic', 'Ataxic': 'Ataxic', 'Dyskinetic': 'Dyskinetic', 'Mixed': 'Mixed'},
    'GMFCS': {'I': 'GMFCS-I', 'II': 'GMFCS-II', 'III': 'GMFCS-III', 'IV': 'GMFCS-IV'}
}

# --- VOS COULEURS RGB (AVEC GESTION DE L'INCONNU PAR BLOC) ---
couleurs_perso = {
    'SEX': {
        'Boys': rgb(199, 21, 133),
        'Girls': rgb(255, 105, 180),
        'Unknown': rgb(255, 204, 229)  # <--- Inconnu Sexe (Bleu très pâle)
    },
    'TOPOGRAPHY': {
        'Hemiplegia': rgb(150, 120, 0),
        'Diplegia': rgb(200, 160, 0),
        'Quadriplegia': rgb(240, 200, 20),
        'Unknown': rgb(255, 230, 80)  # <--- Inconnu Topo (Crème/Jaune pâle)
    },
    'CP SUBTYPE': {
        'Spastic': rgb(120, 45, 0),
        'Dyskinetic': rgb(165, 60, 0),
        'Ataxic': rgb(210, 85, 0),
        'Mixed': rgb(240, 120, 30),
        'Unknown': rgb(255, 160, 80)  # <--- Inconnu Subtype (Mauve très pâle)
    },
    'GMFCS': {
        'I': rgb(120, 0, 0),
        'II': rgb(160, 20, 20),
        'III': rgb(200, 40, 40),
        'IV': rgb(230, 80, 80),
        'Unknown': rgb(255, 150, 150)  # <--- Inconnu GMFCS (Rose très pâle)
    }
}

# Couleur de secours (si on a oublié de définir 'Unknown' dans un bloc)
couleur_inconnu_defaut = rgb(220, 220, 220)


# --- CHARGEMENT ---
def nettoyer_valeur(val):
    try:
        return float(val) if not np.isnan(float(val)) else 0
    except:
        return 0


try:
    df = pd.read_excel(fichier_excel)
except:
    df = pd.DataFrame(columns=[col_total_article])
cols_to_clean = [col_total_article]
for cat in config.values(): cols_to_clean.extend(cat.values())
for col in cols_to_clean:
    if col in df.columns:
        df[col] = df[col].apply(nettoyer_valeur)
    else:
        df[col] = 0
if not df.empty:
    GRAND_TOTAL = df[col_total_article].sum()
else:
    GRAND_TOTAL = 1

# ==========================================
# 3. PRÉPARATION DES DONNÉES
# ==========================================

valeurs_interne = []
labels_interne = []
valeurs_externe = []
labels_externe = []
couleurs_externe = []

for nom_categorie, sous_categories in config.items():
    valeurs_interne.append(1)
    labels_interne.append(nom_categorie)

    palette_actuelle = couleurs_perso.get(nom_categorie, {})
    somme_connue = 0

    # Données connues
    for label_sub, col_name in sous_categories.items():
        total_sub = df[col_name].sum()
        somme_connue += total_sub
        valeurs_externe.append(total_sub / len(config) / GRAND_TOTAL)
        labels_externe.append(f"{label_sub}\n({int(total_sub)})")

        # Couleur connue
        code_couleur = palette_actuelle.get(label_sub, (0, 0, 0))
        couleurs_externe.append(code_couleur)

    # Données inconnues
    total_inconnu = GRAND_TOTAL - somme_connue
    if total_inconnu > 0:
        valeurs_externe.append(total_inconnu / len(config) / GRAND_TOTAL)
        labels_externe.append(f"Unknown\n({int(total_inconnu)})")

        # --- NOUVELLE LOGIQUE ICI ---
        # On cherche la couleur 'Unknown' spécifique au bloc.
        # Si elle n'existe pas, on prend le gris par défaut.
        couleur_specifique = palette_actuelle.get('Unknown', couleur_inconnu_defaut)
        couleurs_externe.append(couleur_specifique)

# ==========================================
# 4. AFFICHAGE FINAL
# ==========================================

fig, ax = plt.subplots(figsize=(12, 12))
ax.axis('equal')

# ANNEAU 1
wedges_interne, _ = ax.pie(valeurs_interne, radius=0.65,
                           colors=['white'] * len(config),
                           wedgeprops=dict(width=0.25, edgecolor='white'),
                           startangle=90)

# ANNEAU 2
wedges_externe, text_externe = ax.pie(valeurs_externe, radius=1.0, labels=labels_externe,
                                      labeldistance=1.1, colors=couleurs_externe,
                                      rotatelabels=False, startangle=90,
                                      wedgeprops=dict(width=0.35, edgecolor='white'))

# GRANDS RAYONS NOIRS
for wedge in wedges_interne:
    theta_rad = np.deg2rad(wedge.theta2)
    x1, y1 = 0.40 * np.cos(theta_rad), 0.40 * np.sin(theta_rad)
    x2, y2 = 1.00 * np.cos(theta_rad), 1.00 * np.sin(theta_rad)
    ax.plot([x1, x2], [y1, y2], color='black', linewidth=2)

# FINITIONS
ax.add_patch(plt.Circle((0, 0), 0.40, fill=False, edgecolor='black', linewidth=2))

for i, wedge in enumerate(wedges_interne):
    angle_centre = (wedge.theta1 + wedge.theta2) / 2
    draw_curved_text(labels_interne[i], radius=0.53, center_angle=angle_centre,
                     ax=ax, fontsize=13, facteur_espacement=0.12)

for t in text_externe:
    t.set_fontsize(13)

centre_circle = plt.Circle((0, 0), 0.39, fc='white')
fig.gca().add_artist(centre_circle)

plt.text(0, 0, f"TOTAL CHILDREN\nN = {int(GRAND_TOTAL)}", ha='center', va='center', fontsize=14, fontweight='bold')
plt.title("Clinical Characteristics Overview", fontsize=16, pad=20)
plt.tight_layout()
plt.savefig(r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Plot\Figure_2.svg",
            format='svg', bbox_inches='tight')
plt.show()