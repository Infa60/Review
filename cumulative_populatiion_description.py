#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
clinical_overview_figure.py
===========================

Génère la figure « Clinical Characteristics Overview » (donut à 4 quadrants)
à partir du classeur Excel d'inclusion de la revue systématique.

Adapté pour une exécution directe depuis PyCharm. Sortie en SVG.
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge, Circle

# --------------------------------------------------------------------------- #
# 1. CONFIGURATION DES CHEMINS ET OPTIONS (POUR PYCHARM)                      #
# --------------------------------------------------------------------------- #

# Modifiez ces chemins selon l'emplacement de vos fichiers
EXCEL_FILE_PATH = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
OUTPUT_IMAGE_PATH = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Plot\Clinical_overview.svg"

# Options d'exécution
DROP_UPDATES = False  # Mettre à True pour exclure les "zzadd update"
SHOW_FIGURE = True  # Mettre à True pour afficher la figure dans PyCharm après création

# --------------------------------------------------------------------------- #
# 2. PARAMÈTRES DU GRAPHIQUE ET DES DONNÉES                                   #
# --------------------------------------------------------------------------- #

SHEET_NAME = "Global_overview"
HEADER_ROW = 0

COL_TOTAL = "N_CP"  # effectif d'enfants PC par article
COL_AGE_MEAN = "Mean_age_CP"  # âge moyen par article
COL_AGE_SD = "SD_age_CP"  # écart-type d'âge par article
COL_FLAG = "Exclude reason"  # annotations de workflow (filtre optionnel)

AGE_SD_METHOD = "pooled"

# Définition des 4 groupes affichés
GROUPS = [
    {
        "title": "GMFCS level",
        "label_color": "#E60000",
        "quadrant": (0, 90),
        "columns": {"I": "GMFCS-I", "II": "GMFCS-II", "III": "GMFCS-III",
                    "IV": "GMFCS-IV", "V": "GMFCS-V"},
        "colors": {"I": "#6E0000", "II": "#B01B1B", "III": "#D62728",
                   "IV": "#E8483C", "V": "#F07A6E", "Unknown": "#F6A7A0"},
    },
    {
        "title": "Sex",
        "label_color": "#C71585",
        "quadrant": (90, 180),
        "columns": {"Boys": "Boy_with_CP", "Girls": "Girl_with_CP"},
        "colors": {"Boys": "#C71585", "Girls": "#FF6EB4", "Unknown": "#FBCCE7"},
    },
    {
        "title": "Topography",
        "label_color": "#C9A227",
        "quadrant": (180, 270),
        "columns": {"Hemiplegia": "Hemiplegic", "Diplegia": "Diplegic",
                    "Quadriplegia": "Quadriplegic", "NA": "NA"},
        "colors": {"Hemiplegia": "#8A7A00", "Diplegia": "#C9A227",
                   "Quadriplegia": "#E8C547","NA": "#C4B98E", "Unknown": "#FFE34D"},
    },
    {
        "title": "CP subtype",
        "label_color": "#E8730C",
        "quadrant": (270, 360),
        "columns": {"Spastic": "Spastic", "Ataxic": "Ataxic",
                    "Dyskinetic": "Dyskinetic", "Mixed": "Mixed"},
        "colors": {"Spastic": "#5E3312", "Ataxic": "#8A5A2B",
                   "Dyskinetic": "#B5763A", "Mixed": "#D68A4A", "Unknown": "#F4A24C"},
    },
]

# --- Géométrie du donut ---
R_OUTER = 1.00
R_INNER = 0.64
R_CENTER = 0.36
R_GROUP_LABEL = 0.505
R_SEG_LABEL = 1.14

# ---> CHANGEMENTS ICI : Taille augmentée et espacement élargi <---
GROUP_LABEL_SIZE = 14  # Avant: 11. Taille de la police
DEG_PER_CHAR_FACTOR = 0.28  # Avant: 0.160. Facteur d'espacement entre les lettres
MIN_LABEL_GAP_DEG = 9.0
THIN_THRESHOLD_DEG = 6


# --------------------------------------------------------------------------- #
# 3. LECTURE, AGRÉGATION, ÂGE                                                 #
# --------------------------------------------------------------------------- #

def to_numeric(series: pd.Series) -> pd.Series:
    """'???', 'X', vides -> NaN."""
    return pd.to_numeric(series, errors="coerce")


def compute_age(df: pd.DataFrame, method: str = AGE_SD_METHOD):
    """Âge moyen pondéré par l'effectif + écart-type d'ensemble."""
    n = to_numeric(df[COL_TOTAL])
    m = to_numeric(df[COL_AGE_MEAN])
    sd = to_numeric(df[COL_AGE_SD])

    valid_m = n.notna() & m.notna() & (n > 0)
    if valid_m.sum() == 0:
        return None
    N = n[valid_m].sum()
    grand_mean = (n[valid_m] * m[valid_m]).sum() / N

    valid_s = valid_m & sd.notna()
    ns, ms, sds = n[valid_s], m[valid_s], sd[valid_s]
    if method == "weighted":
        overall_sd = (ns * sds).sum() / ns.sum()
    else:  # "pooled" : combine variance intra et inter-études
        ss_within = ((ns - 1) * sds ** 2).sum()
        ss_between = (ns * (ms - grand_mean) ** 2).sum()
        overall_sd = np.sqrt((ss_within + ss_between) / (N - 1))

    return {"mean": float(grand_mean), "sd": float(overall_sd),
            "n_studies": int(valid_m.sum())}


def load_and_aggregate(excel_path: str, drop_updates: bool = False):
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME, header=HEADER_ROW)

    if drop_updates and COL_FLAG in df.columns:
        mask = df[COL_FLAG].astype(str).str.contains("zzadd update", na=False)
        df = df[~mask]

    total = int(to_numeric(df[COL_TOTAL]).fillna(0).sum())

    groups_data = []
    for g in GROUPS:
        counts = {}
        for label, col in g["columns"].items():
            if col not in df.columns:
                raise KeyError(f"Colonne manquante dans l'Excel : {col!r}")
            counts[label] = int(to_numeric(df[col]).fillna(0).sum())
        known = sum(counts.values())
        counts["Unknown"] = max(total - known, 0)
        counts = {k: v for k, v in counts.items() if v > 0}
        groups_data.append({**g, "counts": counts})

    age = compute_age(df)
    return total, groups_data, age


# --------------------------------------------------------------------------- #
# 4. OUTILS DE TRACÉ                                                          #
# --------------------------------------------------------------------------- #

def _polar(radius: float, angle_deg: float):
    a = np.radians(angle_deg)
    return radius * np.cos(a), radius * np.sin(a)


def _declutter(mid_angles: list, min_gap: float) -> list:
    """Écarte les ancres d'étiquettes trop proches."""
    anchors = list(mid_angles)
    for _ in range(400):
        moved = False
        for i in range(len(anchors) - 1):
            gap = anchors[i + 1] - anchors[i]
            if gap < min_gap:
                shift = (min_gap - gap) / 2.0
                anchors[i] -= shift
                anchors[i + 1] += shift
                moved = True
        if not moved:
            break
    return anchors


def draw_curved_text(ax, text, radius, center_angle, color, fontsize, zorder=3.5):
    """Écrit `text` lettre par lettre le long d'un arc."""
    deg_per_char = DEG_PER_CHAR_FACTOR * fontsize / radius
    total = len(text) * deg_per_char
    top = np.sin(np.radians(center_angle)) >= 0
    for i, ch in enumerate(text):
        if top:
            ang = center_angle + total / 2 - (i + 0.5) * deg_per_char
            rot = ang - 90
        else:
            ang = center_angle - total / 2 + (i + 0.5) * deg_per_char
            rot = ang + 90
        x, y = _polar(radius, ang)
        ax.text(x, y, ch, ha="center", va="center", rotation=rot,
                rotation_mode="anchor", fontsize=fontsize, fontweight="bold",
                color=color, zorder=zorder)


# --------------------------------------------------------------------------- #
# 5. GÉNÉRATION DE LA FIGURE                                                  #
# --------------------------------------------------------------------------- #

def draw_figure(total, groups_data, age, out_path,
                title="Clinical Characteristics Overview"):
    fig, ax = plt.subplots(figsize=(10.5, 10.5))
    ax.set_aspect("equal")
    ax.axis("off")

    for g in groups_data:
        a0, a1 = g["quadrant"]
        span = a1 - a0
        counts = g["counts"]
        group_total = sum(counts.values()) or 1

        segments, cursor = [], a0
        for label, value in counts.items():
            seg = span * value / group_total
            segments.append((label, value, cursor, cursor + seg,
                             cursor + seg / 2.0, seg))
            cursor += seg

        for label, value, start, end, mid, seg in segments:
            # Si le segment fait moins de 2 degrés, on affine énormément la bordure
            edge_width = 1.2 if seg > 2.0 else 0.1
            ax.add_patch(Wedge((0, 0), R_OUTER, start, end, width=R_OUTER - R_INNER,
                               facecolor=g["colors"].get(label, "#CCCCCC"),
                               edgecolor="white", linewidth=edge_width, zorder=1))

        anchors = _declutter([s[4] for s in segments], MIN_LABEL_GAP_DEG)
        for (label, value, start, end, mid, seg), anchor in zip(segments, anchors):
            lx, ly = _polar(R_SEG_LABEL, anchor)
            ha = "left" if np.cos(np.radians(anchor)) >= 0 else "right"
            va = "bottom" if np.sin(np.radians(anchor)) >= 0 else "top"
            ax.text(lx, ly, f"{label}\n({value})", ha=ha, va=va,
                    fontsize=11, color="#111111", zorder=6)
            if abs(anchor - mid) > 1.0 or seg < THIN_THRESHOLD_DEG:
                x0, y0 = _polar(R_OUTER, mid)
                x1, y1 = _polar(R_SEG_LABEL - 0.03, anchor)
                ax.plot([x0, x1], [y0, y1], color="#888888", lw=0.6, zorder=1)

        draw_curved_text(ax, g["title"], R_GROUP_LABEL, (a0 + a1) / 2.0,
                         g["label_color"], GROUP_LABEL_SIZE)

    for ang in (0, 90, 180, 270):
        x, y = _polar(R_INNER, ang)
        ax.plot([0, x], [0, y], color="black", lw=1.4, zorder=2.5)

    ax.add_patch(Circle((0, 0), R_INNER, facecolor="none",
                        edgecolor="black", lw=1.4, zorder=3))
    ax.add_patch(Circle((0, 0), R_CENTER, facecolor="white",
                        edgecolor="black", lw=1.6, zorder=4))

    ax.text(0, 0.13, f"TOTAL CHILDREN\nN = {total}", ha="center", va="center",
            fontsize=13, fontweight="bold", zorder=5)
    if age is not None:
        ax.text(0, -0.055, "Mean age", ha="center", va="center",
                fontsize=11.5, fontweight="bold", zorder=5)
        ax.text(0, -0.145, f"{age['mean']:.1f} \u00b1 {age['sd']:.1f} yrs",
                ha="center", va="center", fontsize=11.5, fontweight="bold", zorder=5)

    ax.set_title(title, fontsize=17, pad=18)
    lim = R_SEG_LABEL + 0.30
    ax.set_xlim(-lim, lim)
    ax.set_ylim(-lim, lim)

    fig.tight_layout()

    # Création du dossier de destination s'il n'existe pas
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)

    fig.savefig(out_path, bbox_inches="tight", format="svg")
    print(f"Figure enregistrée avec succès : {out_path}")
    return fig


# --------------------------------------------------------------------------- #
# 6. MAIN                                                                     #
# --------------------------------------------------------------------------- #

def main():
    print(f"Lecture du fichier : {EXCEL_FILE_PATH}")
    if not Path(EXCEL_FILE_PATH).exists():
        raise SystemExit(f"Erreur : Le fichier est introuvable. Veuillez vérifier EXCEL_FILE_PATH.")

    total, groups_data, age = load_and_aggregate(EXCEL_FILE_PATH, drop_updates=DROP_UPDATES)

    print(f"\nTOTAL CHILDREN (N_CP) = {total}")
    if age is not None:
        print(f"Âge moyen = {age['mean']:.1f} ± {age['sd']:.1f} ans "
              f"({age['n_studies']} études)\n")

    for g in groups_data:
        print(f"{g['title']}:")
        for label, value in g["counts"].items():
            print(f"    {label:<13} {value}")
        print()

    draw_figure(total, groups_data, age, OUTPUT_IMAGE_PATH)

    if SHOW_FIGURE:
        plt.show()


if __name__ == "__main__":
    main()