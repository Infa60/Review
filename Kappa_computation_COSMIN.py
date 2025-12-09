import pandas as pd
import numpy as np
from sklearn.metrics import cohen_kappa_score
import os
import warnings

# --- CONFIGURATION ---
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\COSMIN_kappa.xlsx"

domains = [
    "Reliability",
    "Measurement error",
    "Criterion validity",
    "Convergent validity",
    "Discriminative validity"
]

suffixes = ("_MB", "_NH")

# Stockage pour le Grand Total
grand_total_mb = []
grand_total_nh = []


def calculate_metrics(v1, v2):
    """Calcule Kappa et Accord en gérant les exceptions"""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            if len(v1) == 0: return 0.0, 0.0

            # Pourcentage d'Accord
            accord = np.mean(v1 == v2) * 100

            # Kappa
            # Si accord parfait constant (ex: que des '2' partout)
            #if np.array_equal(v1, v2) and len(np.unique(v1)) == 1:
                #kappa = 1.0
            #else:
            kappa = cohen_kappa_score(v1, v2, weights='linear')
            if np.isnan(kappa): kappa = 0.0

            return kappa, accord
        except:
            return 0.0, 0.0


def analyze_domain_clean(xls, domain_name):
    sheet_mb = f"{domain_name}{suffixes[0]}"
    sheet_nh = f"{domain_name}{suffixes[1]}"

    try:
        # Lecture (header=0 suppose que la ligne 1 contient les noms d'articles)
        df_mb = pd.read_excel(xls, sheet_name=sheet_mb, header=0)
        df_nh = pd.read_excel(xls, sheet_name=sheet_nh, header=0)

        # --- TRANSPOSITION AUTOMATIQUE ---
        # Si vos colonnes sont des articles (ex: "Smith et al."), on transpose.
        # Ainsi : Lignes = Articles, Colonnes = Items. C'est plus facile à traiter.
        cols_check = [str(c).lower() for c in df_mb.columns[:5]]
        if any("et al" in c for c in cols_check) or any("20" in c for c in cols_check):
            # On ne garde que les données numériques/texte (pas les index bizarres)
            # On transpose pour avoir : Colonne = Item 1, Item 2...
            df_mb = df_mb.set_index(df_mb.columns[0])  # On suppose que col 1 = Noms des items
            df_nh = df_nh.set_index(df_nh.columns[0])

            # Transposition
            df_mb = df_mb.T
            df_nh = df_nh.T

        # Maintenant :
        # Lignes = Articles
        # Colonnes = Questions (Items)

        # Identification des colonnes communes (Questions)
        common_cols = [c for c in df_mb.columns if c in df_nh.columns]

        # Filtrage : On ne garde que les colonnes qui ressemblent à des questions
        # (On évite les métadonnées si elles ont survécu)
        question_cols = []
        for col in common_cols:
            # On garde tout ce qui est commun pour l'instant
            question_cols.append(col)

        if not question_cols:
            print(f"WARNING: No matching questions found for {domain_name}.")
            return

        # Listes pour le pool du domaine
        domain_mb = []
        domain_nh = []

        # --- ANALYSE QUESTION PAR QUESTION ---
        # (Utile pour voir quel item pose problème, même si N est petit)
        print(f"\n--- {domain_name.upper()} : DETAIL PAR QUESTION ---")
        print(f"{'QUESTION':<40} | {'N (Valid)'} | {'KAPPA'} | {'ACCORD'}")

        for col in question_cols:
            # Récupération des vecteurs bruts (tous les articles pour cette question)
            # On force en numérique (les "NA" ou texte deviennent NaN)
            s1 = pd.to_numeric(df_mb[col], errors='coerce')
            s2 = pd.to_numeric(df_nh[col], errors='coerce')

            # --- NETTOYAGE CHIRURGICAL ---
            # On ne garde l'article QUE SI les deux juges ont mis un chiffre
            mask = s1.notna() & s2.notna()

            v1 = s1[mask].values
            v2 = s2[mask].values

            if len(v1) > 0:
                # Ajout au pool
                domain_mb.extend(v1)
                domain_nh.extend(v2)
                grand_total_mb.extend(v1)
                grand_total_nh.extend(v2)

                # Calcul local (juste pour info)
                k, acc = calculate_metrics(v1, v2)
                print(f"{str(col)[:38]:<40} | {len(v1):<9} | {k:.3f} | {acc:.1f}%")

        # --- RÉSULTAT DU DOMAINE ---
        if len(domain_mb) > 0:
            arr_mb = np.array(domain_mb)
            arr_nh = np.array(domain_nh)
            k_dom, acc_dom = calculate_metrics(arr_mb, arr_nh)

            print("-" * 70)
            print(f">>> DOMAIN RESULT: {domain_name}")
            print(f"    Weighted Kappa : {k_dom:.4f}")
            print(f"    Agreement      : {acc_dom:.2f}%")
            print(f"    Valid Pairs    : {len(arr_mb)}")
            print("-" * 70)
        else:
            print(f"No valid data pairs found for {domain_name} (All NA).")

    except Exception as e:
        print(f"Skipping {domain_name} (Structure error or missing sheet): {e}")


# --- MAIN ---
if __name__ == "__main__":
    if os.path.exists(file_path):
        print("Loading Excel file...")
        try:
            xls = pd.ExcelFile(file_path)

            # 1. Analyse par Domaine
            for domain in domains:
                analyze_domain_clean(xls, domain)

            # 2. Calcul du GRAND TOTAL
            print("\n" + "=" * 80)
            print("GRAND TOTAL (ALL DOMAINS POOLED)")
            print("=" * 80)

            if len(grand_total_mb) > 0:
                gt_mb = np.array(grand_total_mb)
                gt_nh = np.array(grand_total_nh)

                gt_kappa, gt_accord = calculate_metrics(gt_mb, gt_nh)

                print(f"-> OVERALL Linear Weighted Kappa : {gt_kappa:.4f}")
                print(f"-> OVERALL Percent Agreement     : {gt_accord:.2f}%")
                print(f"-> TOTAL Valid Ratings           : {len(gt_mb)}")

                if gt_kappa <= 0.0:
                    verdict = "Poor"
                elif gt_kappa <= 0.2:
                    verdict = "Slight"
                elif gt_kappa <= 0.4:
                    verdict = "Fair"
                elif gt_kappa <= 0.6:
                    verdict = "Moderate"
                elif gt_kappa <= 0.8:
                    verdict = "Substantial"
                else:
                    verdict = "Almost perfect"
                print(f"-> Interpretation                : {verdict}")
            else:
                print("No data found anywhere.")

        except Exception as e:
            print(f"Critical Error: {e}")
    else:
        print("File not found.")