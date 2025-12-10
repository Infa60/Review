import pandas as pd
import numpy as np
from sklearn.metrics import cohen_kappa_score
import os
import warnings

# --- CONFIGURATION ---
# Update this path to your actual file location
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\COSMIN_kappa.xlsx"

domains = [
    "Reliability",
    "Measurement error",
    "Criterion validity",
    "Convergent validity",
    "Discriminative validity"
]

# Sheet name suffixes for each rater
suffixes = ("_MB", "_NH")

# Storage for the Grand Total (Pooled data)
grand_total_mb_final = []
grand_total_nh_final = []


def calculate_metrics(v1, v2):
    """
    Calculates Linear Weighted Kappa and Raw Percent Agreement.
    Handles exceptions for empty vectors or constant values.
    """
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            if len(v1) == 0: return np.nan, 0.0

            # Percent Agreement
            accord = np.mean(v1 == v2) * 100

            # Kappa (Linear Weighted for ordinal scales)
            # If there is no variance (e.g., all scores are '2'), Kappa might return NaN or 0
            kappa = cohen_kappa_score(v1, v2, weights='linear')

            if np.isnan(kappa): kappa = 0.0

            return kappa, accord
        except:
            return 0.0, 0.0


def get_worst_score_per_article(df):
    """
    Applies the COSMIN 'Worst Score Counts' principle for each article.
    Returns a Series with the minimum score (0-3) per row (article).

    Logic:
    1. Convert all data to numeric.
    2. Calculate min value across columns (questions) for each row.
    3. Ignore NA values unless the entire row is NA.
    """
    # Ensure data is numeric
    df_numeric = df.apply(pd.to_numeric, errors='coerce')

    # Calculate min per row (axis=1), skipping NaNs
    worst_scores = df_numeric.min(axis=1, skipna=True)

    return worst_scores


def analyze_domain_worst_score(xls, domain_name):
    sheet_mb = f"{domain_name}{suffixes[0]}"
    sheet_nh = f"{domain_name}{suffixes[1]}"

    try:
        # Load sheets
        df_mb = pd.read_excel(xls, sheet_name=sheet_mb, header=0)
        df_nh = pd.read_excel(xls, sheet_name=sheet_nh, header=0)

        # --- TRANSPOSITION & CLEANING ---
        # Check if the first column contains article names (e.g., "Smith et al.")
        cols_check = [str(c).lower() for c in df_mb.columns[:5]]
        if any("et al" in c for c in cols_check) or any("20" in c for c in cols_check):
            # Set the first column as index (Article Names)
            df_mb = df_mb.set_index(df_mb.columns[0])
            df_nh = df_nh.set_index(df_nh.columns[0])
            # Transpose: Rows = Articles, Columns = Questions
            df_mb = df_mb.T
            df_nh = df_nh.T

        # --- FILTER COMMON COLUMNS ---
        # Keep only columns (questions) that exist in both sheets
        common_cols = [c for c in df_mb.columns if c in df_nh.columns]

        if not common_cols:
            print(f"WARNING: No matching questions found for {domain_name}.")
            return

        # Restrict DataFrames to common questions only
        df_mb_clean = df_mb[common_cols]
        df_nh_clean = df_nh[common_cols]

        # --- APPLY 'WORST SCORE COUNTS' ---
        # Result: One single score per article per rater
        scores_mb = get_worst_score_per_article(df_mb_clean)
        scores_nh = get_worst_score_per_article(df_nh_clean)

        # --- ALIGNMENT & FINAL CLEANING ---
        # Create a temp DataFrame to align articles by index
        comparison_df = pd.DataFrame({'MB': scores_mb, 'NH': scores_nh})

        # Drop articles where one or both raters have no valid score (NaN)
        # (e.g., Article not evaluated for this specific domain)
        valid_comparison = comparison_df.dropna()

        v1 = valid_comparison['MB'].values
        v2 = valid_comparison['NH'].values

        # --- CALCULATION & PRINTING ---
        print(f"\n--- {domain_name.upper()} : FINAL SCORES (Worst Score Counts) ---")

        if len(v1) > 0:
            # Add to the global pool for final calculation
            grand_total_mb_final.extend(v1)
            grand_total_nh_final.extend(v2)

            kappa, accord = calculate_metrics(v1, v2)

            print(f"    Valid Articles : {len(v1)}")
            print(f"    Raw Agreement  : {accord:.2f}%")

            # Display 'NC' if N < 4 to avoid statistically meaningless Kappas
            if len(v1) < 4:
                print(f"    Weighted Kappa : NC (N={len(v1)} too small)")
            else:
                print(f"    Weighted Kappa : {kappa:.4f}")

        else:
            print("    No valid articles found (All NA).")

    except Exception as e:
        print(f"Skipping {domain_name} (Error or missing sheet): {e}")


# --- MAIN EXECUTION ---
if __name__ == "__main__":
    if os.path.exists(file_path):
        print("Loading Excel file...")
        print("Applying COSMIN 'Worst Score Counts' logic per article...")
        try:
            xls = pd.ExcelFile(file_path)

            # 1. Analyze per Domain (Final Score only)
            for domain in domains:
                analyze_domain_worst_score(xls, domain)

            # 2. Calculate GRAND TOTAL (POOLED FINAL SCORES)
            print("\n" + "=" * 80)
            print("GRAND TOTAL (POOLED FINAL SCORES)")
            print("=" * 80)

            if len(grand_total_mb_final) > 0:
                gt_mb = np.array(grand_total_mb_final)
                gt_nh = np.array(grand_total_nh_final)

                gt_kappa, gt_accord = calculate_metrics(gt_mb, gt_nh)

                print(f"-> POOLED Linear Weighted Kappa : {gt_kappa:.4f}")
                print(f"-> POOLED Percent Agreement     : {gt_accord:.2f}%")
                print(f"-> TOTAL Valid Domain Ratings   : {len(gt_mb)}")

                # Standard interpretation of Kappa
                if gt_kappa < 0:
                    verdict = "No agreement"
                elif gt_kappa <= 0.20:
                    verdict = "Slight"
                elif gt_kappa <= 0.40:
                    verdict = "Fair"
                elif gt_kappa <= 0.60:
                    verdict = "Moderate"
                elif gt_kappa <= 0.80:
                    verdict = "Substantial"
                else:
                    verdict = "Almost perfect"

                print(f"-> Interpretation               : {verdict}")
            else:
                print("No data found anywhere.")

        except Exception as e:
            print(f"Critical Error: {e}")
    else:
        print(f"File not found at: {file_path}")