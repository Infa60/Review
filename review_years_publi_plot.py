import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os

# --- CONFIGURATION ---
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
sheet_name = "Global_overview"
save_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Reports_per_years_plot.png"


def plot_publication_trend():
    if not os.path.exists(file_path):
        print(f"Erreur : Le fichier est introuvable à l'adresse : {file_path}")
        return

    try:
        print("Chargement des données...")
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Normalisation des colonnes
        df.columns = [c.lower().strip() for c in df.columns]

        if 'year' not in df.columns:
            print("Erreur : Colonne 'year' introuvable.")
            return

        # Nettoyage des années
        years = pd.to_numeric(df['year'], errors='coerce').dropna().astype(int)

        if len(years) == 0:
            print("Aucune année valide trouvée.")
            return

        # Compte par année
        year_counts = years.value_counts().sort_index()

        # --- ASTUCE : REMPLISSAGE DES ANNÉES MANQUANTES ---
        # On crée une plage complète de la première à la dernière année
        # Cela permet d'avoir des 0 pour les années sans publication
        full_range = range(year_counts.index.min(), year_counts.index.max() + 1)
        year_counts = year_counts.reindex(full_range, fill_value=0)

        # --- GÉNÉRATION DU GRAPHIQUE ---
        print("Génération de la courbe...")
        fig, ax = plt.subplots(figsize=(6, 3))

        # Tracé de la courbe
        # marker='o' : met un point sur chaque année
        # linestyle='-' : relie les points
        ax.plot(year_counts.index, year_counts.values,
                marker='o', linestyle='-', linewidth=2, color='#4c72b0', label='Included studies')

        # Mise en forme
        plt.title('Timeline of included studies', fontsize=12, pad=15)
        plt.xlabel('Year', fontsize=10)
        plt.ylabel('Number of articles', fontsize=10)

        # Configuration de l'axe X (Années)
        # On force l'affichage de nombres entiers uniquement
        ax.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
        plt.xticks(rotation=45)

        # Configuration de l'axe Y (Nombre d'articles)
        # On force des entiers (pas de "1.5 article")
        ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))

        # Grille légère
        plt.grid(True, linestyle='--', alpha=0.6)


        # Remplissage sous la courbe (Optionnel, pour le style)
        ax.fill_between(year_counts.index, year_counts.values, color='#4c72b0', alpha=0.1)

        # Ajustement des marges
        plt.tight_layout()
        plt.savefig(save_path)

        plt.show()
        print("Graphique généré avec succès.")

    except Exception as e:
        print(f"Une erreur est survenue : {e}")


if __name__ == "__main__":
    plot_publication_trend()