import pandas as pd
import plotly.graph_objects as go

# 1. Chargement du fichier Excel
# Note : Assurez-vous que le chemin est correct sur votre ordinateur
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Task_usage.xlsx"
xls = pd.ExcelFile(file_path)
output_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Plot\Sankey_task_usage.png"

# Liste des catégories cibles et attribution des couleurs
# Vous pouvez changer les codes HEX pour d'autres couleurs
category_settings = {
    "Monitoring": "#1f77b4",  # Bleu
    "Assessment": "#ff7f0e",  # Orange
    "Daily life improvement": "#2ca02c",  # Vert
    "Rehabilitation": "#d62728",  # Rouge
    #"Ecology": "#9467bd"  # Violet
}

categories = list(category_settings.keys())
data_links = []

# 2. Parcours de chaque feuille
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name)

    for cat in categories:
        if cat in df.columns:
            # Compte les 'X' (insensible à la casse et aux espaces)
            count = df[df[cat].astype(str).str.upper().str.strip() == 'X'].shape[0]

            if count > 0:
                data_links.append({
                    'source': sheet_name,
                    'target': cat,
                    'value': count
                })

if not data_links:
    print("Aucune donnée 'X' trouvée.")
else:
    # 3. Préparation des nœuds et des indices
    sources_unique = sorted(list(set(d['source'] for d in data_links)))
    targets_unique = categories  # On garde l'ordre défini au début

    all_nodes = sources_unique + targets_unique
    node_indices = {name: i for i, name in enumerate(all_nodes)}

    # Couleurs des nœuds
    node_colors = []
    for node in all_nodes:
        if node in category_settings:
            node_colors.append(category_settings[node])  # Couleur de l'usage
        else:
            node_colors.append("#A9A9A9")  # Gris pour les tâches (sources)


    # 4. Préparation des liens et de leurs couleurs
    # Fonction pour transformer HEX en RGBA (pour la transparence des liens)
    def hex_to_rgba(hex_code, opacity=0.3):
        hex_code = hex_code.lstrip('#')
        rgb = tuple(int(hex_code[i:i + 2], 16) for i in (0, 2, 4))
        return f'rgba({rgb[0]}, {rgb[1]}, {rgb[2]}, {opacity})'


    sources = [node_indices[d['source']] for d in data_links]
    targets = [node_indices[d['target']] for d in data_links]
    values = [d['value'] for d in data_links]

    # La couleur du lien correspond à la couleur de la destination (target)
    link_colors = [hex_to_rgba(category_settings[d['target']], 0.3) for d in data_links]

    # 5. Création du diagramme de Sankey
    fig = go.Figure(data=[go.Sankey(
        node=dict(
            pad=20,
            thickness=20,
            line=dict(color="white", width=1),
            label=all_nodes,
            color=node_colors
        ),
        link=dict(
            source=sources,
            target=targets,
            value=values,
            color=link_colors
        )
    )])

    fig.update_layout(
        width=1000,
        height=700
    )
    fig.write_image(output_path, scale=2)  # scale=2 pour une meilleure netteté

    fig.show()
