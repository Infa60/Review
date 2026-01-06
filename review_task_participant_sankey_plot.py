import pandas as pd
import plotly.graph_objects as go


def main():
    # ==========================================
    # 1. CONFIGURATION
    # ==========================================
    file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
    sheet_name = 'Global_overview'

    L_TACHES = [
        'Sit-to-stand', 'Running', 'Cycling', 'Stair-negotiation',
        'Time-Up-and-Go', 'Obstacle-clearance', 'Game', 'One-leg-standing',
        'Jumping', 'Squat', 'Stepping-target', 'Hopping', 'Kicking-a-ball'
    ]

    L_GMFCS = ['GMFCS-I', 'GMFCS-II', 'GMFCS-III', 'GMFCS-IV']
    UNK_GMFCS = 'GMFCS Unknown'

    # ==========================================
    # 2. CHARGEMENT ET NETTOYAGE
    # ==========================================
    print(">>> Chargement des données...")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        df.columns = [str(c).strip() for c in df.columns]
    except Exception as e:
        print(f"ERREUR : {e}")
        return

    def clean(x):
        s = str(x).strip().upper()
        if s in ['X', 'YES', '1', '1.0', 'TRUE']: return 1
        try:
            return 1 if float(x) > 0 else 0
        except:
            return 0

    for col in L_TACHES + L_GMFCS:
        if col in df.columns:
            df[col] = df[col].apply(clean)
        else:
            df[col] = 0

    if 'GMFCS-I' in df.columns:
        df['is_gmfcs_unk'] = df['GMFCS-I'].apply(lambda x: 1 if str(x).strip() == '???' else 0)
    else:
        df['is_gmfcs_unk'] = 0

    # ==========================================
    # 3. CONSTRUCTION DES NOEUDS (COULEURS DISTINCTES)
    # ==========================================
    labels = []
    x_pos = []
    y_pos = []
    colors = []
    node_map = {}

    # --- TÂCHES (Bleu Acier neutre) ---
    for i, name in enumerate(L_TACHES):
        labels.append(name)
        x_pos.append(0.01)
        y_pos.append(0.01 + (i / (len(L_TACHES) - 1)) * 0.98)
        colors.append("rgba(70, 130, 180, 0.8)")
        node_map[name] = len(labels) - 1

    # --- GMFCS (Couleurs Distinctes) ---
    groupe_droite = L_GMFCS + [UNK_GMFCS]
    for i, name in enumerate(groupe_droite):
        labels.append(name)
        x_pos.append(0.99)
        y_pos.append(0.01 + (i / (len(groupe_droite) - 1)) * 0.98)

        # --- DEFINITION DES COULEURS ICI ---
        c = "rgba(160, 160, 160, 0.8)"  # Gris par défaut

        if "I" in name and "Unknown" not in name:
            c = "rgba(46, 204, 113, 0.8)"  # VERT (Emerald)

        if "II" in name:
            c = "rgba(52, 152, 219, 0.8)"  # BLEU (Peter River)

        if "III" in name:
            c = "rgba(243, 156, 18, 0.8)"  # ORANGE

        if "IV" in name:
            c = "rgba(231, 76, 60, 0.8)"  # ROUGE (Alizarin)

        if "Unknown" in name:
            c = "rgba(149, 165, 166, 0.8)"  # GRIS (Concrete)

        colors.append(c)
        node_map[name] = len(labels) - 1

    # ==========================================
    # 4. CRÉATION DES LIENS COLORES
    # ==========================================
    sources = []
    targets = []
    values = []
    link_colors = []

    def add_link(n1, n2):
        s, t = node_map[n1], node_map[n2]
        for k in range(len(sources)):
            if sources[k] == s and targets[k] == t:
                values[k] += 1
                return
        sources.append(s)
        targets.append(t)
        values.append(1)

        # On prend la couleur de la cible (ex: Rouge 0.8)
        target_c = colors[t]

        # On change juste l'opacité pour le lien (0.8 -> 0.3)
        # La méthode split/join est plus sûre que replace
        parts = target_c.split(',')
        parts[-1] = " 0.3)"  # Nouvelle opacité
        new_link_color = ",".join(parts)

        link_colors.append(new_link_color)

    print(">>> Création des liens...")

    for _, row in df.iterrows():
        tasks = [t for t in L_TACHES if row[t] == 1]
        if not tasks: continue

        gmfcs = [g for g in L_GMFCS if row[g] == 1]
        if row['is_gmfcs_unk'] == 1 or not gmfcs:
            gmfcs = [UNK_GMFCS]

        for t in tasks:
            for g in gmfcs:
                add_link(t, g)

    # ==========================================
    # 5. AFFICHAGE
    # ==========================================
    fig = go.Figure(data=[go.Sankey(
        arrangement="snap",
        node=dict(
            pad=15, thickness=20,
            line=dict(color="black", width=0.5),
            label=labels,
            color=colors,
            x=x_pos, y=y_pos
        ),
        link=dict(
            source=sources, target=targets, value=values,
            color=link_colors
        )
    )])

    fig.update_layout(title="Mapping Tasks -> GMFCS (Multi-Colors)", font_size=12, width=1100, height=800)

    fig.write_html("Sankey_MultiColors.html")
    try:
        fig.write_image("Sankey_MultiColors.png", scale=2)
        print("✅ Image PNG générée.")
    except:
        pass

    fig.show()


if __name__ == "__main__":
    main()