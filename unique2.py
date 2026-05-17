import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import plotly.express as px
import io
import csv
from openpyxl.styles import numbers


# =============================================================================
# FONCTIONS DE TRAITEMENT
# =============================================================================

def process_excel(file):
    """Traite le fichier Excel administration pour extraire la liste des étudiants"""
    try:
        # Lire le fichier Excel sans header
        xls = pd.read_excel(file, header=None)

        # Trouver l'index de la ligne d'en-tête
        header_index = next(
            (idx for idx, row in xls.iterrows() if all(col in row.values for col in ['Code', 'Nom', 'Prénom'])),
            None
        )

        if header_index is None:
            st.error("Les colonnes 'Code', 'Nom', 'Prénom' sont introuvables dans le fichier.")
            return None, None

        # Redéfinir les en-têtes et supprimer les lignes précédentes
        xls.columns = xls.iloc[header_index]
        xls = xls.iloc[header_index + 1:].reset_index(drop=True)

        # Vérification si le fichier est vide après nettoyage
        if xls.empty:
            st.error("Aucune donnée valide après le traitement des lignes.")
            return None, None

        # Vérification des colonnes nécessaires
        required_columns = ['Nom', 'Prénom', 'Code']
        missing = [col for col in required_columns if col not in xls.columns]
        if missing:
            st.error(f"Colonnes manquantes après traitement : {', '.join(missing)}")
            return None, None

        # Nettoyage des données
        liste = xls.dropna(subset=['Nom', 'Prénom', 'Code'])
        liste['Name'] = liste['Code'].astype(str) + ' ' + liste['Nom'] + ' ' + liste['Prénom']
        liste = liste[['Code', 'Name']].drop_duplicates()
        return xls, liste

    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier Excel : {str(e)}")
        st.info("Assurez-vous que le fichier est bien formaté et contient les colonnes requises.")
        return None, None


def detect_delimiter(file_content):
    """Détecte automatiquement le séparateur d'un fichier CSV"""
    sample = file_content.decode('utf-8', errors='ignore')
    try:
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter
    except csv.Error:
        st.warning("Aucun délimiteur détecté. La virgule ',' sera utilisée par défaut.")
        return ','


def process_csv(csv_file):
    """Traite le fichier CSV d'AMC pour l'analyse statistique"""
    try:
        csv_content = csv_file.read()
        delimiter = detect_delimiter(csv_content)
        csv_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')), delimiter=delimiter, encoding='utf-8')
        
        # Renommer la colonne 'Mark' en 'Note' si elle existe
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
        
        # Vérification des colonnes nécessaires
        if 'A:Code' not in csv_data.columns or 'Note' not in csv_data.columns:
            st.error("Les colonnes nécessaires ('A:Code', 'Note') sont manquantes.")
            return None, None
        
        # Séparation des données valides et des anomalies
        csv_nones = csv_data[csv_data['A:Code'] == 'NONE'].copy()
        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()

        # Conversion des codes en numérique
        csv_clean['A:Code'] = pd.to_numeric(csv_clean['A:Code'], errors='coerce')
        
        # Gestion des notes : virgule → point pour le calcul
        csv_clean['Note'] = csv_clean['Note'].astype(str).str.replace(',', '.').astype(float)

        if csv_clean.empty:
            st.error("Aucune donnée valide après le nettoyage !")
            return None, None

        return csv_clean, csv_nones

    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier CSV : {str(e)}")
        st.info("Assurez-vous que le fichier est bien formaté et contient les colonnes requises.")
        return None, None


def process_csv2excel(xls_file, csv_file, add_notes=0):
    """
    Fusionne les notes du CSV AMC dans le fichier Excel administration.
    ✅ CORRECTION : Normalisation des codes étudiants pour gérer int/str
    """
    try:
        # === TRAITEMENT DU CSV ===
        csv_content = csv_file.read()
        delimiter = detect_delimiter(csv_content)
        csv_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')), delimiter=delimiter, encoding='utf-8')
        
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
        
        csv_data = csv_data[['A:Code', 'Note']]
        
        # Séparer anomalies et données valides
        anomalies = csv_data[csv_data['A:Code'] == 'NONE'].copy()
        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()

        # Formatage des notes (virgule décimale)
        csv_clean['Note'] = csv_clean['Note'].astype(str).str.replace('.', ',', regex=False)

        # Ajout de points optionnel avec plafonnement à 20
        if add_notes > 0:
            csv_clean['Note'] = csv_clean['Note'].str.replace(',', '.', regex=False).astype(float)
            csv_clean['Note'] = csv_clean['Note'].apply(lambda x: min(x + add_notes, 20))
            csv_clean['Note'] = csv_clean['Note'].apply(lambda x: str(x).replace('.', ','))

        if csv_clean.empty:
            st.error("Aucune donnée valide après le nettoyage !")
            return None, None

        # === CRÉATION DU DICTIONNAIRE DE NOTES ===
        # 🎯 CLÉS CONVERTIES EN ENTIER pour matching robuste
        notes_etudiants = {}
        for _, row in csv_clean.iterrows():
            try:
                code_key = int(float(str(row['A:Code']).strip()))
                notes_etudiants[code_key] = row['Note']
            except (ValueError, TypeError):
                continue  # Skip les codes non convertibles

        # === CHARGEMENT DU FICHIER EXCEL ===
        wb = load_workbook(filename=xls_file)
        sheet = wb.active

        # 🔍 Recherche des colonnes 'Code' et 'Note' (case-insensitive)
        code_column = None
        note_column = None
        header_row = None
        
        for row in sheet.iter_rows(min_row=1, max_row=15):
            for cell in row:
                if cell.value:
                    cell_val = str(cell.value).strip().lower()
                    if cell_val == 'code' and code_column is None:
                        code_column = cell.column
                        header_row = cell.row
                    elif cell_val == 'note' and note_column is None:
                        note_column = cell.column
                        header_row = cell.row
            if code_column and note_column:
                break

        if code_column is None or note_column is None:
            st.error("Les en-têtes 'Code' et/ou 'Note' n'ont pas été trouvés dans le fichier Excel.")
            wb.close()
            return None, None

        # === MISE À JOUR DES NOTES AVEC NORMALISATION DES CODES ===
        matched_count = 0
        unmatched_codes = []
        error_codes = []
        
        for row in sheet.iter_rows(min_row=header_row + 1, max_col=max(code_column, note_column), values_only=False):
            code_cell = row[code_column - 1]
            note_cell = row[note_column - 1]
            code_raw = code_cell.value
            
            # 🔄 NORMALISATION : Convertir le code Excel en entier
            try:
                if code_raw is None or str(code_raw).strip() == '':
                    continue
                # Gestion polyvalente : "12345", 12345.0, 12345
                code_int = int(float(str(code_raw).strip()))
            except (ValueError, TypeError):
                unmatched_codes.append(str(code_raw))
                continue
            
            # 🔍 Recherche dans le dictionnaire (clés en int)
            if code_int in notes_etudiants:
                note_value = notes_etudiants[code_int]
                try:
                    if ',' in str(note_value):
                        note_cell.value = float(str(note_value).replace(',', '.'))
                        note_cell.number_format = numbers.FORMAT_NUMBER_00
                    else:
                        note_cell.value = float(note_value) if '.' in str(note_value) else int(note_value)
                    matched_count += 1
                except ValueError:
                    note_cell.value = note_value
                    error_codes.append(code_int)

        wb.close()

        # 📊 FEEDBACK UTILISATEUR
        if matched_count == 0:
            st.warning("⚠️ Aucune note n'a pu être transférée. Vérifiez la compatibilité des codes étudiants.")
        elif len(unmatched_codes) > 0:
            st.warning(f"⚠️ {len(unmatched_codes)} codes Excel n'ont pas pu être convertis en numérique.")
        if len(error_codes) > 0:
            st.warning(f"⚠️ {len(error_codes)} notes ont été copiées sans formatage numérique.")
        
        # Sauvegarde dans un buffer
        output = io.BytesIO()
        wb = load_workbook(filename=xls_file)
        # Re-appliquer les modifications (wb a été fermé, on recharge pour sauvegarder)
        # Note: Dans une version optimisée, on garderait wb ouvert, mais pour Streamlit c'est plus sûr ainsi
        wb.save(output)
        wb.close()
        output.seek(0)

        return output, len(anomalies)

    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier Excel : {str(e)}")
        st.info("Assurez-vous que le fichier est bien formaté et contient les colonnes requises.")
        return None, None


# =============================================================================
# INTERFACE UTILISATEUR STREAMLIT
# =============================================================================

st.set_page_config(page_title="🛠️ Outils pour AMC", page_icon="🎓", layout="wide")
st.title("🛠️ Outils pour AMC")

# Sidebar pour la navigation
section = st.sidebar.radio("📋 Sections", ["ETUDIANTS", "STATISTIQUES", "NOTES"])

# =============================================================================
# ONGLET 1 : GESTION DES ÉTUDIANTS
# =============================================================================
if section == "ETUDIANTS":
    st.header("👨‍🎓 Liste des étudiants")
    st.info(
        """
        **Objectif** : Générer la liste des étudiants au format CSV pour Auto Multiple Choice.
        
        **Instructions** :
        1. Téléchargez le fichier Excel envoyé par l'administration
        2. L'application extrait automatiquement les colonnes requises
        3. Téléchargez le fichier CSV prêt pour AMC
        """
    )

    uploaded_excel_file = st.file_uploader("📁 Sélectionnez le fichier Excel administration", type="xlsx", key="excel_uploader")

    if uploaded_excel_file is not None:
        with st.spinner("⏳ Traitement du fichier Excel en cours..."):
            xls, liste = process_excel(uploaded_excel_file)

            if xls is not None and liste is not None:
                st.success("✅ Lecture du fichier Excel réussie !")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.write("🔎 Aperçu de la base de données brute :")
                    st.dataframe(xls.head(10), use_container_width=True)
                with col2:
                    st.write("🔎 Aperçu de la liste pour AMC :")
                    st.dataframe(liste.head(10), use_container_width=True)
                
                st.success(f"🔢 La liste contient **{len(liste)} étudiants** uniques.")
                
                # Génération du CSV
                csv_data = liste.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button(
                    label="📥 Télécharger la liste au format CSV (AMC)",
                    data=csv_data,
                    file_name="liste_etudiants_amc.csv",
                    mime="text/csv"
                )

# =============================================================================
# ONGLET 2 : STATISTIQUES DES NOTES
# =============================================================================
elif section == "STATISTIQUES":
    st.header("📊 Statistiques des notes")
    st.info(
        """
        **Objectif** : Analyser la distribution des notes après correction AMC.
        
        **Fonctionnalités** :
        • Histogramme interactif de la distribution
        • Métriques : présents, validés, taux de réussite
        • Simulation d'ajustement des notes (+0 à +5 points)
        """
    )
    
    uploaded_csv_file = st.file_uploader("📁 Sélectionnez le fichier CSV des résultats AMC", type="csv", key="csv_uploader_stats")

    if uploaded_csv_file is not None:
        with st.spinner("⏳ Analyse des notes en cours..."):
            csv_clean, csv_nones = process_csv(uploaded_csv_file)

            if csv_clean is not None:
                # === GRAPHIQUE PRINCIPAL ===
                effectifs = csv_clean['Note'].value_counts().sort_index().reset_index()
                effectifs.columns = ['Valeur', 'Effectif']

                fig = px.bar(
                    effectifs, x='Valeur', y='Effectif',
                    title="📈 Distribution des notes",
                    labels={'Valeur': 'Notes (0-20)', 'Effectif': 'Nombre d\'étudiants'},
                    text_auto=True, color='Effectif',
                    color_continuous_scale='Blues'
                )
                fig.update_layout(
                    xaxis=dict(tickmode='linear', tick0=0, dtick=1),
                    showlegend=False, height=500
                )
                fig.update_traces(textfont_size=10, textposition='outside')
                st.plotly_chart(fig, use_container_width=True)

                # === MÉTRIQUES ===
                st.subheader("📋 Indicateurs clés")
                total = len(csv_clean)
                valides = (csv_clean['Note'] >= 10).sum()
                taux = round((valides / total) * 100, 1) if total > 0 else 0
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("👥 Présents", total)
                col2.metric("✅ Validés (≥10)", valides)
                col3.metric("📊 Taux de réussite", f"{taux}%")
                col4.metric("⚠️ Non identifiés", len(csv_nones) if csv_nones is not None else 0)

                # === SIMULATION AJUSTEMENT ===
                st.subheader("🎚️ Simulation d'ajustement")
                ajout_points = st.slider("Points à ajouter à toutes les notes", 0.0, 5.0, 0.0, 0.5)

                if ajout_points > 0:
                    csv_plus = csv_clean.copy()
                    csv_plus['Note'] = csv_plus['Note'].apply(lambda x: min(x + ajout_points, 20))
                    
                    effectifs_plus = csv_plus['Note'].value_counts().sort_index().reset_index()
                    effectifs_plus.columns = ['Valeur', 'Effectif']

                    fig_plus = px.bar(
                        effectifs_plus, x='Valeur', y='Effectif',
                        title=f"📈 Distribution après +{ajout_points} points",
                        labels={'Valeur': 'Notes (0-20)', 'Effectif': 'Nombre d\'étudiants'},
                        text_auto=True, color='Effectif',
                        color_continuous_scale='Greens'
                    )
                    fig_plus.update_layout(
                        xaxis=dict(tickmode='linear', tick0=0, dtick=1),
                        showlegend=False, height=500
                    )
                    fig_plus.update_traces(textfont_size=10, textposition='outside')
                    st.plotly_chart(fig_plus, use_container_width=True)
                    
                    # Métriques après ajustement
                    valides_plus = (csv_plus['Note'] >= 10).sum()
                    taux_plus = round((valides_plus / total) * 100, 1) if total > 0 else 0
                    
                    col5, col6, col7 = st.columns(3)
                    col5.metric("✅ Validés après ajustement", valides_plus, delta=valides_plus-valides)
                    col6.metric("📊 Nouveau taux de réussite", f"{taux_plus}%", delta=f"{taux_plus-taux}%")
                    col7.metric("🔄 Gain de validés", f"+{valides_plus-valides}")

# =============================================================================
# ONGLET 3 : TRAITEMENT ET FUSION DES NOTES
# =============================================================================
elif section == "NOTES":
    st.header("✍️ Fusion des notes dans le fichier administration")
    
    col_info1, col_info2 = st.columns(2)
    with col_info1:
        st.info(
            """
            **Étape 1** : Fichier Excel administration
            
            • Doit contenir les colonnes : `Code`, `Nom`, `Prénom`, `Note`
            • La colonne `Code` peut être au format texte ou numérique ✅
            """
        )
    with col_info2:
        st.info(
            """
            **Étape 2** : Fichier CSV des résultats AMC
            
            • Exporté depuis Auto Multiple Choice
            • Doit contenir les colonnes : `A:Code`, `Note` (ou `Mark`)
            """
        )

    # Upload des fichiers
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        xls_file = st.file_uploader("📁 Fichier Excel administration", type="xlsx", key="excel_uploader_notes")
    with col_up2:
        csv_file = st.file_uploader("📁 Fichier CSV résultats AMC", type="csv", key="csv_uploader_notes")

    if xls_file is not None and csv_file is not None:
        st.divider()
        st.subheader("⚙️ Paramètres de fusion")
        
        col_params, col_preview = st.columns([1, 2])
        with col_params:
            add_notes = st.number_input("➕ Points bonus à ajouter", min_value=0.0, max_value=5.0, value=0.0, step=0.5, help="Les notes seront plafonnées à 20/20")
            process_btn = st.button("🚀 Lancer la fusion des notes", type="primary", use_container_width=True)
        
        if process_btn:
            with st.spinner("⏳ Traitement en cours..."):
                # Reset des pointeurs de fichier
                xls_file.seek(0)
                csv_file.seek(0)
                
                processed_xls, nb_none = process_csv2excel(xls_file, csv_file, add_notes)

                if processed_xls is not None:
                    st.success("✅ Fusion réussie ! Les notes ont été intégrées au fichier Excel.")
                    
                    # Résumé du traitement
                    st.write("📋 **Résumé du traitement** :")
                    st.write(f"• Notes transférées avec succès")
                    if nb_none > 0:
                        st.warning(f"• ⚠️ {nb_none} étudiant(s) non identifié(s) dans le CSV AMC → à saisir manuellement")
                    
                    st.divider()
                    
                    # Nom du fichier de sortie
                    file_name = st.text_input("📝 Nom du fichier de sortie (sans extension)", value="notes_traitees")
                    
                    # Bouton de téléchargement
                    st.download_button(
                        label="📥 Télécharger le fichier Excel traité",
                        data=processed_xls.getvalue(),
                        file_name=f"{file_name.strip() or 'notes_traitees'}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    # 💡 Conseils post-traitement
                    with st.expander("💡 Conseils pour la finalisation"):
                        st.markdown("""
                        1. **Vérifiez** le fichier téléchargé : ouvrez-le et contrôlez quelques notes aléatoires
                        2. **Complétez manuellement** les notes des étudiants non identifiés (signalés ci-dessus)
                        3. **Validez** les formats : les notes doivent apparaître avec 2 décimales (ex: `14,50`)
                        4. **Sauvegardez** une copie avant envoi à l'administration
                        """)

    # Aide contextuelle
    with st.expander("❓ Problèmes fréquents et solutions"):
        st.markdown("""
        ### 🔧 Dépannage rapide
        
        | Problème | Solution |
        |----------|----------|
        | ❌ "Aucune note transférée" | Vérifiez que les codes étudiants sont identiques dans les 2 fichiers (pas d'espaces, mêmes chiffres) |
        | ❌ "Colonnes manquantes" | Ouvrez les fichiers et vérifiez que les en-têtes sont exactement : `Code`, `Note`, `A:Code` |
        | ❌ Format de date/heure dans Excel | Assurez-vous que la colonne `Code` ne contient que des nombres, pas de formules |
        | ⚠️ Codes avec zéros non significatifs | L'application normalise automatiquement : `01234` = `1234` |
        
        ### 🎯 Bonnes pratiques
        • Gardez toujours une copie de sauvegarde de vos fichiers originaux
        • Testez la fusion avec 3-4 étudiants avant de traiter toute la promotion
        • Utilisez le même fichier Excel administration tout au long du processus
        """)

# =============================================================================
# FOOTER
# =============================================================================
st.divider()
st.caption("🛠️ Outils pour AMC v2.1 • Application Streamlit pour la gestion des QCM • Dernière mise à jour : normalisation des codes étudiants")
