import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import plotly.express as px
import io
import csv
from openpyxl.styles import numbers

# =============================================================================
# CONFIGURATION DE LA PAGE
# =============================================================================
st.set_page_config(page_title="🛠️ Outils pour AMC", page_icon="🎓", layout="wide")
st.title("🛠️ Outils pour AMC")

# Sidebar pour les sections
section = st.sidebar.radio(" ", ["ETUDIANTS", "STATISTIQUES", "NOTES"])

# =============================================================================
# FONCTIONS DE TRAITEMENT
# =============================================================================

def process_excel(file):
    """Traite le fichier Excel administration pour extraire la liste des étudiants"""
    try:
        xls = pd.read_excel(file, header=None)
        
        header_index = next(
            (idx for idx, row in xls.iterrows() if all(col in row.values for col in ['Code', 'Nom', 'Prénom'])),
            None
        )
        
        if header_index is None:
            st.error("Les colonnes 'Code', 'Nom', 'Prénom' sont introuvables dans le fichier.")
            return None, None
            
        xls.columns = xls.iloc[header_index]
        xls = xls.iloc[header_index + 1:].reset_index(drop=True)
        
        if xls.empty:
            st.error("Aucune donnée valide après le traitement des lignes.")
            return None, None
            
        required_columns = ['Nom', 'Prénom', 'Code']
        missing = [col for col in required_columns if col not in xls.columns]
        if missing:
            st.error(f"Colonnes manquantes après traitement : {', '.join(missing)}")
            return None, None
            
        liste = xls.dropna(subset=['Nom', 'Prénom', 'Code'])
        liste['Name'] = liste['Code'].astype(str) + ' ' + liste['Nom'] + ' ' + liste['Prénom']
        liste = liste[['Code', 'Name']].drop_duplicates()
        return xls, liste
        
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier Excel : {str(e)}")
        st.info("Assurez-vous que le fichier est bien formaté et contient les colonnes requises.")
        return None, None


def detect_delimiter(file_content):
    """Détecte le séparateur dans un fichier CSV"""
    sample = file_content.decode('utf-8', errors='ignore')
    try:
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter
    except csv.Error:
        st.warning("Aucun délimiteur détecté. La virgule ',' sera utilisée par défaut.")
        return ','


def process_csv(csv_file):
    """Traite le fichier CSV pour l'analyse statistique"""
    try:
        csv_content = csv_file.read()
        delimiter = detect_delimiter(csv_content)
        csv_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')), delimiter=delimiter, encoding='utf-8')
        
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
            
        if 'A:Code' not in csv_data.columns or 'Note' not in csv_data.columns:
            st.error("Les colonnes nécessaires ('A:Code', 'Note') sont manquantes.")
            return None, None
            
        csv_nones = csv_data[csv_data['A:Code'] == 'NONE'].copy()
        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()
        
        csv_clean['A:Code'] = pd.to_numeric(csv_clean['A:Code'], errors='coerce')
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
    ✅ CORRECTION : Normalisation stricte des codes pour garantir la correspondance
    quel que soit le format (texte, entier, float, espaces).
    """
    try:
        # === 1. LECTURE ET NETTOYAGE DU CSV ===
        csv_content = csv_file.read()
        delimiter = detect_delimiter(csv_content)
        csv_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')), delimiter=delimiter, encoding='utf-8')
        
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
            
        csv_data = csv_data[['A:Code', 'Note']]
        anomalies = csv_data[csv_data['A:Code'] == 'NONE'].copy()
        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()
        
        # Gestion des notes (virgule/point)
        csv_clean['Note'] = csv_clean['Note'].astype(str).str.replace('.', ',', regex=False)
        
        if add_notes > 0:
            csv_clean['Note'] = csv_clean['Note'].str.replace(',', '.', regex=False).astype(float)
            csv_clean['Note'] = csv_clean['Note'].apply(lambda x: min(x + add_notes, 20))
            csv_clean['Note'] = csv_clean['Note'].apply(lambda x: str(x).replace('.', ','))
            
        if csv_clean.empty:
            st.error("Aucune donnée valide après le nettoyage !")
            return None, None

        # === 2. CRÉATION DU DICTIONNAIRE DE NOTES ===
        # On normalise TOUTES les clés en CHAÎNES DE CARACTÈRES sans espaces
        notes_dict = {}
        for _, row in csv_clean.iterrows():
            code_key = str(row['A:Code']).strip()
            notes_dict[code_key] = row['Note']

        # === 3. OUVERTURE ET PARCOURS DE L'EXCEL ===
        wb = load_workbook(filename=xls_file)
        sheet = wb.active
        
        code_col = None
        note_col = None
        header_row = None
        
        # Recherche dynamique des en-têtes (max 15 premières lignes)
        for i, row in enumerate(sheet.iter_rows(min_row=1, max_row=15), 1):
            for cell in row:
                if cell.value:
                    val = str(cell.value).strip().lower()
                    if val == 'code' and code_col is None:
                        code_col = cell.column
                        header_row = i
                    elif val == 'note' and note_col is None:
                        note_col = cell.column
            if code_col and note_col:
                break
                
        if code_col is None or note_col is None:
            st.error("Les colonnes 'Code' et/ou 'Note' sont introuvables dans l'Excel.")
            wb.close()
            return None, None

        # === 4. TRANSFERT DES NOTES ===
        matched = 0
        for row in sheet.iter_rows(min_row=header_row + 1, max_col=max(code_col, note_col)):
            code_cell = row[code_col - 1]
            note_cell = row[note_col - 1]
            
            # Normalisation du code Excel exactement comme dans le dictionnaire
            if code_cell.value is not None:
                excel_code = str(code_cell.value).strip()
                
                if excel_code in notes_dict:
                    note_val = notes_dict[excel_code]
                    try:
                        if ',' in str(note_val):
                            note_cell.value = float(str(note_val).replace(',', '.'))
                            note_cell.number_format = numbers.FORMAT_NUMBER_00
                        else:
                            note_cell.value = float(note_val) if '.' in str(note_val) else int(note_val)
                        matched += 1
                    except ValueError:
                        note_cell.value = note_val  # Fallback si conversion échoue
                        
        wb.close()
        
        # Feedback
        if matched == 0:
            st.warning("⚠️ Aucune note n'a été transférée. Vérifiez la correspondance des codes.")
        else:
            st.success(f"✅ {matched} notes transférées avec succès.")
            
        # Sauvegarde
        output = io.BytesIO()
        wb = load_workbook(filename=xls_file)
        wb.save(output)
        wb.close()
        output.seek(0)
        
        return output, len(anomalies)

    except Exception as e:
        st.error(f"Erreur critique lors du traitement : {str(e)}")
        return None, None

# =============================================================================
# INTERFACE UTILISATEUR
# =============================================================================

if section == "ETUDIANTS":
    st.header("👨‍ Liste des étudiants")
    st.info("Pour générer la liste des étudiants au format CSV à charger dans Auto Multiple Choice, téléchargez le fichier Excel envoyé par l'administration.")

    uploaded_excel_file = st.file_uploader(" ", type="xlsx", key="excel_uploader")

    if uploaded_excel_file is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            xls, liste = process_excel(uploaded_excel_file)

            if xls is not None:
                st.success(f"✅  Lecture du fichier Excel réussie !")
                st.write("🔎 Aperçu de la base de données des étudiants:")
                st.write(xls.head(10))
                st.write("🔎 Aperçu de la liste des étudiants à fournir à Auto Multiple Choice:")
                st.write(liste.head(10))
                st.success(f"🔢 La liste contient {len(xls)} étudiants.")
                
                csv_data = liste.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Télécharger la liste des étudiants au format CSV",
                    data=csv_data,
                    file_name="liste.csv",
                    mime="text/csv"
                )

elif section == "STATISTIQUES":
    st.header("📊 Statistiques des notes")
    st.info("Pour avoir une idée sur la distribution statistique des notes, télécharger le fichier des notes calculées par Auto Multiple Choice au format CSV.")
    
    uploaded_csv_file = st.file_uploader(" ", type="csv", key="csv_uploader")

    if uploaded_csv_file is not None:
        with st.spinner("Intégration des notes aux étudiants..."):
            csv_clean, csv_nones = process_csv(uploaded_csv_file)

            if csv_clean is not None:
                # Calcul des effectifs
                effectifs = csv_clean['Note'].value_counts().reset_index()
                modalites = csv_clean['Note'].unique()
                effectifs.columns = ['Valeur', 'Effectif']

                fig = px.bar(effectifs,
                             x='Valeur',
                             y='Effectif',
                             title=" ",
                             labels={'Valeur': 'Notes', 'Effectif': 'Effectifs'},
                             text_auto=True
                             )
                fig.update_layout(title_font_size=20, xaxis_title_font=dict(size=14), yaxis_title_font=dict(size=14), showlegend=False)
                fig.update_traces(textfont_size=14, textangle=0, textposition="outside", width=0.5)
                fig.update_xaxes(tickmode='array', tickvals=list(range(21)), ticktext=[str(i) for i in range(21)])
                fig.update_layout(width=800, height=600)
                st.plotly_chart(fig)

                st.info("Quelques statistiques des notes.")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Présents", len(csv_clean) if csv_clean is not None else 0)
                with col2:
                    st.metric("Validés", (csv_clean['Note'] >= 10).sum() if csv_clean is not None else 0)
                with col3:
                    st.metric("Taux de réussite (%)", round(((csv_clean['Note'] >= 10).sum() / len(csv_clean)) * 100, 2) if csv_clean is not None else 0)
                with col4:
                    st.metric("Mal identifiés", len(csv_nones) if csv_nones is not None else 0)

                ajout_points = st.slider("Ajouter des points", min_value=0.0, max_value=5.0, value=0.0, step=0.5)

                if ajout_points > 0:
                    csv_plus = csv_clean.copy()
                    st.info("Simulation de la distribution des notes et du taux de réussite après ajout de points.")
                    csv_plus['Note'] = csv_plus['Note'].apply(lambda x: min(x + ajout_points, 20))
                    
                    effectifs_plus = csv_plus['Note'].value_counts().reset_index()
                    effectifs_plus.columns = ['Valeur', 'Effectif']
                    
                    fig_plus = px.bar(effectifs_plus, x='Valeur', y='Effectif', title=" ", labels={'Valeur': 'Notes', 'Effectif': 'Effectifs'}, text_auto=True)
                    fig_plus.update_layout(title_font_size=20, xaxis_title_font=dict(size=14), yaxis_title_font=dict(size=14), showlegend=False)
                    fig_plus.update_traces(textfont_size=14, textangle=0, textposition="outside", width=0.5)
                    fig_plus.update_xaxes(tickmode='array', tickvals=list(range(21)), ticktext=[str(i) for i in range(21)])
                    fig_plus.update_layout(width=800, height=600)
                    
                    col5, col6, col7, col8 = st.columns(4)
                    with col5:
                        st.metric("Présents", len(csv_plus) if csv_plus is not None else 0)
                    with col6:
                        st.metric("Validés", (csv_plus['Note'] >= 10).sum() if csv_plus is not None else 0)
                    with col7:
                        st.metric("Taux de réussite (%)", round(((csv_plus['Note'] >= 10).sum() / len(csv_plus)) * 100, 2) if csv_plus is not None else 0)
                    with col8:
                        st.metric("Mal identifiés", len(csv_nones) if csv_nones is not None else 0)
                        
                    st.plotly_chart(fig_plus)

elif section == "NOTES":
    st.header("✍️ Traitement des notes")
    st.info("Télécharger le fichier Excel envoyé par l'administration pour la saisie des notes.")
    xls_file = st.file_uploader(" ", type="xlsx", key="excel_uploader2")

    st.info("Télécharger le fichier des notes calculées par Auto Multiple Choice au format CSV.")
    csv_file = st.file_uploader("", type="csv", key="csv_uploader")
    
    if xls_file is not None and csv_file is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            csv_content = io.BytesIO(csv_file.getvalue())
            xls_content = io.BytesIO(xls_file.getvalue())
            
            add_notes = st.number_input("Combien voulez-vous ajouter de points à l'ensemble des étudiants?", step=0.5)
            processed_xls, nb_none = process_csv2excel(xls_content, csv_content, add_notes)

        if processed_xls is not None:
            wb = load_workbook(processed_xls)
            output = io.BytesIO()
            wb.save(output)
            wb.close()
            output.seek(0)
            file_data = output.getvalue()

            st.success("✅ Félicitations ! Les notes ont été saisies avec succès.")
            file_name = st.text_input("Saisir le nom du fichier Excel de l'adminstration sans l'extension '.xlsx' puis valider.")
            st.download_button(
                label="📥 Télécharger le fichier Excel traité",
                data=file_data,
                file_name=file_name + ".xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        if nb_none is not None and nb_none > 0:
            st.warning(f"Attention! {nb_none} étudiants ont été mal identifiés. Vérifiez leurs copies et saisissez leurs notes manuellement.")
