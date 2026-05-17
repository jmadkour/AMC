import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import plotly.express as px
import io
import csv
from openpyxl.styles import numbers

# =============================================================================
# CONFIGURATION
# =============================================================================
st.set_page_config(page_title="🛠️ Outils pour AMC", page_icon="🎓", layout="wide")
st.title("🛠️ Outils pour AMC")
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
            st.error("Les colonnes 'Code', 'Nom', 'Prénom' sont introuvables.")
            return None, None
            
        xls.columns = xls.iloc[header_index]
        xls = xls.iloc[header_index + 1:].reset_index(drop=True)
        
        if xls.empty:
            st.error("Aucune donnée valide après le traitement.")
            return None, None
            
        required_columns = ['Nom', 'Prénom', 'Code']
        missing = [col for col in required_columns if col not in xls.columns]
        if missing:
            st.error(f"Colonnes manquantes : {', '.join(missing)}")
            return None, None
            
        liste = xls.dropna(subset=['Nom', 'Prénom', 'Code'])
        liste['Name'] = liste['Code'].astype(str) + ' ' + liste['Nom'] + ' ' + liste['Prénom']
        liste = liste[['Code', 'Name']].drop_duplicates()
        return xls, liste
        
    except Exception as e:
        st.error(f"Erreur Excel : {str(e)}")
        return None, None


def detect_delimiter(file_content):
    sample = file_content.decode('utf-8', errors='ignore')
    try:
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter
    except csv.Error:
        return ','


def process_csv(csv_file):
    try:
        csv_content = csv_file.read()
        delimiter = detect_delimiter(csv_content)
        csv_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')), delimiter=delimiter, encoding='utf-8')
        
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
            
        if 'A:Code' not in csv_data.columns or 'Note' not in csv_data.columns:
            st.error("Colonnes 'A:Code' ou 'Note' manquantes dans le CSV.")
            return None, None
            
        csv_nones = csv_data[csv_data['A:Code'] == 'NONE'].copy()
        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()
        
        csv_clean['A:Code'] = csv_clean['A:Code'].astype(str).str.strip()
        csv_clean['Note'] = csv_clean['Note'].astype(str).str.replace(',', '.').astype(float)
        
        if csv_clean.empty:
            st.error("Aucune donnée valide.")
            return None, None
            
        return csv_clean, csv_nones
    except Exception as e:
        st.error(f"Erreur CSV : {str(e)}")
        return None, None


def process_csv2excel(xls_file, csv_file, add_notes=0):
    """
    Fusionne les notes du CSV vers l'Excel.
    ✅ CORRECTION : Normalisation stricte en STRING pour éviter les échecs int/float/text
    """
    try:
        # 1. Lecture et nettoyage du CSV
        csv_content = csv_file.read()
        delimiter = detect_delimiter(csv_content)
        csv_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')), delimiter=delimiter, encoding='utf-8')
        
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
            
        csv_data = csv_data[['A:Code', 'Note']]
        anomalies = csv_data[csv_data['A:Code'] == 'NONE'].copy()
        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()

        # Nettoyage des notes
        csv_clean['Note'] = csv_clean['Note'].astype(str).str.replace('.', ',', regex=False)
        if add_notes > 0:
            csv_clean['Note'] = csv_clean['Note'].str.replace(',', '.', regex=False).astype(float)
            csv_clean['Note'] = csv_clean['Note'].apply(lambda x: min(x + add_notes, 20))
            csv_clean['Note'] = csv_clean['Note'].apply(lambda x: str(x).replace('.', ','))

        if csv_clean.empty:
            st.error("Aucune note valide à transférer !")
            return None, None

        # 2. Création du dictionnaire de correspondance (CLÉS EN STRING)
        notes_dict = {}
        for _, row in csv_clean.iterrows():
            code_key = str(row['A:Code']).strip()
            notes_dict[code_key] = row['Note']

        # 3. Chargement et parcours de l'Excel
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
                
        if not code_col or not note_col:
            st.error("Impossible de trouver les colonnes 'Code' et 'Note' dans l'Excel.")
            wb.close()
            return None, None

        # 4. Transfert des notes avec correspondance STRING
        matched_count = 0
        for row in sheet.iter_rows(min_row=header_row + 1, max_col=max(code_col, note_col)):
            code_cell = row[code_col - 1]
            note_cell = row[note_col - 1]
            
            # Conversion sécurisée en string pour la comparaison
            excel_code = str(code_cell.value).strip() if code_cell.value else ""
            
            if excel_code in notes_dict:
                note_val = notes_dict[excel_code]
                try:
                    # Écriture propre dans Excel
                    if ',' in str(note_val):
                        note_cell.value = float(str(note_val).replace(',', '.'))
                        note_cell.number_format = numbers.FORMAT_NUMBER_00
                    else:
                        note_cell.value = float(note_val) if '.' in str(note_val) else int(note_val)
                    matched_count += 1
                except ValueError:
                    note_cell.value = note_val  # Fallback si conversion échoue

        wb.close()
        
        # 5. Feedback et sauvegarde
        if matched_count == 0:
            st.warning("⚠️ Aucune note transférée. Vérifiez que les codes correspondent exactement.")
        else:
            st.success(f"✅ {matched_count} notes transférées avec succès !")
            
        output = io.BytesIO()
        wb = load_workbook(filename=xls_file)
        wb.save(output)
        wb.close()
        output.seek(0)
        
        return output, len(anomalies)

    except Exception as e:
        st.error(f"Erreur critique : {str(e)}")
        return None, None


# =============================================================================
# INTERFACE UTILISATEUR
# =============================================================================

if section == "ETUDIANTS":
    st.header("👨‍ Liste des étudiants")
    st.info("Téléchargez le fichier Excel administration pour générer la liste CSV compatible AMC.")
    
    uploaded_excel = st.file_uploader(" ", type="xlsx", key="excel_uploader")
    
    if uploaded_excel is not None:
        with st.spinner("Traitement en cours..."):
            xls, liste = process_excel(uploaded_excel)
            
            if xls is not None:
                st.success(f"✅ {len(liste)} étudiants trouvés.")
                st.write("🔎 Aperçu des données brutes :")
                st.dataframe(xls.head(), use_container_width=True)
                
                st.write("🔎 Format final pour AMC :")
                st.dataframe(liste.head(), use_container_width=True)
                
                # ✅ BOUTON DE TÉLÉCHARGEMENT RESTAURÉ
                csv_data = liste.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button(
                    label="📥 Télécharger la liste au format CSV",
                    data=csv_data,
                    file_name="liste_etudiants_amc.csv",
                    mime="text/csv",
                    use_container_width=True
                )

elif section == "STATISTIQUES":
    st.header("📊 Statistiques des notes")
    st.info("Analysez la distribution des notes depuis le fichier CSV d'AMC.")
    
    uploaded_csv = st.file_uploader(" ", type="csv", key="csv_uploader_stats")
    
    if uploaded_csv is not None:
        with st.spinner("Analyse en cours..."):
            csv_clean, csv_nones = process_csv(uploaded_csv)
            
            if csv_clean is not None:
                effectifs = csv_clean['Note'].value_counts().sort_index().reset_index()
                effectifs.columns = ['Valeur', 'Effectif']
                
                fig = px.bar(effectifs, x='Valeur', y='Effectif', text_auto=True)
                fig.update_layout(xaxis=dict(tickmode='linear', tick0=0, dtick=1), height=500)
                st.plotly_chart(fig, use_container_width=True)
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Présents", len(csv_clean))
                col2.metric("Validés (≥10)", (csv_clean['Note'] >= 10).sum())
                col3.metric("Taux réussite", f"{round(((csv_clean['Note'] >= 10).sum() / len(csv_clean)) * 100, 1)}%")
                col4.metric("Mal identifiés", len(csv_nones) if csv_nones is not None else 0)
                
                ajout_points = st.slider("Ajouter des points", 0.0, 5.0, 0.0, 0.5)
                if ajout_points > 0:
                    csv_plus = csv_clean.copy()
                    csv_plus['Note'] = csv_plus['Note'].apply(lambda x: min(x + ajout_points, 20))
                    effectifs_plus = csv_plus['Note'].value_counts().sort_index().reset_index()
                    effectifs_plus.columns = ['Valeur', 'Effectif']
                    
                    fig_plus = px.bar(effectifs_plus, x='Valeur', y='Effectif', text_auto=True)
                    fig_plus.update_layout(xaxis=dict(tickmode='linear', tick0=0, dtick=1), height=500)
                    st.plotly_chart(fig_plus, use_container_width=True)
                    
                    col5, col6, col7 = st.columns(3)
                    col5.metric("Validés après ajout", (csv_plus['Note'] >= 10).sum())
                    col6.metric("Nouveau taux", f"{round(((csv_plus['Note'] >= 10).sum() / len(csv_plus)) * 100, 1)}%")
                    col7.metric("Gain", f"+{(csv_plus['Note'] >= 10).sum() - (csv_clean['Note'] >= 10).sum()}")

elif section == "NOTES":
    st.header("✍️ Transfert des notes vers l'Excel administration")
    
    col1, col2 = st.columns(2)
    with col1:
        st.info("📄 Fichier Excel administration (contient la colonne `Code`)")
        xls_file = st.file_uploader(" ", type="xlsx", key="excel_notes")
    with col2:
        st.info("📄 Fichier CSV résultats AMC (contient `A:Code` et `Note`)")
        csv_file = st.file_uploader(" ", type="csv", key="csv_notes")
        
    if xls_file is not None and csv_file is not None:
        add_notes = st.number_input("Points bonus à ajouter", min_value=0.0, max_value=5.0, value=0.0, step=0.5)
        
        if st.button("🚀 Lancer le transfert", type="primary", use_container_width=True):
            with st.spinner("Fusion en cours..."):
                xls_file.seek(0)
                csv_file.seek(0)
                
                processed_xls, nb_none = process_csv2excel(xls_file, csv_file, add_notes)
                
                if processed_xls is not None:
                    file_name = st.text_input("Nom du fichier de sortie (sans extension)", value="notes_traitees")
                    st.download_button(
                        label="📥 Télécharger le fichier Excel traité",
                        data=processed_xls.getvalue(),
                        file_name=f"{file_name.strip() or 'notes_traitees'}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    if nb_none > 0:
                        st.warning(f"⚠️ {nb_none} étudiant(s) non identifié(s) (Code = NONE) → à saisir manuellement.")
