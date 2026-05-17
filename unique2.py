import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import plotly.express as px
import io
import csv
from openpyxl.styles import numbers

# =============================================================================
# CONFIGURATION DE PAGE
# =============================================================================
st.set_page_config(
    page_title="🛠️ Outils pour AMC", 
    page_icon="🎓", 
    layout="wide",
    initial_sidebar_state="expanded"
)

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
            st.error("Aucune donnée valide.")
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
    sample = file_content.decode('utf-8', errors='ignore')[:10000]
    try:
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter
    except csv.Error:
        return ','


def process_csv(csv_file):
    """Traite le CSV pour statistiques"""
    try:
        csv_content = csv_file.read()
        delimiter = detect_delimiter(csv_content)
        csv_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')), delimiter=delimiter, encoding='utf-8')
        
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
        
        if 'A:Code' not in csv_data.columns or 'Note' not in csv_data.columns:
            st.error("Colonnes manquantes ('A:Code', 'Note').")
            return None, None
        
        csv_nones = csv_data[csv_data['A:Code'] == 'NONE'].copy()
        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()
        
        csv_clean['A:Code'] = pd.to_numeric(csv_clean['A:Code'], errors='coerce')
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
    🎯 FONCTION CORRIGÉE : Fusionne les notes avec normalisation des codes
    ✅ Résout le problème int/str pour les codes étudiants
    """
    try:
        # === TRAITEMENT CSV ===
        csv_content = csv_file.read()
        delimiter = detect_delimiter(csv_content)
        csv_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')), delimiter=delimiter, encoding='utf-8')
        
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
        
        csv_data = csv_data[['A:Code', 'Note']]
        anomalies = csv_data[csv_data['A:Code'] == 'NONE'].copy()
        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()
        
        csv_clean['Note'] = csv_clean['Note'].astype(str).str.replace('.', ',', regex=False)
        
        if add_notes > 0:
            csv_clean['Note'] = csv_clean['Note'].str.replace(',', '.', regex=False).astype(float)
            csv_clean['Note'] = csv_clean['Note'].apply(lambda x: min(x + add_notes, 20))
            csv_clean['Note'] = csv_clean['Note'].apply(lambda x: str(x).replace('.', ','))
        
        if csv_clean.empty:
            st.error("Aucune donnée valide !")
            return None, None

        # 🎯 CRÉATION DU DICTIONNAIRE DE NOTES - CODE NORMALISÉ EN ENTIER
        notes_etudiants = {}
        for _, row in csv_clean.iterrows():
            try:
                # Convertit en entier pour matcher avec l'Excel
                code_key = int(float(str(row['A:Code']).strip()))
                notes_etudiants[code_key] = row['Note']
            except (ValueError, TypeError):
                continue  # Ignore les codes invalides

        # === CHARGEMENT EXCEL ===
        wb = load_workbook(filename=xls_file)
        sheet = wb.active
        
        # 🔍 RECHERCHE DES COLONNES
        code_column = None
        note_column = None
        header_row = None
        
        for i, row in enumerate(sheet.iter_rows(min_row=1, max_row=20), 1):
            for cell in row:
                if cell.value:
                    cell_val = str(cell.value).strip().lower()
                    if cell_val == 'code' and code_column is None:
                        code_column = cell.column
                        header_row = i
                    elif cell_val == 'note' and note_column is None:
                        note_column = cell.column
            if code_column and note_column:
                break
        
        if not code_column or not note_column:
            st.error("Colonnes 'Code' ou 'Note' introuvables !")
            wb.close()
            return None, None

        #  MISE À JOUR DES NOTES AVEC NORMALISATION
        matched_count = 0
        unmatched_codes = []
        updates_to_apply = []
        
        for row_idx, row in enumerate(sheet.iter_rows(min_row=header_row + 1, max_col=max(code_column, note_column)), header_row + 1):
            code_cell = row[code_column - 1]
            note_cell = row[note_column - 1]
            code_raw = code_cell.value
            
            # 🔄 NORMALISATION DU CODE EXCEL EN ENTIER
            try:
                if code_raw is None or str(code_raw).strip() == '':
                    continue
                # Gère "24011624", 24011624, 24011624.0, etc.
                code_int = int(float(str(code_raw).strip()))
            except (ValueError, TypeError):
                unmatched_codes.append(str(code_raw))
                continue
            
            # 🔍 RECHERCHE DANS LE DICTIONNAIRE
            if code_int in notes_etudiants:
                note_value = notes_etudiants[code_int]
                try:
                    if ',' in str(note_value):
                        updates_to_apply.append((note_cell, float(str(note_value).replace(',', '.')), numbers.FORMAT_NUMBER_00))
                    else:
                        updates_to_apply.append((note_cell, int(note_value) if '.' not in str(note_value) else float(note_value), None))
                    matched_count += 1
                except ValueError:
                    pass

        # 📊 FEEDBACK
        if matched_count == 0:
            st.warning("⚠️ Aucune note transférée ! Vérifiez les codes.")
        else:
            st.success(f"✅ {matched_count} notes transférées avec succès !")
        
        if len(unmatched_codes) > 0 and len(unmatched_codes) < 10:
            st.info(f"ℹ️ Codes non reconnus : {unmatched_codes[:5]}...")
        
        # Application des mises à jour
        for cell, value, fmt in updates_to_apply:
            cell.value = value
            if fmt:
                cell.number_format = fmt
        
        wb.close()
        
        # Sauvegarde
        output = io.BytesIO()
        wb = load_workbook(filename=xls_file)
        wb.save(output)
        wb.close()
        output.seek(0)
        
        return output, len(anomalies)

    except Exception as e:
        st.error(f"Erreur traitement : {str(e)}")
        return None, None


# =============================================================================
# INTERFACE UTILISATEUR
# =============================================================================

st.title("🛠️ Outils pour AMC")

section = st.sidebar.radio("📋 Sections", ["ETUDIANTS", "STATISTIQUES", "NOTES"])

# =============================================================================
# ONGLET ETUDIANTS
# =============================================================================
if section == "ETUDIANTS":
    st.header("👨‍🎓 Liste des étudiants")
    st.info("Téléchargez le fichier Excel administration pour générer la liste CSV AMC")
    
    uploaded_file = st.file_uploader("📁 Fichier Excel", type="xlsx", key="excel_uploader")
    
    if uploaded_file is not None:
        with st.spinner("Traitement..."):
            xls, liste = process_excel(uploaded_file)
            
            if xls is not None:
                st.success(f"✅ {len(liste)} étudiants trouvés")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.write("📋 Aperçu Excel:")
                    st.dataframe(xls.head(), use_container_width=True)
                with col2:
                    st.write("📋 Format AMC:")
                    st.dataframe(liste.head(), use_container_width=True)
                
                csv_data = liste.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button(
                    label="📥 Télécharger CSV AMC",
                    data=csv_data,
                    file_name="liste_etudiants.csv",
                    mime="text/csv"
                )

# =============================================================================
# ONGLET STATISTIQUES
# =============================================================================
elif section == "STATISTIQUES":
    st.header("📊 Statistiques des notes")
    st.info("Analysez la distribution des notes AMC")
    
    uploaded_csv = st.file_uploader("📁 Fichier CSV", type="csv", key="csv_stats")
    
    if uploaded_csv is not None:
        with st.spinner("Analyse..."):
            csv_clean, csv_nones = process_csv(uploaded_csv)
            
            if csv_clean is not None:
                effectifs = csv_clean['Note'].value_counts().sort_index().reset_index()
                effectifs.columns = ['Valeur', 'Effectif']
                
                fig = px.bar(effectifs, x='Valeur', y='Effectif', text_auto=True)
                fig.update_layout(xaxis=dict(tickmode='linear', tick0=0, dtick=1), height=400)
                st.plotly_chart(fig, use_container_width=True)
                
                total = len(csv_clean)
                valides = (csv_clean['Note'] >= 10).sum()
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Présents", total)
                col2.metric("Validés", valides)
                col3.metric("Taux réussite", f"{round(valides/total*100, 1)}%")
                col4.metric("Non identifiés", len(csv_nones))

# =============================================================================
# ONGLET NOTES
# =============================================================================
elif section == "NOTES":
    st.header("✍️ Transfert des notes")
    
    col1, col2 = st.columns(2)
    with col1:
        st.info("📄 **Fichier Excel administration**\n\nContient la colonne `Code` au format texte")
        xls_file = st.file_uploader("📁 Excel admin", type="xlsx", key="excel_notes")
    with col2:
        st.info("📄 **Fichier CSV AMC**\n\nContient la colonne `A:Code` au format numérique")
        csv_file = st.file_uploader("📁 CSV AMC", type="csv", key="csv_notes")
    
    if xls_file is not None and csv_file is not None:
        add_notes = st.slider("➕ Points bonus", 0.0, 5.0, 0.0, 0.5)
        
        if st.button("🚀 Transférer les notes", type="primary"):
            xls_file.seek(0)
            csv_file.seek(0)
            
            with st.spinner("Transfert en cours..."):
                result, nb_none = process_csv2excel(xls_file, csv_file, add_notes)
                
                if result is not None:
                    st.download_button(
                        label="📥 Télécharger Excel avec notes",
                        data=result.getvalue(),
                        file_name="notes_transferees.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    if nb_none > 0:
                        st.warning(f"⚠️ {nb_none} étudiants non identifiés")
