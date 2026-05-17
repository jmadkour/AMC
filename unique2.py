
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import plotly.express as px
import io
import csv
from openpyxl.styles import numbers
import hashlib

# =============================================================================
# CONFIGURATION DE PAGE - À METTRE EN PREMIER
# =============================================================================
st.set_page_config(
    page_title="🛠️ Outils pour AMC", 
    page_icon="🎓", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============================================================================
# FONCTIONS AVEC CACHING POUR PERFORMANCE
# =============================================================================

@st.cache_data(show_spinner=False)
def detect_delimiter_cached(file_bytes):
    """Détection du délimiteur avec cache"""
    sample = file_bytes.decode('utf-8', errors='ignore')[:10000]  # Premier 10KB seulement
    try:
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter
    except csv.Error:
        return ','


@st.cache_data(show_spinner=False)
def process_excel_cached(file_bytes):
    """Version cachée du traitement Excel"""
    try:
        xls = pd.read_excel(io.BytesIO(file_bytes), header=None)
        
        header_index = next(
            (idx for idx, row in xls.iterrows() if all(col in row.values for col in ['Code', 'Nom', 'Prénom'])),
            None
        )
        
        if header_index is None:
            return None, None, "En-têtes introuvables"
        
        xls.columns = xls.iloc[header_index]
        xls = xls.iloc[header_index + 1:].reset_index(drop=True)
        
        if xls.empty:
            return None, None, "Aucune donnée"
        
        required_columns = ['Nom', 'Prénom', 'Code']
        missing = [col for col in required_columns if col not in xls.columns]
        if missing:
            return None, None, f"Colonnes manquantes: {', '.join(missing)}"
        
        liste = xls.dropna(subset=['Nom', 'Prénom', 'Code'])
        liste['Name'] = liste['Code'].astype(str) + ' ' + liste['Nom'] + ' ' + liste['Prénom']
        liste = liste[['Code', 'Name']].drop_duplicates()
        
        return xls, liste, None
        
    except Exception as e:
        return None, None, str(e)


@st.cache_data(show_spinner=False)
def process_csv_cached(file_bytes):
    """Version cachée du traitement CSV"""
    try:
        delimiter = detect_delimiter_cached(file_bytes)
        csv_data = pd.read_csv(
            io.StringIO(file_bytes.decode('utf-8')), 
            delimiter=delimiter, 
            encoding='utf-8'
        )
        
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
        
        if 'A:Code' not in csv_data.columns or 'Note' not in csv_data.columns:
            return None, None, "Colonnes manquantes"
        
        csv_nones = csv_data[csv_data['A:Code'] == 'NONE'].copy()
        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()
        
        csv_clean['A:Code'] = pd.to_numeric(csv_clean['A:Code'], errors='coerce')
        csv_clean['Note'] = csv_clean['Note'].astype(str).str.replace(',', '.').astype(float)
        
        if csv_clean.empty:
            return None, None, "Aucune donnée valide"
        
        return csv_clean, csv_nones, None
        
    except Exception as e:
        return None, None, str(e)


def process_csv2excel_fast(xls_file, csv_file, add_notes=0):
    """
    Version optimisée sans cache (car modification de fichier)
    Mais avec traitement plus efficace
    """
    try:
        # Lecture rapide du CSV
        csv_content = csv_file.read()
        delimiter = detect_delimiter_cached(csv_content)
        csv_data = pd.read_csv(
            io.StringIO(csv_content.decode('utf-8')), 
            delimiter=delimiter, 
            encoding='utf-8'
        )
        
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
        
        # Création du dictionnaire optimisée
        notes_etudiants = {}
        for _, row in csv_clean.iterrows():
            try:
                code_key = int(float(str(row['A:Code']).strip()))
                notes_etudiants[code_key] = row['Note']
            except:
                continue
        
        # Chargement Excel avec read_only pour performance
        wb = load_workbook(filename=xls_file, read_only=False, data_only=True)
        sheet = wb.active
        
        # Recherche optimisée des colonnes
        code_column = None
        note_column = None
        header_row = None
        
        # Limiter la recherche aux 20 premières lignes
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
            st.error("Colonnes 'Code' et/ou 'Note' introuvables")
            wb.close()
            return None, None
        
        # Mise à jour batch des notes
        matched_count = 0
        updates_to_apply = []
        
        for row_idx, row in enumerate(sheet.iter_rows(min_row=header_row + 1, max_col=max(code_column, note_column)), header_row + 1):
            code_cell = row[code_column - 1]
            note_cell = row[note_column - 1]
            code_raw = code_cell.value
            
            try:
                if code_raw is None or str(code_raw).strip() == '':
                    continue
                code_int = int(float(str(code_raw).strip()))
            except:
                continue
            
            if code_int in notes_etudiants:
                note_value = notes_etudiants[code_int]
                try:
                    if ',' in str(note_value):
                        updates_to_apply.append((note_cell, float(str(note_value).replace(',', '.')), numbers.FORMAT_NUMBER_00))
                    else:
                        updates_to_apply.append((note_cell, int(note_value) if '.' not in str(note_value) else float(note_value), None))
                    matched_count += 1
                except:
                    pass
        
        # Application batch des mises à jour
        for cell, value, fmt in updates_to_apply:
            cell.value = value
            if fmt:
                cell.number_format = fmt
        
        wb.close()
        
        # Feedback
        if matched_count == 0:
            st.warning("⚠️ Aucune note transférée")
        else:
            st.success(f"✅ {matched_count} notes transférées avec succès")
        
        # Rechargement pour sauvegarde
        wb = load_workbook(filename=xls_file)
        output = io.BytesIO()
        wb.save(output)
        wb.close()
        output.seek(0)
        
        return output, len(anomalies)
        
    except Exception as e:
        st.error(f"Erreur: {str(e)}")
        return None, None


# =============================================================================
# INTERFACE UTILISATEUR
# =============================================================================

st.title("🛠️ Outils pour AMC")

# Sidebar
section = st.sidebar.radio("📋 Sections", ["ETUDIANTS", "STATISTIQUES", "NOTES"])

# =============================================================================
# ONGLET ETUDIANTS
# =============================================================================
if section == "ETUDIANTS":
    st.header("👨‍🎓 Liste des étudiants")
    st.info("Téléchargez le fichier Excel administration pour générer la liste CSV pour AMC")
    
    uploaded_file = st.file_uploader("📁 Fichier Excel", type="xlsx", key="excel_uploader")
    
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        
        with st.spinner("⏳ Traitement en cours..."):
            xls, liste, error = process_excel_cached(file_bytes)
            
            if error:
                st.error(f"❌ {error}")
            elif xls is not None:
                st.success(f"✅ {len(liste)} étudiants trouvés")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.write("📋 Aperçu:")
                    st.dataframe(xls.head(), use_container_width=True)
                with col2:
                    st.write("📋 Format AMC:")
                    st.dataframe(liste.head(), use_container_width=True)
                
                csv_data = liste.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button(
                    label="📥 Télécharger CSV",
                    data=csv_data,
                    file_name="liste_etudiants.csv",
                    mime="text/csv"
                )

# =============================================================================
# ONGLET STATISTIQUES
# =============================================================================
elif section == "STATISTIQUES":
    st.header("📊 Statistiques des notes")
    st.info("Analysez la distribution des notes depuis le fichier CSV AMC")
    
    uploaded_csv = st.file_uploader("📁 Fichier CSV", type="csv", key="csv_stats")
    
    if uploaded_csv is not None:
        file_bytes = uploaded_csv.getvalue()
        
        with st.spinner("⏳ Analyse en cours..."):
            csv_clean, csv_nones, error = process_csv_cached(file_bytes)
            
            if error:
                st.error(f"❌ {error}")
            elif csv_clean is not None:
                # Graphique
                effectifs = csv_clean['Note'].value_counts().sort_index().reset_index()
                effectifs.columns = ['Valeur', 'Effectif']
                
                fig = px.bar(effectifs, x='Valeur', y='Effectif', text_auto=True)
                fig.update_layout(xaxis=dict(tickmode='linear', tick0=0, dtick=1), height=400)
                st.plotly_chart(fig, use_container_width=True)
                
                # Métriques
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
    st.header("✍️ Fusion des notes")
    
    col1, col2 = st.columns(2)
    with col1:
        xls_file = st.file_uploader("📁 Excel admin", type="xlsx", key="excel_notes")
    with col2:
        csv_file = st.file_uploader("📁 CSV AMC", type="csv", key="csv_notes")
    
    if xls_file and csv_file:
        add_notes = st.number_input("Points bonus", 0.0, 5.0, 0.0, 0.5)
        
        if st.button("🚀 Fusionner", type="primary"):
            xls_file.seek(0)
            csv_file.seek(0)
            
            with st.spinner("⏳ Fusion en cours..."):
                result, nb_none = process_csv2excel_fast(xls_file, csv_file, add_notes)
                
                if result:
                    st.download_button(
                        label="📥 Télécharger Excel traité",
                        data=result.getvalue(),
                        file_name="notes_traitees.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    if nb_none > 0:
                        st.warning(f"⚠️ {nb_none} étudiants non identifiés")

# Footer
st.divider()
st.caption("🛠️ Outils AMC optimisés v2.2")
