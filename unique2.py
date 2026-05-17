import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import plotly.express as px
import io
import csv

# =============================================================================
# CONFIGURATION DE LA PAGE
# =============================================================================
st.set_page_config(page_title="🛠️ Outils AMC", layout="wide")

# =============================================================================
# FONCTIONS UTILITAIRES
# =============================================================================

def detect_delimiter(file_content):
    sample = file_content.decode('utf-8', errors='ignore')
    try:
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter
    except csv.Error:
        return ','

# =============================================================================
# FONCTION PRINCIPALE DE TRANSFERT (CORRIGÉE)
# =============================================================================

def process_csv2excel(xls_file, csv_file, add_notes=0):
    """
    Fusionne les notes du CSV AMC vers le fichier Excel.
    CORRECTION : Conversion systématique des codes en CHAÎNE DE CARACTÈRES (String)
    pour garantir la correspondance, quel que soit le format (int, float, text).
    """
    try:
        # 1. Lecture et nettoyage du CSV
        csv_content = csv_file.read()
        delimiter = detect_delimiter(csv_content)
        csv_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')), delimiter=delimiter)
        
        # Renommer Mark en Note si nécessaire
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
        
        # Vérifier les colonnes requises
        if 'A:Code' not in csv_data.columns or 'Note' not in csv_data.columns:
            st.error("Le fichier CSV ne contient pas les colonnes 'A:Code' ou 'Note'.")
            return None, 0

        # Filtrer les lignes valides (Code != NONE)
        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()
        anomalies_count = len(csv_data) - len(csv_clean)

        # --- CORRECTION MAJEURE ICI ---
        # Création du dictionnaire de notes avec des CLÉS EN TEXTE (String)
        # str(x).strip() enlève les espaces invisibles et convertit les nombres en texte
        notes_dict = {}
        for _, row in csv_clean.iterrows():
            code_raw = row['A:Code']
            note_raw = row['Note']
            
            # Nettoyage du code
            code_key = str(code_raw).strip()
            
            # Nettoyage de la note (gestion des virgules/points)
            if isinstance(note_raw, str):
                note_val = note_raw.replace(',', '.').strip()
                try:
                    note_val = float(note_val)
                except:
                    note_val = note_raw # Garder le texte si ce n'est pas un nombre
            else:
                note_val = note_raw
            
            # Ajouter les points bonus si demandé
            if add_notes > 0 and isinstance(note_val, (int, float)):
                note_val = min(note_val + add_notes, 20)
            
            notes_dict[code_key] = note_val

        if not notes_dict:
            st.warning("Aucune note valide trouvée dans le CSV.")
            return None, 0

        # 2. Ouverture du fichier Excel
        wb = load_workbook(xls_file)
        ws = wb.active

        # 3. Recherche dynamique des colonnes "Code" et "Note"
        code_col_idx = None
        note_col_idx = None
        header_row_idx = None

        # On scanne les 15 premières lignes pour trouver les en-têtes
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=15), 1):
            for cell in row:
                if cell.value:
                    val = str(cell.value).strip().lower()
                    if val == 'code' and code_col_idx is None:
                        code_col_idx = cell.column
                        header_row_idx = row_idx
                    elif val == 'note' and note_col_idx is None:
                        note_col_idx = cell.column
                        header_row_idx = row_idx
            
            if code_col_idx and note_col_idx:
                break

        if not code_col_idx or not note_col_idx:
            st.error("Impossible de trouver les colonnes 'Code' et 'Note' dans l'Excel.")
            return None, 0

        # 4. Transfert des notes
        matched_count = 0
        # On commence à la ligne juste après les en-têtes
        start_row = header_row_idx + 1

        for row in ws.iter_rows(min_row=start_row, max_col=max(code_col_idx, note_col_idx)):
            code_cell = row[code_col_idx - 1]
            note_cell = row[note_col_idx - 1]
            
            if code_cell.value is None:
                continue

            # --- CORRECTION MAJEURE ICI ---
            # Conversion du code Excel en TEXTE pour la comparaison
            excel_code_key = str(code_cell.value).strip()

            # Recherche dans le dictionnaire
            if excel_code_key in notes_dict:
                final_note = notes_dict[excel_code_key]
                
                # Écriture dans Excel
                note_cell.value = final_note
                matched_count += 1

        wb.close() # Important de fermer avant de sauvegarder si on recharge

        # 5. Sauvegarde
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return output, anomalies_count, matched_count, len(notes_dict)

    except Exception as e:
        st.error(f"Erreur technique : {str(e)}")
        return None, 0, 0, 0

# =============================================================================
# INTERFACE UTILISATEUR
# =============================================================================

st.title("🛠️ Transfert des Notes AMC vers Excel")

tab1, tab2 = st.tabs(["📤 Transfert de Notes", "ℹ️ Aide"])

with tab1:
    st.subheader("1. Configuration")
    
    col1, col2 = st.columns(2)
    with col1:
        xls_file = st.file_uploader("📄 Fichier Excel (Liste Admin)", type=['xlsx', 'xls'])
    with col2:
        csv_file = st.file_uploader("📄 Fichier CSV (Résultats AMC)", type=['csv'])

    add_notes = st.number_input("➕ Ajouter des points (Bonus)", min_value=0.0, max_value=5.0, value=0.0, step=0.5)

    if st.button("🚀 Lancer le transfert", type="primary", disabled=not (xls_file and csv_file)):
        with st.spinner("Traitement en cours..."):
            # Reset des fichiers pour lecture
            xls_file.seek(0)
            csv_file.seek(0)
            
            result_file, anomalies, matched, total_notes = process_csv2excel(xls_file, csv_file, add_notes)
            
            if result_file:
                st.success(f"✅ Succès ! {matched} notes transférées sur {total_notes} disponibles.")
                if anomalies > 0:
                    st.warning(f"⚠️ {anomalies} étudiants ignorés (Code = NONE).")
                
                st.download_button(
                    label="📥 Télécharger le fichier Excel final",
                    data=result_file,
                    file_name="notes_transferees.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("❌ Le transfert a échoué.")

with tab2:
    st.markdown("""
    ### Comment ça marche ?
    1. **Fichier Excel** : Doit contenir une colonne nommée `Code` et une colonne `Note`.
    2. **Fichier CSV** : Doit être l'export standard d'AMC (contient `A:Code` et `Note`).
    3. **Correspondance** : Le script compare le `Code` Excel avec le `A:Code` du CSV.
    
    ### Pourquoi cela ne marchait pas avant ?
    Souvent, Excel stocke les codes comme du **Texte** ("123") tandis que le CSV les donne en **Nombre** (123). 
    Cette version du script force la conversion en texte pour les deux, garantissant la correspondance.
    """)
