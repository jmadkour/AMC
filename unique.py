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
st.set_page_config(page_title="🛠️ Outils AMC", page_icon="🎓", layout="wide")
st.title("🛠️ Outils pour AMC")
section = st.sidebar.radio(" ", ["ETUDIANTS", "STATISTIQUES", "NOTES"])

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

def process_excel(file):
    try:
        file.seek(0)
        xls = pd.read_excel(file, header=None)
        header_index = next((idx for idx, row in xls.iterrows() if all(col in row.values for col in ['Code', 'Nom', 'Prénom'])), None)
        if header_index is None:
            st.error("En-têtes introuvables.")
            return None, None
        xls.columns = xls.iloc[header_index]
        xls = xls.iloc[header_index + 1:].reset_index(drop=True)
        if xls.empty:
            st.error("Aucune donnée valide.")
            return None, None
        required = ['Nom', 'Prénom', 'Code']
        missing = [c for c in required if c not in xls.columns]
        if missing:
            st.error(f"Colonnes manquantes : {', '.join(missing)}")
            return None, None
        liste = xls.dropna(subset=['Nom', 'Prénom', 'Code'])
        liste['Name'] = liste['Code'].astype(str) + ' ' + liste['Nom'] + ' ' + liste['Prénom']
        liste = liste[['Code', 'Name']].drop_duplicates()
        return xls, liste
    except Exception as e:
        st.error(f"Erreur Excel : {e}")
        return None, None

def process_csv(csv_file):
    try:
        csv_file.seek(0)
        content = csv_file.read()
        delimiter = detect_delimiter(content)
        data = pd.read_csv(io.StringIO(content.decode('utf-8')), delimiter=delimiter)
        if 'Mark' in data.columns:
            data = data.rename(columns={'Mark': 'Note'})
        if 'A:Code' not in data.columns or 'Note' not in data.columns:
            st.error("Colonnes 'A:Code' ou 'Note' manquantes.")
            return None, None
        clean = data[data['A:Code'] != 'NONE'].copy()
        nones = data[data['A:Code'] == 'NONE'].copy()
        clean['Note'] = pd.to_numeric(clean['Note'].astype(str).str.replace(',', '.'), errors='coerce')
        if clean.empty:
            st.error("Aucune donnée valide.")
            return None, None
        return clean, nones
    except Exception as e:
        st.error(f"Erreur CSV : {e}")
        return None, None

# =============================================================================
# FONCTION PRINCIPALE DE TRANSFERT (CORRIGÉE)
# =============================================================================

def process_csv2excel(xls_file, csv_file, add_notes=0):
    """
    Fusionne les notes du CSV AMC vers le fichier Excel.
    FIX : Gestion correcte du flux, sauvegarde avant fermeture, et matching string strict.
    """
    try:
        # 1. Réinitialisation des pointeurs
        xls_file.seek(0)
        csv_file.seek(0)

        # 2. Lecture CSV
        csv_content = csv_file.read()
        delimiter = detect_delimiter(csv_content)
        csv_data = pd.read_csv(io.StringIO(csv_content.decode('utf-8')), delimiter=delimiter)
        
        if 'Mark' in csv_data.columns:
            csv_data = csv_data.rename(columns={'Mark': 'Note'})
        
        if 'A:Code' not in csv_data.columns or 'Note' not in csv_data.columns:
            st.error("CSV invalide : colonnes 'A:Code' ou 'Note' manquantes.")
            return None, 0

        csv_clean = csv_data[csv_data['A:Code'] != 'NONE'].copy()
        anomalies_count = len(csv_data) - len(csv_clean)

        # 3. Dictionnaire de notes (Clés strictement en STRING)
        notes_dict = {}
        for _, row in csv_clean.iterrows():
            code_key = str(row['A:Code']).strip()
            note_val = row['Note']
            
            # Nettoyage note
            if isinstance(note_val, str):
                try: note_val = float(note_val.replace(',', '.'))
                except: pass
                
            if add_notes > 0 and isinstance(note_val, (int, float)):
                note_val = min(note_val + add_notes, 20)
                
            notes_dict[code_key] = note_val

        if not notes_dict:
            st.warning("Aucune note valide.")
            return None, 0

        # 4. Ouverture Excel
        wb = load_workbook(xls_file)
        ws = wb.active

        # 5. Recherche en-têtes
        code_col_idx = note_col_idx = header_row_idx = None
        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=15), 1):
            for cell in row:
                if cell.value:
                    val = str(cell.value).strip().lower()
                    if val == 'code' and code_col_idx is None:
                        code_col_idx = cell.column
                        header_row_idx = r_idx
                    elif val == 'note' and note_col_idx is None:
                        note_col_idx = cell.column
            if code_col_idx and note_col_idx:
                break

        if not code_col_idx or not note_col_idx:
            st.error("Colonnes 'Code' ou 'Note' introuvables.")
            wb.close()
            return None, 0

        # 6. Transfert
        matched_count = 0
        for row in ws.iter_rows(min_row=header_row_idx + 1, max_col=max(code_col_idx, note_col_idx)):
            code_cell = row[code_col_idx - 1]
            note_cell = row[note_col_idx - 1]
            
            if code_cell.value is None: continue
            
            # Matching strict en string
            excel_code_key = str(code_cell.value).strip()
            if excel_code_key in notes_dict:
                note_cell.value = notes_dict[excel_code_key]
                matched_count += 1

        # 7. SAUVEGARDE & FERMETURE (CORRECTION CRITIQUE)
        output = io.BytesIO()
        wb.save(output)  # Sauvegarder AVANT de fermer
        wb.close()
        output.seek(0)

        return output, anomalies_count, matched_count, len(notes_dict)

    except Exception as e:
        st.error(f"Erreur technique : {str(e)}")
        return None, 0, 0, 0

# =============================================================================
# INTERFACE UTILISATEUR
# =============================================================================

if section == "ETUDIANTS":
    st.header("👨‍🎓 Liste des étudiants")
    st.info("Générez le fichier CSV compatible AMC depuis le fichier administration.")
    up_xls = st.file_uploader("Fichier Excel administration", type=["xlsx", "xls"], key="up_etud")
    if up_xls:
        with st.spinner("Traitement..."):
            df, lst = process_excel(up_xls)
            if df is not None:
                st.success(f"✅ {len(lst)} étudiants trouvés")
                st.dataframe(lst.head(), use_container_width=True)
                csv_bytes = lst.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button("📥 Télécharger CSV AMC", csv_bytes, "liste_amc.csv", "text/csv")

elif section == "STATISTIQUES":
    st.header("📊 Statistiques des notes")
    st.info("Visualisez la distribution et simulez des bonus.")
    up_csv = st.file_uploader("Fichier CSV résultats AMC", type="csv", key="up_stat")
    if up_csv:
        with st.spinner("Analyse..."):
            clean, nones = process_csv(up_csv)
            if clean is not None:
                fig = px.bar(clean['Note'].value_counts().reset_index(), x='index', y='Note', 
                             labels={'index': 'Note', 'Note': 'Effectif'}, text_auto=True)
                fig.update_xaxes(type='linear', dtick=1)
                st.plotly_chart(fig, use_container_width=True)
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Présents", len(clean))
                c2.metric("Validés (≥10)", (clean['Note'] >= 10).sum())
                c3.metric("Taux réussite", f"{round((clean['Note'] >= 10).sum()/len(clean)*100, 1)}%")
                c4.metric("Mal identifiés", len(nones))
                
                bonus = st.slider("Bonus à simuler", 0.0, 5.0, 0.0, 0.5)
                if bonus > 0:
                    sim = clean.copy()
                    sim['Note'] = sim['Note'].apply(lambda x: min(x + bonus, 20))
                    val_sim = (sim['Note'] >= 10).sum()
                    st.metric(f"Validés avec +{bonus}", val_sim, delta=val_sim - (clean['Note'] >= 10).sum())

elif section == "NOTES":
    st.header("✍️ Transfert des notes vers Excel")
    st.info("Importez les deux fichiers. Le script matche les codes en texte pour éviter les échecs int/float.")
    
    c1, c2 = st.columns(2)
    with c1: xls_f = st.file_uploader("📄 Excel Administration", type=["xlsx", "xls"], key="up_xls_note")
    with c2: csv_f = st.file_uploader("📄 CSV Résultats AMC", type="csv", key="up_csv_note")
    
    bonus_pts = st.number_input("Points bonus à ajouter", 0.0, 5.0, 0.0, 0.5)
    
    if xls_f and csv_f and st.button("🚀 Lancer le transfert", type="primary"):
        with st.spinner("Fusion en cours..."):
            res, anom, matched, total = process_csv2excel(xls_f, csv_f, bonus_pts)
            if res:
                st.success(f"✅ Transfert réussi ! {matched} notes appliquées sur {total} disponibles.")
                if anom > 0: st.warning(f"⚠️ {anom} étudiants ignorés (Code = NONE).")
                st.download_button("📥 Télécharger Excel final", res.getvalue(), "notes_transferees.xlsx", 
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.error("❌ Échec du transfert. Vérifiez les formats et les en-têtes.")
