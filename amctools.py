import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font
import plotly.express as px
import plotly.graph_objects as go


# Fonction de traitement pour le fichier Excel
def process_excel(file):
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
        
        # Vérification des colonnes nécessaires (double vérification)
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
        return None, 

# Fonction de traitement pour le fichier CSV
def process_csv(csv_file):
    try:
        # Lire le fichier CSV
        csv = pd.read_csv(csv_file, delimiter=';', encoding='utf-8')
        
        # Nettoyer les données : supprimer les lignes où 'A:Code' == 'NONE'
        anomalies = csv[csv['A:Code'] == 'NONE'].copy()
        csv_clean = csv[csv['A:Code'] != 'NONE'].copy()
        
        # Vérifier si le fichier nettoyé est vide
        if csv_clean.empty:
            st.error("Aucune donnée valide après le nettoyage !")
            return None, None, None
        
        # Construire le dictionnaire Notes
        Notes = {row['A:Code'].strip().upper(): row['Note'] for _, row in csv_clean.iterrows()}
        
        return csv_clean, anomalies, Notes
    
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier CSV : {str(e)}")
        return None, None, None

def update_excel_with_notes(file_path, notes):
    try:
        # Lire le fichier Excel sans supposer les en-têtes
        df = pd.read_excel(file_path, header=None)
        
        # Identifier la ligne contenant les en-têtes
        header_row = None
        for i, row in df.iterrows():
            if set(['Code', 'CNE', 'Nom', 'Prénom', 'DATE_NAI_IND', 'Groupe', 'N° Exam', 'Note']).issubset(row.dropna().astype(str).str.strip()):
                header_row = i
                break
        
        if header_row is None:
            st.error("Les en-têtes attendus n'ont pas été trouvés dans le fichier Excel.")
        
        # Créer une copie du DataFrame pour conserver toutes les lignes
        updated_df = df.copy()
        
        # Définir les en-têtes correctement (nettoyage des espaces)
        updated_df.columns = [col.strip() if isinstance(col, str) else f"Unnamed_{j}" for j, col in enumerate(updated_df.iloc[header_row])]
        updated_df.columns = updated_df.columns.str.strip()  # Supprimer les espaces inutiles dans les noms de colonnes
        
        # Filtrer les lignes après les en-têtes
        data_rows = updated_df.iloc[header_row + 1:].reset_index(drop=True)
        
        # Nettoyer la colonne 'Code' pour faciliter la correspondance
        data_rows['Code'] = data_rows['Code'].astype(str).str.strip().str.upper()
        
        # Mettre à jour la colonne 'Note' avec les valeurs du dictionnaire 'Notes'
        data_rows['Note'] = data_rows['Code'].map(notes)
        
        # Remplacer les lignes modifiées dans le DataFrame original
        updated_df.iloc[header_row + 1:] = data_rows.values
        
        return updated_df
    
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier Excel : {str(e)}")
        return None

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Feuille1', header=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data



# ----------------- Interface utilisateur -----------------
st.title("Traitements de fichiers Excel et CSV pour AMC")

# Sidebar pour les sections
section = st.sidebar.radio("Choisir une section", ["Liste des étudiants", "Traitement des notes", "Statistiques"])

if section == "Liste des étudiants":
    st.header("Préparation de la liste des étudiants")
    st.info(
        """
        - Télécharger le fichier Excel de l'administration. 
        - Les en-têtes 'Nom', 'Prénom' et 'Code' seront détectés.
        - Les lignes avant les en-têtes seront automatiquement supprimées.
        - La liste des étudiants à fournir à AMC sera préparée au format Excel avec un en-tête personnalisé.
        """
    )
    
    uploaded_excel_file = st.file_uploader(
        "Télécharger le fichier Excel de l'administration", 
        type="xlsx", 
        key="excel_uploader"
    )
    
    if uploaded_excel_file is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            xls, liste = process_excel(uploaded_excel_file)
            
            if xls is not None:
                st.success(f"Lecture du fichier Excel réussie ! {len(xls)} étudiants trouvés.")  
                st.write("Aperçu de la base de données des étudiants avant traitement automatique :")
                st.write(xls.head(10))              
                st.write("Aperçu de la liste des étudiants à fournir à AMC:")
                st.write(liste.head(10))
                st.success(f"La liste contient {len(xls)} étudiants.")
                # Générer le fichier CSV
                csv_data = liste.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Télécharger la liste des étudiants au format CSV",
                    data=csv_data,
                    file_name="liste_etudiants.csv",
                    mime="text/csv"
                )

elif section == "Traitement des notes":
    st.header("Traitement des notes")
    st.info(
        """
        - Télécharger le fichier Excel de l'administration. 
        - Télécharger le fichier CSV des notes calculées par AMC.
        - Les notes seront automatiquement associées aux étudiants.
        - Le nombre d'étudiants mal identifiés sera indiqué.
        """
    )
    uploaded_excel_file2 = st.file_uploader(
        "Télécharger le fichier Excel de l'administration", 
        type="xlsx", 
        key="excel_uploader2"
    )
    uploaded_csv_file = st.file_uploader(
        "Télécharger le fichier CSV des notes calculées par AMC", 
        type="csv", 
        key="csv_uploader"
    )
    if uploaded_excel_file2 is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            xls, liste = process_excel(uploaded_excel_file2)
    if uploaded_csv_file is not None and uploaded_excel_file2 is not None:
        with st.spinner("Intégration des notes aux étudiants..."):
            csv_clean, anomalies, Notes = process_csv(uploaded_csv_file)
            st.write("Aperçu de la base de données des étudiants :")
            st.write(xls.head(10))
            
            st.write("Aperçu du fichier des notes :")
            st.write(csv_clean.head(10))
                
            # Générer le fichier Excel final avec en-tête personnalisé
            if Notes is not None:
                # Mettre à jour le fichier Excel avec les notes
                updated_df = update_excel_with_notes(uploaded_excel_file2, Notes)
    
            if updated_df is not None:
                # Afficher le DataFrame mis à jour
                st.write("notes des étudiants prêtes à l'envoi :")
                st.write(updated_df)
                # Exporter le résultat dans un nouveau fichier Excel

                processed_data = to_excel(updated_df)
                st.download_button(
                label="📥 Télécharger le fichier final des notes au format Excel",
                data=processed_data,
                file_name="etudiants_avec_notes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            else:
               st.error("La mise à jour du fichier Excel a échoué.")
            
        if len(anomalies) > 0:   
            st.warning(f"Attention! {len(anomalies)} étudiants ont été mal identifiés. Vérifiez leurs copies.")

elif section == "Statistiques":
    st.header("Statistiques des notes")
    st.info(
        """
        - Télécharger le fichier Excel de l'administration. 
        - Télécharger le fichier CSV des notes calculées par AMC.
        - Les notes seront automatiquement associées aux étudiants.
        - Le nombre d'étudiants mal identifiés sera indiqué.
        """
    )
    uploaded_excel_file2 = st.file_uploader(
        "Télécharger le fichier Excel de l'administration", 
        type="xlsx", 
        key="excel_uploader2"
    )
    uploaded_csv_file = st.file_uploader(
        "Télécharger le fichier CSV des notes calculées par AMC", 
        type="csv", 
        key="csv_uploader"
    )
    if uploaded_excel_file2 is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            xls, liste = process_excel(uploaded_excel_file2)
    if uploaded_csv_file is not None and uploaded_excel_file2 is not None:
        with st.spinner("Intégration des notes aux étudiants..."):
            csv_clean, anomalies, Notes = process_csv(uploaded_csv_file)
                
            # Générer le fichier Excel final avec en-tête personnalisé
            if Notes is not None:
                # Mettre à jour le fichier Excel avec les notes
                updated_df = update_excel_with_notes(uploaded_excel_file2, Notes)


            # Affichage des statistiques
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Effectif total", len(xls) if xls is not None else 0)
            with col2:
                st.metric("Présents", len(csv_clean) if csv_clean is not None else 0)
            with col3:
                st.metric("Taux de réussite (%)", round((csv_clean['Note'] >= 10).mean()*100,2)
 if xls is not None and csv_clean is not None else 0)
            with col4:
                st.metric("Mal identifiés", len(anomalies) if anomalies is not None else 0)


            # Calcul des effectifs
            effectifs = csv_clean['Note'].value_counts().reset_index()#.sort_index()
            modalites = csv_clean['Note'].unique()
            effectifs.columns = ['Valeur', 'Effectif']


            # Création du graphique Plotly avec les effectifs affichés sur les barres
            fig = px.bar(effectifs,
                x='Valeur',
                y='Effectif',
                title=" ",
                labels={'Valeur': 'Notes', 'Effectif': 'Effectifs'},
                text_auto=True 
            )

            # Personnalisation du layout
            fig.update_layout(
                title_font_size=20,
                xaxis_title_font=dict(size=14),
                yaxis_title_font=dict(size=14),
                showlegend=False
            )

            # Ajuster la position et le style des étiquettes (optionnel)
            fig.update_traces(textfont_size=14, textangle=0, textposition="outside", width=0.5)
            fig.update_xaxes(tickmode='array', tickvals=list(range(21)), ticktext=[str(i) for i in range(21)])
            fig.update_layout(width=800, height=600)

            st.plotly_chart(fig)

            cola, colb = st.columns(2)
            with cola:
                # Slider pour ajouter des points
                ajout_points = st.slider("Ajouter des points", min_value=0.0, max_value=5.0, value=0.0, step=0.5)

            if ajout_points > 0:
                csv_plus = csv_clean.copy()

                # Ajout des points avec validation (limite maximale de 20)
                csv_plus['Note'] = csv_plus['Note'].apply(lambda x: min(x + ajout_points, 20))

                # Calcul des effectifs après modification
                effectifs_plus = csv_plus['Note'].value_counts().reset_index()
                effectifs_plus.columns = ['Valeur', 'Effectif']

                # Création du graphique Plotly avec les effectifs affichés sur les barres
                fig_plus = px.bar(
                    effectifs_plus,
                    x='Valeur',
                    y='Effectif',
                    title=" ",
                    labels={'Valeur': 'Notes', 'Effectif': 'Effectifs'},
                    text_auto=True
                )

                # Personnalisation du layout
                fig_plus.update_layout(
                    title_font_size=20,
                    xaxis_title_font=dict(size=14),
                    yaxis_title_font=dict(size=14),
                    showlegend=False
                )

                # Ajuster la position et le style des étiquettes
                fig_plus.update_traces(textfont_size=14, textangle=0, textposition="outside", width=0.5)

                # Configuration de l'axe des abscisses pour inclure toutes les valeurs de 0 à 20
                fig_plus.update_xaxes(tickmode='array', tickvals=list(range(21)), ticktext=[str(i) for i in range(21)])

                # Définir la taille du graphique
                fig_plus.update_layout(width=800, height=600)

                # Affichage du taux de réussite mis à jour
                with colb:
                    st.metric("Nouveau taux de réussite (%)", round((csv_plus['Note'] >= 10).mean() * 100, 2))

                # Affichage du graphique
                st.plotly_chart(fig_plus)






    

            #st.download_button(
            #    label="📥 Télécharger le fichier final des notes au format Excel",
            #    data=excel_data,
            #    file_name="etudiants_avec_notes.xlsx",
            #    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            #)
            