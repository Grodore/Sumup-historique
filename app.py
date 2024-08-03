import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Vos fonctions existantes
def filter_and_write_to_excel(dataframe, description_list, excel_file_stream, date_JJ_MM):
    workbook = openpyxl.load_workbook(excel_file_stream)
    sheet = workbook[date_JJ_MM]
    filtered_df = dataframe[dataframe['Description'].isin(description_list)]
    ligne_excel = 2
    for index, sell in filtered_df.iterrows():
        sheet['A'+str(ligne_excel)] = sell["Time"]
        sheet['B'+str(ligne_excel)] = sell["Description"]
        sheet['C'+str(ligne_excel)] = sell["Price (Gross)"]
        ligne_excel += 1
    # Enregistrer le fichier dans un BytesIO pour un téléchargement ultérieur
    output_stream = BytesIO()
    workbook.save(output_stream)
    output_stream.seek(0)
    return output_stream

def sum_transaction(dataframe, description_list, excel_file_stream, date_JJ_MM):
    filtered_df = dataframe[dataframe['Description'].isin(description_list)]
    df = pd.DataFrame({'Description': description_list, 'Total': 0})
    workbook = openpyxl.load_workbook(excel_file_stream)
    sheet = workbook[date_JJ_MM]
    total_index = 0
    for index, row in df.iterrows():
        df.loc[index, 'Total'] = filtered_df[filtered_df['Description'] == row['Description']]['Price (Gross)'].sum()
        sheet['E'+str(index+2)] = row['Description']
        sheet['F'+str(index+2)] = df.loc[index, 'Total']
        total_index += 1
    sheet['E'+str(total_index+2)] = "TOTAL"
    sheet['F'+str(total_index+2)] = df['Total'].sum()
    output_stream = BytesIO()
    workbook.save(output_stream)
    output_stream.seek(0)
    return output_stream

# Fonction pour convertir un fichier Excel en lien de téléchargement
def to_excel(file_stream):
    return file_stream.getvalue()

# Application Streamlit
st.title('Historique des Transactions SumUp')

# Téléchargement du fichier CSV
uploaded_file = st.file_uploader("Choisissez un fichier CSV", type="csv")
template_excel_file = st.file_uploader("Choisissez un fichier Excel modèle", type="xlsx")
date_JJ_MM = st.text_input("Entrez la date (JJ_MM) pour sélectionner la feuille Excel correspondante")

if uploaded_file and template_excel_file and date_JJ_MM:
    # Lire les données du CSV
    df = pd.read_csv(uploaded_file)

    # Récupérer la liste unique des descriptions
    unique_descriptions = df['Description'].unique()
    
    # Afficher les checkboxes pour chaque description unique
    st.header("Sélectionnez les descriptions à inclure")
    selected_descriptions = [desc for desc in unique_descriptions if st.checkbox(desc)]
    
    if st.button("Appliquer les filtres et télécharger le fichier Excel"):
        if selected_descriptions:
            # Sauvegarde le fichier Excel mis à jour
            excel_file_stream = BytesIO(template_excel_file.read())
            updated_excel_stream = filter_and_write_to_excel(df, selected_descriptions, excel_file_stream, date_JJ_MM)
            updated_excel_stream = sum_transaction(df, selected_descriptions, updated_excel_stream, date_JJ_MM)
            
            # Convertir en fichier téléchargeable
            excel_data = to_excel(updated_excel_stream)
            
            # Lien de téléchargement du fichier Excel mis à jour
            st.download_button(label="Télécharger le fichier Excel mis à jour", 
                               data=excel_data,
                               file_name="result.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("Veuillez sélectionner au moins une description.")
