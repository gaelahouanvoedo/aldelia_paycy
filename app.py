import streamlit as st
import pandas as pd
from io import BytesIO
from PIL import Image

# Configuration de la page
st.set_page_config(
    page_title="Analyse de la Paie - Avril 2024",
    page_icon="📊",
    initial_sidebar_state="expanded",
)

def process_data(file):
    # Lecture du fichier Excel
    df = pd.read_excel(file, header=2)
    
    # Filtrage des données pour ne conserver que certains codes
    filtered_codes = ['RPT1003', 'RTCP1725', 'RBI5000']
    df_filtered = df[df['Code'].isin(filtered_codes)].drop_duplicates(subset='Code')
    df_filtered = df_filtered.drop(columns=['Code', 'Description'])
    
    # Calcul des sommes
    df_sum = df_filtered.sum().to_frame().reset_index()
    df_sum.columns = ['Employé', 'Montant']
    df_sum['Montant'] = df_sum['Montant'].round().astype(int)
    
    # Ajout du total général
    total_sum = df_sum['Montant'].sum()
    df_sum.loc[len(df_sum)] = ['Total', total_sum]
    
    return df_sum

def convert_df_to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()
    processed_data = output.getvalue()
    return processed_data

# Chargement de l'image de logo
image = Image.open(r'C:\Users\GaelAHOUANVOEDO\DATA\ORNELLY\log.png')

# Barre latérale
with st.sidebar:
    st.image(image, width=180)
    st.success("Lancez l'application ici 👇")
    menu = st.selectbox("Menu", ('Introduction', "Lancer l'app"))
    st.subheader("Informations")
    st.write("Cette application permet d'extraire les données de paie.", unsafe_allow_html=True)
    '***'
    '**Conçu avec ♥ par Gael Ahouanvoedo**'

# Page d'introduction
if menu == "Introduction":
    st.write("""
    # Extraction de donnée de paie
    
    Cette application permet d'analyser, de traiter et d'extraire les données de paie.
                   
    **👈 Pour démarrer, sélectionnez "Lancer l'app" dans la barre latérale.**
    """)

    st.write("""
    ### Crédits
    Gael Ahouanvoedo, gael.ahouanvoedo@aldelia.com
    """)

    st.write("""
    ### Avertissement
    Il s'agit d'une micro-application web créée pour un besoin spécifique. Elle peut ne pas répondre à vos attentes dans tous vos contextes. Veuillez donc ne pas vous fier entièrement aux résultats issus de son exploitation.
    """)

# Page principale de l'application
if menu == "Lancer l'app":
    st.title("Extraire les données de paie")

    # Téléchargement du fichier Excel par l'utilisateur
    uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx"])

    if uploaded_file is not None:
        # Traitement des données
        df_sum = process_data(uploaded_file)
        
        # Exclure la ligne Total avant d'afficher le DataFrame
        df_display = df_sum[df_sum['Employé'] != 'Total']
        
        # Affichage des résultats
        st.write("Résultats des calculs :")
        st.dataframe(df_display)
        
        # Conversion du DataFrame en fichier Excel
        df_xlsx = convert_df_to_excel(df_display)
        
        # Bouton de téléchargement
        st.download_button(
            label="Télécharger les résultats en Excel",
            data=df_xlsx,
            file_name='bilan_facturation.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.write("Veuillez télécharger un fichier Excel pour commencer l'analyse.")
