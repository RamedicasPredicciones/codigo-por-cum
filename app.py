import streamlit as st
import pandas as pd
from io import BytesIO
from rapidfuzz import fuzz, process

# Cargar datos de Ramedicas desde Google Drive
@st.cache_data
def load_ramedicas_data():
    # URL del archivo Excel en Google Drive
    ramedicas_url = (
        "https://docs.google.com/spreadsheets/d/1Y9SgliayP_J5Vi2SdtZmGxKWwf1iY7ma/export?format=xlsx&sheet=Hoja1"
    )
    # Leer el archivo Excel desde la URL
    ramedicas_df = pd.read_excel(ramedicas_url, sheet_name="Hoja1")
    return ramedicas_df[['cum', 'codart', 'nomart']]

# Preprocesar CUM para una mejor comparación
def preprocess_cum(cum):
    if not isinstance(cum, str):  # Verificar que el CUM sea una cadena
        return ""
    
    replacements = {
        "(": "", ")": "", "+": " ", "/": " ", "-": " ", ",": "", ";": "",
        ".": "", "mg": " mg", "ml": " ml", "capsula": " capsulas",
        "tablet": " tableta", "tableta": " tableta", "parches": " parche", "parche": " parche"
    }
    for old, new in replacements.items():
        cum = cum.lower().replace(old, new)
    
    stopwords = {"de", "el", "la", "los", "las", "un", "una", "y", "en", "por"}
    words = [word for word in cum.split() if word not in stopwords]
    return " ".join(sorted(words))

# Buscar la mejor coincidencia entre el CUM del cliente y los productos de Ramedicas
def find_best_match(client_cum, ramedicas_df):
    client_cum_processed = preprocess_cum(client_cum)
    ramedicas_df['processed_cum'] = ramedicas_df['cum'].apply(preprocess_cum)

    if client_cum_processed in ramedicas_df['processed_cum'].values:
        exact_match = ramedicas_df[ramedicas_df['processed_cum'] == client_cum_processed].iloc[0]
        return {
            'cum_cliente': client_cum,
            'cum_ramedicas': exact_match['cum'],
            'codart': exact_match['codart'],
            'nomart': exact_match['nomart'],
            'score': 100
        }

    client_terms = set(client_cum_processed.split())
    matches = process.extract(client_cum_processed, ramedicas_df['processed_cum'], scorer=fuzz.token_set_ratio, limit=10)

    best_match = None
    highest_score = 0

    for match, score, idx in matches:
        candidate_row = ramedicas_df.iloc[idx]
        candidate_terms = set(match.split())
        if client_terms.issubset(candidate_terms) and score > highest_score:
            highest_score = score
            best_match = {
                'cum_cliente': client_cum,
                'cum_ramedicas': candidate_row['cum'],
                'codart': candidate_row['codart'],
                'nomart': candidate_row['nomart'],
                'score': score
            }

    if not best_match and matches:
        best_match = {
            'cum_cliente': client_cum,
            'cum_ramedicas': matches[0][0],
            'codart': ramedicas_df.iloc[matches[0][2]]['codart'],
            'score': matches[0][1]
        }

    return best_match

# Convertir DataFrame a archivo Excel
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Homologación")
    return output.getvalue()

# Interfaz de Streamlit
st.markdown(
    """
    <h1 style="text-align: center; color: orange;">RAMEDICAS S.A.S.</h1>
    <h3 style="text-align: center;">Homologador de CUMs</h3>
    <p style="text-align: center;">
    Esta herramienta permite realizar la homologación eficiente de los CUMs de productos que nos envíen los clientes, 
    con los productos registrados en la base de datos de Ramedicas S.A.S.
    </p>
    """, unsafe_allow_html=True
)

if st.button("Actualizar base de datos"):
    st.cache_data.clear()

# Subir archivo
uploaded_file = st.file_uploader("O sube tu archivo de excel con la columna CUM que contiene los productos:", type="xlsx")

# Procesar manualmente
client_cums_manual = st.text_area("Ingresa los CUMs de los productos que envió el cliente, separados por saltos de línea:")

ramedicas_df = load_ramedicas_data()

if uploaded_file:
    client_cums_df = pd.read_excel(uploaded_file)
    if 'cum' not in client_cums_df.columns:
        st.error("El archivo debe tener una columna llamada 'cum'.")
    else:
        client_cums = client_cums_df['cum'].tolist()
        matches = []

        for client_cum in client_cums:
            if client_cum.strip():
                match = find_best_match(client_cum, ramedicas_df)
                matches.append(match)

        results_df = pd.DataFrame(matches)
        st.dataframe(results_df)

        excel_data = to_excel(results_df)
        st.download_button(
            label="📥 Descargar resultados en Excel",
            data=excel_data,
            file_name="homologacion_resultados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Procesar texto manual
if client_cums_manual:
    client_cums = client_cums_manual.split("\n")
    matches = []

    for client_cum in client_cums:
        if client_cum.strip():
            match = find_best_match(client_cum, ramedicas_df)
            matches.append(match)

    results_df = pd.DataFrame(matches)
    st.dataframe(results_df)

    excel_data = to_excel(results_df)
    st.download_button(
        label="📥 Descargar resultados en Excel",
        data=excel_data,
        file_name="homologacion_resultados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
