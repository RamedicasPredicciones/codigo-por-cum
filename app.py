import streamlit as st
import pandas as pd
from io import BytesIO
from googleapiclient.discovery import build
from google.oauth2 import service_account
from rapidfuzz import fuzz, process

# Cargar datos desde Google Sheets (CUMs de clientes)
def load_client_cums():
    # URL del archivo de Google Sheets con los CUMs de los clientes
    sheet_url = "https://docs.google.com/spreadsheets/d/1CP-mR-ZeuI2ga8R7_CCXB0aiV_vqsrpK/edit?usp=sharing&ouid=109532697276677589725&rtpof=true&sd=true"
    sheet_id = sheet_url.split('/d/')[1].split('/')[0]  # Extraer el ID del sheet
    range_name = 'Hoja1!A:A'  # Asumimos que los CUMs están en la columna A

    # Autenticación con Google Sheets
    creds = service_account.Credentials.from_service_account_file('path_to_your_service_account.json')
    service = build('sheets', 'v4', credentials=creds)

    # Leer datos de Google Sheets
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get('values', [])

    # Convertir los CUMs a un DataFrame
    client_cums_df = pd.DataFrame(values, columns=["cum"])
    return client_cums_df

# Cargar datos de Ramedicas
def load_ramedicas_data():
    # URL del archivo Excel de Ramedicas
    ramedicas_url = "https://docs.google.com/spreadsheets/d/1Y9SgliayP_J5Vi2SdtZmGxKWwf1iY7ma/export?format=xlsx&sheet=Hoja1"
    ramedicas_df = pd.read_excel(ramedicas_url, sheet_name="Hoja1")
    return ramedicas_df[['cum', 'codart', 'nomart']]  # 'cum', 'codart' y 'nomart'

# Preprocesar CUM para asegurar que siempre esté en minúsculas
def preprocess_cum(cum):
    if isinstance(cum, str):  # Verificar que cum sea una cadena de texto
        return cum.strip().lower()  # Limpiar y poner el CUM en minúsculas
    return ""  # Si no es una cadena, devolver una cadena vacía

# Buscar la mejor coincidencia entre el CUM del cliente y los productos de Ramedicas
def find_best_match(client_cum, ramedicas_df):
    # Preprocesar el CUM del cliente
    client_cum_processed = preprocess_cum(client_cum)
    # Aplicar el preprocesamiento a todos los productos de Ramedicas
    ramedicas_df['processed_cum'] = ramedicas_df['cum'].apply(preprocess_cum)

    # Buscar coincidencia exacta primero
    if client_cum_processed in ramedicas_df['processed_cum'].values:
        exact_match = ramedicas_df[ramedicas_df['processed_cum'] == client_cum_processed].iloc[0]
        return {
            'cum_cliente': client_cum,
            'cum_ramedicas': exact_match['cum'],  # CUM de Ramedicas
            'codart_ramedicas': exact_match['codart'],  # Código del producto (codart)
            'nomart_ramedicas': exact_match['nomart'],  # Nombre del producto (nomart)
            'score': 100  # Puntaje 100 para coincidencia exacta
        }

    # Buscar las mejores coincidencias utilizando RapidFuzz (algoritmo de fuzzy matching)
    matches = process.extract(
        client_cum_processed,
        ramedicas_df['processed_cum'],
        scorer=fuzz.token_set_ratio,  # Usar el método de puntuación token_set_ratio
        limit=10  # Limitar el número de coincidencias a 10
    )

    best_match = None
    highest_score = 0

    # Iterar sobre las coincidencias encontradas
    for match, score, idx in matches:
        candidate_row = ramedicas_df.iloc[idx]  # Obtener la fila candidata

        if score > highest_score:  # Si la puntuación es mejor que la anterior, se actualiza
            highest_score = score
            best_match = {
                'cum_cliente': client_cum,
                'cum_ramedicas': candidate_row['cum'],  # CUM de Ramedicas
                'codart_ramedicas': candidate_row['codart'],  # Código del producto (codart)
                'nomart_ramedicas': candidate_row['nomart'],  # Nombre del producto (nomart)
                'score': score  # Puntaje de la coincidencia
            }

    return best_match

# Interfaz de Streamlit
st.title("Homologador de Productos - CUM a Ramedicas")  # El título de la aplicación

# Opción para actualizar la base de datos y limpiar el caché
if st.button("Actualizar base de datos"):
    st.cache_data.clear()  # Limpiar el caché para cargar los datos de nuevo

# Espacio de entrada para subir archivo con los CUMs de los clientes
uploaded_file = st.file_uploader("Sube tu archivo con los CUMs de los clientes", type="xlsx")

if uploaded_file:
    # Leer el archivo subido con los CUMs de los clientes
    client_cums_df = pd.read_excel(uploaded_file)

    # Verificar si la columna 'cum' está presente en el archivo subido
    if 'cum' not in client_cums_df.columns:
        st.error("El archivo debe contener una columna llamada 'cum'.")
    else:
        # Cargar los datos de productos de Ramedicas
        ramedicas_df = load_ramedicas_data()

        # Lista para almacenar los resultados de la homologación
        results = []

        # Se crea la iteración sobre cada CUM para encontrar la mejor coincidencia
        for client_cum in client_cums_df['cum']:
            match = find_best_match(client_cum, ramedicas_df)
            if match:
                results.append(match)
            else:
                results.append({
                    'cum_cliente': client_cum,
                    'cum_ramedicas': None,
                    'codart_ramedicas': None,
                    'nomart_ramedicas': None,
                    'score': 0
                })

        # Crear un DataFrame con los resultados
        results_df = pd.DataFrame(results)
        st.write("Resultados de homologación:")
        st.dataframe(results_df)  # Mostrar los resultados en una tabla

        # Función para convertir el DataFrame a un archivo Excel
        def to_excel(df):
            output = BytesIO()  # Usamos BytesIO para manejar el archivo en memoria
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Homologación")  # Escribir el DataFrame al archivo
            return output.getvalue()

        # Botón para descargar los resultados como un archivo Excel
        st.download_button(
            label="Descargar archivo con resultados",
            data=to_excel(results_df),
            file_name="homologacion_productos.xlsx",  # Nombre del archivo de descarga
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",  # Tipo de archivo excel 
        )
