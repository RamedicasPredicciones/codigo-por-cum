import streamlit as st
import pandas as pd
from io import BytesIO
from rapidfuzz import fuzz, process

# Cargar datos de Ramedicas desde Google Drive
@st.cache_data  # Cachear los datos para evitar recargarlos en cada ejecución
def load_ramedicas_data():
    # URL del archivo Excel en Google Drive
    ramedicas_url = (
        "https://docs.google.com/spreadsheets/d/1Y9SgliayP_J5Vi2SdtZmGxKWwf1iY7ma/export?format=xlsx&sheet=Hoja1"
    )
    # Leer el archivo Excel desde la URL
    ramedicas_df = pd.read_excel(ramedicas_url, sheet_name="Hoja1")
    # Retornar solo las columnas relevantes
    return ramedicas_df[['cum', 'codart', 'nomart']]  # Incluir 'cum', 'codart' y 'nomart'

# Preprocesar CUM para que siempre esté en minúsculas
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
                'codart_ramedicas': candidate_row['codart'],  # Código del producto
                'nomart_ramedicas': candidate_row['nomart'],  # Nombre del producto
                'score': score
            }

    # Si no hay coincidencias completas, se devuelve la mejor aproximación
    if not best_match and matches:
        best_match = {
            'cum_cliente': client_cum,
            'cum_ramedicas': ramedicas_df.iloc[matches[0][2]]['cum'],  # CUM de Ramedicas
            'codart_ramedicas': ramedicas_df.iloc[matches[0][2]]['codart'],  # Código del producto
            'nomart_ramedicas': ramedicas_df.iloc[matches[0][2]]['nomart'],  # Nombre del producto
            'score': matches[0][1]  # Puntuación de la mejor coincidencia
        }

    return best_match

# Interfaz de Streamlit
st.markdown(
    """
    <h1 style="text-align: center; color: #FF5800; font-family: Arial, sans-serif;">
        RAMEDICAS S.A.S.
    </h1>
    <h3 style="text-align: center; font-family: Arial, sans-serif; color: #3A86FF;">
        Homologador de Productos por CUM
    </h3>
    <p style="text-align: center; font-family: Arial, sans-serif; color: #6B6B6B;">
         Esta herramienta te permite buscar y consultar los códigos de productos por medio de su CUM.
    </p>
    """, unsafe_allow_html=True
)

# URL de la imagen
image_url = "https://drive.google.com/uc?export=view&id=1CnszOM0U5iq-tmP_b_OoeGhMvZChyOLt"

# CSS para posicionar la imagen
st.markdown(
    f"""
    <style>
        .custom-image {{
            position: fixed;
            bottom: 10px;
            left: 10px;
            width: 150px; /* Ajusta el tamaño de la imagen */
            border-radius: 10px; /* Esquinas redondeadas opcionales */
            z-index: 100;
        }}
    </style>
    <img src="{image_url}" class="custom-image">
    """,
    unsafe_allow_html=True
)

# Opción para actualizar la base de datos y limpiar el caché
if st.button("Actualizar base de datos"):
    st.cache_data.clear()  # Limpiar el caché para cargar los datos de nuevo

# Espacio de entrada para subir archivo con los CUM de los clientes
uploaded_file = st.file_uploader("Sube tu archivo con los cum a buscar (El nombre de la columna debe ser  \"cum\")", type="xlsx")

if uploaded_file:
    # Leer el archivo subido con los CUM de los clientes
    client_cums_df = pd.read_excel(uploaded_file)
    
    # Verificar si la columna 'cum' está presente en el archivo subido
    if 'cum' not in client_cums_df.columns:
        st.error("El archivo debe contener una columna llamada 'cum'.")
    else:
        # Cargar los datos de productos de Ramedicas
        ramedicas_df = load_ramedicas_data()
        
        # Lista para almacenar los resultados de la homologación
        results = []
        
        # Se crea la iteración sobre cada CUM de cliente para encontrar la mejor coincidencia
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
