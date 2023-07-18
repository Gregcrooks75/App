import re
from openpyxl import load_workbook
import base64
import pandas as pd
import streamlit as st

def generar_link_descarga(df):
    """
    Genera un enlace que permite descargar los datos de un dataframe de pandas dado
    in:  dataframe
    out: string href
    """
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()  # algunas conversiones de cadenas <-> bytes necesarias aquí
    href = f'<a href="data:file/csv;base64,{b64}" download="datos_procesados.csv">Descargar Datos Procesados</a>'
    return href

# Inicialización del DataFrame
def init_df():
    Tabla_Estructura_Datos = pd.DataFrame({
    'SUBHOLDING':[
        'CORP-ESP', 'SPW', 'AGR SAP', 'AGR', 'MEXICO', 'NEO', 'ROKAS', 'IIC', 'INMOB'],
    'Filas_Inicio_Eliminadas':[
        4, 5, 4, 4, 5, 7, 4, 5, 5],
    'Filas_Final_Eliminadas':[
        3, 7, 5, 0, 5, 0, 4, 4, 4],
    'Cabecera':[
        ['Usuario','Nombre','Pagos(SÍ/NO)','Pagos(Tipo)','Pagos(Ámbito-Área/Negocio)','Pagos(Ámbito-Subholding)','Pagos Confidenciales','Garantías','Nuevos Roles'],
        ['Pagos(Ámbito-Subholding)','Pagos(Ámbito-Área/Negocio)','Nombre', 'Usuario','Pagos(SÍ/NO)','Pagos(Tipo)','Pagos(Ámbito)','Pagos(Límite Máximo)','Garantías','Garantías(Límite Máximo)'],
        ['Pagos(Ámbito-Subholding)','Pagos(Ámbito-Área/Negocio)', 'Título','Nombre','Usuario','Aprobación Avangrid','Aprobación UIL','Pagos(SÍ/NO)','Pagos(Tipo)','Pagos(Ámbito)','Garantías','Revisión James Jenkins'],
        ['Pagos(Ámbito-Subholding)','Pagos(Ámbito-Área/Negocio)', 'Título','Nombre','Usuario','Pagos(Límite Máximo)','Aprobación UIL','Pagos(SÍ/NO)','Pagos(Tipo)','Pagos(Ámbito)','Garantías','Revisión James Jenkins'],
        ['Pagos(Ámbito-Subholding)','Pagos(Ámbito-Área/Negocio)','Nombre','Usuario','Pagos(SÍ/NO)','Pagos(Tipo)','Pagos(Ámbito)','Garantías'],
        ['Pagos(Ámbito-Subholding)','Pagos(Ámbito-Área/Negocio)','Nombre','Usuario','Descripción','Pagos(SÍ/NO)','Pagos(Tipo)','Pagos(Ámbito)','Garantías','Info_No_Importante1','Info_No_Importante2','Info_No_Importante3','Info_No_Importante4','Info_No_Importante5'],
        ['Pagos(Ámbito-Subholding)','Nombre','Usuario','Pagos(SÍ/NO)','Pagos(Tipo)','Pagos(Ámbito)','Garantías'],
        ['Pagos(Ámbito-Subholding)','Pagos(Ámbito-Área/Negocio)','Nombre','Usuario','Descripción','Pagos(SÍ/NO)','Pagos(Tipo)','Pagos(Ámbito)','Garantías'],
        ['Pagos(Ámbito-Subholding)','Pagos(Ámbito-Área/Negocio)','Nombre','Usuario','Pagos(SÍ/NO)','Pagos(Tipo)','Pagos(Ámbito)','Garantías']],
    'Rellenar_NANs_SI_NO':[
        'NO', 'SI', 'NO', 'NO', 'SI', 'NO', 'SI', 'NO', 'SI'],
    'Columnas_Rellenar_NANs':[
        '', ['Pagos(Ámbito-Subholding)', 'Pagos(Ámbito-Área/Negocio)'], '', '', ['Pagos(Ámbito-Subholding)','Pagos(Ámbito-Área/Negocio)','Nombre','Usuario','Pagos(SÍ/NO)','Pagos(Tipo)','Garantías'], '', ['Pagos(Ámbito-Subholding)'], '', ['Pagos(Ámbito-Subholding)', 'Pagos(Ámbito-Área/Negocio)']],
    'Eliminar_Columnas_SI_NO':[
        'NO', 'NO', 'NO', 'NO', 'NO', 'SI', 'NO', 'NO', 'NO'],
    'Columnas_Eliminar':[
        '', '', '', '', '', ['Info_No_Importante1','Info_No_Importante2','Info_No_Importante3','Info_No_Importante4','Info_No_Importante5'], '', '', '']
    })

    return Tabla_Estructura_Datos

def procesar_excel(Tabla_Estructura_Datos, RUTA_ARCHIVO):
    Datos= pd.read_excel(RUTA_ARCHIVO, sheet_name=None, header=None)
    DataFrames= []
    for Hoja, data in Datos.items():
        Hojas_Excluir=['VW', 'AGR SAP', 'Tipo-Subtipo de pago', 'Ámbito']
        if Hoja in Hojas_Excluir:
            continue
        Filas_INICIO_Eliminar= Tabla_Estructura_Datos.loc[Tabla_Estructura_Datos['SUBHOLDING']==Hoja, 'Filas_Inicio_Eliminadas'].item()
        data= data.drop(index=range(Filas_INICIO_Eliminar))
        Filas_FINAL_Eliminar= Tabla_Estructura_Datos.loc[Tabla_Estructura_Datos['SUBHOLDING']==Hoja, 'Filas_Final_Eliminadas'].item()
        data= data.drop(data.tail(Filas_FINAL_Eliminar).index).reset_index(drop=True)
        data.columns= Tabla_Estructura_Datos.loc[Tabla_Estructura_Datos['SUBHOLDING']==Hoja, 'Cabecera'].item()
        if Tabla_Estructura_Datos.loc[Tabla_Estructura_Datos['SUBHOLDING']==Hoja, 'Rellenar_NANs_SI_NO'].item()=='SI':
            for Columna in Tabla_Estructura_Datos.loc[Tabla_Estructura_Datos['SUBHOLDING']==Hoja, 'Columnas_Rellenar_NANs'].item():
                data[Columna].fillna(method='ffill', inplace=True)
        data['Usuario']= data['Usuario'].astype(str).str.extract(r'(\d+)', expand=False)
        def extract_substrings(input_str):
            cuatro_digitos_str= re.findall(r'\d{4}', str(input_str))
            p_dos_digitos_str= re.findall(r'P\d{2}', str(input_str))
            return ', '.join(cuatro_digitos_str + p_dos_digitos_str)
        if Hoja=='AGR':
            data['Pagos(Tipo)']= data['Pagos(Límite Máximo)'].apply(extract_substrings)
        if 'Pagos(SÍ/NO)' in data.columns:
            data['Pagos(SÍ/NO)'] = data['Pagos(SÍ/NO)'].apply(lambda x: 'SI' if str(x).replace(" ", "").replace("'", "") == 'SI' or str(x).replace(" ", "").replace("'", "") == 'SÍ' or x == 1 else 'NO')
        if 'Garantías' in data.columns:
            data['Garantías'] = data['Garantías'].apply(lambda x: 'SI' if str(x).replace(" ", "").replace("'", "") == 'SI' or str(x).replace(" ", "").replace("'", "") == 'SÍ' or x == 1 else 'NO')
        if Hoja=='ROKAS':
            data['Pagos(Ámbito-Área/Negocio)']=''
        if Hoja=='CORP-ESP':
            data['Pagos(Ámbito)']=''
        if 'Pagos Confidenciales' not in data.columns:
            data['Pagos Confidenciales']=''
        if 'Pagos(Límite Máximo)' not in data.columns:
            data['Pagos(Límite Máximo)']='NO INDICADO'
        if Tabla_Estructura_Datos.loc[Tabla_Estructura_Datos['SUBHOLDING']==Hoja, 'Eliminar_Columnas_SI_NO'].item()=='SI':
            for Columna in Tabla_Estructura_Datos.loc[Tabla_Estructura_Datos['SUBHOLDING']==Hoja, 'Columnas_Eliminar'].item():
                data.drop(labels=[Columna], axis=1, inplace=True)
        data['Nombre Hoja']= Hoja
        DataFrames.append(data)
    Columnas_UNION= ['Nombre Hoja', 'Usuario', 'Nombre', 'Pagos(SÍ/NO)', 'Pagos(Tipo)', 'Pagos(Ámbito-Área/Negocio)', 'Pagos Confidenciales', 'Garantías', 'Pagos(Ámbito)', 'Pagos(Límite Máximo)']
    UNION= pd.concat([df[Columnas_UNION] for df in DataFrames], ignore_index=True)
    UNION.to_excel('Hojas_unidas_y_procesadas.xlsx', index=False)
    return UNION

# Inicializar DataFrame
df = init_df()

# Configurar la disposición de la página
st.set_page_config(layout="wide")

st.title("Procesador de Datos Excel")

# Barra lateral para subir el archivo Excel
st.sidebar.title("Subir archivo Excel")
ruta_archivo = st.sidebar.file_uploader("", type=['xlsx'])

# Mostrar los parámetros para la selección del usuario
st.subheader("Parámetros")
df_mostrar = df.copy()
df_mostrar = df.drop(columns='Columnas_Rellenar_NANs')

indice_seleccionado = st.selectbox("Seleccione fila para editar:", df.index)
fila_seleccionada = df.iloc[indice_seleccionado]

if st.button("Edit selected row", key="edit_button"):
    with st.form("Edit Form"):
        new_values = []
        for column in df.columns:
            new_value = st.text_input(f"New value for {column}", selected_row[column])
            new_values.append(new_value)

        submit_button = st.form_submit_button("Submit Changes")
        if submit_button:
            df_mostrar.loc[indice_seleccionado] = new_values
            df.at[indice_seleccionado] = new_values
    st.table(st.session_state.df)

# Si se sube un archivo, proceder con el procesamiento y visualización
if ruta_archivo is not None:
    st.sidebar.subheader("Procesar archivo Excel")
    if st.sidebar.button("Procesar", key="boton_procesar"):
        datos_procesados = procesar_excel(df, ruta_archivo)
        st.subheader("Datos Procesados")
        
        # Para hacer la visualización compacta, puedes usar una rejilla de datos en lugar de una tabla.
        st.dataframe(datos_procesados)

        # Descargar datos procesados
        st.sidebar.subheader("Descargar Datos Procesados")
        st.sidebar.markdown(generar_link_descarga(datos_procesados), unsafe_allow_html=True)
else:
    st.warning("Por favor, sube un archivo Excel para continuar.")
