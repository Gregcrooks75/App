import re
from openpyxl import load_workbook
import base64
import pandas as pd
import streamlit as st
import io

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
    Estructura_Tabla_Datos = pd.DataFrame({
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
            [], ['Pagos(Ámbito-Subholding)', 'Pagos(Ámbito-Área/Negocio)'], [], [], ['Pagos(Ámbito-Subholding)','Pagos(Ámbito-Área/Negocio)','Nombre','Usuario','Pagos(SÍ/NO)','Pagos(Tipo)','Garantías'], [], ['Pagos(Ámbito-Subholding)'], [], ['Pagos(Ámbito-Subholding)', 'Pagos(Ámbito-Área/Negocio)']],
    })

    return Estructura_Tabla_Datos

def procesar_excel(Estructura_Tabla_Datos, RUTA_ARCHIVO):
    Datos= pd.read_excel(RUTA_ARCHIVO, sheet_name=None, header=None)
    DataFrames= []
    for Hoja, data in Datos.items():
        Hojas_Excluir=['VW', 'AGR SAP', 'Tipo-Subtipo de pago', 'Ámbito']
        if Hoja in Hojas_Excluir:
            continue
        Filas_INICIO_Eliminar= Estructura_Tabla_Datos.loc[Estructura_Tabla_Datos['SUBHOLDING']==Hoja, 'Filas_Inicio_Eliminadas'].item()
        data= data.drop(index=range(Filas_INICIO_Eliminar))
        Filas_FINAL_Eliminar= Estructura_Tabla_Datos.loc[Estructura_Tabla_Datos['SUBHOLDING']==Hoja, 'Filas_Final_Eliminadas'].item()
        data= data.drop(data.tail(Filas_FINAL_Eliminar).index).reset_index(drop=True)
        data.columns= Estructura_Tabla_Datos.loc[Estructura_Tabla_Datos['SUBHOLDING']==Hoja, 'Cabecera'].item()
        data['Usuario']= data['Usuario'].astype(str).str.extract(r'(\d+)', expand=False)
        def extraer_subcadenas(cadena_entrada):
            cuatro_digitos_str= re.findall(r'\d{4}', str(cadena_entrada))
            p_dos_digitos_str= re.findall(r'P\d{2}', str(cadena_entrada))
            return ', '.join(cuatro_digitos_str + p_dos_digitos_str)
        if Hoja=='AGR':
            data['Pagos(Tipo)']= data['Pagos(Límite Máximo)'].apply(extraer_subcadenas)
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
        data['Nombre Hoja']= Hoja
        DataFrames.append(data)
    Columnas_UNION= ['Nombre Hoja', 'Usuario', 'Nombre', 'Pagos(SÍ/NO)', 'Pagos(Tipo)', 'Pagos(Ámbito-Área/Negocio)', 'Pagos Confidenciales', 'Garantías', 'Pagos(Ámbito)', 'Pagos(Límite Máximo)']
    UNION= pd.concat([df[Columnas_UNION] for df in DataFrames], ignore_index=True)
    UNION.to_excel('Hojas_unidas_y_procesadas.xlsx', index=False, sheet_name='Hojas_Unificadas')
    return UNION

# Initialize the dataframe and display it using st.experimental_data_editor
df = init_df()

# Configurar la disposición de la página
st.set_page_config(layout="wide")

st.title("Procesador de Datos Excel")
# Create an editable dataframe
edited_df = st.data_editor(df)


# Barra lateral para subir el archivo Excel
st.sidebar.title("Subir archivo Excel")
ruta_archivo = st.sidebar.file_uploader("", type=['xlsx'])

if ruta_archivo is not None:
    st.sidebar.subheader("Procesar archivo Excel")
    if st.sidebar.button("Procesar", key="boton_procesar"):
        # Convert the BytesIO object to an io.BytesIO object before passing it to your processing function
        datos_procesados = procesar_excel(edited_df, io.BytesIO(ruta_archivo.read()))
        st.subheader("Datos Procesados")

        # Para hacer la visualización compacta, puedes usar una rejilla de datos en lugar de una tabla.
        st.dataframe(datos_procesados)

        # Descargar datos procesados
        st.sidebar.subheader("Descargar Datos Procesados")
        st.sidebar.markdown(generar_link_descarga(datos_procesados), unsafe_allow_html=True)
else:
    st.warning("Por favor, sube un archivo Excel para continuar.")



