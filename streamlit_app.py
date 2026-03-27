import streamlit as st
import pandas as pd
import os
from processor import generar_paquete_zip

st.set_page_config(page_title="Generador de Pagarés | Kapitaliza", layout="wide")

st.title("🚀 Sistema de Pagarés Automáticos - Kapitaliza")

# Nombre del archivo que debe estar en tu repositorio
NOMBRE_PLANTILLA = "plantilla_pagare_kapitaliza.docx"

# Verificamos si la plantilla existe en la raíz
if not os.path.exists(NOMBRE_PLANTILLA):
    st.error(f"❌ Error: No se encontró el archivo '{NOMBRE_PLANTILLA}' en el servidor.")
    st.stop()

# Lista de columnas que el usuario debe tener en su Excel
columnas_requeridas = [
    'Cuota', 
    'Nombre Cliente', 
    'Calle Cliente', 
    'Número Exterior Cliente', 
    'Colonia Cliente', 
    'Municipio Cliente', 
    'Localidad Cliente', 
    'Código Postal Cliente', 
    'Estado Cliente',
    'Pais Cliente'
]

# Nota: Dejamos 'Número Interior Cliente' fuera de la validación obligatoria 
# para que el sistema sea más flexible si no todos tienen interior.

st.info(f"Sube el Excel con las siguientes columnas: {', '.join(columnas_requeridas)}")

uploaded_data = st.file_uploader("Sube tu Base de Datos (.xlsx)", type="xlsx")

if uploaded_data:
    df = pd.read_excel(uploaded_data)
    
    # Validación de columnas obligatorias
    missing_cols = [col for col in columnas_requeridas if col not in df.columns]
    
    if not missing_cols:
        st.success(f"✅ Se cargaron {len(df)} registros correctamente.")
        
        # Vista previa para que el usuario confirme
        with st.expander("Ver vista previa de los datos"):
            st.write(df.head())
        
        if st.button("Generar y Descargar Paquete ZIP"):
            with st.spinner("Procesando documentos y convirtiendo a PDF..."):
                try:
                    # Leemos la plantilla estática del disco
                    with open(NOMBRE_PLANTILLA, "rb") as f:
                        plantilla_bytes = f.read()
                    
                    # Llamamos al procesador que genera Word y PDF
                    zip_final = generar_paquete_zip(df, plantilla_bytes)
                    
                    st.download_button(
                        label="📥 Descargar todos los Pagarés (Word y PDF)",
                        data=zip_final,
                        file_name="Paquete_Pagares_Kapitaliza.zip",
                        mime="application/zip"
                    )
                except Exception as e:
                    st.error(f"Hubo un error al generar el paquete: {e}")
    else:
        st.error(f"⚠️ Al Excel le faltan columnas obligatorias: {', '.join(missing_cols)}")