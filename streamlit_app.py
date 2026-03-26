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

st.info("Sube el Excel con las columnas: 'Cuota', 'Nombre Cliente' y 'Dirección'")

uploaded_data = st.file_uploader("Sube tu Base de Datos (.xlsx)", type="xlsx")

if uploaded_data:
    df = pd.read_excel(uploaded_data)
    
    # Validación rápida de columnas para evitar errores de ejecución
    columnas_requeridas = ['Cuota', 'Nombre Cliente', 'Dirección']
    if all(col in df.columns for col in columnas_requeridas):
        st.success(f"✅ Se cargaron {len(df)} registros correctamente.")
        
        if st.button("Generar y Descargar Paquete ZIP"):
            with st.spinner("Procesando documentos..."):
                # Leemos la plantilla estática del disco
                with open(NOMBRE_PLANTILLA, "rb") as f:
                    plantilla_bytes = f.read()
                
                zip_final = generar_paquete_zip(df, plantilla_bytes)
                
                st.download_button(
                    label="📥 Descargar todos los Pagarés",
                    data=zip_final,
                    file_name="Paquete_Pagares_Kapitaliza.zip",
                    mime="application/zip"
                )
    else:
        st.warning(f"⚠️ El Excel debe contener exactamente las columnas: {', '.join(columnas_requeridas)}")