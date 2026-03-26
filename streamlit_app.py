import streamlit as st
import pandas as pd
from mailmerge import MailMerge
from io import BytesIO

st.title("Generador de Documentos Kapitaliza")

# 1. Carga de archivos
uploaded_template = st.file_uploader("Sube tu plantilla Word (.docx)", type="docx")
uploaded_data = st.file_uploader("Sube tu archivo Excel (.xlsx)", type="xlsx")

if uploaded_template and uploaded_data:
    df = pd.read_excel(uploaded_data)
    
    # Mostrar vista previa de los datos
    st.write("Datos detectados:", df.head())
    
    if st.button("Generar Documentos"):
        # Leer la plantilla desde memoria
        template_bytes = BytesIO(uploaded_template.read())
        
        # Crear un archivo Word final en memoria para contener todos los cruces
        output_stream = BytesIO()
        
        with MailMerge(template_bytes) as document:
            # Convertir el DataFrame a una lista de diccionarios (clave = nombre del MergeField)
            data_to_merge = df.to_dict(orient='records')
            
            # Realizar el cruce (crea una página nueva por cada fila del Excel)
            document.merge_pages(data_to_merge)
            document.write(output_stream)
        
        # Botón para descargar el Word resultante
        st.download_button(
            label="Descargar Documentos Generados",
            data=output_stream.getvalue(),
            file_name="correspondencia_final.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )