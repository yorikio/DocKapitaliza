import zipfile
import subprocess
import os
import tempfile
from io import BytesIO
from mailmerge import MailMerge
from utils import monto_a_letra, obtener_fecha_formal
import pandas as pd

def generar_paquete_zip(df, plantilla_bytes):
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        with tempfile.TemporaryDirectory() as temp_dir:
            for _, fila in df.iterrows():
                nombre_raw = str(fila['Nombre Cliente']).upper()
                nombre_archivo = nombre_raw.replace(" ", "_")
                # Construcción dinámica de la dirección
                partes_direccion = [
                    str(fila['Calle Cliente']),
                    f"#{fila['Número Exterior Cliente']}",
                    f"Int. {fila['Número Interior Cliente']}" if pd.notna(fila['Número Interior Cliente']) else "",
                    f"Col. {fila['Colonia Cliente']}",
                    str(fila['Localidad Cliente']),
                    str(fila['Municipio Cliente']),
                    f"C.P. {fila['Código Postal Cliente']}",
                    str(fila['Estado Cliente']),
                    str(fila['Pais Cliente'])
                ]
                
                # Unimos solo las partes que no estén vacías con una coma y espacio
                direccion_completa = ", ".join([p for p in partes_direccion if str(p).strip() and p != "nan"])
                
                datos_contrato = {
                    'Nombre_Cliente': nombre_raw,
                    'Monto_Numero': f"${fila['Cuota']:,.2f}",
                    'Monto_Letra': monto_a_letra(fila['Cuota']),
                    'Direccion_Cliente': direccion_completa.upper(),
                    'Fecha_Emision': obtener_fecha_formal()
                }
                
                # 1. Crear el archivo Word en el directorio temporal
                path_word = os.path.join(temp_dir, f"PAGARE_{nombre_archivo}.docx")
                with MailMerge(BytesIO(plantilla_bytes)) as document:
                    document.merge(**datos_contrato)
                    document.write(path_word)
                
                # --- CAMBIO AQUÍ: Guardar el Word en el ZIP siempre ---
                with open(path_word, "rb") as f_word:
                    zip_file.writestr(f"PAGARE_{nombre_archivo}.docx", f_word.read())
                
                # 2. Intentar la conversión a PDF
                try:
                    # En Streamlit Cloud (Linux) el comando suele ser 'libreoffice'
                    subprocess.run([
                        'libreoffice', '--headless', '--convert-to', 'pdf', 
                        '--outdir', temp_dir, path_word
                    ], check=True, capture_output=True)
                    
                    path_pdf = path_word.replace(".docx", ".pdf")
                    
                    # 3. Guardar el PDF generado en el ZIP
                    if os.path.exists(path_pdf):
                        with open(path_pdf, "rb") as f_pdf:
                            zip_file.writestr(f"PAGARE_{nombre_archivo}.pdf", f_pdf.read())
                        
                except Exception as e:
                    # Si el PDF falla, ya tenemos el Word arriba, 
                    # podrías opcionalmente registrar el error para debugging
                    print(f"No se pudo generar PDF para {nombre_raw}: {e}")

    return zip_buffer.getvalue()