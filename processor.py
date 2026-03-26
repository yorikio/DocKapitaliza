import zipfile
import subprocess
import os
import tempfile
from io import BytesIO
from mailmerge import MailMerge
from utils import monto_a_letra, obtener_fecha_formal

def generar_paquete_zip(df, plantilla_bytes):
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        # Usamos un directorio temporal para la conversión
        with tempfile.TemporaryDirectory() as temp_dir:
            for _, fila in df.iterrows():
                nombre_raw = str(fila['Nombre Cliente']).upper()
                nombre_archivo = nombre_raw.replace(" ", "_")
                
                datos_contrato = {
                    'Nombre_Cliente': nombre_raw,
                    'Monto_Numero': f"${fila['Cuota']:,.2f}",
                    'Monto_Letra': monto_a_letra(fila['Cuota']),
                    'Direccion_Cliente': str(fila['Dirección']).upper(),
                    'Fecha_Emision': obtener_fecha_formal()
                }
                
                # 1. Crear el Word temporal
                path_word = os.path.join(temp_dir, f"PAGARE_{nombre_archivo}.docx")
                with MailMerge(BytesIO(plantilla_bytes)) as document:
                    document.merge(**datos_contrato)
                    document.write(path_word)
                
                # 2. Convertir a PDF usando LibreOffice (Comando de sistema)
                try:
                    subprocess.run([
                        'lowriter', '--headless', '--convert-to', 'pdf', 
                        '--outdir', temp_dir, path_word
                    ], check=True, capture_output=True)
                    
                    path_pdf = path_word.replace(".docx", ".pdf")
                    
                    # 3. Leer el PDF generado y meterlo al ZIP
                    with open(path_pdf, "rb") as f:
                        zip_file.writestr(f"PAGARE_{nombre_archivo}.pdf", f.read())
                        
                except Exception as e:
                    # Si falla el PDF, al menos metemos el Word para no perder el trabajo
                    with open(path_word, "rb") as f:
                        zip_file.writestr(f"PAGARE_{nombre_archivo}.docx", f.read())

    return zip_buffer.getvalue()