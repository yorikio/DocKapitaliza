import zipfile
from io import BytesIO
from mailmerge import MailMerge
from utils import monto_a_letra, obtener_fecha_formal

def generar_paquete_zip(df, plantilla_bytes):
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for _, fila in df.iterrows():
            # Limpieza de nombre para el archivo
            nombre_raw = str(fila['Nombre Cliente']).upper()
            nombre_archivo = nombre_raw.replace(" ", "_")
            
            # Mapeo de campos según tu Excel y Plantilla
            datos_contrato = {
                'Nombre_Cliente': nombre_raw,
                'Monto_Numero': f"${fila['Cuota']:,.2f}",
                'Monto_Letra': monto_a_letra(fila['Cuota']),
                'Direccion_Cliente': str(fila['Dirección']).upper(),
                'Fecha_Emision': obtener_fecha_formal()
            }
            
            # Generación individual en memoria
            doc_io = BytesIO(plantilla_bytes)
            with MailMerge(doc_io) as document:
                document.merge(**datos_contrato)
                cliente_buffer = BytesIO()
                document.write(cliente_buffer)
                
                # Agregar al ZIP
                zip_file.writestr(f"PAGARE_{nombre_archivo}.docx", cliente_buffer.getvalue())

    return zip_buffer.getvalue()