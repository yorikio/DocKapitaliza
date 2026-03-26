from datetime import datetime
from num2words import num2words

def monto_a_letra(numero):
    """Convierte un float a formato moneda legal mexicana."""
    entero = int(numero)
    centavos = int(round((numero - entero) * 100))
    texto_entero = num2words(entero, lang='es').upper()
    return f"({texto_entero} PESOS {centavos:02d}/100 M.N.)"

def obtener_fecha_formal():
    """Genera la fecha actual en el formato requerido por el pagaré."""
    meses = [
        "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
        "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
    ]
    ahora = datetime.now()
    return f"{ahora.day} DE {meses[ahora.month - 1]} DEL {ahora.year}"