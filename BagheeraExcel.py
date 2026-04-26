import pandas as pd
import os
from datetime import datetime, timedelta

# 📁 ruta del Excel
RUTA = os.path.join(os.path.dirname(__file__), 'CONCENTRADO_EMPLEADOS.xlsx')


import unicodedata

def limpiar(texto):
    texto = str(texto).lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('utf-8')
    return texto


def buscar_empleado(nombre):
    df = pd.read_excel(RUTA)
    df.columns = df.columns.str.strip().str.upper()

    palabras_busqueda = limpiar(nombre).split()

    for _, row in df.iterrows():
        nombre_excel = limpiar(row['NOMBRE'])
        
        coincidencias = sum(1 for p in palabras_busqueda if p in nombre_excel)

        # 🔥 si al menos 2 palabras coinciden → es la persona
        if coincidencias >= 2:
            return row

    return None

# 🔄 ACTUALIZAR VIGENCIA
def actualizar_vigencia(nombre, nueva_fecha):
    df = pd.read_excel(RUTA)
    df.columns = df.columns.str.strip().str.upper()

    for i, row in df.iterrows():
        if nombre.lower() in str(row['NOMBRE']).lower():
            df.at[i, 'VIGENCIA'] = nueva_fecha

    df.to_excel(RUTA, index=False)


# 🤖 FUNCIÓN PRINCIPAL DE BAGIRA
def contrato_desde_excel(nombre, generar_contrato_func, personalidad_func):

    persona = buscar_empleado(nombre)

    if persona is None:
        return personalidad_func("❌ No encontré al empleado")

    hoy = datetime.now()
    nueva_vigencia = hoy + timedelta(days=180)

    datos = {
        "tipo": "TEMPORAL",
        "jornada": "COMPLETA",
        "duracion": "6 MESES",
        "fecha_inicio": hoy.strftime("%Y-%m-%d"),
        "fecha_termino": nueva_vigencia.strftime("%Y-%m-%d"),
        "nombre": persona['NOMBRE'],
        "nacionalidad": persona.get('NACIONALIDAD', 'MEXICANA'),
        "sexo": persona.get('SEXO', 'M'),
        "curp": persona.get('CURP', ''),
        "domicilio": persona.get('DOMICILIO', ''),
        "puesto": persona.get('PUESTO', 'EMPLEADO'),
        "dias": "LUNES A SABADO"
    }

    pdf = generar_contrato_func(datos)

    actualizar_vigencia(nombre, nueva_vigencia)

    return pdf