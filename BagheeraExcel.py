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


# 🆔 OBTENER PRÓXIMO ID
def obtener_proximo_id():
    """Obtiene el ID del próximo empleado (máximo ID + 1)"""
    df = pd.read_excel(RUTA)
    df.columns = df.columns.str.strip().str.upper()
    
    # Si la tabla está vacía, comenzar con ID 1
    if df.empty:
        return 1
    
    # Obtener el máximo ID y sumarle 1
    try:
        max_id = pd.to_numeric(df['ID'], errors='coerce').max()
        if pd.isna(max_id):
            return 1
        return int(max_id) + 1
    except:
        return 1


# ➕ AGREGAR NUEVO EMPLEADO
def agregar_empleado(empleado_data):
    """Agrega un nuevo empleado al Excel"""
    try:
        df = pd.read_excel(RUTA)
        df.columns = df.columns.str.strip().str.upper()

        # Obtener próximo ID
        proximo_id = obtener_proximo_id()
        empleado_data['ID'] = proximo_id

        # Crear una fila nueva con los datos
        nueva_fila = pd.DataFrame([empleado_data])

        # Asegurar que las columnas coincidan
        for col in df.columns:
            if col not in nueva_fila.columns:
                nueva_fila[col] = ""

        # Reordenar las columnas de la nueva fila igual al DataFrame original
        nueva_fila = nueva_fila[df.columns]

        # Concatenar y guardar
        df = pd.concat([df, nueva_fila], ignore_index=True)
        df.to_excel(RUTA, index=False)

        # Verificar que se guardó correctamente
        df_verif = pd.read_excel(RUTA)
        if len(df_verif) == len(df):
            print(f"✅ Empleado agregado con ID: {proximo_id}")
            print(f"📊 Total empleados ahora: {len(df)}")
            return proximo_id
        else:
            raise Exception("Error: El empleado no se guardó correctamente")

    except Exception as e:
        print(f"❌ ERROR al agregar empleado: {str(e)}")
        raise e