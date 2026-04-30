import pandas as pd
import os
from datetime import datetime, timedelta
import unicodedata

# 📁 ruta del Excel
RUTA = os.path.join(os.path.dirname(__file__), 'CONCENTRADO_EMPLEADOS.xlsx')


# =========================================================
# 🔤 LIMPIAR TEXTO
# =========================================================
def limpiar(texto):
    texto = str(texto).lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('utf-8')
    return texto.strip()


# =========================================================
# 🔍 BUSCAR EMPLEADO (PRECISO)
# =========================================================
def buscar_empleado(nombre):
    df = pd.read_excel(RUTA)
    df.columns = df.columns.str.strip().str.upper()

    nombre_limpio = limpiar(nombre)

    # 🔥 1. MATCH EXACTO
    for _, row in df.iterrows():
        if limpiar(row['NOMBRE']) == nombre_limpio:
            return row

    # 🔥 2. MATCH POR SCORE
    palabras = nombre_limpio.split()

    mejor_match = None
    mejor_score = 0

    for _, row in df.iterrows():
        nombre_excel = limpiar(row['NOMBRE'])
        score = sum(1 for p in palabras if p in nombre_excel)

        if score > mejor_score:
            mejor_score = score
            mejor_match = row

    if mejor_score >= 2:
        return mejor_match

    return None


# =========================================================
# 🔄 ACTUALIZAR VIGENCIA (SEGURO)
# =========================================================
def actualizar_vigencia(nombre_excel, nueva_fecha):
    df = pd.read_excel(RUTA)
    df.columns = df.columns.str.strip().str.upper()

    actualizado = False

    for i, row in df.iterrows():
        if limpiar(row['NOMBRE']) == limpiar(nombre_excel):
            df.at[i, 'VIGENCIA'] = nueva_fecha
            actualizado = True
            break

    if actualizado:
        df.to_excel(RUTA, index=False)
        print(f"✅ Vigencia actualizada: {nombre_excel}")
    else:
        print("⚠️ No se encontró empleado para actualizar")


# =========================================================
# 🆔 OBTENER PRÓXIMO ID
# =========================================================
def obtener_proximo_id():
    if not os.path.exists(RUTA):
        return 1

    df = pd.read_excel(RUTA)
    df.columns = df.columns.str.strip().str.upper()

    if df.empty:
        return 1

    max_id = pd.to_numeric(df['ID'], errors='coerce').max()

    if pd.isna(max_id):
        return 1

    return int(max_id) + 1


# =========================================================
# ➕ AGREGAR EMPLEADO (ROBUSTO)
# =========================================================
def agregar_empleado(empleado_data):
    try:
        columnas = [
            'ID','NOMBRE','AREA','PUESTO','FECHA_INGRESO',
            'TIPO_CONTRATO','VIGENCIA','NACIONALIDAD',
            'SEXO','CURP','DOMICILIO'
        ]

        # 🔥 crear archivo si no existe
        if not os.path.exists(RUTA):
            df = pd.DataFrame(columns=columnas)
        else:
            df = pd.read_excel(RUTA)
            df.columns = df.columns.str.strip().str.upper()

        # 🔥 normalizar datos
        empleado_data = {k.upper(): v for k, v in empleado_data.items()}

        # 🔥 ID automático
        nuevo_id = obtener_proximo_id()
        empleado_data['ID'] = nuevo_id

        # 🔥 asegurar columnas
        for col in columnas:
            if col not in empleado_data:
                empleado_data[col] = ""

        nueva_fila = pd.DataFrame([empleado_data])[columnas]

        df = pd.concat([df, nueva_fila], ignore_index=True)

        df.to_excel(RUTA, index=False)

        print(f"✅ Empleado agregado ID: {nuevo_id}")
        return nuevo_id

    except Exception as e:
        print(f"❌ Error al agregar empleado: {e}")
        return None


# =========================================================
# 🤖 CONTRATO DESDE EXCEL
# =========================================================
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
        "nacionalidad": persona.get('NACIONALIDAD', ''),
        "sexo": persona.get('SEXO', ''),
        "curp": persona.get('CURP', ''),
        "domicilio": persona.get('DOMICILIO', ''),
        "puesto": persona.get('PUESTO', ''),
        "dias": "LUNES A SABADO"
    }

    pdf = generar_contrato_func(datos)

    # 🔥 actualizar con nombre EXACTO
    actualizar_vigencia(persona['NOMBRE'], nueva_vigencia)

    return pdf