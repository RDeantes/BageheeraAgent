import unicodedata
from datetime import datetime, timedelta
from google_sheets import sheet


# =========================================================
# 🔤 LIMPIAR TEXTO
# =========================================================
def limpiar(texto):
    texto = str(texto).lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('utf-8')
    return texto.strip()


# =========================================================
# 🔍 BUSCAR EMPLEADO
# =========================================================
def buscar_empleado(nombre):
    nombre = limpiar(nombre)
    registros = sheet.get_all_records()

    # 🔥 match exacto primero
    for row in registros:
        if limpiar(row["NOMBRE"]) == nombre:
            return row

    # 🔥 match flexible
    palabras = nombre.split()

    mejor = None
    mejor_score = 0

    for row in registros:
        nombre_excel = limpiar(row["NOMBRE"])
        score = sum(1 for p in palabras if p in nombre_excel)

        if score > mejor_score:
            mejor_score = score
            mejor = row

    if mejor_score >= 2:
        return mejor

    return None


# =========================================================
# 🔄 ACTUALIZAR VIGENCIA
# =========================================================
def actualizar_vigencia(nombre, nueva_fecha):
    registros = sheet.get_all_records()

    for i, row in enumerate(registros):
        if limpiar(row["NOMBRE"]) == limpiar(nombre):
            sheet.update_cell(i + 2, 7, str(nueva_fecha))
            print(f"✅ Vigencia actualizada: {nombre}")
            return True

    print("⚠️ No se encontró empleado para actualizar")
    return False


# =========================================================
# 🆔 OBTENER PRÓXIMO ID
# =========================================================
def obtener_proximo_id():
    registros = sheet.get_all_records()

    if not registros:
        return 1

    try:
        ids = [int(r["ID"]) for r in registros if str(r["ID"]).isdigit()]
        return max(ids) + 1 if ids else 1
    except:
        return 1


# =========================================================
# ➕ AGREGAR EMPLEADO
# =========================================================
def agregar_empleado(data):

    nuevo_id = obtener_proximo_id()

    fila = [
        nuevo_id,
        data.get("NOMBRE", ""),
        data.get("AREA", ""),
        data.get("PUESTO", ""),
        data.get("FECHA_INGRESO", ""),
        data.get("TIPO_CONTRATO", ""),
        data.get("VIGENCIA", ""),
        data.get("NACIONALIDAD", ""),
        data.get("SEXO", ""),
        data.get("CURP", ""),
        data.get("DOMICILIO", "")
    ]

    sheet.append_row(fila)

    print(f"✅ Empleado agregado ID: {nuevo_id}")
    return nuevo_id


# =========================================================
# 🤖 CONTRATO DESDE EXCEL (AHORA SHEETS)
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

    actualizar_vigencia(persona['NOMBRE'], nueva_vigencia)

    return pdf