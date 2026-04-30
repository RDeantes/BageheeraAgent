import unicodedata
from datetime import datetime, timedelta
from google_sheets import get_sheet


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
    sheet = get_sheet()
    registros = sheet.get_all_records()

    nombre_limpio = limpiar(nombre)
    palabras_input = nombre_limpio.split()

    # =========================================================
    # 🔥 1. MATCH EXACTO (PRIORIDAD MÁXIMA)
    # =========================================================
    for row in registros:
        nombre_excel = limpiar(row.get("NOMBRE", ""))
        if nombre_excel == nombre_limpio:
            return row

    # =========================================================
    # 🔥 2. MATCH POR PALABRAS COMPLETAS (MÍNIMO 2)
    # =========================================================
    mejor_match = None
    mejor_score = 0

    for row in registros:
        nombre_excel = limpiar(row.get("NOMBRE", ""))
        palabras_excel = nombre_excel.split()

        # contar coincidencias reales (palabras completas)
        coincidencias = sum(1 for p in palabras_input if p in palabras_excel)

        if coincidencias > mejor_score:
            mejor_score = coincidencias
            mejor_match = row

    # =========================================================
    # 🔥 VALIDACIÓN FINAL
    # =========================================================
    if mejor_score >= 2:
        return mejor_match

    return None


# =========================================================
# 🔄 ACTUALIZAR VIGENCIA
# =========================================================
def actualizar_vigencia(nombre, nueva_fecha):
    sheet = get_sheet()
    registros = sheet.get_all_records()

    for i, row in enumerate(registros):
        if limpiar(row.get("NOMBRE", "")) == limpiar(nombre):
            sheet.update_cell(i + 2, 7, str(nueva_fecha))
            print(f"✅ Vigencia actualizada: {nombre}")
            return True

    print("⚠️ No se encontró empleado")
    return False


# =========================================================
# 🆔 OBTENER ID
# =========================================================
def obtener_proximo_id():
    sheet = get_sheet()
    registros = sheet.get_all_records()

    if not registros:
        return 1

    try:
        ids = [int(r["ID"]) for r in registros if str(r.get("ID", "")).isdigit()]
        return max(ids) + 1 if ids else 1
    except:
        return 1


# =========================================================
# ➕ AGREGAR EMPLEADO
# =========================================================
def agregar_empleado(data):
    sheet = get_sheet()

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

    print("✅ GUARDADO EN GOOGLE SHEETS:", fila)
    return nuevo_id


# =========================================================
# 📄 CONTRATO
# =========================================================
def contrato_desde_excel(nombre, generar_contrato_func, personalidad_func):

    persona = buscar_empleado(nombre)

    if persona is None:
        return personalidad_func("❌ No encontré al empleado")

    hoy = datetime.now()
    nueva_vigencia = hoy + timedelta(days=180)

    datos = {
         "tipo": datos.get("tipo", ""),
        "jornada": datos.get("jornada", ""),
        "duracion": datos.get("duracion", ""),  # 🔥 FIX
        "fecha_inicio": hoy.strftime("%Y-%m-%d"),
        "fecha_termino": nueva_vigencia.strftime("%Y-%m-%d"),
        "nombre": persona.get("NOMBRE", ""),
        "nacionalidad": persona.get("NACIONALIDAD", ""),
        "sexo": persona.get("SEXO", ""),
        "curp": persona.get("CURP", ""),
        "domicilio": persona.get("DOMICILIO", ""),
        "puesto": persona.get("PUESTO", ""),
        "dias": "LUNES A SABADO"
    }

    pdf = generar_contrato_func(datos)

    actualizar_vigencia(persona.get("NOMBRE", ""), nueva_vigencia)

    return pdf