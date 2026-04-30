import os
import subprocess
from datetime import datetime
from docx import Document

from BagheeraExcel import contrato_desde_excel, agregar_empleado


# 🧠 MEMORIA GLOBAL
estado_contrato = {}
estado_empleado = {}


# =========================================================
# 🐱 PERSONALIDAD
# =========================================================
def personalidad_bagheera(respuesta):
    return f"🐱:\n{respuesta}\n\n“La fuerza no está en rugir… está en saber cuándo moverte en silencio.” 🌑"


# =========================================================
# 🔤 NORMALIZAR TEXTO (PRO)
# =========================================================
def normalizar(texto):
    texto = texto.lower().strip()
    reemplazos = {
        "á": "a", "é": "e", "í": "i",
        "ó": "o", "ú": "u"
    }
    for k, v in reemplazos.items():
        texto = texto.replace(k, v)
    return texto


# =========================================================
# 🎯 ROUTER PRINCIPAL
# =========================================================
def procesar(mensaje):
    global estado_contrato, estado_empleado

    mensaje = normalizar(mensaje)

    # 🔥 PRIORIDAD: EMPLEADO
    if estado_empleado.get("activo"):
        return flujo_agregar_empleado(mensaje)

    # 🔥 DESPUÉS CONTRATO
    if estado_contrato.get("activo"):
        return flujo_contrato(mensaje)

    # =====================================================
    # COMANDOS
    # =====================================================

    # 🌴 VACACIONES (tolerante)
    if "vacacion" in mensaje or "vacac" in mensaje:
        return revisar_vacaciones_por_mes(mensaje)

    # 👤 AGREGAR EMPLEADO
    if "agregar empleado" in mensaje:
        estado_contrato = {}
        estado_empleado = {"activo": True, "paso": 1}
        return personalidad_bagheera("Nombre del empleado:")

    # 📄 NUEVO CONTRATO
    if "nuevo contrato" in mensaje:
        estado_empleado = {}
        estado_contrato = {"activo": True}
        return personalidad_bagheera("¿Tipo de contrato? (temporal / permanente)")

    return personalidad_bagheera("No entendí la orden.")


# =========================================================
# 🧠 FLUJO CONTRATO
# =========================================================
def flujo_contrato(mensaje):
    global estado_contrato

    if "tipo" not in estado_contrato:
        if "temporal" in mensaje:
            estado_contrato["tipo"] = "TEMPORAL"
        elif "permanente" in mensaje:
            estado_contrato["tipo"] = "PERMANENTE"
        else:
            return personalidad_bagheera("Responde: temporal o permanente")

        return personalidad_bagheera("¿Jornada completa o parcial?")

    elif "jornada" not in estado_contrato:
        if "completa" in mensaje:
            estado_contrato["jornada"] = "COMPLETA"
        elif "parcial" in mensaje:
            estado_contrato["jornada"] = "PARCIAL"
        else:
            return personalidad_bagheera("Responde: completa o parcial")

        return personalidad_bagheera("Duración del contrato:")

    elif "duracion" not in estado_contrato:
        estado_contrato["duracion"] = mensaje.upper()
        return personalidad_bagheera("Fecha inicio (YYYY-MM-DD):")

    elif "fecha_inicio" not in estado_contrato:
        estado_contrato["fecha_inicio"] = mensaje
        return personalidad_bagheera("Fecha fin (YYYY-MM-DD):")

    elif "fecha_termino" not in estado_contrato:
        estado_contrato["fecha_termino"] = mensaje
        return personalidad_bagheera("Nombre del empleado:")

    elif "nombre" not in estado_contrato:
        estado_contrato["nombre"] = mensaje.upper()

        datos = estado_contrato.copy()
        estado_contrato = {}

        return contrato_desde_excel(datos, generar_contrato, personalidad_bagheera)


# =========================================================
# 👤 FLUJO AGREGAR EMPLEADO
# =========================================================
def flujo_agregar_empleado(mensaje):
    global estado_empleado

    paso = estado_empleado["paso"]

    if paso == 1:
        estado_empleado["NOMBRE"] = mensaje.upper()
        estado_empleado["paso"] = 2
        return personalidad_bagheera("Área:")

    elif paso == 2:
        estado_empleado["AREA"] = mensaje.upper()
        estado_empleado["paso"] = 3
        return personalidad_bagheera("Puesto:")

    elif paso == 3:
        estado_empleado["PUESTO"] = mensaje.upper()
        estado_empleado["paso"] = 4
        return personalidad_bagheera("Fecha ingreso (YYYY-MM-DD):")

    elif paso == 4:
        estado_empleado["FECHA_INGRESO"] = mensaje
        estado_empleado["paso"] = 5
        return personalidad_bagheera("Tipo contrato:")

    elif paso == 5:
        estado_empleado["TIPO_CONTRATO"] = mensaje.upper()
        estado_empleado["paso"] = 6
        return personalidad_bagheera("Vigencia:")

    elif paso == 6:
        estado_empleado["VIGENCIA"] = mensaje
        estado_empleado["paso"] = 7
        return personalidad_bagheera("Nacionalidad:")

    elif paso == 7:
        estado_empleado["NACIONALIDAD"] = mensaje.upper()
        estado_empleado["paso"] = 8
        return personalidad_bagheera("Sexo:")

    elif paso == 8:
        estado_empleado["SEXO"] = mensaje.upper()
        estado_empleado["paso"] = 9
        return personalidad_bagheera("CURP:")

    elif paso == 9:
        estado_empleado["CURP"] = mensaje.upper()
        estado_empleado["paso"] = 10
        return personalidad_bagheera("Domicilio:")

    elif paso == 10:
        estado_empleado["DOMICILIO"] = mensaje.upper()

        agregar_empleado(estado_empleado)

        estado_empleado = {}

        return personalidad_bagheera("✅ Empleado agregado correctamente")


# =========================================================
# 📄 GENERAR CONTRATO
# =========================================================
def generar_contrato(datos):
    base_path = os.path.dirname(__file__)

    plantilla = os.path.join(base_path, "CONTRATO FORMATO TEMPORAL.docx")

    doc = Document(plantilla)

    for p in doc.paragraphs:
        for k, v in datos.items():
            if f"{{{{{k}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{k}}}}}", str(v))

    ruta_docx = os.path.join(base_path, f"{datos['nombre']}.docx")
    doc.save(ruta_docx)

    ruta_pdf = ruta_docx.replace(".docx", ".pdf")

    subprocess.run([
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", base_path,
        ruta_docx
    ])

    return ruta_pdf


#def revisar_vacaciones_por_mes(mensaje):
    from google_sheets import get_sheet
    from datetime import datetime, timedelta

    meses = {
        "enero": 1, "febrero": 2, "marzo": 3,
        "abril": 4, "mayo": 5, "junio": 6,
        "julio": 7, "agosto": 8, "septiembre": 9,
        "octubre": 10, "noviembre": 11, "diciembre": 12
    }

    mes_detectado = None

    for mes, num in meses.items():
        if mes in mensaje:
            mes_detectado = num
            break

    if not mes_detectado:
        return personalidad_bagheera("❌ Mes no válido")

    sheet = get_sheet()
    registros = sheet.get_all_records()

    resultado = []

    # 🔥 función robusta de parseo
    def parse_fecha(valor):
        if not valor:
            return None

        # caso datetime real
        if isinstance(valor, datetime):
            return valor

        texto = str(valor).strip()

        # 🔥 limpiar formato tipo ISO
        if "T" in texto:
            texto = texto.split("T")[0]

        # intentar formatos comunes
        formatos = [
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%m/%d/%Y",
            "%d-%m-%Y"
        ]

        for f in formatos:
            try:
                return datetime.strptime(texto, f)
            except:
                continue

        # 🔥 caso número (Excel serial date)
        try:
            num = float(texto)
            base = datetime(1899, 12, 30)
            return base + timedelta(days=num)
        except:
            pass

        return None

    # 🔍 procesamiento real
    for row in registros:
        fecha_raw = row.get("FECHA_DE_INGRESO", "")
        fecha = parse_fecha(fecha_raw)

        if fecha:
            if fecha.month == mes_detectado:
                resultado.append(row.get("NOMBRE", ""))

    if not resultado:
        return personalidad_bagheera("No hay empleados con aniversario en ese mes")

    return personalidad_bagheera(
        "Empleados con aniversario:\n\n" + "\n".join(resultado)
    )