import os
import subprocess
from datetime import datetime
from docx import Document

from BagheeraExcel import (
    contrato_desde_excel,
    buscar_empleado,
    actualizar_vigencia,
    obtener_proximo_id,
    agregar_empleado
)

# 🧠 MEMORIA GLOBAL
estado_contrato = {}
estado_empleado = {}

# =========================================================
# 🐱 PERSONALIDAD
# =========================================================
def personalidad_bagheera(respuesta):
    return f"🐱:\n{respuesta}\n\n“La fuerza no está en rugir… está en saber cuándo moverte en silencio.” 🌑"


# =========================================================
# 🎯 ROUTER
# =========================================================
def procesar(mensaje):
    global estado_contrato, estado_empleado

    mensaje = mensaje.strip()

    # 🔥 PRIORIDAD: flujos activos
    if estado_contrato.get("activo"):
        return flujo_contrato(mensaje)

    if estado_empleado.get("activo"):
        return flujo_agregar_empleado(mensaje)

    # VENCIDOS
    if "VENCIDO" in mensaje:
        return revisar_contratos_vencidos()

    # VACACIONES
    if "VACACION" in mensaje:
        meses = {
            "ENERO": "enero","FEBRERO": "febrero","MARZO": "marzo",
            "ABRIL": "abril","MAYO": "mayo","JUNIO": "junio",
            "JULIO": "julio","AGOSTO": "agosto","SEPTIEMBRE": "septiembre",
            "OCTUBRE": "octubre","NOVIEMBRE": "noviembre","DICIEMBRE": "diciembre"
        }

        for mes_texto, mes_param in meses.items():
            if mes_texto in mensaje:
                return revisar_vacaciones_por_mes(mes_param)

        return personalidad_bagheera("Usa: VACACIONES + MES")

    # NUEVO CONTRATO
    if "NUEVO CONTRATO" in mensaje:
        estado_contrato = {"activo": True}
        return personalidad_bagheera("¿Qué tipo de contrato?\n(TEMPORAL / PERMANENTE)")

    # AGREGAR EMPLEADO
    if "AGREGAR EMPLEADO" in mensaje:
        proximo_id = obtener_proximo_id()
        estado_empleado = {"activo": True, "id": proximo_id, "paso": 1}
        return personalidad_bagheera(f"ID: {proximo_id}\n¿Nombre del empleado?")

    return personalidad_bagheera("NO ENTENDÍ LA ORDEN.")


# =========================================================
# 🧠 FLUJO CONTRATO
# =========================================================
def flujo_contrato(mensaje):
    global estado_contrato

    # TIPO
    if "tipo" not in estado_contrato:
        if "TEMPORAL" in mensaje:
            estado_contrato["tipo"] = "TEMPORAL"
        elif "PERMANENTE" in mensaje:
            estado_contrato["tipo"] = "PERMANENTE"
        else:
            return personalidad_bagheera("Responde: TEMPORAL o PERMANENTE")

        return personalidad_bagheera("¿Jornada COMPLETA o PARCIAL?")

    # JORNADA
    elif "jornada" not in estado_contrato:
        estado_contrato["jornada"] = mensaje
        return personalidad_bagheera("Duración del contrato (ej: 3 MESES):")

    # DURACION
    elif "duracion" not in estado_contrato:
        estado_contrato["duracion"] = mensaje
        return personalidad_bagheera("Fecha de inicio (YYYY-MM-DD):")

    # FECHA INICIO
    elif "fecha_inicio" not in estado_contrato:
        estado_contrato["fecha_inicio"] = mensaje
        return personalidad_bagheera("Fecha de término (YYYY-MM-DD):")

    # FECHA TERMINO
    elif "fecha_termino" not in estado_contrato:
        estado_contrato["fecha_termino"] = mensaje
        return personalidad_bagheera("Nombre del empleado:")

    # NOMBRE
    elif "nombre" not in estado_contrato:
        estado_contrato["nombre"] = mensaje

        datos = estado_contrato.copy()
        estado_contrato = {}

        return generar_contrato_desde_excel(datos)


# =========================================================
# 📄 GENERAR CONTRATO
# =========================================================
def generar_contrato(datos):

    base_path = os.path.dirname(__file__)

    tipo = datos["tipo"].lower()
    jornada = datos["jornada"].lower()

    if tipo == "temporal" and "COMPLETA" in jornada:
        archivo = "CONTRATO FORMATO TEMPORAL.docx"

    elif tipo == "temporal" and "PARCIAL" in jornada:
        archivo = "CONTRATO FORMATO TEMPORAL Y PARCIAL.docx"

    elif tipo == "permanente" and "COMPLETA" in jornada:
        archivo = "CONTRATO FORMATO TIEMPO INDETERMINADO.docx"

    elif tipo == "permanente" and "PARCIAL" in jornada:
        archivo = "CONTRATO FORMATO INDETERMINADO Y PARCIAL.docx"

    else:
        return personalidad_bagheera("❌ Error en tipo o jornada")

    plantilla = os.path.join(base_path, archivo)

    if not os.path.exists(plantilla):
        return personalidad_bagheera(f"❌ No encontré plantilla: {archivo}")

    doc = Document(plantilla)

    reemplazos = {
        "{{nombre}}": datos.get("nombre", ""),
        "{{duracion}}": datos.get("duracion", ""),
        "{{fecha_inicio}}": datos.get("fecha_inicio", ""),
        "{{fecha_termino}}": datos.get("fecha_termino", ""),
        "{{nacionalidad}}": datos.get("nacionalidad", ""),
        "{{sexo}}": datos.get("sexo", ""),
        "{{curp}}": datos.get("curp", ""),
        "{{domicilio}}": datos.get("domicilio", ""),
        "{{puesto}}": datos.get("puesto", ""),
        "{{dias}}": datos.get("dias", "LUNES A SABADO")
    }

    for p in doc.paragraphs:
        for key, val in reemplazos.items():
            if key in p.text:
                p.text = p.text.replace(key, str(val))

    # 🔥 GUARDAR DOCX
    nombre_archivo = f"Contrato_{datos['nombre'].replace(' ', '_')}.docx"
    ruta_docx = os.path.join(base_path, nombre_archivo)
    doc.save(ruta_docx)

    # 🔥 CONVERTIR A PDF (LIBREOFFICE)
    ruta_pdf = ruta_docx.replace(".docx", ".pdf")

    subprocess.run([
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", base_path,
        ruta_docx
    ])

    return ruta_pdf


# =========================================================
# 📄 DESDE EXCEL
# =========================================================
def generar_contrato_desde_excel(datos):

    persona = buscar_empleado(datos["nombre"])

    if persona is None:
        return personalidad_bagheera("❌ No encontré al empleado")

    datos_finales = {
        **datos,
        "nombre": persona["NOMBRE"],
        "nacionalidad": persona.get("NACIONALIDAD", ""),
        "sexo": persona.get("SEXO", ""),
        "curp": persona.get("CURP", ""),
        "domicilio": persona.get("DOMICILIO", ""),
        "puesto": persona.get("PUESTO", "")
    }

    pdf = generar_contrato(datos_finales)

    # 🔥 ACTUALIZA BIEN EL EXCEL
    actualizar_vigencia(persona["NOMBRE"], datos["fecha_termino"])

    return pdf


# =========================================================
# 📊 VENCIDOS
# =========================================================
def revisar_contratos_vencidos():
    res = subprocess.run(['python', 'LeeArchivo.py'], capture_output=True, text=True)
    return personalidad_bagheera(res.stdout)


# =========================================================
# 🌴 VACACIONES
# =========================================================
def revisar_vacaciones_por_mes(mes):
    res = subprocess.run(['python', 'RevisarVacaciones.py', mes], capture_output=True, text=True)
    return personalidad_bagheera(res.stdout)


# =========================================================
# 👤 AGREGAR EMPLEADO
# =========================================================
def flujo_agregar_empleado(mensaje):
    global estado_empleado

    if estado_empleado["paso"] == 1:
        estado_empleado["nombre"] = mensaje
        estado_empleado["paso"] = 2
        return personalidad_bagheera("Área:")

    elif estado_empleado["paso"] == 2:
        estado_empleado["area"] = mensaje
        estado_empleado["paso"] = 3
        return personalidad_bagheera("Puesto:")

    elif estado_empleado["paso"] == 3:
        estado_empleado["puesto"] = mensaje
        estado_empleado["paso"] = 4
        return personalidad_bagheera("Fecha ingreso (YYYY-MM-DD):")

    elif estado_empleado["paso"] == 4:
        estado_empleado["fecha_ingreso"] = mensaje
        estado_empleado["paso"] = 5
        return personalidad_bagheera("Tipo contrato:")

    elif estado_empleado["paso"] == 5:
        estado_empleado["tipo_contrato"] = mensaje
        estado_empleado["paso"] = 6
        return personalidad_bagheera("Vigencia:")

    elif estado_empleado["paso"] == 6:
        estado_empleado["vigencia"] = mensaje
        estado_empleado["paso"] = 7
        return personalidad_bagheera("Nacionalidad:")

    elif estado_empleado["paso"] == 7:
        estado_empleado["nacionalidad"] = mensaje
        estado_empleado["paso"] = 8
        return personalidad_bagheera("Sexo:")

    elif estado_empleado["paso"] == 8:
        estado_empleado["sexo"] = mensaje
        estado_empleado["paso"] = 9
        return personalidad_bagheera("CURP:")

    elif estado_empleado["paso"] == 9:
        estado_empleado["curp"] = mensaje
        estado_empleado["paso"] = 10
        return personalidad_bagheera("Domicilio:")

    elif estado_empleado["paso"] == 10:
        estado_empleado["domicilio"] = mensaje

        datos = {
            'NOMBRE': estado_empleado['nombre'],
            'AREA': estado_empleado['area'],
            'PUESTO': estado_empleado['puesto'],
            'FECHA_DE_INGRESO': estado_empleado['fecha_ingreso'],
            'TIPO DE CONTRATO': estado_empleado['tipo_contrato'],
            'VIGENCIA': estado_empleado['vigencia'],
            'NACIONALIDAD': estado_empleado['nacionalidad'],
            'SEXO': estado_empleado['sexo'],
            'CURP': estado_empleado['curp'],
            'DOMICILIO': estado_empleado['domicilio']
        }

        try:
            id_agregado = agregar_empleado(datos)
            estado_empleado = {}
            return personalidad_bagheera(f"✅ Empleado agregado correctamente con ID: {id_agregado}")
        except Exception as e:
            estado_empleado = {}
            return personalidad_bagheera(f"❌ Error al agregar empleado: {str(e)}")