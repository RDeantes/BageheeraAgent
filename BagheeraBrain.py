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
# 🎯 ROUTER PRINCIPAL
# =========================================================
def procesar(mensaje):
    global estado_contrato, estado_empleado

    mensaje = mensaje.lower().strip()

    # 🔥 PRIORIDAD: EMPLEADO (evita cruce de flujos)
    if estado_empleado.get("activo"):
        return flujo_agregar_empleado(mensaje)

    # 🔥 DESPUÉS CONTRATO
    if estado_contrato.get("activo"):
        return flujo_contrato(mensaje)

    # -----------------------------------------------------
    # COMANDOS
    # -----------------------------------------------------

    if "vencido" in mensaje:
        return revisar_contratos_vencidos()

    if "vacacion" in mensaje:
        return revisar_vacaciones_por_mes(mensaje)

    if "agregar empleado" in mensaje:
        estado_contrato = {}  # 🔥 limpiar contrato
        estado_empleado = {"activo": True, "paso": 1}
        return personalidad_bagheera("Nombre del empleado:")

    if "nuevo contrato" in mensaje:
        estado_empleado = {}  # 🔥 limpiar empleado
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
        estado_contrato = {}  # 🔥 limpiar

        return generar_contrato_desde_excel(datos)


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

        estado_empleado = {}  # 🔥 limpiar flujo

        return personalidad_bagheera("✅ Empleado agregado correctamente")


# =========================================================
# 📄 GENERAR CONTRATO
# =========================================================
def generar_contrato_desde_excel(datos):
    return contrato_desde_excel(datos["nombre"], generar_contrato, personalidad_bagheera)


def generar_contrato(datos):
    base_path = os.path.dirname(__file__)

    archivo = "CONTRATO FORMATO TEMPORAL.docx"
    plantilla = os.path.join(base_path, archivo)

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


# =========================================================
# OTROS
# =========================================================
def revisar_contratos_vencidos():
    return personalidad_bagheera("Función pendiente")


def revisar_vacaciones_por_mes(mensaje):
    return personalidad_bagheera("Función pendiente")