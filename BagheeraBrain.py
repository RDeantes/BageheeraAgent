import os
import subprocess
from datetime import datetime
from docx import Document
from docx2pdf import convert
from BagheeraExcel import contrato_desde_excel
from BagheeraExcel import buscar_empleado, actualizar_vigencia

# 🧠 MEMORIA GLOBAL
estado_contrato = {}

# =========================================================
# 🐱 PERSONALIDAD
# =========================================================
def personalidad_bagheera(respuesta):
    return f"🐱:\n{respuesta}\n\n“La fuerza no está en rugir… está en saber cuándo moverte en silencio.” 🌑"


# =========================================================
# 🎯 PROCESAR MENSAJES (ROUTER)
# =========================================================
def procesar(mensaje):
    global estado_contrato

    mensaje = mensaje.lower().strip()

    # 🔥 1. SI YA ESTAMOS EN FLUJO → CONTINUAR
    if estado_contrato.get("activo"):
        return flujo_contrato(mensaje)
    # =========================================================
    # 1️⃣ VENCIDOS
    # =========================================================
    if "vencido" in mensaje:
        return revisar_contratos_vencidos()

    # =========================================================
    # 2️⃣ VACACIONES
    # =========================================================
    if "vacacion" in mensaje:

        meses = {
            "enero": "enero",
            "febrero": "febrero",
            "marzo": "marzo",
            "abril": "abril",
            "mayo": "mayo",
            "junio": "junio",
            "julio": "julio",
            "agosto": "agosto",
            "septiembre": "septiembre",
            "octubre": "octubre",
            "noviembre": "noviembre",
            "diciembre": "diciembre"
        }

        for mes_texto, mes_param in meses.items():
            if mes_texto in mensaje:
                return revisar_vacaciones_por_mes(mes_param)

        return personalidad_bagheera("USA SIEMPRE EL FORMATO VACACIONES + EL MES A BUSCAR")

    # 
# =========================================================
# 3️⃣ NUEVO CONTRATO (FLUJO MANUAL)
# =========================================================
   # 🔥 4. NUEVO CONTRATO
    if "nuevo contrato" in mensaje:
        estado_contrato = {"activo": True}
        return personalidad_bagheera("¿Qué tipo de contrato?\n(temporal / permanente)")

    return personalidad_bagheera("No entendí la orden.")
    
    
    # =========================================================
    # 4️⃣ CONTINUAR FLUJO
    # =========================================================
    if estado_contrato.get("activo"):
        return flujo_contrato(mensaje)

    return personalidad_bagheera("No entendí la orden.")


# =========================================================
# 🧠 FLUJO DE CONTRATO
# =========================================================
def flujo_contrato(mensaje):
    global estado_contrato

    mensaje = mensaje.lower().strip()

    # 🔥 TIPO DE CONTRATO
    if "tipo" not in estado_contrato:

        if "temporal" in mensaje:
            estado_contrato["tipo"] = "TEMPORAL"
        elif "permanente" in mensaje:
            estado_contrato["tipo"] = "PERMANENTE"
        else:
            return personalidad_bagheera("Responde: temporal o permanente")

        return personalidad_bagheera("¿La jornada es parcial o completa?")
    elif "jornada" not in estado_contrato:
        estado_contrato["jornada"] = mensaje.upper()
        return personalidad_bagheera("Duración del contrato:")

    elif "duracion" not in estado_contrato:
        estado_contrato["duracion"] = mensaje.upper()
        return personalidad_bagheera("Fecha de inicio (YYYY-MM-DD):")

    elif "fecha_inicio" not in estado_contrato:
        estado_contrato["fecha_inicio"] = mensaje.upper()
        return personalidad_bagheera("Fecha de término (YYYY-MM-DD):")

    elif "fecha_termino" not in estado_contrato:
        estado_contrato["fecha_termino"] = mensaje.upper()
        return personalidad_bagheera("Nombre del empleado:")

    elif "nombre" not in estado_contrato:
        estado_contrato["nombre"] = mensaje.upper()

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

    if tipo == "temporal" and "completa" in jornada:
        archivo = "CONTRATO FORMATO TEMPORAL.docx"

    elif tipo == "temporal" and "parcial" in jornada:
        archivo = "CONTRATO FORMATO TEMPORAL PARCIAL.docx"

    elif tipo == "permanente" and "completa" in jornada:
        archivo = "CONTRATO FORMATO TIEMPO INDETERMINADO.docx"

    elif tipo == "permanente" and "parcial" in jornada:
        archivo = "CONTRATO FORMATO INDETERMINADO PARCIAL.docx"

    else:
        return personalidad_bagheera("❌ No entendí el tipo o la jornada")

    plantilla = os.path.join(base_path, archivo)

    if not os.path.exists(plantilla):
        return personalidad_bagheera(f"❌ No encontré la plantilla: {archivo}")

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
        "{{dias}}": datos.get("dias", "")
    }

    for p in doc.paragraphs:
        for key, value in reemplazos.items():
            if key in p.text:
                p.text = p.text.replace(key, str(value))

    nombre_archivo = f"Contrato_{datos['nombre'].replace(' ', '_')}.docx"
    ruta_salida = os.path.join(base_path, nombre_archivo)
    doc.save(ruta_salida)

    # 🔥 PDF
    ruta_pdf = ruta_salida.replace(".docx", ".pdf")
    convert(ruta_salida, ruta_pdf)

    return ruta_pdf


# =========================================================
# 📄 CONTRATOS VENCIDOS
# =========================================================
def revisar_contratos_vencidos():
    resultado = subprocess.run(['python', 'LeeArchivo.py'], capture_output=True, text=True)

    if resultado.returncode == 0:
        return personalidad_bagheera(resultado.stdout)
    else:
        return personalidad_bagheera(f"Error: {resultado.stderr}")


# =========================================================
# 🌴 VACACIONES
# =========================================================
def revisar_vacaciones_por_mes(mes):
    resultado = subprocess.run(['python', 'RevisarVacaciones.py', mes], capture_output=True, text=True)

    if resultado.returncode == 0:
        return personalidad_bagheera(resultado.stdout)
    else:
        return personalidad_bagheera(f"Error: {resultado.stderr}")
    
    
def generar_contrato_desde_excel(datos):

    persona = buscar_empleado(datos["nombre"])

    if persona is None:
        return personalidad_bagheera("❌ No encontré al empleado en el Excel")

    datos_finales = {
        "tipo": datos["tipo"],
        "jornada": datos["jornada"],
        "duracion": datos["duracion"],
        "fecha_inicio": datos["fecha_inicio"],
        "fecha_termino": datos["fecha_termino"],
        "nombre": persona['NOMBRE'],
        "nacionalidad": persona.get('NACIONALIDAD', 'MEXICANA'),
        "sexo": persona.get('SEXO', 'M'),
        "curp": persona.get('CURP', ''),
        "domicilio": persona.get('DOMICILIO', ''),
        "puesto": persona.get('PUESTO', 'EMPLEADO'),
        "dias": "LUNES A SABADO"
    }

    pdf = generar_contrato(datos_finales)

    # 🔥 ACTUALIZA EXCEL
    actualizar_vigencia(datos["nombre"], datos["fecha_termino"])

    return pdf