import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 🔐 Permisos
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

def _get_client():
    # 🔥 Lee credenciales desde variable de entorno
    raw = os.environ.get("GOOGLE_CREDS")
    if not raw:
        raise RuntimeError("❌ Falta la variable de entorno GOOGLE_CREDS en Railway")

    try:
        creds_dict = json.loads(raw)
    except Exception as e:
        raise RuntimeError(f"❌ GOOGLE_CREDS no es JSON válido: {e}")

    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
    return gspread.authorize(creds)

def get_sheet():
    client = _get_client()

    # ⚠️ El nombre debe ser EXACTO
    spreadsheet = client.open("CONCENTRADO_EMPLEADOS")
    return spreadsheet.sheet1