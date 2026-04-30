import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

def get_sheet():
    creds_json = os.environ.get("GOOGLE_CREDS")

    if not creds_json:
        raise Exception("❌ No existe GOOGLE_CREDS en Railway")

    creds_dict = json.loads(creds_json)

    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
    client = gspread.authorize(creds)

    return client.open("CONCENTRADO_EMPLEADOS").sheet1