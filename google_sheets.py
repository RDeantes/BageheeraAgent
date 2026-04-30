import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 🔐 permisos
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

# 🔑 credenciales
creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
client = gspread.authorize(creds)

# 📊 abrir sheet
sheet = client.open("CONCENTRADO_EMPLEADOS").sheet1

