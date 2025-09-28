from playwright.sync_api import sync_playwright
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import time

# ====== Einstellungen ======
URL = "https://coast.hhla.de/report?id=Standard-Report-Segelliste"
SPREADSHEET_ID = "1Q_Dvufm0LCUxYtktMtM18Xz30sXQxCnGfI9SSDFPUNw"
RANGE = "Tabelle1!A1"
SERVICE_ACCOUNT_FILE = "segelliste-b7e6393d9533.json"
# ============================

def extract_table(page):
    page.wait_for_load_state("networkidle", timeout=30000)
    time.sleep(5)
    table = page.query_selector("table")
    if not table:
        raise RuntimeError("Tabelle nicht gefunden")
    rows = table.query_selector_all("tr")
    data = []
    for r in rows:
        cells = r.query_selector_all("th", "td")
        row = [c.inner_text().strip() for c in cells]
        if row:
            data.append(row)
    return data

def sync_sheet(sheet, new_data):
    # Überschreibt alles, damit das Sheet exakt wie die Webseite ist
    sheet.values().clear(spreadsheetId=SPREADSHEET_ID, range=RANGE).execute()
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=RANGE,
        valueInputOption="RAW",
        body={"values": new_data}
    ).execute()
    print(f"Sheet auf neuesten Stand gebracht: {len(new_data)} Zeilen.")

# Google Sheets API verbinden
creds = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
service = build("sheets", "v4", credentials=creds)
sheet = service.spreadsheets()

# Hauptschleife entfällt, da die Windows Aufgabenplanung das Script zu festen Uhrzeiten startet
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    page.goto(URL, timeout=60000)
    data = extract_table(page)
    browser.close()

sync_sheet(sheet, data)
