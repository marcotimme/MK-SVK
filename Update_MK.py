import imaplib
import email
from email.header import decode_header
import re
import os
from datetime import datetime
from google.oauth2.service_account import Credentials
import gspread

# Service-Account-JSON als GitHub Secret (aus Umgebungsvariable im Workflow)
GSHEET_CREDS_JSON = os.environ.get('GSHEET_CREDENTIALS_JSON')  # Der Wert ist die JSON als String

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Schreibe die JSON temporär ins Dateisystem, falls nötig
with open("service_account.json", "w") as f:
    f.write(GSHEET_CREDS_JSON)

creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
gc = gspread.authorize(creds)
SHEET_NAME = "Mannschaftskasse_2025_2026"
worksheet = gc.open(SHEET_NAME).sheet1

# Transaktionscodes aus Sheet auslesen
existing_codes = set(worksheet.col_values(6)[4:])  # 6. Spalte, ab Zeile 5
existing_codes.add("14G77435X57269737")
print(existing_codes)

# Gmail-Login (IMAP - Vorsicht: Passwort niemals hart coden)
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(os.environ.get("GMAIL_USER"), os.environ.get("GMAIL_PASSWORD"))  # als Secret setzen!
mail.select("Paypal")
print("Login erfolgreich")

status, messages = mail.search(None, "ALL")
email_ids = messages[0].split()
print("Emails ausgelesen.")

# Regex-Muster...
transaktionscode_pattern = r"\b[A-Z0-9]{17}\b"
betrag_pattern = r"\b\d{1,3}(?:\.\d{3})*,\d{2}\s*€\s*EUR\b"
name_pattern = r"(?:Mitteilung von|DEINE MITTEILUNG AN|Deine Mitteilung an )([^\d<:\n]+)"
zahlungsmuster_pattern = r"Von dir bezahlt"
freitext_pattern = r"(?i)(?<![a-zA-Z])(MK|Mannschaftskasse|@MK|Strafe)(?:[-\s]+(.+))?"
datum_pattern = r"\b\d{1,2}\.\s*[A-Za-z]+\s*\d{4}\b"

for email_id in email_ids:
    status, msg_data = mail.fetch(email_id, "(RFC822)")
    print(msg_data)
    msg = email.message_from_bytes(msg_data[0])
    subject = decode_header(msg["Subject"])
    if isinstance(subject, bytes):
        subject = subject.decode()
    if msg.is_multipart():
        body = ""
        for part in msg.walk():
            if part.get_content_type() == "text/plain" or part.get_content_type() == "text/html":
                body += part.get_payload(decode=True).decode()
    else:
        body = msg.get_payload(decode=True).decode()
    datum_match = re.search(datum_pattern, body)
    if datum_match:
        datum_str = datum_match.group()
        monate = {
            "Januar": "01", "Februar": "02", "März": "03", "April": "04",
            "Mai": "05", "Juni": "06", "Juli": "07", "August": "08",
            "September": "09", "Oktober": "10", "November": "11", "Dezember": "12"}
        for monat, zahl in monate.items():
            if monat in datum_str:
                datum_str = datum_str.replace(monat, zahl)
                break
        try:
            datum = datetime.strptime(datum_str, "%d. %m %Y").strftime("%d.%m.%Y")
        except ValueError:
            datum = "Ungültiges Datum"
    else:
        datum = "Datum nicht gefunden"
    freitext_match = re.search(freitext_pattern, body)
    if not freitext_match:
        print(f"E-Mail ignoriert: kein MK/Strafe.")
        continue
    elif freitext_match.group(2):
        nachricht = freitext_match.group(2).strip()
    else:
        nachricht = freitext_match.group(1).strip()
    nachricht = nachricht.split('</')
    kategorie = "Beitrag"
    if nachricht.lower().startswith("strafe"):
        kategorie = "Strafe"
    transaktionscode_match = re.search(transaktionscode_pattern, body)
    transaktionscode = transaktionscode_match.group() if transaktionscode_match else "Nicht gefunden"
    if transaktionscode in existing_codes:
        print(f"Transaktionscode {transaktionscode} bereits vorhanden.")
        continue
    name_match = re.search(name_pattern, body)
    name = name_match.group(1).strip() if name_match else "Nicht gefunden"
    betrag_match = re.search(betrag_pattern, body)
    if betrag_match:
        betrag_str = betrag_match.group().replace('.', '').replace(',', '.').replace('€', '').replace('EUR', '').strip()
        betrag = float(betrag_str)
        if re.search(zahlungsmuster_pattern, body):
            betrag = -betrag
            kategorie = "Ausgabe"
            nachricht = f"An {name} - {nachricht}"
    else:
        betrag = "Nicht gefunden"
    # Neue Zeile ins Google Sheet
    worksheet.append_row([datum, name, kategorie, betrag, nachricht, transaktionscode])
    print(f"Row added: Datum: {datum}, Name: {name}, Kategorie: {kategorie}, Betrag: {betrag}, Nachricht: {nachricht}, ID: {transaktionscode}")
    existing_codes.add(transaktionscode)
mail.logout()

