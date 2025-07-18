import os
import sys
import pyodbc
import logging
from dotenv import load_dotenv

# Logging-Konfiguration
logging.basicConfig(
    filename='rpa_log.txt',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s',
    encoding='utf-8'
)

# .env laden
load_dotenv()
cad_password = os.getenv("CAD_PASSWORD")

# Verbindung zur Datenbank
server = "10.15.14.20"
database = "CAD"
username = "CAD"
password = cad_password
table = "CAD.dbo.Bestellungen_Test"

conn_str = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password};"
)

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # 1. Anzahl Datensätze prüfen (mindestens 100)
    cursor.execute(f"SELECT COUNT(*) FROM {table}")
    anzahl = cursor.fetchone()[0]
    if anzahl < 100:
        logging.error(f"Integrationstest FEHLER: Zu wenige Datensätze in der Tabelle {table}: {anzahl} gefunden, mindestens 100 erwartet.")
        sys.exit(1)
    else:
        logging.info(f"Integrationstest: Anzahl Datensätze in {table} geprüft: {anzahl} vorhanden.")

    # 2. Testdatensatz suchen (über Bestellnummer)
    test_bestellnummer = "2128186 "
    cursor.execute(f"SELECT * FROM {table} WHERE [Bestellnummer] = ?", test_bestellnummer)
    row = cursor.fetchone()
    if not row:
        logging.error(f"Integrationstest FEHLER: Testdatensatz mit Bestellnummer '{test_bestellnummer}' nicht gefunden!")
        sys.exit(1)
    else:
        logging.info(f"Integrationstest: Testdatensatz mit Bestellnummer '{test_bestellnummer}' gefunden.")

    # 3. Einzelne Spalten prüfen
    sollwerte = {
        'Bestellnummer': "2128186 ",
        'Bestelldatum': "31.07.2019 ",
        'Status': "berechnet      ",
        'Liefertermin': "00.00.0000 ",
        'Bauvorhabennummer': "190002       ",
        'Bauvorhabenname': "MEURER,88279, ",
        'Kuerzel': "A  ",
        'Kostenstellennummer': "991100 ",
        'Kostenstellenname': "Ausgliederungsk   ",
        'Lieferantennummer': "    748 ",
        'Lieferantenname': "BauGrund Süd    ",
        'Titel': "ERDSONDENBOHRUN"
    }
    columns = [column[0] for column in cursor.description]
    for spalte, soll in sollwerte.items():
        ist = row[columns.index(spalte)]
        if ist != soll:
            logging.error(f"Integrationstest FEHLER: Abweichung in Spalte '{spalte}': erwartet '{soll}', gefunden '{ist}'")
            sys.exit(1)
        else:
            logging.info(f"Integrationstest: Spalte '{spalte}' korrekt: '{ist}'")

    logging.info("Integrationstest ERFOLGREICH: Alle Prüfungen bestanden und Testdatensatz vollständig und korrekt in der Datenbank vorhanden.")
    sys.exit(0)

except Exception as e:
    logging.error(f"Integrationstest FEHLER: Ausnahme aufgetreten: {e}")
    sys.exit(1)
