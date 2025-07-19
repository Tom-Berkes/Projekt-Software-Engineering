import psutil
import os
import time
import subprocess
import shutil
import smtplib
import pyautogui
from dotenv import load_dotenv
import logging

# .env-Datei laden
load_dotenv()

# Variablen aus der .env-Datei zuweisen
pps_user = os.getenv("PPS_USER")
pps_password = os.getenv("PPS_PASSWORD")
mail_user = os.getenv("MAIL_USER")
mail_receiver = os.getenv("MAIL_RECEIVER")
mail_password = os.getenv("MAIL_PASSWORD")

# Logging-Konfiguration
logging.basicConfig(
    filename='rpa_log.txt',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s',
    encoding='utf-8'
)

def send_error_mail(subject, body):
    try:
        mailserver = smtplib.SMTP('smtp.office365.com', 587)
        mailserver.ehlo()
        mailserver.starttls()
        mailserver.login(mail_user, mail_password)
        mailserver.sendmail(
            mail_user,
            mail_receiver,
              f'Subject: {subject}\n\n{body}'.encode('utf-8')
        )
        mailserver.quit()
        logging.info("Fehler-E-Mail erfolgreich versendet.")
    except Exception as e:
        logging.error(f"Fehler beim Senden der Fehler-E-Mail: {e}")

def main():
    try:
        logging.info("Skript gestartet.")

        # fpps.exe ggf. beenden
        fpps_check = "fpps.exe" in (p.name() for p in psutil.process_iter())
        if fpps_check:
            os.system('taskkill /f /im fpps.exe')
            logging.info("fpps.exe wurde beendet.")
            time.sleep(2)

        # f-pps starten
        os.startfile(r"C:\Users\Public\Desktop\f-pps.lnk")
        logging.info("f-pps wurde gestartet.")

        time.sleep(2)

        # Prüfen, ob fpps.exe läuft
        fpps_check = "fpps.exe" in (p.name() for p in psutil.process_iter())
        if not fpps_check:
            logging.error("fpps.exe konnte nicht gestartet werden.")
            send_error_mail(
                "RPA Fehler: fpps.exe Start",
                "PPS-Bestellpos.-Export RPA ist fehlgeschlagen: fpps.exe konnte nicht gestartet werden."
            )
            return

        time.sleep(6)

        # Fenster aktivieren
        pyautogui.getWindowsWithTitle("i-soft F-PPS")[0].activate()
        logging.info("Fenster i-soft F-PPS aktiviert.")

        # Login
        try:
            User = pyautogui.locateOnScreen(
                r"C:\Users\tom.berkes\OneDrive - Schwabenhaus GmbH\Desktop\Projekt Software Engineering\pictures\User.PNG",
                grayscale=True, confidence=0.5
            )
            if User is None:
                raise pyautogui.ImageNotFoundException("User-Icon nicht gefunden!")
        except pyautogui.ImageNotFoundException as e:
            logging.error(str(e))
            send_error_mail(
                "RPA Fehler: User-Icon",
                "User-Icon nicht gefunden!"
            )
            return

        pyautogui.moveTo(User)
        pyautogui.click()
        pyautogui.write(pps_user)
        logging.info("Benutzername eingegeben.")

        time.sleep(2)
        pyautogui.press('tab')
        time.sleep(2)
        pyautogui.write(pps_password)
        time.sleep(2)
        pyautogui.press('enter')
        logging.info("Passwort eingegeben und Login abgeschickt.")

        time.sleep(2)

        # Bestellwesen öffnen
        try:
            Bestellwesen = pyautogui.locateOnScreen(
                r"C:\Users\tom.berkes\OneDrive - Schwabenhaus GmbH\Desktop\Projekt Software Engineering\pictures\Bestellwesen.PNG",
                grayscale=True, confidence=0.8
            )
            if Bestellwesen is None:
                raise pyautogui.ImageNotFoundException("Bestellwesen-Icon nicht gefunden!")
        except pyautogui.ImageNotFoundException as e:
            logging.error(str(e))
            send_error_mail(
                "RPA Fehler: Bestellwesen-Icon",
                "Bestellwesen-Icon nicht gefunden!"
            )
            return

        pyautogui.moveTo(Bestellwesen)
        pyautogui.click()
        logging.info("Bestellwesen geöffnet.")

        time.sleep(2)

        try:
            Bestellung = pyautogui.locateOnScreen(
                r"C:\Users\tom.berkes\OneDrive - Schwabenhaus GmbH\Desktop\Projekt Software Engineering\pictures\Bestellung.PNG",
                grayscale=True, confidence=0.9
            )
            if Bestellung is None:
                raise pyautogui.ImageNotFoundException("Bestellung-Icon nicht gefunden!")
        except pyautogui.ImageNotFoundException as e:
            logging.error(str(e))
            send_error_mail(
                "RPA Fehler: Bestellung-Icon",
                "Bestellung-Icon nicht gefunden!"
            )
            return

        pyautogui.moveTo(Bestellung)
        pyautogui.click()
        logging.info("Bestellung geöffnet.")

        time.sleep(2)

        # bsterf.exe prüfen
        bsterf_check = "bsterf.exe" in (p.name() for p in psutil.process_iter())
        if not bsterf_check:
            os.system('taskkill /f /im bsterf.exe')
            os.system('taskkill /f /im fpps.exe')
            logging.error("bsterf.exe läuft nicht, Prozesse beendet.")
            return

        try:
            Listen = pyautogui.locateOnScreen(
                r"C:\Users\tom.berkes\OneDrive - Schwabenhaus GmbH\Desktop\Projekt Software Engineering\pictures\Listen.PNG",
                grayscale=True, confidence=0.8
            )
            if Listen is None:
                raise pyautogui.ImageNotFoundException("Listen-Icon nicht gefunden!")
        except pyautogui.ImageNotFoundException as e:
            logging.error(str(e))
            send_error_mail(
                "RPA Fehler: Listen-Icon",
                "Listen-Icon nicht gefunden!"
            )
            return

        pyautogui.moveTo(Listen)
        pyautogui.click()
        logging.info("Listen geöffnet.")

        time.sleep(2)
        pyautogui.press('down', presses=4)
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(2)

        try:
            Ascii = pyautogui.locateOnScreen(
                r"C:\Users\tom.berkes\OneDrive - Schwabenhaus GmbH\Desktop\Projekt Software Engineering\pictures\Ascii.PNG",
                grayscale=True, confidence=0.8
            )
            if Ascii is None:
                raise pyautogui.ImageNotFoundException("Ascii-Icon nicht gefunden!")
        except pyautogui.ImageNotFoundException as e:
            logging.error(str(e))
            send_error_mail(
                "RPA Fehler: Ascii-Icon",
                "Ascii-Icon nicht gefunden!"
            )
            return

        pyautogui.moveTo(Ascii)
        pyautogui.click()
        logging.info("Ascii gewählt.")

        time.sleep(2)

        try:
            Ausgeben = pyautogui.locateOnScreen(
                r"C:\Users\tom.berkes\OneDrive - Schwabenhaus GmbH\Desktop\Projekt Software Engineering\pictures\Ausgeben.PNG",
                grayscale=True, confidence=0.8
            )
            if Ausgeben is None:
                raise pyautogui.ImageNotFoundException("Ausgeben-Icon nicht gefunden!")
        except pyautogui.ImageNotFoundException as e:
            logging.error(str(e))
            send_error_mail(
                "RPA Fehler: Ausgeben-Icon",
                "Ausgeben-Icon nicht gefunden!"
            )
            return

        pyautogui.moveTo(Ausgeben)
        pyautogui.click()
        logging.info("Ausgeben ausgeführt.")

        # Warten auf Export
        time.sleep(600)  # 10 Minuten

        # Prüfen, ob bsterf.exe noch läuft
        bsterf_check = "bsterf.exe" in (p.name() for p in psutil.process_iter())
        if bsterf_check:
            os.system('taskkill /f /im bsterf.exe')
            os.system('taskkill /f /im fpps.exe')
            logging.info("bsterf.exe und fpps.exe beendet.")

            # Datei kopieren
            quellpfad = r"C:\isoft\f-pps\userdat\BILDSC01.TXT"
            export_pfad = r"C:\Users\tom.berkes\OneDrive - Schwabenhaus GmbH\Desktop\Projekt Software Engineering\BILDSC01.TXT"
            try:
                shutil.copyfile(quellpfad, export_pfad)
                logging.info("Datei exportiert und kopiert.")
            except Exception as e:
                logging.error(f"Fehler beim Kopieren der Datei: {e}")
                send_error_mail(
                    "RPA Fehler: Kopieren der Datei",
                    f"Fehler beim Kopieren der Datei:\n{e}"
                )
                return

            # Statusprüfung der Datei
            if os.path.exists(export_pfad) and os.path.getsize(export_pfad) > 0:
                logging.info("Exportierte Datei gefunden und ist nicht leer.")
            else:
                logging.error("Exportierte Datei fehlt oder ist leer!")
                send_error_mail(
                    "RPA Fehler: Exportierte Datei",
                    "Die exportierte Datei fehlt oder ist leer!"
                )
                return

            # SSIS-Paket starten
            ssis_result = subprocess.run(
                ["python", "Modul_2.py"],
                capture_output=True,
                text=True
            )

            if ssis_result.returncode != 0:
                error_message = (
                    f"Fehler beim Ausführen von Modul_2.py (SSIS):\n"
                    f"stdout:\n{ssis_result.stdout}\n"
                    f"stderr:\n{ssis_result.stderr}"
                )
                logging.error(error_message)
                send_error_mail("RPA Fehler: SSIS-Paket", error_message)
                raise Exception(error_message)  # <-- Exception werfen
            else:
                logging.info("SSIS-Paket erfolgreich abgeschlossen.")

            if ssis_result.returncode == 0:
                # Nach erfolgreichem SSIS-Lauf: Datei löschen
                try:
                    os.remove(export_pfad)
                    logging.info("Datei nach erfolgreichem SSIS-Lauf gelöscht.")
                except Exception as e:
                    logging.error(f"Fehler beim Löschen der Datei nach SSIS-Lauf: {e}")
                    send_error_mail(
                        "RPA Fehler: Löschen der Datei",
                        f"Fehler beim Löschen der Datei nach SSIS-Lauf:\n{e}"
                    )

                # Integrationstest aufrufen!
                integration_result = subprocess.run(
                    ["python", "Modul_3.py"],
                    capture_output=True,
                    text=True
                )

                if integration_result.returncode != 0:
                    error_message = (
                        f"Fehler beim Ausführen von Modul_3.py (Integrationstest):\n"
                        f"stdout:\n{integration_result.stdout}\n"
                        f"stderr:\n{integration_result.stderr}"
                    )
                    logging.error(error_message)
                    send_error_mail("RPA Fehler: Integrationstest", error_message)
                    raise Exception(error_message)  # <-- Exception werfen
                else:
                    logging.info("Integrationstest erfolgreich abgeschlossen.")
            else:
                logging.error(f"SSIS-Paket Fehler: stdout: {integration_result.stdout} stderr: {integration_result.stderr}")
                send_error_mail(
                    "RPA Fehler: SSIS-Paket",
                    f"SSIS-Paket Fehler:\nstdout:\n{integration_result.stdout}\nstderr:\n{integration_result.stderr}"
                )

        else:
            logging.error("bsterf.exe läuft nicht mehr nach Export.")
            send_error_mail(
                "RPA Fehler: bsterf.exe",
                "bsterf.exe läuft nicht mehr nach Export."
            )
            return

        logging.info("Skript erfolgreich abgeschlossen.")

    except Exception as e:
        logging.error(f"Unbehandelter Fehler: {e}", exc_info=True)
        send_error_mail(
            "RPA Fehler: Unbehandelter Fehler",
            f"Es ist ein unbehandelter Fehler aufgetreten:\n{e}"
        )
        return

if __name__ == "__main__":
    main()
