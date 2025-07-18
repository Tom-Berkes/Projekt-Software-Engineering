import subprocess
import sys
import logging

# Logging-Konfiguration (wie in den anderen Modulen)
logging.basicConfig(
    filename='rpa_log.txt',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s',
    encoding='utf-8'
)

def run_ssis_package():
    ssis_package = r"C:\Users\tom.berkes\OneDrive - Schwabenhaus GmbH\Desktop\Projekt Software Engineering\Projekt Software Engineering\Package.dtsx"
    dtexec_path = r"C:\Program Files\Microsoft SQL Server\130\DTS\Binn\dtexec.exe"
    command = [dtexec_path, "/f", ssis_package]

    try:
        result = subprocess.run(command, capture_output=True, text=True, encoding='utf-8', errors='replace')
        logging.info(f"SSIS-Paket ausgeführt. Rückgabe: {result.returncode}")
        if result.returncode == 0:
            logging.info("SSIS-Paket erfolgreich abgeschlossen.")
            sys.exit(0)
        else:
            logging.error(f"SSIS-Paket Fehler: stdout: {result.stdout} stderr: {result.stderr}")
            sys.exit(1)
    except Exception as e:
        logging.error(f"Fehler beim Ausführen des SSIS-Pakets: {e}")
        sys.exit(1)

if __name__ == "__main__":
    run_ssis_package()
