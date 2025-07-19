## Softwarearchitektur (C4-Modell)

Im Folgenden sind die wichtigsten Ebenen der Softwarearchitektur als C4-Diagramme dargestellt.
### C4-Kontextdiagramm (Mermaid)
```mermaid
graph TD
    %% Akteure
    User([Systemverantwortlicher<br/>/ Entwickler])
    
    %% Systeme
    APS([APS<br/>(Auftragsplanungssystem)])
    Automation([Automatisierte Datenextraktion<br/>& Integration])
    SQL([SQL-Datenbank<br/>(CAD)])
    Mail([E-Mail-Server])
    VM([Virtuelle Maschine<br/>(Windows-Server)])
    
    %% Beziehungen
    User -- "Startet, überwacht und konfiguriert" --> Automation
    Automation -- "Automatisiert GUI-Interaktion<br/>(pyautogui)" --> APS
    APS -- "Exportiert CSV-Report<br/>(BILDSC01.TXT)" --> Automation
    Automation -- "Überträgt Daten<br/>(SSIS/ETL)" --> SQL
    Automation -- "Sendet Fehler-/Statusmails" --> Mail
    Automation -- "Läuft auf" --> VM
    VM -- "Hostet" --> APS
    VM -- "Hostet" --> Automation
    User -- "Erhält Benachrichtigungen" --> Mail
    SQL -- "Stellt Daten bereit für<br/>BI, Produktion, Auswertung" --> User
