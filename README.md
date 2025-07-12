# easyEXO - Exchange Online Verwaltungstool

![Screenshot](https://github.com/PS-easyIT/easyEXO/blob/main/%23%20Screenshots/easyEXO-V0.0.13_Dashboard.jpg)

## Übersicht

**easyEXO** ist ein leistungsstarkes PowerShell-basiertes Tool mit einer grafischen Benutzeroberfläche (WPF), das die Verwaltung von Microsoft Exchange Online vereinfacht. Es bündelt eine Vielzahl von administrativen Aufgaben in einer zentralen Konsole und richtet sich an IT-Administratoren, die eine effiziente Alternative zur webbasierten Exchange-Verwaltungskonsole und zur reinen Kommandozeile suchen.

Das Tool bietet einen modularen Aufbau mit verschiedenen Tabs für spezifische Verwaltungsbereiche, von der grundlegenden Postfach- und Kalenderverwaltung bis hin zu komplexen Mailflow-Regeln, Sicherheitsrichtlinien und Fehlerbehebungsdiagnosen.

## Hauptfunktionen

### 📊 Dashboard
- **Live-Statistiken**: Zeigt eine dynamische Übersicht über wichtige Exchange-Objekte wie Postfächer, Gruppen, Kontakte und Ressourcen.
- **Verbindungsstatus**: Klare visuelle Anzeige, ob eine Verbindung zu Exchange Online besteht.

### 🗂️ Grundlegende Verwaltung
- **Kalenderberechtigungen**: Einfaches Anzeigen, Hinzufügen, Ändern und Entfernen von Berechtigungen für Benutzerkalender. Setzen von Standard- und anonymen Berechtigungen.
- **Postfachberechtigungen**: Verwaltung von `FullAccess`, `SendAs` und `SendOnBehalf` Berechtigungen.
- **Freigegebene Postfächer**: Erstellen, Konvertieren und Verwalten von freigegebenen Postfächern und deren Berechtigungen.
- **Gruppen**: Verwaltung von Verteilergruppen, inklusive Mitgliedschaften und Einstellungen.
- **Ressourcen**: Verwaltung von Raum- und Gerätepostfächern (Erstellen, Suchen, Berechtigungen bearbeiten).
- **Kontakte**: Suchen und Bearbeiten von Mail-Kontakten.

### ⚙️ Mailflow
- **Transportregeln**: Erstellen, Anzeigen, Aktivieren/Deaktivieren und Exportieren/Importieren von Mailflow-Regeln.
- **Posteingangsregeln**: Verwaltung von Posteingangsregeln für einzelne Benutzerpostfächer.
- **Nachrichtenverfolgung**: Detaillierte Suche und Analyse von E-Mail-Zustellungen.
- **Automatische Antworten**: Konfiguration von Abwesenheitsnotizen für Benutzer.

### 🛡️ Sicherheit & Compliance
- **Microsoft Defender (ATP)**: Verwaltung von Anti-Phishing-, sicheren Anlagen- und sicheren Links-Richtlinien.
- **Quarantäne**: Anzeigen, Freigeben und Löschen von Nachrichten in Quarantäne.
- **Mobile Device Management (MDM)**: Verwaltung von Geräterichtlinien und Geräten in Quarantäne.

### 🔧 Systemkonfiguration
- **Regionaleinstellungen**: Anpassen von Sprache, Zeitzone sowie Datums- und Zeitformaten für Postfächer.
- **Mail-Routing (Cross-Premises)**: Anzeigen von Mail-Connectors, akzeptierten und Remote-Domänen.

### 📈 Monitoring & Support
- **Health Check**: Umfassende Überprüfung des Exchange Online-Dienststatus, der Konnektivität und wichtiger Konfigurationen.
- **Troubleshooting**: Ausführen von Diagnosen für Postfächer, Abrufen von Drosselungsinformationen (Throttling) und Audit-Logs.

## Voraussetzungen

- **Windows PowerShell 5.1** oder **PowerShell 7**
- **ExchangeOnlineManagement Modul**: Version 3.0.0 oder höher. Das Skript prüft beim Start, ob das Modul installiert ist.
- **Administratorrechte**: Das Skript muss mit erhöhten Rechten ausgeführt werden, um eine Verbindung zu Exchange Online herstellen und Konfigurationen ändern zu können.
- **Internetverbindung**: Für die Verbindung zu Exchange Online.

## Anwendung

1.  **Herunterladen**: Laden Sie das Skript `easyEXO_V0.1.1.ps1` herunter.
2.  **Ausführen**: Starten Sie das Skript in einer PowerShell-Konsole mit Administratorrechten.
    ```powershell
    .\easyEXO_V0.1.1.ps1
    ```
3.  **Verbinden**: Klicken Sie auf den Button "Mit Exchange Online verbinden". Nach erfolgreicher Authentifizierung werden die GUI-Elemente aktiviert.
4.  **Verwalten**: Navigieren Sie durch die Tabs, um die gewünschten Aktionen auszuführen.

## Konfiguration

Das Skript speichert grundlegende Einstellungen in der Windows-Registrierung unter:
`HKCU:\Software\easyIT\easyEXO`

Hier kann z.B. der **Debug-Modus** aktiviert werden, um detailliertere Log-Ausgaben zu erhalten.

## Logging

Alle Aktionen, Fehler und wichtige Informationen werden in einer Log-Datei im Unterordner `Logs` gespeichert (`ExchangeTool.log`). Dies erleichtert die Nachverfolgung und Fehlerbehebung.

