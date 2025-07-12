Dieses Repository enthält das PowerShell Script **easyEXO_V0.0.12.ps1** zur Verwaltung und Konfiguration von Exchange Online. 
easyEXO bietet eine übersichtliche Benutzeroberfläche mit vielfältigen Tabs und Funktionen.

## 📚 Inhalt / Table of Contents

- [🇩🇪 Deutsch](#-deutsch)
  - [🔧 Übersicht](#-übersicht)
  - [⚙️ Voraussetzungen](#-voraussetzungen)
  - [🚀 Installation](#-installation)
  - [🖥️ Starten](#-starten)
  - [📋 Tabs & Funktionen](#-tabs--funktionen)
    - [Dashboard](#dashboard)
    - [Grundlegende Verwaltung](#grundlegende-verwaltung)
    - [Mail Flow](#mail-flow)
    - [Systemkonfiguration](#systemkonfiguration)
    - [Monitoring & Support](#monitoring--support)
  - [📂 Logs](#-logs)
  - [🔗 Weiterführende Links](#-weiterführende-links)
- [🇬🇧 English](#-english)
  - [🔧 Overview](#-overview)
  - [⚙️ Prerequisites](#-prerequisites)
  - [🚀 Installation](#-installation-1)
  - [🖥️ Launching](#-launching)
  - [📋 Tabs & Features](#-tabs--features)
    - [Dashboard](#dashboard-1)
    - [Basic Management](#basic-management)
    - [Mail Flow](#mail-flow-1)
    - [System Configuration](#system-configuration)
    - [Monitoring & Support](#monitoring--support-1)
  - [📂 Logs](#-logs-1)
  - [🔗 References](#-references)

## 🇩🇪 Deutsch

### 🔧 Übersicht
easyEXO ist ein PowerShell-Skript mit WPF-GUI, das zentrale Exchange Online-Verwaltungsaufgaben in einer grafischen Oberfläche bündelt.

### ⚙️ Voraussetzungen
- Windows mit PowerShell 7 oder höher
- PowerShell-Modul **ExchangeOnlineManagement** (`Install-Module ExchangeOnlineManagement`)
- Exchange Online-Administratorrechte (z.B. Global Admin)
- Ausführungsrichtlinie `RemoteSigned` oder strenger

### 🖥️ Starten
- Doppelklick auf `easyEXO_V0.0.12.ps1`  
- Oder im PowerShell (Administrator):
  ```powershell
  pwsh .\easyEXO_V0.0.12.ps1
  ```

### 📋 Tabs & Funktionen

#### Dashboard
Zeigt Skriptversion, Verbindungsstatus und Schnellstatistiken.

#### Grundlegende Verwaltung
- **Calendar**: Kalenderberechtigungen und Freigaben verwalten  
- **Mailbox**: Postfach-Eigenschaften und Delegierungen  
- **Shared Mailbox**: Freigegebene Postfächer konfigurieren  
- **Groups**: Office 365-Gruppen und Sicherheitsgruppen verwalten  
- **Resources**: Ressourcenpostfächer (Raum/Equipment) verwalten  
- **Contacts**: Kontakte außerhalb der Organisation verwalten  

#### Mail Flow
- **Mail Flow Rules**: Transportregeln anzeigen und bearbeiten  
- **Inbox Rules**: Posteingangsregeln für Benutzerpostfächer  
- **Message Trace**: Nachrichtennachverfolgung und Details (inkl. Export)  
- **Auto Reply**: Automatische Antworten (Abwesenheitsnotizen) erstellen und verwalten  

#### Systemkonfiguration
- **Regionsettings**: Regional und Zeitzonen Einstellungen auslesen und anpassen 
- **EXO Settings**: Global Organization Settings auslesen und anpassen  

#### Monitoring & Support
- **Health Check**: Systemgesundheit und Service-Status prüfen  
- **Mailbox Audit**: Prüfprotokolle und Auditing-Einstellungen  
- **Reports**: Standardberichte generieren (z.B. Mailbox-Größen)  
- **Troubleshooting**: Hilfsfunktionen und Log-Analyse  

### 📂 Logs
- Ordner: `Logs`  

### 🔗 Weiterführende Links
- [easyEXO auf GitHub](https://github.com/PS-easyIT/easyEXO)  
- [Exchange Online PowerShell Docs](https://aka.ms/exops-docs)

---

## 🇬🇧 English

### 🔧 Overview
easyEXO is a PowerShell WPF GUI tool grouping key Exchange Online management tasks into a single interface.

### ⚙️ Prerequisites
- Windows with PowerShell 7 or newer  
- **ExchangeOnlineManagement** PowerShell module (`Install-Module ExchangeOnlineManagement`)  
- Exchange Online admin permissions  
- Execution policy `RemoteSigned` or stricter  

### 🖥️ Launching
- Double-click `easyEXO_V0.0.12.ps1`  
- Or run in PowerShell (admin):
  ```powershell
  pwsh .\easyEXO_V0.0.12.ps1
  ```

### 📋 Tabs & Features

#### Dashboard
Displays script version, connection status, and quick stats.

#### Basic Management
- **Calendar**: Manage calendar permissions and sharing  
- **Mailbox**: Manage mailbox properties and delegations  
- **Shared Mailbox**: Configure shared mailboxes  
- **Groups**: Manage Office 365 and security groups  
- **Resources**: Manage resource mailboxes (room/equipment)  
- **Contacts**: Manage external contacts  

#### Mail Flow
- **Mail Flow Rules**: View/edit transport rules  
- **Inbox Rules**: User mailbox inbox rules  
- **Message Trace**: Track messages and view details (exportable)  
- **Auto Reply**: Create/manage automatic replies (out-of-office)  

#### System Configuration
- **Regionsettings**:
- **EXO Settings**: Read and adjust global organization settings  

#### Monitoring & Support
- **Health Check**: Check system health and service status  
- **Mailbox Audit**: Audit logs and auditing settings  
- **Reports**: Generate standard reports (e.g., mailbox sizes)  
- **Troubleshooting**: Utilities and log analysis  

### 📂 Logs
- Folder: `Logs`  

### 🔗 References
- [easyEXO project](https://github.com/PS-easyIT/easyEXO)  
- [Exchange Online PowerShell docs](https://aka.ms/exops-docs)
