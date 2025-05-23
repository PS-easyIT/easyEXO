<# 
=======================================================================
Exchange Berechtigungen Verwaltung Tool
=======================================================================
Beschreibung:
Dieses PowerShell-Skript implementiert ein grafisches Tool (WPF) zur 
Verwaltung von Exchange Online Kalender- und Postfachberechtigungen.
Alle Operationen erfolgen ausschließlich über PowerShell-Cmdlets 
des ExchangeOnlineManagement-Moduls.

Voraussetzungen:
- PowerShell 5.1 oder höher
- ExchangeOnlineManagement Modul installiert
- Internetverbindung

Nutzung:
Führen Sie das Skript aus. Klicken Sie auf "Mit Exchange verbinden" und 
geben Sie den UserPrincipalName (UPN) ein, um die Verbindung zu 
Exchange Online herzustellen. Wählen Sie anschließend den entsprechenden 
Tab (Kalender- oder Postfachberechtigungen) und füllen Sie die Felder aus, 
um Berechtigungen hinzuzufügen oder zu entfernen.
Alle Benutzeraktionen werden in der Datei Logs\ExchangeTool.log protokolliert.

Änderungshistorie:
v0.0.4 - Ursprüngliche Version
v0.0.5 - Korrektur für Get-ThrottlingPolicy und modernere EXO-Befehle
v0.0.6 - Fehlerbehebung bei der GUI-Initialisierung und Logging und weitere Funktionen
=======================================================================
#>
# Überprüfen, ob das Skript aus easyStartUp_V0.0.3.ps1 gestartet wurde
param (
    [Parameter(Mandatory=$false)]
    [hashtable]$InheritedConnectionState
)

# Globale Verbindungsvariable
if ($null -eq $Global:IsConnectedToExo) {
    $Global:IsConnectedToExo = $false
}

# Verbindungserkennung vom Hauptskript übernehmen
if ($null -ne $InheritedConnectionState) {
    Write-Host "Verbindungsstatus übernommen: Exchange Online $(if($InheritedConnectionState.ExchangeOnline){"verbunden"}else{"nicht verbunden"})" -ForegroundColor Cyan
    
    if ($InheritedConnectionState.ExchangeOnline) {
        $Global:IsConnectedToExo = $true
    }
    
    # Prüfungsfunktion übernehmen wenn vorhanden
    if ($null -ne ${global:EasyStartup_TestExchangeOnlineConnection}) {
        function Test-ExchangeOnlineConnection {
            & ${global:EasyStartup_TestExchangeOnlineConnection}
        }
    }
}
# -------------------------------------------------
# Check for PowerShell 7
# -------------------------------------------------
# Simple internal logging function to ensure logging works before the main Log-Action is defined
function Write-LogEntry {
    param([string]$Message)
    $logFolder = "$PSScriptRoot\Logs"
    if (-not (Test-Path $logFolder)) {
        New-Item -ItemType Directory -Path $logFolder | Out-Null
    }
    $logFile = Join-Path $logFolder "ExchangeTool.log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$timestamp - $Message"
}

function Test-PowerShell7 {
    [CmdletBinding()]
    param()
    
    try {
        # Check if running in PowerShell 7+
        $psVersion = $PSVersionTable.PSVersion
        $isPSCore = $psVersion.Major -ge 7
        
        if ($isPSCore) {
            Write-Host "PowerShell 7+ detected. Version: $($psVersion.ToString())" -ForegroundColor Green
            Write-LogEntry "Using PowerShell version: $($psVersion.ToString())"
            return $true
        }
        
        Write-Host "Current PowerShell version: $($psVersion.ToString())" -ForegroundColor Yellow
        Write-Host "This script works best with PowerShell 7 or later." -ForegroundColor Yellow
        
        # Check if PowerShell 7 is installed but not being used
        $ps7Paths = @(
            "${env:ProgramFiles}\PowerShell\7\pwsh.exe", 
            "${env:ProgramFiles(x86)}\PowerShell\7\pwsh.exe",
            "$env:LOCALAPPDATA\Programs\PowerShell\7\pwsh.exe"
        )
        
        $ps7Path = $ps7Paths | Where-Object { Test-Path $_ } | Select-Object -First 1
        
        if ($ps7Path) {
            $message = "PowerShell 7 ist installiert, aber Sie verwenden derzeit PowerShell $($psVersion.Major).$($psVersion.Minor). Möchten Sie dieses Skript mit PowerShell 7 neu starten?"
            $result = [System.Windows.MessageBox]::Show(
                $message,
                "PowerShell Versionsüberprüfung",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                # Start PowerShell 7 with the current script
                $scriptPath = $MyInvocation.MyCommand.Path
                
                # Start PowerShell 7 with the current script
                $argList = "-File `"$scriptPath`""
                Start-Process -FilePath $ps7Path -ArgumentList $argList
                
                # Exit current PowerShell session
                exit
            }
        }
        else {
            $message = "PowerShell 7 wurde auf Ihrem System nicht gefunden. Möchten Sie es jetzt installieren? (Empfohlen)"
                    # Check if winget is available for easier installation
                    $useWinget = $false
                    try {
                        # Run winget command and check exit code without storing unused output
                        winget --version > $null 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            $useWinget = $true
                        }
                    }
                    catch {
                        $useWinget = $false
                    }
                    
                    # Check if winget is available for easier installation
                    $useWinget = $false
                    try {
                        $wingetVersion = (winget --version) 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            $useWinget = $true
                        }
                    }
                    catch {
                        $useWinget = $false
                    }
                    
                    if ($useWinget) {
                        # Use winget for installation
                        Write-Host "Verwende winget zur Installation von PowerShell 7..." -ForegroundColor Cyan
                        Start-Process -FilePath "winget" -ArgumentList "install Microsoft.PowerShell" -Wait -NoNewWindow
                    }
                    else {
                        # Fall back to the MSI installer
                        Write-Host "Lade PowerShell 7 Installer herunter..." -ForegroundColor Cyan
                        $installerUrl = "https://github.com/PowerShell/PowerShell/releases/download/v7.3.6/PowerShell-7.3.6-win-x64.msi"
                        $installerPath = "$env:TEMP\PowerShell-7.3.6-win-x64.msi"
                        
                        # Download the installer
                        Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
                        
                        # Install PowerShell 7
                        Write-Host "Führe Installer aus. Bitte folgen Sie den Anweisungen..." -ForegroundColor Cyan
                        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$installerPath`" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1" -Wait
                        
                        # Clean up
                        Remove-Item -Path $installerPath -Force
                    }
                    
                    Write-Host "PowerShell 7 Installation abgeschlossen. Bitte starten Sie dieses Skript mit PowerShell 7 neu." -ForegroundColor Green
                    
                    # Check if PowerShell 7 was successfully installed
                    $ps7Path = $ps7Paths | Where-Object { Test-Path $_ } | Select-Object -First 1
                    
                    if ($ps7Path) {
                        $startNow = [System.Windows.MessageBox]::Show(
                            "PowerShell 7 wurde installiert. Möchten Sie dieses Skript jetzt mit PowerShell 7 neu starten?",
                            "Installation abgeschlossen",
                            [System.Windows.MessageBoxButton]::YesNo,
                            [System.Windows.MessageBoxImage]::Question
                        )
                        
                        if ($startNow -eq [System.Windows.MessageBoxResult]::Yes) {
                            # Start PowerShell 7 with the current script
                            $scriptPath = $MyInvocation.MyCommand.Path
                            Start-Process -FilePath $ps7Path -ArgumentList "-File `"$scriptPath`""
                            
                            # Exit current PowerShell session
                            exit
                        }
                    }
                }
                
        # If we got here, we're continuing with the current PowerShell version
        Write-Host "Fahre mit PowerShell $($psVersion.ToString()) fort." -ForegroundColor Yellow
        Write-LogEntry "Verwende PowerShell Version: $($psVersion.ToString()) (nicht empfohlen)"
        return $false
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Host "Fehler bei der Überprüfung der PowerShell Version: $errorMsg" -ForegroundColor Red
        Write-LogEntry "Fehler bei der Überprüfung der PowerShell Version: $errorMsg"
        return $false
    }
}

# Check for PowerShell 7 at startup
Test-PowerShell7
# --------------------------------------------------------------
# Initialisiere Debugging und Logging für das Script
# --------------------------------------------------------------
$script:debugMode = $false
$script:logFilePath = Join-Path -Path "$PSScriptRoot\Logs" -ChildPath "ExchangeTool.log"

# Assembly für WPF-Komponenten laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Definiere Farben für GUI
$script:connectedBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Green)
$script:disconnectedBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Red)
$script:isConnected = $false

# INI-Datei für Konfigurationseinstellungen
$script:configFilePath = Join-Path -Path $PSScriptRoot -ChildPath "config.ini"

# Erstelle INI-Datei, falls sie nicht existiert
if (-not (Test-Path -Path $script:configFilePath)) {
    try {
        $configFolder = Split-Path -Path $script:configFilePath -Parent
        if (-not (Test-Path -Path $configFolder)) {
            New-Item -ItemType Directory -Path $configFolder -Force | Out-Null
        }
        
        $defaultConfig = @"
[General]
Debug = 1
AppName = Exchange Online Verwaltung
Version = 0.0.6
ThemeColor = #0078D7
DarkMode = 0

[Paths]
LogPath = $PSScriptRoot\Logs

[UI]
HeaderLogoURL = https://www.microsoft.com/de-de/microsoft-365/exchange/email
"@
        Set-Content -Path $script:configFilePath -Value $defaultConfig -Encoding UTF8
        Write-Host "Konfigurationsdatei wurde erstellt: $script:configFilePath" -ForegroundColor Green
    }
    catch {
        Write-Host "Fehler beim Erstellen der Konfigurationsdatei: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Lade Konfiguration aus INI-Datei
function Get-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        $ini = @{
        }
        switch -regex -file $FilePath {
            "^\[(.+)\]" {
                $section = $matches[1]
                $ini[$section] = @{
                }
                continue
            }
            "^\s*([^#].+?)\s*=\s*(.*)" {
                $name, $value = $matches[1..2]
                $ini[$section][$name] = $value
                continue
            }
        }
        return $ini
    }
    catch {
        Write-Host "Fehler beim Lesen der INI-Datei: $($_.Exception.Message)" -ForegroundColor Red
        return @{
        }
    }
}

# Lade Konfiguration
try {
    $script:config = Get-IniContent -FilePath $script:configFilePath
    
    # Debug-Modus einschalten, wenn in INI aktiviert
    if ($script:config["General"]["Debug"] -eq "1") {
        $script:debugMode = $true
        Write-Host "Debug-Modus ist aktiviert" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "Fehler beim Laden der Konfiguration: $($_.Exception.Message)" -ForegroundColor Red
    # Fallback zu Standardwerten
    $script:config = @{
        "General" = @{
            "Debug" = "1"
            "AppName" = "Exchange Berechtigungen Verwaltung"
            "Version" = "0.0.1"
            "ThemeColor" = "#0078D7"
            "DarkMode" = "0"
        }
        "Paths" = @{
            "LogPath" = "$PSScriptRoot\Logs"
        }
    }
    $script:debugMode = $true
}

# --------------------------------------------------------------
# Verbesserte Debug- und Logging-Funktionen
# --------------------------------------------------------------
function Write-DebugMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Type = "Info"
    )
    
    try {
        if ($script:debugMode) {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $colorMap = @{
                "Info" = "Cyan"
                "Warning" = "Yellow"
                "Error" = "Red"
                "Success" = "Green"
            }
            
            # Sicherstellen, dass nur druckbare ASCII-Zeichen verwendet werden
            $sanitizedMessage = $Message -replace '[^\x20-\x7E]', '?'
            
            # Ausgabe auf der Konsole
            Write-Host "[$timestamp] [DEBUG] [$Type] $sanitizedMessage" -ForegroundColor $colorMap[$Type]
            
            # Auch ins Log schreiben
            Log-Action "DEBUG: $Type - $sanitizedMessage"
        }
    }
    catch {
        # Fallback für Fehler in der Debug-Funktion - schreibe direkt ins Log
        try {
            $errorMsg = $_.Exception.Message -replace '[^\x20-\x7E]', '?'
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logFolder = "$PSScriptRoot\Logs"
            if (-not (Test-Path $logFolder)) {
                New-Item -ItemType Directory -Path $logFolder | Out-Null
            }
            $fallbackLogFile = Join-Path $logFolder "debug_fallback.log"
            Add-Content -Path $fallbackLogFile -Value "[$timestamp] Fehler in Write-DebugMessage: $errorMsg" -Encoding UTF8
        }
        catch {
            # Absoluter Fallback - ignoriere Fehler um Programmablauf nicht zu stören
        }
    }
}

function Log-Action {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message
    )
    
    try {
        # Sicherstellen, dass nur druckbare ASCII-Zeichen verwendet werden
        $sanitizedMessage = $Message -replace '[^\x20-\x7E]', '?'
        
        # Zeitstempel erzeugen
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        # Logverzeichnis erstellen, falls nicht vorhanden
        $logFolder = Split-Path -Path $script:logFilePath -Parent
        if (-not (Test-Path $logFolder)) {
            New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
            Write-DebugMessage "Logverzeichnis wurde erstellt: $logFolder" -Type "Info"
        }
        
        # Log-Eintrag schreiben
        Add-Content -Path $script:logFilePath -Value "[$timestamp] $sanitizedMessage" -Encoding UTF8
        
        # Bei zu langer Logdatei (>10 MB) rotieren
        $logFile = Get-Item -Path $script:logFilePath -ErrorAction SilentlyContinue
        if ($logFile -and $logFile.Length -gt 10MB) {
            $backupLogPath = "$($script:logFilePath)_$(Get-Date -Format 'yyyyMMdd_HHmmss').bak"
            Move-Item -Path $script:logFilePath -Destination $backupLogPath -Force
            Write-DebugMessage "Logdatei wurde rotiert: $backupLogPath" -Type "Info"
        }
    }
    catch {
        # Fallback für Fehler in der Log-Funktion
        try {
            $errorMsg = $_.Exception.Message -replace '[^\x20-\x7E]', '?'
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $fallbackLogFile = Join-Path -Path "$PSScriptRoot\Logs" -ChildPath "log_fallback.log"
            $fallbackLogFolder = Split-Path -Path $fallbackLogFile -Parent
            
            if (-not (Test-Path $fallbackLogFolder)) {
                New-Item -ItemType Directory -Path $fallbackLogFolder -Force | Out-Null
            }
            
            Add-Content -Path $fallbackLogFile -Value "[$timestamp] Fehler in Log-Action: $errorMsg" -Encoding UTF8
            Add-Content -Path $fallbackLogFile -Value "[$timestamp] Ursprüngliche Nachricht: $sanitizedMessage" -Encoding UTF8
        }
        catch {
            # Absoluter Fallback - ignoriere Fehler um Programmablauf nicht zu stören
        }
    }
}

# Funktion zum Aktualisieren der GUI-Textanzeige mit Fehlerbehandlung
function Update-GuiText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.TextBlock]$TextElement,
        
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [System.Windows.Media.Brush]$Color = $null,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxLength = 10000
    )
    
    try {
        if ($null -eq $TextElement) {
            Write-DebugMessage "GUI-Element ist null in Update-GuiText" -Type "Warning"
            return
        }
        
        # Sicherstellen, dass nur druckbare ASCII-Zeichen verwendet werden
        $sanitizedMessage = $Message -replace '[^\x20-\x7E]', '?'
        
        # Nachricht auf maximale Länge begrenzen
        if ($sanitizedMessage.Length -gt $MaxLength) {
            $sanitizedMessage = $sanitizedMessage.Substring(0, $MaxLength) + "..."
        }
        
        # GUI-Element im UI-Thread aktualisieren
        $TextElement.Dispatcher.Invoke([Action]{
            $TextElement.Text = $sanitizedMessage
            if ($null -ne $Color) {
                $TextElement.Foreground = $Color
            }
        }, "Normal")
    }
    catch {
        try {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler in Update-GuiText: $errorMsg" -Type "Error"
            Log-Action "GUI-Ausgabefehler: $errorMsg"
        }
        catch {
            # Ignoriere Fehler in der Fehlerbehandlung
        }
    }
}

# -------------------------------------------------
# Abschnitt: Selbstdiagnose
# -------------------------------------------------
function Test-ModuleInstalled {
    param([string]$ModuleName)
    try {
        if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
            # Return false without throwing an error
            return $false
        }
        return $true
    } catch {
        # Log the error silently without Write-Error
        $errorMessage = $_.Exception.Message
        Log-Action "Fehler beim Prüfen des Moduls $ModuleName - $errorMessage"
        return $false
    }
}

function Test-InternetConnection {
    try {
        $ping = Test-Connection -ComputerName "www.google.com" -Count 1 -Quiet
        if (-not $ping) { throw "Keine Internetverbindung." }
        return $true
    } catch {
        Write-Error $_.Exception.Message
        return $false
    }
}

# -------------------------------------------------
# Abschnitt: Logging
# -------------------------------------------------
function Log-Action {
    param([string]$Message)
    $logFolder = "$PSScriptRoot\Logs"
    if (-not (Test-Path $logFolder)) {
        New-Item -ItemType Directory -Path $logFolder | Out-Null
    }
    $logFile = Join-Path $logFolder "ExchangeTool.log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$timestamp - $Message"
}

# -------------------------------------------------
# Abschnitt: Eingabevalidierung
# -------------------------------------------------
function Validate-Email {
    param([string]$Email)
    $regex = '^[\w\.\-]+@([\w\-]+\.)+[a-zA-Z]{2,}$'
    return $Email -match $regex
}

# -------------------------------------------------
# Abschnitt: Exchange Online Verbindung
# -------------------------------------------------
function Connect-ExchangeOnlineSession {
    [CmdletBinding()]
    param()
    
    try {
        # Status aktualisieren und loggen
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Verbindung wird hergestellt..."
        }
        Log-Action "Verbindungsversuch zu Exchange Online mit ModernAuth"
        
        # Prüfen, ob bereits eine Verbindung vom Hauptskript besteht
        if ($Global:IsConnectedToExo -eq $true) {
            Write-DebugMessage "Bestehende Exchange Online-Verbindung vom Hauptskript erkannt" -Type "Info"
            
            # Verbindung testen falls möglich
            if ($null -ne ${global:EasyStartup_TestExchangeOnlineConnection}) {
                $connectionValid = & ${global:EasyStartup_TestExchangeOnlineConnection}
                if (-not $connectionValid) {
                    Write-DebugMessage "Bestehende Verbindung ist nicht mehr gültig, stelle neue Verbindung her" -Type "Warning"
                    $Global:IsConnectedToExo = $false
                }
            }
            else {
                # Alternative Prüfung wenn keine Test-Funktion übergeben wurde
                try {
                    # Einfache Prüfung durch Ausführen eines kleinen Exchange-Befehls
                    Get-OrganizationConfig -ErrorAction Stop | Out-Null
                    $connectionValid = $true
                }
                catch {
                    Write-DebugMessage "Bestehende Verbindung ist nicht mehr gültig, stelle neue Verbindung her" -Type "Warning"
                    $Global:IsConnectedToExo = $false
                }
            }
        }
        
        # Wenn keine gültige Verbindung besteht, neue herstellen
        if ($Global:IsConnectedToExo -ne $true) {
            # Prüfen, ob Modul installiert ist
            if (-not (Test-ModuleInstalled -ModuleName "ExchangeOnlineManagement")) {
                throw "ExchangeOnlineManagement Modul ist nicht installiert. Bitte installieren Sie das Modul über den 'Installiere Module' Button."
            }
            
            # Modul laden
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            
            # ModernAuth-Verbindung herstellen (nutzt automatisch die Standardbrowser-Authentifizierung)
            Connect-ExchangeOnline -ShowBanner:$false -ShowProgress $true -ErrorAction Stop
            
            # Bei erfolgreicher Verbindung
            Log-Action "Exchange Online Verbindung hergestellt mit ModernAuth (MFA)"
            $Global:IsConnectedToExo = $true
        }
        else {
            Log-Action "Bestehende Exchange Online Verbindung vom Hauptskript übernommen"
        }
        
        # GUI-Elemente aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Mit Exchange verbunden"
        }
        if ($null -ne $txtConnectionStatus) {
            $txtConnectionStatus.Text = "Verbunden"
            $txtConnectionStatus.Foreground = $script:connectedBrush
        }
        $script:isConnected = $true
        
        # Button-Status aktualisieren
        if ($null -ne $btnConnect) {
            $btnConnect.Content = "Verbindung trennen"
            $btnConnect.Tag = "disconnect"
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler beim Verbinden: $errorMsg"
        }
        Log-Action "Fehler beim Verbinden: $errorMsg"
        
        # Zeige Fehlermeldung an den Benutzer
        try {
            [System.Windows.MessageBox]::Show(
                "Fehler bei der Verbindung zu Exchange Online: $errorMsg", 
                "Verbindungsfehler", 
                [System.Windows.MessageBoxButton]::OK, 
                [System.Windows.MessageBoxImage]::Error)
        }
        catch {
            # Fallback, falls MessageBox fehlschlägt
            Write-Host "Fehler bei der Verbindung zu Exchange Online: $errorMsg" -ForegroundColor Red
        }
        
        return $false
    }
}

function Disconnect-ExchangeOnlineSession {
    [CmdletBinding()]
    param()
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Log-Action "Exchange Online Verbindung getrennt"
        
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Exchange Verbindung getrennt"
        }
        if ($null -ne $txtConnectionStatus) {
            $txtConnectionStatus.Text = "Nicht verbunden"
            $txtConnectionStatus.Foreground = $script:disconnectedBrush
        }
        $script:isConnected = $false
        
        # Button-Status aktualisieren
        if ($null -ne $btnConnect) {
            $btnConnect.Content = "Mit Exchange verbinden"
            $btnConnect.Tag = "connect"
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler beim Trennen der Verbindung: $errorMsg"
        }
        Log-Action "Fehler beim Trennen der Verbindung: $errorMsg"
        
        # Zeige Fehlermeldung an den Benutzer
        try {
            [System.Windows.MessageBox]::Show(
                "Fehler beim Trennen der Verbindung: $errorMsg", 
                "Fehler", 
                [System.Windows.MessageBoxButton]::OK, 
                [System.Windows.MessageBoxImage]::Error)
        }
        catch {
            # Fallback, falls MessageBox fehlschlägt
            Write-Host "Fehler beim Trennen der Verbindung: $errorMsg" -ForegroundColor Red
        }
        
        return $false
    }
}

# Funktion zum Überprüfen der Voraussetzungen (Module)
function Check-Prerequisites {
    [CmdletBinding()]
    param()
    
    try {
        Write-DebugMessage "Überprüfe benötigte PowerShell-Module" -Type "Info"
        
        $missingModules = @()
        $requiredModules = @(
            @{Name = "ExchangeOnlineManagement"; MinVersion = "3.0.0"; Description = "Exchange Online Management"}
        )
        
        $results = @()
        $allModulesInstalled = $true
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Überprüfe installierte Module..."
        }
        
        foreach ($moduleInfo in $requiredModules) {
            $moduleName = $moduleInfo.Name
            $minVersion = $moduleInfo.MinVersion
            $description = $moduleInfo.Description
            
            # Prüfe, ob Modul installiert ist
            $module = Get-Module -Name $moduleName -ListAvailable -ErrorAction SilentlyContinue
            
            if ($null -ne $module) {
                # Prüfe Modul-Version, falls erforderlich
                $latestVersion = ($module | Sort-Object Version -Descending | Select-Object -First 1).Version
                
                if ($null -ne $minVersion -and $latestVersion -lt [Version]$minVersion) {
                    $results += [PSCustomObject]@{
                        Module = $moduleName
                        Status = "Update erforderlich"
                        Installiert = $latestVersion
                        Erforderlich = $minVersion
                        Beschreibung = $description
                    }
                    $missingModules += $moduleInfo
                    $allModulesInstalled = $false
                } else {
                    $results += [PSCustomObject]@{
                        Module = $moduleName
                        Status = "Installiert"
                        Installiert = $latestVersion
                        Erforderlich = $minVersion
                        Beschreibung = $description
                    }
                }
            } else {
                $results += [PSCustomObject]@{
                    Module = $moduleName
                    Status = "Nicht installiert"
                    Installiert = "---"
                    Erforderlich = $minVersion
                    Beschreibung = $description
                }
                $missingModules += $moduleInfo
                $allModulesInstalled = $false
            }
        }
        
        # Ergebnis anzeigen
        $resultText = "Prüfergebnis der benötigten Module:`n`n"
        foreach ($result in $results) {
            $statusIcon = switch ($result.Status) {
                "Installiert" { "✅" }
                "Update erforderlich" { "⚠️" }
                "Nicht installiert" { "❌" }
                default { "❓" }
            }
            
            $resultText += "$statusIcon $($result.Module): $($result.Status)"
            if ($result.Status -ne "Installiert") {
                $resultText += " (Installiert: $($result.Installiert), Erforderlich: $($result.Erforderlich))"
            } else {
                $resultText += " (Version: $($result.Installiert))"
            }
            $resultText += " - $($result.Beschreibung)`n"
        }
        
        $resultText += "`n"
        
        if ($allModulesInstalled) {
            $resultText += "Alle erforderlichen Module sind installiert. Sie können Exchange Online verwenden."
            
            if ($null -ne $txtStatus) {
                $txtStatus.Text = "Alle Module erfolgreich installiert."
                $txtStatus.Foreground = $script:connectedBrush
            }
        } else {
            $resultText += "Es fehlen erforderliche Module. Bitte klicken Sie auf 'Installiere Module', um diese zu installieren."
            
            if ($null -ne $txtStatus) {
                $txtStatus.Text = "Es fehlen erforderliche Module."
            }
        }
        
        # Ergebnis in einem MessageBox anzeigen
        [System.Windows.MessageBox]::Show(
            $resultText,
            "Modul-Überprüfung",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        
        # Return-Wert (für Skript-Logik)
        return @{
            AllInstalled = $allModulesInstalled
            MissingModules = $missingModules
            Results = $results
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler bei der Überprüfung der Module: $errorMsg" -Type "Error"
        
        [System.Windows.MessageBox]::Show(
            "Fehler bei der Überprüfung der Module: $errorMsg",
            "Fehler",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler bei der Überprüfung der Module."
        }
        
        return @{
            AllInstalled = $false
            Error = $errorMsg
        }
    }
}

# Funktion zum Installieren der fehlenden Module
function Install-Prerequisites {
    [CmdletBinding()]
    param()
    
    try {
        Write-DebugMessage "Installiere benötigte PowerShell-Module" -Type "Info"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Überprüfe und installiere Module..."
        }
        
        # Benötigte Module definieren
        $requiredModules = @(
            @{Name = "ExchangeOnlineManagement"; MinVersion = "3.0.0"; Description = "Exchange Online Management"}
        )
        
        # Überprüfe, ob PowerShellGet aktuell ist
        $psGetVersion = (Get-Module PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version
        
        if ($null -eq $psGetVersion -or $psGetVersion -lt [Version]"2.0.0") {
            Write-DebugMessage "PowerShellGet-Modul ist veraltet oder nicht installiert, versuche zu aktualisieren" -Type "Warning"
            
            # Prüfen, ob bereits als Administrator ausgeführt
            $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
            
            if (-not $isAdmin) {
                $message = "Für die Installation/Aktualisierung des PowerShellGet-Moduls werden Administratorrechte benötigt.`n`n" + 
                           "Möchten Sie PowerShell als Administrator neu starten und anschließend das Tool erneut ausführen?"
                           
                $result = [System.Windows.MessageBox]::Show(
                    $message, 
                    "Administratorrechte erforderlich", 
                    [System.Windows.MessageBoxButton]::YesNo, 
                    [System.Windows.MessageBoxImage]::Question
                )
                
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    # Starte PowerShell als Administrator neu
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName = "powershell.exe"
                    $psi.Arguments = "-ExecutionPolicy Bypass -Command ""Start-Process PowerShell -Verb RunAs -ArgumentList '-ExecutionPolicy Bypass -NoExit -Command ""Install-Module PowerShellGet -Force -AllowClobber; Install-Module ExchangeOnlineManagement -Force -AllowClobber""'"""
                    $psi.Verb = "runas"
                    [System.Diagnostics.Process]::Start($psi)
                    
                    # Aktuelle Instanz schließen
                    if ($null -ne $script:Form) {
                        $script:Form.Close()
                    }
                    
                    return @{
                        Success = $false
                        AdminRestartRequired = $true
                    }
                } else {
                    Write-DebugMessage "Benutzer hat Administrator-Neustart abgelehnt" -Type "Warning"
                    
                    [System.Windows.MessageBox]::Show(
                        "Die Modulinstallation wird ohne Administratorrechte fortgesetzt, es können jedoch Fehler auftreten.",
                        "Fortsetzen ohne Administratorrechte",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                }
            } else {
                # Versuche, PowerShellGet zu aktualisieren
                try {
                    Install-Module PowerShellGet -Force -AllowClobber
                    Write-DebugMessage "PowerShellGet erfolgreich aktualisiert" -Type "Success"
                } catch {
                    Write-DebugMessage "Fehler beim Aktualisieren von PowerShellGet: $($_.Exception.Message)" -Type "Error"
                    # Fortfahren trotz Fehler
                }
            }
        }
        
        # Installiere jedes Modul
        $results = @()
        $allSuccess = $true
        
        foreach ($moduleInfo in $requiredModules) {
            $moduleName = $moduleInfo.Name
            $minVersion = $moduleInfo.MinVersion
            
            Write-DebugMessage "Installiere/Aktualisiere Modul: $moduleName" -Type "Info"
            
            try {
                # Prüfe, ob Modul bereits installiert ist
                $module = Get-Module -Name $moduleName -ListAvailable -ErrorAction SilentlyContinue
                
                if ($null -ne $module) {
                    $latestVersion = ($module | Sort-Object Version -Descending | Select-Object -First 1).Version
                    
                    # Prüfe, ob Update notwendig ist
                    if ($null -ne $minVersion -and $latestVersion -lt [Version]$minVersion) {
                        Write-DebugMessage "Aktualisiere Modul $moduleName von $latestVersion auf mindestens $minVersion" -Type "Info"
                        Install-Module -Name $moduleName -Force -AllowClobber -MinimumVersion $minVersion
                        $newVersion = (Get-Module -Name $moduleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version
                        
                        $results += [PSCustomObject]@{
                            Module = $moduleName
                            Status = "Aktualisiert"
                            AlteVersion = $latestVersion
                            NeueVersion = $newVersion
                        }
                    } else {
                        Write-DebugMessage "Modul $moduleName ist bereits in ausreichender Version ($latestVersion) installiert" -Type "Info"
                        
                        $results += [PSCustomObject]@{
                            Module = $moduleName
                            Status = "Bereits aktuell"
                            AlteVersion = $latestVersion
                            NeueVersion = $latestVersion
                        }
                    }
                } else {
                    # Installiere Modul
                    Write-DebugMessage "Installiere Modul $moduleName" -Type "Info"
                    Install-Module -Name $moduleName -Force -AllowClobber
                    $newVersion = (Get-Module -Name $moduleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version
                    
                    $results += [PSCustomObject]@{
                        Module = $moduleName
                        Status = "Neu installiert"
                        AlteVersion = "---"
                        NeueVersion = $newVersion
                    }
                }
            } catch {
                $errorMsg = $_.Exception.Message
                Write-DebugMessage "Fehler beim Installieren/Aktualisieren von $moduleName - $errorMsg" -Type "Error"
                
                $results += [PSCustomObject]@{
                    Module = $moduleName
                    Status = "Fehler"
                    AlteVersion = "---"
                    NeueVersion = "---"
                    Fehler = $errorMsg
                }
                
                $allSuccess = $false
            }
        }
        
        # Ergebnis anzeigen
        $resultText = "Ergebnis der Modulinstallation:`n`n"
        foreach ($result in $results) {
            $statusIcon = switch ($result.Status) {
                "Neu installiert" { "✅" }
                "Aktualisiert" { "✅" }
                "Bereits aktuell" { "✅" }
                "Fehler" { "❌" }
                default { "❓" }
            }
            
            $resultText += "$statusIcon $($result.Module): $($result.Status)"
            if ($result.Status -eq "Aktualisiert") {
                $resultText += " (Von Version $($result.AlteVersion) auf $($result.NeueVersion))"
            } elseif ($result.Status -eq "Neu installiert") {
                $resultText += " (Version $($result.NeueVersion))"
            } elseif ($result.Status -eq "Fehler") {
                $resultText += " - Fehler: $($result.Fehler)"
            }
            $resultText += "`n"
        }
        
        $resultText += "`n"
        
        if ($allSuccess) {
            $resultText += "Alle Module wurden erfolgreich installiert oder waren bereits aktuell.`n"
            $resultText += "Sie können das Tool verwenden."
            
            if ($null -ne $txtStatus) {
                $txtStatus.Text = "Alle Module erfolgreich installiert."
                $txtStatus.Foreground = $script:connectedBrush
            }
        } else {
            $resultText += "Bei der Installation einiger Module sind Fehler aufgetreten.`n"
            $resultText += "Starten Sie PowerShell mit Administratorrechten und versuchen Sie es erneut."
            
            if ($null -ne $txtStatus) {
                $txtStatus.Text = "Fehler bei der Modulinstallation aufgetreten."
            }
        }
        
        # Ergebnis in einem MessageBox anzeigen
        [System.Windows.MessageBox]::Show(
            $resultText,
            "Modul-Installation",
            [System.Windows.MessageBoxButton]::OK,
            $allSuccess ? [System.Windows.MessageBoxImage]::Information : [System.Windows.MessageBoxImage]::Warning
        )
        
        # Return-Wert (für Skript-Logik)
        return @{
            Success = $allSuccess
            Results = $results
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler bei der Modulinstallation: $errorMsg" -Type "Error"
        
        [System.Windows.MessageBox]::Show(
            "Fehler bei der Modulinstallation: $errorMsg`n`nVersuchen Sie, PowerShell als Administrator auszuführen und wiederholen Sie den Vorgang.",
            "Fehler",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler bei der Modulinstallation."
        }
        
        return @{
            Success = $false
            Error = $errorMsg
        }
    }
}

# -------------------------------------------------
# Abschnitt: Kalenderberechtigungen
# -------------------------------------------------
function Get-CalendarPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $MailboxUser)) {
            throw "Ungültige E-Mail-Adresse für Postfach."
        }
        
        Write-DebugMessage "Rufe Kalenderberechtigungen ab für: $MailboxUser" -Type "Info"
        
        # Prüfe deutsche und englische Kalenderordner
        $permissions = $null
        try {
            # Versuche mit deutschem Pfad
            $identity = "${MailboxUser}:\Kalender"
            Write-DebugMessage "Versuche deutschen Kalenderpfad: $identity" -Type "Info"
            $permissions = Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop
        } 
        catch {
            try {
                # Versuche mit englischem Pfad
                $identity = "${MailboxUser}:\Calendar"
                Write-DebugMessage "Versuche englischen Kalenderpfad: $identity" -Type "Info"
                $permissions = Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop
            } 
            catch {
                $errorMsg = $_.Exception.Message
                Write-DebugMessage "Beide Kalenderpfade fehlgeschlagen: $errorMsg" -Type "Error"
                throw "Kalenderordner konnte nicht gefunden werden. Weder 'Kalender' noch 'Calendar' sind zugänglich."
            }
        }
        
        Write-DebugMessage "Kalenderberechtigungen abgerufen: $($permissions.Count) Einträge gefunden" -Type "Success"
        Log-Action "Kalenderberechtigungen für $MailboxUser erfolgreich abgerufen: $($permissions.Count) Einträge."
        return $permissions
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Kalenderberechtigungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Kalenderberechtigungen: $errorMsg"
        throw $errorMsg
    }
}

# Funktion für das Anzeigen aller Kalenderberechtigungen
function Show-CalendarPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxUser
    )
    
    try {
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange verbunden. Bitte stellen Sie zuerst eine Verbindung her."
        }
        
        # Prüfe, ob eine gültige E-Mail-Adresse eingegeben wurde
        if (-not (Validate-Email -Email $MailboxUser)) {
            throw "Bitte geben Sie eine gültige E-Mail-Adresse ein."
        }
        
        # Status aktualisieren
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Rufe Kalenderberechtigungen ab..."
        }
        
        # Versuche Kalenderberechtigungen abzurufen
        $permissions = Get-CalendarPermission -MailboxUser $MailboxUser
        
        # Aufbereiten der Berechtigungsdaten für die DataGrid-Anzeige
        $permissionsForGrid = @()
        foreach ($permission in $permissions) {
            # Extrahiere die relevanten Informationen und erstelle ein neues Objekt
            $permObj = [PSCustomObject]@{
                User = $permission.User.DisplayName
                AccessRights = ($permission.AccessRights -join ", ")
                IsInherited = $permission.IsInherited
            }
            $permissionsForGrid += $permObj
        }
        
        # Aktualisiere das DataGrid mit den aufbereiteten Daten
        if ($null -ne $script:lstCalendarPermissions) {
            $script:lstCalendarPermissions.Dispatcher.Invoke([Action]{
                $script:lstCalendarPermissions.ItemsSource = $permissionsForGrid
            }, "Normal")
        }
        
        # Status aktualisieren
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Kalenderberechtigungen erfolgreich abgerufen."
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Anzeigen der Kalenderberechtigungen: $errorMsg" -Type "Error"
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler: $errorMsg"
        }
        
        return $false
    }
}

# Fix for Set-CalendarDefaultPermissionsAction function
function Set-CalendarDefaultPermissionsAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Standard", "Anonym", "Beides")]
        [string]$PermissionType,
        
        [Parameter(Mandatory = $true)]
        [string]$AccessRights,
        
        [Parameter(Mandatory = $false)]
        [switch]$ForAllMailboxes = $false,
        
        [Parameter(Mandatory = $false)]
        [string]$MailboxUser = ""
    )
    
    try {
        Write-DebugMessage "Setze Standardberechtigungen für Kalender: $PermissionType mit $AccessRights" -Type "Info"
        
        if ($ForAllMailboxes) {
            # Frage den Benutzer ob er das wirklich tun möchte (kann lange dauern)
            $confirmResult = [System.Windows.MessageBox]::Show(
                "Möchten Sie wirklich die $PermissionType-Berechtigungen für ALLE Postfächer setzen? Diese Aktion kann bei vielen Postfächern lange dauern.",
                "Massenänderung bestätigen",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning)
                
            if ($confirmResult -eq [System.Windows.MessageBoxResult]::No) {
                Write-DebugMessage "Massenänderung vom Benutzer abgebrochen" -Type "Info"
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Operation abgebrochen."
                }
                return $false
            }
            
            Log-Action "Starte Setzen von Standardberechtigungen für alle Postfächer: $PermissionType"
            
            if ($PermissionType -eq "Standard" -or $PermissionType -eq "Beides") {
                Set-DefaultCalendarPermissionForAll -AccessRights $AccessRights
            }
            if ($PermissionType -eq "Anonym" -or $PermissionType -eq "Beides") {
                Set-AnonymousCalendarPermissionForAll -AccessRights $AccessRights
            }
        }
        else {         
            if ($null -eq $script:txtCalendarMailboxUser -or [string]::IsNullOrEmpty($script:txtCalendarMailboxUser.Text)) {
                throw "Keine Postfach-E-Mail-Adresse angegeben"
            }
            
            $mailboxUser = $script:txtCalendarMailboxUser.Text
            
            if ($PermissionType -eq "Standard") {
                Set-DefaultCalendarPermission -MailboxUser $mailboxUser -AccessRights $AccessRights
            }
            elseif ($PermissionType -eq "Anonym") {
                Set-AnonymousCalendarPermission -MailboxUser $mailboxUser -AccessRights $AccessRights
            }
            elseif ($PermissionType -eq "Beides") {
                Set-DefaultCalendarPermission -MailboxUser $mailboxUser -AccessRights $AccessRights
                Set-AnonymousCalendarPermission -MailboxUser $mailboxUser -AccessRights $AccessRights
            }
        }
        
        Write-DebugMessage "Standardberechtigungen für Kalender erfolgreich gesetzt: $PermissionType mit $AccessRights" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Standardberechtigungen gesetzt: $PermissionType mit $AccessRights" -Color $script:connectedBrush
        }
        Log-Action "Standardberechtigungen für Kalender gesetzt: $PermissionType mit $AccessRights"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Setzen der Standardberechtigungen für Kalender: $errorMsg" -Type "Error"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        Log-Action "Fehler beim Setzen der Standardberechtigungen für Kalender: $errorMsg"
        return $false
    }
}

function Add-CalendarPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUser,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUser,
        
        [Parameter(Mandatory = $true)]
        [string]$Permission
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Validate-Email -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage "Füge Kalenderberechtigung hinzu/aktualisiere: $SourceUser -> $TargetUser ($Permission)" -Type "Info"
        
        # Prüfe ob Berechtigung bereits existiert
        $calendarExists = $false
        $identityDE = "${SourceUser}:\Kalender"
        $identityEN = "${SourceUser}:\Calendar"
        $identity = $null  # Explizit als null initialisieren
        
        try {
            # Versuche vorhandene Berechtigung mit Get-MailboxFolderPermission zu prüfen
            Write-DebugMessage "Prüfe bestehende Berechtigungen (DE): $identityDE" -Type "Info"
            $existingPerm = Get-MailboxFolderPermission -Identity $identityDE -User $TargetUser -ErrorAction SilentlyContinue
            
            if ($existingPerm) {
                $calendarExists = $true
                $identity = $identityDE
                Write-DebugMessage "Bestehende Berechtigung gefunden (DE): $($existingPerm.AccessRights)" -Type "Info"
            } 
            else {
                Write-DebugMessage "Prüfe bestehende Berechtigungen (EN): $identityEN" -Type "Info"
                $existingPerm = Get-MailboxFolderPermission -Identity $identityEN -User $TargetUser -ErrorAction SilentlyContinue
                
                if ($existingPerm) {
                    $calendarExists = $true
                    $identity = $identityEN
                    Write-DebugMessage "Bestehende Berechtigung gefunden (EN): $($existingPerm.AccessRights)" -Type "Info"
                } 
                else {
                    Write-DebugMessage "Keine bestehende Berechtigung gefunden, füge neu hinzu" -Type "Info"
                    # Hier fehlt die Zuweisung zu $identity
                }
            }
        } 
        catch {
            # Bei Fehler wird automatisch Add-MailboxFolderPermission versucht
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler bei der Prüfung bestehender Berechtigungen: $errorMsg" -Type "Warning"
            
            Write-DebugMessage "Versuche Kalender-Existenz zu prüfen" -Type "Info"
        }
        
        # Wenn $identity immer noch null ist, versuche explizit zu bestimmen, welcher Kalenderordner existiert
        if ($null -eq $identity) {
            try {
                $deExists = Get-MailboxFolderPermission -Identity $identityDE -ErrorAction SilentlyContinue
                if ($deExists) {
                    $identity = $identityDE
                    Write-DebugMessage "Deutscher Kalenderordner gefunden: $identityDE" -Type "Info"
                }
                else {
                    $enExists = Get-MailboxFolderPermission -Identity $identityEN -ErrorAction SilentlyContinue
                    if ($enExists) {
                        $identity = $identityEN
                        Write-DebugMessage "Englischer Kalenderordner gefunden: $identityEN" -Type "Info"
                    }
                    else {
                        # Beide nicht gefunden - versuche Standard-Ansatz mit Kalenderpfaden zu ermitteln
                        try {
                            $folderStats = Get-MailboxFolderStatistics -Identity $SourceUser -FolderScope Calendar -ErrorAction Stop
                            foreach ($folder in $folderStats) {
                                if ($folder.FolderType -eq "Calendar" -or $folder.Name -eq "Kalender" -or $folder.Name -eq "Calendar") {
                                    $identity = "$SourceUser`:" + $folder.FolderPath.Replace("/", "\")
                                    Write-DebugMessage "Kalenderordner über FolderStatistics gefunden: $identity" -Type "Info"
                                    break
                                }
                            }
                        }
                        catch {
                            Write-DebugMessage "Fehler bei der Suche nach Kalenderordnern über FolderStatistics: $($_.Exception.Message)" -Type "Warning"
                        }
                    }
                }
            }
            catch {
                Write-DebugMessage "Fehler beim Abrufen der Kalenderordner: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Wenn immer noch kein Kalender gefunden wurde, Exception werfen
        if ($null -eq $identity) {
            throw "Kein Kalenderordner für $SourceUser gefunden. Bitte stellen Sie sicher, dass das Postfach existiert und Sie Zugriff haben."
        }
        
        # Je nachdem ob Berechtigung existiert, update oder add
        if ($calendarExists) {
            Write-DebugMessage "Aktualisiere bestehende Berechtigung: $identity ($Permission)" -Type "Info"
            Set-MailboxFolderPermission -Identity $identity -User $TargetUser -AccessRights $Permission -ErrorAction Stop
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Kalenderberechtigung aktualisiert." -Color $script:connectedBrush
            }
            
            Write-DebugMessage "Kalenderberechtigung erfolgreich aktualisiert" -Type "Success"
            Log-Action "Kalenderberechtigung aktualisiert: $SourceUser -> $TargetUser mit $Permission"
        } 
        else {
            Write-DebugMessage "Füge neue Berechtigung hinzu: $identity ($Permission)" -Type "Info"
            Add-MailboxFolderPermission -Identity $identity -User $TargetUser -AccessRights $Permission -ErrorAction Stop
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Kalenderberechtigung hinzugefügt." -Color $script:connectedBrush
            }
            
            Write-DebugMessage "Kalenderberechtigung erfolgreich hinzugefügt" -Type "Success"
            Log-Action "Kalenderberechtigung hinzugefügt: $SourceUser -> $TargetUser mit $Permission"
        }
        
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen/Aktualisieren der Kalenderberechtigung: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Hinzufügen/Aktualisieren der Kalenderberechtigung: $errorMsg"
        return $false
    }
}

function Remove-CalendarPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUser,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Validate-Email -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage "Entferne Kalenderberechtigung: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfe deutsche und englische Kalenderordner
        $removed = $false
        
        try {
            $identityDE = "${SourceUser}:\Kalender"
            Write-DebugMessage "Prüfe deutsche Kalenderberechtigungen: $identityDE" -Type "Info"
            
            # Prüfe ob Berechtigung existiert
            $existingPerm = Get-MailboxFolderPermission -Identity $identityDE -User $TargetUser -ErrorAction SilentlyContinue
            
            if ($existingPerm) {
                Write-DebugMessage "Gefundene Berechtigung wird entfernt (DE): $($existingPerm.AccessRights)" -Type "Info"
                Remove-MailboxFolderPermission -Identity $identityDE -User $TargetUser -Confirm:$false -ErrorAction Stop
                $removed = $true
                Write-DebugMessage "Berechtigung erfolgreich entfernt (DE)" -Type "Success"
            }
            else {
                Write-DebugMessage "Keine Berechtigung gefunden für deutschen Kalender" -Type "Info"
            }
        } 
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Entfernen der deutschen Kalenderberechtigungen: $errorMsg" -Type "Warning"
            # Bei Fehler einfach weitermachen und englischen Pfad versuchen
        }
        
        if (-not $removed) {
            try {
                $identityEN = "${SourceUser}:\Calendar"
                Write-DebugMessage "Prüfe englische Kalenderberechtigungen: $identityEN" -Type "Info"
                
                # Prüfe ob Berechtigung existiert
                $existingPerm = Get-MailboxFolderPermission -Identity $identityEN -User $TargetUser -ErrorAction SilentlyContinue
                
                if ($existingPerm) {
                    Write-DebugMessage "Gefundene Berechtigung wird entfernt (EN): $($existingPerm.AccessRights)" -Type "Info"
                    Remove-MailboxFolderPermission -Identity $identityEN -User $TargetUser -Confirm:$false -ErrorAction Stop
                    $removed = $true
                    Write-DebugMessage "Berechtigung erfolgreich entfernt (EN)" -Type "Success"
                }
                else {
                    Write-DebugMessage "Keine Berechtigung gefunden für englischen Kalender" -Type "Info"
                }
            } 
            catch {
                if (-not $removed) {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Entfernen der englischen Kalenderberechtigungen: $errorMsg" -Type "Error"
                    throw "Fehler beim Entfernen der Kalenderberechtigung: $errorMsg"
                }
            }
        }
        
        if ($removed) {
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Kalenderberechtigung entfernt." -Color $script:connectedBrush
            }
            
            Log-Action "Kalenderberechtigung entfernt: $SourceUser -> $TargetUser"
            return $true
        } 
        else {
            Write-DebugMessage "Keine Kalenderberechtigung zum Entfernen gefunden" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine Kalenderberechtigung gefunden zum Entfernen."
            }
            
            Log-Action "Keine Kalenderberechtigung gefunden zum Entfernen: $SourceUser -> $TargetUser"
            return $false
        }
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen der Kalenderberechtigung: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Entfernen der Kalenderberechtigung: $errorMsg"
        return $false
    }
}

# -------------------------------------------------
# Abschnitt: Postfachberechtigungen
# -------------------------------------------------
function Add-MailboxPermissionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUser,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Validate-Email -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage "Füge Postfachberechtigung hinzu: $SourceUser -> $TargetUser (FullAccess)" -Type "Info"
        
        # Prüfen, ob die Berechtigung bereits existiert
        $existingPermissions = Get-MailboxPermission -Identity $SourceUser -User $TargetUser -ErrorAction SilentlyContinue
        $fullAccessExists = $existingPermissions | Where-Object { $_.AccessRights -like "*FullAccess*" }
        
        if ($fullAccessExists) {
            Write-DebugMessage "Berechtigung existiert bereits, keine Änderung notwendig" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Postfachberechtigung bereits vorhanden." -Color $script:connectedBrush
            }
            Log-Action "Postfachberechtigung bereits vorhanden: $SourceUser -> $TargetUser"
            return $true
        }
        
        # Berechtigung hinzufügen
        Add-MailboxPermission -Identity $SourceUser -User $TargetUser -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop
        
        Write-DebugMessage "Postfachberechtigung erfolgreich hinzugefügt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Postfachberechtigung hinzugefügt." -Color $script:connectedBrush
        }
        Log-Action "Postfachberechtigung hinzugefügt: $SourceUser -> $TargetUser (FullAccess)"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen der Postfachberechtigung: $errorMsg" -Type "Error"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        Log-Action "Fehler beim Hinzufügen der Postfachberechtigung: $errorMsg"
        return $false
    }
}

function Remove-MailboxPermissionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUser,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Validate-Email -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage "Entferne Postfachberechtigung: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfen, ob die Berechtigung existiert
        $existingPermissions = Get-MailboxPermission -Identity $SourceUser -User $TargetUser -ErrorAction SilentlyContinue
        if (-not $existingPermissions) {
            Write-DebugMessage "Keine Berechtigung zum Entfernen gefunden" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine Postfachberechtigung zum Entfernen gefunden."
            }
            Log-Action "Keine Postfachberechtigung zum Entfernen gefunden: $SourceUser -> $TargetUser"
            return $false
        }
        
        # Berechtigung entfernen
        Remove-MailboxPermission -Identity $SourceUser -User $TargetUser -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
        
        Write-DebugMessage "Postfachberechtigung erfolgreich entfernt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Postfachberechtigung entfernt."
        }
        Log-Action "Postfachberechtigung entfernt: $SourceUser -> $TargetUser"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen der Postfachberechtigung: $errorMsg" -Type "Error"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        Log-Action "Fehler beim Entfernen der Postfachberechtigung: $errorMsg"
        return $false
    }
}

function Get-MailboxPermissionsAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxUser
    )
    
    try {
        Write-DebugMessage "Postfachberechtigungen abrufen: Validiere Benutzereingabe" -Type "Info"
        
        if ([string]::IsNullOrEmpty($MailboxUser)) {
            Write-DebugMessage "Keine gültige E-Mail-Adresse angegeben" -Type "Error"
            return $null
        }
        
        Write-DebugMessage "Postfachberechtigungen abrufen für: $MailboxUser" -Type "Info"
        Write-DebugMessage "Rufe Postfachberechtigungen ab für: $MailboxUser" -Type "Info"
        
        # Postfachberechtigungen abrufen
        $mailboxPermissions = Get-MailboxPermission -Identity $MailboxUser | 
            Where-Object { $_.User -notlike "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false } | 
            Select-Object Identity, User, AccessRights, IsInherited, Deny
        
        # SendAs-Berechtigungen abrufen
        $sendAsPermissions = Get-RecipientPermission -Identity $MailboxUser | 
            Where-Object { $_.Trustee -notlike "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false } |
            Select-Object @{Name="Identity"; Expression={$_.Identity}}, 
                        @{Name="User"; Expression={$_.Trustee}}, 
                        @{Name="AccessRights"; Expression={"SendAs"}}, 
                        @{Name="IsInherited"; Expression={$_.IsInherited}},
                        @{Name="Deny"; Expression={$false}}
        
        # Ergebnisse in eine Liste konvertieren und zusammenführen
        $allPermissions = @()
        
        if ($mailboxPermissions) {
            foreach ($perm in $mailboxPermissions) {
                $permObj = [PSCustomObject]@{
                    Identity = $perm.Identity
                    User = $perm.User
                    AccessRights = $perm.AccessRights -join ", "
                    IsInherited = $perm.IsInherited
                    Deny = $perm.Deny
                }
                $allPermissions += $permObj
                Write-DebugMessage "Postfachberechtigung verarbeitet: $($perm.User) -> $($perm.AccessRights)" -Type "Info"
            }
        }
        
        if ($sendAsPermissions) {
            foreach ($perm in $sendAsPermissions) {
                $permObj = [PSCustomObject]@{
                    Identity = $perm.Identity
                    User = $perm.User
                    AccessRights = $perm.AccessRights
                    IsInherited = $perm.IsInherited
                    Deny = $perm.Deny
                }
                $allPermissions += $permObj
                Write-DebugMessage "SendAs-Berechtigung verarbeitet: $($perm.User) -> SendAs" -Type "Info"
            }
        }
        
        $count = $allPermissions.Count
        Write-DebugMessage "Postfachberechtigungen abgerufen und verarbeitet: $count Einträge gefunden" -Type "Success"
        
        return $allPermissions
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Postfachberechtigungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Postfachberechtigungen für $MailboxUser`: $errorMsg"
        return $null
    }
}

function Get-MailboxPermissions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Mailbox
    )
    
    try {
        Write-DebugMessage "Postfachberechtigungen abrufen: Validiere Benutzereingabe" -Type "Info"
        
        # E-Mail-Format überprüfen
        if (-not ($Mailbox -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")) {
            if (-not ($Mailbox -match "^[a-zA-Z0-9\s.-]+$")) {
                throw "Ungültige E-Mail-Adresse oder Benutzername: $Mailbox"
            }
        }
        
        Write-DebugMessage "Postfachberechtigungen abrufen für: $Mailbox" -Type "Info"
        
        # Postfachberechtigungen abrufen
        Write-DebugMessage "Rufe Postfachberechtigungen ab für: $Mailbox" -Type "Info"
        $permissions = Get-MailboxPermission -Identity $Mailbox | Where-Object { 
            $_.User -notlike "NT AUTHORITY\SELF" -and 
            $_.User -notlike "S-1-5*" -and 
            $_.User -notlike "NT AUTHORITY\SYSTEM" -and
            $_.IsInherited -eq $false 
        }
        
        # SendAs-Berechtigungen abrufen
        $sendAsPermissions = Get-RecipientPermission -Identity $Mailbox | Where-Object { 
            $_.Trustee -notlike "NT AUTHORITY\SELF" -and 
            $_.Trustee -notlike "S-1-5*" -and
            $_.Trustee -notlike "NT AUTHORITY\SYSTEM" 
        }
        
        # Ergebnissammlung vorbereiten
        $resultCollection = @()
        
        # Postfachberechtigungen verarbeiten
        foreach ($perm in $permissions) {
            $hasSendAs = $false
            $sendAsEntry = $sendAsPermissions | Where-Object { $_.Trustee -eq $perm.User }
            
            if ($null -ne $sendAsEntry) {
                $hasSendAs = $true
            }
            
            $entry = [PSCustomObject]@{
                Identity = $Mailbox
                User = $perm.User.ToString()
                AccessRights = $perm.AccessRights -join ", "
            }
            
            Write-DebugMessage "Postfachberechtigung verarbeitet: $($perm.User) -> $($perm.AccessRights -join ', ')" -Type "Info"
            $resultCollection += $entry
        }
        
        # SendAs-Berechtigungen hinzufügen, die nicht bereits in Postfachberechtigungen enthalten sind
        foreach ($sendPerm in $sendAsPermissions) {
            $existingEntry = $resultCollection | Where-Object { $_.User -eq $sendPerm.Trustee }
            
            if ($null -eq $existingEntry) {
                $entry = [PSCustomObject]@{
                    Identity = $Mailbox
                    User = $sendPerm.Trustee.ToString()
                    AccessRights = "SendAs"
                }
                $resultCollection += $entry
                Write-DebugMessage "Separate SendAs-Berechtigung verarbeitet: $($sendPerm.Trustee)" -Type "Info"
            }
        }
        
        # Wenn keine Berechtigungen gefunden wurden
        if ($resultCollection.Count -eq 0) {
            # Füge "NT AUTHORITY\SELF" hinzu, der normalerweise vorhanden ist
            $selfPerm = Get-MailboxPermission -Identity $Mailbox | Where-Object { $_.User -like "NT AUTHORITY\SELF" } | Select-Object -First 1
            
            if ($null -ne $selfPerm) {
                $entry = [PSCustomObject]@{
                    Identity = $Mailbox
                    User = "Keine benutzerdefinierten Berechtigungen"
                    AccessRights = "Nur Standardberechtigungen"
                }
                $resultCollection += $entry
                Write-DebugMessage "Keine benutzerdefinierten Berechtigungen gefunden, nur Standardzugriff" -Type "Info"
            }
            else {
                $entry = [PSCustomObject]@{
                    Identity = $Mailbox
                    User = "Keine Berechtigungen gefunden"
                    AccessRights = "Unbekannt"
                }
                $resultCollection += $entry
                Write-DebugMessage "Keine Berechtigungen gefunden" -Type "Warning"
            }
        }
        
        Write-DebugMessage "Postfachberechtigungen abgerufen und verarbeitet: $($resultCollection.Count) Einträge gefunden" -Type "Success"
        
        # Wichtig: Rückgabe als Array für die GUI-Darstellung
        return ,$resultCollection
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Postfachberechtigungen: $errorMsg" -Type "Error"
        return @()
    }
}

# -------------------------------------------------
# Abschnitt: Verwaltung der Standard und Anonym Berechtigungen
# -------------------------------------------------
function Set-DefaultCalendarPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxUser,
        
        [Parameter(Mandatory = $true)]
        [string]$AccessRights
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $MailboxUser)) {
            throw "Ungültige E-Mail-Adresse für Postfach."
        }
        
        Write-DebugMessage "Setze Standard-Kalenderberechtigungen für: $MailboxUser auf: $AccessRights" -Type "Info"
        
        # Prüfe deutsche und englische Kalenderordner
        $identityDE = "${MailboxUser}:\Kalender"
        $identityEN = "${MailboxUser}:\Calendar"
        $identity = $null
        
        # Prüfe, welcher Pfad existiert
        try {
            if (Get-MailboxFolderPermission -Identity $identityDE -User Default -ErrorAction SilentlyContinue) {
                $identity = $identityDE
                Write-DebugMessage "Deutscher Kalenderpfad gefunden: $identity" -Type "Info"
            } else {
                $identity = $identityEN
                Write-DebugMessage "Englischer Kalenderpfad wird verwendet: $identity" -Type "Info"
            }
        } catch {
            $identity = $identityEN
            Write-DebugMessage "Fehler beim Prüfen des deutschen Pfads, verwende englischen Pfad: $identity" -Type "Warning"
        }
        
        # Standard-Berechtigungen setzen
        Write-DebugMessage "Aktualisiere Standard-Berechtigungen für: $identity" -Type "Info"
        Set-MailboxFolderPermission -Identity $identity -User Default -AccessRights $AccessRights -ErrorAction Stop
        
        Write-DebugMessage "Standard-Kalenderberechtigungen erfolgreich gesetzt" -Type "Success"
        Log-Action "Standard-Kalenderberechtigungen für $MailboxUser auf $AccessRights gesetzt"
        return $true
    } catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Setzen der Standard-Kalenderberechtigungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Setzen der Standard-Kalenderberechtigungen: $errorMsg"
        throw $errorMsg
    }
}

function Set-AnonymousCalendarPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxUser,
        
        [Parameter(Mandatory = $true)]
        [string]$AccessRights
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $MailboxUser)) {
            throw "Ungültige E-Mail-Adresse für Postfach."
        }
        
        Write-DebugMessage "Setze Anonym-Kalenderberechtigungen für: $MailboxUser auf: $AccessRights" -Type "Info"
        
        # Prüfe deutsche und englische Kalenderordner
        $identityDE = "${MailboxUser}:\Kalender"
        $identityEN = "${MailboxUser}:\Calendar"
        $identity = $null
        
        # Prüfe, welcher Pfad existiert
        try {
            if (Get-MailboxFolderPermission -Identity $identityDE -User Anonymous -ErrorAction SilentlyContinue) {
                $identity = $identityDE
                Write-DebugMessage "Deutscher Kalenderpfad gefunden: $identity" -Type "Info"
            } else {
                $identity = $identityEN
                Write-DebugMessage "Englischer Kalenderpfad wird verwendet: $identity" -Type "Info"
            }
        } catch {
            $identity = $identityEN
            Write-DebugMessage "Fehler beim Prüfen des deutschen Pfads, verwende englischen Pfad: $identity" -Type "Warning"
        }
        
        # Anonym-Berechtigungen setzen
        Write-DebugMessage "Aktualisiere Anonymous-Berechtigungen für: $identity" -Type "Info"
        Set-MailboxFolderPermission -Identity $identity -User Anonymous -AccessRights $AccessRights -ErrorAction Stop
        
        Write-DebugMessage "Anonymous-Kalenderberechtigungen erfolgreich gesetzt" -Type "Success"
        Log-Action "Anonymous-Kalenderberechtigungen für $MailboxUser auf $AccessRights gesetzt"
        return $true
    } catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Setzen der Anonymous-Kalenderberechtigungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Setzen der Anonymous-Kalenderberechtigungen: $errorMsg"
        throw $errorMsg
    }
}

# -------------------------------------------------
# Abschnitt: Standard und Anonym Berechtigungen für alle Postfächer
# -------------------------------------------------
function Set-DefaultCalendarPermissionForAll {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessRights
    )
    
    try {
        Write-DebugMessage "Setze Standard-Kalenderberechtigungen für alle Postfächer auf: $AccessRights" -Type "Info"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Setze Standard-Kalenderberechtigungen für alle Postfächer..."
        }
        
        # Alle Mailboxen abrufen
        Write-DebugMessage "Rufe alle Mailboxen ab" -Type "Info"
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
        $totalCount = $mailboxes.Count
        $successCount = 0
        $errorCount = 0
        
        Write-DebugMessage "$totalCount Mailboxen gefunden" -Type "Info"
        
        # Fortschrittsanzeige vorbereiten
        $progressIndex = 0
        
        foreach ($mailbox in $mailboxes) {
            $progressIndex++
            $progressPercentage = [math]::Round(($progressIndex / $totalCount) * 100)
            
            try {
                # Status aktualisieren
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Setze Standard-Kalenderberechtigungen ($progressIndex von $totalCount, $progressPercentage%)..."
                }
                
                # Berechtigungen für dieses Postfach setzen
                $mailboxAddress = $mailbox.PrimarySmtpAddress.ToString()
                Write-DebugMessage "Bearbeite Postfach $progressIndex/$totalCount - $mailboxAddress" -Type "Info"
                
                Set-DefaultCalendarPermission -MailboxUser $mailboxAddress -AccessRights $AccessRights
                $successCount++
                Write-DebugMessage "Standard-Kalenderberechtigungen erfolgreich für $mailboxAddress gesetzt" -Type "Success"
            }
            catch {
                $errorCount++
                $errorMsg = $_.Exception.Message
                Write-DebugMessage "Fehler bei Postfach $mailboxAddress - $errorMsg" -Type "Error"
                Log-Action "Fehler beim Setzen der Standard-Kalenderberechtigungen für $mailboxAddress`: $errorMsg"
            }
        }
        
        $statusMessage = "Standard-Kalenderberechtigungen für alle Postfächer gesetzt. Erfolgreich - $successCount, Fehler: $errorCount"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message $statusMessage -Color $script:connectedBrush
        }
        
        Write-DebugMessage $statusMessage -Type "Success"
        Log-Action $statusMessage
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Setzen der Standard-Kalenderberechtigungen für alle - $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Setzen der Standard-Kalenderberechtigungen für alle: $errorMsg"
        return $false
    }
}

function Set-AnonymousCalendarPermissionForAll {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessRights
    )
    
    try {
        Write-DebugMessage "Setze Anonym-Kalenderberechtigungen für alle Postfächer auf: $AccessRights" -Type "Info"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Setze Anonym-Kalenderberechtigungen für alle Postfächer..."
        }
        
        # Alle Mailboxen abrufen
        Write-DebugMessage "Rufe alle Mailboxen ab" -Type "Info"
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
        $totalCount = $mailboxes.Count
        $successCount = 0
        $errorCount = 0
        
        Write-DebugMessage "$totalCount Mailboxen gefunden" -Type "Info"
        
        # Fortschrittsanzeige vorbereiten
        $progressIndex = 0
        
        foreach ($mailbox in $mailboxes) {
            $progressIndex++
            $progressPercentage = [math]::Round(($progressIndex / $totalCount) * 100)
            
            try {
                # Status aktualisieren
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Setze Anonym-Kalenderberechtigungen ($progressIndex von $totalCount, $progressPercentage%)..."
                }
                
                # Berechtigungen für dieses Postfach setzen
                $mailboxAddress = $mailbox.PrimarySmtpAddress.ToString()
                Write-DebugMessage "Bearbeite Postfach $progressIndex/$totalCount - $mailboxAddress" -Type "Info"
                
                Set-AnonymousCalendarPermission -MailboxUser $mailboxAddress -AccessRights $AccessRights
                $successCount++
                Write-DebugMessage "Anonym-Kalenderberechtigungen erfolgreich für $mailboxAddress gesetzt" -Type "Success"
            }
            catch {
                $errorCount++
                $errorMsg = $_.Exception.Message
                Write-DebugMessage "Fehler bei Postfach $mailboxAddress - $errorMsg" -Type "Error"
                Log-Action "Fehler beim Setzen der Anonym-Kalenderberechtigungen für $mailboxAddress`: $errorMsg"
            }
        }
        
        $statusMessage = "Anonym-Kalenderberechtigungen für alle Postfächer gesetzt. Erfolgreich - $successCount, Fehler: $errorCount"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message $statusMessage -Color $script:connectedBrush
        }
        
        Write-DebugMessage $statusMessage -Type "Success"
        Log-Action $statusMessage
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Setzen der Anonym-Kalenderberechtigungen für alle - $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Setzen der Anonym-Kalenderberechtigungen für alle: $errorMsg"
        return $false
    }
}
# -------------------------------------------------
# Abschnitt: Neue Hilfsfunktion für Throttling-Informationen
# -------------------------------------------------
function Get-ExchangeThrottlingInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$InfoType = "General"
    )
    
    try {
        Write-DebugMessage "Rufe Exchange Throttling Informationen ab: $InfoType" -Type "Info"
        
        # Modern Exchange Online hat kein Get-ThrottlingPolicy, aber wir können alternative Informationen sammeln
        $result = @"
## Exchange Online Throttling Informationen

Hinweis: Der Befehl 'Get-ThrottlingPolicy' ist in modernen Exchange Online PowerShell-Verbindungen nicht mehr verfügbar.
Es wurden alternative Informationen zusammengestellt:

"@
        # Je nach gewünschtem Informationstyp unterschiedliche Alternativen anbieten
        switch ($InfoType) {
            "EWSPolicy" {
                $result += @"
### EWS Throttling Informationen

Die EWS-Throttling-Einstellungen können nicht direkt abgefragt werden, jedoch gelten folgende Standardlimits:
- EWSMaxConcurrency: 27 Anfragen/Benutzer
- EWSPercentTimeInAD: 75 (Prozentsatz der Zeit, die EWS in AD verbringen kann)
- EWSPercentTimeInCAS: 150 (Prozentsatz der Zeit, die EWS mit Client Access Services verbringen kann)
- EWSMaxSubscriptions: 20 Ereignisabonnements/Benutzer
- EWSFastSearchTimeoutInSeconds: 60 (Suchzeitlimit)

Weitere Informationen: https://learn.microsoft.com/de-de/exchange/clients-and-mobile-in-exchange-online/exchange-web-services/ews-throttling-in-exchange-online

Um die Werte für einen bestimmten Benutzer zu überprüfen, können Sie einen Test-Fall erstellen und das Verhalten analysieren.
"@
            }
            "PowerShell" {
                # Sammeln von verfügbaren Informationen über Remote PowerShell Limits
                try {
                    $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
                    $result += @"
### PowerShell Throttling Informationen

PowerShell Verbindungslimits aus OrganizationConfig:
- PowerShellMaxConcurrency: $($orgConfig.PowerShellMaxConcurrency)
- PowerShellMaxCmdletQueueDepth: $($orgConfig.PowerShellMaxCmdletQueueDepth)
- PowerShellMaxCmdletsExecutionDuration: $($orgConfig.PowerShellMaxCmdletsExecutionDuration)

Standard Remote PowerShell Limits:
- 3 Remote PowerShell Verbindungen pro Benutzer
- 18 Remote PowerShell Verbindungen pro Mandant
- Timeoutzeit: 15 Minuten Leerlaufzeit

Weitere Informationen: https://learn.microsoft.com/de-de/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps
"@
                }
                catch {
                    $result += "Fehler beim Abrufen von PowerShell-Konfigurationen: $($_.Exception.Message)`n`n"
                    $result += "Dies könnte auf fehlende Berechtigungen oder Verbindungsprobleme hindeuten."
                }
            }
            default {
                # Allgemeine Informationen zu Exchange Online Throttling
                $result += @"
### Allgemeine Throttling Informationen für Exchange Online

Exchange Online implementiert verschiedene Throttling-Mechanismen, um die Dienstverfügbarkeit für alle Benutzer sicherzustellen:

1. **EWS (Exchange Web Services)**
   - Begrenzte gleichzeitige Verbindungen und Anfragen pro Benutzer
   - Für Migrationstools empfohlene Werte: 
     * EWSMaxConcurrency: 20-50
     * EWSMaxConnections: 10-20
     * EWSMaxBatchSize: 500-1000

2. **Remote PowerShell**
   - Standardmäßig 3 Verbindungen pro Benutzer
   - 18 gleichzeitige Verbindungen pro Mandant
   - Timeout nach 15 Minuten Inaktivität

3. **REST APIs**
   - Verschiedene Limits je nach Endpunkt 
   - Microsoft Graph API hat eigene Throttling-Regeln

4. **SMTP**
   - Begrenzte ausgehende Nachrichten pro Tag
   - Begrenzte Empfänger pro Nachricht

Hinweis: Die genauen Throttling-Werte können sich ändern und werden nicht mehr über PowerShell-Cmdlets offengelegt. Bei anhaltenden Problemen mit Throttling wenden Sie sich an den Microsoft Support.

Weitere Informationen:
- https://learn.microsoft.com/de-de/exchange/clients-and-mobile-in-exchange-online/exchange-web-services/ews-throttling-in-exchange-online
- https://learn.microsoft.com/de-de/exchange/client-developer/exchange-web-services/how-to-maintain-affinity-between-group-of-subscriptions-and-mailbox-server
"@
            }
        }

        Write-DebugMessage "Exchange Throttling Information erfolgreich erstellt" -Type "Success"
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Exchange Throttling Informationen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Exchange Throttling Informationen: $errorMsg"
        return "Fehler beim Abrufen der Exchange Throttling Informationen: $errorMsg"
    }
}

# -------------------------------------------------
# Abschnitt: Aktualisierte Throttling Policy Funktionen
# -------------------------------------------------
function Get-ThrottlingPolicyAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$PolicyName = "",
        
        [Parameter(Mandatory = $false)]
        [switch]$ShowEWSOnly,
        
        [Parameter(Mandatory = $false)]
        [switch]$DetailedView
    )
    
    try {
        Write-DebugMessage "Rufe alternative Throttling-Informationen ab" -Type "Info"
        
        if ($ShowEWSOnly) {
            return Get-ExchangeThrottlingInfo -InfoType "EWSPolicy"
        }
        elseif ($DetailedView) {
            return Get-ExchangeThrottlingInfo -InfoType "PowerShell"
        }
        else {
            return Get-ExchangeThrottlingInfo -InfoType "General"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Throttling-Informationen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Throttling-Informationen: $errorMsg"
        return "Fehler beim Abrufen der Throttling-Informationen: $errorMsg"
    }
}

# Funktion zur Unterstützung des Troubleshooting-Bereichs
function Get-ThrottlingPolicyForTroubleshooting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$PolicyType
    )
    
    try {
        Write-DebugMessage "Führe Throttling Policy Troubleshooting aus: $PolicyType" -Type "Info"
        
        switch ($PolicyType) {
            "EWSPolicy" {
                # Spezifisch für EWS Throttling (für Migrationstools)
                return Get-ExchangeThrottlingInfo -InfoType "EWSPolicy"
            }
            "PowerShell" {
                # Für Remote PowerShell Throttling
                return Get-ExchangeThrottlingInfo -InfoType "PowerShell"
            }
            "All" {
                # Zeige alle Policies in der Übersicht
                return Get-ExchangeThrottlingInfo -InfoType "General"
            }
            default {
                # Standard-Ansicht
                return Get-ExchangeThrottlingInfo -InfoType "General"
            }
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Throttling Policy Troubleshooting: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Throttling Policy Troubleshooting: $errorMsg"
        return "Fehler beim Abrufen der Throttling Policy Informationen: $errorMsg"
    }
}

# Erweitere die Diagnostics-Funktionen um einen speziellen Throttling-Test
function Test-EWSThrottlingPolicy {
    [CmdletBinding()]
    param()
    
    try {
        Write-DebugMessage "Prüfe EWS Throttling Policy für Migration" -Type "Info"
        
        # EWS Policy Informationen abrufen
        $ewsPolicy = Get-ExchangeThrottlingInfo -InfoType "EWSPolicy"
        
        # Ergebnis formatieren mit Empfehlungen
        $result = $ewsPolicy
        $result += "`n`nZusätzliche Empfehlungen für Migrationen:`n"
        $result += "- Verwenden Sie für große Migrationen mehrere Benutzerkonten, um die Throttling-Limits zu umgehen\n"
        $result += "- Planen Sie Migrationen zu Zeiten geringerer Exchange-Auslastung\n"
        $result += "- Implementieren Sie bei Throttling-Fehlern eine exponentielle Backoff-Strategie\n\n"
        
        $result += "Hinweis: Microsoft passt Throttling-Limits regelmäßig an, ohne dies zu dokumentieren.\n"
        $result += "Bei anhaltenden Problemen wenden Sie sich an den Microsoft Support."
        
        Write-DebugMessage "EWS Throttling Policy Test abgeschlossen" -Type "Success"
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Testen der EWS Throttling Policy: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Testen der EWS Throttling Policy: $errorMsg"
        return "Fehler beim Testen der EWS Throttling Policy: $errorMsg"
    }
}

# -------------------------------------------------
# Abschnitt: Exchange Online Troubleshooting Diagnostics
# -------------------------------------------------

# Aktualisierte Diagnostics-Datenstruktur
$script:exchangeDiagnostics = @(
    @{
        Name = "Migration EWS Throttling Policy"
        Description = "Informationen über EWS-Throttling-Einstellungen für Mailbox-Migrationen (auch für Drittanbieter-Tools relevant)."
        PowerShellCheck = "Get-ExchangeThrottlingInfo -InfoType 'EWSPolicy'"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/settings/services"
        Tooltip = "Öffnen Sie die Exchange Admin Center-Einstellungen"
    },
    @{
        Name = "Exchange Online Accepted Domain diagnostics"
        Description = "Prüfen Sie, ob eine Domain korrekt als akzeptierte Domain in Exchange Online konfiguriert ist."
        PowerShellCheck = "Get-AcceptedDomain"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/accepted-domains"
        Tooltip = "Überprüfen Sie die Konfiguration akzeptierter Domains"
    },
    @{
        Name = "Test a user's Exchange Online RBAC permissions"
        Description = "Überprüfen Sie, ob ein Benutzer die erforderlichen RBAC-Rollen besitzt, um bestimmte Exchange Online-Cmdlets auszuführen."
        PowerShellCheck = "Get-ManagementRoleAssignment –RoleAssignee '[USER]'"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/permissions"
        Tooltip = "Öffnen Sie die RBAC-Einstellungen zur Verwaltung von Benutzerberechtigungen"
        RequiresUser = $true
    },
    @{
        Name = "Compare EXO RBAC Permissions for Two Users"
        Description = "Vergleichen Sie die RBAC-Rollen zweier Benutzer, um Unterschiede zu identifizieren, wenn ein Benutzer Cmdlet-Fehler erhält."
        PowerShellCheck = "Compare-Object (Get-ManagementRoleAssignment –RoleAssignee '[USER1]') (Get-ManagementRoleAssignment –RoleAssignee '[USER2]')"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/permissions"
        Tooltip = "Überprüfen und vergleichen Sie RBAC-Zuweisungen zur Fehlerbehebung"
        RequiresTwoUsers = $true
    },
    @{
        Name = "Recipient failure"
        Description = "Überprüfen Sie den Status und die Konfiguration eines Exchange Online-Empfängers, um Bereitstellungs- oder Synchronisierungsprobleme zu beheben."
        PowerShellCheck = "Get-EXORecipient –Identity '[USER]'"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/recipients"
        Tooltip = "Überprüfen Sie die Empfängerkonfiguration und den Bereitstellungsstatus"
        RequiresUser = $true
    },
    @{
        Name = "Exchange Organization Object check"
        Description = "Diagnostizieren Sie Probleme mit dem Exchange Online-Organisationsobjekt, wie Mandantenbereitstellung oder RBAC-Fehlkonfigurationen."
        PowerShellCheck = "Get-OrganizationConfig | Format-List"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/organization"
        Tooltip = "Überprüfen Sie die Organisationskonfigurationseinstellungen"
    },
    @{
        Name = "Mailbox or message size"
        Description = "Überprüfen Sie die Postfachgröße und Nachrichtengröße (einschließlich Anhänge), um Speicherprobleme zu identifizieren."
        PowerShellCheck = "Get-EXOMailboxStatistics –Identity '[USER]' | Select-Object DisplayName, TotalItemSize, ItemCount"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/mailboxes"
        Tooltip = "Überprüfen Sie Postfachgröße und Speicherstatistiken"
        RequiresUser = $true
    },
    @{
        Name = "Deleted mailbox diagnostics"
        Description = "Überprüfen Sie den Status kürzlich gelöschter (soft-deleted) Postfächer zur Wiederherstellung oder Bereinigung."
        PowerShellCheck = "Get-Mailbox –SoftDeletedMailbox"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/deletedmailboxes"
        Tooltip = "Verwalten Sie gelöschte Postfächer"
    },
    @{
        Name = "Exchange Remote PowerShell throttling information"
        Description = "Bewerten Sie Remote PowerShell Throttling-Einstellungen, um Verbindungsprobleme zu minimieren."
        PowerShellCheck = "Get-ExchangeThrottlingInfo -InfoType 'PowerShell'"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/settings"
        Tooltip = "Überprüfen Sie Remote PowerShell Throttling-Einstellungen"
    },
    @{
        Name = "Email delivery troubleshooter"
        Description = "Diagnostizieren Sie E-Mail-Zustellungsprobleme durch Nachverfolgung von Nachrichtenpfaden und Identifizierung von Fehlern."
        PowerShellCheck = "Get-MessageTrace –StartDate (Get-Date).AddDays(-7) –EndDate (Get-Date)"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/mailflow"
        Tooltip = "Überprüfen Sie Mail-Flow und Zustellungsprobleme"
    },
    @{
        Name = "Archive mailbox diagnostics"
        Description = "Überprüfen Sie die Konfiguration und den Status von Archiv-Postfächern, um sicherzustellen, dass die Archivierung aktiviert ist und funktioniert."
        PowerShellCheck = "Get-EXOMailbox –Identity '[USER]' | Select-Object DisplayName, ArchiveStatus"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/archivemailboxes"
        Tooltip = "Überprüfen Sie die Konfiguration des Archivpostfachs"
        RequiresUser = $true
    },
    @{
        Name = "Retention policy diagnostics for a user mailbox"
        Description = "Überprüfen Sie die Aufbewahrungsrichtlinieneinstellungen (einschließlich Tags und Richtlinien) für ein Benutzerpostfach, um die Einhaltung der organisatorischen Richtlinien zu gewährleisten."
        PowerShellCheck = "Get-EXOMailbox –Identity '[USER]' | Select-Object DisplayName, RetentionPolicy; Get-RetentionPolicy"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/compliance"
        Tooltip = "Überprüfen Sie die Einstellungen der Aufbewahrungsrichtlinie"
        RequiresUser = $true
    },
    @{
        Name = "DomainKeys Identified Mail (DKIM) diagnostics"
        Description = "Überprüfen Sie, ob die DKIM-Signierung korrekt konfiguriert ist und die richtigen DNS-Einträge veröffentlicht wurden."
        PowerShellCheck = "Get-DkimSigningConfig"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/dkim"
        Tooltip = "Verwalten Sie DKIM-Einstellungen und überprüfen Sie die DNS-Konfiguration"
    },
    @{
        Name = "Proxy address conflict diagnostics"
        Description = "Identifizieren Sie den Exchange-Empfänger, der eine bestimmte Proxy-Adresse (E-Mail-Adresse) verwendet, die Konflikte verursacht, z.B. Fehler bei der Postfacherstellung."
        PowerShellCheck = "Get-Recipient –Filter {EmailAddresses -like '[EMAIL]'}"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/recipients"
        Tooltip = "Identifizieren und beheben Sie Proxy-Adresskonflikte"
        RequiresEmail = $true
    },
    @{
        Name = "Mailbox safe/blocked sender list diagnostics"
        Description = "Überprüfen Sie sichere Absender und blockierte Absender/Domains in den Junk-E-Mail-Einstellungen eines Postfachs, um potenzielle Zustellungsprobleme zu beheben."
        PowerShellCheck = "Get-MailboxJunkEmailConfiguration –Identity '[USER]' | Select-Object SafeSenders, BlockedSenders"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/mailboxes"
        Tooltip = "Überprüfen Sie Listen sicherer und blockierter Absender"
        RequiresUser = $true
    }
)

# Verbesserte Run-ExchangeDiagnostic Funktion mit noch robusterer Fehlerbehandlung
function Run-ExchangeDiagnostic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$DiagnosticIndex,
        
        [Parameter(Mandatory = $false)]
        [string]$User = "",
        
        [Parameter(Mandatory = $false)]
        [string]$User2 = "",
        
        [Parameter(Mandatory = $false)]
        [string]$Email = ""
    )
    
    try {
        if (-not $script:isConnected) {
            throw "Sie müssen zuerst mit Exchange Online verbunden sein."
        }
        
        $diagnostic = $script:exchangeDiagnostics[$DiagnosticIndex]
        Write-DebugMessage "Führe Diagnose aus: $($diagnostic.Name)" -Type "Info"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Diagnose wird ausgeführt: $($diagnostic.Name)..."
        }
        
        # Befehl vorbereiten
        $command = $diagnostic.PowerShellCheck
        
        # Platzhalter ersetzen
        if ($diagnostic.RequiresUser -and -not [string]::IsNullOrEmpty($User)) {
            $command = $command -replace '\[USER\]', $User
        } elseif ($diagnostic.RequiresUser -and [string]::IsNullOrEmpty($User)) {
            throw "Diese Diagnose erfordert einen Benutzernamen."
        }
        
        if ($diagnostic.RequiresTwoUsers) {
            if ([string]::IsNullOrEmpty($User) -or [string]::IsNullOrEmpty($User2)) {
                throw "Diese Diagnose erfordert zwei Benutzernamen."
            }
            $command = $command -replace '\[USER1\]', $User
            $command = $command -replace '\[USER2\]', $User2
        }
        
        if ($diagnostic.RequiresEmail) {
            if ([string]::IsNullOrEmpty($Email)) {
                throw "Diese Diagnose erfordert eine E-Mail-Adresse."
            }
            $command = $command -replace '\[EMAIL\]', $Email
        }
        
        # Befehl ausführen
        Write-DebugMessage "Führe PowerShell-Befehl aus: $command" -Type "Info"
        
        # Create ScriptBlock from command string and execute
        try {
            $scriptBlock = [Scriptblock]::Create($command)
            $result = & $scriptBlock | Out-String
            
            # Behandlung für leere Ergebnisse
            if ([string]::IsNullOrWhiteSpace($result)) {
                $result = "Die Abfrage wurde erfolgreich ausgeführt, lieferte aber keine Ergebnisse. Dies kann bedeuten, dass keine Daten vorhanden sind oder dass ein Filter keine Übereinstimmungen fand."
            }
        }
        catch {
            # Spezielle Behandlung für bekannte Fehler
            if ($_.Exception.Message -like "*Get-ThrottlingPolicy*") {
                Write-DebugMessage "Get-ThrottlingPolicy ist nicht verfügbar, verwende alternative Informationsquellen" -Type "Warning"
                $result = Get-ExchangeThrottlingInfo -InfoType $(if ($command -like "*EWS*") { "EWSPolicy" } elseif ($command -like "*PowerShell*") { "PowerShell" } else { "General" })
            }
            elseif ($_.Exception.Message -like "*not recognized as the name of a cmdlet*") {
                Write-DebugMessage "Cmdlet wird nicht erkannt: $($_.Exception.Message)" -Type "Warning"
                
                # Spezifische Behandlung für bekannte alte Cmdlets und deren Ersatz
                if ($command -like "*Get-EXORecipient*") {
                    Write-DebugMessage "Versuche Get-Recipient als Alternative zu Get-EXORecipient" -Type "Info"
                    $alternativeCommand = $command -replace "Get-EXORecipient", "Get-Recipient"
                    try {
                        $scriptBlock = [Scriptblock]::Create($alternativeCommand)
                        $result = & $scriptBlock | Out-String
                    } catch {
                        throw "Fehler beim Ausführen des alternativen Befehls: $($_.Exception.Message)"
                    }
                }
                elseif ($command -like "*Get-EXOMailboxStatistics*") {
                    Write-DebugMessage "Versuche Get-MailboxStatistics als Alternative zu Get-EXOMailboxStatistics" -Type "Info"
                    $alternativeCommand = $command -replace "Get-EXOMailboxStatistics", "Get-MailboxStatistics"
                    try {
                        $scriptBlock = [Scriptblock]::Create($alternativeCommand)
                        $result = & $scriptBlock | Out-String
                    } catch {
                        throw "Fehler beim Ausführen des alternativen Befehls: $($_.Exception.Message)"
                    }
                }
                elseif ($command -like "*Get-EXOMailbox*") {
                    Write-DebugMessage "Versuche Get-Mailbox als Alternative zu Get-EXOMailbox" -Type "Info"
                    $alternativeCommand = $command -replace "Get-EXOMailbox", "Get-Mailbox"
                    try {
                        $scriptBlock = [Scriptblock]::Create($alternativeCommand)
                        $result = & $scriptBlock | Out-String
                    } catch {
                        throw "Fehler beim Ausführen des alternativen Befehls: $($_.Exception.Message)"
                    }
                }
                else {
                    throw
                }
            }
            else {
                throw
            }
        }
        
        Log-Action "Exchange-Diagnose ausgeführt: $($diagnostic.Name)"
        Write-DebugMessage "Diagnose abgeschlossen: $($diagnostic.Name)" -Type "Success"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Diagnose abgeschlossen: $($diagnostic.Name)" -Color $script:connectedBrush
        }
        
        # Format und Filter Ausgabe für bessere Lesbarkeit
        if ($result.Length -gt 30000) {
            $result = $result.Substring(0, 30000) + "`n`n... (Output gekürzt - zu viele Daten)`n"
        }
        
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler bei der Diagnose: $errorMsg" -Type "Error"
        Log-Action "Fehler bei der Exchange-Diagnose '$($diagnostic.Name)': $errorMsg"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler bei der Diagnose: $errorMsg"
        }
        
        # Hilfreiche Fehlermeldung für häufige Probleme
        $result = "Fehler bei der Ausführung der Diagnose: $errorMsg`n`n"
        
        if ($errorMsg -like "*is not recognized as the name of a cmdlet*") {
            $result += "Dies könnte folgende Gründe haben:`n"
            $result += "1. Das verwendete Cmdlet ist in modernen Exchange Online PowerShell-Verbindungen nicht mehr verfügbar.`n"
            $result += "2. Sie benötigen möglicherweise zusätzliche Berechtigungen.`n"
            $result += "3. Die Exchange Online Verbindung wurde getrennt oder ist eingeschränkt.`n`n"
            $result += "Empfehlung: Versuchen Sie die Verbindung zu trennen und neu herzustellen oder verwenden Sie die Exchange Admin Center-Website für diese Aufgabe."
        }
        elseif ($errorMsg -like "*Benutzer nicht gefunden*" -or $errorMsg -like "*user not found*") {
            $result += "Der angegebene Benutzer konnte nicht gefunden werden. Mögliche Ursachen:`n"
            $result += "1. Der Benutzername oder die E-Mail-Adresse wurde falsch eingegeben.`n"
            $result += "2. Der Benutzer existiert nicht in der Exchange Online-Umgebung.`n"
            $result += "3. Es kann bis zu 24 Stunden dauern, bis ein neuer Benutzer in allen Exchange-Systemen sichtbar ist.`n`n"
            $result += "Empfehlung: Überprüfen Sie die Schreibweise oder verwenden Sie den vollständigen UPN (user@domain.com)."
        }
        elseif ($errorMsg -like "*insufficient*" -or $errorMsg -like "*not authorized*" -or $errorMsg -like "*access*denied*") {
            $result += "Sie haben nicht die erforderlichen Berechtigungen für diesen Vorgang. Mögliche Lösungen:`n"
            $result += "1. Versuchen Sie, sich mit einem Exchange Admin-Konto anzumelden.`n"
            $result += "2. Lassen Sie sich die erforderlichen RBAC-Rollen zuweisen.`n"
            $result += "3. Verwenden Sie das Exchange Admin Center für diese Aufgabe.`n"
        }
        elseif ($errorMsg -like "*timeout*" -or $errorMsg -like "*Zeitüberschreitung*") {
            $result += "Die Verbindung wurde durch ein Timeout unterbrochen. Mögliche Lösungen:`n"
            $result += "1. Stellen Sie die Verbindung zu Exchange Online neu her.`n"
            $result += "2. Die Abfrage könnte zu umfangreich sein, versuchen Sie einen spezifischeren Filter.`n"
            $result += "3. Exchange Online kann bei hoher Last gedrosselt werden, versuchen Sie es später erneut.`n"
        }
        
        return $result
    }
}

# Fortsetzung des Codes - Vervollständigung der fehlenden Funktionen

function Get-MailboxAuditConfigForUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [int]$InfoType
    )
    
    try {
        # Mailbox-Audit-Konfiguration abrufen
        $mailboxObj = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        
        switch ($InfoType) {
            1 { # Audit-Konfiguration
                $result = "### Audit-Konfiguration für Postfach: $($mailboxObj.DisplayName)`n`n"
                $result += "Audit aktiviert: $($mailboxObj.AuditEnabled)`n"
                $result += "Audit-Administratoraktionen: $($mailboxObj.AuditAdmin -join ', ')`n"
                $result += "Audit-Stellvertreteraktionen: $($mailboxObj.AuditDelegate -join ', ')`n"
                $result += "Audit-Benutzeraktionen: $($mailboxObj.AuditOwner -join ', ')`n"
                $result += "Aufbewahrungszeitraum für Audit-Logs: $($mailboxObj.AuditLogAgeLimit)`n"
                return $result
            }
            2 { # Empfehlungen zur Audit-Konfiguration
                $result = "### Audit-Empfehlungen für Postfach: $($mailboxObj.DisplayName)`n`n"
                
                if (-not $mailboxObj.AuditEnabled) {
                    $result += "⚠️ Audit ist derzeit deaktiviert. Empfehlung: Aktivieren Sie das Audit für dieses Postfach.`n`n"
                }
                
                $recommendedAdminActions = @("Copy", "Create", "FolderBind", "HardDelete", "MessageBind", "Move", "MoveToDeletedItems", "SendAs", "SendOnBehalf", "SoftDelete", "Update")
                $recommendedDelegateActions = @("FolderBind", "SendAs", "SendOnBehalf", "SoftDelete", "HardDelete", "Update", "Move", "MoveToDeletedItems")
                $recommendedOwnerActions = @("HardDelete", "SoftDelete", "Update", "Move", "MoveToDeletedItems")
                
                $missingAdminActions = $recommendedAdminActions | Where-Object { $mailboxObj.AuditAdmin -notcontains $_ }
                $missingDelegateActions = $recommendedDelegateActions | Where-Object { $mailboxObj.AuditDelegate -notcontains $_ }
                $missingOwnerActions = $recommendedOwnerActions | Where-Object { $mailboxObj.AuditOwner -notcontains $_ }
                
                if ($missingAdminActions.Count -gt 0) {
                    $result += "⚠️ Fehlende empfohlene Administrator-Audit-Aktionen: $($missingAdminActions -join ', ')`n"
                }
                
                if ($missingDelegateActions.Count -gt 0) {
                    $result += "⚠️ Fehlende empfohlene Stellvertreter-Audit-Aktionen: $($missingDelegateActions -join ', ')`n"
                }
                
                if ($missingOwnerActions.Count -gt 0) {
                    $result += "⚠️ Fehlende empfohlene Benutzer-Audit-Aktionen: $($missingOwnerActions -join ', ')`n"
                }
                
                if ($mailboxObj.AuditLogAgeLimit -lt "180.00:00:00") {
                    $result += "⚠️ Audit-Aufbewahrungszeitraum ist kürzer als empfohlen. Aktuell: $($mailboxObj.AuditLogAgeLimit), Empfohlen: 180 Tage oder mehr.`n"
                }
                
                return $result
            }
            3 { # Audit-Status prüfen
                $result = "### Audit-Status für Postfach: $($mailboxObj.DisplayName)`n`n"
                
                try {
                    # Versuche, Audit-Logs für die letzte Woche abzurufen
                    $endDate = Get-Date
                    $startDate = $endDate.AddDays(-7)
                    $auditLogs = Search-MailboxAuditLog -Identity $Mailbox -StartDate $startDate -EndDate $endDate -LogonTypes Owner,Delegate,Admin -ResultSize 10 -ErrorAction Stop
                    
                    if ($auditLogs -and $auditLogs.Count -gt 0) {
                        $result += "✅ Audit-Logging funktioniert, $($auditLogs.Count) Einträge in den letzten 7 Tagen gefunden.`n`n"
                        $result += "Die neuesten 10 Audit-Ereignisse:`n"
                        foreach ($log in $auditLogs) {
                            $result += "- $($log.LastAccessed) - $($log.Operation) von $($log.LogonUserDisplayName)`n"
                        }
                    } else {
                        $result += "⚠️ Keine Audit-Logs für die letzten 7 Tage gefunden. Mögliche Ursachen:`n"
                        $result += "1. Das Postfach wurde nicht verwendet`n"
                        $result += "2. Audit ist nicht richtig konfiguriert`n"
                        $result += "3. Die zu auditierenden Aktionen sind nicht konfiguriert`n"
                    }
                } catch {
                    $result += "Fehler beim Abrufen der Audit-Logs: $($_.Exception.Message)`n"
                    $result += "Möglicherweise haben Sie nicht ausreichende Berechtigungen für diese Operation.`n"
                }
                
                return $result
            }
            4 { # Audit-Konfiguration anpassen
                $result = "### Audit-Konfigurationsbefehle für Postfach: $($mailboxObj.DisplayName)`n`n"
                $result += "Um Audit zu aktivieren und alle empfohlenen Einstellungen vorzunehmen, können Sie folgende PowerShell-Befehle verwenden:`n`n"
                $result += "```powershell`n"
                $result += "Set-Mailbox -Identity '$Mailbox' -AuditEnabled `$true`n"
                $result += "Set-Mailbox -Identity '$Mailbox' -AuditAdmin Copy,Create,FolderBind,HardDelete,MessageBind,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update`n"
                $result += "Set-Mailbox -Identity '$Mailbox' -AuditDelegate FolderBind,SendAs,SendOnBehalf,SoftDelete,HardDelete,Update,Move,MoveToDeletedItems`n"
                $result += "Set-Mailbox -Identity '$Mailbox' -AuditOwner HardDelete,SoftDelete,Update,Move,MoveToDeletedItems`n"
                $result += "Set-Mailbox -Identity '$Mailbox' -AuditLogAgeLimit 180`n"
                $result += "```"
                
                return $result
            }
            5 { # Vollständige Audit-Details
                $result = "### Vollständige Audit-Details für Postfach: $($mailboxObj.DisplayName)`n`n"
                
                $result += "AuditEnabled: $($mailboxObj.AuditEnabled)`n"
                $result += "AuditLogAgeLimit: $($mailboxObj.AuditLogAgeLimit)`n`n"
                
                $result += "AuditAdmin: `n"
                if ($mailboxObj.AuditAdmin -and $mailboxObj.AuditAdmin.Count -gt 0) {
                    foreach ($action in $mailboxObj.AuditAdmin) {
                        $result += "- $action`n"
                    }
                } else {
                    $result += "- Keine Aktionen konfiguriert`n"
                }
                
                $result += "`nAuditDelegate: `n"
                if ($mailboxObj.AuditDelegate -and $mailboxObj.AuditDelegate.Count -gt 0) {
                    foreach ($action in $mailboxObj.AuditDelegate) {
                        $result += "- $action`n"
                    }
                } else {
                    $result += "- Keine Aktionen konfiguriert`n"
                }
                
                $result += "`nAuditOwner: `n"
                if ($mailboxObj.AuditOwner -and $mailboxObj.AuditOwner.Count -gt 0) {
                    foreach ($action in $mailboxObj.AuditOwner) {
                        $result += "- $action`n"
                    }
                } else {
                    $result += "- Keine Aktionen konfiguriert`n"
                }
                
                return $result
            }
            default {
                return "Unbekannter Informationstyp: $InfoType"
            }
        }
    }
    catch {
        Write-DebugMessage "Fehler beim Abrufen der Audit-Konfiguration: $($_.Exception.Message)" -Type "Error"
        return "Fehler beim Abrufen der Audit-Konfiguration: $($_.Exception.Message)"
    }
}

function Get-MailboxForwardingForUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [int]$InfoType
    )
    
    try {
        # Mailbox-Objekt und Weiterleitungsregeln abrufen
        $mailboxObj = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        
        switch ($InfoType) {
            1 { # Grundlegende Weiterleitungsinformationen
                $result = "### Weiterleitungseinstellungen für Postfach: $($mailboxObj.DisplayName)`n`n"
                
                # Überprüfe ForwardingAddress und ForwardingSmtpAddress
                if (-not [string]::IsNullOrEmpty($mailboxObj.ForwardingAddress) -or -not [string]::IsNullOrEmpty($mailboxObj.ForwardingSmtpAddress)) {
                    $result += "⚠️ Das Postfach hat eine aktive Weiterleitung konfiguriert!`n`n"
                    
                    if (-not [string]::IsNullOrEmpty($mailboxObj.ForwardingAddress)) {
                        $result += "ForwardingAddress: $($mailboxObj.ForwardingAddress)`n"
                    }
                    
                    if (-not [string]::IsNullOrEmpty($mailboxObj.ForwardingSmtpAddress)) {
                        $result += "ForwardingSmtpAddress: $($mailboxObj.ForwardingSmtpAddress)`n"
                    }
                    
                    $result += "DeliverToMailboxAndForward: $($mailboxObj.DeliverToMailboxAndForward)`n"
                    
                    if ($mailboxObj.DeliverToMailboxAndForward -eq $false) {
                        $result += "⚠️ Nachrichten werden NUR weitergeleitet (keine Kopie im Postfach)`n"
                    } else {
                        $result += "✓ Nachrichten werden weitergeleitet UND im Postfach behalten`n"
                    }
                } else {
                    $result += "✓ Keine direkte Postfach-Weiterleitung konfiguriert.`n"
                }
                
                # Überprüfe Inbox-Regeln
                try {
                    $inboxRules = Get-InboxRule -Mailbox $Mailbox -ErrorAction Stop | Where-Object { 
                        -not [string]::IsNullOrEmpty($_.ForwardTo) -or 
                        -not [string]::IsNullOrEmpty($_.ForwardAsAttachmentTo) -or 
                        -not [string]::IsNullOrEmpty($_.RedirectTo) 
                    }
                    
                    if ($inboxRules -and $inboxRules.Count -gt 0) {
                        $result += "`n⚠️ Das Postfach hat $($inboxRules.Count) Weiterleitungsregeln konfiguriert:`n"
                        
                        foreach ($rule in $inboxRules) {
                            $result += "`n- Regel: $($rule.Name)`n"
                            
                            if ($rule.Enabled -eq $true) {
                                $result += "  Status: Aktiv`n"
                            } else {
                                $result += "  Status: Deaktiviert`n"
                            }
                            
                            if (-not [string]::IsNullOrEmpty($rule.ForwardTo)) {
                                $result += "  ForwardTo: $($rule.ForwardTo -join ', ')`n"
                            }
                            
                            if (-not [string]::IsNullOrEmpty($rule.ForwardAsAttachmentTo)) {
                                $result += "  ForwardAsAttachmentTo: $($rule.ForwardAsAttachmentTo -join ', ')`n"
                            }
                            
                            if (-not [string]::IsNullOrEmpty($rule.RedirectTo)) {
                                $result += "  RedirectTo: $($rule.RedirectTo -join ', ')`n"
                            }
                        }
                    } else {
                        $result += "`n✓ Keine Weiterleitungsregeln im Posteingang gefunden.`n"
                    }
                } catch {
                    $result += "`n⚠️ Fehler beim Prüfen der Inbox-Regeln: $($_.Exception.Message)`n"
                }
                
                return $result
            }
            2 { # Detaillierte Analyse externer Weiterleitungen
                $result = "### Analyse externer Weiterleitungen für: $($mailboxObj.DisplayName)`n`n"
                
                # Liste der internen Domains abrufen
                try {
                    $internalDomains = (Get-AcceptedDomain).DomainName
                    $result += "Interne Domains für E-Mail-Weiterleitungs-Analyse:`n"
                    foreach ($domain in $internalDomains) {
                        $result += "- $domain`n"
                    }
                    $result += "`n"
                } catch {
                    $result += "⚠️ Fehler beim Abrufen der internen Domains: $($_.Exception.Message)`n`n"
                    $internalDomains = @()
                }
                
                # Überprüfe ForwardingSmtpAddress auf externe Weiterleitung
                if (-not [string]::IsNullOrEmpty($mailboxObj.ForwardingSmtpAddress)) {
                    $forwardingDomain = ($mailboxObj.ForwardingSmtpAddress -split "@")[1].TrimEnd(">")
                    $isExternal = $true
                    
                    foreach ($domain in $internalDomains) {
                        if ($forwardingDomain -like "*$domain*") {
                            $isExternal = $false
                            break
                        }
                    }
                    
                    if ($isExternal) {
                        $result += "⚠️ EXTERNE WEITERLEITUNG GEFUNDEN!`n"
                        $result += "ForwardingSmtpAddress: $($mailboxObj.ForwardingSmtpAddress)`n"
                        $result += "Die Weiterleitung erfolgt an eine externe Domain: $forwardingDomain`n`n"
                    } else {
                        $result += "✓ Die konfigurierte Weiterleitung erfolgt an eine interne Domain.`n`n"
                    }
                }
                
                # Überprüfe Inbox-Regeln auf externe Weiterleitungen
                try {
                    $inboxRules = Get-InboxRule -Mailbox $Mailbox -ErrorAction Stop | Where-Object { 
                        -not [string]::IsNullOrEmpty($_.ForwardTo) -or 
                        -not [string]::IsNullOrEmpty($_.ForwardAsAttachmentTo) -or 
                        -not [string]::IsNullOrEmpty($_.RedirectTo) 
                    }
                    
                    if ($inboxRules -and $inboxRules.Count -gt 0) {
                        $externalRules = @()
                        
                        foreach ($rule in $inboxRules) {
                            $destinations = @()
                            if (-not [string]::IsNullOrEmpty($rule.ForwardTo)) { $destinations += $rule.ForwardTo }
                            if (-not [string]::IsNullOrEmpty($rule.ForwardAsAttachmentTo)) { $destinations += $rule.ForwardAsAttachmentTo }
                            if (-not [string]::IsNullOrEmpty($rule.RedirectTo)) { $destinations += $rule.RedirectTo }
                            
                            foreach ($dest in $destinations) {
                                if ($dest -match "SMTP:([^@]+@([^>]+))") {
                                    $email = $matches[1]
                                    $domain = $matches[2]
                                    
                                    $isExternal = $true
                                    foreach ($intDomain in $internalDomains) {
                                        if ($domain -like "*$intDomain*") {
                                            $isExternal = $false
                                            break
                                        }
                                    }
                                    
                                    if ($isExternal) {
                                        $externalRules += [PSCustomObject]@{
                                            RuleName = $rule.Name
                                            Enabled = $rule.Enabled
                                            Destination = $email
                                            Domain = $domain
                                        }
                                    }
                                }
                            }
                        }
                        
                        if ($externalRules.Count -gt 0) {
                            $result += "⚠️ EXTERNE WEITERLEITUNGSREGELN GEFUNDEN!`n"
                            $result += "Es wurden $($externalRules.Count) Regeln mit Weiterleitungen an externe E-Mail-Adressen gefunden:`n`n"
                            
                            foreach ($extRule in $externalRules) {
                                $status = if ($extRule.Enabled) { "Aktiv" } else { "Deaktiviert" }
                                $result += "- Regel: $($extRule.RuleName) ($status)`n"
                                $result += "  Weiterleitung an: $($extRule.Destination)`n"
                                $result += "  Externe Domain: $($extRule.Domain)`n`n"
                            }
                        } else {
                            $result += "✓ Keine externen Weiterleitungsregeln gefunden.`n"
                        }
                    }
                } catch {
                    $result += "⚠️ Fehler beim Prüfen der Inbox-Regeln: $($_.Exception.Message)`n"
                }
                
                return $result
            }
            3 { # Aktionen zum Entfernen von Weiterleitungen
                $result = "### Aktionen zum Entfernen von Weiterleitungen für: $($mailboxObj.DisplayName)`n`n"
                $hasForwarding = $false
                
                # PowerShell-Befehle zum Entfernen von Weiterleitungen erstellen
                if (-not [string]::IsNullOrEmpty($mailboxObj.ForwardingAddress) -or -not [string]::IsNullOrEmpty($mailboxObj.ForwardingSmtpAddress)) {
                    $hasForwarding = $true
                    $result += "Befehl zum Entfernen der Postfach-Weiterleitung:`n`n"
                    $result += "```powershell`n"
                    $result += "Set-Mailbox -Identity '$Mailbox' -ForwardingAddress `$null -ForwardingSmtpAddress `$null`n"
                    $result += "```\n\n"
                }
                
                # Inbox-Regeln überprüfen und Befehle zum Entfernen erstellen
                try {
                    $inboxRules = Get-InboxRule -Mailbox $Mailbox -ErrorAction Stop | Where-Object { 
                        -not [string]::IsNullOrEmpty($_.ForwardTo) -or 
                        -not [string]::IsNullOrEmpty($_.ForwardAsAttachmentTo) -or 
                        -not [string]::IsNullOrEmpty($_.RedirectTo) 
                    }
                    
                    if ($inboxRules -and $inboxRules.Count -gt 0) {
                        $hasForwarding = $true
                        $result += "Befehle zum Entfernen der Weiterleitungsregeln:`n`n"
                        $result += "```powershell`n"
                        
                        foreach ($rule in $inboxRules) {
                            $result += "# Regel '${$rule.Name}' entfernen`n"
                            $result += "Remove-InboxRule -Identity '$($rule.Identity)' -Confirm:`$false`n`n"
                        }
                        
                        $result += "```\n\n"
                    }
                } catch {
                    $result += "⚠️ Fehler beim Prüfen der Inbox-Regeln: $($_.Exception.Message)`n"
                }
                
                if (-not $hasForwarding) {
                    $result += "✓ Keine Weiterleitungen gefunden. Keine Aktionen notwendig.`n"
                }
                
                return $result
            }
            4 { # Transport-Regeln auf externe Weiterleitungen prüfen
                $result = "### Transport-Regeln und Externe Weiterleitungen für: $($mailboxObj.DisplayName)`n`n"
                
                # Postfach-Weiterleitungen prüfen
                $hasMailboxForwarding = $false
                if (-not [string]::IsNullOrEmpty($mailboxObj.ForwardingAddress) -or -not [string]::IsNullOrEmpty($mailboxObj.ForwardingSmtpAddress)) {
                    $hasMailboxForwarding = $true
                    $result += "⚠️ Postfach hat aktive Weiterleitungen konfiguriert`n`n"
                }
                
                # Transport-Regeln prüfen, die mit diesem Postfach zusammenhängen könnten
                try {
                    $result += "Organisationsweite Transport-Regeln, die E-Mails umleiten (kann alle Benutzer betreffen):`n`n"
                    
                    $transportRules = Get-TransportRule -ErrorAction Stop | Where-Object {
                        $_.RedirectMessageTo -ne $null -or 
                        $_.BlindCopyTo -ne $null -or 
                        $_.CopyTo -ne $null -or
                        $_.AddToRecipients -ne $null -or
                        $_.AddBcc -ne $null -or
                        $_.AddCc -ne $null
                    }
                    
                    if ($transportRules -and $transportRules.Count -gt 0) {
                        foreach ($rule in $transportRules) {
                            $result += "- Regel: $($rule.Name)`n"
                            $result += "  Status: $(if ($rule.Enabled) { "Aktiv" } else { "Deaktiviert" })`n"
                            
                            if ($rule.RedirectMessageTo) {
                                $result += "  RedirectMessageTo: $($rule.RedirectMessageTo -join ", ")`n"
                            }
                            if ($rule.BlindCopyTo) {
                                $result += "  BlindCopyTo: $($rule.BlindCopyTo -join ", ")`n"
                            }
                            if ($rule.CopyTo) {
                                $result += "  CopyTo: $($rule.CopyTo -join ", ")`n"
                            }
                            if ($rule.AddToRecipients) {
                                $result += "  AddToRecipients: $($rule.AddToRecipients -join ", ")`n"
                            }
                            if ($rule.AddBcc) {
                                $result += "  AddBcc: $($rule.AddBcc -join ", ")`n"
                            }
                            if ($rule.AddCc) {
                                $result += "  AddCc: $($rule.AddCc -join ", ")`n"
                            }
                            
                            # Zeige die Bedingung, wenn nach bestimmten Empfängern gefiltert wird
                            if ($rule.SentTo -or $rule.From -or $rule.FromMemberOf -or $rule.SentToMemberOf) {
                                $result += "  Bedingungen:`n"
                                if ($rule.SentTo) { $result += "    - SentTo: $($rule.SentTo -join ", ")`n" }
                                if ($rule.From) { $result += "    - From: $($rule.From -join ", ")`n" }
                                if ($rule.FromMemberOf) { $result += "    - FromMemberOf: $($rule.FromMemberOf -join ", ")`n" }
                                if ($rule.SentToMemberOf) { $result += "    - SentToMemberOf: $($rule.SentToMemberOf -join ", ")`n" }
                            }
                            
                            $result += "`n"
                        }
                    } else {
                        $result += "✓ Keine Transport-Regeln mit Weiterleitungen gefunden.`n`n"
                    }
                } catch {
                    $result += "⚠️ Fehler beim Prüfen der Transport-Regeln: $($_.Exception.Message)`n`n"
                }
                
                # Liste der bekannten Spammethoden
                $result += "### Bekannte Methoden für E-Mail-Weiterleitungen:`n`n"
                $result += "1. **Postfach-Weiterleitung** - Direkt über Exchange-Einstellungen für das Postfach`n"
                $result += "2. **Inbox-Regel** - Vom Benutzer erstellte Regeln im Posteingang`n"
                $result += "3. **Transport-Regeln** - Auf Organisationsebene konfigurierte Umleitungen`n"
                $result += "4. **Outlook-Regeln** - Clientseitig im Outlook des Benutzers gespeicherte Regeln`n"
                $result += "5. **Mail-Flow-Connector** - Angepasste Connectors zur Weiterleitung an externe Systeme`n"
                $result += "6. **Automatische Antworten** - Automatische Weiterleitungen über Out-of-Office Nachrichten`n"
                
                return $result
            }
            5 { # Vollständige Weiterleitungsdetails
                $result = "### Vollständige Weiterleitungsdetails für: $($mailboxObj.DisplayName)`n`n"
                
                $result += "**Postfacheinstellungen:**`n"
                $result += "ForwardingAddress: $($mailboxObj.ForwardingAddress)`n"
                $result += "ForwardingSmtpAddress: $($mailboxObj.ForwardingSmtpAddress)`n"
                $result += "DeliverToMailboxAndForward: $($mailboxObj.DeliverToMailboxAndForward)`n`n"
                
                $result += "**E-Mail-Aliase und alternative Adressen:**`n"
                $result += "PrimarySmtpAddress: $($mailboxObj.PrimarySmtpAddress)`n"
                
                if ($mailboxObj.EmailAddresses) {
                    $result += "EmailAddresses:`n"
                    foreach ($address in $mailboxObj.EmailAddresses) {
                        $result += "- $address`n"
                    }
                }
                
                $result += "`n**Weiterleitungsregeln:**`n"
                
                try {
                    $allInboxRules = Get-InboxRule -Mailbox $Mailbox -ErrorAction Stop
                    
                    if ($allInboxRules -and $allInboxRules.Count -gt 0) {
                        foreach ($rule in $allInboxRules) {
                            $result += "- Regel: $($rule.Name)`n"
                            $result += "  Aktiviert: $($rule.Enabled)`n"
                            $result += "  Priorität: $($rule.Priority)`n"
                            
                            if ($rule.ForwardTo) {
                                $result += "  ForwardTo: $($rule.ForwardTo -join ', ')`n"
                            }
                            if ($rule.ForwardAsAttachmentTo) {
                                $result += "  ForwardAsAttachmentTo: $($rule.ForwardAsAttachmentTo -join ', ')`n"
                            }
                            if ($rule.RedirectTo) {
                                $result += "  RedirectTo: $($rule.RedirectTo -join ', ')`n"
                            }
                            
                            $result += "  Bedingungen:`n"
                            $properties = $rule | Get-Member -MemberType Properties | Where-Object { $_.Name -ne 'ForwardTo' -and $_.Name -ne 'ForwardAsAttachmentTo' -and $_.Name -ne 'RedirectTo' -and $_.Name -ne 'Name' -and $_.Name -ne 'Enabled' -and $_.Name -ne 'Priority' }
                            
                            foreach ($prop in $properties) {
                                $value = $rule.($prop.Name)
                                if ($null -ne $value -and $value -ne '') {
                                    $result += "    $($prop.Name): $value`n"
                                }
                            }
                            
                            $result += "`n"
                        }
                    } else {
                        $result += "Keine Inbox-Regeln gefunden.`n"
                    }
                } catch {
                    $result += "⚠️ Fehler beim Abrufen der Inbox-Regeln: $($_.Exception.Message)`n"
                }
                
                return $result
            }
            default {
                return "Unbekannter Informationstyp: $InfoType"
            }
        }
    }
    catch {
        Write-DebugMessage "Fehler beim Abrufen der Weiterleitungsinformationen: $($_.Exception.Message)" -Type "Error"
        return "Fehler beim Abrufen der Weiterleitungsinformationen: $($_.Exception.Message)"
    }
}

function Get-FormattedMailboxInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [int]$InfoType,

        [Parameter(Mandatory = $false)]
        [string]$NavigationType = ""
    )
    
    try {
        # Status in der GUI aktualisieren
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Rufe Postfachinformationen ab..."
        }
        
        # Bestimme NavigationType anhand des ausgewählten Dropdown-Eintrags wenn nicht explizit übergeben
        if ([string]::IsNullOrEmpty($NavigationType) -and $null -ne $cmbAuditCategory) {
            $NavigationType = $cmbAuditCategory.SelectedItem.Content
        }
        
        Write-DebugMessage "Führe Mailbox-Audit aus. NavigationType: $NavigationType, InfoType: $InfoType, Mailbox: $Mailbox" -Type "Info"
        
        switch ($NavigationType) {
            "Postfach-Informationen" {
                return Get-MailboxInfoForUser -Mailbox $Mailbox -InfoType $InfoType
            }
            "Postfach-Statistiken" {
                return Get-MailboxStatisticsForUser -Mailbox $Mailbox -InfoType $InfoType
            }
            "Postfach-Berechtigungen" {
                return Get-MailboxPermissionsSummary -Mailbox $Mailbox -InfoType $InfoType
            }
            "Audit-Konfiguration" {
                return Get-MailboxAuditConfigForUser -Mailbox $Mailbox -InfoType $InfoType
            }
            "E-Mail-Weiterleitung" {
                return Get-MailboxForwardingForUser -Mailbox $Mailbox -InfoType $InfoType
            }
            default {
                return "Nicht unterstützter Navigationstyp: $NavigationType"
            }
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Informationen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Informationen (Typ $InfoType, Navigation $NavigationType): $errorMsg"
        return "Fehler: $errorMsg`n`nBitte überprüfen Sie die Eingabe und die Verbindung zu Exchange Online."
    }
}

function Get-MailboxInfoForUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [int]$InfoType
    )
    
    try {
        # Exchange-Postfachobjekt abrufen
        $mailboxObj = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        
        # Verschiedene Informationstypen basierend auf InfoType
        switch ($InfoType) {
            1 { # Grundlegende Informationen
                $result = "### Grundlegende Postfach-Informationen für: $($mailboxObj.DisplayName)`n`n"
                $result += "Name: $($mailboxObj.DisplayName)`n"
                $result += "E-Mail: $($mailboxObj.PrimarySmtpAddress)`n"
                $result += "Typ: $($mailboxObj.RecipientTypeDetails)`n"
                $result += "Alias: $($mailboxObj.Alias)`n"
                $result += "ExchangeGUID: $($mailboxObj.ExchangeGUID)`n"
                $result += "Archiv aktiviert: $($mailboxObj.ArchiveStatus)`n"
                $result += "Mailbox aktiviert: $($mailboxObj.IsMailboxEnabled)`n"
                $result += "LitigationHold: $($mailboxObj.LitigationHoldEnabled)`n"
                $result += "Single Item Recovery: $($mailboxObj.SingleItemRecoveryEnabled)`n"
                $result += "Gelöschte Elemente aufbewahren: $($mailboxObj.RetainDeletedItemsFor)`n"
                $result += "Organisation: $($mailboxObj.OrganizationalUnit)`n"
                return $result
            }
            2 { # Speicherbegrenzungen
                $result = "### Speicherbegrenzungen für Postfach: $($mailboxObj.DisplayName)`n`n"
                $result += "ProhibitSendQuota: $($mailboxObj.ProhibitSendQuota)`n"
                $result += "ProhibitSendReceiveQuota: $($mailboxObj.ProhibitSendReceiveQuota)`n"
                $result += "IssueWarningQuota: $($mailboxObj.IssueWarningQuota)`n"
                $result += "UseDatabaseQuotaDefaults: $($mailboxObj.UseDatabaseQuotaDefaults)`n"
                $result += "Archiv Status: $($mailboxObj.ArchiveStatus)`n"
                $result += "Archiv Quota: $($mailboxObj.ArchiveQuota)`n"
                $result += "Archiv Warning Quota: $($mailboxObj.ArchiveWarningQuota)`n"
                return $result
            }
            3 { # E-Mail-Adressen
                $result = "### E-Mail-Adressen für Postfach: $($mailboxObj.DisplayName)`n`n"
                $result += "Primäre SMTP-Adresse: $($mailboxObj.PrimarySmtpAddress)`n`n"
                $result += "Alle E-Mail-Adressen:`n"
                foreach ($address in $mailboxObj.EmailAddresses) {
                    $result += "- $address`n"
                }
                return $result
            }
            4 { # Funktion/Rolle
                $result = "### Funktion und Rolle für Postfach: $($mailboxObj.DisplayName)`n`n"
                $result += "RecipientType: $($mailboxObj.RecipientType)`n"
                $result += "RecipientTypeDetails: $($mailboxObj.RecipientTypeDetails)`n"
                $result += "CustomAttribute1: $($mailboxObj.CustomAttribute1)`n"
                $result += "CustomAttribute2: $($mailboxObj.CustomAttribute2)`n"
                $result += "CustomAttribute3: $($mailboxObj.CustomAttribute3)`n"
                $result += "Department: $($mailboxObj.Department)`n"
                $result += "Office: $($mailboxObj.Office)`n"
                $result += "Title: $($mailboxObj.Title)`n"
                return $result
            }
            5 { # Alle Details (detaillierte Ausgabe)
                $result = "### Vollständige Postfach-Details für: $($mailboxObj.DisplayName)`n`n"
                $mailboxDetails = $mailboxObj | Format-List | Out-String
                $result += $mailboxDetails
                return $result
            }
            default {
                return "Unbekannter Informationstyp: $InfoType"
            }
        }
    }
    catch {
        Write-DebugMessage "Fehler beim Abrufen der Postfachinformationen: $($_.Exception.Message)" -Type "Error"
        return "Fehler beim Abrufen der Postfachinformationen: $($_.Exception.Message)"
    }
}

function Get-MailboxStatisticsForUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [int]$InfoType
    )
    
    try {
        # Mailbox-Statistiken abrufen
        $stats = Get-MailboxStatistics -Identity $Mailbox -ErrorAction Stop
        
        switch ($InfoType) {
            1 { # Größeninformationen
                $result = "### Größeninformationen für Postfach: $($stats.DisplayName)`n`n"
                $result += "Gesamtgröße: $($stats.TotalItemSize)`n"
                $result += "Elemente-Anzahl: $($stats.ItemCount)`n"
                $result += "Gelöschte Elemente: $($stats.DeletedItemCount)`n"
                $result += "Gelöschte Elemente Größe: $($stats.TotalDeletedItemSize)`n"
                $result += "Letzte Anmeldung: $($stats.LastLogonTime)`n"
                $result += "Letzte logoff Zeit: $($stats.LastLogoffTime)`n"
                return $result
            }
            2 { # Ordnerinformationen
                $result = "### Ordnerinformationen für Postfach: $Mailbox`n`n"
                $result += "Die Ordnerinformationen werden gesammelt...`n`n"
                
                try {
                    $folders = Get-MailboxFolderStatistics -Identity $Mailbox -ErrorAction Stop
                    $result += "Anzahl der Ordner: $($folders.Count)`n`n"
                    $result += "Top 15 Ordner nach Größe:`n"
                    $result += "----------------------------`n"
                    $folders | Sort-Object -Property FolderSize -Descending | Select-Object -First 15 | ForEach-Object {
                        $result += "$($_.Name): $($_.FolderSize) ($($_.ItemsInFolder) Elemente)`n"
                    }
                }
                catch {
                    $result += "Fehler beim Abrufen der Ordnerinformationen: $($_.Exception.Message)`n"
                }
                
                return $result
            }
            3 { # Nutzungsstatistiken
                $result = "### Nutzungsstatistiken für Postfach: $($stats.DisplayName)`n`n"
                $result += "Letzte Anmeldung: $($stats.LastLogonTime)`n"
                $result += "Letzte Abmeldung: $($stats.LastLogoffTime)`n"
                $result += "Zugriffe gesamt: $($stats.LogonCount)`n"
                
                $daysInactiveSinceLogon = $null
                if ($stats.LastLogonTime) {
                    $daysInactiveSinceLogon = (New-TimeSpan -Start $stats.LastLogonTime -End (Get-Date)).Days
                    $result += "Tage seit letzter Anmeldung: $daysInactiveSinceLogon`n"
                } else {
                    $result += "Tage seit letzter Anmeldung: Keine Anmeldung aufgezeichnet`n"
                }
                
                if ($daysInactiveSinceLogon -gt 90) {
                    $result += "`n⚠️ Hinweis: Dieses Postfach wurde seit mehr als 90 Tagen nicht genutzt.`n"
                }
                
                return $result
            }
            4 { # Zusammenfassende Statistiken
                try {
                    $stats = Get-MailboxStatistics -Identity $Mailbox -IncludeMoveReport -IncludeMoveHistory -ErrorAction Stop
                    $result = "### Zusammenfassende Statistiken für Postfach: $($stats.DisplayName)`n`n"
                    $result += "Gesamtgröße: $($stats.TotalItemSize)`n"
                    $result += "Anzahl Elemente: $($stats.ItemCount)`n"
                    $result += "Letzte Anmeldung: $($stats.LastLogonTime)`n`n"
                    
                    # Migrationsinfos, falls vorhanden
                    if ($stats.MoveHistory) {
                        $result += "### Migrationshistorie:`n"
                        foreach ($move in $stats.MoveHistory) {
                            $result += "- Status: $($move.Status), Gestartet: $($move.StartTime), Beendet: $($move.EndTime)`n"
                        }
                    }
                    
                    return $result
                }
                catch {
                    return "Fehler beim Abrufen der erweiterten Statistiken: $($_.Exception.Message)"
                }
            }
            5 { # Alle Statistiken (detaillierte Ausgabe)
                try {
                    $stats = Get-MailboxStatistics -Identity $Mailbox -IncludeMoveReport -IncludeMoveHistory -ErrorAction Stop
                    $result = "### Vollständige Statistiken für Postfach: $($stats.DisplayName)`n`n"
                    $statDetails = $stats | Format-List | Out-String
                    $result += $statDetails
                    return $result
                }
                catch {
                    return "Fehler beim Abrufen der vollständigen Statistiken: $($_.Exception.Message)"
                }
            }
            default {
                return "Unbekannter Statistiktyp: $InfoType"
            }
        }
    }
    catch {
        Write-DebugMessage "Fehler beim Abrufen der Postfach-Statistiken: $($_.Exception.Message)" -Type "Error"
        return "Fehler beim Abrufen der Postfach-Statistiken: $($_.Exception.Message)"
    }
}

function Get-MailboxPermissionsSummary {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [int]$InfoType
    )
    
    try {
        switch ($InfoType) {
            1 { # Postfach-Berechtigungen
                $result = "### Postfach-Berechtigungen für: $Mailbox`n`n"
                
                try {
                    $permissions = Get-MailboxPermission -Identity $Mailbox | Where-Object {
                        $_.User -notlike "NT AUTHORITY\SELF" -and
                        $_.User -notlike "S-1-5*" -and 
                        $_.User -notlike "NT AUTHORITY\SYSTEM" -and
                        $_.IsInherited -eq $false
                    }
                    
                    if ($permissions.Count -gt 0) {
                        $result += "**Direkte Berechtigungen:**`n"
                        foreach ($perm in $permissions) {
                            $result += "- Benutzer: $($perm.User)`n"
                            $result += "  Rechte: $($perm.AccessRights -join ', ')`n"
                            $result += "  Verweigert: $($perm.Deny)`n"
                        }
                    } else {
                        $result += "Keine direkten Postfach-Berechtigungen gefunden.`n"
                    }
                }
                catch {
                    $result += "Fehler beim Abrufen der Postfach-Berechtigungen: $($_.Exception.Message)`n"
                }
                
                return $result
            }
            2 { # SendAs-Berechtigungen
                $result = "### SendAs-Berechtigungen für: $Mailbox`n`n"
                
                try {
                    $sendAsPermissions = Get-RecipientPermission -Identity $Mailbox | Where-Object {
                        $_.Trustee -notlike "NT AUTHORITY\SELF" -and
                        $_.Trustee -notlike "S-1-5*" -and 
                        $_.Trustee -notlike "NT AUTHORITY\SYSTEM"
                    }
                    
                    if ($sendAsPermissions.Count -gt 0) {
                        $result += "**SendAs-Berechtigungen:**`n"
                        foreach ($perm in $sendAsPermissions) {
                            $result += "- Benutzer: $($perm.Trustee)`n"
                            $result += "  Rechte: $($perm.AccessRights -join ', ')`n"
                            $result += "  Verweigert: $($perm.Deny)`n"
                        }
                    } else {
                        $result += "Keine SendAs-Berechtigungen gefunden.`n"
                    }
                }
                catch {
                    $result += "Fehler beim Abrufen der SendAs-Berechtigungen: $($_.Exception.Message)`n"
                }
                
                return $result
            }
            3 { # SendOnBehalf-Berechtigungen
                $result = "### SendOnBehalf-Berechtigungen für: $Mailbox`n`n"
                
                try {
                    $mailboxObj = Get-Mailbox -Identity $Mailbox
                    $onBehalfPermissions = $mailboxObj.GrantSendOnBehalfTo
                    
                    if ($onBehalfPermissions -and $onBehalfPermissions.Count -gt 0) {
                        $result += "**SendOnBehalf-Berechtigungen:**`n"
                        foreach ($perm in $onBehalfPermissions) {
                            $result += "- Benutzer: $perm`n"
                        }
                    } else {
                        $result += "Keine SendOnBehalf-Berechtigungen gefunden.`n"
                    }
                }
                catch {
                    $result += "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $($_.Exception.Message)`n"
                }
                
                return $result
            }
            4 { # Kalenderberechtigungen
                $result = "### Kalenderberechtigungen für: $Mailbox`n`n"
                
                try {
                    # Prüfe deutsche und englische Kalenderordner
                    $calendarPermissions = $null
                    
                    try {
                        # Deutsche Version zuerst versuchen
                        $calendarPermissions = Get-MailboxFolderPermission -Identity "${Mailbox}:\Kalender" -ErrorAction Stop
                    }
                    catch {
                        try {
                            # Englische Version als Fallback
                            $calendarPermissions = Get-MailboxFolderPermission -Identity "${Mailbox}:\Calendar" -ErrorAction Stop
                        }
                        catch {
                            $result += "Fehler beim Abrufen der Kalenderberechtigungen: $($_.Exception.Message)`n"
                            return $result
                        }
                    }
                    
                    if ($calendarPermissions -and $calendarPermissions.Count -gt 0) {
                        $result += "**Kalenderberechtigungen:**`n"
                        foreach ($perm in $calendarPermissions) {
                            if ($perm.User.DisplayName -ne "Standard" -and $perm.User.DisplayName -ne "Default" -and 
                                $perm.User.DisplayName -ne "Anonym" -and $perm.User.DisplayName -ne "Anonymous") {
                                $result += "- Benutzer: $($perm.User.DisplayName)`n"
                                $result += "  Rechte: $($perm.AccessRights -join ', ')`n"
                                $result += "  SharingPermissionFlags: $($perm.SharingPermissionFlags)`n"
                            }
                        }
                        
                        $result += "`n**Standard- und Anonymberechtigungen:**`n"
                        foreach ($perm in $calendarPermissions) {
                            if ($perm.User.DisplayName -eq "Standard" -or $perm.User.DisplayName -eq "Default" -or 
                                $perm.User.DisplayName -eq "Anonym" -or $perm.User.DisplayName -eq "Anonymous") {
                                $result += "- $($perm.User.DisplayName): $($perm.AccessRights -join ', ')`n"
                            }
                        }
                    } else {
                        $result += "Keine Kalenderberechtigungen gefunden.`n"
                    }
                }
                catch {
                    $result += "Fehler beim Verarbeiten der Kalenderberechtigungen: $($_.Exception.Message)`n"
                }
                
                return $result
            }
            5 { # Alle Berechtigungen (kombiniert)
                $result = "### Zusammenfassung aller Berechtigungen für: $Mailbox`n`n"
                
                # Postfach-Berechtigungen
                $result += "#### Postfach-Berechtigungen`n"
                try {
                    $permissions = Get-MailboxPermission -Identity $Mailbox | Where-Object {
                        $_.User -notlike "NT AUTHORITY\SELF" -and
                        $_.User -notlike "S-1-5*" -and 
                        $_.User -notlike "NT AUTHORITY\SYSTEM" -and
                        $_.IsInherited -eq $false
                    }
                    
                    if ($permissions.Count -gt 0) {
                        foreach ($perm in $permissions) {
                            $result += "- Benutzer: $($perm.User)`n"
                            $result += "  Rechte: $($perm.AccessRights -join ', ')`n"
                        }
                    } else {
                        $result += "Keine direkten Postfach-Berechtigungen gefunden.`n"
                    }
                }
                catch {
                    $result += "Fehler beim Abrufen der Postfach-Berechtigungen: $($_.Exception.Message)`n"
                }
                
                # SendAs-Berechtigungen
                $result += "`n#### SendAs-Berechtigungen`n"
                try {
                    $sendAsPermissions = Get-RecipientPermission -Identity $Mailbox | Where-Object {
                        $_.Trustee -notlike "NT AUTHORITY\SELF" -and
                        $_.Trustee -notlike "S-1-5*" -and 
                        $_.Trustee -notlike "NT AUTHORITY\SYSTEM"
                    }
                    
                    if ($sendAsPermissions.Count -gt 0) {
                        foreach ($perm in $sendAsPermissions) {
                            $result += "- Benutzer: $($perm.Trustee)`n"
                        }
                    } else {
                        $result += "Keine SendAs-Berechtigungen gefunden.`n"
                    }
                }
                catch {
                    $result += "Fehler beim Abrufen der SendAs-Berechtigungen: $($_.Exception.Message)`n"
                }
                
                # SendOnBehalf-Berechtigungen
                $result += "`n#### SendOnBehalf-Berechtigungen`n"
                try {
                    $mailboxObj = Get-Mailbox -Identity $Mailbox
                    $onBehalfPermissions = $mailboxObj.GrantSendOnBehalfTo
                    
                    if ($onBehalfPermissions -and $onBehalfPermissions.Count -gt 0) {
                        foreach ($perm in $onBehalfPermissions) {
                            $result += "- Benutzer: $perm`n"
                        }
                    } else {
                        $result += "Keine SendOnBehalf-Berechtigungen gefunden.`n"
                    }
                }
                catch {
                    $result += "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $($_.Exception.Message)`n"
                }
                
                return $result
            }
            default {
                return "Unbekannter Berechtigungstyp: $InfoType"
            }
        }
    }
    catch {
        Write-DebugMessage "Fehler beim Abrufen der Berechtigungszusammenfassung: $($_.Exception.Message)" -Type "Error"
        return "Fehler beim Abrufen der Berechtigungszusammenfassung: $($_.Exception.Message)"
    }
}

# -------------------------------------------------
# Abschnitt: SendAs Berechtigungen
# -------------------------------------------------
function Add-SendAsPermissionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUser,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Validate-Email -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage "SendAs-Berechtigung hinzufügen: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfen, ob die Berechtigung bereits existiert
        $existingPermissions = Get-RecipientPermission -Identity $SourceUser -Trustee $TargetUser -ErrorAction SilentlyContinue
        
        if ($existingPermissions) {
            Write-DebugMessage "SendAs-Berechtigung existiert bereits, keine Änderung notwendig" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "SendAs-Berechtigung bereits vorhanden." -Color $script:connectedBrush
            }
            Log-Action "SendAs-Berechtigung bereits vorhanden: $SourceUser -> $TargetUser"
            return $true
        }
        
        # Berechtigung hinzufügen
        Add-RecipientPermission -Identity $SourceUser -Trustee $TargetUser -AccessRights SendAs -Confirm:$false -ErrorAction Stop
        
        Write-DebugMessage "SendAs-Berechtigung erfolgreich hinzugefügt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "SendAs-Berechtigung hinzugefügt." -Color $script:connectedBrush
        }
        Log-Action "SendAs-Berechtigung hinzugefügt: $SourceUser -> $TargetUser"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen der SendAs-Berechtigung: $errorMsg" -Type "Error"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        Log-Action "Fehler beim Hinzufügen der SendAs-Berechtigung: $errorMsg"
        return $false
    }
}

function Remove-SendAsPermissionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUser,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Validate-Email -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage "Entferne SendAs-Berechtigung: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfen, ob die Berechtigung existiert
        $existingPermissions = Get-RecipientPermission -Identity $SourceUser -Trustee $TargetUser -ErrorAction SilentlyContinue
        if (-not $existingPermissions) {
            Write-DebugMessage "Keine SendAs-Berechtigung zum Entfernen gefunden" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine SendAs-Berechtigung zum Entfernen gefunden."
            }
            Log-Action "Keine SendAs-Berechtigung zum Entfernen gefunden: $SourceUser -> $TargetUser"
            return $false
        }
        
        # Berechtigung entfernen
        Remove-RecipientPermission -Identity $SourceUser -Trustee $TargetUser -AccessRights SendAs -Confirm:$false -ErrorAction Stop
        
        Write-DebugMessage "SendAs-Berechtigung erfolgreich entfernt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "SendAs-Berechtigung entfernt." -Color $script:connectedBrush
        }
        Log-Action "SendAs-Berechtigung entfernt: $SourceUser -> $TargetUser"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen der SendAs-Berechtigung: $errorMsg" -Type "Error"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        Log-Action "Fehler beim Entfernen der SendAs-Berechtigung: $errorMsg"
        return $false
    }
}

function Get-SendAsPermissionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $MailboxUser)) {
            throw "Ungültige E-Mail-Adresse für Postfach."
        }
        
        Write-DebugMessage "Rufe SendAs-Berechtigungen ab für: $MailboxUser" -Type "Info"
        
        # Berechtigungen abrufen
        $permissions = Get-RecipientPermission -Identity $MailboxUser -ErrorAction Stop
        
        # Wichtig: Deserialisierungsproblem beheben, indem wir die Daten in ein neues Array umwandeln
        $processedPermissions = @()
        
        foreach ($permission in $permissions) {
            $permObj = [PSCustomObject]@{
                Identity = $permission.Identity
                User = $permission.Trustee.ToString()
                AccessRights = "SendAs"
                IsInherited = $permission.IsInherited
                Deny = $false
            }
            $processedPermissions += $permObj
            Write-DebugMessage "SendAs-Berechtigung verarbeitet: $($permission.Trustee)" -Type "Info"
        }
        
        Write-DebugMessage "SendAs-Berechtigungen abgerufen und verarbeitet: $($processedPermissions.Count) Einträge gefunden" -Type "Success"
        Log-Action "SendAs-Berechtigungen für $MailboxUser abgerufen: $($processedPermissions.Count) Einträge gefunden"
        return $processedPermissions
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der SendAs-Berechtigungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der SendAs-Berechtigungen: $errorMsg"
        
        # Bei Fehler ein leeres Array zurückgeben, damit die GUI nicht abstürzt
        return @()
    }
}

# -------------------------------------------------
# Abschnitt: SendOnBehalf Berechtigungen
# -------------------------------------------------
function Add-SendOnBehalfPermissionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUser,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Validate-Email -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage "Füge SendOnBehalf-Berechtigung hinzu: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfen, ob die Berechtigung bereits existiert
        $mailbox = Get-Mailbox -Identity $SourceUser -ErrorAction Stop
        $currentDelegates = $mailbox.GrantSendOnBehalfTo
        
        if ($currentDelegates -contains $TargetUser) {
            Write-DebugMessage "SendOnBehalf-Berechtigung existiert bereits, keine Änderung notwendig" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "SendOnBehalf-Berechtigung bereits vorhanden." -Color $script:connectedBrush
            }
            Log-Action "SendOnBehalf-Berechtigung bereits vorhanden: $SourceUser -> $TargetUser"
            return $true
        }
        
        # Berechtigung hinzufügen (bestehende Berechtigungen beibehalten)
        $newDelegates = $currentDelegates + $TargetUser
        Set-Mailbox -Identity $SourceUser -GrantSendOnBehalfTo $newDelegates -ErrorAction Stop
        
        Write-DebugMessage "SendOnBehalf-Berechtigung erfolgreich hinzugefügt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "SendOnBehalf-Berechtigung hinzugefügt." -Color $script:connectedBrush
        }
        Log-Action "SendOnBehalf-Berechtigung hinzugefügt: $SourceUser -> $TargetUser"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen der SendOnBehalf-Berechtigung: $errorMsg" -Type "Error"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        Log-Action "Fehler beim Hinzufügen der SendOnBehalf-Berechtigung: $errorMsg"
        return $false
    }
}

function Remove-SendOnBehalfPermissionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUser,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Validate-Email -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage "Entferne SendOnBehalf-Berechtigung: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfen, ob die Berechtigung existiert
        $mailbox = Get-Mailbox -Identity $SourceUser -ErrorAction Stop
        $currentDelegates = $mailbox.GrantSendOnBehalfTo
        
        if (-not ($currentDelegates -contains $TargetUser)) {
            Write-DebugMessage "Keine SendOnBehalf-Berechtigung zum Entfernen gefunden" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine SendOnBehalf-Berechtigung zum Entfernen gefunden."
            }
            Log-Action "Keine SendOnBehalf-Berechtigung zum Entfernen gefunden: $SourceUser -> $TargetUser"
            return $false
        }
        
        # Berechtigung entfernen
        $newDelegates = $currentDelegates | Where-Object { $_ -ne $TargetUser }
        Set-Mailbox -Identity $SourceUser -GrantSendOnBehalfTo $newDelegates -ErrorAction Stop
        
        Write-DebugMessage "SendOnBehalf-Berechtigung erfolgreich entfernt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "SendOnBehalf-Berechtigung entfernt." -Color $script:connectedBrush
        }
        Log-Action "SendOnBehalf-Berechtigung entfernt: $SourceUser -> $TargetUser"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen der SendOnBehalf-Berechtigung: $errorMsg" -Type "Error"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        Log-Action "Fehler beim Entfernen der SendOnBehalf-Berechtigung: $errorMsg"
        return $false
    }
}

function Get-SendOnBehalfPermissionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Validate-Email -Email $MailboxUser)) {
            throw "Ungültige E-Mail-Adresse für Postfach."
        }
        
        Write-DebugMessage "Rufe SendOnBehalf-Berechtigungen ab für: $MailboxUser" -Type "Info"
        
        # Mailbox abrufen
        $mailbox = Get-Mailbox -Identity $MailboxUser -ErrorAction Stop
        
        # SendOnBehalf-Berechtigungen extrahieren
        $delegates = $mailbox.GrantSendOnBehalfTo
        
        # Ergebnisse in ein Array von Objekten umwandeln
        $processedDelegates = @()
        
        if ($delegates -and $delegates.Count -gt 0) {
            foreach ($delegate in $delegates) {
                # Versuche, den anzeigenamen des Delegated zu bekommen
                try {
                    $delegateUser = Get-User -Identity $delegate -ErrorAction SilentlyContinue
                    $displayName = if ($delegateUser) { $delegateUser.DisplayName } else { $delegate }
                }
                catch {
                    $displayName = $delegate
                }
                
                $permObj = [PSCustomObject]@{
                    Identity = $MailboxUser
                    User = $delegate
                    DisplayName = $displayName
                    AccessRights = "SendOnBehalf"
                    IsInherited = $false
                    Deny = $false
                }
                
                $processedDelegates += $permObj
                Write-DebugMessage "SendOnBehalf-Berechtigung verarbeitet: $delegate" -Type "Info"
            }
        }
        
        Write-DebugMessage "SendOnBehalf-Berechtigungen abgerufen: $($processedDelegates.Count) Einträge gefunden" -Type "Success"
        Log-Action "SendOnBehalf-Berechtigungen für $MailboxUser abgerufen: $($processedDelegates.Count) Einträge gefunden"
        
        return $processedDelegates
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $errorMsg"
        
        # Bei Fehler ein leeres Array zurückgeben, damit die GUI nicht abstürzt
        return @()
    }
}

# -------------------------------------------------
# Abschnitt: Gruppen/Verteiler-Funktionen
# -------------------------------------------------
function New-DistributionGroupAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        
        [Parameter(Mandatory = $true)]
        [string]$GroupEmail,
        
        [Parameter(Mandatory = $true)]
        [string]$GroupType,
        
        [Parameter(Mandatory = $false)]
        [string]$Members = "",
        
        [Parameter(Mandatory = $false)]
        [string]$Description = ""
    )
    
    try {
        Write-DebugMessage "Erstelle neue Gruppe: $GroupName ($GroupType)" -Type "Info"
        
        # Parameter für die Gruppenerstellung vorbereiten
        $params = @{
            Name = $GroupName
            PrimarySmtpAddress = $GroupEmail
            DisplayName = $GroupName
        }
        
        if (-not [string]::IsNullOrWhiteSpace($Description)) {
            $params.Add("Notes", $Description)
        }
        
        # Je nach Gruppentyp die passende Funktion aufrufen
        switch ($GroupType) {
            "Verteilergruppe" {
                New-DistributionGroup @params -Type "Distribution" -ErrorAction Stop
                $groupCreated = $true
            }
            "Sicherheitsgruppe" {
                New-DistributionGroup @params -Type "Security" -ErrorAction Stop
                $groupCreated = $true
            }
            "Microsoft 365-Gruppe" {
                # Für Microsoft 365-Gruppen ggf. einen anderen Cmdlet verwenden
                New-UnifiedGroup @params -ErrorAction Stop
                $groupCreated = $true
            }
            "E-Mail-aktivierte Sicherheitsgruppe" {
                New-DistributionGroup @params -Type "Security" -ErrorAction Stop
                $groupCreated = $true
            }
            default {
                throw "Unbekannter Gruppentyp: $GroupType"
            }
        }
        
        # Wenn die Gruppe erstellt wurde und Members angegeben wurden, diese hinzufügen
        if ($groupCreated -and -not [string]::IsNullOrWhiteSpace($Members)) {
            $memberList = $Members -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            
            foreach ($member in $memberList) {
                try {
                    Add-DistributionGroupMember -Identity $GroupEmail -Member $member.Trim() -ErrorAction Stop
                    Write-DebugMessage "Mitglied $member zu Gruppe $GroupName hinzugefügt" -Type "Info"
                }
                catch {
                    Write-DebugMessage "Fehler beim Hinzufügen von $member zu Gruppe $GroupName - $($_.Exception.Message)" -Type "Warning"
                }
            }
        }
        
        Write-DebugMessage "Gruppe $GroupName erfolgreich erstellt" -Type "Success"
        Log-Action "Gruppe $GroupName ($GroupType) mit E-Mail $GroupEmail erstellt"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Gruppe $GroupName erfolgreich erstellt."
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Erstellen der Gruppe: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Erstellen der Gruppe $GroupName - $errorMsg"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler beim Erstellen der Gruppe: $errorMsg"
        }
        
        return $false
    }
}

function Remove-DistributionGroupAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )
    
    try {
        Write-DebugMessage "Lösche Gruppe: $GroupName" -Type "Info"
        
        # Prüfen, welcher Gruppentyp vorliegt
        $isUnifiedGroup = $false
        try {
            $group = Get-UnifiedGroup -Identity $GroupName -ErrorAction SilentlyContinue
            if ($null -ne $group) {
                $isUnifiedGroup = $true
            }
        }
        catch {
            $isUnifiedGroup = $false
        }
        
        # Je nach Gruppentyp die passende Funktion aufrufen
        if ($isUnifiedGroup) {
            Remove-UnifiedGroup -Identity $GroupName -Confirm:$false -ErrorAction Stop
            Write-DebugMessage "Microsoft 365-Gruppe $GroupName erfolgreich gelöscht" -Type "Success"
        }
        else {
            Remove-DistributionGroup -Identity $GroupName -Confirm:$false -ErrorAction Stop
            Write-DebugMessage "Verteilerliste/Sicherheitsgruppe $GroupName erfolgreich gelöscht" -Type "Success"
        }
        
        Log-Action "Gruppe $GroupName wurde gelöscht"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Gruppe $GroupName erfolgreich gelöscht."
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Löschen der Gruppe: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Löschen der Gruppe $GroupName - $errorMsg"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler beim Löschen der Gruppe: $errorMsg"
        }
        
        return $false
    }
}

function Add-GroupMemberAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        
        [Parameter(Mandatory = $true)]
        [string]$MemberIdentity
    )
    
    try {
        Write-DebugMessage "Füge $MemberIdentity zu Gruppe $GroupName hinzu" -Type "Info"
        
        # Prüfen, welcher Gruppentyp vorliegt
        $isUnifiedGroup = $false
        try {
            $group = Get-UnifiedGroup -Identity $GroupName -ErrorAction SilentlyContinue
            if ($null -ne $group) {
                $isUnifiedGroup = $true
            }
        }
        catch {
            $isUnifiedGroup = $false
        }
        
        # Je nach Gruppentyp die passende Funktion aufrufen
        if ($isUnifiedGroup) {
            Add-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $MemberIdentity -ErrorAction Stop
            Write-DebugMessage "$MemberIdentity erfolgreich zur Microsoft 365-Gruppe $GroupName hinzugefügt" -Type "Success"
        }
        else {
            Add-DistributionGroupMember -Identity $GroupName -Member $MemberIdentity -ErrorAction Stop
            Write-DebugMessage "$MemberIdentity erfolgreich zur Gruppe $GroupName hinzugefügt" -Type "Success"
        }
        
        Log-Action "Benutzer $MemberIdentity zur Gruppe $GroupName hinzugefügt"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Benutzer $MemberIdentity erfolgreich zur Gruppe hinzugefügt."
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen des Benutzers zur Gruppe: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Hinzufügen von $MemberIdentity zu $GroupName - $errorMsg"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler beim Hinzufügen des Benutzers zur Gruppe: $errorMsg"
        }
        
        return $false
    }
}

function Remove-GroupMemberAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        
        [Parameter(Mandatory = $true)]
        [string]$MemberIdentity
    )
    
    try {
        Write-DebugMessage "Entferne $MemberIdentity aus Gruppe $GroupName" -Type "Info"
        
        # Prüfen, welcher Gruppentyp vorliegt
        $isUnifiedGroup = $false
        try {
            $group = Get-UnifiedGroup -Identity $GroupName -ErrorAction SilentlyContinue
            if ($null -ne $group) {
                $isUnifiedGroup = $true
            }
        }
        catch {
            $isUnifiedGroup = $false
        }
        
        # Je nach Gruppentyp die passende Funktion aufrufen
        if ($isUnifiedGroup) {
            Remove-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $MemberIdentity -Confirm:$false -ErrorAction Stop
            Write-DebugMessage "$MemberIdentity erfolgreich aus Microsoft 365-Gruppe $GroupName entfernt" -Type "Success"
        }
        else {
            Remove-DistributionGroupMember -Identity $GroupName -Member $MemberIdentity -Confirm:$false -ErrorAction Stop
            Write-DebugMessage "$MemberIdentity erfolgreich aus Gruppe $GroupName entfernt" -Type "Success"
        }
        
        Log-Action "Benutzer $MemberIdentity aus Gruppe $GroupName entfernt"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Benutzer $MemberIdentity erfolgreich aus Gruppe entfernt."
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen des Benutzers aus der Gruppe: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Entfernen von $MemberIdentity aus $GroupName - $errorMsg"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler beim Entfernen des Benutzers aus der Gruppe: $errorMsg"
        }
        
        return $false
    }
}

function Get-GroupMembersAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )
    
    try {
        Write-DebugMessage "Rufe Mitglieder der Gruppe $GroupName ab" -Type "Info"
        
        # Prüfen, welcher Gruppentyp vorliegt
        $isUnifiedGroup = $false
        try {
            $group = Get-UnifiedGroup -Identity $GroupName -ErrorAction SilentlyContinue
            if ($null -ne $group) {
                $isUnifiedGroup = $true
                $groupObj = $group
            }
        }
        catch {
            $isUnifiedGroup = $false
            try {
                $groupObj = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop
            }
            catch {
                throw "Gruppe $GroupName nicht gefunden."
            }
        }
        
        # Je nach Gruppentyp die passende Funktion aufrufen
        if ($isUnifiedGroup) {
            $members = Get-UnifiedGroupLinks -Identity $GroupName -LinkType Members -ErrorAction Stop
        }
        else {
            $members = Get-DistributionGroupMember -Identity $GroupName -ErrorAction Stop
        }
        
        # Objekte für die DataGrid-Anzeige aufbereiten
        $memberList = @()
        foreach ($member in $members) {
            $memberObj = [PSCustomObject]@{
                DisplayName = $member.DisplayName
                PrimarySmtpAddress = $member.PrimarySmtpAddress
                RecipientType = $member.RecipientType
                HiddenFromAddressListsEnabled = $member.HiddenFromAddressListsEnabled
            }
            $memberList += $memberObj
        }
        
        Write-DebugMessage "Mitglieder der Gruppe $GroupName erfolgreich abgerufen: $($memberList.Count)" -Type "Success"
        
        return $memberList
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Gruppenmitglieder: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Mitglieder von $GroupName - $errorMsg"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler beim Abrufen der Gruppenmitglieder: $errorMsg"
        }
        
        return @()
    }
}

function Get-GroupSettingsAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )
    
    try {
        Write-DebugMessage "Rufe Einstellungen der Gruppe $GroupName ab" -Type "Info"
        
        # Prüfen, welcher Gruppentyp vorliegt
        $isUnifiedGroup = $false
        try {
            $group = Get-UnifiedGroup -Identity $GroupName -ErrorAction SilentlyContinue
            if ($null -ne $group) {
                $isUnifiedGroup = $true
                return $group
            }
        }
        catch {
            $isUnifiedGroup = $false
        }
        
        # Für normale Verteilerlisten/Sicherheitsgruppen
        $group = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop
        return $group
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Gruppeneinstellungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Einstellungen von $GroupName - $errorMsg"
        return $null
    }
}

function Update-GroupSettingsAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        
        [Parameter(Mandatory = $false)]
        [bool]$HiddenFromAddressListsEnabled,
        
        [Parameter(Mandatory = $false)]
        [bool]$RequireSenderAuthenticationEnabled,
        
        [Parameter(Mandatory = $false)]
        [bool]$AllowExternalSenders
    )
    
    try {
        Write-DebugMessage "Aktualisiere Einstellungen für Gruppe $GroupName" -Type "Info"
        
        # Prüfen, welcher Gruppentyp vorliegt
        $isUnifiedGroup = $false
        try {
            $group = Get-UnifiedGroup -Identity $GroupName -ErrorAction SilentlyContinue
            if ($null -ne $group) {
                $isUnifiedGroup = $true
                
                # Parameter für das Update vorbereiten
                $params = @{
                    Identity = $GroupName
                    HiddenFromAddressListsEnabled = $HiddenFromAddressListsEnabled
                }
                
                # Microsoft 365-Gruppe aktualisieren
                Set-UnifiedGroup @params -ErrorAction Stop
                
                # AllowExternalSenders für Microsoft 365-Gruppen setzen
                Set-UnifiedGroup -Identity $GroupName -AcceptMessagesOnlyFromSendersOrMembers $(-not $AllowExternalSenders) -ErrorAction Stop
                
                Write-DebugMessage "Microsoft 365-Gruppe $GroupName erfolgreich aktualisiert" -Type "Success"
            }
        }
        catch {
            $isUnifiedGroup = $false
        }
        
        if (-not $isUnifiedGroup) {
            # Parameter für das Update vorbereiten
            $params = @{
                Identity = $GroupName
                HiddenFromAddressListsEnabled = $HiddenFromAddressListsEnabled
                RequireSenderAuthenticationEnabled = $RequireSenderAuthenticationEnabled
            }
            
            # Verteilerliste/Sicherheitsgruppe aktualisieren
            Set-DistributionGroup @params -ErrorAction Stop
            
            # AcceptMessagesOnlyFromSendersOrMembers für normale Gruppen setzen
            if (-not $AllowExternalSenders) {
                Set-DistributionGroup -Identity $GroupName -AcceptMessagesOnlyFromSendersOrMembers @() -ErrorAction Stop
            } else {
                Set-DistributionGroup -Identity $GroupName -AcceptMessagesOnlyFrom $null -ErrorAction Stop
            }
            
            Write-DebugMessage "Gruppe $GroupName erfolgreich aktualisiert" -Type "Success"
        }
        
        Log-Action "Einstellungen für Gruppe $GroupName aktualisiert"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Gruppeneinstellungen für $GroupName erfolgreich aktualisiert."
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Aktualisieren der Gruppeneinstellungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Aktualisieren der Einstellungen von $GroupName - $errorMsg"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler beim Aktualisieren der Gruppeneinstellungen: $errorMsg"
        }
        
        return $false
    }
}

function New-SharedMailboxAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress
    )
    
    try {
        Write-DebugMessage "Erstelle neue Shared Mailbox: $Name mit Adresse $EmailAddress" -Type "Info"
        New-Mailbox -Name $Name -PrimarySmtpAddress $EmailAddress -Shared -ErrorAction Stop
        Write-DebugMessage "Shared Mailbox $Name erfolgreich erstellt" -Type "Success"
        Log-Action "Shared Mailbox $Name ($EmailAddress) erfolgreich erstellt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Erstellen der Shared Mailbox: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Erstellen der Shared Mailbox $Name - $errorMsg"
        return $false
    }
}

function Convert-ToSharedMailboxAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    
    try {
        Write-DebugMessage "Konvertiere Postfach zu Shared Mailbox: $Identity" -Type "Info"
        Set-Mailbox -Identity $Identity -Type Shared -ErrorAction Stop
        Write-DebugMessage "Postfach $Identity erfolgreich zu Shared Mailbox konvertiert" -Type "Success"
        Log-Action "Postfach $Identity erfolgreich zu Shared Mailbox konvertiert"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Konvertieren des Postfachs: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Konvertieren des Postfachs  - $errorMsg"
        return $false
    }
}

function Add-SharedMailboxPermissionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$PermissionType,
        
        [Parameter(Mandatory = $false)]
        [bool]$AutoMapping = $true
    )
    
    try {
        Write-DebugMessage "Füge Shared Mailbox Berechtigung hinzu: $PermissionType für $User auf $Mailbox" -Type "Info"
        
        switch ($PermissionType) {
            "FullAccess" {
                Add-MailboxPermission -Identity $Mailbox -User $User -AccessRights FullAccess -AutoMapping $AutoMapping -ErrorAction Stop
            }
            "SendAs" {
                Add-RecipientPermission -Identity $Mailbox -Trustee $User -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            }
            "SendOnBehalf" {
                $currentGrantSendOnBehalf = (Get-Mailbox -Identity $Mailbox).GrantSendOnBehalfTo
                if ($currentGrantSendOnBehalf -notcontains $User) {
                    $currentGrantSendOnBehalf += $User
                    Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo $currentGrantSendOnBehalf -ErrorAction Stop
                }
            }
        }
        
        Write-DebugMessage "Shared Mailbox Berechtigung erfolgreich hinzugefügt" -Type "Success"
        Log-Action "Shared Mailbox Berechtigung $PermissionType für $User auf $Mailbox hinzugefügt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen der Shared Mailbox Berechtigung: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Hinzufügen der Shared Mailbox Berechtigung: $errorMsg"
        return $false
    }
}

function Remove-SharedMailboxPermissionAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$PermissionType
    )
    
    try {
        Write-DebugMessage "Entferne Shared Mailbox Berechtigung: $PermissionType für $User auf $Mailbox" -Type "Info"
        
        switch ($PermissionType) {
            "FullAccess" {
                Remove-MailboxPermission -Identity $Mailbox -User $User -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
            }
            "SendAs" {
                Remove-RecipientPermission -Identity $Mailbox -Trustee $User -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            }
            "SendOnBehalf" {
                $currentGrantSendOnBehalf = (Get-Mailbox -Identity $Mailbox).GrantSendOnBehalfTo
                if ($currentGrantSendOnBehalf -contains $User) {
                    $newGrantSendOnBehalf = $currentGrantSendOnBehalf | Where-Object { $_ -ne $User }
                    Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo $newGrantSendOnBehalf -ErrorAction Stop
                }
            }
        }
        
        Write-DebugMessage "Shared Mailbox Berechtigung erfolgreich entfernt" -Type "Success"
        Log-Action "Shared Mailbox Berechtigung $PermissionType für $User auf $Mailbox entfernt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen der Shared Mailbox Berechtigung: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Entfernen der Shared Mailbox Berechtigung: $errorMsg"
        return $false
    }
}

function Get-SharedMailboxPermissionsAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mailbox
    )
    
    try {
        Write-DebugMessage "Rufe Berechtigungen für Shared Mailbox ab: $Mailbox" -Type "Info"
        
        $permissions = @()
        
        # FullAccess-Berechtigungen abrufen
        $fullAccessPerms = Get-MailboxPermission -Identity $Mailbox | Where-Object {
            $_.User -notlike "NT AUTHORITY\SELF" -and 
            $_.AccessRights -like "*FullAccess*" -and 
            $_.IsInherited -eq $false
        }
        
        foreach ($perm in $fullAccessPerms) {
            $permissions += [PSCustomObject]@{
                User = $perm.User.ToString()
                AccessRights = "FullAccess"
                PermissionType = "FullAccess"
            }
        }
        
        # SendAs-Berechtigungen abrufen
        $sendAsPerms = Get-RecipientPermission -Identity $Mailbox | Where-Object {
            $_.Trustee -notlike "NT AUTHORITY\SELF" -and 
            $_.IsInherited -eq $false
        }
        
        foreach ($perm in $sendAsPerms) {
            $permissions += [PSCustomObject]@{
                User = $perm.Trustee.ToString()
                AccessRights = "SendAs"
                PermissionType = "SendAs"
            }
        }
        
        # SendOnBehalf-Berechtigungen abrufen
        $mailboxObj = Get-Mailbox -Identity $Mailbox
        $sendOnBehalfPerms = $mailboxObj.GrantSendOnBehalfTo
        
        foreach ($perm in $sendOnBehalfPerms) {
            $permissions += [PSCustomObject]@{
                User = $perm.ToString()
                AccessRights = "SendOnBehalf"
                PermissionType = "SendOnBehalf"
            }
        }
        
        Write-DebugMessage "Shared Mailbox Berechtigungen erfolgreich abgerufen: $($permissions.Count) Einträge" -Type "Success"
        Log-Action "Shared Mailbox Berechtigungen für $Mailbox abgerufen: $($permissions.Count) Einträge"
        
        return $permissions
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Shared Mailbox Berechtigungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Shared Mailbox Berechtigungen: $errorMsg"
        return @()
    }
}

function Update-SharedMailboxAutoMappingAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [bool]$AutoMapping
    )
    
    try {
        Write-DebugMessage "Aktualisiere AutoMapping für Shared Mailbox $Mailbox auf $AutoMapping" -Type "Info"
        
        # Bestehende FullAccess-Berechtigungen abrufen und neu setzen mit AutoMapping-Parameter
        $fullAccessPerms = Get-MailboxPermission -Identity $Mailbox | Where-Object {
            $_.User -notlike "NT AUTHORITY\SELF" -and 
            $_.AccessRights -like "*FullAccess*" -and 
            $_.IsInherited -eq $false
        }
        
        foreach ($perm in $fullAccessPerms) {
            $user = $perm.User.ToString()
            Remove-MailboxPermission -Identity $Mailbox -User $user -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
            Add-MailboxPermission -Identity $Mailbox -User $user -AccessRights FullAccess -AutoMapping $AutoMapping -ErrorAction Stop
            Write-DebugMessage "AutoMapping für $user auf $Mailbox aktualisiert" -Type "Info"
        }
        
        Write-DebugMessage "AutoMapping für Shared Mailbox erfolgreich aktualisiert" -Type "Success"
        Log-Action "AutoMapping für Shared Mailbox $Mailbox auf $AutoMapping gesetzt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Aktualisieren des AutoMapping: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Aktualisieren des AutoMapping: $errorMsg"
        return $false
    }
}

function Set-SharedMailboxForwardingAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [string]$ForwardingAddress
    )
    
    try {
        Write-DebugMessage "Setze Weiterleitung für Shared Mailbox $Mailbox auf $ForwardingAddress" -Type "Info"
        
        if ([string]::IsNullOrEmpty($ForwardingAddress)) {
            # Weiterleitung entfernen
            Set-Mailbox -Identity $Mailbox -ForwardingAddress $null -ForwardingSmtpAddress $null -ErrorAction Stop
            Write-DebugMessage "Weiterleitung für Shared Mailbox erfolgreich entfernt" -Type "Success"
        } else {
            # Weiterleitung setzen
            Set-Mailbox -Identity $Mailbox -ForwardingSmtpAddress $ForwardingAddress -DeliverToMailboxAndForward $true -ErrorAction Stop
            Write-DebugMessage "Weiterleitung für Shared Mailbox erfolgreich gesetzt" -Type "Success"
        }
        
        Log-Action "Weiterleitung für Shared Mailbox $Mailbox auf $ForwardingAddress gesetzt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Setzen der Weiterleitung: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Setzen der Weiterleitung: $errorMsg"
        return $false
    }
}

function Set-SharedMailboxGALVisibilityAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [bool]$HideFromGAL
    )
    
    try {
        Write-DebugMessage "Setze GAL-Sichtbarkeit für Shared Mailbox $Mailbox auf HideFromGAL=$HideFromGAL" -Type "Info"
        
        Set-Mailbox -Identity $Mailbox -HiddenFromAddressListsEnabled $HideFromGAL -ErrorAction Stop
        
        $visibilityStatus = if ($HideFromGAL) { "ausgeblendet" } else { "sichtbar" }
        Write-DebugMessage "GAL-Sichtbarkeit für Shared Mailbox erfolgreich gesetzt - $visibilityStatus" -Type "Success"
        Log-Action "Shared Mailbox $Mailbox wurde in GAL $visibilityStatus gesetzt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Setzen der GAL-Sichtbarkeit: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Setzen der GAL-Sichtbarkeit: $errorMsg"
        return $false
    }
}

function Remove-SharedMailboxAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mailbox
    )
    
    try {
        Write-DebugMessage "Lösche Shared Mailbox: $Mailbox" -Type "Info"
        
        Remove-Mailbox -Identity $Mailbox -Confirm:$false -ErrorAction Stop
        
        Write-DebugMessage "Shared Mailbox erfolgreich gelöscht" -Type "Success"
        Log-Action "Shared Mailbox $Mailbox wurde gelöscht"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Löschen der Shared Mailbox: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Löschen der Shared Mailbox: $errorMsg"
        return $false
    }
}
function Open-AdminCenterLink {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$DiagnosticIndex
    )
    
    try {
        $diagnostic = $script:exchangeDiagnostics[$DiagnosticIndex]
        
        if (-not [string]::IsNullOrEmpty($diagnostic.AdminCenterLink)) {
            Write-DebugMessage "Öffne Admin Center Link: $($diagnostic.AdminCenterLink)" -Type "Info"
            
            # Status aktualisieren
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Öffne Exchange Admin Center..."
            }
            
            # Link öffnen mit Standard-Browser
            Start-Process $diagnostic.AdminCenterLink
            
            Log-Action "Admin Center Link geöffnet: $($diagnostic.Name)"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Exchange Admin Center geöffnet." -Color $script:connectedBrush
            }
            
            return $true
        }
        else {
            Write-DebugMessage "Kein Admin Center Link für diese Diagnose vorhanden" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Kein Admin Center Link für diese Diagnose vorhanden."
            }
            
            return $false
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Öffnen des Admin Center Links: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler beim Öffnen des Admin Center Links: $errorMsg"
        }
        
        Log-Action "Fehler beim Öffnen des Admin Center Links: $errorMsg"
        return $false
    }
}

# Hilfsfunktion zur Erstellung eines Hilfetexts für die Audit-Funktion
function Get-HelpText {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$Topic = "General"
    )
    
    switch ($Topic) {
        "Calendar" {
            return @"
## Hilfe: Kalenderberechtigungen

Mit diesem Tab können Sie Kalenderberechtigungen für Exchange Online-Postfächer verwalten.

### Hinzufügen von Berechtigungen:
1. Geben Sie die E-Mail-Adresse des Quellpostfachs ein (der Kalender, auf den zugegriffen werden soll)
2. Geben Sie die E-Mail-Adresse des Zielpostfachs ein (der Benutzer, der Zugriff erhalten soll)
3. Wählen Sie die gewünschte Berechtigungsstufe aus
4. Klicken Sie auf "Berechtigung hinzufügen"

### Entfernen von Berechtigungen:
1. Geben Sie die E-Mail-Adresse des Quellpostfachs ein
2. Geben Sie die E-Mail-Adresse des Zielpostfachs ein
3. Klicken Sie auf "Berechtigung entfernen"

### Berechtigungen anzeigen:
1. Geben Sie die E-Mail-Adresse des zu prüfenden Postfachs ein
2. Klicken Sie auf "Berechtigungen anzeigen"

### Berechtigungsstufen:
- Owner: Vollständige Kontrolle über den Kalender
- PublishingEditor: Bearbeiten und Erstellen von Elementen
- Editor: Bearbeiten von Elementen
- PublishingAuthor: Erstellen von Elementen und eigene Elemente bearbeiten
- Author: Eigene Elemente erstellen und bearbeiten
- Reviewer: Nur Lesezugriff
- AvailabilityOnly: Nur Verfügbarkeitsinformationen sehen
- None: Keine Berechtigungen
"@
        }
        "Mailbox" {
            return @"
## Hilfe: Postfachberechtigungen

Mit diesem Tab können Sie Postfachberechtigungen für Exchange Online-Postfächer verwalten.

### Hinzufügen von Berechtigungen:
1. Geben Sie die E-Mail-Adresse des Quellpostfachs ein (das Postfach, auf das zugegriffen werden soll)
2. Geben Sie die E-Mail-Adresse des Zielpostfachs ein (der Benutzer, der Zugriff erhalten soll)
3. Klicken Sie auf die gewünschte Berechtigung hinzufügen:
   - "FullAccess hinzufügen" - Vollzugriff auf das gesamte Postfach
   - "SendAs hinzufügen" - Berechtigung zum Senden als dieses Postfach
   - "SendOnBehalf hinzufügen" - Berechtigung zum Senden im Namen dieses Postfachs

### Entfernen von Berechtigungen:
1. Geben Sie die E-Mail-Adresse des Quellpostfachs ein
2. Geben Sie die E-Mail-Adresse des Zielpostfachs ein
3. Klicken Sie auf die entsprechende Schaltfläche zum Entfernen der Berechtigung

### Berechtigungen anzeigen:
1. Geben Sie die E-Mail-Adresse des zu prüfenden Postfachs ein
2. Klicken Sie auf die entsprechende Schaltfläche zum Anzeigen der Berechtigungen

### Berechtigungstypen:
- FullAccess: Vollständige Kontrolle über das Postfach
- SendAs: Berechtigung, E-Mails als dieser Benutzer zu senden
- SendOnBehalf: Berechtigung, E-Mails im Namen dieses Benutzers zu senden
"@
        }
        "Audit" {
            return @"
## Hilfe: Postfach-Audit und -Information

Mit diesem Tab können Sie verschiedene Informationen zu Exchange Online-Postfächern abrufen:

### Verwendung:
1. Geben Sie die E-Mail-Adresse des zu prüfenden Postfachs ein
2. Wählen Sie die Kategorie der Informationen, die Sie abrufen möchten
3. Wählen Sie den spezifischen Informationstyp
4. Klicken Sie auf "Ausführen", um die Informationen abzurufen

### Verfügbare Kategorien:
- **Postfach-Informationen**: Grundlegende Daten zum Postfach, Speicherbegrenzungen, E-Mail-Adressen
- **Postfach-Statistiken**: Größeninformationen, Ordnerstrukturen, Nutzung
- **Postfach-Berechtigungen**: Zugriffsberechtigungen, SendAs, SendOnBehalf, Kalender
- **Audit-Konfiguration**: Einstellungen zur Überwachung von Postfachaktivitäten
- **E-Mail-Weiterleitung**: Analyse von Weiterleitungseinstellungen und -regeln

### Tipps:
- Die abgerufenen Informationen können für die Diagnose von Problemen und die Überprüfung von Konfigurationen verwendet werden
- Die Audit-Konfiguration zeigt an, ob Postfachaktivitäten protokolliert werden und welche Aktivitäten überwacht werden
- Bei E-Mail-Weiterleitungen werden auch externe Weiterleitungen identifiziert, die ein Sicherheitsrisiko darstellen können
"@
        }
        "Troubleshooting" {
            return @"
## Hilfe: Exchange Online Troubleshooting

Mit diesem Tab können Sie verschiedene Diagnosen für Exchange Online-Probleme durchführen:

### Verwendung:
1. Wählen Sie eine Diagnose aus der Liste aus
2. Geben Sie bei Bedarf Benutzer-E-Mail-Adressen oder andere Parameter ein
3. Klicken Sie auf "Diagnose ausführen", um die Diagnose durchzuführen
4. Die Ergebnisse werden im Ergebnisfenster angezeigt
5. Bei einigen Diagnosen können Sie auch das Exchange Admin Center öffnen, um weitere Aktionen durchzuführen

### Arten von Diagnosen:
- Throttling-Policy-Informationen
- Akzeptierte Domains
- RBAC-Berechtigungen
- Postfachkonfiguration
- Organisationskonfiguration
- Mail-Flow und Zustellungsprobleme
- DKIM-Einstellungen und DNS
- Weiterleitungskonfiguration

### Hinweise:
- Einige Diagnosen erfordern administrative Berechtigungen
- Ergebnisse sind oft technisch und können für die Problemlösung mit dem Microsoft Support hilfreich sein
- Bei Fehlern versuchen Sie, die Verbindung zu trennen und neu herzustellen
"@
        }
        default {
            return @"
## Exchange Berechtigungen Verwaltung Tool - Hilfe

Dies ist ein Tool zur Verwaltung von Exchange Online-Berechtigungen und -Konfigurationen.

### Hauptfunktionen:
- Kalenderberechtigungen verwalten
- Postfachberechtigungen verwalten
- Postfachinformationen und Audit-Konfigurationen abrufen
- Exchange Online-Diagnostik durchführen

### Erste Schritte:
1. Klicken Sie auf "Mit Exchange verbinden", um eine Verbindung zu Exchange Online herzustellen
2. Wählen Sie nach erfolgreicher Verbindung den gewünschten Funktionsbereich aus
3. Folgen Sie den Anweisungen im jeweiligen Bereich

### Hinweise:
- Alle Aktionen werden protokolliert und können im Log-Ordner eingesehen werden
- Bei Fehlern prüfen Sie bitte Ihre Internetverbindung und Ihre Exchange Online-Berechtigungen
- Die Verbindung wird nach einer längeren Inaktivität automatisch getrennt

Für detailliertere Hilfe zu einem bestimmten Bereich wählen Sie bitte den entsprechenden Tab aus.
"@
        }
    }
}

# Event-Handler für die Hilfe-Schaltflächen
function Show-HelpDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Topic = "General"
    )
    
    try {
        $helpText = Get-HelpText -Topic $Topic
        
        # Erstelle ein WPF-Fenster zur Anzeige des Hilfetexts
        $helpWindow = New-Object System.Windows.Window
        $helpWindow.Title = "Exchange Tool - Hilfe"
        $helpWindow.SizeToContent = "WidthAndHeight"
        $helpWindow.MinWidth = 600
        $helpWindow.MinHeight = 400
        $helpWindow.WindowStartupLocation = "CenterScreen"
        
        # Grid erstellen
        $grid = New-Object System.Windows.Controls.Grid
        $helpWindow.Content = $grid
        
        # Zeilen definieren
        $rowDefinition1 = New-Object System.Windows.Controls.RowDefinition
        $rowDefinition1.Height = New-Object System.Windows.GridLength 1, "Star"
        $rowDefinition2 = New-Object System.Windows.Controls.RowDefinition
        $rowDefinition2.Height = New-Object System.Windows.GridLength 40
        $grid.RowDefinitions.Add($rowDefinition1)
        $grid.RowDefinitions.Add($rowDefinition2)
        
        # TextBox erstellen
        $textBox = New-Object System.Windows.Controls.TextBox
        $textBox.IsReadOnly = $true
        $textBox.TextWrapping = "Wrap"
        $textBox.AcceptsReturn = $true
        $textBox.VerticalScrollBarVisibility = "Auto"
        $textBox.HorizontalScrollBarVisibility = "Auto"
        $textBox.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
        $textBox.Margin = New-Object System.Windows.Thickness(10)
        $textBox.Text = $helpText
        $grid.Children.Add($textBox)
        [System.Windows.Controls.Grid]::SetRow($textBox, 0)
        
        # Button-Panel erstellen
        $buttonPanel = New-Object System.Windows.Controls.StackPanel
        $buttonPanel.Orientation = "Horizontal"
        $buttonPanel.HorizontalAlignment = "Right"
        $buttonPanel.Margin = New-Object System.Windows.Thickness(10)
        $grid.Children.Add($buttonPanel)
        [System.Windows.Controls.Grid]::SetRow($buttonPanel, 1)
        
        # Schließen-Button erstellen
        $closeButton = New-Object System.Windows.Controls.Button
        $closeButton.Content = "Schließen"
        $closeButton.Width = 100
        $closeButton.Height = 30
        $closeButton.Margin = New-Object System.Windows.Thickness(5)
        $buttonPanel.Children.Add($closeButton)
        
        # Event-Handler für den Schließen-Button
        $closeButton.Add_Click({
            $helpWindow.Close()
        })
        
        # Fenster anzeigen
        [void]$helpWindow.ShowDialog()
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Anzeigen des Hilfedialogs: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Anzeigen des Hilfedialogs: $errorMsg"
        return $false
    }
}

# -------------------------------------------------
# Event-Handler für die Programmeinstellungen
# -------------------------------------------------
function Show-SettingsDialog {
    [CmdletBinding()]
    param()
    
    try {
        # Einstellungen aus der INI-Datei laden
        $settings = Get-IniContent -FilePath $script:configFilePath
        
        # Erstelle ein WPF-Fenster für die Einstellungen
        $settingsWindow = New-Object System.Windows.Window
        $settingsWindow.Title = "Exchange Tool - Einstellungen"
        $settingsWindow.Width = 650
        $settingsWindow.Height = 565
        $settingsWindow.MinWidth = 500
        $settingsWindow.MinHeight = 400
        $settingsWindow.WindowStartupLocation = "CenterScreen"
        $settingsWindow.Padding = New-Object System.Windows.Thickness(0)
        $settingsWindow.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::WhiteSmoke)
        
        # Haupt-Grid erstellen (enthält ScrollViewer und Button-Bereich)
        $mainGrid = New-Object System.Windows.Controls.Grid
        $settingsWindow.Content = $mainGrid
        
        # Zwei Zeilen definieren: Inhalt und Buttons
        $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))
        $mainGrid.RowDefinitions[0].Height = New-Object System.Windows.GridLength 1, "Star"
        $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))
        $mainGrid.RowDefinitions[1].Height = New-Object System.Windows.GridLength 60
        
        # ScrollViewer erstellen
        $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
        $scrollViewer.VerticalScrollBarVisibility = "Auto"
        $scrollViewer.HorizontalScrollBarVisibility = "Disabled"
        $scrollViewer.Padding = New-Object System.Windows.Thickness(0)
        $scrollViewer.Margin = New-Object System.Windows.Thickness(0)
        $mainGrid.Children.Add($scrollViewer)
        [System.Windows.Controls.Grid]::SetRow($scrollViewer, 0)
        
        # StackPanel für den Inhalt im ScrollViewer
        $contentStackPanel = New-Object System.Windows.Controls.StackPanel
        $contentStackPanel.Margin = New-Object System.Windows.Thickness(20)
        $scrollViewer.Content = $contentStackPanel
        
        # Titel hinzufügen
        $titleBlock = New-Object System.Windows.Controls.TextBlock
        $titleBlock.Text = "Programmeinstellungen"
        $titleBlock.FontSize = 20
        $titleBlock.FontWeight = "Bold"
        $titleBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 20)
        $contentStackPanel.Children.Add($titleBlock)
        
        # ---- ALLGEMEINE EINSTELLUNGEN BEREICH ----
        $generalGroupBox = New-Object System.Windows.Controls.GroupBox
        $generalGroupBox.Header = "Allgemeine Einstellungen"
        $generalGroupBox.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
        $generalGroupBox.Padding = New-Object System.Windows.Thickness(15, 15, 15, 15)
        $contentStackPanel.Children.Add($generalGroupBox)
        
        $generalStack = New-Object System.Windows.Controls.StackPanel
        $generalStack.Margin = New-Object System.Windows.Thickness(0)
        $generalGroupBox.Content = $generalStack
        
        # Debug-Modus
        $debugPanel = New-Object System.Windows.Controls.DockPanel
        $debugPanel.Margin = New-Object System.Windows.Thickness(0, 5, 0, 10)
        $generalStack.Children.Add($debugPanel)
        
        $debugLabel = New-Object System.Windows.Controls.Label
        $debugLabel.Content = "Debug-Modus aktivieren"
        $debugLabel.VerticalAlignment = "Center"
        $debugLabel.FontWeight = "Normal"
        [System.Windows.Controls.DockPanel]::SetDock($debugLabel, "Left")
        $debugPanel.Children.Add($debugLabel)
        
        $debugCheckBox = New-Object System.Windows.Controls.CheckBox
        $debugCheckBox.IsChecked = ($settings["General"]["Debug"] -eq "1")
        $debugCheckBox.VerticalAlignment = "Center"
        $debugCheckBox.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
        $debugCheckBox.ToolTip = "Aktiviert detaillierte Protokollierung für Fehlerbehebung"
        [System.Windows.Controls.DockPanel]::SetDock($debugCheckBox, "Right")
        $debugPanel.Children.Add($debugCheckBox)
        
        # Dark Mode (deaktiviert)
        $darkModePanel = New-Object System.Windows.Controls.DockPanel
        $darkModePanel.Margin = New-Object System.Windows.Thickness(0, 5, 0, 10)
        $generalStack.Children.Add($darkModePanel)
        
        $darkModeLabel = New-Object System.Windows.Controls.Label
        $darkModeLabel.Content = "Dark Mode aktivieren"
        $darkModeLabel.VerticalAlignment = "Center"
        $darkModeLabel.FontWeight = "Normal"
        [System.Windows.Controls.DockPanel]::SetDock($darkModeLabel, "Left")
        $darkModePanel.Children.Add($darkModeLabel)
        
        $darkModeCheckBox = New-Object System.Windows.Controls.CheckBox
        $darkModeCheckBox.IsChecked = ($settings["General"]["DarkMode"] -eq "1")
        $darkModeCheckBox.VerticalAlignment = "Center"
        $darkModeCheckBox.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
        $darkModeCheckBox.IsEnabled = $false
        $darkModeCheckBox.ToolTip = "Diese Funktion ist noch nicht verfügbar"
        [System.Windows.Controls.DockPanel]::SetDock($darkModeCheckBox, "Right")
        $darkModePanel.Children.Add($darkModeCheckBox)
        
        # Dark Mode Info
        $darkModeInfo = New-Object System.Windows.Controls.TextBlock
        $darkModeInfo.Text = "Hinweis: Dark Mode ist derzeit noch nicht implementiert."
        $darkModeInfo.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Gray)
        $darkModeInfo.FontStyle = "Italic"
        $darkModeInfo.Margin = New-Object System.Windows.Thickness(25, 0, 0, 15)
        $darkModeInfo.TextWrapping = "Wrap"
        $generalStack.Children.Add($darkModeInfo)
        
        # ---- DATEIPFAD EINSTELLUNGEN BEREICH ----
        $pathsGroupBox = New-Object System.Windows.Controls.GroupBox
        $pathsGroupBox.Header = "Dateipfade"
        $pathsGroupBox.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
        $pathsGroupBox.Padding = New-Object System.Windows.Thickness(15, 15, 15, 15)
        $contentStackPanel.Children.Add($pathsGroupBox)
        
        $pathsStack = New-Object System.Windows.Controls.StackPanel
        $pathsStack.Margin = New-Object System.Windows.Thickness(0)
        $pathsGroupBox.Content = $pathsStack
        
        # Log-Pfad
        $logPathLabel = New-Object System.Windows.Controls.Label
        $logPathLabel.Content = "Protokollverzeichnis:"
        $logPathLabel.Margin = New-Object System.Windows.Thickness(0, 5, 0, 5)
        $pathsStack.Children.Add($logPathLabel)
        
        $logPathPanel = New-Object System.Windows.Controls.Grid
        $logPathPanel.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
        $pathsStack.Children.Add($logPathPanel)
        
        # Spalten definieren für das Log-Pfad-Grid
        $logPathPanel.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
        $logPathPanel.ColumnDefinitions[0].Width = New-Object System.Windows.GridLength 1, "Star"
        $logPathPanel.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
        $logPathPanel.ColumnDefinitions[1].Width = New-Object System.Windows.GridLength 80
        
        $logPathTextBox = New-Object System.Windows.Controls.TextBox
        $logPathTextBox.Text = $settings["Paths"]["LogPath"]
        $logPathTextBox.VerticalAlignment = "Center"
        $logPathTextBox.Height = 30
        $logPathTextBox.Padding = New-Object System.Windows.Thickness(5, 5, 5, 5)
        $logPathTextBox.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
        $logPathTextBox.ToolTip = "Pfad zum Verzeichnis, in dem Protokolldateien gespeichert werden"
        $logPathPanel.Children.Add($logPathTextBox)
        [System.Windows.Controls.Grid]::SetColumn($logPathTextBox, 0)
        
        $browseButton = New-Object System.Windows.Controls.Button
        $browseButton.Content = "Durchsuchen"
        $browseButton.Height = 30
        $browseButton.Padding = New-Object System.Windows.Thickness(5, 0, 5, 0)
        $logPathPanel.Children.Add($browseButton)
        [System.Windows.Controls.Grid]::SetColumn($browseButton, 1)
        
        # Log-Pfad Info
        $logPathInfo = New-Object System.Windows.Controls.TextBlock
        $logPathInfo.Text = "Hier wird das Verzeichnis festgelegt, in dem alle Protokolldateien gespeichert werden. Stellen Sie sicher, dass genügend Speicherplatz vorhanden ist."
        $logPathInfo.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Gray)
        $logPathInfo.TextWrapping = "Wrap"
        $logPathInfo.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
        $pathsStack.Children.Add($logPathInfo)
        

        
        # ---- BUTTON-BEREICH ----
        $buttonPanel = New-Object System.Windows.Controls.Border
        $buttonPanel.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::WhiteSmoke)
        $buttonPanel.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::LightGray)
        $buttonPanel.BorderThickness = New-Object System.Windows.Thickness(0, 1, 0, 0)
        $buttonPanel.Padding = New-Object System.Windows.Thickness(20, 0, 20, 0)
        $mainGrid.Children.Add($buttonPanel)
        [System.Windows.Controls.Grid]::SetRow($buttonPanel, 1)
        
        $buttonStack = New-Object System.Windows.Controls.StackPanel
        $buttonStack.Orientation = "Horizontal"
        $buttonStack.HorizontalAlignment = "Right"
        $buttonStack.VerticalAlignment = "Center"
        $buttonPanel.Child = $buttonStack
        
        $saveButton = New-Object System.Windows.Controls.Button
        $saveButton.Content = "Speichern"
        $saveButton.Width = 120
        $saveButton.Height = 35
        $saveButton.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
        $saveButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::DodgerBlue)
        $saveButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::White)
        $saveButton.FontWeight = "SemiBold"
        $buttonStack.Children.Add($saveButton)
        
        $cancelButton = New-Object System.Windows.Controls.Button
        $cancelButton.Content = "Abbrechen"
        $cancelButton.Width = 120
        $cancelButton.Height = 35
        $cancelButton.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
        $buttonStack.Children.Add($cancelButton)
        
        # Event-Handler für Browse-Button
        $browseButton.Add_Click({
            $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderBrowser.Description = "Log-Verzeichnis auswählen"
            $folderBrowser.SelectedPath = $logPathTextBox.Text
            
            if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $logPathTextBox.Text = $folderBrowser.SelectedPath
            }
        })
        
        # Event-Handler für den Speichern-Button
        $saveButton.Add_Click({
            try {
            # Einstellungen aktualisieren
            $settings["General"]["Debug"] = if ($debugCheckBox.IsChecked) { "1" } else { "0" }
            $settings["General"]["DarkMode"] = if ($darkModeCheckBox.IsChecked) { "1" } else { "0" }
            $settings["Paths"]["LogPath"] = $logPathTextBox.Text
            
            # Einstellungen speichern
            $iniContent = "[General]`n"
            foreach ($key in $settings["General"].Keys) {
                $iniContent += "$key = $($settings["General"][$key])`n"
            }
            
            $iniContent += "`n[Paths]`n"
            foreach ($key in $settings["Paths"].Keys) {
                $iniContent += "$key = $($settings["Paths"][$key])`n"
            }
            
            if ($settings.ContainsKey("UI")) {
                $iniContent += "`n[UI]`n"
                foreach ($key in $settings["UI"].Keys) {
                $iniContent += "$key = $($settings["UI"][$key])`n"
                }
            }
            
            # Datei schreiben
            Set-Content -Path $script:configFilePath -Value $iniContent -Encoding UTF8
            
            # Debug-Modus aktualisieren
            $script:debugMode = ($settings["General"]["Debug"] -eq "1")
            
            # Logpfad aktualisieren
            $script:logFilePath = Join-Path -Path $settings["Paths"]["LogPath"] -ChildPath "ExchangeTool.log"
            
            Write-DebugMessage "Einstellungen gespeichert" -Type "Success"
            Log-Action "Einstellungen wurden aktualisiert"
            
            [System.Windows.MessageBox]::Show("Einstellungen wurden gespeichert.", "Erfolg", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            $settingsWindow.Close()
            }
            catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Speichern der Einstellungen: $errorMsg" -Type "Error"
            [System.Windows.MessageBox]::Show("Fehler beim Speichern der Einstellungen: $errorMsg", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })
        
        # Event-Handler für den Abbrechen-Button
        $cancelButton.Add_Click({
            $settingsWindow.Close()
        })
        
        # Fenster anzeigen
        [void]$settingsWindow.ShowDialog()
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Anzeigen des Einstellungsdialogs: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Anzeigen des Einstellungsdialogs: $errorMsg"
        return $false
    }
}

# Füge den Event-Handler für den Settings-Button hinzu
if ($null -ne $btnSettings) {
    $btnSettings.Add_Click({
        Show-SettingsDialog
    })
}

# Funktion zum Anzeigen des Hilfe-Menüs
function Show-HelpMenu {
    [CmdletBinding()]
    param()
    
    try {
        # Erstelle ein Kontextmenü für die Hilfe
        $contextMenu = New-Object System.Windows.Controls.ContextMenu
        
        # Füge Menüpunkte hinzu
        $menuItemGeneral = New-Object System.Windows.Controls.MenuItem
        $menuItemGeneral.Header = "Allgemeine Hilfe"
        $menuItemGeneral.Add_Click({ Show-HelpDialog -Topic "General" })
        $contextMenu.Items.Add($menuItemGeneral)
        
        $menuItemCalendar = New-Object System.Windows.Controls.MenuItem
        $menuItemCalendar.Header = "Kalenderberechtigungen"
        $menuItemCalendar.Add_Click({ Show-HelpDialog -Topic "Calendar" })
        $contextMenu.Items.Add($menuItemCalendar)
        
        $menuItemMailbox = New-Object System.Windows.Controls.MenuItem
        $menuItemMailbox.Header = "Postfachberechtigungen"
        $menuItemMailbox.Add_Click({ Show-HelpDialog -Topic "Mailbox" })
        $contextMenu.Items.Add($menuItemMailbox)
        
        $menuItemAudit = New-Object System.Windows.Controls.MenuItem
        $menuItemAudit.Header = "Postfach-Audit"
        $menuItemAudit.Add_Click({ Show-HelpDialog -Topic "Audit" })
        $contextMenu.Items.Add($menuItemAudit)
        
        $menuItemTroubleshooting = New-Object System.Windows.Controls.MenuItem
        $menuItemTroubleshooting.Header = "Troubleshooting"
        $menuItemTroubleshooting.Add_Click({ Show-HelpDialog -Topic "Troubleshooting" })
        $contextMenu.Items.Add($menuItemTroubleshooting)
        
        $menuItemAbout = New-Object System.Windows.Controls.MenuItem
        $menuItemAbout.Header = "Über"
        $menuItemAbout.Add_Click({ 
            [System.Windows.MessageBox]::Show(
                "Exchange Berechtigungen Verwaltung Tool v$($script:config["General"]["Version"])`n`nEntwickelt für die einfache Verwaltung von Exchange Online-Berechtigungen.", 
                "Über Exchange Tool", 
                [System.Windows.MessageBoxButton]::OK, 
                [System.Windows.MessageBoxImage]::Information
            )
        })
        $contextMenu.Items.Add($menuItemAbout)
        
        # Zeige das Kontextmenü
        $contextMenu.IsOpen = $true
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Anzeigen des Hilfe-Menüs: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Anzeigen des Hilfe-Menüs: $errorMsg"
        return $false
    }
}

# Füge den Event-Handler für den Help-Button hinzu
if ($null -ne $btnInfo) {
    $btnInfo.Add_Click({
        Show-HelpMenu
    })
}

# Funktion zum Initialisieren der Hilfe-Links in den Tabs
function Initialize-HelpLinks {
    [CmdletBinding()]
    param()
    
    try {
        # Prüfe, ob die Form-Variable existiert und nicht null ist
        if (-not (Get-Variable -Name Form -ErrorAction SilentlyContinue) -or $null -eq $Form) {
            Write-DebugMessage "Formular-Variable ist nicht definiert oder null" -Type "Warning"
            return $false
        }
        
        # Hole alle Hilfe-Links im Formular (durch Namenskonvention)
        $helpLinks = $Form.FindName("HelpLinks")
        
        if ($null -ne $helpLinks) {
            # Event-Handler für alle Hilfe-Links hinzufügen
            foreach ($link in $helpLinks.Children) {
                if ($link -is [System.Windows.Controls.TextBlock] -and $link.Name -like "helpLink*") {
                    $link.Add_MouseLeftButtonDown({
                        $linkName = $this.Name
                        $topic = "General"
                        
                        # Bestimme das Hilfe-Thema basierend auf dem Link-Namen
                        switch -Wildcard ($linkName) {
                            "*Calendar*" { $topic = "Calendar" }
                            "*Mailbox*" { $topic = "Mailbox" }
                            "*Audit*" { $topic = "Audit" }
                            "*Trouble*" { $topic = "Troubleshooting" }
                        }
                        
                        Show-HelpDialog -Topic $topic
                    })
                    
                    # Mauszeiger ändern, wenn auf den Link gezeigt wird
                    $link.Add_MouseEnter({
                        $this.Cursor = [System.Windows.Input.Cursors]::Hand
                        $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                    })
                    
                    $link.Add_MouseLeave({
                        $this.TextDecorations = $null
                    })
                }
            }
        }
        else {
            Write-DebugMessage "Keine Hilfe-Links im Formular gefunden (HelpLinks-Element nicht gefunden)" -Type "Warning"
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Initialisieren der Hilfe-Links: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Initialisieren der Hilfe-Links: $errorMsg"
        return $false
    }
}

# -------------------------------------------------
# Abschnitt: GUI Design (WPF/XAML) und Initialisierung
# -------------------------------------------------

function Load-XAML {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$XamlFilePath
    )
    
    try {
        # Überprüfen, ob die XAML-Datei existiert
        if (-not (Test-Path -Path $XamlFilePath)) {
            throw "XAML-Datei nicht gefunden: $XamlFilePath"
        }
        
        # XAML-Datei laden
        [xml]$xamlContent = Get-Content -Path $XamlFilePath -Encoding UTF8
        
        # XML-Namespace hinzufügen für WPF
        $xamlContent.SelectNodes("//*[@*[contains(name(), 'Canvas.')]]") | ForEach-Object {
            $_.RemoveAttribute("Canvas.Left")
            $_.RemoveAttribute("Canvas.Top")
        }
        
        # Parse XAML
        $reader = New-Object System.Xml.XmlNodeReader $xamlContent
        $window = [System.Windows.Markup.XamlReader]::Load($reader)
        
        # Erfolg loggen
        Write-DebugMessage "XAML-GUI erfolgreich geladen: $XamlFilePath" -Type "Success"
        Log-Action "XAML-GUI erfolgreich geladen"
        
        return $window
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Laden der XAML-Datei: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Laden der XAML-Datei: $errorMsg"
        throw "Fehler beim Laden der XAML-GUI: $errorMsg"
    }
}

# -------------------------------------------------
# Abschnitt: GUI Laden
# -------------------------------------------------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Pfad zur XAML-Datei
$script:xamlFilePath = Join-Path -Path $PSScriptRoot -ChildPath "EXOGUI.xaml"

# Fallback-Pfad, falls die Datei nicht im Hauptverzeichnis ist
if (-not (Test-Path -Path $script:xamlFilePath)) {
    $script:xamlFilePath = Join-Path -Path $PSScriptRoot -ChildPath "assets\EXOGUI.xaml"
    Write-DebugMessage "Primärer XAML-Pfad nicht gefunden, versuche Fallback-Pfad: $script:xamlFilePath" -Type "Warning"
}

# Prüfen, ob XAML-Datei gefunden wurde
if (-not (Test-Path -Path $script:xamlFilePath)) {
    Write-Host "KRITISCHER FEHLER: XAML-Datei nicht gefunden an beiden Standardpfaden!" -ForegroundColor Red
    Write-Host "Gesucht wurde in: $PSScriptRoot und $PSScriptRoot\assets" -ForegroundColor Red
    
    # Versuche eine temporäre minimal-GUI zu erstellen
    $minimalXaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Exchange Tool - ERROR" Width="600" Height="300">
    <Grid>
        <StackPanel Margin="20">
            <TextBlock FontSize="18" Foreground="Red" FontWeight="Bold" TextWrapping="Wrap">
                FEHLER: Die XAML-GUI-Datei 'EXOGUI.xaml' wurde nicht gefunden!
            </TextBlock>
            <TextBlock Margin="0,20,0,0" TextWrapping="Wrap">
                Bitte stellen Sie sicher, dass die Datei 'EXOGUI.xaml' im gleichen Verzeichnis wie dieses Skript oder im Unterverzeichnis 'assets' vorhanden ist.
            </TextBlock>
            <TextBlock Margin="0,20,0,0" TextWrapping="Wrap">
                Gesucht wurde in:
            </TextBlock>
            <TextBlock Margin="20,5,0,0">
                • $PSScriptRoot\EXOGUI.xaml
            </TextBlock>
            <TextBlock Margin="20,5,0,0">
                • $PSScriptRoot\assets\EXOGUI.xaml
            </TextBlock>
            <Button Content="OK" Width="100" Height="30" Margin="0,20,0,0" HorizontalAlignment="Center" x:Name="btnOk"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    try {
        $tempXamlPath = [System.IO.Path]::GetTempFileName() + ".xaml"
        Set-Content -Path $tempXamlPath -Value $minimalXaml -Encoding UTF8
        
        $script:xamlFilePath = $tempXamlPath
        Write-DebugMessage "Erstelle Notfall-GUI: $tempXamlPath" -Type "Warning"
    }
    catch {
        Write-Host "Konnte keine Notfall-GUI erstellen. Das Programm wird beendet." -ForegroundColor Red
        exit
    }
}

try {
    # GUI aus externer XAML-Datei laden
    $script:Form = Load-XAML -XamlFilePath $script:xamlFilePath
    
    # Wichtig: Ab hier können GUI-Elemente referenziert werden
    Write-DebugMessage "Hauptfenster erfolgreich geladen, fahre mit GUI-Initialisierung fort" -Type "Info"
    
    # -------------------------------------------------
    # Abschnitt: GUI-Elemente referenzieren
    # -------------------------------------------------
    function Get-XamlElement {
        param (
            [Parameter(Mandatory = $true)]
            [string]$ElementName,
            
            [Parameter(Mandatory = $false)]
            [switch]$Required = $false
        )
        
        try {
            $element = $script:Form.FindName($ElementName)
            if ($null -eq $element -and $Required.IsPresent) {
                Write-DebugMessage "Element nicht gefunden (ERFORDERLICH): $ElementName" -Type "Error"
                throw "Erforderliches Element nicht gefunden: $ElementName"
            }
            elseif ($null -eq $element) {
                Write-DebugMessage "Element nicht gefunden: $ElementName" -Type "Warning"
            }
            return $element
        }
        catch {
            Write-DebugMessage "Fehler beim Suchen des Elements $ElementName : $($_.Exception.Message)" -Type "Error"
            if ($Required.IsPresent) { throw }
            return $null
        }
    }
    
    # Hauptelemente referenzieren
    $script:btnConnect          = Get-XamlElement -ElementName "btnConnect" -Required
    $script:tabContent          = Get-XamlElement -ElementName "tabContent" -Required
    $script:tabCalendar         = Get-XamlElement -ElementName "tabCalendar"
    $script:tabMailbox          = Get-XamlElement -ElementName "tabMailbox"
    $script:tabMailboxAudit     = Get-XamlElement -ElementName "tabMailboxAudit"
    $script:tabTroubleshooting  = Get-XamlElement -ElementName "tabTroubleshooting"
    $script:txtStatus           = Get-XamlElement -ElementName "txtStatus" -Required
    $script:txtVersion          = Get-XamlElement -ElementName "txtVersion"
    $script:txtConnectionStatus = Get-XamlElement -ElementName "txtConnectionStatus" -Required
    
    # Referenzierung der Navigationselemente
    $script:btnNavCalendar        = Get-XamlElement -ElementName "btnNavCalendar"
    $script:btnNavMailbox         = Get-XamlElement -ElementName "btnNavMailbox"
    $script:btnNavGroups          = Get-XamlElement -ElementName "btnNavGroups"
    $script:btnNavSharedMailbox   = Get-XamlElement -ElementName "btnNavSharedMailbox"
    $script:btnNavAudit           = Get-XamlElement -ElementName "btnNavAudit"
    $script:btnNavTroubleshooting = Get-XamlElement -ElementName "btnNavTroubleshooting"
    $script:btnInfo               = Get-XamlElement -ElementName "btnInfo"
    $script:btnSettings           = Get-XamlElement -ElementName "btnSettings"
    $script:btnClose              = Get-XamlElement -ElementName "btnClose" -Required
    
    # Referenzierung weiterer wichtiger UI-Elemente
    $script:btnCheckPrerequisites   = Get-XamlElement -ElementName "btnCheckPrerequisites"
    
    # Funktion zum sicheren Hinzufügen eines Event-Handlers
    function Register-EventHandler {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$false)]
            $Control,
            
            [Parameter(Mandatory=$true)]
            $Handler,
            
            [Parameter(Mandatory=$false)]
            [string]$ControlName = "UnknownControl",
            
            [Parameter(Mandatory=$false)]
            [string]$EventName = "Click"
        )
        
        if ($null -eq $Control) {
            Write-DebugMessage "Control nicht gefunden: $ControlName" -Type "Warning"
            return $false
        }
        
        try {
            # Event-Handler hinzufügen
            $event = "Add_$EventName"
            $Control.$event($Handler)
            Write-DebugMessage "Event-Handler für $ControlName.$EventName registriert" -Type "Info"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Registrieren des Event-Handlers für $ControlName - $errorMsg" -Type "Error"
            return $false
        }
    }
    
    # -------------------------------------------------
    # Abschnitt: Event-Handler registrieren
    # -------------------------------------------------
    
    # Event-Handler für Connect-Button
    if ($null -ne $script:btnConnect) {
        $script:btnConnect.Add_Click({
            try {
                Write-DebugMessage "Connect-Button wurde geklickt" -Type "Info"
                if (-not $script:isConnected) {
                    # Verbindung herstellen
                    $success = Connect-ExchangeOnlineSession
                    if ($success) {
                        Write-DebugMessage "Verbindung zu Exchange Online hergestellt" -Type "Success"
                        if ($null -ne $script:txtStatus) {
                            $script:txtStatus.Text = "Mit Exchange Online verbunden"
                            $script:txtStatus.Foreground = $script:connectedBrush
                        }
                    }
                }
                else {
                    # Verbindung trennen
                    Disconnect-ExchangeOnlineSession
                    Write-DebugMessage "Verbindung zu Exchange Online getrennt" -Type "Info"
                    if ($null -ne $script:txtStatus) {
                        $script:txtStatus.Text = "Verbindung zu Exchange Online getrennt"
                    }
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-DebugMessage "Fehler beim Verbinden/Trennen: $errorMsg" -Type "Error"
                if ($null -ne $script:txtStatus) {
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            }
        })
        Write-DebugMessage "Event-Handler für Connect-Button registriert" -Type "Info"
    }
    else {
        Write-DebugMessage "Connect-Button nicht gefunden!" -Type "Warning"
    }
    
    # Event-Handler für Navigation
    Register-EventHandler -Control $script:btnNavCalendar -Handler {
        if ($null -ne $script:tabContent) {
            $script:tabContent.SelectedIndex = 0
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Text = "Kalenderberechtigungen-Tab ausgewählt"
            }
        }
    } -ControlName "btnNavCalendar"
    
    Register-EventHandler -Control $script:btnNavMailbox -Handler {
        if ($null -ne $script:tabContent) {
            $script:tabContent.SelectedIndex = 1
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Text = "Postfachberechtigungen-Tab ausgewählt"
            }
        }
    } -ControlName "btnNavMailbox"
    
    Register-EventHandler -Control $script:btnNavGroups -Handler {
        if ($null -ne $script:tabContent) {
            $script:tabContent.SelectedIndex = 4
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Text = "Gruppen/Verteiler-Tab ausgewählt"
            }
        }
    } -ControlName "btnNavGroups"
    
    Register-EventHandler -Control $script:btnNavSharedMailbox -Handler {
        if ($null -ne $script:tabContent) {
            $script:tabContent.SelectedIndex = 5
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Text = "Shared Mailbox-Tab ausgewählt"
            }
        }
    } -ControlName "btnNavSharedMailbox"
    
    # Handler für Audit-Navigation
    $auditHandler = {
        if ($null -ne $script:tabContent) {
            $script:tabContent.SelectedIndex = 2
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Text = "Audit-Tab ausgewählt"
            }
        }
    }
    
    # Audit-Navigation-Buttons
    Register-EventHandler -Control $script:btnNavAudit -Handler $auditHandler -ControlName "btnNavAudit"
    
    # Troubleshooting-Navigation
    Register-EventHandler -Control $script:btnNavTroubleshooting -Handler {
        if ($null -ne $script:tabContent) {
            $script:tabContent.SelectedIndex = 3
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Text = "Troubleshooting-Tab ausgewählt"
            }
        }
    } -ControlName "btnNavTroubleshooting"
    
    Register-EventHandler -Control $script:btnCheckPrerequisites -Handler {
        Check-Prerequisites
    } -ControlName "btnCheckPrerequisites"
    
    # Schließen-Button
    Register-EventHandler -Control $script:btnClose -Handler {
        $script:Form.Close()
    } -ControlName "btnClose"
    
    # Settings-Button
    Register-EventHandler -Control $script:btnSettings -Handler {
        Show-SettingsDialog
    } -ControlName "btnSettings"
    
    # Info-Button
    Register-EventHandler -Control $script:btnInfo -Handler {
        Show-HelpMenu
    } -ControlName "btnInfo"
    
    # -------------------------------------------------
    # Abschnitt: UI Initialisierung und Element-Referenzierung
    # -------------------------------------------------
    function Initialize-CalendarTab {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Kalender-Tab" -Type "Info"
            
            # Referenzieren der UI-Elemente im Kalender-Tab
            $txtCalendarSource = Get-XamlElement -ElementName "txtCalendarSource"
            $txtCalendarTarget = Get-XamlElement -ElementName "txtCalendarTarget"
            $cmbCalendarPermission = Get-XamlElement -ElementName "cmbCalendarPermission"
            $btnAddCalendarPermission = Get-XamlElement -ElementName "btnAddCalendarPermission"
            $btnRemoveCalendarPermission = Get-XamlElement -ElementName "btnRemoveCalendarPermission"
            $btnShowCalendarPermissions = Get-XamlElement -ElementName "btnShowCalendarPermissions"
            $lstCalendarPermissions = Get-XamlElement -ElementName "lstCalendarPermissions"
            $txtCalendarMailboxUser = Get-XamlElement -ElementName "txtCalendarMailboxUser"
            $cmbDefaultPermission = Get-XamlElement -ElementName "cmbDefaultPermission"
            $btnSetDefaultPermission = Get-XamlElement -ElementName "btnSetDefaultPermission"
            $btnSetAnonymousPermission = Get-XamlElement -ElementName "btnSetAnonymousPermission"
            $btnSetAllCalPermission = Get-XamlElement -ElementName "btnSetAllCalPermission"
            $helpLinkCalendar = Get-XamlElement -ElementName "helpLinkCalendar"

            # Globale Variablen für spätere Verwendung setzen
            $script:txtCalendarSource = $txtCalendarSource
            $script:txtCalendarTarget = $txtCalendarTarget
            $script:cmbCalendarPermission = $cmbCalendarPermission
            $script:lstCalendarPermissions = $lstCalendarPermissions
            $script:txtCalendarMailboxUser = $txtCalendarMailboxUser
            $script:cmbDefaultPermission = $cmbDefaultPermission
            
            # Event-Handler für Kalender-Buttons registrieren
            Register-EventHandler -Control $btnAddCalendarPermission -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtCalendarSource.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtCalendarTarget.Text) -or
                        $null -eq $script:cmbCalendarPermission.SelectedItem) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte füllen Sie alle erforderlichen Felder aus.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Parameter sammeln und Funktion aufrufen
                    $sourceUser = $script:txtCalendarSource.Text
                    $targetUser = $script:txtCalendarTarget.Text
                    $permission = $script:cmbCalendarPermission.SelectedItem.Content
                    
                    Write-DebugMessage "Füge Kalenderberechtigung hinzu: $sourceUser -> $targetUser ($permission)" -Type "Info"
                    $result = Add-CalendarPermission -SourceUser $sourceUser -TargetUser $targetUser -Permission $permission
                    
                    if ($result) {
                        $script:txtStatus.Text = "Kalenderberechtigung erfolgreich hinzugefügt/aktualisiert."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Hinzufügen der Kalenderberechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnAddCalendarPermission"
            
            Register-EventHandler -Control $btnRemoveCalendarPermission -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtCalendarSource.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtCalendarTarget.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie Quell- und Zielbenutzer an.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Parameter sammeln und Funktion aufrufen
                    $sourceUser = $script:txtCalendarSource.Text
                    $targetUser = $script:txtCalendarTarget.Text
                    
                    Write-DebugMessage "Entferne Kalenderberechtigung: $sourceUser -> $targetUser" -Type "Info"
                    $result = Remove-CalendarPermission -SourceUser $sourceUser -TargetUser $targetUser
                    
                    if ($result) {
                        $script:txtStatus.Text = "Kalenderberechtigung erfolgreich entfernt."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Entfernen der Kalenderberechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnRemoveCalendarPermission"
            
            Register-EventHandler -Control $btnShowCalendarPermissions -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtCalendarMailboxUser.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie eine E-Mail-Adresse ein.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $mailboxUser = $script:txtCalendarMailboxUser.Text
                    
                    Write-DebugMessage "Zeige Kalenderberechtigungen für: $mailboxUser" -Type "Info"
                    $permissions = Get-CalendarPermission -MailboxUser $mailboxUser
                    
                    # ListView leeren und mit neuen Daten füllen
                    if ($null -ne $script:lstCalendarPermissions) {
                        $script:lstCalendarPermissions.Items.Clear()
                        
                        foreach ($perm in $permissions) {
                            $item = New-Object PSObject -Property @{
                                User = $perm.User.ToString()
                                AccessRights = $perm.AccessRights -join ", "
                                SharingPermissionFlags = $perm.SharingPermissionFlags
                            }
                            [void]$script:lstCalendarPermissions.Items.Add($item)
                        }
                        
                        $script:txtStatus.Text = "Kalenderberechtigungen erfolgreich abgerufen."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Abrufen der Kalenderberechtigungen: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnShowCalendarPermissions"
            
            Register-EventHandler -Control $btnSetDefaultPermission -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtCalendarMailboxUser.Text) -or
                        $null -eq $script:cmbDefaultPermission.SelectedItem) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte wählen Sie einen Benutzer und eine Berechtigung aus.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $mailboxUser = $script:txtCalendarMailboxUser.Text
                    $permission = $script:cmbDefaultPermission.SelectedItem.Content
                    
                    Write-DebugMessage "Setze Standard-Kalenderberechtigung für $mailboxUser auf $permission" -Type "Info"
                    Set-CalendarDefaultPermissionsAction -PermissionType "Standard" -AccessRights $permission
                    
                    $script:txtStatus.Text = "Standard-Kalenderberechtigung erfolgreich gesetzt."
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Setzen der Standard-Kalenderberechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnSetDefaultPermission"
            
            Register-EventHandler -Control $btnSetAnonymousPermission -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtCalendarMailboxUser.Text) -or
                        $null -eq $script:cmbDefaultPermission.SelectedItem) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte wählen Sie einen Benutzer und eine Berechtigung aus.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $mailboxUser = $script:txtCalendarMailboxUser.Text
                    $permission = $script:cmbDefaultPermission.SelectedItem.Content
                    
                    Write-DebugMessage "Setze Anonymous-Kalenderberechtigung für $mailboxUser auf $permission" -Type "Info"
                    Set-CalendarDefaultPermissionsAction -PermissionType "Anonym" -AccessRights $permission
                    
                    $script:txtStatus.Text = "Anonymous-Kalenderberechtigung erfolgreich gesetzt."
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Setzen der Anonymous-Kalenderberechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnSetAnonymousPermission"
            
            Register-EventHandler -Control $btnSetAllCalPermission -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ($null -eq $script:cmbDefaultPermission.SelectedItem) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte wählen Sie eine Berechtigung aus.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Sicherheitsabfrage
                    $confirmation = [System.Windows.MessageBox]::Show(
                        "Möchten Sie wirklich die Standard-Kalenderberechtigungen für ALLE Postfächer ändern? Diese Aktion kann lange dauern.",
                        "Bestätigung",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Question
                    )
                    
                    if ($confirmation -eq [System.Windows.MessageBoxResult]::No) {
                        return
                    }
                    
                    $permission = $script:cmbDefaultPermission.SelectedItem.Content
                    
                    Write-DebugMessage "Setze Standard-Kalenderberechtigung für ALLE Postfächer auf $permission" -Type "Info"
                    Set-CalendarDefaultPermissionsAction -PermissionType "Beides" -AccessRights $permission -ForAllMailboxes
                    
                    $script:txtStatus.Text = "Standard-Kalenderberechtigungen für alle Postfächer werden gesetzt..."
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Setzen der Standard-Kalenderberechtigungen für alle Postfächer: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnSetAllCalPermission"
            
            if ($null -ne $helpLinkCalendar) {
                $helpLinkCalendar.Add_MouseLeftButtonDown({
                    Show-HelpDialog -Topic "Calendar"
                })
                
                # Mauszeiger ändern, wenn auf den Link gezeigt wird
                $helpLinkCalendar.Add_MouseEnter({
                    $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                    $this.Cursor = [System.Windows.Input.Cursors]::Hand
                })
                
                $helpLinkCalendar.Add_MouseLeave({
                    $this.TextDecorations = $null
                    $this.Cursor = [System.Windows.Input.Cursors]::Arrow
                })
            }
            
            # Initialisierung der ComboBoxen
            if ($null -ne $cmbCalendarPermission) {
                $calendarPermissions = @("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author", "Reviewer", "AvailabilityOnly", "None")
                $cmbCalendarPermission.Items.Clear()
                foreach ($perm in $calendarPermissions) {
                    $item = New-Object System.Windows.Controls.ComboBoxItem
                    $item.Content = $perm
                    [void]$cmbCalendarPermission.Items.Add($item)
                }
                if ($cmbCalendarPermission.Items.Count -gt 0) {
                    $cmbCalendarPermission.SelectedIndex = 0
                }
            }
            
            if ($null -ne $cmbDefaultPermission) {
                $defaultPermissions = @("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author", "Reviewer", "AvailabilityOnly", "None")
                $cmbDefaultPermission.Items.Clear()
                foreach ($perm in $defaultPermissions) {
                    $item = New-Object System.Windows.Controls.ComboBoxItem
                    $item.Content = $perm
                    [void]$cmbDefaultPermission.Items.Add($item)
                }
                if ($cmbDefaultPermission.Items.Count -gt 0) {
                    $cmbDefaultPermission.SelectedIndex = 5  # Default auf Reviewer
                }
            }
            
            Write-DebugMessage "Kalender-Tab erfolgreich initialisiert" -Type "Success"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Initialisieren des Kalender-Tabs: $errorMsg" -Type "Error"
            return $false
        }
    }

    function Initialize-MailboxTab {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Postfach-Tab" -Type "Info"
            
            # Referenzieren der UI-Elemente im Postfach-Tab
            $txtMailboxSource = Get-XamlElement -ElementName "txtMailboxSource"
            $txtMailboxTarget = Get-XamlElement -ElementName "txtMailboxTarget"
            $btnAddMailboxPermission = Get-XamlElement -ElementName "btnAddMailboxPermission"
            $btnRemoveMailboxPermission = Get-XamlElement -ElementName "btnRemoveMailboxPermission"
            $btnShowMailboxPermissions = Get-XamlElement -ElementName "btnShowMailboxPermissions"
            $btnAddSendAs = Get-XamlElement -ElementName "btnAddSendAs"
            $btnRemoveSendAs = Get-XamlElement -ElementName "btnRemoveSendAs"
            $btnShowSendAs = Get-XamlElement -ElementName "btnShowSendAs"
            $btnAddSendOnBehalf = Get-XamlElement -ElementName "btnAddSendOnBehalf"
            $btnRemoveSendOnBehalf = Get-XamlElement -ElementName "btnRemoveSendOnBehalf"
            $btnShowSendOnBehalf = Get-XamlElement -ElementName "btnShowSendOnBehalf"
            $txtMailboxUser = Get-XamlElement -ElementName "txtMailboxUser"
            $lstMailboxPermissions = Get-XamlElement -ElementName "lstMailboxPermissions"
            $helpLinkMailbox = Get-XamlElement -ElementName "helpLinkMailbox"
            
            # Globale Variablen für spätere Verwendung setzen
            $script:txtMailboxSource = $txtMailboxSource
            $script:txtMailboxTarget = $txtMailboxTarget
            $script:txtMailboxUser = $txtMailboxUser
            $script:lstMailboxPermissions = $lstMailboxPermissions
            
            # Event-Handler für Postfach-Buttons registrieren
            Register-EventHandler -Control $btnAddMailboxPermission -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtMailboxSource.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtMailboxTarget.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie Quell- und Zielbenutzer an.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $sourceUser = $script:txtMailboxSource.Text
                    $targetUser = $script:txtMailboxTarget.Text
                    
                    Write-DebugMessage "Füge Postfachberechtigung hinzu: $sourceUser -> $targetUser" -Type "Info"
                    $result = Add-MailboxPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                    
                    if ($result) {
                        $script:txtStatus.Text = "Postfachberechtigung erfolgreich hinzugefügt."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Hinzufügen der Postfachberechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnAddMailboxPermission"
            
            # Event-Handler für Entfernen von Postfachberechtigungen
            Register-EventHandler -Control $btnRemoveMailboxPermission -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtMailboxSource.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtMailboxTarget.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie Quell- und Zielbenutzer an.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $sourceUser = $script:txtMailboxSource.Text
                    $targetUser = $script:txtMailboxTarget.Text
                    
                    Write-DebugMessage "Entferne Postfachberechtigung: $sourceUser -> $targetUser" -Type "Info"
                    $result = Remove-MailboxPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                    
                    if ($result) {
                        $script:txtStatus.Text = "Postfachberechtigung erfolgreich entfernt."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Entfernen der Postfachberechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnRemoveMailboxPermission"
            
            # Event-Handler für Anzeigen von Postfachberechtigungen
            Register-EventHandler -Control $btnShowMailboxPermissions -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtMailboxUser.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie eine E-Mail-Adresse ein.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $mailboxUser = $script:txtMailboxUser.Text
                    
                    Write-DebugMessage "Zeige Postfachberechtigungen für: $mailboxUser" -Type "Info"
                    $permissions = Get-MailboxPermissions -Mailbox $mailboxUser
                    
                    # ListView leeren und mit neuen Daten füllen
                    if ($null -ne $script:lstMailboxPermissions) {
                        $script:lstMailboxPermissions.Items.Clear()
                        
                        foreach ($perm in $permissions) {
                            [void]$script:lstMailboxPermissions.Items.Add($perm)
                        }
                        
                        $script:txtStatus.Text = "Postfachberechtigungen erfolgreich abgerufen."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Abrufen der Postfachberechtigungen: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnShowMailboxPermissions"
            
            # Event-Handler für SendAs-Berechtigungen hinzufügen
            Register-EventHandler -Control $btnAddSendAs -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtMailboxSource.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtMailboxTarget.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie Quell- und Zielbenutzer an.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $sourceUser = $script:txtMailboxSource.Text
                    $targetUser = $script:txtMailboxTarget.Text
                    
                    Write-DebugMessage "Füge SendAs-Berechtigung hinzu: $sourceUser -> $targetUser" -Type "Info"
                    $result = Add-SendAsPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                    
                    if ($result) {
                        $script:txtStatus.Text = "SendAs-Berechtigung erfolgreich hinzugefügt."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Hinzufügen der SendAs-Berechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnAddSendAs"
            
            # Event-Handler für SendAs-Berechtigungen entfernen
            Register-EventHandler -Control $btnRemoveSendAs -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtMailboxSource.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtMailboxTarget.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie Quell- und Zielbenutzer an.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $sourceUser = $script:txtMailboxSource.Text
                    $targetUser = $script:txtMailboxTarget.Text
                    
                    Write-DebugMessage "Entferne SendAs-Berechtigung: $sourceUser -> $targetUser" -Type "Info"
                    $result = Remove-SendAsPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                    
                    if ($result) {
                        $script:txtStatus.Text = "SendAs-Berechtigung erfolgreich entfernt."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Entfernen der SendAs-Berechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnRemoveSendAs"
            
            # Event-Handler für Anzeigen von SendAs-Berechtigungen
            Register-EventHandler -Control $btnShowSendAs -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtMailboxUser.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie eine E-Mail-Adresse ein.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $mailboxUser = $script:txtMailboxUser.Text
                    
                    Write-DebugMessage "Zeige SendAs-Berechtigungen für: $mailboxUser" -Type "Info"
                    $permissions = Get-SendAsPermissionAction -MailboxUser $mailboxUser
                    
                    # ListView leeren und mit neuen Daten füllen
                    if ($null -ne $script:lstMailboxPermissions) {
                        $script:lstMailboxPermissions.Items.Clear()
                        
                        foreach ($perm in $permissions) {
                            [void]$script:lstMailboxPermissions.Items.Add($perm)
                        }
                        
                        $script:txtStatus.Text = "SendAs-Berechtigungen erfolgreich abgerufen."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Abrufen der SendAs-Berechtigungen: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnShowSendAs"
            
            # Event-Handler für SendOnBehalf-Berechtigungen hinzufügen
            Register-EventHandler -Control $btnAddSendOnBehalf -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtMailboxSource.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtMailboxTarget.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie Quell- und Zielbenutzer an.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $sourceUser = $script:txtMailboxSource.Text
                    $targetUser = $script:txtMailboxTarget.Text
                    
                    Write-DebugMessage "Füge SendOnBehalf-Berechtigung hinzu: $sourceUser -> $targetUser" -Type "Info"
                    $result = Add-SendOnBehalfPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                    
                    if ($result) {
                        $script:txtStatus.Text = "SendOnBehalf-Berechtigung erfolgreich hinzugefügt."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Hinzufügen der SendOnBehalf-Berechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnAddSendOnBehalf"
            
            # Event-Handler für SendOnBehalf-Berechtigungen entfernen
            Register-EventHandler -Control $btnRemoveSendOnBehalf -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtMailboxSource.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtMailboxTarget.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie Quell- und Zielbenutzer an.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $sourceUser = $script:txtMailboxSource.Text
                    $targetUser = $script:txtMailboxTarget.Text
                    
                    Write-DebugMessage "Entferne SendOnBehalf-Berechtigung: $sourceUser -> $targetUser" -Type "Info"
                    $result = Remove-SendOnBehalfPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                    
                    if ($result) {
                        $script:txtStatus.Text = "SendOnBehalf-Berechtigung erfolgreich entfernt."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Entfernen der SendOnBehalf-Berechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnRemoveSendOnBehalf"
            
            # Event-Handler für Anzeigen von SendOnBehalf-Berechtigungen
            Register-EventHandler -Control $btnShowSendOnBehalf -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtMailboxUser.Text)) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie eine E-Mail-Adresse ein.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $mailboxUser = $script:txtMailboxUser.Text
                    
                    Write-DebugMessage "Zeige SendOnBehalf-Berechtigungen für: $mailboxUser" -Type "Info"
                    $permissions = Get-SendOnBehalfPermissionAction -MailboxUser $mailboxUser
                    
                    # ListView leeren und mit neuen Daten füllen
                    if ($null -ne $script:lstMailboxPermissions) {
                        $script:lstMailboxPermissions.Items.Clear()
                        
                        foreach ($perm in $permissions) {
                            [void]$script:lstMailboxPermissions.Items.Add($perm)
                        }
                        
                        $script:txtStatus.Text = "SendOnBehalf-Berechtigungen erfolgreich abgerufen."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnShowSendOnBehalf"
            
            # Event-Handler für Hilfe-Link
            if ($null -ne $helpLinkMailbox) {
                $helpLinkMailbox.Add_MouseLeftButtonDown({
                    Show-HelpDialog -Topic "Mailbox"
                })
                
                $helpLinkMailbox.Add_MouseEnter({
                    $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                    $this.Cursor = [System.Windows.Input.Cursors]::Hand
                })
                
                $helpLinkMailbox.Add_MouseLeave({
                    $this.TextDecorations = $null
                    $this.Cursor = [System.Windows.Input.Cursors]::Arrow
                })
            }
            
            Write-DebugMessage "Postfach-Tab erfolgreich initialisiert" -Type "Success"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Initialisieren des Postfach-Tabs: $errorMsg" -Type "Error"
            return $false
        }
    }
    function Initialize-GroupsTab {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Gruppen/Verteiler-Tab" -Type "Info"
            
            # UI-Elemente referenzieren
            $cmbGroupType = Get-XamlElement -ElementName "cmbGroupType"
            $txtGroupName = Get-XamlElement -ElementName "txtGroupName"
            $txtGroupEmail = Get-XamlElement -ElementName "txtGroupEmail"
            $txtGroupMembers = Get-XamlElement -ElementName "txtGroupMembers"
            $txtGroupDescription = Get-XamlElement -ElementName "txtGroupDescription"
            $btnCreateGroup = Get-XamlElement -ElementName "btnCreateGroup"
            $btnDeleteGroup = Get-XamlElement -ElementName "btnDeleteGroup"
            
            $txtExistingGroupName = Get-XamlElement -ElementName "txtExistingGroupName"
            $txtGroupUser = Get-XamlElement -ElementName "txtGroupUser"
            $btnAddUserToGroup = Get-XamlElement -ElementName "btnAddUserToGroup"
            $btnRemoveUserFromGroup = Get-XamlElement -ElementName "btnRemoveUserFromGroup"
            $chkHiddenFromGAL = Get-XamlElement -ElementName "chkHiddenFromGAL"
            $chkRequireSenderAuth = Get-XamlElement -ElementName "chkRequireSenderAuth"
            $chkAllowExternalSenders = Get-XamlElement -ElementName "chkAllowExternalSenders"
            $btnShowGroupMembers = Get-XamlElement -ElementName "btnShowGroupMembers"
            $btnUpdateGroupSettings = Get-XamlElement -ElementName "btnUpdateGroupSettings"
            
            $lstGroupMembers = Get-XamlElement -ElementName "lstGroupMembers"
            $helpLinkGroups = Get-XamlElement -ElementName "helpLinkGroups"
            
            # Globale Variablen setzen
            $script:cmbGroupType = $cmbGroupType
            $script:txtGroupName = $txtGroupName
            $script:txtGroupEmail = $txtGroupEmail
            $script:txtGroupMembers = $txtGroupMembers
            $script:txtGroupDescription = $txtGroupDescription
            $script:txtExistingGroupName = $txtExistingGroupName
            $script:txtGroupUser = $txtGroupUser
            $script:chkHiddenFromGAL = $chkHiddenFromGAL
            $script:chkRequireSenderAuth = $chkRequireSenderAuth
            $script:chkAllowExternalSenders = $chkAllowExternalSenders
            $script:lstGroupMembers = $lstGroupMembers
            
            # ComboBox für Group Type initialisieren
            if ($null -ne $cmbGroupType) {
                $groupTypes = @("Distribution", "Security", "Mail-enabled Security")
                $cmbGroupType.Items.Clear()
                foreach ($type in $groupTypes) {
                    $item = New-Object System.Windows.Controls.ComboBoxItem
                    $item.Content = $type
                    [void]$cmbGroupType.Items.Add($item)
                }
                if ($cmbGroupType.Items.Count -gt 0) {
                    $cmbGroupType.SelectedIndex = 0
                }
            }
            
            # Event-Handler registrieren
            Register-EventHandler -Control $btnCreateGroup -Handler {
                try {
                    # Verbindungsprüfung
                    # Überprüfen, ob GroupType ausgewählt ist
                    if ($null -eq $script:cmbGroupType.SelectedItem) {
                        [System.Windows.MessageBox]::Show("Bitte wählen Sie einen Gruppentyp aus.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $groupType = $script:cmbGroupType.SelectedItem.Content
                    $groupName = $script:txtGroupName.Text
                    $groupEmail = $script:txtGroupEmail.Text
                    $members = $script:txtGroupMembers.Text
                    $description = $script:txtGroupDescription.Text
                    # Eingabeprüfung
                    if ([string]::IsNullOrWhiteSpace($script:txtGroupName.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtGroupEmail.Text)) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie einen Namen und eine E-Mail-Adresse für die Gruppe an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $groupType = $script:cmbGroupType.SelectedItem.Content
                    $groupName = $script:txtGroupName.Text
                    $groupEmail = $script:txtGroupEmail.Text
                    $members = $script:txtGroupMembers.Text
                    $description = $script:txtGroupDescription.Text
                    
                    # Funktion zur Gruppenerstellung aufrufen
                    $result = New-DistributionGroupAction -GroupName $groupName -GroupEmail $groupEmail `
                        -GroupType $groupType -Members $members -Description $description
                    
                    if ($result) {
                        $script:txtStatus.Text = "Gruppe wurde erfolgreich erstellt."
                        
                        # Felder zurücksetzen
                        $script:txtGroupName.Text = ""
                        $script:txtGroupEmail.Text = ""
                        $script:txtGroupMembers.Text = ""
                        $script:txtGroupDescription.Text = ""
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Erstellen der Gruppe: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnCreateGroup"
            
            Register-EventHandler -Control $btnDeleteGroup -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ([string]::IsNullOrWhiteSpace($script:txtGroupName.Text)) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie den Namen der zu löschenden Gruppe an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Sicherheitsabfrage
                    $result = [System.Windows.MessageBox]::Show(
                        "Sind Sie sicher, dass Sie die Gruppe '$($script:txtGroupName.Text)' löschen möchten? Diese Aktion kann nicht rückgängig gemacht werden.",
                        "Gruppe löschen",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Warning)
                    
                    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                        # Funktion zum Löschen der Gruppe aufrufen
                        $result = Remove-DistributionGroupAction -GroupName $script:txtGroupName.Text
                        
                        if ($result) {
                            $script:txtStatus.Text = "Gruppe wurde erfolgreich gelöscht."
                            
                            # Felder zurücksetzen
                            $script:txtGroupName.Text = ""
                            $script:txtGroupEmail.Text = ""
                            $script:txtGroupMembers.Text = ""
                            $script:txtGroupDescription.Text = ""
                        }
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Löschen der Gruppe: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnDeleteGroup"
            
            Register-EventHandler -Control $btnAddUserToGroup -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ([string]::IsNullOrWhiteSpace($script:txtExistingGroupName.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtGroupUser.Text)) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie den Namen der Gruppe und des hinzuzufügenden Benutzers an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }

                    # Funktion zum Hinzufügen eines Benutzers aufrufen
                    $result = Add-GroupMemberAction -GroupName $script:txtExistingGroupName.Text -MemberIdentity $script:txtGroupUser.Text
                    
                    if ($result) {
                        $script:txtStatus.Text = "Benutzer wurde erfolgreich zur Gruppe hinzugefügt."
                        
                        # Benutzerfeld zurücksetzen
                        $script:txtGroupUser.Text = ""
                        
                        # Gruppenmitglieder neu laden, wenn sie bereits angezeigt werden
                        if ($null -ne $script:lstGroupMembers -and $script:lstGroupMembers.Items.Count -gt 0) {
                            $members = Get-GroupMembersAction -GroupName $script:txtExistingGroupName.Text
                            if ($null -ne $members) {
                                # Überprüfen, ob Control die ItemsSource-Eigenschaft unterstützt
                                if ($script:lstGroupMembers | Get-Member -Name "ItemsSource") {
                                    $script:lstGroupMembers.ItemsSource = $members
                                } else {
                                    # Alternative Methode, falls keine ItemsSource-Eigenschaft verfügbar
                                    $script:lstGroupMembers.Items.Clear()
                                    foreach ($member in $members) {
                                        [void]$script:lstGroupMembers.Items.Add($member)
                                    }
                                }
                            }
                        }
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Hinzufügen des Benutzers zur Gruppe: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnAddUserToGroup"
            
            Register-EventHandler -Control $btnRemoveUserFromGroup -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ([string]::IsNullOrWhiteSpace($script:txtExistingGroupName.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtGroupUser.Text)) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie den Namen der Gruppe und des zu entfernenden Benutzers an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Funktion zum Entfernen eines Benutzers aufrufen
                    $result = Remove-GroupMemberAction -GroupName $script:txtExistingGroupName.Text -MemberIdentity $script:txtGroupUser.Text
                    
                    if ($result) {
                        $script:txtStatus.Text = "Benutzer wurde erfolgreich aus der Gruppe entfernt."
                        
                        # Benutzerfeld zurücksetzen
                        $script:txtGroupUser.Text = ""
                        
                        # Gruppenmitglieder neu laden, wenn sie bereits angezeigt werden
                        if ($null -ne $script:lstGroupMembers) {
                            $members = Get-GroupMembersAction -GroupName $script:txtExistingGroupName.Text
                            if ($null -ne $members) {
                                # Überprüfen, ob Control die ItemsSource-Eigenschaft unterstützt
                                if ($script:lstGroupMembers | Get-Member -Name "ItemsSource") {
                                    $script:lstGroupMembers.ItemsSource = $members
                                } else {
                                    # Alternative Methode, falls keine ItemsSource-Eigenschaft verfügbar
                                    $script:lstGroupMembers.Items.Clear()
                                    foreach ($member in $members) {
                                        [void]$script:lstGroupMembers.Items.Add($member)
                                    }
                                }
                                
                                if ($members.Count -gt 0) {
                                    $script:txtStatus.Text = "$($members.Count) Gruppenmitglieder gefunden."
                                } else {
                                    $script:txtStatus.Text = "Gruppe ist leer."
                                }
                            } else {
                                $script:txtStatus.Text = "Gruppe wurde nicht gefunden oder ein Fehler ist aufgetreten."
                            }
                        }
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Entfernen des Benutzers aus der Gruppe: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnRemoveUserFromGroup"

            Register-EventHandler -Control $btnShowGroupMembers -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ([string]::IsNullOrWhiteSpace($script:txtExistingGroupName.Text)) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie den Namen der Gruppe an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Funktion zum Anzeigen der Gruppenmitglieder aufrufen
                    $members = Get-GroupMembersAction -GroupName $script:txtExistingGroupName.Text
                    
                    # DataGrid aktualisieren
                    $script:lstGroupMembers.ItemsSource = $members
                    
                    if ($members.Count -gt 0) {
                        $script:txtStatus.Text = "$($members.Count) Gruppenmitglieder gefunden."
                    } else {
                        $script:txtStatus.Text = "Gruppe ist leer oder wurde nicht gefunden."
                    }
                    
                    # Gruppeneinstellungen laden
                    $groupSettings = Get-GroupSettingsAction -GroupName $script:txtExistingGroupName.Text
                    if ($null -ne $groupSettings) {
                        $script:chkHiddenFromGAL.IsChecked = $groupSettings.HiddenFromAddressListsEnabled
                        $script:chkRequireSenderAuth.IsChecked = $groupSettings.RequireSenderAuthenticationEnabled
                        $script:chkAllowExternalSenders.IsChecked = (-not $groupSettings.RejectMessagesFromSendersOrMembers)
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Anzeigen der Gruppenmitglieder: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnShowGroupMembers"
            
            Register-EventHandler -Control $btnUpdateGroupSettings -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ([string]::IsNullOrWhiteSpace($script:txtExistingGroupName.Text)) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie den Namen der Gruppe an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $params = @{
                        GroupName = $script:txtExistingGroupName.Text
                        HiddenFromAddressListsEnabled = $script:chkHiddenFromGAL.IsChecked
                        RequireSenderAuthenticationEnabled = $script:chkRequireSenderAuth.IsChecked
                        AllowExternalSenders = $script:chkAllowExternalSenders.IsChecked
                    }
                    
                    # Funktion zum Aktualisieren der Gruppeneinstellungen aufrufen
                    $result = Update-GroupSettingsAction @params
                    
                    if ($result) {
                        $script:txtStatus.Text = "Gruppeneinstellungen wurden erfolgreich aktualisiert."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Aktualisieren der Gruppeneinstellungen: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnUpdateGroupSettings"
            
            # Hilfe-Link initialisieren
            if ($null -ne $helpLinkGroups) {
                $helpLinkGroups.Add_MouseLeftButtonDown({
                    Show-HelpDialog -Topic "Groups"
                })
                
                $helpLinkGroups.Add_MouseEnter({
                    $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                    $this.Cursor = [System.Windows.Input.Cursors]::Hand
                })
                
                $helpLinkGroups.Add_MouseLeave({
                    $this.TextDecorations = $null
                    $this.Cursor = [System.Windows.Input.Cursors]::Arrow
                })
            }
            
            Write-DebugMessage "Gruppen/Verteiler-Tab erfolgreich initialisiert" -Type "Success"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Initialisieren des Gruppen/Verteiler-Tabs: $errorMsg" -Type "Error"
            return $false
        }
    }

    function Initialize-SharedMailboxTab {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Shared Mailbox-Tab" -Type "Info"
            
            # UI-Elemente referenzieren
            $txtSharedMailboxName = Get-XamlElement -ElementName "txtSharedMailboxName"
            $txtSharedMailboxEmail = Get-XamlElement -ElementName "txtSharedMailboxEmail"
            $cmbSharedMailboxDomain = Get-XamlElement -ElementName "cmbSharedMailboxDomain"
            $btnCreateSharedMailbox = Get-XamlElement -ElementName "btnCreateSharedMailbox"
            $btnConvertToShared = Get-XamlElement -ElementName "btnConvertToShared"
            
            $txtSharedMailboxPermSource = Get-XamlElement -ElementName "txtSharedMailboxPermSource"
            $txtSharedMailboxPermUser = Get-XamlElement -ElementName "txtSharedMailboxPermUser"
            $cmbSharedMailboxPermType = Get-XamlElement -ElementName "cmbSharedMailboxPermType"
            $btnAddSharedMailboxPermission = Get-XamlElement -ElementName "btnAddSharedMailboxPermission"
            $btnRemoveSharedMailboxPermission = Get-XamlElement -ElementName "btnRemoveSharedMailboxPermission"
            
            $cmbSharedMailboxSelect = Get-XamlElement -ElementName "cmbSharedMailboxSelect"
            $btnShowSharedMailboxes = Get-XamlElement -ElementName "btnShowSharedMailboxes"
            $btnShowSharedMailboxPerms = Get-XamlElement -ElementName "btnShowSharedMailboxPerms"
            $chkAutoMapping = Get-XamlElement -ElementName "chkAutoMapping"
            $btnUpdateAutoMapping = Get-XamlElement -ElementName "btnUpdateAutoMapping"
            $txtForwardingAddress = Get-XamlElement -ElementName "txtForwardingAddress"
            $btnSetForwarding = Get-XamlElement -ElementName "btnSetForwarding"
            $btnHideFromGAL = Get-XamlElement -ElementName "btnHideFromGAL"
            $btnShowInGAL = Get-XamlElement -ElementName "btnShowInGAL"
            $btnRemoveSharedMailbox = Get-XamlElement -ElementName "btnRemoveSharedMailbox"
            $helpLinkShared = Get-XamlElement -ElementName "helpLinkShared"
            
            # Diese Elemente werden für zukünftige Funktionalität referenziert, aber aktuell noch nicht verwendet
            
            # Globale Variablen setzen
            $script:txtSharedMailboxName = $txtSharedMailboxName
            $script:txtSharedMailboxEmail = $txtSharedMailboxEmail
            $script:cmbSharedMailboxDomain = $cmbSharedMailboxDomain
            $script:txtSharedMailboxPermSource = $txtSharedMailboxPermSource
            $script:txtSharedMailboxPermUser = $txtSharedMailboxPermUser
            $script:cmbSharedMailboxPermType = $cmbSharedMailboxPermType
            $script:cmbSharedMailboxSelect = $cmbSharedMailboxSelect
            $script:chkAutoMapping = $chkAutoMapping
            $script:txtForwardingAddress = $txtForwardingAddress
            
            # Hilfsfunktion zum Aktualisieren der Shared Mailbox-Liste
            function RefreshSharedMailboxList {
                try {
                    if ($null -ne $script:cmbSharedMailboxSelect) {
                        $script:cmbSharedMailboxSelect.Items.Clear()
                        
                        # Abrufen aller Shared Mailboxen
                        $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | 
                                           Select-Object -ExpandProperty PrimarySmtpAddress
                        
                        foreach ($mailbox in $sharedMailboxes) {
                            [void]$script:cmbSharedMailboxSelect.Items.Add($mailbox)
                        }
                        
                        if ($script:cmbSharedMailboxSelect.Items.Count -gt 0) {
                            $script:cmbSharedMailboxSelect.SelectedIndex = 0
                        }
                        
                        Write-DebugMessage "Shared Mailbox-Liste wurde aktualisiert: $($script:cmbSharedMailboxSelect.Items.Count) gefunden" -Type "Info"
                    }
                    return $true
                }
                catch {
                    Write-DebugMessage "Fehler beim Aktualisieren der Shared Mailbox-Liste: $($_.Exception.Message)" -Type "Error"
                    return $false
                }
            }
            
            # Berechtigungstypen für die ComboBox initialisieren
            if ($null -ne $cmbSharedMailboxPermType) {
                $permissionTypes = @("FullAccess", "SendAs", "SendOnBehalf")
                $cmbSharedMailboxPermType.Items.Clear()
                foreach ($perm in $permissionTypes) {
                    $item = New-Object System.Windows.Controls.ComboBoxItem
                    $item.Content = $perm
                    [void]$cmbSharedMailboxPermType.Items.Add($item)
                }
                if ($cmbSharedMailboxPermType.Items.Count -gt 0) {
                    $cmbSharedMailboxPermType.SelectedIndex = 0
                }
            }
            
            # Domains für ComboBox laden
            if ($null -ne $cmbSharedMailboxDomain) {
                try {
                    # Sichere Ausführung der Domain-Abfrage
                    if ($script:isConnected) {
                        $domains = Get-AcceptedDomain | Select-Object -ExpandProperty DomainName
                        $cmbSharedMailboxDomain.Items.Clear()
                        foreach ($domain in $domains) {
                            [void]$cmbSharedMailboxDomain.Items.Add($domain)
                        }
                        if ($cmbSharedMailboxDomain.Items.Count -gt 0) {
                            $cmbSharedMailboxDomain.SelectedIndex = 0
                        }
                    }
                }
                catch {
                    Write-DebugMessage "Fehler beim Laden der Domains: $($_.Exception.Message)" -Type "Warning"
                }
            }
            
            # Event-Handler registrieren
            Register-EventHandler -Control $btnCreateSharedMailbox -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ([string]::IsNullOrWhiteSpace($script:txtSharedMailboxName.Text)) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie einen Namen für die Shared Mailbox an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # E-Mail-Adresse zusammensetzen, falls nicht angegeben
                    $mailboxEmail = $script:txtSharedMailboxEmail.Text
                    if ([string]::IsNullOrWhiteSpace($mailboxEmail) -and $null -ne $script:cmbSharedMailboxDomain.SelectedItem) {
                        # Einfache Normalisierung des Namens für die E-Mail-Adresse
                        $mailboxName = $script:txtSharedMailboxName.Text.ToLower() -replace '[^\w\d]', '.'
                        $domain = $script:cmbSharedMailboxDomain.SelectedItem.ToString()
                        $mailboxEmail = "$mailboxName@$domain"
                    }
                    
                    # Funktion zur Erstellung der Shared Mailbox aufrufen
                    $result = New-SharedMailboxAction -Name $script:txtSharedMailboxName.Text -EmailAddress $mailboxEmail
                    
                    if ($result) {
                        $script:txtStatus.Text = "Shared Mailbox wurde erfolgreich erstellt."
                        
                        # Felder zurücksetzen
                        $script:txtSharedMailboxName.Text = ""
                        $script:txtSharedMailboxEmail.Text = ""
                        
                        # Shared Mailbox Liste aktualisieren, falls geöffnet
                        if ($null -ne $script:cmbSharedMailboxSelect -and $script:cmbSharedMailboxSelect.Items.Count -gt 0) {
                            RefreshSharedMailboxList
                        }
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Erstellen der Shared Mailbox: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnCreateSharedMailbox"
            
            Register-EventHandler -Control $btnConvertToShared -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ([string]::IsNullOrWhiteSpace($script:txtSharedMailboxEmail.Text)) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie die E-Mail-Adresse des zu konvertierenden Postfachs an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Sicherheitsabfrage
                    $result = [System.Windows.MessageBox]::Show(
                        "Sind Sie sicher, dass Sie das Postfach '$($script:txtSharedMailboxEmail.Text)' in eine Shared Mailbox umwandeln möchten? " +
                        "Der Benutzer kann sich danach nicht mehr direkt an diesem Postfach anmelden.",
                        "Postfach konvertieren",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Warning)
                    
                    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                        # Funktion zum Konvertieren des Postfachs aufrufen
                        $result = Convert-ToSharedMailboxAction -Identity $script:txtSharedMailboxEmail.Text
                        
                        if ($result) {
                            $script:txtStatus.Text = "Postfach wurde erfolgreich in eine Shared Mailbox umgewandelt."
                            
                            # Shared Mailbox Liste aktualisieren, falls geöffnet
                            if ($null -ne $script:cmbSharedMailboxSelect -and $script:cmbSharedMailboxSelect.Items.Count -gt 0) {
                                RefreshSharedMailboxList
                            }
                        }
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Konvertieren des Postfachs: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnConvertToShared"
            
            Register-EventHandler -Control $btnAddSharedMailboxPermission -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ([string]::IsNullOrWhiteSpace($script:txtSharedMailboxPermSource.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtSharedMailboxPermUser.Text) -or 
                        $null -eq $script:cmbSharedMailboxPermType.SelectedItem) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie Shared Mailbox, Benutzer und Berechtigungstyp an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $mailbox = $script:txtSharedMailboxPermSource.Text
                    $user = $script:txtSharedMailboxPermUser.Text
                    $permType = $script:cmbSharedMailboxPermType.SelectedItem.Content.ToString()
                    $autoMapping = $script:chkAutoMapping.IsChecked
                    
                    # Funktion zum Hinzufügen der Berechtigung aufrufen
                    $result = Add-SharedMailboxPermissionAction -Mailbox $mailbox -User $user -PermissionType $permType -AutoMapping $autoMapping
                    
                    if ($result) {
                        $script:txtStatus.Text = "Berechtigung wurde erfolgreich hinzugefügt."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Hinzufügen der Berechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnAddSharedMailboxPermission"
            
            Register-EventHandler -Control $btnRemoveSharedMailboxPermission -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ([string]::IsNullOrWhiteSpace($script:txtSharedMailboxPermSource.Text) -or 
                        [string]::IsNullOrWhiteSpace($script:txtSharedMailboxPermUser.Text) -or 
                        $null -eq $script:cmbSharedMailboxPermType.SelectedItem) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie Shared Mailbox, Benutzer und Berechtigungstyp an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $mailbox = $script:txtSharedMailboxPermSource.Text
                    $user = $script:txtSharedMailboxPermUser.Text
                    $permType = $script:cmbSharedMailboxPermType.SelectedItem.Content.ToString()
                    
                    # Funktion zum Entfernen der Berechtigung aufrufen
                    $result = Remove-SharedMailboxPermissionAction -Mailbox $mailbox -User $user -PermissionType $permType
                    
                    if ($result) {
                        $script:txtStatus.Text = "Berechtigung wurde erfolgreich entfernt."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Entfernen der Berechtigung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnRemoveSharedMailboxPermission"
            
            Register-EventHandler -Control $btnShowSharedMailboxes -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Funktion zum Anzeigen der Shared Mailboxes aufrufen
                    RefreshSharedMailboxList
                    
                    if ($script:cmbSharedMailboxSelect.Items.Count -gt 0) {
                        $script:txtStatus.Text = "$($script:cmbSharedMailboxSelect.Items.Count) Shared Mailboxes gefunden."
                    } else {
                        $script:txtStatus.Text = "Keine Shared Mailboxes gefunden."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Anzeigen der Shared Mailboxes: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnShowSharedMailboxes"
            
            Register-EventHandler -Control $btnShowSharedMailboxPerms -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ($null -eq $script:cmbSharedMailboxSelect.SelectedItem) {
                        [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Shared Mailbox aus.", 
                            "Keine Auswahl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $mailbox = $script:cmbSharedMailboxSelect.SelectedItem.ToString()
                    
                    # Funktion zum Anzeigen der Berechtigungen aufrufen
                    $permissions = Get-SharedMailboxPermissionsAction -Mailbox $mailbox
                    
                    # MessageBox mit den Berechtigungen anzeigen
                    if ($permissions.Count -gt 0) {
                        $permText = "Berechtigungen für Shared Mailbox '$mailbox':`n`n"
                        foreach ($perm in $permissions) {
                            $permText += "- $($perm.User): $($perm.AccessRights) ($($perm.PermissionType))`n"
                        }
                        
                        [System.Windows.MessageBox]::Show(
                            $permText,
                            "Shared Mailbox Berechtigungen",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Information)
                        
                        $script:txtStatus.Text = "$($permissions.Count) Berechtigungen gefunden."
                    } else {
                        [System.Windows.MessageBox]::Show(
                            "Keine Berechtigungen für die ausgewählte Shared Mailbox gefunden.",
                            "Shared Mailbox Berechtigungen",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Information)
                        
                        $script:txtStatus.Text = "Keine Berechtigungen gefunden."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Anzeigen der Shared Mailbox Berechtigungen: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnShowSharedMailboxPerms"
            
            Register-EventHandler -Control $btnUpdateAutoMapping -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ($null -eq $script:cmbSharedMailboxSelect.SelectedItem) {
                        [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Shared Mailbox aus.", 
                            "Keine Auswahl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $mailbox = $script:cmbSharedMailboxSelect.SelectedItem.ToString()
                    $autoMapping = $script:chkAutoMapping.IsChecked
                    
                    # Funktion zum Aktualisieren des Auto-Mappings aufrufen
                    $result = Update-SharedMailboxAutoMappingAction -Mailbox $mailbox -AutoMapping $autoMapping
                    
                    if ($result) {
                        $script:txtStatus.Text = "Auto-Mapping wurde erfolgreich aktualisiert."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Aktualisieren des Auto-Mappings: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnUpdateAutoMapping"
            
            Register-EventHandler -Control $btnSetForwarding -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ($null -eq $script:cmbSharedMailboxSelect.SelectedItem -or 
                        [string]::IsNullOrWhiteSpace($script:txtForwardingAddress.Text)) {
                        [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Shared Mailbox aus und geben Sie eine Weiterleitungsadresse an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $mailbox = $script:cmbSharedMailboxSelect.SelectedItem.ToString()
                    $forwardingAddress = $script:txtForwardingAddress.Text
                    
                    # Funktion zum Einrichten der Weiterleitung aufrufen
                    $result = Set-SharedMailboxForwardingAction -Mailbox $mailbox -ForwardingAddress $forwardingAddress
                    
                    if ($result) {
                        $script:txtStatus.Text = "Weiterleitung wurde erfolgreich eingerichtet."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Einrichten der Weiterleitung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnSetForwarding"
            
            Register-EventHandler -Control $btnHideFromGAL -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ($null -eq $script:cmbSharedMailboxSelect.SelectedItem) {
                        [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Shared Mailbox aus.", 
                            "Keine Auswahl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $mailbox = $script:cmbSharedMailboxSelect.SelectedItem.ToString()
                    
                    # Funktion zum Ausblenden aus GAL aufrufen
                    $result = Set-SharedMailboxGALVisibilityAction -Mailbox $mailbox -HideFromGAL $true
                    
                    if ($result) {
                        $script:txtStatus.Text = "Shared Mailbox wurde aus der GAL ausgeblendet."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Ausblenden aus der GAL: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnHideFromGAL"
            
            Register-EventHandler -Control $btnShowInGAL -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ($null -eq $script:cmbSharedMailboxSelect.SelectedItem) {
                        [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Shared Mailbox aus.", 
                            "Keine Auswahl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $mailbox = $script:cmbSharedMailboxSelect.SelectedItem.ToString()
                    
                    # Funktion zum Einblenden in GAL aufrufen
                    $result = Set-SharedMailboxGALVisibilityAction -Mailbox $mailbox -HideFromGAL $false
                    
                    if ($result) {
                        $script:txtStatus.Text = "Shared Mailbox wurde in der GAL sichtbar gemacht."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Anzeigen in der GAL: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnShowInGAL"
            
            # Event-Handler für Löschen einer Shared Mailbox
            Register-EventHandler -Control $btnRemoveSharedMailbox -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Eingabeprüfung
                    if ($null -eq $script:cmbSharedMailboxSelect.SelectedItem) {
                        [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Shared Mailbox aus.", 
                            "Keine Auswahl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Sicherheitsabfrage
                    $mailbox = $script:cmbSharedMailboxSelect.SelectedItem.ToString()
                    $result = [System.Windows.MessageBox]::Show(
                        "Sind Sie sicher, dass Sie die Shared Mailbox '$mailbox' löschen möchten? Diese Aktion kann nicht rückgängig gemacht werden.",
                        "Shared Mailbox löschen",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Warning)
                    
                    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                        # Funktion zum Löschen der Shared Mailbox aufrufen
                        $result = Remove-SharedMailboxAction -Mailbox $mailbox
                        
                        if ($result) {
                            $script:txtStatus.Text = "Shared Mailbox wurde erfolgreich gelöscht."
                            RefreshSharedMailboxList
                        }
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Löschen der Shared Mailbox: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnRemoveSharedMailbox"
            
            # Hilfe-Link initialisieren
            if ($null -ne $helpLinkShared) {
                $helpLinkShared.Add_MouseLeftButtonDown({
                    Show-HelpDialog -Topic "SharedMailbox"
                })
                
                $helpLinkShared.Add_MouseEnter({
                    $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                    $this.Cursor = [System.Windows.Input.Cursors]::Hand
                })
                
                $helpLinkShared.Add_MouseLeave({
                    $this.TextDecorations = $null
                    $this.Cursor = [System.Windows.Input.Cursors]::Arrow
                })
            }
            
            Write-DebugMessage "Shared Mailbox-Tab erfolgreich initialisiert" -Type "Success"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Initialisieren des Shared Mailbox-Tabs: $errorMsg" -Type "Error"
            return $false
        }
    }

    function Initialize-AuditTab {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Audit-Tab" -Type "Info"
            
            # Referenzieren der UI-Elemente im Audit-Tab
            $txtAuditMailbox = Get-XamlElement -ElementName "txtAuditMailbox" 
            $cmbAuditCategory = Get-XamlElement -ElementName "cmbAuditCategory"
            $cmbAuditType = Get-XamlElement -ElementName "cmbAuditType"
            $btnRunAudit = Get-XamlElement -ElementName "btnRunAudit"
            $txtAuditResult = Get-XamlElement -ElementName "txtAuditResult"
            $helpLinkAudit = Get-XamlElement -ElementName "helpLinkAudit"
            
            # Globale Variablen für spätere Verwendung setzen
            $script:txtAuditMailbox = $txtAuditMailbox
            $script:cmbAuditCategory = $cmbAuditCategory
            $script:cmbAuditType = $cmbAuditType
            $script:txtAuditResult = $txtAuditResult
            
            # Event-Handler für Audit-Buttons registrieren
            Register-EventHandler -Control $btnRunAudit -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob alle erforderlichen Eingaben vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtAuditMailbox.Text) -or 
                        $null -eq $script:cmbAuditCategory.SelectedItem -or
                        $null -eq $script:cmbAuditType.SelectedItem) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte geben Sie eine E-Mail-Adresse ein und wählen Sie eine Kategorie und einen Typ.",
                            "Unvollständige Eingabe",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $mailbox = $script:txtAuditMailbox.Text
                    $navigationType = $script:cmbAuditCategory.SelectedItem.Content
                    $infoType = $script:cmbAuditType.SelectedIndex + 1
                    
                    Write-DebugMessage "Führe Audit aus: $navigationType / $infoType für $mailbox" -Type "Info"
                    $result = Get-FormattedMailboxInfo -Mailbox $mailbox -InfoType $infoType -NavigationType $navigationType
                    
                    if ($null -ne $script:txtAuditResult) {
                        $script:txtAuditResult.Text = $result
                    }
                    
                    $script:txtStatus.Text = "Audit erfolgreich ausgeführt."
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Ausführen des Audits: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                    
                    if ($null -ne $script:txtAuditResult) {
                        $script:txtAuditResult.Text = "Fehler beim Ausführen des Audits: $errorMsg"
                    }
                }
            } -ControlName "btnRunAudit"
            
            if ($null -ne $helpLinkAudit) {
                $helpLinkAudit.Add_MouseLeftButtonDown({
                    Show-HelpDialog -Topic "Audit"
                })
                
                $helpLinkAudit.Add_MouseEnter({
                    $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                    $this.Cursor = [System.Windows.Input.Cursors]::Hand
                })
                
                $helpLinkAudit.Add_MouseLeave({
                    $this.TextDecorations = $null
                    $this.Cursor = [System.Windows.Input.Cursors]::Arrow
                })
            }
            
            # Initialisierung der ComboBoxen
            if ($null -ne $cmbAuditCategory) {
                $auditCategories = @("Postfach-Informationen", "Postfach-Statistiken", "Postfach-Berechtigungen", "Audit-Konfiguration", "E-Mail-Weiterleitung")
                $cmbAuditCategory.Items.Clear()
                foreach ($category in $auditCategories) {
                    $item = New-Object System.Windows.Controls.ComboBoxItem
                    $item.Content = $category
                    [void]$cmbAuditCategory.Items.Add($item)
                }
                if ($cmbAuditCategory.Items.Count -gt 0) {
                    $cmbAuditCategory.SelectedIndex = 0
                }
            }
            
            if ($null -ne $cmbAuditCategory -and $null -ne $cmbAuditType) {
                # Event-Handler für die Kategorie-ComboBox
                $cmbAuditCategory.Add_SelectionChanged({
                    try {
                        if ($null -ne $script:cmbAuditType) {
                            $script:cmbAuditType.Items.Clear()
                            
                            switch ($script:cmbAuditCategory.SelectedIndex) {
                                0 { # Postfach-Informationen
                                    $infoTypes = @("Grundlegende Informationen", "Speicherbegrenzungen", "E-Mail-Adressen", "Funktion/Rolle", "Alle Details")
                                }
                                1 { # Postfach-Statistiken
                                    $infoTypes = @("Größeninformationen", "Ordnerinformationen", "Nutzungsstatistiken", "Zusammenfassende Statistiken", "Alle Statistiken")
                                }
                                2 { # Postfach-Berechtigungen
                                    $infoTypes = @("Postfach-Berechtigungen", "SendAs-Berechtigungen", "SendOnBehalf-Berechtigungen", "Kalenderberechtigungen", "Alle Berechtigungen")
                                }
                                3 { # Audit-Konfiguration
                                    $infoTypes = @("Audit-Konfiguration", "Audit-Empfehlungen", "Audit-Status prüfen", "Audit-Konfiguration anpassen", "Vollständige Audit-Details")
                                }
                                4 { # E-Mail-Weiterleitung
                                    $infoTypes = @("Weiterleitungseinstellungen", "Analyse externer Weiterleitungen", "Entfernen von Weiterleitungen", "Transport-Regeln prüfen", "Vollständige Weiterleitungsdetails")
                                }
                                default {
                                    $infoTypes = @("Option 1", "Option 2", "Option 3", "Option 4", "Option 5")
                                }
                            }
                            
                            foreach ($type in $infoTypes) {
                                $item = New-Object System.Windows.Controls.ComboBoxItem
                                $item.Content = $type
                                [void]$script:cmbAuditType.Items.Add($item)
                            }
                            
                            if ($script:cmbAuditType.Items.Count -gt 0) {
                                $script:cmbAuditType.SelectedIndex = 0
                            }
                        }
                    }
                    catch {
                        $errorMsg = $_.Exception.Message
                        Write-DebugMessage "Fehler beim Aktualisieren der Audit-Typen: $errorMsg" -Type "Error"
                    }
                })
                
                # Initial die erste Kategorie auswählen, um die Typen zu laden
                if ($cmbAuditCategory.Items.Count -gt 0) {
                    $cmbAuditCategory.SelectedIndex = 0
                }
            }
            
            Write-DebugMessage "Audit-Tab erfolgreich initialisiert" -Type "Success"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Initialisieren des Audit-Tabs: $errorMsg" -Type "Error"
            return $false
        }
    }

    function Initialize-TroubleshootingTab {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Troubleshooting-Tab" -Type "Info"
            
            # Referenzieren der UI-Elemente im Troubleshooting-Tab
            $lstDiagnostics = Get-XamlElement -ElementName "lstDiagnostics"
            $txtDiagnosticUser = Get-XamlElement -ElementName "txtDiagnosticUser"
            $txtDiagnosticUser2 = Get-XamlElement -ElementName "txtDiagnosticUser2"
            $txtDiagnosticEmail = Get-XamlElement -ElementName "txtDiagnosticEmail"
            $btnRunDiagnostic = Get-XamlElement -ElementName "btnRunDiagnostic"
            $btnAdminCenter = Get-XamlElement -ElementName "btnAdminCenter"
            $txtDiagnosticResult = Get-XamlElement -ElementName "txtDiagnosticResult"
            $helpLinkTroubleshooting = Get-XamlElement -ElementName "helpLinkTroubleshooting"
            
            # Globale Variablen für spätere Verwendung setzen
            $script:lstDiagnostics = $lstDiagnostics
            $script:txtDiagnosticUser = $txtDiagnosticUser
            $script:txtDiagnosticUser2 = $txtDiagnosticUser2
            $script:txtDiagnosticEmail = $txtDiagnosticEmail
            $script:txtDiagnosticResult = $txtDiagnosticResult
            
            # Diagnostics-Liste befüllen
            if ($null -ne $lstDiagnostics) {
                $lstDiagnostics.Items.Clear()
                for ($i = 0; $i -lt $script:exchangeDiagnostics.Count; $i++) {
                    $diagnostic = $script:exchangeDiagnostics[$i]
                    $item = New-Object PSObject -Property @{
                        Name = $diagnostic.Name
                        Description = $diagnostic.Description
                        Index = $i
                    }
                    [void]$lstDiagnostics.Items.Add($item)
                }
                
                # Set DisplayMemberPath to only show the Name property
                $lstDiagnostics.DisplayMemberPath = "Name"
            }
            
            # Event-Handler für Diagnose-Button registrieren
            Register-EventHandler -Control $btnRunDiagnostic -Handler {
                try {
                    # Prüfen, ob eine Exchange-Verbindung besteht
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.",
                            "Keine Verbindung",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    # Prüfen, ob ein Eintrag in der Liste ausgewählt wurde
                    if ($null -eq $script:lstDiagnostics.SelectedItem) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte wählen Sie eine Diagnose aus der Liste.",
                            "Keine Diagnose ausgewählt",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $diagnosticIndex = $script:lstDiagnostics.SelectedItem.Index
                    $user = $script:txtDiagnosticUser.Text
                    $user2 = $script:txtDiagnosticUser2.Text
                    $email = $script:txtDiagnosticEmail.Text
                    
                    Write-DebugMessage "Führe Diagnose aus: Index=$diagnosticIndex, User=$user, User2=$user2, Email=$email" -Type "Info"
                    
                    $result = Run-ExchangeDiagnostic -DiagnosticIndex $diagnosticIndex -User $user -User2 $user2 -Email $email
                    
                    if ($null -ne $script:txtDiagnosticResult) {
                        $script:txtDiagnosticResult.Text = $result
                        $script:txtStatus.Text = "Diagnose erfolgreich ausgeführt."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler bei der Diagnose: $errorMsg" -Type "Error"
                    if ($null -ne $script:txtStatus) {
                        $script:txtStatus.Text = "Fehler: $errorMsg"
                    }
                    
                    if ($null -ne $script:txtDiagnosticResult) {
                        $script:txtDiagnosticResult.Text = "Fehler bei der Diagnose: $errorMsg"
                    }
                }
            } -ControlName "btnRunDiagnostic"
            
            # Event-Handler für Admin-Center-Button
            Register-EventHandler -Control $btnAdminCenter -Handler {
                try {
                    # Prüfen, ob ein Eintrag in der Liste ausgewählt wurde
                    if ($null -eq $script:lstDiagnostics.SelectedItem) {
                        [System.Windows.MessageBox]::Show(
                            "Bitte wählen Sie eine Diagnose aus der Liste.",
                            "Keine Diagnose ausgewählt",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                        return
                    }
                    
                    $diagnosticIndex = $script:lstDiagnostics.SelectedItem.Index
                    
                    Write-DebugMessage "Öffne Admin-Center für Diagnose: Index=$diagnosticIndex" -Type "Info"
                    $result = Open-AdminCenterLink -DiagnosticIndex $diagnosticIndex
                    
                    if (-not $result) {
                        [System.Windows.MessageBox]::Show(
                            "Kein Admin-Center-Link für diese Diagnose verfügbar.",
                            "Kein Link verfügbar",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Information
                        )
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Öffnen des Admin-Centers: $errorMsg" -Type "Error"
                    if ($null -ne $script:txtStatus) {
                        $script:txtStatus.Text = "Fehler: $errorMsg"
                    }
                }
            } -ControlName "btnAdminCenter"
            
            # Event-Handler für Hilfe-Link
            if ($null -ne $helpLinkTroubleshooting) {
                $helpLinkTroubleshooting.Add_MouseLeftButtonDown({
                    Show-HelpDialog -Topic "Troubleshooting"
                })
                
                $helpLinkTroubleshooting.Add_MouseEnter({
                    $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                    $this.Cursor = [System.Windows.Input.Cursors]::Hand
                })
                
                $helpLinkTroubleshooting.Add_MouseLeave({
                    $this.TextDecorations = $null
                    $this.Cursor = [System.Windows.Input.Cursors]::Arrow
                })
            }
            
            Write-DebugMessage "Troubleshooting-Tab erfolgreich initialisiert" -Type "Success"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Initialisieren des Troubleshooting-Tabs: $errorMsg" -Type "Error"
            return $false
        }
    }
    
    # Verbesserte Version der HelpLinks-Initialisierung
    function Initialize-HelpLinks {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Hilfe-Links" -Type "Info"
            
            # Da möglicherweise kein HelpLinks-Container existiert,
            # suche nach individuellen Hilfe-Links
            $helpLinks = @(
                $script:Form.FindName("helpLinkCalendar"),
                $script:Form.FindName("helpLinkMailbox"),
                $script:Form.FindName("helpLinkAudit"),
                $script:Form.FindName("helpLinkTroubleshooting")
            )
            
            $foundLinks = 0
            foreach ($link in $helpLinks) {
                if ($null -ne $link) {
                    $foundLinks++
                    $linkName = $link.Name
                    $topic = "General"
                    
                    # Bestimme das Hilfe-Thema basierend auf dem Link-Namen
                    switch -Wildcard ($linkName) {
                        "*Calendar*" { $topic = "Calendar" }
                        "*Mailbox*" { $topic = "Mailbox" }
                        "*Audit*" { $topic = "Audit" }
                        "*Trouble*" { $topic = "Troubleshooting" }
                    }
                    
                    # Event-Handler hinzufügen
                    $topicClosure = $topic  # Variable-Capture für den Closure
                    $link.Add_MouseLeftButtonDown({
                        Show-HelpDialog -Topic $topicClosure
                    })
                    
                    # Mauszeiger ändern, wenn auf den Link gezeigt wird
                    $link.Add_MouseEnter({
                        $this.Cursor = [System.Windows.Input.Cursors]::Hand
                        $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                    })
                    
                    $link.Add_MouseLeave({
                        $this.TextDecorations = $null
                    })
                    
                    Write-DebugMessage "Hilfe-Link initialisiert: $linkName für Thema $topic" -Type "Info"
                }
            }
            
            if ($foundLinks -eq 0) {
                Write-DebugMessage "Keine Hilfe-Links in der XAML gefunden - dies ist kein kritischer Fehler" -Type "Warning"
                return $false
            }
            
            Write-DebugMessage "$foundLinks Hilfe-Links erfolgreich initialisiert" -Type "Success"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Initialisieren der Hilfe-Links: $errorMsg" -Type "Error"
            Log-Action "Fehler beim Initialisieren der Hilfe-Links: $errorMsg"
            return $false
        }
    }

    # Initialisiere alle Tabs
    function Initialize-AllTabs {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere alle Tabs" -Type "Info"
            
            $results = @{
                Calendar = Initialize-CalendarTab
                Mailbox = Initialize-MailboxTab
                Audit = Initialize-AuditTab
                Troubleshooting = Initialize-TroubleshootingTab
            }
            
            $successCount = ($results.Values | Where-Object {$_ -eq $true}).Count
            $totalCount = $results.Count
            
            Write-DebugMessage "Tab-Initialisierung abgeschlossen: $successCount von $totalCount Tabs erfolgreich initialisiert" -Type "Info"
            
            foreach ($tab in $results.Keys) {
                $status = if ($results[$tab]) { "erfolgreich" } else { "fehlgeschlagen" }
                Write-DebugMessage "Tab $tab - Initialisierung $status" -Type "Info"
            }
            
            if ($successCount -eq $totalCount) {
                Write-DebugMessage "Alle Tabs erfolgreich initialisiert" -Type "Success"
                return $true
            } else {
                Write-DebugMessage "Einige Tabs konnten nicht initialisiert werden" -Type "Warning"
                return ($successCount -gt 0) # Wenn mindestens ein Tab initialisiert wurde, gilt es als teilweise erfolgreich
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Initialisieren der Tabs: $errorMsg" -Type "Error"
            return $false
        }
    }

    # Event-Handler für das Loaded-Event des Formulars
    $script:Form.Add_Loaded({
        Write-DebugMessage "GUI-Loaded-Event ausgelöst, initialisiere Komponenten" -Type "Info"
        
        # Version anzeigen
        if ($null -ne $script:txtVersion) {
            try {
                if ($null -ne $script:config -and 
                    $null -ne $script:config["General"] -and 
                    $null -ne $script:config["General"]["Version"]) {
                    $script:txtVersion.Text = "v" + $script:config["General"]["Version"]
                } else {
                    $script:txtVersion.Text = "v0.0.5" # Fallback auf Standardversion
                }
            } catch {
                $script:txtVersion.Text = "v0.0.5" # Fallback bei Fehler
                Write-DebugMessage "Fehler beim Setzen der Version: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Verbindungsstatus initialisieren
        if ($null -ne $script:txtConnectionStatus) {
            $script:txtConnectionStatus.Text = "Nicht verbunden"
            $script:txtConnectionStatus.Foreground = $script:disconnectedBrush
        }
        
        # Statusmeldung aktualisieren
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Bereit. Bitte stellen Sie eine Verbindung zu Exchange Online her."
        }
        
        # WICHTIG: Tabs initialisieren
        $result = Initialize-AllTabs
        Write-DebugMessage "Initialize-AllTabs Ergebnis: $result" -Type "Info"
        
        # Hilfe-Links initialisieren
        $result = Initialize-HelpLinks
        Write-DebugMessage "Initialize-HelpLinks Ergebnis: $result" -Type "Info"
    })

    # -------------------------------------------------
    # Fenster anzeigen
    # -------------------------------------------------
    Write-DebugMessage "Öffne GUI-Fenster" -Type "Info"
    [void]$script:Form.ShowDialog()
    Write-DebugMessage "GUI-Fenster wurde geschlossen" -Type "Info"
    
    # Aufräumen nach Schließen des Fensters
    if ($script:isConnected) {
        Write-DebugMessage "Trenne Exchange Online-Verbindung..." -Type "Info"
        Disconnect-ExchangeOnlineSession
    }
}
catch {
    $errorMsg = $_.Exception.Message
    Write-Host "Kritischer Fehler beim Laden oder Anzeigen der GUI: $errorMsg" -ForegroundColor Red
    
    # Falls die XAML-Datei nicht geladen werden kann, zeige einen alternativen Fehlerhinweis
    try {
        [System.Windows.MessageBox]::Show(
            "Kritischer Fehler beim Laden oder Anzeigen der GUI: $errorMsg`n`nBitte stellen Sie sicher, dass die XAML-Datei vorhanden und korrekt formatiert ist.", 
            "Fehler", 
            [System.Windows.MessageBoxButton]::OK, 
            [System.Windows.MessageBoxImage]::Error
        )
    }
    catch {
        Write-Host "Konnte keine MessageBox anzeigen. Zusätzlicher Fehler: $($_.Exception.Message)" -ForegroundColor Red
    }
}
finally {
    # Aufräumarbeiten
    if ($null -ne $script:Form) {
        $script:Form.Close()
        $script:Form = $null
    }
    
    Write-DebugMessage "Aufräumarbeiten abgeschlossen" -Type "Info"
}