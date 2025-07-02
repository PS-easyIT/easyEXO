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

# MessageBox-Funktion
function Show-MessageBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Title = "Information",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "Question")]
        [string]$Type = "Info"
    )
    
    try {
        $icon = switch ($Type) {
            "Info" { [System.Windows.MessageBoxImage]::Information }
            "Warning" { [System.Windows.MessageBoxImage]::Warning }
            "Error" { [System.Windows.MessageBoxImage]::Error }
            "Question" { [System.Windows.MessageBoxImage]::Question }
        }
        
        $buttons = if ($Type -eq "Question") { 
            [System.Windows.MessageBoxButton]::YesNo 
        } else { 
            [System.Windows.MessageBoxButton]::OK 
        }
        
        $result = [System.Windows.MessageBox]::Show($Message, $Title, $buttons, $icon)
        
        # Erfolg loggen
        Write-DebugMessage -Message "$Title - $Type - $Message" -Type "Info"
        
        # Ergebnis zurückgeben (wichtig für Ja/Nein-Fragen)
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage -Message "Fehler beim Anzeigen der MessageBox: $errorMsg" -Type "Error"
        
        # Fallback-Ausgabe
        Write-Host "Meldung ($Type): $Title - $Message" -ForegroundColor Red
        
        if ($Type -eq "Question") {
            return [System.Windows.MessageBoxResult]::No
        }
    }
}

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
        
        # GUI-Element im UI-Thread aktualisieren mit Überprüfung des Dispatcher-Status
        if ($null -ne $TextElement.Dispatcher -and $TextElement.Dispatcher.CheckAccess()) {
            # Wir sind bereits im UI-Thread
            $TextElement.Text = $sanitizedMessage
            if ($null -ne $Color) {
                $TextElement.Foreground = $Color
            }
        } 
        else {
            # Dispatcher verwenden für Thread-Sicherheit
            $TextElement.Dispatcher.Invoke([Action]{
                $TextElement.Text = $sanitizedMessage
                if ($null -ne $Color) {
                    $TextElement.Foreground = $Color
                }
            }, "Normal")
        }
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

# Funktion zum Aktualisieren des Status in der GUI
function Write-StatusMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Type = "Info"
    )
    
    try {
        # Logge die Nachricht auch
        Write-DebugMessage -Message $Message -Type $Type
        
        # Bestimme die Farbe basierend auf dem Nachrichtentyp
        $color = switch ($Type) {
            "Success" { $script:connectedBrush }
            "Error" { $script:disconnectedBrush }
            "Warning" { New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Orange) }
            "Info" { $null }
            default { $null }
        }
        
        # Aktualisiere das Status-Textfeld in der GUI
        if ($null -ne $script:txtStatus) {
            Update-GuiText -TextElement $script:txtStatus -Message $Message -Color $color
        }
    } 
    catch {
        # Bei Fehler einfach eine Debug-Meldung ausgeben
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler in Write-StatusMessage: $errorMsg" -Type "Error"
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
# Abschnitt: Eingabevalidierung
# -------------------------------------------------
function  Validate-Email{
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
            
            # Verbindung immer überprüfen
            try {
                # Einfache Prüfung durch Ausführen eines kleinen Exchange-Befehls
                Get-OrganizationConfig -ErrorAction Stop | Out-Null
                Write-DebugMessage "Bestehende Exchange-Verbindung ist gültig" -Type "Info"
            }
            catch {
                Write-DebugMessage "Bestehende Verbindung ist nicht mehr gültig, stelle neue Verbindung her" -Type "Warning"
                $Global:IsConnectedToExo = $false
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

function Test-ExchangeOnlineConnection {
    [CmdletBinding()]
    param()
    
    try {
        # Zuerst prüfen, ob die globale Testfunktion verfügbar ist
        if ($null -ne ${global:EasyStartup_TestExchangeOnlineConnection}) {
            return & ${global:EasyStartup_TestExchangeOnlineConnection}
        }
        
        # Eigene Verbindungsprüfung als Fallback
        try {
            # Einfacher Test durch Ausführung eines Exchange-Befehls
            Get-OrganizationConfig -ErrorAction Stop | Out-Null
            Write-DebugMessage "Exchange Online-Verbindung ist aktiv" -Type "Info"
            return $true
        }
        catch {
            Write-DebugMessage "Exchange Online-Verbindung ist nicht mehr gültig: $($_.Exception.Message)" -Type "Warning"
            return $false
        }
    }
    catch {
        Write-DebugMessage "Fehler beim Testen der Exchange Online-Verbindung: $($_.Exception.Message)" -Type "Error"
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
function Initialize-CalendarTab {
    [CmdletBinding()]
    param()
    
    try {
        Write-DebugMessage "Initialisiere Calendar Tab..." -Type "Info"
        
        # Überprüfe ob die notwendigen UI-Elemente existieren
        $calendarTab = Get-XamlElement -ElementName "tabCalendar"
        if ($null -eq $calendarTab) {
            Write-DebugMessage "Calendar Tab Element nicht gefunden" -Type "Error"
            return $false
        }
        
        # Initialisiere die Calendar-spezifischen UI-Elemente
        $calendarElements = @(
            "calendarView",
            "btnRefreshCalendar",
            "btnAddCalendarPermission",
            "btnRemoveCalendarPermission"
        )
        
        foreach ($elementName in $calendarElements) {
            $element = Get-XamlElement -ElementName $elementName
            if ($null -eq $element) {
                Write-DebugMessage "Calendar Element '$elementName' nicht gefunden" -Type "Warning"
                continue
            }
            
            # Registriere Event-Handler für die Elemente
            switch ($elementName) {
                "btnRefreshCalendar" {
                    $element.Add_Click({
                        Write-DebugMessage "Calendar Refresh angefordert" -Type "Info"
                        # Implementiere Refresh-Logik hier
                    })
                }
                "btnAddCalendarPermission" {
                    $element.Add_Click({
                        Write-DebugMessage "Kalenderberechtigung hinzufügen angefordert" -Type "Info"
                        # Implementiere Add-Permission-Logik hier
                    })
                }
                "btnRemoveCalendarPermission" {
                    $element.Add_Click({
                        Write-DebugMessage "Kalenderberechtigung entfernen angefordert" -Type "Info"
                        # Implementiere Remove-Permission-Logik hier
                    })
                }
            }
            
            Write-DebugMessage "Event-Handler für $elementName registriert" -Type "Info"
        }
        
        Write-DebugMessage "Calendar Tab erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        Write-DebugMessage "Fehler beim Initialisieren des Calendar Tabs: $($_.Exception.Message)" -Type "Error"
        return $false
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
            # Frage den Benutzer ob er das wirklich tun möchte
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
            
            $successCount = 0
            $errorCount = 0

            if ($PermissionType -eq "Standard" -or $PermissionType -eq "Beides") {
                $result = Set-DefaultCalendarPermissionForAll -AccessRights $AccessRights
                if ($result) { $successCount++ } else { $errorCount++ }
            }
            if ($PermissionType -eq "Anonym" -or $PermissionType -eq "Beides") {
                $result = Set-AnonymousCalendarPermissionForAll -AccessRights $AccessRights
                if ($result) { $successCount++ } else { $errorCount++ }
            }
        }
        else {         
            if ([string]::IsNullOrWhiteSpace($MailboxUser) -and 
                $null -ne $script:txtCalendarMailboxUser -and 
                -not [string]::IsNullOrWhiteSpace($script:txtCalendarMailboxUser.Text)) {
                $mailboxUser = $script:txtCalendarMailboxUser.Text.Trim()
            }
            
            if ([string]::IsNullOrWhiteSpace($mailboxUser)) {
                throw "Keine Postfach-E-Mail-Adresse angegeben"
            }
            
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
        
        # Prüfe ob Berechtigung bereits existiert und ermittle den korrekten Kalenderordner
        $calendarExists = $false
        $identityDE = "${SourceUser}:\Kalender"
        $identityEN = "${SourceUser}:\Calendar"
        $identity = $null
        
        # Systematisch nach dem richtigen Kalender suchen
        try {
            # Zuerst versuchen wir den deutschen Kalender
            $existingPermDE = Get-MailboxFolderPermission -Identity $identityDE -User $TargetUser -ErrorAction SilentlyContinue
            if ($null -ne $existingPermDE) {
                $calendarExists = $true
                $identity = $identityDE
                Write-DebugMessage "Bestehende Berechtigung gefunden (DE): $($existingPermDE.AccessRights)" -Type "Info"
            }
            else {
                # Dann den englischen Kalender probieren
                $existingPermEN = Get-MailboxFolderPermission -Identity $identityEN -User $TargetUser -ErrorAction SilentlyContinue
                if ($null -ne $existingPermEN) {
                    $calendarExists = $true
                    $identity = $identityEN
                    Write-DebugMessage "Bestehende Berechtigung gefunden (EN): $($existingPermEN.AccessRights)" -Type "Info"
                }
            }
    }
    catch {
            Write-DebugMessage "Fehler bei der Prüfung bestehender Berechtigungen: $($_.Exception.Message)" -Type "Warning"
        }
        
        # Falls noch kein identifizierter Kalender, versuchen wir die Kalender zu prüfen ohne Benutzerberechtigungen
        if ($null -eq $identity) {
            try {
                # Prüfen, ob der deutsche Kalender existiert
                $deExists = Get-MailboxFolderPermission -Identity $identityDE -ErrorAction SilentlyContinue
                if ($null -ne $deExists) {
                    $identity = $identityDE
                    Write-DebugMessage "Deutscher Kalenderordner gefunden: $identityDE" -Type "Info"
                }
                else {
                    # Prüfen, ob der englische Kalender existiert
                    $enExists = Get-MailboxFolderPermission -Identity $identityEN -ErrorAction SilentlyContinue
                    if ($null -ne $enExists) {
                        $identity = $identityEN
                        Write-DebugMessage "Englischer Kalenderordner gefunden: $identityEN" -Type "Info"
                    }
                }
            }
            catch {
                Write-DebugMessage "Fehler beim Prüfen der Kalenderordner: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Falls immer noch kein Kalender gefunden, über Statistiken suchen
        if ($null -eq $identity) {
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
                Write-DebugMessage "Fehler beim Suchen des Kalenderordners über FolderStatistics: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Wenn immer noch kein Kalender gefunden, Exception werfen
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
        $success = $false
        switch ($GroupType) {
            "Verteilergruppe" {
                New-DistributionGroup @params -Type "Distribution" -ErrorAction Stop
                $success = $true
            }
            "Sicherheitsgruppe" {
                New-DistributionGroup @params -Type "Security" -ErrorAction Stop
                $success = $true
            }
            "Microsoft 365-Gruppe" {
                # Für Microsoft 365-Gruppen ggf. einen anderen Cmdlet verwenden
                New-UnifiedGroup @params -ErrorAction Stop
                $success = $true
            }
            "E-Mail-aktivierte Sicherheitsgruppe" {
                New-DistributionGroup @params -Type "Security" -ErrorAction Stop
                $success = $true
            }
            default {
                throw "Unbekannter Gruppentyp: $GroupType"
            }
        }
        
        # Wenn die Gruppe erstellt wurde und Members angegeben wurden, diese hinzufügen
        if ($success -and -not [string]::IsNullOrWhiteSpace($Members)) {
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

# Neue Funktion zum Aktualisieren der Domain-Liste
function Update-DomainList {
    [CmdletBinding()]
    param()
    
    try {
        Write-DebugMessage "Aktualisiere Domain-Liste für die ComboBox" -Type "Info"
        
        # Prüfen, ob die ComboBox existiert
        if ($null -eq $script:cmbSharedMailboxDomain) {
            $script:cmbSharedMailboxDomain = Get-XamlElement -ElementName "cmbSharedMailboxDomain"
            if ($null -eq $script:cmbSharedMailboxDomain) {
                Write-DebugMessage "Domain-ComboBox nicht gefunden" -Type "Warning"
                return $false
            }
        }
        
        # Prüfen, ob eine Verbindung besteht
        if (-not $script:isConnected) {
            Write-DebugMessage "Keine Exchange-Verbindung für Domain-Abfrage" -Type "Warning"
            return $false
        }
        
        # Domains abrufen und ComboBox befüllen
        $domains = Get-AcceptedDomain | Select-Object -ExpandProperty DomainName
        
        # Dispatcher verwenden für Thread-Sicherheit
        $script:cmbSharedMailboxDomain.Dispatcher.Invoke([Action]{
            $script:cmbSharedMailboxDomain.Items.Clear()
            foreach ($domain in $domains) {
                [void]$script:cmbSharedMailboxDomain.Items.Add($domain)
            }
            if ($script:cmbSharedMailboxDomain.Items.Count -gt 0) {
                $script:cmbSharedMailboxDomain.SelectedIndex = 0
            }
        }, "Normal")
        
        Write-DebugMessage "Domain-Liste erfolgreich aktualisiert: $($domains.Count) Domains geladen" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Aktualisieren der Domain-Liste: $errorMsg" -Type "Error"
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
    function Initialize-ContactsTab {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Kontakte-Tab" -Type "Info"
            
            # UI-Elemente referenzieren
            $txtContactExternalEmail = Get-XamlElement -ElementName "txtContactExternalEmail" -Required
            $txtContactSearch = Get-XamlElement -ElementName "txtContactSearch" -Required
            $cmbContactType = Get-XamlElement -ElementName "cmbContactType" -Required
            $btnCreateContact = Get-XamlElement -ElementName "btnCreateContact" -Required
            $btnShowMailContacts = Get-XamlElement -ElementName "btnShowMailContacts" -Required
            $btnShowMailUsers = Get-XamlElement -ElementName "btnShowMailUsers" -Required
            $btnRemoveContact = Get-XamlElement -ElementName "btnRemoveContact" -Required
            $lstContacts = Get-XamlElement -ElementName "lstContacts" -Required
            $btnExportContacts = Get-XamlElement -ElementName "btnExportContacts" -Required
            
            # Globale Variablen setzen
            $script:txtContactExternalEmail = $txtContactExternalEmail
            $script:txtContactSearch = $txtContactSearch
            $script:cmbContactType = $cmbContactType
            $script:lstContacts = $lstContacts
            
            # Event-Handler für Buttons registrieren
            Write-DebugMessage "Event-Handler für btnCreateContact registrieren" -Type "Info"
            Register-EventHandler -Control $btnCreateContact -Handler {
                    $externalEmail = $script:txtContactExternalEmail.Text.Trim()
                    if ([string]::IsNullOrEmpty($externalEmail)) {
                        Show-MessageBox -Message "Bitte geben Sie eine externe E-Mail-Adresse ein." -Title "Fehlende Angaben" -Type "Warning"
                        return
                    }
                    
                    $contactType = $script:cmbContactType.SelectedIndex
                    if ($contactType -eq 0) { # MailContact
                        try {
                            $displayName = $externalEmail.Split('@')[0]
                            $script:txtStatus.Text = "Erstelle MailContact $displayName..."
                            New-MailContact -Name $displayName -ExternalEmailAddress $externalEmail -FirstName $displayName -DisplayName $displayName -ErrorAction Stop
                            $script:txtStatus.Text = "MailContact erfolgreich erstellt: $displayName"
                            Show-MessageBox -Message "MailContact $displayName wurde erfolgreich erstellt." -Title "Erfolg" -Type "Info"
                            $script:txtContactExternalEmail.Text = ""
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Erstellen des MailContact: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Erstellen des MailContact: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                    }
                    else { # MailUser
                        try {
                            $displayName = $externalEmail.Split('@')[0]
                            $script:txtStatus.Text = "Erstelle MailUser $displayName..."
                            New-MailUser -Name $displayName -ExternalEmailAddress $externalEmail -MicrosoftOnlineServicesID "$displayName@$($script:primaryDomain)" -FirstName $displayName -DisplayName $displayName -ErrorAction Stop
                            $script:txtStatus.Text = "MailUser erfolgreich erstellt: $displayName"
                            Show-MessageBox -Message "MailUser $displayName wurde erfolgreich erstellt." -Title "Erfolg" -Type "Info"
                            $script:txtContactExternalEmail.Text = ""
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Erstellen des MailUser: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Erstellen des MailUser: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                }
            } -ControlName "btnCreateContact"
            Write-DebugMessage "Event-Handler für btnShowMailContacts registrieren" -Type "Info"
            Register-EventHandler -Control $btnShowMailContacts -Handler {
                # Verbindungsprüfung - nur bei Bedarf
                if (-not $script:isConnected) {
                    [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                
                $script:txtStatus.Text = "Lade MailContacts..."
                try {
                    # Versuche einen einfachen Exchange-Befehl zur Verbindungsprüfung
                    Get-OrganizationConfig -ErrorAction Stop | Out-Null
                    
                    # Exchange-Befehle ausführen
                    $contacts = Get-Recipient -RecipientType MailContact -ResultSize Unlimited | 
                        Select-Object DisplayName, PrimarySmtpAddress, ExternalEmailAddress, RecipientTypeDetails
                    
                    # Als Array behandeln
                    if ($null -ne $contacts -and -not ($contacts -is [Array])) {
                        $contacts = @($contacts)
                    }
                    
                    $script:lstContacts.ItemsSource = $contacts
                    $script:txtStatus.Text = "$($contacts.Count) MailContacts gefunden"
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    
                    # Wenn der Fehler auf eine abgelaufene Verbindung hindeutet
                    if ($errorMsg -like "*Connect-ExchangeOnline*" -or 
                        $errorMsg -like "*session*" -or 
                        $errorMsg -like "*authentication*") {
                        
                        # Status korrigieren
                        $script:isConnected = $false
                        $script:txtConnectionStatus.Text = "Nicht verbunden"
                        $script:txtConnectionStatus.Foreground = $script:disconnectedBrush
                        
                        [System.Windows.MessageBox]::Show("Die Verbindung zu Exchange Online ist unterbrochen. Bitte stellen Sie die Verbindung erneut her.", 
                            "Verbindungsfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                    else {
                        $script:txtStatus.Text = "Fehler beim Laden der MailContacts: $errorMsg"
                        [System.Windows.MessageBox]::Show("Fehler beim Laden der MailContacts: $errorMsg", 
                            "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    }
                }
            } -ControlName "btnShowMailContacts"
            
            Write-DebugMessage "Event-Handler für btnShowMailUsers registrieren" -Type "Info"

            Register-EventHandler -Control $btnShowMailUsers -Handler {
                # Überprüfe, ob wir mit Exchange verbunden sind
                if (-not (Confirm-ExchangeConnection)) {
                    Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                    return
                }
                
                $script:txtStatus.Text = "Lade MailUsers..."
                try {
                    $users = Get-MailUser -ResultSize Unlimited | 
                        Select-Object DisplayName, PrimarySmtpAddress, ExternalEmailAddress, RecipientTypeDetails
                    
                    $script:lstContacts.ItemsSource = $users
                    $script:txtStatus.Text = "$($users.Count) MailUsers gefunden"
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $script:txtStatus.Text = "Fehler beim Laden der MailUsers: $errorMsg"
                    Show-MessageBox -Message "Fehler beim Laden der MailUsers: $errorMsg" -Title "Fehler" -Type "Error"
                }
            } -ControlName "btnShowMailUsers"
            
            Write-DebugMessage "Event-Handler für btnRemoveContact registrieren" -Type "Info"
            Register-EventHandler -Control $btnRemoveContact -Handler {
                    if ($script:lstContacts.SelectedItem -eq $null) {
                        Show-MessageBox -Message "Bitte wählen Sie zuerst einen Kontakt aus." -Title "Kein Kontakt ausgewählt" -Type "Warning"
                        return
                    }
                    
                    $selectedContact = $script:lstContacts.SelectedItem
                    $contactType = $selectedContact.RecipientTypeDetails
                    $contactName = $selectedContact.DisplayName
                    $contactEmail = $selectedContact.PrimarySmtpAddress
                    
                    $confirmResult = Show-MessageBox -Message "Sind Sie sicher, dass Sie den Kontakt '$contactName' löschen möchten?" `
                        -Title "Kontakt löschen" -Type "YesNo"
                    
                    if ($confirmResult -eq "Yes") {
                        try {
                            $script:txtStatus.Text = "Lösche Kontakt $contactName..."
                            
                            if ($contactType -like "*MailContact*") {
                                Remove-MailContact -Identity $contactEmail -Confirm:$false -ErrorAction Stop
                            }
                            elseif ($contactType -like "*MailUser*") {
                                Remove-MailUser -Identity $contactEmail -Confirm:$false -ErrorAction Stop
                            }
                            
                            $script:txtStatus.Text = "Kontakt $contactName erfolgreich gelöscht"
                            Show-MessageBox -Message "Der Kontakt '$contactName' wurde erfolgreich gelöscht." -Title "Kontakt gelöscht" -Type "Info"
                            
                            # Liste aktualisieren
                            if ($contactType -like "*MailContact*") {
                                $script:btnShowMailContacts.RaiseEvent([System.Windows.RoutedEventArgs]::Click)
                            }
                            else {
                                $script:btnShowMailUsers.RaiseEvent([System.Windows.RoutedEventArgs]::Click)
                            }
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Löschen des Kontakts: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Löschen des Kontakts: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                }
            } -ControlName "btnRemoveContact"
            
            Write-DebugMessage "Event-Handler für btnExportContacts registrieren" -Type "Info"
            Register-EventHandler -Control $btnExportContacts -Handler {
                    if ($script:lstContacts.Items.Count -eq 0) {
                        Show-MessageBox -Message "Es gibt keine Kontakte zum Exportieren." -Title "Keine Daten" -Type "Info"
                        return
                    }
                    
                    $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
                    $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv"
                    $saveFileDialog.Title = "Kontakte exportieren"
                    $saveFileDialog.FileName = "Kontakte_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                    
                    if ($saveFileDialog.ShowDialog() -eq $true) {
                        try {
                            $script:txtStatus.Text = "Exportiere Kontakte nach $($saveFileDialog.FileName)..."
                            $script:lstContacts.ItemsSource | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                            $script:txtStatus.Text = "Kontakte erfolgreich exportiert: $($saveFileDialog.FileName)"
                            Show-MessageBox -Message "Die Kontakte wurden erfolgreich nach '$($saveFileDialog.FileName)' exportiert." -Title "Export erfolgreich" -Type "Info"
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Exportieren der Kontakte: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Exportieren der Kontakte: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                    }
            } -ControlName "btnExportContacts"
            
            Write-DebugMessage "Kontakte-Tab wurde initialisiert" -Type "Success"
            return $true
                }
                catch {
                    $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler bei der Initialisierung des Kontakte-Tabs: $errorMsg" -Type "Error"
            return $false
        }
    }
    
    # Implementiert die Hauptfunktionen für den Kontakte-Tab
    
    function New-ExoContact {
        [CmdletBinding()]
        param()
        
        try {
            $name = $script:txtContactName.Text.Trim()
            $email = $script:txtContactEmail.Text.Trim()
            
            if ([string]::IsNullOrEmpty($name) -or [string]::IsNullOrEmpty($email)) {
                Show-MessageBox -Message "Bitte geben Sie einen Namen und eine E-Mail-Adresse ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Erstelle Kontakt $name mit E-Mail $email..."
            
            # Exchange-Befehle ausführen
            $newMailContact = New-MailContact -Name $name -ExternalEmailAddress $email -ErrorAction Stop
            
            if ($newMailContact) {
                Write-StatusMessage "Kontakt $name erfolgreich erstellt."
                Show-MessageBox -Message "Der Kontakt $name wurde erfolgreich erstellt." -Title "Kontakt erstellt" -Type "Info"
                
                # Kontakte aktualisieren
                Get-ExoContacts
                
                # Felder leeren
                $script:txtContactName.Text = ""
                $script:txtContactEmail.Text = ""
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Erstellen des Kontakts: $errorMsg"
            Show-MessageBox -Message "Fehler beim Erstellen des Kontakts: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
    function Update-ExoContact {
        [CmdletBinding()]
        param()
        
        try {
            $name = $script:txtContactName.Text.Trim()
            $email = $script:txtContactEmail.Text.Trim()
            
            if ([string]::IsNullOrEmpty($name) -or [string]::IsNullOrEmpty($email)) {
                Show-MessageBox -Message "Bitte geben Sie einen Namen und eine E-Mail-Adresse ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            if ($script:dgContacts.SelectedItem) {
                $contactObj = $script:dgContacts.SelectedItem
                $contactId = $contactObj.Identity
                
                Write-StatusMessage "Aktualisiere Kontakt $contactId..."
                
                # Exchange-Befehle ausführen
                Set-MailContact -Identity $contactId -Name $name -ExternalEmailAddress $email -ErrorAction Stop
                
                Write-StatusMessage "Kontakt $name erfolgreich aktualisiert."
                Show-MessageBox -Message "Der Kontakt $name wurde erfolgreich aktualisiert." -Title "Kontakt aktualisiert" -Type "Info"
                
                # Kontakte aktualisieren
                Get-ExoContacts
                
                # Felder leeren
                $script:txtContactName.Text = ""
                $script:txtContactEmail.Text = ""
            }
            else {
                Show-MessageBox -Message "Bitte wählen Sie zuerst einen Kontakt aus der Liste aus." -Title "Kein Kontakt ausgewählt" -Type "Warning"
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Aktualisieren des Kontakts: $errorMsg"
            Show-MessageBox -Message "Fehler beim Aktualisieren des Kontakts: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
    function Search-ExoContacts {
        [CmdletBinding()]
        param()
        
        try {
            $searchTerm = $txtContactSearch.Text.Trim()
            
            if ([string]::IsNullOrEmpty($searchTerm)) {
                Show-MessageBox -Message "Bitte geben Sie einen Suchbegriff ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Suche nach Kontakten mit '$searchTerm'..."
            
            # Exchange-Befehle ausführen
            $contacts = Get-MailContact -Filter "Name -like '*$searchTerm*' -or EmailAddresses -like '*$searchTerm*'" -ResultSize Unlimited | 
                Select-Object Name, DisplayName, PrimarySmtpAddress, ExternalEmailAddress, Department, Title
            
            # Daten zum DataGrid hinzufügen
            $script:dgContacts.Dispatcher.Invoke([action]{
                $script:dgContacts.ItemsSource = $contacts
            }, "Normal")
            
            Write-StatusMessage "Suche abgeschlossen. $($contacts.Count) Kontakte gefunden."
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler bei der Kontaktsuche: $errorMsg"
            Show-MessageBox -Message "Fehler bei der Kontaktsuche: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    

# Funktion zur Überprüfung der Exchange Online-Verbindung
function Confirm-ExchangeConnection {
    [CmdletBinding()]
    param()
    
    try {
        # Überprüfen, ob die globale Verbindungsvariable gesetzt ist
        if ($Global:IsConnectedToExo -eq $true) {
            # Verbindung testen durch Abrufen einer Exchange-Information
            try {
                $null = Get-OrganizationConfig -ErrorAction Stop
                return $true
            }
            catch {
                $Global:IsConnectedToExo = $false
                Write-DebugMessage "Exchange Online Verbindung getrennt: $($_.Exception.Message)" -Type "Warning"
                return $false
            }
        }
        else {
            return $false
        }
    }
    catch {
        $Global:IsConnectedToExo = $false
        Write-DebugMessage "Fehler bei der Überprüfung der Exchange Online-Verbindung: $($_.Exception.Message)" -Type "Error"
        return $false
    }
}
function Ensure-ExchangeConnection {
    # Prüfen, ob eine gültige Verbindung besteht
    if (-not (Confirm-ExchangeConnection)) {
        $script:txtStatus.Text = "Verbindung zu Exchange Online wird hergestellt..."
        try {
            # Verbindung herstellen
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $script:isConnected = $true
            $script:txtStatus.Text = "Verbindung zu Exchange Online hergestellt"
            return $true
        }
        catch {
            $script:txtStatus.Text = "Fehler beim Verbinden mit Exchange Online: $($_.Exception.Message)"
            return $false
        }
    }
    return $true
}

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

# MessageBox-Funktion
function Show-MessageBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Title = "Information",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "Question")]
        [string]$Type = "Info"
    )
    
    try {
        $icon = switch ($Type) {
            "Info" { [System.Windows.MessageBoxImage]::Information }
            "Warning" { [System.Windows.MessageBoxImage]::Warning }
            "Error" { [System.Windows.MessageBoxImage]::Error }
            "Question" { [System.Windows.MessageBoxImage]::Question }
        }
        
        $buttons = if ($Type -eq "Question") { 
            [System.Windows.MessageBoxButton]::YesNo 
        } else { 
            [System.Windows.MessageBoxButton]::OK 
        }
        
        $result = [System.Windows.MessageBox]::Show($Message, $Title, $buttons, $icon)
        
        # Erfolg loggen
        Write-DebugMessage -Message "$Title - $Type - $Message" -Type "Info"
        
        # Ergebnis zurückgeben (wichtig für Ja/Nein-Fragen)
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage -Message "Fehler beim Anzeigen der MessageBox: $errorMsg" -Type "Error"
        
        # Fallback-Ausgabe
        Write-Host "Meldung ($Type): $Title - $Message" -ForegroundColor Red
        
        if ($Type -eq "Question") {
            return [System.Windows.MessageBoxResult]::No
        }
    }
}

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
        
        # GUI-Element im UI-Thread aktualisieren mit Überprüfung des Dispatcher-Status
        if ($null -ne $TextElement.Dispatcher -and $TextElement.Dispatcher.CheckAccess()) {
            # Wir sind bereits im UI-Thread
            $TextElement.Text = $sanitizedMessage
            if ($null -ne $Color) {
                $TextElement.Foreground = $Color
            }
        } 
        else {
            # Dispatcher verwenden für Thread-Sicherheit
            $TextElement.Dispatcher.Invoke([Action]{
                $TextElement.Text = $sanitizedMessage
                if ($null -ne $Color) {
                    $TextElement.Foreground = $Color
                }
            }, "Normal")
        }
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

# Funktion zum Aktualisieren des Status in der GUI
function Write-StatusMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Type = "Info"
    )
    
    try {
        # Logge die Nachricht auch
        Write-DebugMessage -Message $Message -Type $Type
        
        # Bestimme die Farbe basierend auf dem Nachrichtentyp
        $color = switch ($Type) {
            "Success" { $script:connectedBrush }
            "Error" { $script:disconnectedBrush }
            "Warning" { New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Orange) }
            "Info" { $null }
            default { $null }
        }
        
        # Aktualisiere das Status-Textfeld in der GUI
        if ($null -ne $script:txtStatus) {
            Update-GuiText -TextElement $script:txtStatus -Message $Message -Color $color
        }
    } 
    catch {
        # Bei Fehler einfach eine Debug-Meldung ausgeben
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler in Write-StatusMessage: $errorMsg" -Type "Error"
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
# Abschnitt: Eingabevalidierung
# -------------------------------------------------
function  Validate-Email{
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
            
            # Verbindung immer überprüfen
            try {
                # Einfache Prüfung durch Ausführen eines kleinen Exchange-Befehls
                Get-OrganizationConfig -ErrorAction Stop | Out-Null
                Write-DebugMessage "Bestehende Exchange-Verbindung ist gültig" -Type "Info"
            }
            catch {
                Write-DebugMessage "Bestehende Verbindung ist nicht mehr gültig, stelle neue Verbindung her" -Type "Warning"
                $Global:IsConnectedToExo = $false
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

function Test-ExchangeOnlineConnection {
    [CmdletBinding()]
    param()
    
    try {
        # Zuerst prüfen, ob die globale Testfunktion verfügbar ist
        if ($null -ne ${global:EasyStartup_TestExchangeOnlineConnection}) {
            return & ${global:EasyStartup_TestExchangeOnlineConnection}
        }
        
        # Eigene Verbindungsprüfung als Fallback
        try {
            # Einfacher Test durch Ausführung eines Exchange-Befehls
            Get-OrganizationConfig -ErrorAction Stop | Out-Null
            Write-DebugMessage "Exchange Online-Verbindung ist aktiv" -Type "Info"
            return $true
        }
        catch {
            Write-DebugMessage "Exchange Online-Verbindung ist nicht mehr gültig: $($_.Exception.Message)" -Type "Warning"
            return $false
        }
    }
    catch {
        Write-DebugMessage "Fehler beim Testen der Exchange Online-Verbindung: $($_.Exception.Message)" -Type "Error"
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
            # Frage den Benutzer ob er das wirklich tun möchte
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
            
            $successCount = 0
            $errorCount = 0

            if ($PermissionType -eq "Standard" -or $PermissionType -eq "Beides") {
                $result = Set-DefaultCalendarPermissionForAll -AccessRights $AccessRights
                if ($result) { $successCount++ } else { $errorCount++ }
            }
            if ($PermissionType -eq "Anonym" -or $PermissionType -eq "Beides") {
                $result = Set-AnonymousCalendarPermissionForAll -AccessRights $AccessRights
                if ($result) { $successCount++ } else { $errorCount++ }
            }
        }
        else {         
            if ([string]::IsNullOrWhiteSpace($MailboxUser) -and 
                $null -ne $script:txtCalendarMailboxUser -and 
                -not [string]::IsNullOrWhiteSpace($script:txtCalendarMailboxUser.Text)) {
                $mailboxUser = $script:txtCalendarMailboxUser.Text.Trim()
            }
            
            if ([string]::IsNullOrWhiteSpace($mailboxUser)) {
                throw "Keine Postfach-E-Mail-Adresse angegeben"
            }
            
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
        
        # Prüfe ob Berechtigung bereits existiert und ermittle den korrekten Kalenderordner
        $calendarExists = $false
        $identityDE = "${SourceUser}:\Kalender"
        $identityEN = "${SourceUser}:\Calendar"
        $identity = $null
        
        # Systematisch nach dem richtigen Kalender suchen
        try {
            # Zuerst versuchen wir den deutschen Kalender
            $existingPermDE = Get-MailboxFolderPermission -Identity $identityDE -User $TargetUser -ErrorAction SilentlyContinue
            if ($null -ne $existingPermDE) {
                $calendarExists = $true
                $identity = $identityDE
                Write-DebugMessage "Bestehende Berechtigung gefunden (DE): $($existingPermDE.AccessRights)" -Type "Info"
            }
            else {
                # Dann den englischen Kalender probieren
                $existingPermEN = Get-MailboxFolderPermission -Identity $identityEN -User $TargetUser -ErrorAction SilentlyContinue
                if ($null -ne $existingPermEN) {
                    $calendarExists = $true
                    $identity = $identityEN
                    Write-DebugMessage "Bestehende Berechtigung gefunden (EN): $($existingPermEN.AccessRights)" -Type "Info"
                }
            }
    }
    catch {
            Write-DebugMessage "Fehler bei der Prüfung bestehender Berechtigungen: $($_.Exception.Message)" -Type "Warning"
        }
        
        # Falls noch kein identifizierter Kalender, versuchen wir die Kalender zu prüfen ohne Benutzerberechtigungen
        if ($null -eq $identity) {
            try {
                # Prüfen, ob der deutsche Kalender existiert
                $deExists = Get-MailboxFolderPermission -Identity $identityDE -ErrorAction SilentlyContinue
                if ($null -ne $deExists) {
                    $identity = $identityDE
                    Write-DebugMessage "Deutscher Kalenderordner gefunden: $identityDE" -Type "Info"
                }
                else {
                    # Prüfen, ob der englische Kalender existiert
                    $enExists = Get-MailboxFolderPermission -Identity $identityEN -ErrorAction SilentlyContinue
                    if ($null -ne $enExists) {
                        $identity = $identityEN
                        Write-DebugMessage "Englischer Kalenderordner gefunden: $identityEN" -Type "Info"
                    }
                }
            }
            catch {
                Write-DebugMessage "Fehler beim Prüfen der Kalenderordner: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Falls immer noch kein Kalender gefunden, über Statistiken suchen
        if ($null -eq $identity) {
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
                Write-DebugMessage "Fehler beim Suchen des Kalenderordners über FolderStatistics: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Wenn immer noch kein Kalender gefunden, Exception werfen
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

# MessageBox-Funktion
function Show-MessageBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Title = "Information",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "Question")]
        [string]$Type = "Info"
    )
    
    try {
        $icon = switch ($Type) {
            "Info" { [System.Windows.MessageBoxImage]::Information }
            "Warning" { [System.Windows.MessageBoxImage]::Warning }
            "Error" { [System.Windows.MessageBoxImage]::Error }
            "Question" { [System.Windows.MessageBoxImage]::Question }
        }
        
        $buttons = if ($Type -eq "Question") { 
            [System.Windows.MessageBoxButton]::YesNo 
        } else { 
            [System.Windows.MessageBoxButton]::OK 
        }
        
        $result = [System.Windows.MessageBox]::Show($Message, $Title, $buttons, $icon)
        
        # Erfolg loggen
        Write-DebugMessage -Message "$Title - $Type - $Message" -Type "Info"
        
        # Ergebnis zurückgeben (wichtig für Ja/Nein-Fragen)
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage -Message "Fehler beim Anzeigen der MessageBox: $errorMsg" -Type "Error"
        
        # Fallback-Ausgabe
        Write-Host "Meldung ($Type): $Title - $Message" -ForegroundColor Red
        
        if ($Type -eq "Question") {
            return [System.Windows.MessageBoxResult]::No
        }
    }
}

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
        
        # GUI-Element im UI-Thread aktualisieren mit Überprüfung des Dispatcher-Status
        if ($null -ne $TextElement.Dispatcher -and $TextElement.Dispatcher.CheckAccess()) {
            # Wir sind bereits im UI-Thread
            $TextElement.Text = $sanitizedMessage
            if ($null -ne $Color) {
                $TextElement.Foreground = $Color
            }
        } 
        else {
            # Dispatcher verwenden für Thread-Sicherheit
            $TextElement.Dispatcher.Invoke([Action]{
                $TextElement.Text = $sanitizedMessage
                if ($null -ne $Color) {
                    $TextElement.Foreground = $Color
                }
            }, "Normal")
        }
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

# Funktion zum Aktualisieren des Status in der GUI
function Write-StatusMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Type = "Info"
    )
    
    try {
        # Logge die Nachricht auch
        Write-DebugMessage -Message $Message -Type $Type
        
        # Bestimme die Farbe basierend auf dem Nachrichtentyp
        $color = switch ($Type) {
            "Success" { $script:connectedBrush }
            "Error" { $script:disconnectedBrush }
            "Warning" { New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Orange) }
            "Info" { $null }
            default { $null }
        }
        
        # Aktualisiere das Status-Textfeld in der GUI
        if ($null -ne $script:txtStatus) {
            Update-GuiText -TextElement $script:txtStatus -Message $Message -Color $color
        }
    } 
    catch {
        # Bei Fehler einfach eine Debug-Meldung ausgeben
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler in Write-StatusMessage: $errorMsg" -Type "Error"
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
# Abschnitt: Eingabevalidierung
# -------------------------------------------------
function  Validate-Email{
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
            
            # Verbindung immer überprüfen
            try {
                # Einfache Prüfung durch Ausführen eines kleinen Exchange-Befehls
                Get-OrganizationConfig -ErrorAction Stop | Out-Null
                Write-DebugMessage "Bestehende Exchange-Verbindung ist gültig" -Type "Info"
            }
            catch {
                Write-DebugMessage "Bestehende Verbindung ist nicht mehr gültig, stelle neue Verbindung her" -Type "Warning"
                $Global:IsConnectedToExo = $false
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

function Test-ExchangeOnlineConnection {
    [CmdletBinding()]
    param()
    
    try {
        # Zuerst prüfen, ob die globale Testfunktion verfügbar ist
        if ($null -ne ${global:EasyStartup_TestExchangeOnlineConnection}) {
            return & ${global:EasyStartup_TestExchangeOnlineConnection}
        }
        
        # Eigene Verbindungsprüfung als Fallback
        try {
            # Einfacher Test durch Ausführung eines Exchange-Befehls
            Get-OrganizationConfig -ErrorAction Stop | Out-Null
            Write-DebugMessage "Exchange Online-Verbindung ist aktiv" -Type "Info"
            return $true
        }
        catch {
            Write-DebugMessage "Exchange Online-Verbindung ist nicht mehr gültig: $($_.Exception.Message)" -Type "Warning"
            return $false
        }
    }
    catch {
        Write-DebugMessage "Fehler beim Testen der Exchange Online-Verbindung: $($_.Exception.Message)" -Type "Error"
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
            # Frage den Benutzer ob er das wirklich tun möchte
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
            
            $successCount = 0
            $errorCount = 0

            if ($PermissionType -eq "Standard" -or $PermissionType -eq "Beides") {
                $result = Set-DefaultCalendarPermissionForAll -AccessRights $AccessRights
                if ($result) { $successCount++ } else { $errorCount++ }
            }
            if ($PermissionType -eq "Anonym" -or $PermissionType -eq "Beides") {
                $result = Set-AnonymousCalendarPermissionForAll -AccessRights $AccessRights
                if ($result) { $successCount++ } else { $errorCount++ }
            }
        }
        else {         
            if ([string]::IsNullOrWhiteSpace($MailboxUser) -and 
                $null -ne $script:txtCalendarMailboxUser -and 
                -not [string]::IsNullOrWhiteSpace($script:txtCalendarMailboxUser.Text)) {
                $mailboxUser = $script:txtCalendarMailboxUser.Text.Trim()
            }
            
            if ([string]::IsNullOrWhiteSpace($mailboxUser)) {
                throw "Keine Postfach-E-Mail-Adresse angegeben"
            }
            
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
        
        # Prüfe ob Berechtigung bereits existiert und ermittle den korrekten Kalenderordner
        $calendarExists = $false
        $identityDE = "${SourceUser}:\Kalender"
        $identityEN = "${SourceUser}:\Calendar"
        $identity = $null
        
        # Systematisch nach dem richtigen Kalender suchen
        try {
            # Zuerst versuchen wir den deutschen Kalender
            $existingPermDE = Get-MailboxFolderPermission -Identity $identityDE -User $TargetUser -ErrorAction SilentlyContinue
            if ($null -ne $existingPermDE) {
                $calendarExists = $true
                $identity = $identityDE
                Write-DebugMessage "Bestehende Berechtigung gefunden (DE): $($existingPermDE.AccessRights)" -Type "Info"
            }
            else {
                # Dann den englischen Kalender probieren
                $existingPermEN = Get-MailboxFolderPermission -Identity $identityEN -User $TargetUser -ErrorAction SilentlyContinue
                if ($null -ne $existingPermEN) {
                    $calendarExists = $true
                    $identity = $identityEN
                    Write-DebugMessage "Bestehende Berechtigung gefunden (EN): $($existingPermEN.AccessRights)" -Type "Info"
                }
            }
    }
    catch {
            Write-DebugMessage "Fehler bei der Prüfung bestehender Berechtigungen: $($_.Exception.Message)" -Type "Warning"
        }
        
        # Falls noch kein identifizierter Kalender, versuchen wir die Kalender zu prüfen ohne Benutzerberechtigungen
        if ($null -eq $identity) {
            try {
                # Prüfen, ob der deutsche Kalender existiert
                $deExists = Get-MailboxFolderPermission -Identity $identityDE -ErrorAction SilentlyContinue
                if ($null -ne $deExists) {
                    $identity = $identityDE
                    Write-DebugMessage "Deutscher Kalenderordner gefunden: $identityDE" -Type "Info"
                }
                else {
                    # Prüfen, ob der englische Kalender existiert
                    $enExists = Get-MailboxFolderPermission -Identity $identityEN -ErrorAction SilentlyContinue
                    if ($null -ne $enExists) {
                        $identity = $identityEN
                        Write-DebugMessage "Englischer Kalenderordner gefunden: $identityEN" -Type "Info"
                    }
                }
            }
            catch {
                Write-DebugMessage "Fehler beim Prüfen der Kalenderordner: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Falls immer noch kein Kalender gefunden, über Statistiken suchen
        if ($null -eq $identity) {
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
                Write-DebugMessage "Fehler beim Suchen des Kalenderordners über FolderStatistics: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Wenn immer noch kein Kalender gefunden, Exception werfen
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
        $success = $false
        switch ($GroupType) {
            "Verteilergruppe" {
                New-DistributionGroup @params -Type "Distribution" -ErrorAction Stop
                $success = $true
            }
            "Sicherheitsgruppe" {
                New-DistributionGroup @params -Type "Security" -ErrorAction Stop
                $success = $true
            }
            "Microsoft 365-Gruppe" {
                # Für Microsoft 365-Gruppen ggf. einen anderen Cmdlet verwenden
                New-UnifiedGroup @params -ErrorAction Stop
                $success = $true
            }
            "E-Mail-aktivierte Sicherheitsgruppe" {
                New-DistributionGroup @params -Type "Security" -ErrorAction Stop
                $success = $true
            }
            default {
                throw "Unbekannter Gruppentyp: $GroupType"
            }
        }
        
        # Wenn die Gruppe erstellt wurde und Members angegeben wurden, diese hinzufügen
        if ($success -and -not [string]::IsNullOrWhiteSpace($Members)) {
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

# Neue Funktion zum Aktualisieren der Domain-Liste
function Update-DomainList {
    [CmdletBinding()]
    param()
    
    try {
        Write-DebugMessage "Aktualisiere Domain-Liste für die ComboBox" -Type "Info"
        
        # Prüfen, ob die ComboBox existiert
        if ($null -eq $script:cmbSharedMailboxDomain) {
            $script:cmbSharedMailboxDomain = Get-XamlElement -ElementName "cmbSharedMailboxDomain"
            if ($null -eq $script:cmbSharedMailboxDomain) {
                Write-DebugMessage "Domain-ComboBox nicht gefunden" -Type "Warning"
                return $false
            }
        }
        
        # Prüfen, ob eine Verbindung besteht
        if (-not $script:isConnected) {
            Write-DebugMessage "Keine Exchange-Verbindung für Domain-Abfrage" -Type "Warning"
            return $false
        }
        
        # Domains abrufen und ComboBox befüllen
        $domains = Get-AcceptedDomain | Select-Object -ExpandProperty DomainName
        
        # Dispatcher verwenden für Thread-Sicherheit
        $script:cmbSharedMailboxDomain.Dispatcher.Invoke([Action]{
            $script:cmbSharedMailboxDomain.Items.Clear()
            foreach ($domain in $domains) {
                [void]$script:cmbSharedMailboxDomain.Items.Add($domain)
            }
            if ($script:cmbSharedMailboxDomain.Items.Count -gt 0) {
                $script:cmbSharedMailboxDomain.SelectedIndex = 0
            }
        }, "Normal")
        
        Write-DebugMessage "Domain-Liste erfolgreich aktualisiert: $($domains.Count) Domains geladen" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Aktualisieren der Domain-Liste: $errorMsg" -Type "Error"
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
    function Initialize-ContactsTab {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Kontakte-Tab" -Type "Info"
            
            # UI-Elemente referenzieren
            $txtContactExternalEmail = Get-XamlElement -ElementName "txtContactExternalEmail" -Required
            $txtContactSearch = Get-XamlElement -ElementName "txtContactSearch" -Required
            $cmbContactType = Get-XamlElement -ElementName "cmbContactType" -Required
            $btnCreateContact = Get-XamlElement -ElementName "btnCreateContact" -Required
            $btnShowMailContacts = Get-XamlElement -ElementName "btnShowMailContacts" -Required
            $btnShowMailUsers = Get-XamlElement -ElementName "btnShowMailUsers" -Required
            $btnRemoveContact = Get-XamlElement -ElementName "btnRemoveContact" -Required
            $lstContacts = Get-XamlElement -ElementName "lstContacts" -Required
            $btnExportContacts = Get-XamlElement -ElementName "btnExportContacts" -Required
            
            # Globale Variablen setzen
            $script:txtContactExternalEmail = $txtContactExternalEmail
            $script:txtContactSearch = $txtContactSearch
            $script:cmbContactType = $cmbContactType
            $script:lstContacts = $lstContacts
            
            # Event-Handler für Buttons registrieren
            Write-DebugMessage "Event-Handler für btnCreateContact registrieren" -Type "Info"
            Register-EventHandler -Control $btnCreateContact -Handler {
                    $externalEmail = $script:txtContactExternalEmail.Text.Trim()
                    if ([string]::IsNullOrEmpty($externalEmail)) {
                        Show-MessageBox -Message "Bitte geben Sie eine externe E-Mail-Adresse ein." -Title "Fehlende Angaben" -Type "Warning"
                        return
                    }
                    
                    $contactType = $script:cmbContactType.SelectedIndex
                    if ($contactType -eq 0) { # MailContact
                        try {
                            $displayName = $externalEmail.Split('@')[0]
                            $script:txtStatus.Text = "Erstelle MailContact $displayName..."
                            New-MailContact -Name $displayName -ExternalEmailAddress $externalEmail -FirstName $displayName -DisplayName $displayName -ErrorAction Stop
                            $script:txtStatus.Text = "MailContact erfolgreich erstellt: $displayName"
                            Show-MessageBox -Message "MailContact $displayName wurde erfolgreich erstellt." -Title "Erfolg" -Type "Info"
                            $script:txtContactExternalEmail.Text = ""
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Erstellen des MailContact: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Erstellen des MailContact: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                    }
                    else { # MailUser
                        try {
                            $displayName = $externalEmail.Split('@')[0]
                            $script:txtStatus.Text = "Erstelle MailUser $displayName..."
                            New-MailUser -Name $displayName -ExternalEmailAddress $externalEmail -MicrosoftOnlineServicesID "$displayName@$($script:primaryDomain)" -FirstName $displayName -DisplayName $displayName -ErrorAction Stop
                            $script:txtStatus.Text = "MailUser erfolgreich erstellt: $displayName"
                            Show-MessageBox -Message "MailUser $displayName wurde erfolgreich erstellt." -Title "Erfolg" -Type "Info"
                            $script:txtContactExternalEmail.Text = ""
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Erstellen des MailUser: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Erstellen des MailUser: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                }
            } -ControlName "btnCreateContact"
            Write-DebugMessage "Event-Handler für btnShowMailContacts registrieren" -Type "Info"
            Register-EventHandler -Control $btnShowMailContacts -Handler {
                # Verbindungsprüfung - nur bei Bedarf
                if (-not $script:isConnected) {
                    [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                
                $script:txtStatus.Text = "Lade MailContacts..."
                try {
                    # Versuche einen einfachen Exchange-Befehl zur Verbindungsprüfung
                    Get-OrganizationConfig -ErrorAction Stop | Out-Null
                    
                    # Exchange-Befehle ausführen
                    $contacts = Get-Recipient -RecipientType MailContact -ResultSize Unlimited | 
                        Select-Object DisplayName, PrimarySmtpAddress, ExternalEmailAddress, RecipientTypeDetails
                    
                    # Als Array behandeln
                    if ($null -ne $contacts -and -not ($contacts -is [Array])) {
                        $contacts = @($contacts)
                    }
                    
                    $script:lstContacts.ItemsSource = $contacts
                    $script:txtStatus.Text = "$($contacts.Count) MailContacts gefunden"
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    
                    # Wenn der Fehler auf eine abgelaufene Verbindung hindeutet
                    if ($errorMsg -like "*Connect-ExchangeOnline*" -or 
                        $errorMsg -like "*session*" -or 
                        $errorMsg -like "*authentication*") {
                        
                        # Status korrigieren
                        $script:isConnected = $false
                        $script:txtConnectionStatus.Text = "Nicht verbunden"
                        $script:txtConnectionStatus.Foreground = $script:disconnectedBrush
                        
                        [System.Windows.MessageBox]::Show("Die Verbindung zu Exchange Online ist unterbrochen. Bitte stellen Sie die Verbindung erneut her.", 
                            "Verbindungsfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                    else {
                        $script:txtStatus.Text = "Fehler beim Laden der MailContacts: $errorMsg"
                        [System.Windows.MessageBox]::Show("Fehler beim Laden der MailContacts: $errorMsg", 
                            "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    }
                }
            } -ControlName "btnShowMailContacts"
            
            Write-DebugMessage "Event-Handler für btnShowMailUsers registrieren" -Type "Info"

            Register-EventHandler -Control $btnShowMailUsers -Handler {
                # Überprüfe, ob wir mit Exchange verbunden sind
                if (-not (Confirm-ExchangeConnection)) {
                    Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                    return
                }
                
                $script:txtStatus.Text = "Lade MailUsers..."
                try {
                    $users = Get-MailUser -ResultSize Unlimited | 
                        Select-Object DisplayName, PrimarySmtpAddress, ExternalEmailAddress, RecipientTypeDetails
                    
                    $script:lstContacts.ItemsSource = $users
                    $script:txtStatus.Text = "$($users.Count) MailUsers gefunden"
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $script:txtStatus.Text = "Fehler beim Laden der MailUsers: $errorMsg"
                    Show-MessageBox -Message "Fehler beim Laden der MailUsers: $errorMsg" -Title "Fehler" -Type "Error"
                }
            } -ControlName "btnShowMailUsers"
            
            Write-DebugMessage "Event-Handler für btnRemoveContact registrieren" -Type "Info"
            Register-EventHandler -Control $btnRemoveContact -Handler {
                    if ($script:lstContacts.SelectedItem -eq $null) {
                        Show-MessageBox -Message "Bitte wählen Sie zuerst einen Kontakt aus." -Title "Kein Kontakt ausgewählt" -Type "Warning"
                        return
                    }
                    
                    $selectedContact = $script:lstContacts.SelectedItem
                    $contactType = $selectedContact.RecipientTypeDetails
                    $contactName = $selectedContact.DisplayName
                    $contactEmail = $selectedContact.PrimarySmtpAddress
                    
                    $confirmResult = Show-MessageBox -Message "Sind Sie sicher, dass Sie den Kontakt '$contactName' löschen möchten?" `
                        -Title "Kontakt löschen" -Type "YesNo"
                    
                    if ($confirmResult -eq "Yes") {
                        try {
                            $script:txtStatus.Text = "Lösche Kontakt $contactName..."
                            
                            if ($contactType -like "*MailContact*") {
                                Remove-MailContact -Identity $contactEmail -Confirm:$false -ErrorAction Stop
                            }
                            elseif ($contactType -like "*MailUser*") {
                                Remove-MailUser -Identity $contactEmail -Confirm:$false -ErrorAction Stop
                            }
                            
                            $script:txtStatus.Text = "Kontakt $contactName erfolgreich gelöscht"
                            Show-MessageBox -Message "Der Kontakt '$contactName' wurde erfolgreich gelöscht." -Title "Kontakt gelöscht" -Type "Info"
                            
                            # Liste aktualisieren
                            if ($contactType -like "*MailContact*") {
                                $script:btnShowMailContacts.RaiseEvent([System.Windows.RoutedEventArgs]::Click)
                            }
                            else {
                                $script:btnShowMailUsers.RaiseEvent([System.Windows.RoutedEventArgs]::Click)
                            }
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Löschen des Kontakts: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Löschen des Kontakts: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                }
            } -ControlName "btnRemoveContact"
            
            Write-DebugMessage "Event-Handler für btnExportContacts registrieren" -Type "Info"
            Register-EventHandler -Control $btnExportContacts -Handler {
                    if ($script:lstContacts.Items.Count -eq 0) {
                        Show-MessageBox -Message "Es gibt keine Kontakte zum Exportieren." -Title "Keine Daten" -Type "Info"
                        return
                    }
                    
                    $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
                    $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv"
                    $saveFileDialog.Title = "Kontakte exportieren"
                    $saveFileDialog.FileName = "Kontakte_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                    
                    if ($saveFileDialog.ShowDialog() -eq $true) {
                        try {
                            $script:txtStatus.Text = "Exportiere Kontakte nach $($saveFileDialog.FileName)..."
                            $script:lstContacts.ItemsSource | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                            $script:txtStatus.Text = "Kontakte erfolgreich exportiert: $($saveFileDialog.FileName)"
                            Show-MessageBox -Message "Die Kontakte wurden erfolgreich nach '$($saveFileDialog.FileName)' exportiert." -Title "Export erfolgreich" -Type "Info"
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Exportieren der Kontakte: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Exportieren der Kontakte: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                    }
            } -ControlName "btnExportContacts"
            
            Write-DebugMessage "Kontakte-Tab wurde initialisiert" -Type "Success"
            return $true
                }
                catch {
                    $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler bei der Initialisierung des Kontakte-Tabs: $errorMsg" -Type "Error"
            return $false
        }
    }
    
    # Implementiert die Hauptfunktionen für den Kontakte-Tab
    
    function New-ExoContact {
        [CmdletBinding()]
        param()
        
        try {
            $name = $script:txtContactName.Text.Trim()
            $email = $script:txtContactEmail.Text.Trim()
            
            if ([string]::IsNullOrEmpty($name) -or [string]::IsNullOrEmpty($email)) {
                Show-MessageBox -Message "Bitte geben Sie einen Namen und eine E-Mail-Adresse ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Erstelle Kontakt $name mit E-Mail $email..."
            
            # Exchange-Befehle ausführen
            $newMailContact = New-MailContact -Name $name -ExternalEmailAddress $email -ErrorAction Stop
            
            if ($newMailContact) {
                Write-StatusMessage "Kontakt $name erfolgreich erstellt."
                Show-MessageBox -Message "Der Kontakt $name wurde erfolgreich erstellt." -Title "Kontakt erstellt" -Type "Info"
                
                # Kontakte aktualisieren
                Get-ExoContacts
                
                # Felder leeren
                $script:txtContactName.Text = ""
                $script:txtContactEmail.Text = ""
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Erstellen des Kontakts: $errorMsg"
            Show-MessageBox -Message "Fehler beim Erstellen des Kontakts: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
    function Update-ExoContact {
        [CmdletBinding()]
        param()
        
        try {
            $name = $script:txtContactName.Text.Trim()
            $email = $script:txtContactEmail.Text.Trim()
            
            if ([string]::IsNullOrEmpty($name) -or [string]::IsNullOrEmpty($email)) {
                Show-MessageBox -Message "Bitte geben Sie einen Namen und eine E-Mail-Adresse ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            if ($script:dgContacts.SelectedItem) {
                $contactObj = $script:dgContacts.SelectedItem
                $contactId = $contactObj.Identity
                
                Write-StatusMessage "Aktualisiere Kontakt $contactId..."
                
                # Exchange-Befehle ausführen
                Set-MailContact -Identity $contactId -Name $name -ExternalEmailAddress $email -ErrorAction Stop
                
                Write-StatusMessage "Kontakt $name erfolgreich aktualisiert."
                Show-MessageBox -Message "Der Kontakt $name wurde erfolgreich aktualisiert." -Title "Kontakt aktualisiert" -Type "Info"
                
                # Kontakte aktualisieren
                Get-ExoContacts
                
                # Felder leeren
                $script:txtContactName.Text = ""
                $script:txtContactEmail.Text = ""
            }
            else {
                Show-MessageBox -Message "Bitte wählen Sie zuerst einen Kontakt aus der Liste aus." -Title "Kein Kontakt ausgewählt" -Type "Warning"
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Aktualisieren des Kontakts: $errorMsg"
            Show-MessageBox -Message "Fehler beim Aktualisieren des Kontakts: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
    function Search-ExoContacts {
        [CmdletBinding()]
        param()
        
        try {
            $searchTerm = $txtContactSearch.Text.Trim()
            
            if ([string]::IsNullOrEmpty($searchTerm)) {
                Show-MessageBox -Message "Bitte geben Sie einen Suchbegriff ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Suche nach Kontakten mit '$searchTerm'..."
            
            # Exchange-Befehle ausführen
            $contacts = Get-MailContact -Filter "Name -like '*$searchTerm*' -or EmailAddresses -like '*$searchTerm*'" -ResultSize Unlimited | 
                Select-Object Name, DisplayName, PrimarySmtpAddress, ExternalEmailAddress, Department, Title
            
            # Daten zum DataGrid hinzufügen
            $script:dgContacts.Dispatcher.Invoke([action]{
                $script:dgContacts.ItemsSource = $contacts
            }, "Normal")
            
            Write-StatusMessage "Suche abgeschlossen. $($contacts.Count) Kontakte gefunden."
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler bei der Kontaktsuche: $errorMsg"
            Show-MessageBox -Message "Fehler bei der Kontaktsuche: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
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

# MessageBox-Funktion
function Show-MessageBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Title = "Information",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "Question")]
        [string]$Type = "Info"
    )
    
    try {
        $icon = switch ($Type) {
            "Info" { [System.Windows.MessageBoxImage]::Information }
            "Warning" { [System.Windows.MessageBoxImage]::Warning }
            "Error" { [System.Windows.MessageBoxImage]::Error }
            "Question" { [System.Windows.MessageBoxImage]::Question }
        }
        
        $buttons = if ($Type -eq "Question") { 
            [System.Windows.MessageBoxButton]::YesNo 
        } else { 
            [System.Windows.MessageBoxButton]::OK 
        }
        
        $result = [System.Windows.MessageBox]::Show($Message, $Title, $buttons, $icon)
        
        # Erfolg loggen
        Write-DebugMessage -Message "$Title - $Type - $Message" -Type "Info"
        
        # Ergebnis zurückgeben (wichtig für Ja/Nein-Fragen)
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage -Message "Fehler beim Anzeigen der MessageBox: $errorMsg" -Type "Error"
        
        # Fallback-Ausgabe
        Write-Host "Meldung ($Type): $Title - $Message" -ForegroundColor Red
        
        if ($Type -eq "Question") {
            return [System.Windows.MessageBoxResult]::No
        }
    }
}

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
        
        # GUI-Element im UI-Thread aktualisieren mit Überprüfung des Dispatcher-Status
        if ($null -ne $TextElement.Dispatcher -and $TextElement.Dispatcher.CheckAccess()) {
            # Wir sind bereits im UI-Thread
            $TextElement.Text = $sanitizedMessage
            if ($null -ne $Color) {
                $TextElement.Foreground = $Color
            }
        } 
        else {
            # Dispatcher verwenden für Thread-Sicherheit
            $TextElement.Dispatcher.Invoke([Action]{
                $TextElement.Text = $sanitizedMessage
                if ($null -ne $Color) {
                    $TextElement.Foreground = $Color
                }
            }, "Normal")
        }
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

# Funktion zum Aktualisieren des Status in der GUI
function Write-StatusMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Type = "Info"
    )
    
    try {
        # Logge die Nachricht auch
        Write-DebugMessage -Message $Message -Type $Type
        
        # Bestimme die Farbe basierend auf dem Nachrichtentyp
        $color = switch ($Type) {
            "Success" { $script:connectedBrush }
            "Error" { $script:disconnectedBrush }
            "Warning" { New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Orange) }
            "Info" { $null }
            default { $null }
        }
        
        # Aktualisiere das Status-Textfeld in der GUI
        if ($null -ne $script:txtStatus) {
            Update-GuiText -TextElement $script:txtStatus -Message $Message -Color $color
        }
    } 
    catch {
        # Bei Fehler einfach eine Debug-Meldung ausgeben
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler in Write-StatusMessage: $errorMsg" -Type "Error"
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
# Abschnitt: Eingabevalidierung
# -------------------------------------------------
function  Validate-Email{
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
            
            # Verbindung immer überprüfen
            try {
                # Einfache Prüfung durch Ausführen eines kleinen Exchange-Befehls
                Get-OrganizationConfig -ErrorAction Stop | Out-Null
                Write-DebugMessage "Bestehende Exchange-Verbindung ist gültig" -Type "Info"
            }
            catch {
                Write-DebugMessage "Bestehende Verbindung ist nicht mehr gültig, stelle neue Verbindung her" -Type "Warning"
                $Global:IsConnectedToExo = $false
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

function Test-ExchangeOnlineConnection {
    [CmdletBinding()]
    param()
    
    try {
        # Zuerst prüfen, ob die globale Testfunktion verfügbar ist
        if ($null -ne ${global:EasyStartup_TestExchangeOnlineConnection}) {
            return & ${global:EasyStartup_TestExchangeOnlineConnection}
        }
        
        # Eigene Verbindungsprüfung als Fallback
        try {
            # Einfacher Test durch Ausführung eines Exchange-Befehls
            Get-OrganizationConfig -ErrorAction Stop | Out-Null
            Write-DebugMessage "Exchange Online-Verbindung ist aktiv" -Type "Info"
            return $true
        }
        catch {
            Write-DebugMessage "Exchange Online-Verbindung ist nicht mehr gültig: $($_.Exception.Message)" -Type "Warning"
            return $false
        }
    }
    catch {
        Write-DebugMessage "Fehler beim Testen der Exchange Online-Verbindung: $($_.Exception.Message)" -Type "Error"
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
            # Frage den Benutzer ob er das wirklich tun möchte
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
            
            $successCount = 0
            $errorCount = 0

            if ($PermissionType -eq "Standard" -or $PermissionType -eq "Beides") {
                $result = Set-DefaultCalendarPermissionForAll -AccessRights $AccessRights
                if ($result) { $successCount++ } else { $errorCount++ }
            }
            if ($PermissionType -eq "Anonym" -or $PermissionType -eq "Beides") {
                $result = Set-AnonymousCalendarPermissionForAll -AccessRights $AccessRights
                if ($result) { $successCount++ } else { $errorCount++ }
            }
        }
        else {         
            if ([string]::IsNullOrWhiteSpace($MailboxUser) -and 
                $null -ne $script:txtCalendarMailboxUser -and 
                -not [string]::IsNullOrWhiteSpace($script:txtCalendarMailboxUser.Text)) {
                $mailboxUser = $script:txtCalendarMailboxUser.Text.Trim()
            }
            
            if ([string]::IsNullOrWhiteSpace($mailboxUser)) {
                throw "Keine Postfach-E-Mail-Adresse angegeben"
            }
            
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
        
        # Prüfe ob Berechtigung bereits existiert und ermittle den korrekten Kalenderordner
        $calendarExists = $false
        $identityDE = "${SourceUser}:\Kalender"
        $identityEN = "${SourceUser}:\Calendar"
        $identity = $null
        
        # Systematisch nach dem richtigen Kalender suchen
        try {
            # Zuerst versuchen wir den deutschen Kalender
            $existingPermDE = Get-MailboxFolderPermission -Identity $identityDE -User $TargetUser -ErrorAction SilentlyContinue
            if ($null -ne $existingPermDE) {
                $calendarExists = $true
                $identity = $identityDE
                Write-DebugMessage "Bestehende Berechtigung gefunden (DE): $($existingPermDE.AccessRights)" -Type "Info"
            }
            else {
                # Dann den englischen Kalender probieren
                $existingPermEN = Get-MailboxFolderPermission -Identity $identityEN -User $TargetUser -ErrorAction SilentlyContinue
                if ($null -ne $existingPermEN) {
                    $calendarExists = $true
                    $identity = $identityEN
                    Write-DebugMessage "Bestehende Berechtigung gefunden (EN): $($existingPermEN.AccessRights)" -Type "Info"
                }
            }
    }
    catch {
            Write-DebugMessage "Fehler bei der Prüfung bestehender Berechtigungen: $($_.Exception.Message)" -Type "Warning"
        }
        
        # Falls noch kein identifizierter Kalender, versuchen wir die Kalender zu prüfen ohne Benutzerberechtigungen
        if ($null -eq $identity) {
            try {
                # Prüfen, ob der deutsche Kalender existiert
                $deExists = Get-MailboxFolderPermission -Identity $identityDE -ErrorAction SilentlyContinue
                if ($null -ne $deExists) {
                    $identity = $identityDE
                    Write-DebugMessage "Deutscher Kalenderordner gefunden: $identityDE" -Type "Info"
                }
                else {
                    # Prüfen, ob der englische Kalender existiert
                    $enExists = Get-MailboxFolderPermission -Identity $identityEN -ErrorAction SilentlyContinue
                    if ($null -ne $enExists) {
                        $identity = $identityEN
                        Write-DebugMessage "Englischer Kalenderordner gefunden: $identityEN" -Type "Info"
                    }
                }
            }
            catch {
                Write-DebugMessage "Fehler beim Prüfen der Kalenderordner: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Falls immer noch kein Kalender gefunden, über Statistiken suchen
        if ($null -eq $identity) {
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
                Write-DebugMessage "Fehler beim Suchen des Kalenderordners über FolderStatistics: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Wenn immer noch kein Kalender gefunden, Exception werfen
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
        $success = $false
        switch ($GroupType) {
            "Verteilergruppe" {
                New-DistributionGroup @params -Type "Distribution" -ErrorAction Stop
                $success = $true
            }
            "Sicherheitsgruppe" {
                New-DistributionGroup @params -Type "Security" -ErrorAction Stop
                $success = $true
            }
            "Microsoft 365-Gruppe" {
                # Für Microsoft 365-Gruppen ggf. einen anderen Cmdlet verwenden
                New-UnifiedGroup @params -ErrorAction Stop
                $success = $true
            }
            "E-Mail-aktivierte Sicherheitsgruppe" {
                New-DistributionGroup @params -Type "Security" -ErrorAction Stop
                $success = $true
            }
            default {
                throw "Unbekannter Gruppentyp: $GroupType"
            }
        }
        
        # Wenn die Gruppe erstellt wurde und Members angegeben wurden, diese hinzufügen
        if ($success -and -not [string]::IsNullOrWhiteSpace($Members)) {
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

# Neue Funktion zum Aktualisieren der Domain-Liste
function Update-DomainList {
    [CmdletBinding()]
    param()
    
    try {
        Write-DebugMessage "Aktualisiere Domain-Liste für die ComboBox" -Type "Info"
        
        # Prüfen, ob die ComboBox existiert
        if ($null -eq $script:cmbSharedMailboxDomain) {
            $script:cmbSharedMailboxDomain = Get-XamlElement -ElementName "cmbSharedMailboxDomain"
            if ($null -eq $script:cmbSharedMailboxDomain) {
                Write-DebugMessage "Domain-ComboBox nicht gefunden" -Type "Warning"
                return $false
            }
        }
        
        # Prüfen, ob eine Verbindung besteht
        if (-not $script:isConnected) {
            Write-DebugMessage "Keine Exchange-Verbindung für Domain-Abfrage" -Type "Warning"
            return $false
        }
        
        # Domains abrufen und ComboBox befüllen
        $domains = Get-AcceptedDomain | Select-Object -ExpandProperty DomainName
        
        # Dispatcher verwenden für Thread-Sicherheit
        $script:cmbSharedMailboxDomain.Dispatcher.Invoke([Action]{
            $script:cmbSharedMailboxDomain.Items.Clear()
            foreach ($domain in $domains) {
                [void]$script:cmbSharedMailboxDomain.Items.Add($domain)
            }
            if ($script:cmbSharedMailboxDomain.Items.Count -gt 0) {
                $script:cmbSharedMailboxDomain.SelectedIndex = 0
            }
        }, "Normal")
        
        Write-DebugMessage "Domain-Liste erfolgreich aktualisiert: $($domains.Count) Domains geladen" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Aktualisieren der Domain-Liste: $errorMsg" -Type "Error"
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
    function Initialize-ContactsTab {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Kontakte-Tab" -Type "Info"
            
            # UI-Elemente referenzieren
            $txtContactExternalEmail = Get-XamlElement -ElementName "txtContactExternalEmail" -Required
            $txtContactSearch = Get-XamlElement -ElementName "txtContactSearch" -Required
            $cmbContactType = Get-XamlElement -ElementName "cmbContactType" -Required
            $btnCreateContact = Get-XamlElement -ElementName "btnCreateContact" -Required
            $btnShowMailContacts = Get-XamlElement -ElementName "btnShowMailContacts" -Required
            $btnShowMailUsers = Get-XamlElement -ElementName "btnShowMailUsers" -Required
            $btnRemoveContact = Get-XamlElement -ElementName "btnRemoveContact" -Required
            $lstContacts = Get-XamlElement -ElementName "lstContacts" -Required
            $btnExportContacts = Get-XamlElement -ElementName "btnExportContacts" -Required
            
            # Globale Variablen setzen
            $script:txtContactExternalEmail = $txtContactExternalEmail
            $script:txtContactSearch = $txtContactSearch
            $script:cmbContactType = $cmbContactType
            $script:lstContacts = $lstContacts
            
            # Event-Handler für Buttons registrieren
            Write-DebugMessage "Event-Handler für btnCreateContact registrieren" -Type "Info"
            Register-EventHandler -Control $btnCreateContact -Handler {
                    $externalEmail = $script:txtContactExternalEmail.Text.Trim()
                    if ([string]::IsNullOrEmpty($externalEmail)) {
                        Show-MessageBox -Message "Bitte geben Sie eine externe E-Mail-Adresse ein." -Title "Fehlende Angaben" -Type "Warning"
                        return
                    }
                    
                    $contactType = $script:cmbContactType.SelectedIndex
                    if ($contactType -eq 0) { # MailContact
                        try {
                            $displayName = $externalEmail.Split('@')[0]
                            $script:txtStatus.Text = "Erstelle MailContact $displayName..."
                            New-MailContact -Name $displayName -ExternalEmailAddress $externalEmail -FirstName $displayName -DisplayName $displayName -ErrorAction Stop
                            $script:txtStatus.Text = "MailContact erfolgreich erstellt: $displayName"
                            Show-MessageBox -Message "MailContact $displayName wurde erfolgreich erstellt." -Title "Erfolg" -Type "Info"
                            $script:txtContactExternalEmail.Text = ""
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Erstellen des MailContact: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Erstellen des MailContact: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                    }
                    else { # MailUser
                        try {
                            $displayName = $externalEmail.Split('@')[0]
                            $script:txtStatus.Text = "Erstelle MailUser $displayName..."
                            New-MailUser -Name $displayName -ExternalEmailAddress $externalEmail -MicrosoftOnlineServicesID "$displayName@$($script:primaryDomain)" -FirstName $displayName -DisplayName $displayName -ErrorAction Stop
                            $script:txtStatus.Text = "MailUser erfolgreich erstellt: $displayName"
                            Show-MessageBox -Message "MailUser $displayName wurde erfolgreich erstellt." -Title "Erfolg" -Type "Info"
                            $script:txtContactExternalEmail.Text = ""
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Erstellen des MailUser: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Erstellen des MailUser: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                }
            } -ControlName "btnCreateContact"
            Write-DebugMessage "Event-Handler für btnShowMailContacts registrieren" -Type "Info"
            Register-EventHandler -Control $btnShowMailContacts -Handler {
                # Verbindungsprüfung - nur bei Bedarf
                if (-not $script:isConnected) {
                    [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                
                $script:txtStatus.Text = "Lade MailContacts..."
                try {
                    # Versuche einen einfachen Exchange-Befehl zur Verbindungsprüfung
                    Get-OrganizationConfig -ErrorAction Stop | Out-Null
                    
                    # Exchange-Befehle ausführen
                    $contacts = Get-Recipient -RecipientType MailContact -ResultSize Unlimited | 
                        Select-Object DisplayName, PrimarySmtpAddress, ExternalEmailAddress, RecipientTypeDetails
                    
                    # Als Array behandeln
                    if ($null -ne $contacts -and -not ($contacts -is [Array])) {
                        $contacts = @($contacts)
                    }
                    
                    $script:lstContacts.ItemsSource = $contacts
                    $script:txtStatus.Text = "$($contacts.Count) MailContacts gefunden"
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    
                    # Wenn der Fehler auf eine abgelaufene Verbindung hindeutet
                    if ($errorMsg -like "*Connect-ExchangeOnline*" -or 
                        $errorMsg -like "*session*" -or 
                        $errorMsg -like "*authentication*") {
                        
                        # Status korrigieren
                        $script:isConnected = $false
                        $script:txtConnectionStatus.Text = "Nicht verbunden"
                        $script:txtConnectionStatus.Foreground = $script:disconnectedBrush
                        
                        [System.Windows.MessageBox]::Show("Die Verbindung zu Exchange Online ist unterbrochen. Bitte stellen Sie die Verbindung erneut her.", 
                            "Verbindungsfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                    else {
                        $script:txtStatus.Text = "Fehler beim Laden der MailContacts: $errorMsg"
                        [System.Windows.MessageBox]::Show("Fehler beim Laden der MailContacts: $errorMsg", 
                            "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    }
                }
            } -ControlName "btnShowMailContacts"
            
            Write-DebugMessage "Event-Handler für btnShowMailUsers registrieren" -Type "Info"

            Register-EventHandler -Control $btnShowMailUsers -Handler {
                # Überprüfe, ob wir mit Exchange verbunden sind
                if (-not (Confirm-ExchangeConnection)) {
                    Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                    return
                }
                
                $script:txtStatus.Text = "Lade MailUsers..."
                try {
                    $users = Get-MailUser -ResultSize Unlimited | 
                        Select-Object DisplayName, PrimarySmtpAddress, ExternalEmailAddress, RecipientTypeDetails
                    
                    $script:lstContacts.ItemsSource = $users
                    $script:txtStatus.Text = "$($users.Count) MailUsers gefunden"
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $script:txtStatus.Text = "Fehler beim Laden der MailUsers: $errorMsg"
                    Show-MessageBox -Message "Fehler beim Laden der MailUsers: $errorMsg" -Title "Fehler" -Type "Error"
                }
            } -ControlName "btnShowMailUsers"
            
            Write-DebugMessage "Event-Handler für btnRemoveContact registrieren" -Type "Info"
            Register-EventHandler -Control $btnRemoveContact -Handler {
                    if ($script:lstContacts.SelectedItem -eq $null) {
                        Show-MessageBox -Message "Bitte wählen Sie zuerst einen Kontakt aus." -Title "Kein Kontakt ausgewählt" -Type "Warning"
                        return
                    }
                    
                    $selectedContact = $script:lstContacts.SelectedItem
                    $contactType = $selectedContact.RecipientTypeDetails
                    $contactName = $selectedContact.DisplayName
                    $contactEmail = $selectedContact.PrimarySmtpAddress
                    
                    $confirmResult = Show-MessageBox -Message "Sind Sie sicher, dass Sie den Kontakt '$contactName' löschen möchten?" `
                        -Title "Kontakt löschen" -Type "YesNo"
                    
                    if ($confirmResult -eq "Yes") {
                        try {
                            $script:txtStatus.Text = "Lösche Kontakt $contactName..."
                            
                            if ($contactType -like "*MailContact*") {
                                Remove-MailContact -Identity $contactEmail -Confirm:$false -ErrorAction Stop
                            }
                            elseif ($contactType -like "*MailUser*") {
                                Remove-MailUser -Identity $contactEmail -Confirm:$false -ErrorAction Stop
                            }
                            
                            $script:txtStatus.Text = "Kontakt $contactName erfolgreich gelöscht"
                            Show-MessageBox -Message "Der Kontakt '$contactName' wurde erfolgreich gelöscht." -Title "Kontakt gelöscht" -Type "Info"
                            
                            # Liste aktualisieren
                            if ($contactType -like "*MailContact*") {
                                $script:btnShowMailContacts.RaiseEvent([System.Windows.RoutedEventArgs]::Click)
                            }
                            else {
                                $script:btnShowMailUsers.RaiseEvent([System.Windows.RoutedEventArgs]::Click)
                            }
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Löschen des Kontakts: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Löschen des Kontakts: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                }
            } -ControlName "btnRemoveContact"
            
            Write-DebugMessage "Event-Handler für btnExportContacts registrieren" -Type "Info"
            Register-EventHandler -Control $btnExportContacts -Handler {
                    if ($script:lstContacts.Items.Count -eq 0) {
                        Show-MessageBox -Message "Es gibt keine Kontakte zum Exportieren." -Title "Keine Daten" -Type "Info"
                        return
                    }
                    
                    $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
                    $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv"
                    $saveFileDialog.Title = "Kontakte exportieren"
                    $saveFileDialog.FileName = "Kontakte_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                    
                    if ($saveFileDialog.ShowDialog() -eq $true) {
                        try {
                            $script:txtStatus.Text = "Exportiere Kontakte nach $($saveFileDialog.FileName)..."
                            $script:lstContacts.ItemsSource | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                            $script:txtStatus.Text = "Kontakte erfolgreich exportiert: $($saveFileDialog.FileName)"
                            Show-MessageBox -Message "Die Kontakte wurden erfolgreich nach '$($saveFileDialog.FileName)' exportiert." -Title "Export erfolgreich" -Type "Info"
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $script:txtStatus.Text = "Fehler beim Exportieren der Kontakte: $errorMsg"
                            Show-MessageBox -Message "Fehler beim Exportieren der Kontakte: $errorMsg" -Title "Fehler" -Type "Error"
                        }
                    }
            } -ControlName "btnExportContacts"
            
            Write-DebugMessage "Kontakte-Tab wurde initialisiert" -Type "Success"
            return $true
                }
                catch {
                    $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler bei der Initialisierung des Kontakte-Tabs: $errorMsg" -Type "Error"
            return $false
        }
    }
    
    # Implementiert die Hauptfunktionen für den Kontakte-Tab
    
    function New-ExoContact {
        [CmdletBinding()]
        param()
        
        try {
            $name = $script:txtContactName.Text.Trim()
            $email = $script:txtContactEmail.Text.Trim()
            
            if ([string]::IsNullOrEmpty($name) -or [string]::IsNullOrEmpty($email)) {
                Show-MessageBox -Message "Bitte geben Sie einen Namen und eine E-Mail-Adresse ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Erstelle Kontakt $name mit E-Mail $email..."
            
            # Exchange-Befehle ausführen
            $newMailContact = New-MailContact -Name $name -ExternalEmailAddress $email -ErrorAction Stop
            
            if ($newMailContact) {
                Write-StatusMessage "Kontakt $name erfolgreich erstellt."
                Show-MessageBox -Message "Der Kontakt $name wurde erfolgreich erstellt." -Title "Kontakt erstellt" -Type "Info"
                
                # Kontakte aktualisieren
                Get-ExoContacts
                
                # Felder leeren
                $script:txtContactName.Text = ""
                $script:txtContactEmail.Text = ""
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Erstellen des Kontakts: $errorMsg"
            Show-MessageBox -Message "Fehler beim Erstellen des Kontakts: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
    function Update-ExoContact {
        [CmdletBinding()]
        param()
        
        try {
            $name = $script:txtContactName.Text.Trim()
            $email = $script:txtContactEmail.Text.Trim()
            
            if ([string]::IsNullOrEmpty($name) -or [string]::IsNullOrEmpty($email)) {
                Show-MessageBox -Message "Bitte geben Sie einen Namen und eine E-Mail-Adresse ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            if ($script:dgContacts.SelectedItem) {
                $contactObj = $script:dgContacts.SelectedItem
                $contactId = $contactObj.Identity
                
                Write-StatusMessage "Aktualisiere Kontakt $contactId..."
                
                # Exchange-Befehle ausführen
                Set-MailContact -Identity $contactId -Name $name -ExternalEmailAddress $email -ErrorAction Stop
                
                Write-StatusMessage "Kontakt $name erfolgreich aktualisiert."
                Show-MessageBox -Message "Der Kontakt $name wurde erfolgreich aktualisiert." -Title "Kontakt aktualisiert" -Type "Info"
                
                # Kontakte aktualisieren
                Get-ExoContacts
                
                # Felder leeren
                $script:txtContactName.Text = ""
                $script:txtContactEmail.Text = ""
            }
            else {
                Show-MessageBox -Message "Bitte wählen Sie zuerst einen Kontakt aus der Liste aus." -Title "Kein Kontakt ausgewählt" -Type "Warning"
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Aktualisieren des Kontakts: $errorMsg"
            Show-MessageBox -Message "Fehler beim Aktualisieren des Kontakts: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
    function Search-ExoContacts {
        [CmdletBinding()]
        param()
        
        try {
            $searchTerm = $txtContactSearch.Text.Trim()
            
            if ([string]::IsNullOrEmpty($searchTerm)) {
                Show-MessageBox -Message "Bitte geben Sie einen Suchbegriff ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Suche nach Kontakten mit '$searchTerm'..."
            
            # Exchange-Befehle ausführen
            $contacts = Get-MailContact -Filter "Name -like '*$searchTerm*' -or EmailAddresses -like '*$searchTerm*'" -ResultSize Unlimited | 
                Select-Object Name, DisplayName, PrimarySmtpAddress, ExternalEmailAddress, Department, Title
            
            # Daten zum DataGrid hinzufügen
            $script:dgContacts.Dispatcher.Invoke([action]{
                $script:dgContacts.ItemsSource = $contacts
            }, "Normal")
            
            Write-StatusMessage "Suche abgeschlossen. $($contacts.Count) Kontakte gefunden."
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler bei der Kontaktsuche: $errorMsg"
            Show-MessageBox -Message "Fehler bei der Kontaktsuche: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
    function Get-ExoContacts {
        [CmdletBinding()]
        param()
        
        try {
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Lade Kontakte..."
            
            # Exchange-Befehle ausführen
            $contacts = Get-MailContact -ResultSize Unlimited | 
                Select-Object Name, DisplayName, PrimarySmtpAddress, ExternalEmailAddress, Department, Title
            
            # Daten zum DataGrid hinzufügen
            $script:dgContacts.Dispatcher.Invoke([action]{
                $script:dgContacts.ItemsSource = $contacts
            }, "Normal")
            
            Write-StatusMessage "Kontakte geladen. $($contacts.Count) Kontakte gefunden."
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Laden der Kontakte: $errorMsg"
        }
    }

    # Funktion zum Initialisieren des Ressourcen-Tabs
    function Initialize-ResourcesTab {
        [CmdletBinding()]
        param()

        try {
            Write-DebugMessage "Initialisiere Ressourcen-Tab" -Type "Info"

            # Steuerelemente finden
            $helpLinkResources = $script:Form.FindName("helpLinkResources") # Beibehalten für Hilfe-Link
            $btnCreateResource = $script:Form.FindName("btnCreateResource")
            $btnShowRoomResources = $script:Form.FindName("btnShowRoomResources")
            $btnShowEquipmentResources = $script:Form.FindName("btnShowEquipmentResources")
            $btnSearchResources = $script:Form.FindName("btnSearchResources")
            $btnRefreshResources = $script:Form.FindName("btnRefreshResources")
            $btnEditResourceSettings = $script:Form.FindName("btnEditResourceSettings")
            $btnRemoveResource = $script:Form.FindName("btnRemoveResource")
            $btnExportResources = $script:Form.FindName("btnExportResources")
            $cmbResourceType = $script:Form.FindName("cmbResourceType")
            $txtResourceName = $script:Form.FindName("txtResourceName")
            $txtResourceSearch = $script:Form.FindName("txtResourceSearch")
            $dgResources = $script:Form.FindName("dgResources")
            $cmbResourceSelect = $script:Form.FindName("cmbResourceSelect")

            # Steuerelemente im Script-Scope speichern
            $script:cmbResourceType = $cmbResourceType
            $script:txtResourceName = $txtResourceName
            $script:txtResourceSearch = $txtResourceSearch
            $script:dgResources = $dgResources
            $script:cmbResourceSelect = $cmbResourceSelect

            # ComboBox für Ressourcentyp füllen
            if ($null -ne $script:cmbResourceType) {
                if ($script:cmbResourceType.Items.Count -eq 0) {
                    $script:cmbResourceType.ItemsSource = @("Room", "Equipment")
                    $script:cmbResourceType.SelectedIndex = 0
                    Write-DebugMessage "cmbResourceType mit Optionen gefüllt." -Type "Info"
                } else {
                    Write-DebugMessage "cmbResourceType enthält bereits Items. ItemsSource wird nicht gesetzt." -Type "Warning"
                }
            }

            # DataGrid und ComboBox leeren, bevor Actions ItemsSource setzen
            if ($null -ne $script:dgResources -and $script:dgResources.Items.Count -gt 0) {
                Write-DebugMessage "dgResources enthält bereits Items. Wird vor dem Binden durch Aktionen geleert." -Type "Info"
            }
            if ($null -ne $script:cmbResourceSelect -and $script:cmbResourceSelect.Items.Count -gt 0) {
                Write-DebugMessage "cmbResourceSelect enthält bereits Items. Wird vor dem Binden durch Aktionen geleert." -Type "Info"
            }

            # Event-Handler registrieren
            Register-EventHandler -Control $btnCreateResource -Handler { Create-ResourceAction -ResourceName $script:txtResourceName.Text -ResourceType $script:cmbResourceType.SelectedItem } -ControlName "btnCreateResource"
            Register-EventHandler -Control $btnShowRoomResources -Handler { Show-RoomResourcesAction } -ControlName "btnShowRoomResources"
            Register-EventHandler -Control $btnShowEquipmentResources -Handler { Show-EquipmentResourcesAction } -ControlName "btnShowEquipmentResources"
            Register-EventHandler -Control $btnSearchResources -Handler { Search-ResourcesAction -SearchTerm $script:txtResourceSearch.Text } -ControlName "btnSearchResources"
            Register-EventHandler -Control $btnRefreshResources -Handler { Refresh-ResourcesAction } -ControlName "btnRefreshResources"
            Register-EventHandler -Control $btnEditResourceSettings -Handler {
                $selectedItem = $script:dgResources.SelectedItem
                if ($null -ne $selectedItem) { Edit-ResourceSettingsAction -SelectedResource $selectedItem }
                else { Show-MessageBox -Message "Bitte wählen Sie zuerst eine Ressource aus." -Title "Keine Auswahl" -Type Info }
            } -ControlName "btnEditResourceSettings"
            Register-EventHandler -Control $btnRemoveResource -Handler {
                $selectedItem = $script:dgResources.SelectedItem
                if ($null -ne $selectedItem) { Remove-ResourceAction -SelectedResource $selectedItem }
                else { Show-MessageBox -Message "Bitte wählen Sie zuerst eine Ressource zum Entfernen aus." -Title "Keine Auswahl" -Type Info }
            } -ControlName "btnRemoveResource"
            Register-EventHandler -Control $btnExportResources -Handler {
                $itemsToExport = $script:dgResources.ItemsSource -as [System.Collections.IList]
                if ($null -ne $itemsToExport) { Export-ResourcesAction -ResourceList $itemsToExport }
                else { Show-MessageBox -Message "Keine Daten zum Exportieren vorhanden." -Title "Export" -Type Info }
            } -ControlName "btnExportResources"

            # Event-Handler für Hilfe-Link (korrigierte Methode)
            if ($null -ne $helpLinkResources) {
                Write-DebugMessage "Registriere Event-Handler für helpLinkResources" -Type Verbose
                $helpLinkResources.Add_MouseLeftButtonDown({
                    # Sicherstellen, dass Show-HelpDialog existiert und korrekt aufgerufen wird
                    if (Get-Command Show-HelpDialog -ErrorAction SilentlyContinue) {
                        Show-HelpDialog -Topic "Resources"
                    } else {
                        Write-DebugMessage "Funktion Show-HelpDialog nicht gefunden." -Type Error
                    }
                })
                $helpLinkResources.Add_MouseEnter({
                    $this.Cursor = [System.Windows.Input.Cursors]::Hand
                    $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                })
                $helpLinkResources.Add_MouseLeave({
                    $this.TextDecorations = $null
                    $this.Cursor = [System.Windows.Input.Cursors]::Arrow
                })
                Write-DebugMessage "Event-Handler für helpLinkResources registriert." -Type Info
            } else {
                Write-DebugMessage "Hilfe-Link 'helpLinkResources' nicht gefunden." -Type Warning
            }


            Write-DebugMessage "Ressourcen-Tab erfolgreich initialisiert" -Type "Success"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message + "`n" + $_.ScriptStackTrace # Füge StackTrace hinzu für mehr Details
            Write-DebugMessage "Fehler beim Initialisieren des Ressourcen-Tabs: $errorMsg" -Type "Error"
            Show-MessageBox -Message "Schwerwiegender Fehler beim Initialisieren des Ressourcen-Tabs: $errorMsg" -Title "Initialisierungsfehler" -Type Error
            return $false
        }
    }    
    # Implementiert die Hauptfunktionen für den Ressourcen-Tab
    
    function New-ExoResource {
        [CmdletBinding()]
        param()
        
        try {
            $name = $script:txtResourceName.Text.Trim()
            $resourceTypeIndex = $script:cmbResourceType.SelectedIndex
            
            if ([string]::IsNullOrEmpty($name)) {
                Show-MessageBox -Message "Bitte geben Sie einen Namen für die Ressource ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            # Ressourcentyp bestimmen
            $resourceType = if ($resourceTypeIndex -eq 0) { "Room" } else { "Equipment" }
            
            Write-StatusMessage "Erstelle Ressource $name vom Typ $resourceType..."
            
            # Exchange-Befehle ausführen
            if ($resourceType -eq "Room") {
                $newResource = New-Mailbox -Name $name -Room -ErrorAction Stop
            } else {
                $newResource = New-Mailbox -Name $name -Equipment -ErrorAction Stop
            }
            
            if ($newResource) {
                # Warte kurz, damit die Änderungen sich in Exchange verbreiten können
                Start-Sleep -Seconds 2
                
                # Standardwerte für Ressourcen festlegen
                if ($resourceType -eq "Room") {
                    Set-CalendarProcessing -Identity $name -AutomateProcessing AutoAccept -AllowConflicts $false -BookingWindowInDays 180 -ErrorAction SilentlyContinue
                } else {
                    Set-CalendarProcessing -Identity $name -AutomateProcessing AutoAccept -AllowConflicts $false -BookingWindowInDays 180 -ErrorAction SilentlyContinue
                }
                
                Write-StatusMessage "Ressource $name erfolgreich erstellt."
                Show-MessageBox -Message "Die Ressource $name wurde erfolgreich erstellt." -Title "Ressource erstellt" -Type "Info"
                
                # Ressourcen aktualisieren
                Get-ExoResources
                
                # Felder leeren
                $script:txtResourceName.Text = ""
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Erstellen der Ressource: $errorMsg"
            Show-MessageBox -Message "Fehler beim Erstellen der Ressource: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
    function Set-ExoResourceBookingOptions {
        [CmdletBinding()]
        param()
        
        try {
            # Überprüfen, ob eine Ressource ausgewählt ist
            if (-not $script:dgResources.SelectedItem) {
                Show-MessageBox -Message "Bitte wählen Sie zuerst eine Ressource aus der Liste aus." -Title "Keine Ressource ausgewählt" -Type "Warning"
                return
            }
            
            $resourceObj = $script:dgResources.SelectedItem
            $resourceEmail = $resourceObj.EmailAddress
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            $scheduleOnlyDuringWorkHours = $script:chkScheduleOnlyDuringWorkHours.IsChecked
            $maxDurationStr = $script:txtMaxDuration.Text.Trim()
            
            # Prüfen, ob MaximumDuration ein gültiger Integer ist
            $maxDuration = 0
            if (-not [string]::IsNullOrEmpty($maxDurationStr) -and [int]::TryParse($maxDurationStr, [ref]$maxDuration)) {
                # Gültiger Wert, weiter
            } else {
                Show-MessageBox -Message "Die maximale Buchungsdauer muss eine gültige Zahl sein." -Title "Ungültige Eingabe" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Aktualisiere Buchungsoptionen für Ressource $($resourceObj.Name)..."
            
            # Exchange-Befehle ausführen
            $params = @{
                Identity = $resourceEmail
                ScheduleOnlyDuringWorkHours = $scheduleOnlyDuringWorkHours
                ErrorAction = "Stop"
            }
            
            if ($maxDuration -gt 0) {
                $params.Add("MaximumDuration", (New-TimeSpan -Minutes $maxDuration))
            }
            
            Set-CalendarProcessing @params
            
            Write-StatusMessage "Buchungsoptionen für $($resourceObj.Name) erfolgreich aktualisiert."
            Show-MessageBox -Message "Die Buchungsoptionen für $($resourceObj.Name) wurden erfolgreich aktualisiert." -Title "Optionen aktualisiert" -Type "Info"
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Aktualisieren der Buchungsoptionen: $errorMsg"
            Show-MessageBox -Message "Fehler beim Aktualisieren der Buchungsoptionen: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
    function Get-ResourceBookingOptions {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [string]$ResourceEmail
        )
        
        try {
            Write-StatusMessage "Lade Buchungsoptionen für $ResourceEmail..."
            
            # Exchange-Befehle ausführen
            $calendarProcessing = Get-CalendarProcessing -Identity $ResourceEmail -ErrorAction Stop
            
            # Setze UI-Elemente
            $script:chkScheduleOnlyDuringWorkHours.IsChecked = $calendarProcessing.ScheduleOnlyDuringWorkHours
            
            # Setze MaximumDuration in Minuten
            if ($calendarProcessing.MaximumDuration) {
                $script:txtMaxDuration.Text = [math]::Round($calendarProcessing.MaximumDuration.TotalMinutes)
            } else {
                $script:txtMaxDuration.Text = ""
            }
            
            Write-StatusMessage "Buchungsoptionen geladen."
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Laden der Buchungsoptionen: $errorMsg"
        }
    }
    
    function Search-ExoResources {
        [CmdletBinding()]
        param()
        
        try {
            $searchTerm = $txtResourceSearch.Text.Trim()
            
            if ([string]::IsNullOrEmpty($searchTerm)) {
                Show-MessageBox -Message "Bitte geben Sie einen Suchbegriff ein." -Title "Fehlende Angaben" -Type "Warning"
                return
            }
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Suche nach Ressourcen mit '$searchTerm'..."
            
            # Exchange-Befehle ausführen
            $rooms = Get-Mailbox -Filter "Name -like '*$searchTerm*' -and RecipientTypeDetails -eq 'RoomMailbox'" -ResultSize Unlimited
            $equipment = Get-Mailbox -Filter "Name -like '*$searchTerm*' -and RecipientTypeDetails -eq 'EquipmentMailbox'" -ResultSize Unlimited
            
            # Konvertiere die Ergebnisse in ein einheitliches Format
            $results = @()
            
            foreach ($room in $rooms) {
                $results += [PSCustomObject]@{
                    Name = $room.DisplayName
                    Type = "Raum"
                    EmailAddress = $room.PrimarySmtpAddress
                    Capacity = (Get-Mailbox -Identity $room.Identity | Select-Object -ExpandProperty ResourceCapacity)
                    Identity = $room.Identity
                }
            }
            
            foreach ($equip in $equipment) {
                $results += [PSCustomObject]@{
                    Name = $equip.DisplayName
                    Type = "Ausstattung"
                    EmailAddress = $equip.PrimarySmtpAddress
                    Capacity = $null
                    Identity = $equip.Identity
                }
            }
            
            # Daten zum DataGrid hinzufügen
            $script:dgResources.Dispatcher.Invoke([action]{
                $script:dgResources.ItemsSource = $results
            }, "Normal")
            
            Write-StatusMessage "Suche abgeschlossen. $($results.Count) Ressourcen gefunden."
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler bei der Ressourcensuche: $errorMsg"
            Show-MessageBox -Message "Fehler bei der Ressourcensuche: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }
    
    function Get-ExoResources {
        [CmdletBinding()]
        param()
        
        try {
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Lade alle Ressourcen..."
            
            # Rufen Sie sowohl Raum- als auch Ausstattungsressourcen ab
            $rooms = Get-Mailbox -Filter "RecipientTypeDetails -eq 'RoomMailbox'" -ResultSize Unlimited
            $equipment = Get-Mailbox -Filter "RecipientTypeDetails -eq 'EquipmentMailbox'" -ResultSize Unlimited
            
            # Konvertiere die Ergebnisse in ein einheitliches Format
            $results = @()
            
            foreach ($room in $rooms) {
                $results += [PSCustomObject]@{
                    Name = $room.DisplayName
                    Type = "Raum"
                    EmailAddress = $room.PrimarySmtpAddress
                    Capacity = (Get-Mailbox -Identity $room.Identity | Select-Object -ExpandProperty ResourceCapacity)
                    Identity = $room.Identity
                }
            }
            
            foreach ($equip in $equipment) {
                $results += [PSCustomObject]@{
                    Name = $equip.DisplayName
                    Type = "Ausstattung"
                    EmailAddress = $equip.PrimarySmtpAddress
                    Capacity = $null
                    Identity = $equip.Identity
                }
            }
            
            # Daten zum DataGrid hinzufügen
            $script:dgResources.Dispatcher.Invoke([action]{
                $script:dgResources.ItemsSource = $results
            }, "Normal")
            
            Write-StatusMessage "Ressourcen geladen. $($results.Count) Ressourcen gefunden."
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Laden der Ressourcen: $errorMsg"
        }
    }
    
    function Get-ExoRoomResources {
        [CmdletBinding()]
        param()
        
        try {
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Lade Raumressourcen..."
            
            # Exchange-Befehle ausführen
            $rooms = Get-Mailbox -Filter "RecipientTypeDetails -eq 'RoomMailbox'" -ResultSize Unlimited
            
            # Konvertiere die Ergebnisse in ein einheitliches Format
            $results = @()
            
            foreach ($room in $rooms) {
                $results += [PSCustomObject]@{
                    Name = $room.DisplayName
                    Type = "Raum"
                    EmailAddress = $room.PrimarySmtpAddress
                    Capacity = (Get-Mailbox -Identity $room.Identity | Select-Object -ExpandProperty ResourceCapacity)
                    Identity = $room.Identity
                }
            }
            
            # Daten zum DataGrid hinzufügen
            $script:dgResources.Dispatcher.Invoke([action]{
                $script:dgResources.ItemsSource = $results
            }, "Normal")
            
            Write-StatusMessage "Raumressourcen geladen. $($results.Count) Ressourcen gefunden."
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Laden der Raumressourcen: $errorMsg"
        }
    }
    
    function Get-ExoEquipmentResources {
        [CmdletBinding()]
        param()
        
        try {
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            Write-StatusMessage "Lade Ausstattungsressourcen..."
            
            # Exchange-Befehle ausführen
            $equipment = Get-Mailbox -Filter "RecipientTypeDetails -eq 'EquipmentMailbox'" -ResultSize Unlimited
            
            # Konvertiere die Ergebnisse in ein einheitliches Format
            $results = @()
            
            foreach ($equip in $equipment) {
                $results += [PSCustomObject]@{
                    Name = $equip.DisplayName
                    Type = "Ausstattung"
                    EmailAddress = $equip.PrimarySmtpAddress
                    Capacity = $null
                    Identity = $equip.Identity
                }
            }
            
            # Daten zum DataGrid hinzufügen
            $script:dgResources.Dispatcher.Invoke([action]{
                $script:dgResources.ItemsSource = $results
            }, "Normal")
            
            Write-StatusMessage "Ausstattungsressourcen geladen. $($results.Count) Ressourcen gefunden."
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Laden der Ausstattungsressourcen: $errorMsg"
        }
    }
    
    function Remove-ExoResource {
        [CmdletBinding()]
        param()
        
        try {
            # Überprüfen, ob eine Ressource ausgewählt ist
            if (-not $script:dgResources.SelectedItem) {
                Show-MessageBox -Message "Bitte wählen Sie zuerst eine Ressource aus der Liste aus." -Title "Keine Ressource ausgewählt" -Type "Warning"
                return
            }
            
            $resourceObj = $script:dgResources.SelectedItem
            $resourceName = $resourceObj.Name
            $resourceEmail = $resourceObj.EmailAddress
            
            # Stellen Sie sicher, dass wir mit Exchange verbunden sind
            if (!(Confirm-ExchangeConnection)) {
                Show-MessageBox -Message "Bitte zuerst mit Exchange Online verbinden." -Title "Nicht verbunden" -Type "Warning"
                return
            }
            
            # Sicherheitsabfrage
            $confirm = Show-MessageBox -Message "Sind Sie sicher, dass Sie die Ressource '$resourceName' löschen möchten?`nDiese Aktion kann nicht rückgängig gemacht werden!" -Title "Ressource löschen" -Type "YesNo"
            
            if ($confirm -eq "Yes") {
                Write-StatusMessage "Lösche Ressource $resourceName..."
                
                # Exchange-Befehle ausführen
                Remove-Mailbox -Identity $resourceObj.Identity -Confirm:$false -ErrorAction Stop
                
                Write-StatusMessage "Ressource $resourceName erfolgreich gelöscht."
                Show-MessageBox -Message "Die Ressource $resourceName wurde erfolgreich gelöscht." -Title "Ressource gelöscht" -Type "Info"
                
                # Ressourcen aktualisieren
                Get-ExoResources
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-StatusMessage "Fehler beim Löschen der Ressource: $errorMsg"
            Show-MessageBox -Message "Fehler beim Löschen der Ressource: $errorMsg" -Title "Fehler" -Type "Error"
        }
    }

    # Initialisiere alle Tabs
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
        $settingsWindow.Width = 600
        $settingsWindow.Height = 450
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
        Write-DebugMessage "Initialisiere Hilfe-Links" -Type "Info"
        
        # Da möglicherweise kein HelpLinks-Container existiert,
        # suche nach individuellen Hilfe-Links
        $helpLinks = @(
            $script:Form.FindName("helpLinkCalendar"),
            $script:Form.FindName("helpLinkMailbox"),
            $script:Form.FindName("helpLinkAudit"),
            $script:Form.FindName("helpLinkTroubleshooting"),
            $script:Form.FindName("helpLinkGroups"),
            $script:Form.FindName("helpLinkShared")
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
                    "*Groups*" { $topic = "Groups" }
                    "*Shared*" { $topic = "SharedMailbox" }
                }
                
                # Event-Handler hinzufügen mit Fehlerbehandlung
                try {
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
                catch {
                    Write-DebugMessage "Fehler beim Hinzufügen von Event-Handlern für $linkName - $($_.Exception.Message)" -Type "Warning"
                }
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
    $script:tabEXOSettings      = Get-XamlElement -ElementName "tabEXOSettings"
    $script:tabCalendar         = Get-XamlElement -ElementName "tabCalendar"
    $script:tabMailbox          = Get-XamlElement -ElementName "tabMailbox"
    $script:tabResources        = Get-XamlElement -ElementName "tabResources"
    $script:tabContacts         = Get-XamlElement -ElementName "tabContacts"
    $script:tabMailboxAudit     = Get-XamlElement -ElementName "tabMailboxAudit"
    $script:tabTroubleshooting  = Get-XamlElement -ElementName "tabTroubleshooting"
    $script:txtStatus           = Get-XamlElement -ElementName "txtStatus" -Required
    $script:txtVersion          = Get-XamlElement -ElementName "txtVersion"
    $script:txtConnectionStatus = Get-XamlElement -ElementName "txtConnectionStatus" -Required
    
    # Referenzierung der Navigationselemente
    $script:btnNavEXOSettings     = Get-XamlElement -ElementName "btnNavEXOSettings"
    $script:btnNavCalendar        = Get-XamlElement -ElementName "btnNavCalendar"
    $script:btnNavMailbox         = Get-XamlElement -ElementName "btnNavMailbox"
    $script:btnNavGroups          = Get-XamlElement -ElementName "btnNavGroups"
    $script:btnNavSharedMailbox   = Get-XamlElement -ElementName "btnNavSharedMailbox"
    $script:btnNavResources       = Get-XamlElement -ElementName "btnNavResources"
    $script:btnNavContacts        = Get-XamlElement -ElementName "btnNavContacts"
    $script:btnNavAudit           = Get-XamlElement -ElementName "btnNavAudit"
    $script:btnNavReports         = Get-XamlElement -ElementName "btnNavReports"
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
# Abschnitt: Exchange Online Settings Module Funktionen
# -------------------------------------------------

# easyEXOSettings Modul integriert in die Hauptdatei
# Ursprünglich ein separates Modul: easyEXOSettings.ps1

#region EXOSettings Global Variables
# -----------------------------------------------
# EXOSettings Global Variables
# -----------------------------------------------
$script:tabEXOSettings = $null
$script:organizationConfigSettings = @{}
$script:currentOrganizationConfig = $null
$script:EXOSettingsLoggingEnabled = $false
$script:EXOSettingsLogFilePath = Join-Path -Path $PSScriptRoot -ChildPath "logs\easyEXOSettings.log"

# Definierte UI-Elemente aus der XAML, die wir erwarten
$script:knownUIElements = @(
    # Benutzereinstellungen Tab
    "chkActivityBasedAuthenticationTimeoutEnabled", 
    "chkActivityBasedAuthenticationTimeoutInterval", # Achtung: Ist eine ComboBox, nicht CheckBox
    "chkActivityBasedAuthenticationTimeoutWithSingleSignOnEnabled", 
    "chkAppsForOfficeEnabled", 
    "chkAsyncSendEnabled", 
    "chkBookingsAddressEntryRestricted", 
    "chkBookingsAuthEnabled", 
    "chkBookingsCreationOfCustomQuestionsRestricted", 
    "chkBookingsExposureOfStaffDetailsRestricted", 
    "chkBookingsMembershipApprovalRequired", 
    "chkBookingsNamingPolicyEnabled", 
    "chkBookingsNamingPolicySuffix", 
    "chkBookingsNamingPolicySuffixEnabled", 
    "chkBookingsNotesEntryRestricted", 
    "chkBookingsPaymentsEnabled", 
    "chkBookingsSocialSharingRestricted", 
    "chkFocusedInboxOn", 
    "chkReadTrackingEnabled", 
    "chkSendFromAliasEnabled",
    
    # Admin/Sicherheit Tab
    "chkAdditionalStorageProvidersBlocked", 
    "chkAuditDisabled", 
    "chkAutodiscoverPartialDirSync", 
    "chkAutoEnableArchiveMailbox", 
    "chkAutoExpandingArchive", 
    "chkCalendarVersionStoreEnabled", 
    "chkCASMailboxHasPermissionsIncludingSubfolders", 
    "chkComplianceEnabled", 
    "chkCustomerLockboxEnabled", 
    "chkEcRequiresTls", 
    "chkElcProcessingDisabled", 
    "chkEnableOutlookEvents", 
    "chkMailTipsExternalRecipientsTipsEnabled", 
    "chkMailTipsGroupMetricsEnabled", 
    "chkMailTipsLargeAudienceThreshold", 
    "cmbLargeAudienceThreshold",
    "chkMailTipsMailboxSourcedTipsEnabled", 
    "chkOwaRedirectToOD4BThisUserEnabled", 
    "chkPublicFolderShowClientControl", 
    "chkPublicComputersDetectionEnabled",
    
    # Authentifizierung Tab
    "chkOAuth2ClientProfileEnabled", 
    "chkInformationBarrierMode", 
    "cmbInformationBarrierMode", 
    "chkImplicitSharingEnabled", 
    "chkOAuthUseBasicAuth", 
    "chkRefreshSessionEnabled", 
    "chkPerTenantSwitchToESTSEnabled", 
    "chkEwsApplicationAccessPolicy", 
    "cmbEwsAppAccessPolicy", 
    "chkEws", 
    "chkEwsAllowList", 
    "chkEwsAllowEntourage", 
    "chkEwsAllowMacOutlook", 
    "chkEwsAllowOutlook", 
    "chkMacOutlook", 
    "chkOutlookMobile", 
    "chkWACDiscoveryEndpoint",
    
    # Mobile & Zugriff Tab
    "chkMobileAppEducationEnabled", 
    "chkEnableUserPowerShell", 
    "chkIsSingleInstance", 
    "chkOnPremisesDownloadDisabled", 
    "chkAcceptApiLicenseAgreement", 
    "chkConnectorsEnabled", 
    "chkConnectorsEnabledForYammer", 
    "chkConnectorsEnabledForTeams", 
    "chkConnectorsEnabledForSharepoint", 
    "chkOfficeFeatures", 
    "cmbOfficeFeatures", 
    "chkMobileToFollowedFolders", 
    "chkDisablePlusAddressInRecipients", 
    "chkDefaultAuthenticationPolicy", 
    "txtDefaultAuthPolicy", 
    "chkHierarchicalAddressBookRoot", 
    "txtHierAddressBookRoot",
    
    # Erweitert Tab
    "chkSIPEnabled", 
    "chkRemotePublicFolderBlobsEnabled", 
    "chkPreferredInternetCodePageForShiftJis", 
    "txtPreferredInternetCodePageForShiftJis", 
    "chkVisibilityEnabled", 
    "chkOnlineMeetingsByDefaultEnabled", 
    "chkSearchQueryLanguage", 
    "cmbSearchQueryLanguage", 
    "chkDirectReportsGroupAutoCreationEnabled", 
    "chkMapiHttpEnabled", 
    "chkUnblockUnsafeSenderPromptEnabled", 
    "chkExecutiveAttestation", 
    "chkPDPLocationEnabled", 
    "txtPowerShellMaxConcurrency", 
    "txtPowerShellMaxCmdletQueueDepth", 
    "txtPowerShellMaxCmdletsExecutionDuration",
    
    # Result Text area
    "txtOrganizationConfig"
)
#endregion EXOSettings Global Variables

#region EXOSettings Main Functions
# -----------------------------------------------
# EXOSettings Main Functions
# -----------------------------------------------

#region EXOSettings Tab Initialization
function Initialize-EXOSettingsTab {
    [CmdletBinding()]
    param()
    
    try {
        Write-DebugMessage "Initialisiere EXO Settings Tab" -Type "Info"
        
        # Textfeld für Status finden
        if ($null -eq $script:txtStatus) {
            $script:txtStatus = Get-XamlElement -ElementName "txtStatus"
        }
        
        # Event-Handler für Help-Link
        $helpLinkEXOSettings = Get-XamlElement -ElementName "helpLinkEXOSettings"
        if ($null -ne $helpLinkEXOSettings) {
            $helpLinkEXOSettings.Add_MouseLeftButtonDown({
                Start-Process "https://learn.microsoft.com/de-de/powershell/module/exchange/set-organizationconfig?view=exchange-ps"
            })
            
            # Mauszeiger-Styling
            $helpLinkEXOSettings.Add_MouseEnter({
                $this.Cursor = [System.Windows.Input.Cursors]::Hand
                $this.TextDecorations = [System.Windows.TextDecorations]::Underline
            })
            
            $helpLinkEXOSettings.Add_MouseLeave({
                $this.Cursor = [System.Windows.Input.Cursors]::Arrow
                $this.TextDecorations = $null
            })
        }
        
        # Event-Handler für "Aktuelle Einstellungen laden" Button
        $btnGetOrganizationConfig = Get-XamlElement -ElementName "btnGetOrganizationConfig"
        if ($null -ne $btnGetOrganizationConfig) {
            $btnGetOrganizationConfig.Add_Click({
                Get-CurrentOrganizationConfig
            })
        }
        
        # Event-Handler für "Einstellungen speichern" Button
        $btnSetOrganizationConfig = Get-XamlElement -ElementName "btnSetOrganizationConfig"
        if ($null -ne $btnSetOrganizationConfig) {
            $btnSetOrganizationConfig.Add_Click({
                Set-CustomOrganizationConfig
            })
        }
        
        # Event-Handler für "Konfiguration exportieren" Button
        $btnExportOrgConfig = Get-XamlElement -ElementName "btnExportOrgConfig"
        if ($null -ne $btnExportOrgConfig) {
            $btnExportOrgConfig.Add_Click({
                Export-OrganizationConfig
            })
        }
        
        # Initialize all UI controls for the OrganizationConfig tab
        Initialize-OrganizationConfigControls
        
        # Erstelle Log-Verzeichnis, falls Logging aktiviert ist
        if ($script:EXOSettingsLoggingEnabled) {
            $logDirectory = Split-Path -Path $script:EXOSettingsLogFilePath -Parent
            if (-not (Test-Path -Path $logDirectory)) {
                New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
            }
        }
        
        Write-DebugMessage "EXO Settings Tab wurde erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Initialisieren des EXO Settings Tabs: $errorMsg" -Type "Error"
        return $false
    }
}
#endregion EXOSettings Tab Initialization

#region EXOSettings UI Elements Initialization
function Initialize-OrganizationConfigControls {
    [CmdletBinding()]
    param()
    
    try {
        Write-DebugMessage "Initialisiere OrganizationConfig UI-Elemente" -Type "Info"
        
        # Eine Liste der Control-Typen und ihrer Event-Handler
        $controlHandlers = @{
            # Checkboxen: Jede Checkbox mit "chk"-Präfix entspricht einer OrganizationConfig-Einstellung
            "CheckBox" = @{
                "EventName" = "Click"
                "Handler" = {
                    param($sender, $e)
                    
                    $checkBox = $sender
                    $checkBoxName = $checkBox.Name
                    
                    if ($checkBoxName -like "chk*" -and $checkBoxName.Length -gt 3) {
                        $propertyName = $checkBoxName.Substring(3)
                        $script:organizationConfigSettings[$propertyName] = $checkBox.IsChecked
                        Write-DebugMessage "Checkbox $checkBoxName wurde auf $($checkBox.IsChecked) gesetzt" -Type "Info"
                    }
                }
            }
            
            # ComboBoxes: Spezifische ComboBoxen für bestimmte Einstellungen
            "ComboBox" = @{
                "EventName" = "SelectionChanged"
                "Handler" = {
                    param($sender, $e)
                    
                    $comboBox = $sender
                    $comboBoxName = $comboBox.Name
                    
                    if ($null -ne $comboBox.SelectedItem) {
                        $selectedContent = $comboBox.SelectedItem.Content.ToString()
                        
                        # Spezielle Verarbeitung für bestimmte ComboBoxen
                        switch ($comboBoxName) {
                            "chkActivityBasedAuthenticationTimeoutInterval" {
                                # Format: "01:00:00 (1h)" zu "01:00:00"
                                $value = ($selectedContent -split ' ')[0]
                                $script:organizationConfigSettings["ActivityBasedAuthenticationTimeoutInterval"] = $value
                                Write-DebugMessage "$comboBoxName geändert zu: $value" -Type "Info"
                            }
                            "cmbLargeAudienceThreshold" {
                                # Konvertieren zu Integer
                                $value = [int]$selectedContent
                                $script:organizationConfigSettings["MailTipsLargeAudienceThreshold"] = $value
                                Write-DebugMessage "$comboBoxName geändert zu: $value" -Type "Info"
                            }
                            "cmbInformationBarrierMode" {
                                $script:organizationConfigSettings["InformationBarrierMode"] = $selectedContent
                                Write-DebugMessage "$comboBoxName geändert zu: $selectedContent" -Type "Info"
                            }
                            "cmbEwsAppAccessPolicy" {
                                $script:organizationConfigSettings["EwsApplicationAccessPolicy"] = $selectedContent
                                Write-DebugMessage "$comboBoxName geändert zu: $selectedContent" -Type "Info"
                            }
                            "cmbOfficeFeatures" {
                                $script:organizationConfigSettings["OfficeFeatures"] = $selectedContent
                                Write-DebugMessage "$comboBoxName geändert zu: $selectedContent" -Type "Info"
                            }
                            "cmbSearchQueryLanguage" {
                                $script:organizationConfigSettings["SearchQueryLanguage"] = $selectedContent
                                Write-DebugMessage "$comboBoxName geändert zu: $selectedContent" -Type "Info"
                            }
                            default {
                                # Generische Behandlung für andere ComboBoxen
                                if ($comboBoxName -like "cmb*" -and $comboBoxName.Length -gt 3) {
                                    $propertyName = $comboBoxName.Substring(3)
                                    $script:organizationConfigSettings[$propertyName] = $selectedContent
                                    Write-DebugMessage "$comboBoxName geändert zu: $selectedContent" -Type "Info"
                                }
                            }
                        }
                    }
                }
            }
            
            # TextBoxes: Für numerische und Text-Einstellungen
            "TextBox" = @{
                "EventName" = "TextChanged"
                "Handler" = {
                    param($sender, $e)
                    
                    $textBox = $sender
                    $textBoxName = $textBox.Name
                    
                    # Spezielle Verarbeitung für bestimmte TextBoxen
                    switch ($textBoxName) {
                        "txtPowerShellMaxConcurrency" {
                            if ([int]::TryParse($textBox.Text, [ref]$null)) {
                                $script:organizationConfigSettings["PowerShellMaxConcurrency"] = [int]$textBox.Text
                                Write-DebugMessage "$textBoxName geändert zu: $($textBox.Text)" -Type "Info"
                            }
                        }
                        "txtPowerShellMaxCmdletQueueDepth" {
                            if ([int]::TryParse($textBox.Text, [ref]$null)) {
                                $script:organizationConfigSettings["PowerShellMaxCmdletQueueDepth"] = [int]$textBox.Text
                                Write-DebugMessage "$textBoxName geändert zu: $($textBox.Text)" -Type "Info"
                            }
                        }
                        "txtPowerShellMaxCmdletsExecutionDuration" {
                            if ([int]::TryParse($textBox.Text, [ref]$null)) {
                                $script:organizationConfigSettings["PowerShellMaxCmdletsExecutionDuration"] = [int]$textBox.Text
                                Write-DebugMessage "$textBoxName geändert zu: $($textBox.Text)" -Type "Info"
                            }
                        }
                        "txtDefaultAuthPolicy" {
                            $script:organizationConfigSettings["DefaultAuthenticationPolicy"] = $textBox.Text
                            Write-DebugMessage "$textBoxName geändert zu: $($textBox.Text)" -Type "Info"
                        }
                        "txtHierAddressBookRoot" {
                            $script:organizationConfigSettings["HierarchicalAddressBookRoot"] = $textBox.Text
                            Write-DebugMessage "$textBoxName geändert zu: $($textBox.Text)" -Type "Info"
                        }
                        "txtPreferredInternetCodePageForShiftJis" {
                            if ([int]::TryParse($textBox.Text, [ref]$null)) {
                                $script:organizationConfigSettings["PreferredInternetCodePageForShiftJis"] = [int]$textBox.Text
                                Write-DebugMessage "$textBoxName geändert zu: $($textBox.Text)" -Type "Info"
                            }
                        }
                        default {
                            # Generische Behandlung für andere TextBoxen
                            if ($textBoxName -like "txt*" -and $textBoxName.Length -gt 3) {
                                $propertyName = $textBoxName.Substring(3)
                                $script:organizationConfigSettings[$propertyName] = $textBox.Text
                                Write-DebugMessage "$textBoxName geändert zu: $($textBox.Text)" -Type "Info"
                            }
                        }
                    }
                }
            }
        }
        
        # TabOrgSettings finden, um alle UI-Elemente zu durchsuchen
        $tabOrgSettings = Get-XamlElement -ElementName "tabOrgSettings"
        if ($null -eq $tabOrgSettings) {
            throw "TabControl 'tabOrgSettings' nicht gefunden"
        }
        
        # Statistiken für die Registrierung
        $registeredControls = @{
            "CheckBox" = 0
            "ComboBox" = 0
            "TextBox" = 0
        }
        
        # Registriere nur Event-Handler für bekannte Elemente, die in der XAML definiert sind
        foreach ($elementName in $script:knownUIElements) {
            $element = Get-XamlElement -ElementName $elementName
            if ($null -ne $element) {
                $elementType = $element.GetType().Name
                if ($controlHandlers.ContainsKey($elementType)) {
                    $handler = $controlHandlers[$elementType]
                    $eventName = "Add_" + $handler.EventName
                    
                    try {
                        # Event-Handler registrieren
                        $element.$eventName($handler.Handler)
                        $registeredControls[$elementType]++
                        Write-DebugMessage "Event-Handler für $elementName ($elementType) registriert" -Type "Info"
                    }
                    catch {
                        Write-DebugMessage "Fehler beim Registrieren des Event-Handlers für $elementName - $($_.Exception.Message)" -Type "Warning"
                    }
                }
            }
        }
        
        # Spezialbehandlung der ActivityTimeout ComboBox, falls nicht bereits registriert
        $cmbActivityTimeout = Get-XamlElement -ElementName "chkActivityBasedAuthenticationTimeoutInterval"
        if ($null -ne $cmbActivityTimeout -and ($cmbActivityTimeout -is [System.Windows.Controls.ComboBox])) {
            if (-not ($registeredControls["ComboBox"] -gt 0)) {
                $cmbActivityTimeout.Add_SelectionChanged({
                    param($sender, $e)
                    
                    if ($null -ne $sender.SelectedItem) {
                        $selectedContent = $sender.SelectedItem.Content.ToString()
                        $value = ($selectedContent -split ' ')[0]  # "01:00:00 (1h)" zu "01:00:00" extrahieren
                        $script:organizationConfigSettings["ActivityBasedAuthenticationTimeoutInterval"] = $value
                        Write-DebugMessage "ActivityTimeout geändert zu: $value" -Type "Info"
                    }
                })
                $registeredControls["ComboBox"]++
                Write-DebugMessage "Event-Handler für ActivityTimeout ComboBox registriert" -Type "Info"
            }
        }
        
        Write-DebugMessage ("UI-Elemente registriert: " + 
                          "$($registeredControls.CheckBox) Checkboxen, " +
                          "$($registeredControls.ComboBox) ComboBoxen, " + 
                          "$($registeredControls.TextBox) TextBoxen") -Type "Info"
        
        Write-DebugMessage "OrganizationConfig UI-Elemente erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Initialisieren der OrganizationConfig UI-Elemente - $errorMsg" -Type "Error"
        return $false
    }
}

#region EXOSettings Organization Config Management
function Get-CurrentOrganizationConfig {
    [CmdletBinding()]
    param()
    
    try {
        # Prüfen, ob wir mit Exchange verbunden sind
        if (-not (Confirm-ExchangeConnection)) {
            Show-MessageBox -Message "Bitte verbinden Sie sich zuerst mit Exchange Online." -Title "Nicht verbunden" -Type "Warning"
            return
        }
        
        Write-DebugMessage "Rufe aktuelle Organisationseinstellungen ab" -Type "Info"
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Lade Organisationseinstellungen..."
        }
        
        # Organisationseinstellungen abrufen
        $script:currentOrganizationConfig = Get-OrganizationConfig -ErrorAction Stop
        $configProperties = $script:currentOrganizationConfig | Get-Member -MemberType Properties | 
                            Where-Object { $_.Name -notlike "__*" } | 
                            Select-Object -ExpandProperty Name
        
        # Aktuelle Einstellungs-Hashtable leeren
        $script:organizationConfigSettings = @{}
        
        Write-DebugMessage "Organisation-Config geladen, aktualisiere UI-Elemente" -Type "Info"
        
        # UI-Elemente nur für Controls aktualisieren, die in der XAML existieren
        foreach ($elementName in $script:knownUIElements) {
            $element = Get-XamlElement -ElementName $elementName
            if ($null -ne $element) {
                # Passende Eigenschaft basierend auf Element-Namen finden
                if ($element -is [System.Windows.Controls.CheckBox] -and $elementName -like "chk*") {
                    $propertyName = $elementName.Substring(3)
                    if ($configProperties -contains $propertyName) {
                        $value = [bool]$script:currentOrganizationConfig.$propertyName
                        $element.IsChecked = $value
                        $script:organizationConfigSettings[$propertyName] = $value
                        Write-DebugMessage "Checkbox $elementName auf $value gesetzt" -Type "Info"
                    }
                }
            }
        }
        
        # Spezielle ComboBoxen behandeln
        
        # ActivityBasedAuthenticationTimeoutInterval - korrekter Name aus der XAML
        $cmbActivityTimeout = Get-XamlElement -ElementName "chkActivityBasedAuthenticationTimeoutInterval"
        if ($null -ne $cmbActivityTimeout -and ($cmbActivityTimeout -is [System.Windows.Controls.ComboBox]) -and
            $null -ne $script:currentOrganizationConfig.ActivityBasedAuthenticationTimeoutInterval) {
            
            $timeoutValue = $script:currentOrganizationConfig.ActivityBasedAuthenticationTimeoutInterval.ToString()
            $matchFound = $false
            
            foreach ($item in $cmbActivityTimeout.Items) {
                $itemContent = $item.Content.ToString()
                if ($itemContent.StartsWith($timeoutValue)) {
                    $cmbActivityTimeout.SelectedItem = $item
                    $script:organizationConfigSettings["ActivityBasedAuthenticationTimeoutInterval"] = $timeoutValue
                    Write-DebugMessage "ActivityTimeout auf $timeoutValue gesetzt" -Type "Info"
                    $matchFound = $true
                    break
                }
            }
            
            # Wenn kein Match gefunden, Standardwert auswählen
            if (-not $matchFound -and $cmbActivityTimeout.Items.Count -gt 0) {
                $cmbActivityTimeout.SelectedIndex = 1 # Default to 6h
                $selectedContent = $cmbActivityTimeout.SelectedItem.Content.ToString()
                $value = ($selectedContent -split ' ')[0]
                $script:organizationConfigSettings["ActivityBasedAuthenticationTimeoutInterval"] = $value
                Write-DebugMessage "ActivityTimeout auf Standard $value gesetzt (kein Wert gefunden)" -Type "Info"
            }
        }
        
        # LargeAudienceThreshold
        $cmbLargeAudienceThreshold = Get-XamlElement -ElementName "cmbLargeAudienceThreshold"
        if ($null -ne $cmbLargeAudienceThreshold && 
            $null -ne $script:currentOrganizationConfig.MailTipsLargeAudienceThreshold) {
            
            $thresholdValue = $script:currentOrganizationConfig.MailTipsLargeAudienceThreshold.ToString()
            $matchFound = $false
            
            foreach ($item in $cmbLargeAudienceThreshold.Items) {
                if ($item.Content.ToString() -eq $thresholdValue) {
                    $cmbLargeAudienceThreshold.SelectedItem = $item
                    $script:organizationConfigSettings["MailTipsLargeAudienceThreshold"] = [int]$thresholdValue
                    Write-DebugMessage "LargeAudienceThreshold auf $thresholdValue gesetzt" -Type "Info"
                    $matchFound = $true
                    break
                }
            }
            
            # Wenn kein Match gefunden, Standardwert auswählen
            if (-not $matchFound -and $cmbLargeAudienceThreshold.Items.Count -gt 0) {
                $cmbLargeAudienceThreshold.SelectedIndex = 1 # Default zu 50
                $script:organizationConfigSettings["MailTipsLargeAudienceThreshold"] = [int]$cmbLargeAudienceThreshold.SelectedItem.Content
                Write-DebugMessage "LargeAudienceThreshold auf Standard gesetzt (kein Wert gefunden)" -Type "Info"
            }
        }
        
        # Andere spezielle ComboBoxen aktualisieren
        $comboBoxMappings = @{
            "cmbInformationBarrierMode" = "InformationBarrierMode"
            "cmbEwsAppAccessPolicy" = "EwsApplicationAccessPolicy"
            "cmbOfficeFeatures" = "OfficeFeatures"
            "cmbSearchQueryLanguage" = "SearchQueryLanguage"
        }
        
        foreach ($comboBoxName in $comboBoxMappings.Keys) {
            $comboBox = Get-XamlElement -ElementName $comboBoxName
            $propertyName = $comboBoxMappings[$comboBoxName]
            
            if ($null -ne $comboBox -and $configProperties -contains $propertyName) {
                $value = $script:currentOrganizationConfig.$propertyName
                if ($null -ne $value) {
                    $matchFound = $false
                    
                    foreach ($item in $comboBox.Items) {
                        if ($item.Content.ToString() -eq $value) {
                            $comboBox.SelectedItem = $item
                            $script:organizationConfigSettings[$propertyName] = $value
                            Write-DebugMessage "$comboBoxName auf $value gesetzt" -Type "Info"
                            $matchFound = $true
                            break
                        }
                    }
                    
                    # Wenn kein Match gefunden, erstes Element als Standard auswählen
                    if (-not $matchFound -and $comboBox.Items.Count -gt 0) {
                        $comboBox.SelectedIndex = 0
                        $script:organizationConfigSettings[$propertyName] = $comboBox.SelectedItem.Content.ToString()
                        Write-DebugMessage "$comboBoxName auf Standard gesetzt (kein passender Wert gefunden)" -Type "Info"
                    }
                }
            }
        }
        
        # TextBox-Werte setzen
        $textBoxMappings = @{
            "txtPowerShellMaxConcurrency" = "PowerShellMaxConcurrency"
            "txtPowerShellMaxCmdletQueueDepth" = "PowerShellMaxCmdletQueueDepth"
            "txtPowerShellMaxCmdletsExecutionDuration" = "PowerShellMaxCmdletsExecutionDuration"
            "txtDefaultAuthPolicy" = "DefaultAuthenticationPolicy"
            "txtHierAddressBookRoot" = "HierarchicalAddressBookRoot"
            "txtPreferredInternetCodePageForShiftJis" = "PreferredInternetCodePageForShiftJis"
        }
        
        foreach ($textBoxName in $textBoxMappings.Keys) {
            $textBox = Get-XamlElement -ElementName $textBoxName
            $propertyName = $textBoxMappings[$textBoxName]
            
            if ($null -ne $textBox -and $configProperties -contains $propertyName) {
                $value = $script:currentOrganizationConfig.$propertyName
                if ($null -ne $value) {
                    $textBox.Text = $value.ToString()
                    $script:organizationConfigSettings[$propertyName] = $value
                    Write-DebugMessage "$textBoxName auf $value gesetzt" -Type "Info"
                }
                else {
                    $textBox.Text = ""
                }
            }
        }
        
        # Vollständige Konfiguration in TextBox anzeigen
        $configText = $script:currentOrganizationConfig | Format-List | Out-String
        $txtOrganizationConfig = Get-XamlElement -ElementName "txtOrganizationConfig"
        if ($null -ne $txtOrganizationConfig) {
            $txtOrganizationConfig.Text = $configText
        }
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organisationseinstellungen erfolgreich geladen."
        }
        Write-DebugMessage "Organisationseinstellungen erfolgreich abgerufen und angezeigt" -Type "Success"
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler beim Laden der Organisationseinstellungen: $errorMsg"
        }
        Write-DebugMessage "Fehler beim Abrufen der Organisationseinstellungen: $errorMsg" -Type "Error"
        Show-MessageBox -Message "Fehler beim Laden der Organisationseinstellungen: $errorMsg" -Title "Fehler" -Type "Error"
    }
}

function Set-CustomOrganizationConfig {
    [CmdletBinding()]
    param()
    
    try {
        # Prüfen, ob wir mit Exchange verbunden sind
        if (-not (Confirm-ExchangeConnection)) {
            Show-MessageBox -Message "Bitte verbinden Sie sich zuerst mit Exchange Online." -Title "Nicht verbunden" -Type "Warning"
            return
        }
        
        # Bestätigungsdialog anzeigen
        $confirmResult = Show-MessageBox -Message "Möchten Sie die Organisationseinstellungen wirklich speichern?" -Title "Einstellungen speichern" -Type "Question"
        if ($confirmResult -ne "Yes") {
            return
        }
        
        Write-DebugMessage "Speichere Organisationseinstellungen" -Type "Info"
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Speichere Organisationseinstellungen..."
        }
        
        # Aktualisiere Einstellungen von allen Eingabefeldern um sicherzustellen, dass sie aktuell sind
        # Behandle CheckBoxen
        foreach ($elementName in $script:knownUIElements) {
            $element = Get-XamlElement -ElementName $elementName
            if ($null -ne $element) {
                if ($element -is [System.Windows.Controls.CheckBox] -and $elementName -like "chk*") {
                    $propertyName = $elementName.Substring(3)
                    $script:organizationConfigSettings[$propertyName] = $element.IsChecked
                }
            }
        }
        
        # Behandle ComboBox - ActivityBasedAuthenticationTimeoutInterval (Spezialfall aufgrund der Benennung)
        # Behandle ComboBox - ActivityBasedAuthenticationTimeoutInterval (Spezialfall aufgrund der Benennung)
        $cmbActivityTimeout = Get-XamlElement -ElementName "chkActivityBasedAuthenticationTimeoutInterval"
        if ($null -ne $cmbActivityTimeout -and 
            $null -ne $cmbActivityTimeout.SelectedItem -and 
            $null -ne $cmbActivityTimeout.SelectedItem.Content) {
            $selectedText = $cmbActivityTimeout.SelectedItem.Content.ToString()
            if ($selectedText -match '^\d+') {
                $timeoutValue = ($selectedText -split ' ')[0]
                $script:organizationConfigSettings["ActivityBasedAuthenticationTimeoutInterval"] = $timeoutValue
                Write-DebugMessage "ActivityTimeout gesetzt auf: $timeoutValue" -Type "Info"
            } else {
                Write-DebugMessage "Ungültiges Format für ActivityTimeout: $selectedText" -Type "Warning"
            }
        }
// ... existing code ...
        
        # Behandle andere ComboBoxen
        $comboBoxMappings = @{
            "cmbLargeAudienceThreshold" = "MailTipsLargeAudienceThreshold" 
            "cmbInformationBarrierMode" = "InformationBarrierMode"
            "cmbEwsAppAccessPolicy" = "EwsApplicationAccessPolicy"
            "cmbOfficeFeatures" = "OfficeFeatures"
            "cmbSearchQueryLanguage" = "SearchQueryLanguage"
        }
        
        foreach ($comboBoxName in $comboBoxMappings.Keys) {
            $comboBox = Get-XamlElement -ElementName $comboBoxName
            $propertyName = $comboBoxMappings[$comboBoxName]
            
            if ($null -ne $comboBox -and $null -ne $comboBox.SelectedItem) {
                $selectedValue = $comboBox.SelectedItem.Content.ToString()
                
                # Spezialbehandlung für numerische Werte
                if ($comboBoxName -eq "cmbLargeAudienceThreshold") {
                    $script:organizationConfigSettings[$propertyName] = [int]$selectedValue
                }
                else {
                    $script:organizationConfigSettings[$propertyName] = $selectedValue
                }
            }
        }
        
        # Behandle TextBoxen
        $textBoxMappings = @{
            "txtPowerShellMaxConcurrency" = "PowerShellMaxConcurrency"
            "txtPowerShellMaxCmdletQueueDepth" = "PowerShellMaxCmdletQueueDepth"
            "txtPowerShellMaxCmdletsExecutionDuration" = "PowerShellMaxCmdletsExecutionDuration"
            "txtDefaultAuthPolicy" = "DefaultAuthenticationPolicy"
            "txtHierAddressBookRoot" = "HierarchicalAddressBookRoot"
            "txtPreferredInternetCodePageForShiftJis" = "PreferredInternetCodePageForShiftJis"
        }
        
        foreach ($textBoxName in $textBoxMappings.Keys) {
            $textBox = Get-XamlElement -ElementName $textBoxName
            $propertyName = $textBoxMappings[$textBoxName]
            
            if ($null -ne $textBox -and -not [string]::IsNullOrWhiteSpace($textBox.Text)) {
                $textValue = $textBox.Text.Trim()
                
                # Behandle numerische Werte
                # Behandle numerische Werte
                if ($textBoxName -eq "txtPowerShellMaxConcurrency" -or 
                    $textBoxName -eq "txtPowerShellMaxCmdletQueueDepth" -or 
                    $textBoxName -eq "txtPowerShellMaxCmdletsExecutionDuration") {
                    [int]$numValue = 0
                    if ([int]::TryParse($textValue, [ref]$numValue)) {
                        if ($numValue -gt 0) {
                            $script:organizationConfigSettings[$propertyName] = $numValue
                            Write-DebugMessage "Numerischer Wert für $propertyName gesetzt: $numValue" -Type "Info"
                        } else {
                            Write-DebugMessage "Wert für $propertyName muss größer als 0 sein" -Type "Warning"
                            Show-MessageBox -Message "Der Wert für $propertyName muss größer als 0 sein." -Title "Ungültiger Wert" -Type "Warning"
                            return
                        }
                    } else {
                        Write-DebugMessage "Ungültiger numerischer Wert für $propertyName - $textValue" -Type "Warning"
                        Show-MessageBox -Message "Ungültiger numerischer Wert für $propertyName" -Title "Fehler" -Type "Warning"
                        return
                    }
                }
            }
        }
        
        # Parameter für Set-OrganizationConfig vorbereiten
        $params = @{}
        foreach ($key in $script:organizationConfigSettings.Keys) {
            if ($null -ne $script:organizationConfigSettings[$key]) {
                $params[$key] = $script:organizationConfigSettings[$key]
                
                # Zusätzliche Validierung für bekannte Parameter
                switch ($key) {
                    "PowerShellMaxConcurrency" {
                        if ($params[$key] -gt 100) {
                            Write-DebugMessage "Warnung: PowerShellMaxConcurrency ist sehr hoch: $($params[$key])" -Type "Warning"
                        }
                    }
                    "PowerShellMaxCmdletQueueDepth" {
                        if ($params[$key] -gt 1000) {
                            Write-DebugMessage "Warnung: PowerShellMaxCmdletQueueDepth ist sehr hoch: $($params[$key])" -Type "Warning"
                        }
                    }
                }
            }
        }
        
        # Debug-Log alle Parameter
        foreach ($key in $params.Keys | Sort-Object) {
            Write-DebugMessage "Parameter: $key = $($params[$key])" -Type "Info"
        }
        
        # Organisationseinstellungen aktualisieren
        Set-OrganizationConfig @params -ErrorAction Stop
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organisationseinstellungen erfolgreich gespeichert."
        }
        Write-DebugMessage "Organisationseinstellungen erfolgreich gespeichert" -Type "Success"
        Show-MessageBox -Message "Die Organisationseinstellungen wurden erfolgreich gespeichert." -Title "Erfolg" -Type "Info"
        
        # Aktuelle Konfiguration neu laden, um Änderungen zu sehen
        Get-CurrentOrganizationConfig
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler beim Speichern der Organisationseinstellungen: $errorMsg"
        }
        Write-DebugMessage "Fehler beim Speichern der Organisationseinstellungen: $errorMsg" -Type "Error"
        Show-MessageBox -Message "Fehler beim Speichern der Organisationseinstellungen: $errorMsg" -Title "Fehler" -Type "Error"
    }
}

function Export-OrganizationConfig {
    [CmdletBinding()]
    param()
    
    try {
        # Prüfen, ob wir mit Exchange verbunden sind
        if (-not (Confirm-ExchangeConnection)) {
            Show-MessageBox -Message "Bitte verbinden Sie sich zuerst mit Exchange Online." -Title "Nicht verbunden" -Type "Warning"
            return
        }
        
        # Prüfen, ob aktuelle Konfiguration verfügbar ist
        if ($null -eq $script:currentOrganizationConfig) {
            # Versuche zuerst die Konfiguration zu laden
            Write-DebugMessage "Keine aktuelle Konfiguration gefunden, lade Konfiguration..." -Type "Info"
            Get-CurrentOrganizationConfig
            
            if ($null -eq $script:currentOrganizationConfig) {
                Show-MessageBox -Message "Die Organisationseinstellungen konnten nicht geladen werden." -Title "Fehler" -Type "Error"
                return
            }
        }
        
        # SaveFileDialog anzeigen
        $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Text-Dateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
        $saveFileDialog.FileName = "ExchangeOnline_OrgConfig_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        $saveFileDialog.DefaultExt = ".csv"
        $saveFileDialog.AddExtension = $true
        
        $result = $saveFileDialog.ShowDialog()
        if ($result -ne $true) {
            return
        }
        
        $exportPath = $saveFileDialog.FileName
        $fileExtension = [System.IO.Path]::GetExtension($exportPath)
        
        if ($fileExtension -eq ".csv") {
            # Als CSV exportieren
            $script:currentOrganizationConfig | 
                Select-Object * -ExcludeProperty RunspaceId, PSComputerName, PSShowComputerName, PSSourceJobInstanceId | 
                Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        }
        else {
            # Als Text exportieren
            $script:currentOrganizationConfig | Format-List | Out-File -FilePath $exportPath -Encoding utf8
        }
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organisationseinstellungen exportiert nach $exportPath"
        }
        Write-DebugMessage "Organisationseinstellungen erfolgreich exportiert nach $exportPath" -Type "Success"
        Show-MessageBox -Message "Die Organisationseinstellungen wurden erfolgreich nach '$exportPath' exportiert." -Title "Export erfolgreich" -Type "Info"
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler beim Exportieren der Organisationseinstellungen: $errorMsg"
        }
        Write-DebugMessage "Fehler beim Exportieren der Organisationseinstellungen: $errorMsg" -Type "Error"
        Show-MessageBox -Message "Fehler beim Exportieren der Organisationseinstellungen: $errorMsg" -Title "Fehler" -Type "Error"
    }
}
#endregion EXOSettings Organization Config Management

function Set-EXOSettingsDebugLogging {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Enable,
        
        [Parameter(Mandatory = $false)]
        [string]$LogFilePath
    )
    
    $script:EXOSettingsLoggingEnabled = $Enable
    
    if ($PSBoundParameters.ContainsKey('LogFilePath')) {
        $script:EXOSettingsLogFilePath = $LogFilePath
    }
    
    if ($Enable) {
        $logDirectory = Split-Path -Path $script:EXOSettingsLogFilePath -Parent
        if (-not (Test-Path -Path $logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }
        
        Write-DebugMessage "Debug-Logging aktiviert. Log-Datei: $script:EXOSettingsLogFilePath" -Type "Info"
    }
    else {
        Write-DebugMessage "Debug-Logging deaktiviert" -Type "Info"
    }
}
#endregion EXOSettings Main Functions

function Set-EXOSettingsDebugLogging {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Enable,
        
        [Parameter(Mandatory = $false)]
        [string]$LogFilePath
    )
    
    $script:EXOSettingsLoggingEnabled = $Enable
    
    if ($PSBoundParameters.ContainsKey('LogFilePath')) {
        $script:EXOSettingsLogFilePath = $LogFilePath
    }
    
    if ($Enable) {
        $logDirectory = Split-Path -Path $script:EXOSettingsLogFilePath -Parent
        if (-not (Test-Path -Path $logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }
        
        Write-DebugMessage "Debug logging enabled. Log file: $script:EXOSettingsLogFilePath" -Type "Info"
    }
    else {
        Write-DebugMessage "Debug logging disabled" -Type "Info"
    }
}
#endregion EXOSettings Main Functions
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
                    if ($null -ne $script:lstGroupMembers -and $script:lstGroupMembers.Items -ne $null && $script:lstGroupMembers.Items.Count -gt 0) {
                        if ($null -ne $script:txtStatus) {
                            $script:txtStatus.Text = "$($members.Count) Gruppenmitglieder gefunden."
                        }
                    } else {
                        if ($null -ne $script:txtStatus) {
                            $script:txtStatus.Text = "Gruppe ist leer oder nicht geladen."
                        }
                    }c
                    
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
    function RefreshSharedMailboxList {
        [CmdletBinding()]
        param()
        
        try {
            if ($null -ne $script:cmbSharedMailboxSelect) {
                $script:cmbSharedMailboxSelect.Items.Clear()
                
                # Prüfen, ob eine Verbindung besteht
                if (-not $script:isConnected) {
                    Write-DebugMessage "Keine Verbindung zu Exchange Online für Shared Mailbox-Liste" -Type "Warning"
                    return $false
                }
                
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
                        
                        # Prüfen, ob eine Verbindung besteht
                        if (-not $script:isConnected) {
                            Write-DebugMessage "Keine Verbindung zu Exchange Online für Shared Mailbox-Liste" -Type "Warning"
                            return $false
                        }
                        
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
                        $result = Convert- ToSharedMailboxAction -Identity $script:txtSharedMailboxEmail.Text
                        
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
                    
                    # Übernehme die ausgewählte Mailbox ins Eingabefeld für besseren Workflow
                    if ($null -ne $script:txtSharedMailboxPermSource) {
                        $script:txtSharedMailboxPermSource.Text = $mailbox
                    }
                    
                    # Übernehme die ausgewählte Mailbox auch in die Weiterleitungs-ComboBox
                    if ($null -ne $script:cmbForwardingMailbox) {
                        # Prüfe, ob der Wert in den Items existiert
                        $found = $false
                        foreach ($item in $script:cmbForwardingMailbox.Items) {
                            if ($item -eq $mailbox) {
                                $script:cmbForwardingMailbox.SelectedItem = $item
                                $found = $true
                                break
                            }
                        }
                        
                        # Falls der Wert noch nicht in der ComboBox ist, füge ihn hinzu
                        if (-not $found) {
                            $script:cmbForwardingMailbox.Items.Add($mailbox)
                            $script:cmbForwardingMailbox.SelectedItem = $mailbox
                        }
                    }
                    
                    # Funktion zum Anzeigen der Berechtigungen aufrufen
                    $permissions = Get-SharedMailboxPermissionsAction -Mailbox $mailbox
                    
                    # Berechtigungen in der Benutzeroberfläche anzeigen
                    if ($permissions.Count -gt 0) {
                        # DataGrid für die Anzeige der Berechtigungen verwenden
                        $lstCurrentPermissions = Get-XamlElement -ElementName "lstCurrentPermissions"
                        
                        if ($null -ne $lstCurrentPermissions) {
                            # DataGrid leeren und mit den Berechtigungsdaten füllen
                            $permissionsCollection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                            
                            foreach ($perm in $permissions) {
                                $permObject = [PSCustomObject]@{
                                    Mailbox = $mailbox
                                    User = $perm.User
                                    PermissionType = $perm.PermissionType
                                    AccessRights = $perm.AccessRights
                                    AutoMapping = "N/A"
                                    AddedDate = (Get-Date).ToString("yyyy-MM-dd")
                                }
                                $permissionsCollection.Add($permObject)
                            }
                            
                            $lstCurrentPermissions.ItemsSource = $permissionsCollection
                            
                            $script:txtStatus.Text = "$($permissions.Count) Berechtigungen gefunden für $mailbox."
                        } else {
                            Write-DebugMessage "DataGrid 'lstCurrentPermissions' wurde nicht gefunden." -Type "Warning"
                            
                            # Fallback: MessageBox anzeigen
                            $permText = "Berechtigungen für Shared Mailbox '$mailbox':`n`n"
                            foreach ($perm in $permissions) {
                                $permText += "- $($perm.User): $($perm.AccessRights) ($($perm.PermissionType))`n"
                            }
                            
                            [System.Windows.MessageBox]::Show(
                                $permText,
                                "Shared Mailbox Berechtigungen",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Information)
                            
                            $script:txtStatus.Text = "$($permissions.Count) Berechtigungen gefunden für $mailbox."
                        }
                    } else {
                        # Berechtigungsliste leeren
                        $lstCurrentPermissions = Get-XamlElement -ElementName "lstCurrentPermissions"
                        if ($null -ne $lstCurrentPermissions) {
                            $lstCurrentPermissions.ItemsSource = $null
                        }
                        
                        $script:txtStatus.Text = "Keine Berechtigungen gefunden für $mailbox."
                    }
        }
        catch {
            $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Anzeigen der Shared Mailbox Berechtigungen: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnShowSharedMailboxPerms"
            
            Register-EventHandler -Control $btnGetForwardingMailboxes -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
                    # Funktion zum Anzeigen der Shared Mailboxes für Weiterleitung aufrufen
                    $cmbForwardingMailbox = Get-XamlElement -ElementName "cmbForwardingMailbox"
                    if ($null -ne $cmbForwardingMailbox) {
                        $cmbForwardingMailbox.Items.Clear()
                        
                        # Abrufen aller Shared Mailboxen
                        $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | 
                                          Select-Object -ExpandProperty PrimarySmtpAddress
                        
                        foreach ($mailbox in $sharedMailboxes) {
                            [void]$cmbForwardingMailbox.Items.Add($mailbox)
                        }
                        
                        if ($cmbForwardingMailbox.Items.Count -gt 0) {
                            $cmbForwardingMailbox.SelectedIndex = 0
                            $script:txtStatus.Text = "$($cmbForwardingMailbox.Items.Count) Shared Mailboxes gefunden."
                        } else {
                            $script:txtStatus.Text = "Keine Shared Mailboxes gefunden."
                        }
                        
                        Write-DebugMessage "Shared Mailbox-Liste für Weiterleitung aktualisiert: $($cmbForwardingMailbox.Items.Count) gefunden" -Type "Info"
                    }
        }
        catch {
            $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Abrufen der Shared Mailboxes für Weiterleitung: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnGetForwardingMailboxes"

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
    function Initialize-ReportsTab {
        [CmdletBinding()]
        param()
        
        try {
            Write-DebugMessage "Initialisiere Berichte-Tab" -Type "Info"
            
            # UI-Elemente referenzieren
            $cmbReportType = Get-XamlElement -ElementName "cmbReportType"
            $dpReportStartDate = Get-XamlElement -ElementName "dpReportStartDate"
            $dpReportEndDate = Get-XamlElement -ElementName "dpReportEndDate"
            $btnGenerateReport = Get-XamlElement -ElementName "btnGenerateReport"
            $lstReportResults = Get-XamlElement -ElementName "lstReportResults"
            $btnExportReport = Get-XamlElement -ElementName "btnExportReport"
            $helpLinkReports = Get-XamlElement -ElementName "helpLinkReports"
            
            # Globale Variablen setzen
            $script:cmbReportType = $cmbReportType
            $script:dpReportStartDate = $dpReportStartDate
            $script:dpReportEndDate = $dpReportEndDate
            $script:lstReportResults = $lstReportResults
            
            # Event-Handler registrieren
            Register-EventHandler -Control $btnGenerateReport -Handler {
                try {
                    # Verbindungsprüfung
                    if (-not $script:isConnected) {
                        [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Parameter sammeln
                    $reportType = $script:cmbReportType.SelectedItem
                    if ($null -eq $reportType) {
                        [System.Windows.MessageBox]::Show("Bitte wählen Sie einen Berichtstyp aus.", 
                            "Keine Auswahl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    $reportName = $reportType.ToString()
                    if ($reportType -is [System.Windows.Controls.ComboBoxItem]) {
                        $reportName = $reportType.Content.ToString()
                    }
                    
                    # Start- und Enddatum abrufen (falls angegeben)
                    $startDate = $null
                    $endDate = $null
                    if ($null -ne $script:dpReportStartDate.SelectedDate -and $script:dpReportStartDate.SelectedDate -ne [DateTime]::MinValue) {
                        $startDate = $script:dpReportStartDate.SelectedDate
                    }
                    if ($null -ne $script:dpReportEndDate.SelectedDate -and $script:dpReportEndDate.SelectedDate -ne [DateTime]::MinValue) {
                        $endDate = $script:dpReportEndDate.SelectedDate
                    }
                    
                    # Status anzeigen
                    $script:txtStatus.Text = "Generiere Bericht: $reportName..."
                    
                    # Bericht basierend auf Typ generieren
                    $reportData = @()
                    
                    switch -Wildcard ($reportName) {
                        "*Postfachgrößen*" {
                            $reportData = Get-Mailbox -ResultSize Unlimited | 
                                Get-MailboxStatistics | 
                                Select-Object DisplayName, TotalItemSize, ItemCount, LastLogonTime, @{
                                    Name = "TotalSizeGB"
                                    Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1GB), 2) }
                                }
                        }
                        "*Letzte Anmeldung*" {
                            $reportData = Get-Mailbox -ResultSize Unlimited | 
                                Get-MailboxStatistics | 
                                Select-Object DisplayName, LastLogonTime, ItemCount
                        }
                        "*Postfachberechtigungen*" {
                            $reportData = @()
                            $mailboxes = Get-Mailbox -ResultSize Unlimited
                            
                            foreach ($mailbox in $mailboxes) {
                                $permissions = Get-MailboxPermission -Identity $mailbox.Identity | 
                                    Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.IsInherited -eq $false }
                                
                                foreach ($perm in $permissions) {
                                    $reportData += [PSCustomObject]@{
                                        Mailbox = $mailbox.DisplayName
                                        MailboxEmail = $mailbox.PrimarySmtpAddress
                                        User = $perm.User
                                        AccessRights = ($perm.AccessRights -join ", ")
                                    }
                                }
                            }
                        }
                        "*Kalenderberechtigungen*" {
                            $reportData = @()
                            $mailboxes = Get-Mailbox -ResultSize Unlimited
                            
                            foreach ($mailbox in $mailboxes) {
                                try {
                                    $calendarFolder = $mailbox.PrimarySmtpAddress.ToString() + ":\Kalender"
                                    $permissions = Get-MailboxFolderPermission -Identity $calendarFolder -ErrorAction SilentlyContinue
                                    
                                    if ($null -eq $permissions) {
                                        # Versuche englischen Kalender-Namen
                                        $calendarFolder = $mailbox.PrimarySmtpAddress.ToString() + ":\Calendar"
                                        $permissions = Get-MailboxFolderPermission -Identity $calendarFolder -ErrorAction SilentlyContinue
                                    }
                                    
                                    if ($null -ne $permissions) {
                                        foreach ($perm in $permissions) {
                                            if ($perm.User.ToString() -ne "Anonymous" -and $perm.User.ToString() -ne "Default") {
                                                $reportData += [PSCustomObject]@{
                                                    Mailbox = $mailbox.DisplayName
                                                    MailboxEmail = $mailbox.PrimarySmtpAddress
                                                    User = $perm.User
                                                    AccessRights = ($perm.AccessRights -join ", ")
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    # Fehler beim Abrufen der Kalenderberechtigungen für dieses Postfach, überspringe
                                    continue
                                }
                            }
                        }
                        "*Shared Mailbox*" {
                            $reportData = @()
                            $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
                            
                            foreach ($mailbox in $sharedMailboxes) {
                                $permissions = Get-MailboxPermission -Identity $mailbox.Identity | 
                                    Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.IsInherited -eq $false }
                                
                                foreach ($perm in $permissions) {
                                    $reportData += [PSCustomObject]@{
                                        Mailbox = $mailbox.DisplayName
                                        MailboxEmail = $mailbox.PrimarySmtpAddress
                                        User = $perm.User
                                        AccessRights = ($perm.AccessRights -join ", ")
                                        AutoMapping = "N/A" # AutoMapping müsste noch implementiert werden
                                    }
                                }
                            }
                        }
                        "*Gruppenmitglieder*" {
                            $reportData = @()
                            $groups = Get-DistributionGroup -ResultSize Unlimited
                            
                            foreach ($group in $groups) {
                                $members = Get-DistributionGroupMember -Identity $group.Identity
                                
                                foreach ($member in $members) {
                                    $reportData += [PSCustomObject]@{
                                        Gruppe = $group.DisplayName
                                        GruppenEmail = $group.PrimarySmtpAddress
                                        Mitglied = $member.DisplayName
                                        MitgliedEmail = $member.PrimarySmtpAddress
                                        MitgliedTyp = $member.RecipientType
                                    }
                                }
                            }
                        }
                        default {
                            [System.Windows.MessageBox]::Show("Der gewählte Berichtstyp wird derzeit nicht unterstützt.", 
                                "Nicht unterstützt", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                            return
                        }
                    }
                    
                    # Ergebnisse in der DataGrid anzeigen
                    $script:lstReportResults.ItemsSource = $reportData
                    
                    # Status aktualisieren
                    $script:txtStatus.Text = "Bericht generiert: $($reportData.Count) Datensätze gefunden."
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Generieren des Berichts: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                    [System.Windows.MessageBox]::Show("Fehler beim Generieren des Berichts: $errorMsg", 
                        "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            } -ControlName "btnGenerateReport"
            
            Register-EventHandler -Control $btnExportReport -Handler {
                try {
                    # Prüfen, ob Daten vorhanden sind
                    $data = $script:lstReportResults.ItemsSource
                    if ($null -eq $data -or $data.Count -eq 0) {
                        [System.Windows.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden. Bitte generieren Sie zuerst einen Bericht.", 
                            "Keine Daten", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # SaveFileDialog erstellen und konfigurieren
                    $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
                    $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv"
                    $saveFileDialog.Title = "Bericht exportieren"
                    $saveFileDialog.FileName = "Exchange_Bericht_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                    
                    # Dialog anzeigen
                    $result = $saveFileDialog.ShowDialog()
                    
                    # Ergebnis prüfen
                    if ($result -eq $true) {
                        # Daten in CSV exportieren
                        $data | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                        
                        # Status aktualisieren
                        $script:txtStatus.Text = "Bericht wurde erfolgreich exportiert: $($saveFileDialog.FileName)"
                        
                        # Erfolgsmeldung anzeigen
                        [System.Windows.MessageBox]::Show("Der Bericht wurde erfolgreich exportiert: $($saveFileDialog.FileName)", 
                            "Export erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Exportieren des Berichts: $errorMsg" -Type "Error"
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                    [System.Windows.MessageBox]::Show("Fehler beim Exportieren des Berichts: $errorMsg", 
                        "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            } -ControlName "btnExportReport"
            
            # Hilfe-Link initialisieren, falls vorhanden
            if ($null -ne $helpLinkReports) {
                $helpLinkReports.Add_MouseLeftButtonDown({
                    Show-HelpDialog -Topic "Reports"
                })
                
                $helpLinkReports.Add_MouseEnter({
                    $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                    $this.Cursor = [System.Windows.Input.Cursors]::Hand
                })
                
                $helpLinkReports.Add_MouseLeave({
                    $this.TextDecorations = $null
                    $this.Cursor = [System.Windows.Input.Cursors]::Arrow
                })
            }
            
            Write-DebugMessage "Berichte-Tab erfolgreich initialisiert" -Type "Success"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Initialisieren des Berichte-Tabs: $errorMsg" -Type "Error"
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
                    $this.Cursor = [System.Windows.Input.Cursors]::Hand
                    $this.TextDecorations = [System.Windows.TextDecorations]::Underline
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
                    $this.Cursor = [System.Windows.Input.Cursors]::Arrow
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
                EXOSettings = Initialize-EXOSettingsTab
                Calendar = Initialize-CalendarTab
                Mailbox = Initialize-MailboxTab
                Audit = Initialize-AuditTab
                Troubleshooting = Initialize-TroubleshootingTab
                Groups = Initialize-GroupsTab  
                SharedMailbox = Initialize-SharedMailboxTab
                Contacts = Initialize-ContactsTab
                Resources = Initialize-ResourcesTab
                Reports = Initialize-ReportsTab
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