# --------------------------------------------------------------
# Initialisiere Debugging und Logging für das Script
# --------------------------------------------------------------
$script:debugMode = $false
$script:logFilePath = Join-Path -Path "$PSScriptRoot\Logs" -ChildPath "ExchangeTool.log"
$script:logFileWritable = $true

# Assembly für WPF-Komponenten laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Definiere Farben für GUI
$script:connectedBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Green)
$script:disconnectedBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Red)
$script:isConnected = $false

# -------------------------------------------------
# Funktion zur Überprüfung der Ausführungsrichtlinie
# -------------------------------------------------
function Check-ExecutionPolicy {
    [CmdletBinding()]
    param()

    try {
        # Aktuelle Richtlinie für den Benutzer ermitteln
        $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser -ErrorAction SilentlyContinue
        # Wenn für CurrentUser 'Undefined', prüfe LocalMachine (effektive Richtlinie)
        if ($currentPolicy -eq 'Undefined' -or -not $currentPolicy) {
             $currentPolicy = Get-ExecutionPolicy -Scope LocalMachine -ErrorAction SilentlyContinue
             # Wenn immer noch undefiniert, prüfe den Prozess (könnte durch GPO gesetzt sein)
             if ($currentPolicy -eq 'Undefined' -or -not $currentPolicy) {
                 $currentPolicy = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue
             }
        }
        # Fallback, falls immer noch nichts gefunden wurde (unwahrscheinlich)
        if (-not $currentPolicy) {
            $currentPolicy = Get-ExecutionPolicy # Gesamt effektive Richtlinie
        }


        Write-LogEntry "Aktuell ermittelte effektive Ausführungsrichtlinie: $currentPolicy"

        # Richtlinien, die die Ausführung blockieren könnten
        $restrictedPolicies = @('Restricted', 'AllSigned')

        if ($currentPolicy -in $restrictedPolicies) {
            Write-Host "Warnung: Die aktuelle Ausführungsrichtlinie '$currentPolicy' könnte die Ausführung dieses Skripts verhindern." -ForegroundColor Yellow
            Write-LogEntry "Warnung: Restriktive Ausführungsrichtlinie '$currentPolicy' erkannt."

            $message = "Die aktuelle PowerShell Ausführungsrichtlinie ist auf '$currentPolicy' gesetzt. Dies könnte die Ausführung von easyEXO verhindern.`n`nMöchten Sie die Richtlinie für diesen Prozess temporär auf 'Bypass' setzen, um die Ausführung zu ermöglichen?`n`n(Dies ändert die Richtlinie nur für diese Sitzung und nicht dauerhaft.)"
            $title = "Ausführungsrichtlinie anpassen?"

            # Lade die Assembly für MessageBox, falls noch nicht geschehen (sollte durch Add-Type oben geschehen, aber sicher ist sicher)
            try {
                 Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
            } catch {
                 Write-Host "Fehler beim Laden der WPF-Assembly für die Dialogbox. Fortfahren nicht möglich." -ForegroundColor Red
                 Write-LogEntry "Fehler: Konnte PresentationFramework nicht laden für MessageBox."
                 return $false # Kritischer Fehler
            }


            $result = [System.Windows.MessageBox]::Show(
                $message,
                $title,
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning # Geändert zu Warning, da es eine potenzielle Sicherheitsänderung ist
            )

            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-LogEntry "Benutzer hat zugestimmt, die Ausführungsrichtlinie temporär zu ändern."
                try {
                    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
                    Write-Host "Ausführungsrichtlinie für diesen Prozess erfolgreich auf 'Bypass' gesetzt." -ForegroundColor Green
                    Write-LogEntry "Ausführungsrichtlinie für Prozess erfolgreich auf 'Bypass' gesetzt."
                    return $true
                }
                catch {
                    $errMsg = $_.Exception.Message
                    Write-Host "Fehler beim Setzen der Ausführungsrichtlinie auf 'Bypass': $errMsg" -ForegroundColor Red
                    Write-LogEntry "Fehler: Konnte Ausführungsrichtlinie nicht auf 'Bypass' setzen. Fehler: $errMsg"
                    [System.Windows.MessageBox]::Show(
                        "Die Ausführungsrichtlinie konnte nicht automatisch angepasst werden.`nFehler: $errMsg`n`nMöglicherweise fehlen Ihnen die nötigen Berechtigungen. Das Skript wird möglicherweise nicht korrekt ausgeführt.",
                        "Fehler beim Anpassen der Richtlinie",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Error
                    )
                    return $false # Richtlinie konnte nicht gesetzt werden
                }
            }
            else {
                Write-Host "Die Ausführungsrichtlinie wurde nicht geändert. Das Skript wird möglicherweise nicht ausgeführt." -ForegroundColor Yellow
                Write-LogEntry "Benutzer hat abgelehnt, die Ausführungsrichtlinie zu ändern."
                 [System.Windows.MessageBox]::Show(
                    "Die Ausführungsrichtlinie wurde nicht geändert. Wenn das Skript nicht startet, müssen Sie die Richtlinie manuell anpassen (z.B. mit 'Set-ExecutionPolicy RemoteSigned -Scope CurrentUser') oder das Skript als Administrator ausführen.",
                    "Hinweis zur Ausführungsrichtlinie",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                )
                return $false # Benutzer hat abgelehnt
            }
        }
        else {
            Write-Host "Die aktuelle Ausführungsrichtlinie '$currentPolicy' erlaubt die Skriptausführung." -ForegroundColor Green
            Write-LogEntry "Ausführungsrichtlinie '$currentPolicy' ist ausreichend."
            return $true # Richtlinie ist in Ordnung
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Host "Fehler beim Überprüfen oder Anpassen der Ausführungsrichtlinie: $errMsg" -ForegroundColor Red
        Write-LogEntry "Fehler: Kritischer Fehler in Check-ExecutionPolicy. Fehler: $errMsg"
         [System.Windows.MessageBox]::Show(
            "Ein unerwarteter Fehler ist bei der Überprüfung der Ausführungsrichtlinie aufgetreten:`n$errMsg",
            "Fehler bei Richtlinienprüfung",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return $false # Fehler bei der Überprüfung
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
        [ValidateSet("Debug", "Info", "Warning", "Error", "Success")] # "Debug" hinzugefügt
        [string]$Type = "Info"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Farb- und Präfixzuordnung basierend auf dem Typ
    $colorMap = @{
        "Debug"   = "Gray"
        "Info"    = "Cyan"
        "Warning" = "Yellow"
        "Error"   = "Red"
        "Success" = "Green"
    }
    $prefixMap = @{
        "Debug"   = "DEBUG"
        "Info"    = "INFO"
        "Warning" = "WARN"
        "Error"   = "ERR!"
        "Success" = "OK!"
    }

    # Standardwerte, falls Typ ungültig (sollte durch ValidateSet verhindert werden)
    $color = $colorMap[$Type] # Direktzugriff, da ValidateSet Gültigkeit sicherstellt
    $prefix = $prefixMap[$Type]

    # Konsolenausgabe basierend auf Debug-Modus
    if ($script:debugMode) {
        # Debug-Modus AN: Alle Meldungen ausgeben
        Write-Host "[$timestamp] [$prefix] $Message" -ForegroundColor $color
    }
    else {
        # Debug-Modus AUS: Nur Fehlermeldungen ausgeben
        if ($Type -eq "Error") {
            Write-Host "[$timestamp] [$prefix] $Message" -ForegroundColor $color
        }
    }

    # Logging (immer durchführen, unabhängig vom Debug-Modus für die Konsole)
    try {
        # Nur loggen, wenn Log-Action verfügbar ist und Logdatei beschreibbar ist
        if ($script:logFileWritable) {
             Log-Action "$prefix - $Message"
        }
    }
    catch {
        # Fallback für Fehler in der Debug-Funktion selbst (z.B. Log-Action schlägt fehl oder ist noch nicht definiert)
        try {
            $errorMsg = $_.Exception.Message -replace '[^\x20-\x7E]', '?' # Sicherstellen, dass die Fehlermeldung druckbar ist
            $timestampFallback = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            # Versuche, den Log-Pfad aus der Konfiguration zu holen, sonst Fallback
            $logFolder = "$PSScriptRoot\Logs" # Standard-Fallback
            try {
                 if ($script:config -and $script:config.ContainsKey("Paths") -and $script:config["Paths"].ContainsKey("LogPath")) {
                     $configuredLogPath = $script:config["Paths"]["LogPath"]
                     if (-not [string]::IsNullOrWhiteSpace($configuredLogPath)) {
                        # Verwende den Ordner aus der Konfiguration
                        $logFolder = Split-Path -Path $configuredLogPath -Parent -ErrorAction SilentlyContinue
                        if (-not $logFolder -or -not (Test-Path $logFolder -PathType Container)) {
                            $logFolder = "$PSScriptRoot\Logs" # Fallback wenn Split-Path fehlschlägt oder Pfad ungültig
                        }
                     }
                 }
            } catch {
                # Fehler beim Zugriff auf Konfig ignorieren, Fallback verwenden
            }

            if (-not (Test-Path $logFolder -PathType Container)) {
                 # Versuche Ordner zu erstellen, wenn er nicht existiert
                New-Item -ItemType Directory -Path $logFolder -Force -ErrorAction SilentlyContinue | Out-Null
            }
            # Nur schreiben, wenn der Ordner existiert oder erstellt werden konnte
            if (Test-Path $logFolder -PathType Container) {
                $fallbackLogFile = Join-Path $logFolder "debug_fallback.log"
                # Verwende -ErrorAction SilentlyContinue auch hier, um Endlosschleifen bei Berechtigungsproblemen zu vermeiden
                Add-Content -Path $fallbackLogFile -Value "[$timestampFallback] Fehler in Write-DebugMessage (Log-Action?): $errorMsg" -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch {
            # Absoluter Fallback - ignoriere Fehler um Programmablauf nicht zu stören
        }
    }
}

# Globale Variable, um den Status der Logdatei zu verfolgen (muss außerhalb initialisiert werden, z.B. $script:logFileWritable = $true)
# Annahme: $script:logFileWritable existiert und ist initial $true

function Log-Action {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message
    )

    # Wenn die Logdatei bereits als nicht beschreibbar markiert wurde, nichts tun (Ausgabe erfolgt über Write-DebugMessage)
    if (-not $script:logFileWritable) {
        return
    }

    try {
        # Sicherstellen, dass nur druckbare ASCII-Zeichen verwendet werden (oder UTF8-kompatible Zeichen)
        # Beibehalten von Umlauten und anderen gängigen Zeichen, Entfernen von Steuerzeichen
        $sanitizedMessage = $Message -replace '[\p{C}]', '?' # Entfernt Steuerzeichen

        # Zeitstempel erzeugen
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Logverzeichnis erstellen, falls nicht vorhanden
        $logFolder = Split-Path -Path $script:logFilePath -Parent
        if (-not (Test-Path $logFolder)) {
            New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
            # Keine Debug-Nachricht hier, da dies zu Rekursion führen könnte, wenn das Loggen selbst fehlschlägt
        }

        # Log-Eintrag schreiben
        Add-Content -Path $script:logFilePath -Value "[$timestamp] $sanitizedMessage" -Encoding UTF8 -ErrorAction Stop # ErrorAction Stop, um den Catch-Block auszulösen

        # Bei zu langer Logdatei (>10 MB) rotieren
        $logFile = Get-Item -Path $script:logFilePath -ErrorAction SilentlyContinue # Fehler hier nicht abfangen, Add-Content prüft Schreibbarkeit
        if ($logFile -and $logFile.Length -gt 10MB) {
            $backupLogPath = "$($script:logFilePath)_$(Get-Date -Format 'yyyyMMdd_HHmmss').bak"
            Move-Item -Path $script:logFilePath -Destination $backupLogPath -Force -ErrorAction Stop # Fehler beim Verschieben auch abfangen
            # Keine Debug-Nachricht hier aus demselben Grund wie oben
        }
    }
    catch {
        # Fehler beim Schreiben oder Rotieren der Logdatei
        if ($script:logFileWritable) {
            # Nur beim ersten Fehler die Meldung ausgeben und Flag setzen
            $script:logFileWritable = $false
            $errorMsg = $_.Exception.Message -replace '[\p{C}]', '?'
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Write-Host "[$timestamp] [WARN] Fehler beim Schreiben/Rotieren der Logdatei '$script:logFilePath'. Fehler: $errorMsg" -ForegroundColor Yellow
            Write-Host "[$timestamp] [WARN] Zukünftige Log-Einträge werden nur noch in der Konsole angezeigt." -ForegroundColor Yellow
        }
        # Kein Fallback-Log mehr, da die Anforderung ist, nur noch in der Konsole zu loggen.
        # Die ursprüngliche Nachricht wird bereits über Write-DebugMessage -> Write-Host ausgegeben.
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
    # Hilfsfunktion zum rekursiven Finden eines Elements im Visual Tree anhand des Namens
    function Find-VisualChildByName {
        param(
            [Parameter(Mandatory=$true)]
            [System.Windows.DependencyObject]$Parent,

            [Parameter(Mandatory=$true)]
            [string]$Name
        )

        if ($Parent -eq $null) { return $null }

        $childrenCount = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent)
        for ($i = 0; $i -lt $childrenCount; $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
            $childName = $child.GetValue([System.Windows.FrameworkElement]::NameProperty)

            if ($childName -eq $Name) {
                # Gefunden!
                return $child
            }
            else {
                # Rekursiv weitersuchen
                $found = Find-VisualChildByName -Parent $child -Name $Name
                if ($found -ne $null) {
                    return $found
                }
            }
        }
        # Nicht in diesem Zweig gefunden
        return $null
    }
    #region Helper Functions Specific to Tabs
# Hilfsfunktion zur formatierten Fehlerausgabe
function Get-FormattedError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord, # Nimmt das $_ aus dem Catch-Block entgegen

        [Parameter(Mandatory=$false)]
        [string]$DefaultText = "Ein Fehler ist aufgetreten."
    )

    # Versucht, die spezifischste Fehlermeldung zu extrahieren
    if ($null -ne $ErrorRecord.Exception) {
        # Beinhaltet oft die tiefere Ursache
        return $ErrorRecord.Exception.Message.Trim()
    }
    elseif ($null -ne $ErrorRecord.InvocationInfo -and $ErrorRecord.CategoryInfo.Reason) {
         # Nimmt den Grund, falls verfügbar
         return "$($ErrorRecord.CategoryInfo.Reason)"
    }
    elseif ($ErrorRecord.ToString()) {
        # Nimmt die Standard-String-Repräsentation des Fehlers
        return $ErrorRecord.ToString()
    }
    else {
        # Fallback
        return $DefaultText
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
# NEUE Funktion zum Laden der akzeptierten Domains
function Load-AcceptedDomains {
    [CmdletBinding()]
    param()

    Write-DebugMessage "Versuche akzeptierte Domains zu laden..." -Type Info

    if (-not $script:isConnected) {
        Write-DebugMessage "Laden der Domains übersprungen - keine Verbindung." -Type Warning
        return
    }

    # --- NEU: Hole Referenz zur ComboBox direkt hier ---
    $cmbSharedDomain = Get-XamlElement -ElementName "cmbSharedMailboxDomain"
    # Füge hier weitere Get-XamlElement Aufrufe für andere Domain-ComboBoxen hinzu, falls benötigt
    # $cmbGroupDomain = Get-XamlElement -ElementName "cmbGroupDomain"

    # Stelle sicher, dass die relevante ComboBox existiert
    if ($null -eq $cmbSharedDomain) {
        Write-DebugMessage "Domain-ComboBox 'cmbSharedMailboxDomain' nicht gefunden beim Laden der Domains." -Type Warning
        # Optional: Hier weitere Prüfungen einfügen
        return # Beenden, wenn das Hauptelement fehlt
    }

    try {
        $domains = Get-AcceptedDomain | Select-Object -ExpandProperty DomainName | Sort-Object
        Write-DebugMessage "$(@($domains).Count) akzeptierte Domains gefunden." -Type Info

        # Shared Mailbox Domain ComboBox füllen
        # Verwende die lokale Variable $cmbSharedDomain statt $script:cmbSharedMailboxDomain
        $cmbSharedDomain.Dispatcher.InvokeAsync({
            $cmbSharedDomain.Items.Clear()
            foreach ($domain in $domains) {
                [void]$cmbSharedDomain.Items.Add($domain)
            }
            if ($cmbSharedDomain.Items.Count -gt 0) {
                $cmbSharedDomain.SelectedIndex = 0
                Write-DebugMessage "Domain-ComboBox für Shared Mailboxes gefüllt." -Type Success
            }
        }) | Out-Null


        # Hier Code zum Füllen anderer Domain-ComboBoxen hinzufügen (verwende lokale Variablen)
        # if ($null -ne $cmbGroupDomain) { ... $cmbGroupDomain.Dispatcher.InvokeAsync ... }

    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Laden der akzeptierten Domains: $errorMsg" -Type Error
        # Zeige die Fehlermeldung im UI-Thread an, falls möglich
        Show-MessageBox -Message "Fehler beim Laden der Domains: $errorMsg" -Title "Fehler" -Type Error
    }
}

    #endregion
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
            # Lade die akzeptierten Domains nach erfolgreicher Verbindung
            Load-AcceptedDomains
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

function Show-HelpDialog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Topic # Der Bezeichner für das Hilfethema (z.B. "Connect", "Reports", "MailboxBasics")
    )

    # ---- NEUE PRÜFUNG ----
    if ([string]::IsNullOrWhiteSpace($Topic)) {
        Write-DebugMessage "Show-HelpDialog wurde mit einem leeren Topic aufgerufen. Breche ab." -Type Error
        # Optional: Benutzer eine Meldung anzeigen, aber das Hauptproblem liegt woanders.
        # [System.Windows.MessageBox]::Show("Ein interner Fehler ist aufgetreten: Hilfe kann nicht ohne Thema angezeigt werden.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return # Funktion sofort verlassen
    }
    # ---- ENDE NEUE PRÜFUNG ----

    Write-DebugMessage "Zeige Hilfe für Thema '$Topic'" -Type Info

    # Standard-Hilfetext, falls kein spezifisches Thema gefunden wird
    $helpTitle = "Hilfe: $Topic"
    # Standardnachricht mit korrekter Behandlung von Anführungszeichen und Zeilenumbrüchen
    $helpMessage = "Für das Thema ""$Topic"" ist noch keine spezifische Hilfe verfügbar.`n`nAllgemeine Informationen: Dieses Tool dient zur Verwaltung von Exchange Online über eine grafische Oberfläche."

    # Spezifische Hilfetexte pro Thema definieren
    # Hier können Sie die Hilfetexte für jeden Tab oder jede wichtige Funktion hinterlegen
    switch ($Topic) {
        "Connect" {
            $helpTitle = "Hilfe: Verbindung"
            $helpMessage = @"
Verbindung zu Exchange Online herstellen:
1.  Klicken Sie auf den Button 'Verbinden'.
2.  Folgen Sie den Anweisungen im Authentifizierungsfenster (Anmeldung mit Ihrem Microsoft 365 Konto).
3.  Der Statusbereich unten im Fenster zeigt an, ob die Verbindung erfolgreich war ('Verbunden') oder ob Fehler aufgetreten sind.

Voraussetzungen:
-   Stellen Sie sicher, dass die notwendigen PowerShell-Module (ExchangeOnlineManagement) installiert sind. Nutzen Sie hierfür den Button 'Module prüfen/installieren'.
-   Sie benötigen die entsprechenden Berechtigungen in Microsoft 365, um eine Verbindung herzustellen und Aktionen durchzuführen.
"@
        }
        "MailboxBasics" {
            $helpTitle = "Hilfe: Postfach Grundlagen"
            $helpMessage = @"
Postfach Grundlagen verwalten:
-   **Suchen:** Geben Sie einen Teil des Namens oder die E-Mail-Adresse in das Suchfeld ein und klicken Sie auf 'Suchen', um Postfächer zu finden.
-   **Details anzeigen:** Wählen Sie ein Postfach aus der Liste aus. Die Details werden automatisch in den entsprechenden Feldern angezeigt.
-   **Erstellen:** Füllen Sie die Felder im Bereich 'Neues Postfach erstellen' aus (Name, E-Mail, Passwort etc.) und klicken Sie auf 'Erstellen'.
-   **Löschen:** Suchen und wählen Sie das zu löschende Postfach aus. Klicken Sie auf 'Postfach löschen' und bestätigen Sie die Sicherheitsabfrage. Achtung: Gelöschte Postfächer können oft nur für eine begrenzte Zeit wiederhergestellt werden.
-   **Eigenschaften ändern:** Suchen und wählen Sie ein Postfach aus. Ändern Sie die gewünschten Werte in den Feldern und klicken Sie auf 'Änderungen speichern'.
"@
        }
        "Reports" {
            $helpTitle = "Hilfe: Berichte & Export"
            $helpMessage = @"
Berichte generieren und exportieren:
1.  **Kategorie auswählen:** Wählen Sie links eine Berichtskategorie (z.B. 'Postfachberichte', 'Berechtigungsberichte').
2.  **Berichtstyp auswählen:** Wählen Sie aus der Dropdown-Liste den spezifischen Bericht aus, den Sie generieren möchten. Die verfügbaren Berichte ändern sich je nach gewählter Kategorie.
3.  **(Optional) Datum festlegen:** Für einige Berichte können Sie einen Datumsbereich festlegen. Wählen Sie Start- und Enddatum über die Kalendersteuerung aus.
4.  **Bericht generieren:** Klicken Sie auf 'Bericht generieren'. Die Ergebnisse werden in der Tabelle im unteren Bereich angezeigt. Dies kann je nach Datenmenge einige Zeit dauern.
5.  **Exportieren:** Nachdem der Bericht generiert wurde, klicken Sie auf 'Bericht exportieren (CSV)', um die angezeigten Daten in eine CSV-Datei zu speichern. Sie werden aufgefordert, einen Speicherort und Dateinamen auszuwählen. Die CSV-Datei ist für die Weiterverarbeitung in Excel optimiert (UTF8, Semikolon-Trenner).
"@
        }
        "Troubleshooting" {
            $helpTitle = "Hilfe: Fehlerbehebung"
            $helpMessage = @"
Fehlerbehebung (Troubleshooting):
Dieser Bereich bietet Zugriff auf integrierte Diagnosewerkzeuge von Microsoft 365.
1.  **Diagnose auswählen:** Wählen Sie eine verfügbare Diagnose aus der Liste (z.B. E-Mail-Zustellung prüfen, Postfachzugriff analysieren).
2.  **Parameter eingeben:** Füllen Sie die erforderlichen Felder für die ausgewählte Diagnose aus (z.B. betroffene E-Mail-Adresse, Absender, Empfänger). Die benötigten Felder werden je nach Diagnose angepasst.
3.  **Diagnose ausführen:** Klicken Sie auf 'Diagnose ausführen'. Das Tool führt die entsprechende Diagnose im Hintergrund über PowerShell aus.
4.  **Ergebnis anzeigen:** Das Ergebnis der Diagnose wird im großen Textfeld unten angezeigt. Dies kann technische Details oder Lösungsvorschläge enthalten.
5.  **Admin Center:** Der Button 'Zum Admin Center' öffnet das relevante Microsoft 365 Admin Center in Ihrem Webbrowser für weiterführende Analysen oder Konfigurationen.
"@
        }
        # Fügen Sie hier weitere 'case'-Blöcke für andere $Topic-Werte hinzu,
        # z.B. für "Calendar", "Groups", "Contacts", "Permissions", "Audit"
        # Beispiel:
         "Calendar" {
            $helpTitle = "Hilfe: Kalenderberechtigungen"
            $helpMessage = "Hier steht der Hilfetext für den Kalender-Tab..."
         }
         "Mailbox" { # Hinzugefügt, da es einen helpLinkMailbox gibt
            $helpTitle = "Hilfe: Postfach"
            $helpMessage = "Verwalten Sie hier Postfächer, Berechtigungen und Einstellungen."
         }
         "Audit" { # Hinzugefügt, um den default-Fall zu vermeiden
             $helpTitle = "Hilfe: Audit & Protokollierung"
             $helpMessage = "Führen Sie hier Audit-Suchen durch oder konfigurieren Sie die Protokollierung."
         }
        # --- NEUE CASE BLÖCKE START ---
         "OrganizationSettings" {
             $helpTitle = "Hilfe: Organisationseinstellungen"
             $helpMessage = "Hier können Sie globale Einstellungen für Ihre Exchange Online Organisation konfigurieren."
         }
         "Groups" {
             $helpTitle = "Hilfe: Gruppen & Verteiler"
             $helpMessage = "Verwalten Sie hier Verteilerlisten, Sicherheitsgruppen und Microsoft 365-Gruppen."
         }
         "SharedMailbox" {
             $helpTitle = "Hilfe: Shared Mailboxes"
             $helpMessage = "Erstellen und verwalten Sie hier Shared Mailboxes und deren Berechtigungen."
         }
         "Resources" {
             $helpTitle = "Hilfe: Ressourcen"
             $helpMessage = "Verwalten Sie hier Raum- und Ausstattungsressourcen für Buchungen."
         }
         "Contacts" {
             $helpTitle = "Hilfe: Kontakte"
             $helpMessage = "Verwalten Sie hier externe Mailkontakte und Mailbenutzer."
         }
         # --- NEUE CASE BLÖCKE ENDE ---
        default {
            # Die oben definierte Standardnachricht wird verwendet, wenn kein spezifischer Case passt.
            Write-DebugMessage "Kein spezifischer Hilfetext für '$Topic' gefunden. Zeige Standardtext." -Type Warning
        }
    }

    # Zeige das Hilfe-Fenster mit dem ermittelten Text an
    try {
        # Korrigierte MessageBox-Anzeige mit Variablen und korrekter Syntax
        [System.Windows.MessageBox]::Show($helpMessage,
                                         $helpTitle,
                                         [System.Windows.MessageBoxButton]::OK,
                                         [System.Windows.MessageBoxImage]::Information)
    } catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Anzeigen des Hilfe-Dialogs für '$Topic': $errorMsg" -Type Error
        # Fallback-Benachrichtigung, falls MessageBox fehlschlägt
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler beim Anzeigen der Hilfe für '$Topic'."
        } else {
             Write-Error "Fehler beim Anzeigen des Hilfe-Dialogs für '$Topic': $errorMsg"
        }
    }
}

# Funktion zum Initialisieren des Kalender-Tabs
function Initialize-CalendarTab {
    [CmdletBinding()]
    param()

    Write-DebugMessage "Initialisiere Calendar Tab..." -Type Info

    try {
        # UI-Elemente referenzieren
        $script:txtCalendarSource = Get-XamlElement -ElementName "txtCalendarSource"
        $script:txtCalendarTarget = Get-XamlElement -ElementName "txtCalendarTarget"
        $script:cmbCalendarPermission = Get-XamlElement -ElementName "cmbCalendarPermission"
        $script:btnAddCalendarPermission = Get-XamlElement -ElementName "btnAddCalendarPermission"
        $script:btnRemoveCalendarPermission = Get-XamlElement -ElementName "btnRemoveCalendarPermission"
        $script:txtCalendarMailboxUser = Get-XamlElement -ElementName "txtCalendarMailboxUser"
        $script:btnShowCalendarPermissions = Get-XamlElement -ElementName "btnShowCalendarPermissions"
        $script:cmbDefaultPermission = Get-XamlElement -ElementName "cmbDefaultPermission"
        $script:btnSetDefaultPermission = Get-XamlElement -ElementName "btnSetDefaultPermission"
        $script:btnSetAnonymousPermission = Get-XamlElement -ElementName "btnSetAnonymousPermission"
        $script:btnSetAllCalPermission = Get-XamlElement -ElementName "btnSetAllCalPermission"
        $script:lstCalendarPermissions = Get-XamlElement -ElementName "lstCalendarPermissions"
        $script:btnExportCalendarPermissions = Get-XamlElement -ElementName "btnExportCalendarPermissions"
        $script:helpLinkCalendar = Get-XamlElement -ElementName "helpLinkCalendar"

        # Berechtigungsstufen für ComboBoxen laden (Beispiel)
        $permissionLevels = @("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author", "NoneditingAuthor", "Reviewer", "Contributor", "AvailabilityOnly", "LimitedDetails", "None")
        if ($null -ne $script:cmbCalendarPermission) {
            $script:cmbCalendarPermission.ItemsSource = $permissionLevels
            if ($script:cmbCalendarPermission.Items.Count -gt 0) { $script:cmbCalendarPermission.SelectedIndex = 7 } # Default Contributor? oder Reviewer?
        }
        if ($null -ne $script:cmbDefaultPermission) {
             $script:cmbDefaultPermission.ItemsSource = $permissionLevels
             if ($script:cmbDefaultPermission.Items.Count -gt 0) { $script:cmbDefaultPermission.SelectedIndex = 8 } # Default AvailabilityOnly?
        }


        # --- Event-Handler registrieren ---

        # Hinzufügen
        Register-EventHandler -Control $script:btnAddCalendarPermission -Handler {
            $source = $script:txtCalendarSource.Text
            $target = $script:txtCalendarTarget.Text
            $level = $script:cmbCalendarPermission.SelectedItem
            if (-not [string]::IsNullOrWhiteSpace($source) -and -not [string]::IsNullOrWhiteSpace($target) -and $null -ne $level) {
                Add-CalendarPermission -SourceMailbox $source -User $target -AccessRights $level
            } else {
                Show-MessageBox -Message "Bitte Quell-Postfach, Ziel-Benutzer und Berechtigungsstufe angeben." -Title "Eingabe fehlt" -Type Warning
            }
        } -ControlName "btnAddCalendarPermission"

        # Entfernen
        Register-EventHandler -Control $script:btnRemoveCalendarPermission -Handler {
            $source = $script:txtCalendarSource.Text
            $target = $script:txtCalendarTarget.Text
             if (-not [string]::IsNullOrWhiteSpace($source) -and -not [string]::IsNullOrWhiteSpace($target)) {
                Remove-CalendarPermission -SourceMailbox $source -User $target
            } else {
                Show-MessageBox -Message "Bitte Quell-Postfach und Ziel-Benutzer angeben." -Title "Eingabe fehlt" -Type Warning
            }
        } -ControlName "btnRemoveCalendarPermission"

        # Anzeigen
        Register-EventHandler -Control $script:btnShowCalendarPermissions -Handler {
            $mailbox = $script:txtCalendarMailboxUser.Text
            if (-not [string]::IsNullOrWhiteSpace($mailbox)) {
                Show-CalendarPermissions -MailboxUser $mailbox
            } else {
                 Show-MessageBox -Message "Bitte Postfach zum Anzeigen der Berechtigungen angeben." -Title "Eingabe fehlt" -Type Warning
            }
        } -ControlName "btnShowCalendarPermissions"

         # Standard setzen (Default)
        Register-EventHandler -Control $script:btnSetDefaultPermission -Handler {
            $mailbox = $script:txtCalendarMailboxUser.Text
            $level = $script:cmbDefaultPermission.SelectedItem
             if (-not [string]::IsNullOrWhiteSpace($mailbox) -and $null -ne $level) {
                Set-DefaultCalendarPermission -MailboxUser $mailbox -AccessRights $level
            } else {
                 Show-MessageBox -Message "Bitte Postfach und Berechtigungsstufe für 'Default' angeben." -Title "Eingabe fehlt" -Type Warning
            }
        } -ControlName "btnSetDefaultPermission"

         # Standard setzen (Anonymous)
        Register-EventHandler -Control $script:btnSetAnonymousPermission -Handler {
             $mailbox = $script:txtCalendarMailboxUser.Text
             $level = $script:cmbDefaultPermission.SelectedItem # Nimmt die gleiche ComboBox wie Default
             if (-not [string]::IsNullOrWhiteSpace($mailbox) -and $null -ne $level) {
                 Set-AnonymousCalendarPermission -MailboxUser $mailbox -AccessRights $level
             } else {
                 Show-MessageBox -Message "Bitte Postfach und Berechtigungsstufe für 'Anonymous' angeben." -Title "Eingabe fehlt" -Type Warning
             }
        } -ControlName "btnSetAnonymousPermission"

        # Standard für Alle setzen
        Register-EventHandler -Control $script:btnSetAllCalPermission -Handler {
             $level = $script:cmbDefaultPermission.SelectedItem
             if ($null -ne $level) {
                 $confirmResult = Show-MessageBox -Message "Möchten Sie wirklich die Standard- und Anonym-Berechtigung für ALLE Postfächer auf '$level' setzen? Dies kann dauern!" -Title "Bestätigung erforderlich" -Type Question
                 if ($confirmResult -eq 'Yes') {
                     Set-DefaultCalendarPermissionForAll -AccessRights $level
                     Set-AnonymousCalendarPermissionForAll -AccessRights $level
                 }
             } else {
                  Show-MessageBox -Message "Bitte wählen Sie eine Berechtigungsstufe für 'Alle Postfächer'." -Title "Eingabe fehlt" -Type Warning
             }
        } -ControlName "btnSetAllCalPermission"

        # Exportieren
        Register-EventHandler -Control $script:btnExportCalendarPermissions -Handler {
            $itemsToExport = $script:lstCalendarPermissions.ItemsSource -as [System.Collections.IList]
            if ($null -ne $itemsToExport -and $itemsToExport.Count -gt 0) {
                # Annahme: Es gibt eine generische Export-Funktion
                Export-DataGridContent -DataGridItemsSource $itemsToExport -DefaultFileName "Kalenderberechtigungen"
            } else {
                Show-MessageBox -Message "Keine Daten zum Exportieren vorhanden. Bitte zuerst Berechtigungen anzeigen." -Title "Export nicht möglich" -Type Info
            }
        } -ControlName "btnExportCalendarPermissions"

        # Hilfe-Link
        if ($null -ne $script:helpLinkCalendar) {
            $script:helpLinkCalendar.Add_MouseLeftButtonDown({ Show-HelpDialog -Topic "Calendar" })
            $script:helpLinkCalendar.Add_MouseEnter({ $this.Cursor = [System.Windows.Input.Cursors]::Hand; $this.TextDecorations = [System.Windows.TextDecorations]::Underline })
            $script:helpLinkCalendar.Add_MouseLeave({ $this.Cursor = [System.Windows.Input.Cursors]::Arrow; $this.TextDecorations = $null })
        }

        Write-DebugMessage "Calendar Tab erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message + "`n" + $_.ScriptStackTrace
        Write-DebugMessage "Fehler beim Initialisieren des Calendar Tabs: $errorMsg" -Type "Error"
        Show-MessageBox -Message "Schwerwiegender Fehler beim Initialisieren des Calendar Tabs: $errorMsg" -Title "Initialisierungsfehler" -Type Error
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
    $script:btnNavRegion          = Get-XamlElement -ElementName "btnNavRegion"
    $script:btnNavAudit           = Get-XamlElement -ElementName "btnNavAudit"
    $script:btnNavReports         = Get-XamlElement -ElementName "btnNavReports"
    $script:btnNavTroubleshooting = Get-XamlElement -ElementName "btnNavTroubleshooting"
    $script:btnInfo               = Get-XamlElement -ElementName "btnInfo"
    $script:btnSettings           = Get-XamlElement -ElementName "btnSettings"
    $script:btnClose              = Get-XamlElement -ElementName "btnClose" -Required
    
    # Referenzierung weiterer wichtiger UI-Elemente
    $script:btnCheckPrerequisites   = Get-XamlElement -ElementName "btnCheckPrerequisites"
    
# -------------------------------------------------
# Event-Handler für Navigationsbuttons
# -------------------------------------------------
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

# Event-Handler für Verbinden-Button
Register-EventHandler -Control $script:btnConnect -Handler {
    try {
        Write-DebugMessage "Button 'Mit Exchange verbinden' geklickt." -Type Info
        # Funktion zum Verbinden aufrufen (Name ggf. anpassen)
        Connect-ExchangeOnlineSession
    } catch {
        Write-DebugMessage "Fehler beim Klick auf Verbinden-Button: $($_.Exception.Message)" -Type Error
        Show-MessageBox -Message "Fehler beim Verbindungsversuch: $($_.Exception.Message)" -Title "Verbindungsfehler" -Type Error
    }
} -ControlName "btnConnect"

# Funktion zum Setzen des aktiven Tabs und Aktualisieren des Status
function Set-ActiveTab {
    param(
        [int]$Index,
        [string]$TabName,
        [string]$StatusText
    )
    try {
        if ($null -ne $script:tabContent) {
            # Sicherstellen, dass der Index gültig ist
            if ($Index -ge 0 -and $Index -lt $script:tabContent.Items.Count) {
                $script:tabContent.SelectedIndex = $Index
                Write-DebugMessage "Wechsle zu Tab: $TabName (Index $Index)" -Type Info
                if ($null -ne $script:txtStatus) {
                    $script:txtStatus.Text = $StatusText
                }
            } else {
                 Write-DebugMessage "Fehler: Ungültiger Tab-Index $Index für Tab '$TabName'." -Type Error
            }
        } else {
            Write-DebugMessage "Fehler: Haupt-TabControl (\$script:tabContent) nicht gefunden." -Type Error
        }
    } catch {
        Write-DebugMessage "Fehler beim Wechseln zum Tab '$TabName': $($_.Exception.Message)" -Type Error
    }
}

# Reihenfolge der TabItems im XAML muss mit den Indizes übereinstimmen:
# 0: tabCalendar
# 1: tabMailbox
# 2: tabMailboxAudit
# 3: tabTroubleshooting
# 4: tabEXOSettings
# 5: tabGroups
# 6: tabSharedMailbox
# 7: tabResources
# 8: tabContacts
# 9: tabReports

# Registriere Handler für jeden Navigationsbutton
Register-EventHandler -Control $script:btnNavCalendar -Handler { Set-ActiveTab -Index 0 -TabName "Kalender" -StatusText "Kalenderberechtigungen ausgewählt" } -ControlName "btnNavCalendar"
Register-EventHandler -Control $script:btnNavMailbox -Handler { Set-ActiveTab -Index 1 -TabName "Postfach" -StatusText "Postfachberechtigungen ausgewählt" } -ControlName "btnNavMailbox"
Register-EventHandler -Control $script:btnNavAudit -Handler { Set-ActiveTab -Index 2 -TabName "Audit" -StatusText "Audit & Information ausgewählt" } -ControlName "btnNavAudit"
Register-EventHandler -Control $script:btnNavTroubleshooting -Handler { Set-ActiveTab -Index 3 -TabName "Troubleshooting" -StatusText "Troubleshooting ausgewählt" } -ControlName "btnNavTroubleshooting"
Register-EventHandler -Control $script:btnNavEXOSettings -Handler { Set-ActiveTab -Index 4 -TabName "EXOSettings" -StatusText "Allgemeine EXO Settings ausgewählt" } -ControlName "btnNavEXOSettings"
Register-EventHandler -Control $script:btnNavGroups -Handler { Set-ActiveTab -Index 5 -TabName "Gruppen" -StatusText "Gruppen / Verteiler ausgewählt" } -ControlName "btnNavGroups"
Register-EventHandler -Control $script:btnNavSharedMailbox -Handler { Set-ActiveTab -Index 6 -TabName "SharedMailbox" -StatusText "Shared Mailboxes ausgewählt" } -ControlName "btnNavSharedMailbox"
Register-EventHandler -Control $script:btnNavResources -Handler { Set-ActiveTab -Index 7 -TabName "Ressourcen" -StatusText "Ressourcen ausgewählt" } -ControlName "btnNavResources"
Register-EventHandler -Control $script:btnNavContacts -Handler { Set-ActiveTab -Index 8 -TabName "Kontakte" -StatusText "Kontakte ausgewählt" } -ControlName "btnNavContacts"
Register-EventHandler -Control $script:btnNavReports -Handler { Set-ActiveTab -Index 9 -TabName "Berichte" -StatusText "Berichte & Export ausgewählt" } -ControlName "btnNavReports"
Register-EventHandler -Control $script:btnNavRegion -Handler { Set-ActiveTab -Index 10 -TabName "Regionaleinstellungen" -StatusText "Regionaleinstellungen ausgewählt" } -ControlName "btnNavRegion" # NEU: Event Handler für Regionaleinstellungen
# --- Ende Event-Handler für Navigationsbuttons ---

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
    "chkActivityBasedAuthenticationTimeoutInterval", # Ist eine ComboBox!
    "chkActivityBasedAuthenticationTimeoutWithSingleSignOnEnabled", 
    "chkAppsForOfficeEnabled", 
    "chkAsyncSendEnabled", 
    "chkBookingsAddressEntryRestricted", 
    "chkBookingsAuthEnabled", 
    "chkBookingsCreationOfCustomQuestionsRestricted", 
    "chkBookingsExposureOfStaffDetailsRestricted", 
    "chkBookingsMembershipApprovalRequired", 
    "chkBookingsNamingPolicyEnabled", 
    "chkBookingsNamingPolicySuffix", # Name irreführend, steuert Präfix
    "chkBookingsNamingPolicySuffixEnabled", # Name irreführend, steuert Präfix-Aktivierung
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
    # "chkInformationBarrierMode", # Redundant -> cmbInformationBarrierMode
    "cmbInformationBarrierMode", 
    "chkImplicitSharingEnabled", 
    "chkOAuthUseBasicAuth", 
    "chkRefreshSessionEnabled", 
    "chkPerTenantSwitchToESTSEnabled", 
    # "chkEwsApplicationAccessPolicy", # Redundant -> cmbEwsAppAccessPolicy
    "cmbEwsAppAccessPolicy", 
    "chkEws", # Korrekt für EwsEnabled
    "chkEwsAllowList", 
    "chkEwsAllowEntourage", 
    "chkEwsAllowMacOutlook", 
    "chkEwsAllowOutlook", 
    "chkMacOutlook", # Korrekt für MAPIEnabled? Nein, eigener Parameter
    "chkOutlookMobile", # Korrekt für OutlookMobileEnabled? Ja
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
    # "chkOfficeFeatures", # Redundant -> cmbOfficeFeatures
    "cmbOfficeFeatures", 
    "chkMobileToFollowedFolders", 
    "chkDisablePlusAddressInRecipients", 
    # "chkDefaultAuthenticationPolicy", # Redundant -> txtDefaultAuthPolicy
    "txtDefaultAuthPolicy", 
    # "chkHierarchicalAddressBookRoot", # Redundant -> txtHierAddressBookRoot
    "txtHierAddressBookRoot",
    
    # Erweitert Tab
    "chkSIPEnabled", 
    "chkRemotePublicFolderBlobsEnabled", 
    "chkMapiHttpEnabled", 
    "chkPreferredInternetCodePageForShiftJis", 
    "txtPreferredInternetCodePageForShiftJis", 
    "chkVisibilityEnabled", 
    "chkOnlineMeetingsByDefaultEnabled", 
    "cmbSearchQueryLanguage", 
    "chkDirectReportsGroupAutoCreationEnabled", 
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

        # Finde das TabItem selbst
        $tabEXOSettings = Get-XamlElement -ElementName "tabEXOSettings"
        if ($null -eq $tabEXOSettings) {
            Write-DebugMessage "Fehler: TabItem 'tabEXOSettings' konnte nicht gefunden werden." -Type Error
            return $false # Oder wirf einen Fehler
        }
        $script:tabEXOSettings = $tabEXOSettings # Speichere es global, falls woanders benötigt
        Write-DebugMessage "TabItem 'tabEXOSettings' gefunden." -Type Info

        # Event-Handler für Help-Link (bleibt global, da er außerhalb des Tabs sein könnte oder einfacher so zu finden ist)
        # Versuche, das Element innerhalb des Tabs zu finden, falls es dort definiert ist
        $helpLinkEXOSettings = $tabEXOSettings.FindName("helpLinkEXOSettings")
        # Fallback auf globale Suche, wenn im Tab nicht gefunden
        if ($null -eq $helpLinkEXOSettings) {
            $helpLinkEXOSettings = Get-XamlElement -ElementName "helpLinkEXOSettings"
        }

        if ($null -ne $helpLinkEXOSettings) {
            $helpLinkEXOSettings.Add_MouseLeftButtonDown({
                try {
                    Start-Process "https://learn.microsoft.com/de-de/powershell/module/exchange/set-organizationconfig?view=exchange-ps"
                } catch {
                    Write-DebugMessage "Fehler beim Öffnen des Help-Links: $($_.Exception.Message)" -Type Error
                }
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
             Write-DebugMessage "Event-Handler für helpLinkEXOSettings hinzugefügt." -Type Info
        } else {
             Write-DebugMessage "Element nicht gefunden: helpLinkEXOSettings (weder im Tab noch global)" -Type Warning
        }

        # Suche die Buttons (versuche zuerst im Tab, dann global)
        $btnGetOrganizationConfig = $tabEXOSettings.FindName("btnGetOrganizationConfig")
        if ($null -eq $btnGetOrganizationConfig) { $btnGetOrganizationConfig = Get-XamlElement -ElementName "btnGetOrganizationConfig" }

        $btnSetOrganizationConfig = $tabEXOSettings.FindName("btnSetOrganizationConfig")
        if ($null -eq $btnSetOrganizationConfig) { $btnSetOrganizationConfig = Get-XamlElement -ElementName "btnSetOrganizationConfig" }

        $btnExportOrgConfig = $tabEXOSettings.FindName("btnExportOrgConfig")
        if ($null -eq $btnExportOrgConfig) { $btnExportOrgConfig = Get-XamlElement -ElementName "btnExportOrgConfig" }


        # Event-Handler für "Aktuelle Einstellungen laden" Button
        if ($null -ne $btnGetOrganizationConfig) {
            $btnGetOrganizationConfig.Add_Click({
                # Führe die Aktion in einem Try/Catch aus, um Fehler abzufangen
                try {
                    Get-CurrentOrganizationConfig # Funktion muss existieren
                } catch {
                    $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Ausführen von Get-CurrentOrganizationConfig."
                    Write-DebugMessage $errMsg -Type Error
                    Show-MessageBox -Message "Fehler beim Laden der Einstellungen: $($_.Exception.Message)" -Title "Fehler" -Type Error
                }
            })
            Write-DebugMessage "Event-Handler für btnGetOrganizationConfig hinzugefügt." -Type Info
        } else {
            Write-DebugMessage "Element nicht gefunden: btnGetOrganizationConfig" -Type Warning
        }

        # Event-Handler für "Einstellungen speichern" Button
        if ($null -ne $btnSetOrganizationConfig) {
            $btnSetOrganizationConfig.Add_Click({
                 # Führe die Aktion in einem Try/Catch aus, um Fehler abzufangen
                try {
                    Set-CustomOrganizationConfig # Funktion muss existieren
                } catch {
                    $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Ausführen von Set-CustomOrganizationConfig."
                    Write-DebugMessage $errMsg -Type Error
                    Show-MessageBox -Message "Fehler beim Speichern der Einstellungen: $($_.Exception.Message)" -Title "Fehler" -Type Error
                }
            })
             Write-DebugMessage "Event-Handler für btnSetOrganizationConfig hinzugefügt." -Type Info
        } else {
            Write-DebugMessage "Element nicht gefunden: btnSetOrganizationConfig" -Type Warning
        }

        # Event-Handler für "Konfiguration exportieren" Button
        if ($null -ne $btnExportOrgConfig) {
            $btnExportOrgConfig.Add_Click({
                 # Führe die Aktion in einem Try/Catch aus, um Fehler abzufangen
                try {
                    Export-OrganizationConfig # Funktion muss existieren
                } catch {
                    $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Ausführen von Export-OrganizationConfig."
                    Write-DebugMessage $errMsg -Type Error
                    Show-MessageBox -Message "Fehler beim Exportieren der Konfiguration: $($_.Exception.Message)" -Title "Fehler" -Type Error
                }
            })
             Write-DebugMessage "Event-Handler für btnExportOrgConfig hinzugefügt." -Type Info
        } else {
            Write-DebugMessage "Element nicht gefunden: btnExportOrgConfig" -Type Warning
        }

        # --- Handler für dynamische Sichtbarkeit der allgemeinen Elemente ---
        try {
            Write-DebugMessage "Initialisiere Handler für dynamische Sichtbarkeit im EXO Settings Tab..." -Type Info
            # TabControl für die Unter-Tabs finden
            # Korrektur: Verwende FindName auf dem Parent-Element, da Get-XamlElement -RootElement nicht unterstützt wird.
            $tabOrgSettingsCtrl = $tabEXOSettings.FindName("tabOrgSettings")
            if ($null -eq $tabOrgSettingsCtrl) {
                # Wenn nicht im Tab gefunden, versuche globale Suche als Fallback (obwohl es wahrscheinlich im Tab sein sollte)
                $tabOrgSettingsCtrl = Get-XamlElement -ElementName "tabOrgSettings"
            }

            if ($null -eq $tabOrgSettingsCtrl) {
                Write-DebugMessage "Fehler: TabControl 'tabOrgSettings' konnte weder innerhalb von 'tabEXOSettings' noch global gefunden werden." -Type Error
                # Frühzeitiger Abbruch oder Standardverhalten definieren, wenn das TabControl fehlt
                # Hier wird angenommen, dass die Funktion ohne dieses Feature fortfahren kann, aber eine Warnung ausgibt.
            }

            # Elemente finden, deren Sichtbarkeit gesteuert werden soll (versuche zuerst im Tab, dann global)
            $gridOrgButtonsCtrl = $tabEXOSettings.FindName("gridOrgButtons")
            if ($null -eq $gridOrgButtonsCtrl) { $gridOrgButtonsCtrl = Get-XamlElement -ElementName "gridOrgButtons" }

            $groupBoxOrgDisplayCtrl = $tabEXOSettings.FindName("groupBoxOrgConfigDisplay")
            if ($null -eq $groupBoxOrgDisplayCtrl) { $groupBoxOrgDisplayCtrl = Get-XamlElement -ElementName "groupBoxOrgConfigDisplay" }


            if ($null -ne $tabOrgSettingsCtrl -and $null -ne $gridOrgButtonsCtrl -and $null -ne $groupBoxOrgDisplayCtrl) {

                # Helfer-Skriptblock zum Aktualisieren der Sichtbarkeit
                # Definiere den Skriptblock außerhalb des Event-Handlers, um Wiederverwendung zu ermöglichen
                $UpdateVisibilityScriptBlock = {
                    param(
                        [System.Windows.Controls.TabItem]$selectedTabItem,
                        [System.Windows.Controls.Grid]$gridButtons,
                        [System.Windows.Controls.GroupBox]$groupBoxDisplay
                    )

                    # Interne Hilfsfunktion für Logging innerhalb des ScriptBlocks
                    $Log = {
                        param($Message, $Type = "Debug")
                        # Sicherstellen, dass Write-DebugMessage verfügbar ist (könnte in anderem Runspace sein)
                        if (Get-Command Write-DebugMessage -ErrorAction SilentlyContinue) {
                            Write-DebugMessage -Message $Message -Type $Type
                        } else {
                            Write-Host "DEBUG [$Type]: $Message" # Fallback
                        }
                    }

                    if ($null -eq $selectedTabItem) {
                        $Log.Invoke("UpdateVisibility: Kein TabItem übergeben.", "Warning")
                        return
                    }
                    if ($null -eq $gridButtons -or $null -eq $groupBoxDisplay) {
                        $Log.Invoke("UpdateVisibility: Grid oder GroupBox nicht übergeben.", "Warning")
                        return
                    }

                    # Sicherstellen, dass Header ein String ist
                    $tabHeader = ""
                    try {
                        if ($selectedTabItem.Header -is [string]) {
                            $tabHeader = $selectedTabItem.Header
                        } elseif ($selectedTabItem.Header -is [System.Windows.Controls.TextBlock]) {
                            $tabHeader = $selectedTabItem.Header.Text
                        } elseif ($null -ne $selectedTabItem.Header) {
                            $tabHeader = $selectedTabItem.Header.ToString()
                        } else {
                            $tabHeader = "[Header ist null]" # Fallback
                        }
                    } catch {
                        $Log.Invoke("Fehler beim Lesen des Tab-Headers: $($_.Exception.Message)", "Warning")
                        $tabHeader = "[Fehler beim Lesen des Headers]"
                    }


                    # Prüfen, ob der "Region"-Tab ausgewählt ist (Groß-/Kleinschreibung ignorieren)
                    # Sicherer Vergleich, der Nullwerte im Header berücksichtigt
                    $isRegionTab = ($null -ne $tabHeader -and $tabHeader -eq "Region")
                    $newVisibility = if($isRegionTab) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }

                    $Log.Invoke("Ausgewählter Sub-Tab: '$tabHeader'. Region-Tab aktiv: $isRegionTab. Setze Sichtbarkeit auf $newVisibility", "Info")

                    # Sichtbarkeit der Steuerelemente anpassen (in Try/Catch zur Sicherheit)
                    try {
                        $gridButtons.Visibility = $newVisibility
                        $groupBoxDisplay.Visibility = $newVisibility
                    } catch {
                        $Log.Invoke("Fehler beim Setzen der Sichtbarkeit: $($_.Exception.Message)", "Error")
                    }
                }

                # Initialen Status beim Laden setzen (wichtig!)
                try {
                     if ($tabOrgSettingsCtrl.Items.Count -gt 0) {
                         $initialTab = $tabOrgSettingsCtrl.SelectedItem -as [System.Windows.Controls.TabItem]
                         # Wenn kein Tab vorausgewählt ist oder SelectedItem nicht vom Typ TabItem ist, nimm den ersten Tab
                         if ($null -eq $initialTab -and $tabOrgSettingsCtrl.SelectedIndex -ge 0) {
                             $initialTab = $tabOrgSettingsCtrl.Items[$tabOrgSettingsCtrl.SelectedIndex] -as [System.Windows.Controls.TabItem]
                         } elseif ($null -eq $initialTab) {
                             # Fallback auf das erste Element, wenn auch kein Index gesetzt ist
                             $initialTab = $tabOrgSettingsCtrl.Items[0] -as [System.Windows.Controls.TabItem]
                         }

                         if ($null -ne $initialTab) {
                            Write-DebugMessage "Setze initiale Sichtbarkeit basierend auf Tab: $($initialTab.Header)" -Type Info
                            # Verwende Invoke() statt '&' und übergebe benötigte Elemente als Argumente
                            $UpdateVisibilityScriptBlock.Invoke($initialTab, $gridOrgButtonsCtrl, $groupBoxOrgDisplayCtrl)
                         } else {
                             Write-DebugMessage "Konnte initialen Tab für Sichtbarkeitsprüfung nicht ermitteln oder er ist kein TabItem." -Type Warning
                         }
                     } else {
                         Write-DebugMessage "Keine Sub-Tabs in 'tabOrgSettings' gefunden für initiale Sichtbarkeitsprüfung." -Type Info
                     }
                } catch {
                     Write-DebugMessage "Fehler beim Setzen der initialen Sichtbarkeit: $($_.Exception.Message)" -Type Warning
                     Write-DebugMessage "Fehlerdetails (Initiale Sichtbarkeit): $($_.ScriptStackTrace)" -Type Warning
                }

                # Event-Handler für SelectionChanged hinzufügen
                $tabOrgSettingsCtrl.Add_SelectionChanged({
                    param($sender, $e)
                    # Nur reagieren, wenn das Event vom TabControl kam
                    if ($e.Source -is [System.Windows.Controls.TabControl]) {
                        $selectedTab = $null
                        # Prüfen, ob $e.AddedItems existiert und Elemente enthält
                        if ($e.AddedItems -ne $null -and $e.AddedItems.Count -gt 0) {
                            $selectedTab = $e.AddedItems[0] -as [System.Windows.Controls.TabItem]
                        } else {
                            # Fallback: Wenn keine AddedItems vorhanden sind (z.B. bei Programmstart oder Clear),
                            # versuche, das aktuell ausgewählte Element zu verwenden.
                            $selectedTab = $sender.SelectedItem -as [System.Windows.Controls.TabItem]
                            Write-DebugMessage "SelectionChanged ohne AddedItems ausgelöst. Verwende Sender.SelectedItem." -Type Debug
                        }

                        if ($null -ne $selectedTab) {
                            # Sichtbarkeit aktualisieren mit Invoke()
                            # Wichtig: $gridOrgButtonsCtrl und $groupBoxOrgDisplayCtrl müssen hier zugänglich sein (Closure)
                            # oder als Argumente übergeben werden. Sicherer ist die Übergabe.
                            # Da sie im übergeordneten Scope definiert sind, sollte die Closure funktionieren.
                            # Zur Sicherheit übergeben wir sie explizit, falls Closure-Verhalten Probleme macht.
                            # $UpdateVisibilityScriptBlock.Invoke($selectedTab, $gridOrgButtonsCtrl, $groupBoxOrgDisplayCtrl, @{LogFunction = $Write-DebugMessage})
                            # Test ohne explizite Übergabe, da sie im Scope sein sollten:
                            try {
                                # Stelle sicher, dass die Variablen im Scope des ScriptBlocks verfügbar sind
                                # oder übergebe sie explizit, wenn nötig. Hier wird Closure angenommen.
                                $localGridButtons = $gridOrgButtonsCtrl # Lokale Kopie für Closure
                                $localGroupBoxDisplay = $groupBoxOrgDisplayCtrl # Lokale Kopie für Closure
                                $UpdateVisibilityScriptBlock.Invoke($selectedTab, $localGridButtons, $localGroupBoxDisplay)
                            } catch {
                                Write-DebugMessage "Fehler beim Aufruf von UpdateVisibilityScriptBlock im Event Handler: $($_.Exception.Message)" -Type Error
                            }

                        } else {
                             Write-DebugMessage "SelectionChanged: Konnte das ausgewählte TabItem nicht ermitteln oder es ist kein TabItem." -Type Warning
                        }
                    }
                })
                Write-DebugMessage "SelectionChanged-Handler für 'tabOrgSettings' erfolgreich hinzugefügt." -Type Success
            } else {
                # Fehler loggen, wenn Elemente nicht gefunden wurden
                $notFound = @()
                if ($null -eq $tabOrgSettingsCtrl) { $notFound += "'tabOrgSettings'" }
                if ($null -eq $gridOrgButtonsCtrl) { $notFound += "'gridOrgButtons'" }
                if ($null -eq $groupBoxOrgDisplayCtrl) { $notFound += "'groupBoxOrgConfigDisplay'" }
                Write-DebugMessage "Konnte Elemente für dynamische Sichtbarkeit nicht finden: $($notFound -join ', ')." -Type Warning
            }
        } catch {
             $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Initialisieren des Handlers für dynamische Sichtbarkeit."
             Write-DebugMessage $errMsg -Type Error
        }
        # --- Ende Handler für dynamische Sichtbarkeit ---


        # Initialize all UI controls for the OrganizationConfig tab
        # Diese Funktion muss sicherstellen, dass sie die Controls auch innerhalb des Tabs findet.
        # Annahme: Initialize-OrganizationConfigControls behandelt dies korrekt.
        try {
            # Es könnte sinnvoll sein, Initialize-OrganizationConfigControls das Tab-Objekt zu übergeben,
            # damit es gezielt darin suchen kann, falls notwendig.
            # Initialize-OrganizationConfigControls -RootElement $tabEXOSettings
            # Aktuell wird keine Übergabe verwendet, Annahme: Sucht global oder hat eigene Logik.
            Initialize-OrganizationConfigControls
        } catch {
            $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Aufruf von Initialize-OrganizationConfigControls."
            Write-DebugMessage $errMsg -Type Error
            # Entscheiden, ob der Fehler kritisch ist und die Initialisierung abbrechen soll
        }


        # Log-Verzeichnis-Prüfung (vereinfacht, da die Erstellung zentral erfolgen sollte)
        if ($script:LoggingEnabled -and $script:LogDir) {
            Write-DebugMessage "Logging ist aktiviert. Log-Verzeichnis: $script:LogDir" -Type Debug
        } else {
             Write-DebugMessage "Logging ist nicht aktiviert oder Log-Verzeichnis nicht konfiguriert." -Type Debug
        }


        Write-DebugMessage "EXO Settings Tab wurde erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        # Fange alle unerwarteten Fehler während der Tab-Initialisierung ab
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Schwerwiegender Fehler beim Initialisieren des EXO Settings Tabs."
        Write-DebugMessage $errorMsg -Type Error # Als Error markieren, da die UI evtl. nicht nutzbar ist
        # Dem Benutzer eine klare Fehlermeldung anzeigen
        Show-MessageBox -Message "Ein kritischer Fehler ist beim Initialisieren des 'EXO Settings'-Tabs aufgetreten: $($_.Exception.Message). Die Funktionalität dieses Tabs ist möglicherweise beeinträchtigt." -Title "Initialisierungsfehler" -Type "Error"
        return $false # Signalisiert, dass die Initialisierung fehlgeschlagen ist
    }
}
#endregion EXOSettings Tab Initialization

#region EXOSettings UI Elements Initialization
function Initialize-OrganizationConfigControls {
    [CmdletBinding()]
    param()

    try {
        Write-DebugMessage "Initialisiere OrganizationConfig UI-Elemente" -Type "Info"

        # Initialisiere die Hashtable für Einstellungen, falls noch nicht geschehen
        if ($null -eq $script:organizationConfigSettings) {
            $script:organizationConfigSettings = @{}
            Write-DebugMessage "organizationConfigSettings Hashtable initialisiert." -Type Debug
        }

        # Eine Liste der Control-Typen und ihrer Event-Handler
        # Die Handler werden als ScriptBlocks definiert
        $controlHandlers = @{
            # Checkboxen
            "CheckBox" = @{
                "EventName" = "Click"
                "Handler" = {
                    param($sender, $e)
                    $checkBox = $sender
                    $checkBoxName = $checkBox.Name
                    # Sicherstellen, dass die Hashtable existiert
                    if ($null -eq $script:organizationConfigSettings) { $script:organizationConfigSettings = @{} }

                    if ($checkBoxName -like "chk*" -and $checkBoxName.Length -gt 3) {
                        $propertyName = $checkBoxName.Substring(3)
                        # Speichere den Wert sicher
                        try {
                            $script:organizationConfigSettings[$propertyName] = $checkBox.IsChecked
                            Write-DebugMessage "Checkbox '$checkBoxName' ($propertyName) Wert auf $($checkBox.IsChecked) gesetzt." -Type "Info"
                        } catch {
                            Write-DebugMessage "Fehler beim Speichern des Werts für Checkbox '$checkBoxName': $($_.Exception.Message)" -Type "Warning"
                        }
                    }
                }
            }

            # ComboBoxes
            "ComboBox" = @{
                "EventName" = "SelectionChanged"
                "Handler" = {
                    param($sender, $e)
                    $comboBox = $sender
                    $comboBoxName = $comboBox.Name
                    # Sicherstellen, dass die Hashtable existiert
                    if ($null -eq $script:organizationConfigSettings) { $script:organizationConfigSettings = @{} }

                    if ($null -ne $comboBox.SelectedItem) {
                        # Inhalt sicher abrufen
                        $selectedContent = ""
                        try {
                            # Annahme: ComboBoxItem -> Content ist der relevante Wert
                            if ($comboBox.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) {
                                $selectedContent = $comboBox.SelectedItem.Content.ToString()
                            } else {
                                # Fallback für direkt hinzugefügte Strings o.ä.
                                $selectedContent = $comboBox.SelectedItem.ToString()
                            }
                        } catch {
                            Write-DebugMessage "Fehler beim Abrufen des SelectedItem.Content für '$comboBoxName': $($_.Exception.Message)" -Type "Warning"
                            return # Verarbeitung abbrechen, wenn Inhalt nicht lesbar
                        }


                        $propertyName = $null
                        $propertyValue = $null

                        # Spezielle Verarbeitung für bestimmte ComboBoxen
                        try {
                            switch -Wildcard ($comboBoxName) {
                                # WICHTIG: Der Name im XAML ist 'chkActivity...' nicht 'cmb...'
                                "chkActivityBasedAuthenticationTimeoutInterval" {
                                    $propertyName = "ActivityBasedAuthenticationTimeoutInterval"
                                    # Format: "01:00:00 (1h)" zu "01:00:00"
                                    $propertyValue = ($selectedContent -split ' ')[0]
                                }
                                "cmbLargeAudienceThreshold" {
                                    $propertyName = "MailTipsLargeAudienceThreshold"
                                    # Konvertieren zu Integer
                                    if ([int]::TryParse($selectedContent, [ref]$null)) {
                                        $propertyValue = [int]$selectedContent
                                    } else {
                                        Write-DebugMessage "Ungültiger Integer-Wert '$selectedContent' für '$comboBoxName'." -Type Warning
                                        return # Ungültigen Wert nicht speichern
                                    }
                                }
                                "cmbInformationBarrierMode" {
                                    $propertyName = "InformationBarrierMode"
                                    $propertyValue = $selectedContent
                                }
                                "cmbEwsAppAccessPolicy" {
                                    $propertyName = "EwsApplicationAccessPolicy"
                                    $propertyValue = $selectedContent
                                }
                                "cmbOfficeFeatures" {
                                    $propertyName = "OfficeFeatures"
                                    $propertyValue = $selectedContent
                                }
                                "cmbSearchQueryLanguage" {
                                    $propertyName = "SearchQueryLanguage"
                                    $propertyValue = $selectedContent
                                }
                                "cmb*" { # Generische Behandlung für andere ComboBoxen mit cmb-Präfix
                                    if ($comboBoxName.Length -gt 3) {
                                        $propertyName = $comboBoxName.Substring(3)
                                        $propertyValue = $selectedContent
                                    }
                                }
                                default {
                                    Write-DebugMessage "Keine spezielle Behandlung für ComboBox '$comboBoxName' definiert." -Type Debug
                                }
                            }

                            # Wert speichern, wenn Eigenschaftsname ermittelt wurde
                            if ($null -ne $propertyName) {
                                $script:organizationConfigSettings[$propertyName] = $propertyValue
                                Write-DebugMessage "ComboBox '$comboBoxName' ($propertyName) Wert auf '$propertyValue' gesetzt." -Type "Info"
                            }

                        } catch {
                            Write-DebugMessage "Fehler in der Switch-Anweisung für ComboBox '$comboBoxName': $($_.Exception.Message)" -Type "Warning"
                        }
                    } else {
                         # Fall behandeln, wenn die Auswahl aufgehoben wird (SelectedItem ist null)
                         # Hier könnte man den Wert in $script:organizationConfigSettings auf $null setzen,
                         # aber das hängt von der gewünschten Logik ab. Vorerst keine Aktion.
                         Write-DebugMessage "ComboBox '$comboBoxName' Auswahl aufgehoben (SelectedItem ist null)." -Type Debug
                    }
                }
            }

            # TextBoxes
            "TextBox" = @{
                "EventName" = "TextChanged" # Oder LostFocus, je nach Anforderung
                "Handler" = {
                    param($sender, $e)
                    $textBox = $sender
                    $textBoxName = $textBox.Name
                    $currentText = $textBox.Text
                    # Sicherstellen, dass die Hashtable existiert
                    if ($null -eq $script:organizationConfigSettings) { $script:organizationConfigSettings = @{} }

                    $propertyName = $null
                    $propertyValue = $null

                    try {
                        # Spezielle Verarbeitung für bestimmte TextBoxen
                        switch ($textBoxName) {
                            "txtPowerShellMaxConcurrency" {
                                $propertyName = "PowerShellMaxConcurrency"
                                if ([int]::TryParse($currentText, [ref]$null)) {
                                    $propertyValue = [int]$currentText
                                } else { Write-DebugMessage "Ungültiger Integer '$currentText' für '$textBoxName'." -Type Warning; return }
                            }
                            "txtPowerShellMaxCmdletQueueDepth" {
                                $propertyName = "PowerShellMaxCmdletQueueDepth"
                                if ([int]::TryParse($currentText, [ref]$null)) {
                                    $propertyValue = [int]$currentText
                                } else { Write-DebugMessage "Ungültiger Integer '$currentText' für '$textBoxName'." -Type Warning; return }
                            }
                            "txtPowerShellMaxCmdletsExecutionDuration" {
                                $propertyName = "PowerShellMaxCmdletsExecutionDuration"
                                if ([int]::TryParse($currentText, [ref]$null)) {
                                    $propertyValue = [int]$currentText
                                } else { Write-DebugMessage "Ungültiger Integer '$currentText' für '$textBoxName'." -Type Warning; return }
                            }
                            "txtDefaultAuthPolicy" {
                                $propertyName = "DefaultAuthenticationPolicy"
                                $propertyValue = $currentText # Text direkt übernehmen
                            }
                            "txtHierAddressBookRoot" {
                                $propertyName = "HierarchicalAddressBookRoot"
                                $propertyValue = $currentText # Text direkt übernehmen
                            }
                            "txtPreferredInternetCodePageForShiftJis" {
                                $propertyName = "PreferredInternetCodePageForShiftJis"
                                if ([int]::TryParse($currentText, [ref]$null)) {
                                    $propertyValue = [int]$currentText
                                } else { Write-DebugMessage "Ungültiger Integer '$currentText' für '$textBoxName'." -Type Warning; return }
                            }
                            default {
                                # Generische Behandlung für andere TextBoxen mit txt-Präfix
                                if ($textBoxName -like "txt*" -and $textBoxName.Length -gt 3) {
                                    $propertyName = $textBoxName.Substring(3)
                                    $propertyValue = $currentText
                                } else {
                                     Write-DebugMessage "Keine spezielle Behandlung für TextBox '$textBoxName' definiert." -Type Debug
                                }
                            }
                        }

                        # Wert speichern, wenn Eigenschaftsname ermittelt wurde
                        if ($null -ne $propertyName) {
                            $script:organizationConfigSettings[$propertyName] = $propertyValue
                            Write-DebugMessage "TextBox '$textBoxName' ($propertyName) Wert auf '$propertyValue' gesetzt." -Type "Info"
                        }
                    } catch {
                        Write-DebugMessage "Fehler in der Switch-Anweisung für TextBox '$textBoxName': $($_.Exception.Message)" -Type "Warning"
                    }
                }
            }
        }

        # Statistiken für die Registrierung
        $registeredControls = @{ CheckBox = 0; ComboBox = 0; TextBox = 0; Failed = 0 }

        # Registriere Event-Handler für bekannte UI-Elemente
        # Stelle sicher, dass $script:knownUIElements definiert und befüllt ist
        if ($null -eq $script:knownUIElements -or $script:knownUIElements.Count -eq 0) {
            Write-DebugMessage "\$script:knownUIElements ist leer oder null. Keine Event-Handler werden registriert." -Type Warning
            # Evtl. hier $script:knownUIElements initialisieren oder aus einer Quelle laden?
        } else {
            Write-DebugMessage "Registriere Event-Handler für $($script:knownUIElements.Count) bekannte UI-Elemente..." -Type Info

            foreach ($elementName in $script:knownUIElements) {
                $element = $null
                $elementType = $null
                try {
                    $element = Get-XamlElement -ElementName $elementName
                    if ($null -ne $element) {
                        $elementType = $element.GetType().Name
                        if ($controlHandlers.ContainsKey($elementType)) {
                            $handlerInfo = $controlHandlers[$elementType]
                            $eventName = "Add_" + $handlerInfo.EventName
                            $handlerScriptBlock = $handlerInfo.Handler

                            # Event-Handler registrieren
                            $element.$eventName($handlerScriptBlock)
                            $registeredControls[$elementType]++
                            # Write-DebugMessage "Event-Handler für $elementName ($elementType) registriert." -Type "Debug" # Optional: Weniger verbose
                        } else {
                            # Write-DebugMessage "Kein Handler für Elementtyp '$elementType' (Element: $elementName) definiert." -Type "Debug"
                        }
                    } else {
                        # Write-DebugMessage "Element '$elementName' nicht gefunden, Handler wird nicht registriert." -Type "Debug" # Optional: Weniger verbose
                    }
                } catch {
                    $registeredControls["Failed"]++
                    Write-DebugMessage "Fehler beim Verarbeiten/Registrieren des Handlers für '$elementName' ($elementType): $($_.Exception.Message)" -Type "Warning"
                }
            }
        }

        # Zusammenfassung der Registrierung
        Write-DebugMessage ("Event-Handler Registrierung abgeschlossen: " +
                          "$($registeredControls.CheckBox) CheckBoxen, " +
                          "$($registeredControls.ComboBox) ComboBoxen, " +
                          "$($registeredControls.TextBox) TextBoxes, " +
                          "$($registeredControls.Failed) Fehler.") -Type "Info"

        # Die spezielle Behandlung für ActivityTimeout ComboBox ist jetzt Teil der generischen Logik,
        # da der XAML-Name 'chkActivityBasedAuthenticationTimeoutInterval' korrekt im Switch behandelt wird.
        # Der separate Block dafür wurde entfernt.

        Write-DebugMessage "OrganizationConfig UI-Elemente erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        # Fange Fehler während der UI-Element-Initialisierung ab
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Initialisieren der OrganizationConfig UI-Elemente."
        Write-DebugMessage $errorMsg -Type "Error"
        # Optional: Dem Benutzer eine Meldung anzeigen, aber die Hauptinitialisierung nicht unbedingt abbrechen
        # Show-MessageBox -Message "Fehler bei UI-Initialisierung: $($_.Exception.Message)" -Title "UI Fehler" -Type Warning
        return $false # Signalisiert, dass dieser Teil fehlgeschlagen ist
    }
}

#region EXOSettings Organization Config Management
function Get-CurrentOrganizationConfig {
    [CmdletBinding()]
    param()

    # Verwende $PSCmdlet, um WriteProgress zu unterstützen
    $cmdlet = $PSCmdlet # Beibehalten, falls $PSCmdlet für andere Zwecke benötigt wird

    try {
        # Prüfen, ob wir mit Exchange verbunden sind
        if (-not (Confirm-ExchangeConnection)) {
            Show-MessageBox -Message "Bitte verbinden Sie sich zuerst mit Exchange Online." -Title "Nicht verbunden" -Type "Warning"
            return
        }

        Write-DebugMessage "Beginne Abruf der aktuellen Organisationseinstellungen..." -Type "Info"
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Lade Organisationseinstellungen..."
        }
        # Fortschritt anzeigen (Initial) - Verwende Write-Progress direkt
        Write-Progress -Activity "Organisationseinstellungen laden" -Status "Rufe Einstellungen von Exchange Online ab..." -PercentComplete 0 -Id 1

        # Organisationseinstellungen abrufen
        # Timeout hinzufügen und Fehler explizit abfangen
        $script:currentOrganizationConfig = $null # Zurücksetzen
        try {
             # Hier könnte man Invoke-Command mit Timeout verwenden, falls Get-OrganizationConfig hängt
             $script:currentOrganizationConfig = Get-OrganizationConfig -ErrorAction Stop
        } catch {
             throw "Fehler beim Abrufen von Get-OrganizationConfig: $($_.Exception.Message)"
        }

        if ($null -eq $script:currentOrganizationConfig) {
            throw "Get-OrganizationConfig hat keine Daten zurückgegeben."
        }

        # Eigenschaften extrahieren
        $configProperties = $script:currentOrganizationConfig.PSObject.Properties | Select-Object -ExpandProperty Name

        # Aktuelle Einstellungs-Hashtable leeren und initialisieren
        $script:organizationConfigSettings = @{}

        # Anzahl Eigenschaften sicher ermitteln
        $propCount = 0
        if ($null -ne $configProperties) {
            # Sicherstellen, dass es sich um eine Sammlung handelt, bevor Count verwendet wird
            if ($configProperties -is [array] -or $configProperties -is [System.Collections.ICollection]) {
                $propCount = $configProperties.Count
            } elseif ($null -ne $configProperties) {
                # Wenn es nur ein einzelnes Element ist (sollte nicht passieren, aber sicher ist sicher)
                $propCount = 1
            }
        }
        Write-DebugMessage "Organisation-Config geladen ($propCount Eigenschaften), aktualisiere UI-Elemente..." -Type "Info"
        Write-Progress -Activity "Organisationseinstellungen laden" -Status "Aktualisiere UI-Elemente..." -PercentComplete 30 -Id 1

        # --- UI-Elemente aktualisieren ---
        # Verwende die bekannte Liste der UI-Elemente, um gezielt zu aktualisieren
        $uiUpdateCounter = 0
        $totalUiElements = $script:knownUIElements.Count

        # Liste der bekannten UI-Elemente, die kein direktes Mapping haben oder anders behandelt werden
        # Diese Liste hilft, unnötige Debug-Meldungen für bekannte Ausnahmen zu vermeiden.
        $ignoredOrSpecialHandledCheckboxes = @(
            "chkMailTipsLargeAudienceThreshold", # Ist ComboBox 'cmbLargeAudienceThreshold'
            "chkInformationBarrierMode",         # Ist ComboBox 'cmbInformationBarrierMode'
            "chkEwsApplicationAccessPolicy",     # Ist ComboBox 'cmbEwsAppAccessPolicy'
            "chkOfficeFeatures",                 # Ist ComboBox 'cmbOfficeFeatures' (Property nicht gefunden)
            "chkDefaultAuthenticationPolicy",    # Ist TextBox 'txtDefaultAuthPolicy'
            "chkHierarchicalAddressBookRoot",    # Ist TextBox 'txtHierAddressBookRoot'
            "chkPreferredInternetCodePageForShiftJis", # Ist TextBox 'txtPreferredInternetCodePageForShiftJis'
            "chkSearchQueryLanguage",            # Ist ComboBox 'cmbSearchQueryLanguage' (Property nicht gefunden)
            'chkAdditionalStorageProvidersBlocked', # Property nicht in Get-OrganizationConfig
            'chkCalendarVersionStoreEnabled',       # Property nicht in Get-OrganizationConfig
            'chkCASMailboxHasPermissionsIncludingSubfolders', # Property nicht in Get-OrganizationConfig
            'chkEcRequiresTls',                     # Property nicht in Get-OrganizationConfig
            'chkOwaRedirectToOD4BThisUserEnabled',  # Property nicht in Get-OrganizationConfig
            'chkImplicitSharingEnabled',            # Property nicht in Get-OrganizationConfig
            'chkOAuthUseBasicAuth',                 # Property nicht in Get-OrganizationConfig
            'chkEwsAllowList',                      # Property ist MultiValued, nicht Bool
            'chkMacOutlook',                        # Kein direktes Mapping (EwsAllowMacOutlook ist bereits gemappt)
            'chkOutlookMobile',                     # Kein direktes Mapping
            'chkWACDiscoveryEndpoint',              # Property ist String, nicht Bool
            'chkEnableUserPowerShell',              # Property nicht in Get-OrganizationConfig
            'chkIsSingleInstance',                  # Property nicht in Get-OrganizationConfig
            'chkOnPremisesDownloadDisabled',        # Property nicht in Get-OrganizationConfig
            'chkAcceptApiLicenseAgreement',         # Property nicht in Get-OrganizationConfig
            'chkMobileToFollowedFolders',           # Property nicht in Get-OrganizationConfig
            'chkVisibilityEnabled',                 # Property nicht in Get-OrganizationConfig
            'chkExecutiveAttestation',              # Property nicht in Get-OrganizationConfig
            'chkPDPLocationEnabled',                # Property nicht in Get-OrganizationConfig
            'chkRemotePublicFolderBlobsEnabled'     # Property nicht gefunden/standard
        )


        foreach ($elementName in $script:knownUIElements) {
            $uiUpdateCounter++
            $progress = 30 + (($uiUpdateCounter / $totalUiElements) * 60) # Fortschritt von 30% bis 90%
            # Fortschritt für jedes Element aktualisieren
            Write-Progress -Activity "Organisationseinstellungen laden" -Status "Aktualisiere UI: $elementName" -PercentComplete ([int]$progress) -Id 1

            $element = Get-XamlElement -ElementName $elementName
            if ($null -eq $element) {
                # Write-DebugMessage "Element '$elementName' zum Aktualisieren nicht gefunden." -Type Debug
                continue # Nächstes Element
            }

            $elementType = $element.GetType().Name
            $propertyName = $null

            try {
                # CheckBoxen
                if ($elementType -eq "CheckBox" -and $elementName -like "chk*") {
                    # Neue Logik: Explizites Mapping für bekannte Checkboxen
                    switch ($elementName) {
                        "chkActivityBasedAuthenticationTimeoutEnabled"                 { $propertyName = "ActivityBasedAuthenticationTimeoutEnabled" }
                        "chkActivityBasedAuthenticationTimeoutWithSingleSignOnEnabled" { $propertyName = "ActivityBasedAuthenticationTimeoutWithSingleSignOnEnabled" }
                        "chkAppsForOfficeEnabled"                      { $propertyName = "AppsForOfficeEnabled" }
                        "chkAsyncSendEnabled"                          { $propertyName = "AsyncSendEnabled" }
                        "chkFocusedInboxOn"                            { $propertyName = "FocusedInboxOn" }
                        "chkReadTrackingEnabled"                       { $propertyName = "ReadTrackingEnabled" }
                        "chkSendFromAliasEnabled"                      { $propertyName = "SendFromAliasEnabled" }
                        "chkBookingsAddressEntryRestricted"            { $propertyName = "BookingsAddressEntryRestricted" }
                        "chkBookingsAuthEnabled"                       { $propertyName = "BookingsAuthEnabled" }
                        "chkBookingsCreationOfCustomQuestionsRestricted" { $propertyName = "BookingsCreationOfCustomQuestionsRestricted" }
                        "chkBookingsExposureOfStaffDetailsRestricted"  { $propertyName = "BookingsExposureOfStaffDetailsRestricted" }
                        "chkBookingsMembershipApprovalRequired"        { $propertyName = "BookingsMembershipApprovalRequired" }
                        "chkBookingsNamingPolicyEnabled"               { $propertyName = "BookingsNamingPolicyEnabled" }
                        # "chkBookingsNamingPolicySuffix" wurde entfernt, da Property String ist
                        "chkBookingsNamingPolicySuffixEnabled"         { $propertyName = "BookingsNamingPolicySuffixEnabled" }
                        "chkBookingsNotesEntryRestricted"              { $propertyName = "BookingsNotesEntryRestricted" }
                        "chkBookingsPaymentsEnabled"                   { $propertyName = "BookingsPaymentsEnabled" }
                        "chkBookingsSocialSharingRestricted"           { $propertyName = "BookingsSocialSharingRestricted" }
                        "chkAuditDisabled"                             { $propertyName = "AuditDisabled" }
                        "chkComplianceEnabled"                         { $propertyName = "ComplianceMLBgdCrawlEnabled" } # Vermutung basierend auf Log
                        "chkCustomerLockboxEnabled"                    { $propertyName = "CustomerLockboxEnabled" }
                        "chkPublicComputersDetectionEnabled"           { $propertyName = "PublicComputersDetectionEnabled" }
                        "chkMailTipsExternalRecipientsTipsEnabled"     { $propertyName = "MailTipsExternalRecipientsTipsEnabled" }
                        "chkMailTipsGroupMetricsEnabled"               { $propertyName = "MailTipsGroupMetricsEnabled" }
                        "chkMailTipsMailboxSourcedTipsEnabled"         { $propertyName = "MailTipsMailboxSourcedTipsEnabled" }
                        "chkAutoEnableArchiveMailbox"                  { $propertyName = "AutoEnableArchiveMailbox" }
                        "chkAutoExpandingArchive"                      { $propertyName = "AutoExpandingArchiveEnabled" }
                        "chkElcProcessingDisabled"                     { $propertyName = "ElcProcessingDisabled" }
                        "chkEnableOutlookEvents"                       { $propertyName = "EnableOutlookEvents" }
                        "chkPublicFolderShowClientControl"             { $propertyName = "PublicFolderShowClientControl" }
                        "chkAutodiscoverPartialDirSync"                { $propertyName = "AutodiscoverPartialDirSync" }
                        "chkOAuth2ClientProfileEnabled"                { $propertyName = "OAuth2ClientProfileEnabled" }
                        "chkRefreshSessionEnabled"                     { $propertyName = "RefreshSessionEnabled" }
                        "chkPerTenantSwitchToESTSEnabled"              { $propertyName = "PerTenantSwitchToESTSEnabled" }
                        "chkEwsEnabled"                                { $propertyName = "EwsEnabled" } # Korrigiert von chkEws
                        "chkEws"                                       { $propertyName = "EwsEnabled" } # Mapping für chkEws hinzugefügt
                        "chkEwsAllowEntourage"                         { $propertyName = "EwsAllowEntourage" }
                        "chkEwsAllowMacOutlook"                        { $propertyName = "EwsAllowMacOutlook" }
                        "chkEwsAllowOutlook"                           { $propertyName = "EwsAllowOutlook" }
                        "chkMobileAppEducationEnabled"                 { $propertyName = "MobileAppEducationEnabled" }
                        "chkConnectorsEnabled"                         { $propertyName = "ConnectorsEnabled" }
                        "chkConnectorsEnabledForYammer"                { $propertyName = "ConnectorsEnabledForYammer" }
                        "chkConnectorsEnabledForTeams"                 { $propertyName = "ConnectorsEnabledForTeams" }
                        "chkConnectorsEnabledForSharepoint"            { $propertyName = "ConnectorsEnabledForSharepoint" }
                        "chkDisablePlusAddressInRecipients"            { $propertyName = "DisablePlusAddressInRecipients" }
                        "chkSIPEnabled"                                { $propertyName = "SIPEnabled" } # Property existiert, aber Log sagt "nicht gefunden" - Mapping beibehalten
                        #"chkRemotePublicFolderBlobsEnabled"            { $propertyName = "RemotePublicFolderBlobsEnabled"} # Property nicht standard/gefunden
                        "chkMapiHttpEnabled"                           { $propertyName = "MapiHttpEnabled" }
                        "chkOnlineMeetingsByDefaultEnabled"            { $propertyName = "OnlineMeetingsByDefaultEnabled" }
                        "chkDirectReportsGroupAutoCreationEnabled"     { $propertyName = "DirectReportsGroupAutoCreationEnabled" }
                        "chkUnblockUnsafeSenderPromptEnabled"          { $propertyName = "UnblockUnsafeSenderPromptEnabled" }
                        default { $propertyName = $null }
                    }

                    if ($null -eq $propertyName) {
                        # Ignoriere Checkboxen aus der definierten Liste oder wenn kein Mapping existiert
                        if ($elementName -notin $ignoredOrSpecialHandledCheckboxes) {
                             Write-DebugMessage "Kein gültiges Mapping für Checkbox '$elementName' definiert oder Property nicht boolesch." -Type Debug
                        }
                        continue # Nächstes UI Element
                    }

                    if ($script:currentOrganizationConfig.PSObject.Properties.Name -contains $propertyName) {
                        $value = $script:currentOrganizationConfig.$propertyName
                        # Sicherstellen, dass der Wert ein Boolean ist oder konvertiert werden kann
                        $boolValue = $false
                        if ($value -is [bool]) {
                            $boolValue = $value
                        } elseif ($null -ne $value) {
                            # Versuche Konvertierung, z.B. von $true/$false Strings oder 0/1
                            try { $boolValue = [System.Convert]::ToBoolean($value) } catch {}
                        }
                        $element.IsChecked = $boolValue
                        $script:organizationConfigSettings[$propertyName] = $boolValue
                        # Write-DebugMessage "Checkbox '$elementName' ($propertyName) auf '$boolValue' gesetzt." -Type Debug
                    } else { Write-DebugMessage "Eigenschaft '$propertyName' für Checkbox '$elementName' nicht in Config gefunden." -Type Debug }
                }
                # ComboBoxen
                elseif ($elementType -eq "ComboBox") {
                    $valueFromConfig = $null
                    $matchFound = $false
                    $propertyName = $null # Initialisieren

                    # Finde den zugehörigen Eigenschaftsnamen basierend auf der ComboBox-Logik
                    # Beachte: Einige CheckBoxen im XAML sind hier als ComboBoxen implementiert
                    switch -Wildcard ($elementName) {
                        "cmbLargeAudienceThreshold" { $propertyName = "MailTipsLargeAudienceThreshold" }
                        "cmbInformationBarrierMode" { $propertyName = "InformationBarrierMode" }
                        "cmbEwsAppAccessPolicy" { $propertyName = "EwsApplicationAccessPolicy" }
                        "chkActivityBasedAuthenticationTimeoutInterval" { $propertyName = "ActivityBasedAuthenticationTimeoutInterval" } # XAML CheckBox, aber ist ComboBox
                        "cmbSearchQueryLanguage" { $propertyName = "SearchQueryLanguage" } # Hinzugefügt
                        default { $propertyName = $null } # Default auf null setzen
                    }

                    if ($null -eq $propertyName) {
                         Write-DebugMessage "Kein gültiges Mapping für ComboBox '$elementName' definiert." -Type Debug
                        continue # Nächstes UI Element
                    }

                    if ($script:currentOrganizationConfig.PSObject.Properties.Name -contains $propertyName) {
                        $valueFromConfig = $script:currentOrganizationConfig.$propertyName
                        $valueString = "" # Initialisieren

                        if ($null -ne $valueFromConfig) {
                            $valueString = $valueFromConfig.ToString()

                            # Sonderfall EWS App Access Policy: Leerer String -> '(Nicht konfiguriert)'
                            if ($elementName -eq "cmbEwsAppAccessPolicy" -and [string]::IsNullOrEmpty($valueString)) {
                                $valueString = "(Nicht konfiguriert)"
                            }
                            # Sonderfall Search Query Language: Leerer/Null String -> '(Nicht konfiguriert)'
                            if ($elementName -eq "cmbSearchQueryLanguage" -and [string]::IsNullOrEmpty($valueString)) {
                                $valueString = "(Nicht konfiguriert)"
                            }

                            # Spezielle Behandlung für Timeout-Intervall (hh:mm:ss)
                            if ($elementName -eq "chkActivityBasedAuthenticationTimeoutInterval") {
                                # Suche nach "hh:mm:ss" am Anfang von "hh:mm:ss (xh)"
                                foreach ($item in $element.Items) {
                                    $itemContent = ""
                                    if ($item -is [System.Windows.Controls.ComboBoxItem]) { $itemContent = $item.Content.ToString() }
                                    else { $itemContent = $item.ToString() }

                                    if ($itemContent.StartsWith($valueString)) {
                                        $element.SelectedItem = $item
                                        $script:organizationConfigSettings[$propertyName] = $valueString # Korrekten Wert speichern (nur hh:mm:ss)
                                        $matchFound = $true; break
                                    }
                                }
                            } else {
                                # Generische Suche nach exaktem String-Match im Content
                                foreach ($item in $element.Items) {
                                    $itemContent = ""
                                    if ($item -is [System.Windows.Controls.ComboBoxItem]) { $itemContent = $item.Content.ToString() }
                                    else { $itemContent = $item.ToString() }

                                    # Vergleich unter Berücksichtigung möglicher Enum-Typen oder String-Repräsentationen
                                    if ($itemContent -eq $valueString) {
                                        $element.SelectedItem = $item
                                        # Wert speichern (ggf. Typkonvertierung und Sonderfälle)
                                        $storedValue = $null # Initialisieren
                                        switch ($propertyName) {
                                            "MailTipsLargeAudienceThreshold" {
                                                if ([int]::TryParse($valueString, [ref]$null)) { $storedValue = [int]$valueString }
                                                else { $storedValue = $null } # Fallback bei ungültigem Int
                                            }
                                            "EwsApplicationAccessPolicy" {
                                                if ($valueString -eq "(Nicht konfiguriert)") { $storedValue = $null }
                                                else { $storedValue = $valueString } # String speichern
                                            }
                                            "SearchQueryLanguage" {
                                                if ($valueString -eq "(Nicht konfiguriert)") { $storedValue = $null }
                                                else { $storedValue = $valueString } # String speichern
                                            }
                                            default { $storedValue = $valueString } # Standard: String speichern
                                        }
                                        $script:organizationConfigSettings[$propertyName] = $storedValue
                                        $matchFound = $true; break
                                    }
                                }
                            }
                        } # Ende if ($null -ne $valueFromConfig)

                        # Wenn kein Match gefunden oder Wert war null, Standard setzen (z.B. erster Eintrag)
                        if (-not $matchFound) {
                            # Loggen, dass kein passender Wert gefunden wurde (Info-Level, da es vorkommen kann)
                            Write-DebugMessage "Kein passender Wert für '$elementName' (Property: '$propertyName', Wert aus Config: '$valueFromConfig') in ComboBox-Items gefunden. Setze Standard (Index 0)." -Type Info
                            if ($element.Items.Count -gt 0) {
                                $element.SelectedIndex = 0
                                # Entsprechenden Standardwert auch in Hashtable speichern
                                if ($null -ne $element.SelectedItem) {
                                    $standardValue = ""
                                    if ($element.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) { $standardValue = $element.SelectedItem.Content.ToString() }
                                    else { $standardValue = $element.SelectedItem.ToString() }

                                    # Speichern des Standardwerts mit korrekter Typkonvertierung/Formatierung/Sonderfällen
                                    $storedValue = $null # Initialisieren
                                    switch ($propertyName) {
                                        "ActivityBasedAuthenticationTimeoutInterval" { $storedValue = ($standardValue -split ' ')[0] } # Nur hh:mm:ss
                                        "MailTipsLargeAudienceThreshold" {
                                            if ([int]::TryParse($standardValue, [ref]$null)) { $storedValue = [int]$standardValue }
                                            else { $storedValue = $null }
                                        }
                                        "EwsApplicationAccessPolicy" {
                                            if ($standardValue -eq "(Nicht konfiguriert)") { $storedValue = $null }
                                            else { $storedValue = $standardValue }
                                        }
                                        "SearchQueryLanguage" {
                                            if ($standardValue -eq "(Nicht konfiguriert)") { $storedValue = $null }
                                            else { $storedValue = $standardValue }
                                        }
                                        default { $storedValue = $standardValue } # Standard: String speichern
                                    }
                                    $script:organizationConfigSettings[$propertyName] = $storedValue
                                    Write-DebugMessage "ComboBox '$elementName' auf Standard '$standardValue' gesetzt." -Type Debug
                                }
                            } else { Write-DebugMessage "ComboBox '$elementName' hat keine Items zum Setzen eines Standards." -Type Warning }
                        } else {
                             Write-DebugMessage "ComboBox '$elementName' ($propertyName) Wert auf '$valueString' gesetzt." -Type Info
                        }
                    } elseif ($null -ne $propertyName) { Write-DebugMessage "Eigenschaft '$propertyName' für ComboBox '$elementName' nicht in Config gefunden." -Type Debug }
                }
                # TextBoxes
                elseif ($elementType -eq "TextBox") {
                    $propertyName = $null
                    switch ($elementName) {
                        "txtDefaultAuthPolicy"    { $propertyName = "DefaultAuthenticationPolicy" }
                        "txtHierAddressBookRoot"  { $propertyName = "HierarchicalAddressBookRoot" }
                        "txtPreferredInternetCodePageForShiftJis" { $propertyName = "PreferredInternetCodePageForShiftJis" }
                        "txtPowerShellMaxConcurrency" { $propertyName = "PowerShellMaxConcurrency" }
                        "txtPowerShellMaxCmdletQueueDepth" { $propertyName = "PowerShellMaxCmdletQueueDepth" }
                        "txtPowerShellMaxCmdletsExecutionDuration" { $propertyName = "PowerShellMaxCmdletsExecutionDuration" }
                        default { $propertyName = $null }
                    }

                    if ($null -ne $propertyName) {
                        if ($script:currentOrganizationConfig.PSObject.Properties.Name -contains $propertyName) {
                            $value = $script:currentOrganizationConfig.$propertyName
                            $valueString = if ($null -ne $value) { $value.ToString() } else { "" }
                            $element.Text = $valueString
                            $script:organizationConfigSettings[$propertyName] = $value # Originalwert speichern (kann null sein)
                            Write-DebugMessage "TextBox '$elementName' ($propertyName) Wert auf '$valueString' gesetzt." -Type Info
                        } else {
                            Write-DebugMessage "Eigenschaft '$propertyName' für TextBox '$elementName' nicht in Config gefunden." -Type Debug
                            $element.Text = "" # Leeren, wenn Property nicht gefunden
                            $script:organizationConfigSettings[$propertyName] = $null
                        }
                    } elseif ($elementName -ne "txtOrganizationConfig") { # Ignoriere bekannte, nicht gemappte Textboxen
                        # Write-DebugMessage "Kein Mapping für TextBox '$elementName' definiert." -Type Debug
                    }
                }
            } catch {
                 Write-DebugMessage "Fehler beim Aktualisieren von UI-Element '$elementName' (Property: '$propertyName'): $($_.Exception.Message)" -Type "Warning"
            }

        } # Ende foreach ($elementName in $script:knownUIElements)

        # Vollständige Konfiguration in TextBox anzeigen (optional)
        $txtOrganizationConfig = Get-XamlElement -ElementName "txtOrganizationConfig"
        if ($null -ne $txtOrganizationConfig) {
            try {
                # Verwende benutzerdefinierte Formatierung für bessere Lesbarkeit
                # Format-List oder Format-Table könnten hier auch Optionen sein
                $configText = $script:currentOrganizationConfig | Out-String -Stream | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                $txtOrganizationConfig.Text = $configText -join [System.Environment]::NewLine
                Write-DebugMessage "Vollständige Konfiguration in txtOrganizationConfig angezeigt." -Type Info
            } catch {
                 Write-DebugMessage "Fehler beim Formatieren/Anzeigen der vollständigen OrganisationConfig: $($_.Exception.Message)" -Type "Warning"
                 $txtOrganizationConfig.Text = "Fehler beim Laden der Konfigurationsdetails."
            }
        }

        # Fortschritt abschließen (Erfolg)
        Write-Progress -Activity "Organisationseinstellungen laden" -Status "Abgeschlossen" -Completed -Id 1
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organisationseinstellungen erfolgreich geladen."
        }
        Write-DebugMessage "Organisationseinstellungen erfolgreich abgerufen und UI aktualisiert." -Type "Success"

    } catch {
        # Fange alle Fehler während des Abrufs und der Verarbeitung ab
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Abrufen oder Verarbeiten der Organisationseinstellungen."
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler beim Laden der Organisationseinstellungen!" # Deutliche Fehlermeldung
        }
        Write-DebugMessage $errorMsg -Type "Error" # Detaillierte Fehlermeldung ins Log
        Show-MessageBox -Message "Fehler beim Laden der Organisationseinstellungen: $($_.Exception.Message)" -Title "Fehler" -Type "Error" # Benutzerdialog
        # Fortschritt abschließen (Fehler)
        # Sicherstellen, dass die ID 1 verwendet wird, um die vorherige Progress Bar zu schließen
        Write-Progress -Activity "Organisationseinstellungen laden" -Status "Fehler" -Completed -Id 1
    } finally {
        # Sicherstellen, dass die Fortschrittsanzeige geschlossen wird, falls sie noch offen ist
        # (Sollte durch -Completed im try/catch bereits geschehen sein, aber als zusätzliche Sicherheit)
        # Write-Progress -Activity "Organisationseinstellungen laden" -Completed -Id 1 # Doppelt gemoppelt, aber sicher.
    }
}
function Set-CustomOrganizationConfig {
    [CmdletBinding()]
    param()

    # Der gesamte Vorgang (Prüfungen, Parameter-Vorbereitung, Speichern) wird von diesem Try/Catch umschlossen.
    try {
        # Prüfen, ob wir mit Exchange verbunden sind
        if (-not (Confirm-ExchangeConnection)) {
            Show-MessageBox -Message "Bitte verbinden Sie sich zuerst mit Exchange Online." -Title "Nicht verbunden" -Type "Warning"
            return
        }

        # Bestätigungsdialog anzeigen
        $confirmResult = Show-MessageBox -Message "Möchten Sie die Organisationseinstellungen wirklich speichern?" -Title "Einstellungen speichern" -Type "Question"
        if ($confirmResult -ne "Yes") {
            Write-DebugMessage "Speichervorgang vom Benutzer abgebrochen." -Type "Info"
            return
        }

        Write-DebugMessage "Speichere Organisationseinstellungen" -Type "Info"
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Speichere Organisationseinstellungen..."
        }

        # Parameter für Set-OrganizationConfig vorbereiten
        $params = @{}
        $validParameterNames = (Get-Command Set-OrganizationConfig).Parameters.Keys

        # Mapping von UI-Namen (ohne Präfix) zu potenziellen Parameternamen
        $controlNameToParameterMap = @{
            # CheckBoxen (Beispiele) -> Boolean Parameter
            "ActivityBasedAuthenticationTimeoutEnabled" = "ActivityBasedAuthenticationTimeoutEnabled"
            "AppsForOfficeEnabled" = "AppsForOfficeEnabled"
            "SendFromAliasEnabled" = "SendFromAliasEnabled"
            "AuditDisabled" = "AuditDisabled"
            "FocusedInboxOn" = "FocusedInboxOn"
            "Ews" = "EwsEnabled" # Annahme: chkEws -> EwsEnabled
            "OutlookMobile" = "OutlookMobileEnabled" # Annahme: chkOutlookMobile -> OutlookMobileEnabled

            # ComboBoxen (Beispiele) -> Verschiedene Typen
            "ActivityBasedAuthenticationTimeoutInterval" = "ActivityBasedAuthenticationTimeoutInterval" # String (TimeSpan Format)
            "LargeAudienceThreshold" = "MailTipsLargeAudienceThreshold" # Int
            "InformationBarrierMode" = "InformationBarrierMode" # String Enum
            "EwsAppAccessPolicy" = "EwsApplicationAccessPolicy" # String Enum
            "SearchQueryLanguage" = "SearchQueryLanguage" # String Enum

            # TextBoxes (Beispiele) -> String oder Int
            "DefaultAuthPolicy" = "DefaultAuthenticationPolicy" # String
            "HierAddressBookRoot" = "HierarchicalAddressBookRoot" # String (DN)
            "PreferredInternetCodePageForShiftJis" = "PreferredInternetCodePageForShiftJis" # Int
            "PowerShellMaxConcurrency" = "PowerShellMaxConcurrency" # Int
            "PowerShellMaxCmdletQueueDepth" = "PowerShellMaxCmdletQueueDepth" # Int
            "PowerShellMaxCmdletsExecutionDuration" = "PowerShellMaxCmdletsExecutionDuration" # String (TimeSpan Format)
        }

        # Iteriere durch die gespeicherten Einstellungen ($script:organizationConfigSettings enthält die Werte aus der UI)
        # und füge sie zu $params hinzu, wenn der Schlüssel einem gültigen Set-OrganizationConfig Parameter entspricht.
        # Dieser Teil ist jetzt innerhalb des Haupt-Try-Blocks
        foreach ($key in $script:organizationConfigSettings.Keys) {
            # Korrekten Parameternamen finden mithilfe des Mappings
            $paramName = $key # Standardannahme, falls kein Mapping existiert
            if ($controlNameToParameterMap.ContainsKey($key)) {
                 $paramName = $controlNameToParameterMap[$key]
            }

            # Nur fortfahren, wenn es ein gültiger Parameter für Set-OrganizationConfig ist
            if ($validParameterNames -contains $paramName) {
                $value = $script:organizationConfigSettings[$key]

                # Allgemeine Behandlung für "(Nicht konfiguriert)" -> $null für alle Typen
                # Dies wird angewendet, bevor spezifische Typkonvertierungen erfolgen.
                if ($value -is [string] -and $value -eq "(Nicht konfiguriert)") {
                    $value = $null
                }

                # Handle spezifische Typkonvertierungen und Validierungen basierend auf dem Zielparameter
                # Diese Switch-Anweisung verarbeitet die Werte, die bereits aus $script:organizationConfigSettings gelesen wurden.
                # GRUPPIERTE CASES WURDEN AUFGETEILT
                switch ($paramName) {
                    "MailTipsLargeAudienceThreshold" {
                        # Wert muss Int sein oder ein String, der zu Int konvertiert werden kann.
                        if ($value -is [string]) {
                            if ([int]::TryParse($value, [ref]$null)) {
                                $value = [int]$value
                            } else {
                                # Ungültiger Int-String (und nicht "(Nicht konfiguriert)") -> $null
                                Write-DebugMessage "Ungültiger Wert für MailTipsLargeAudienceThreshold '$value' wird als \$null behandelt." -Type "Warning"
                                $value = $null
                            }
                        } elseif ($value -ne $null -and $value -isnot [int]) {
                             # Unerwarteter Typ (sollte nicht vorkommen, wenn UI-Logik korrekt ist) -> $null
                             Write-DebugMessage "Unerwarteter Typ für MailTipsLargeAudienceThreshold '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"
                             $value = $null
                        }
                        # $value ist jetzt Int oder $null
                    }
                    "PreferredInternetCodePageForShiftJis" {
                         # Wert muss Int sein oder ein String, der zu Int konvertiert werden kann, oder leer/null.
                         if ($value -is [string]) {
                            if ([string]::IsNullOrWhiteSpace($value)) {
                                $value = $null # Leerer String -> $null
                            } elseif ([int]::TryParse($value, [ref]$null)) {
                                $value = [int]$value
                            } else {
                                # Ungültiger Int-String -> $null
                                Write-DebugMessage "Ungültiger Wert für PreferredInternetCodePageForShiftJis '$value' wird als \$null behandelt." -Type "Warning"
                                $value = $null
                            }
                         } elseif ($value -ne $null -and $value -isnot [int]) {
                             # Unerwarteter Typ -> $null
                             Write-DebugMessage "Unerwarteter Typ für PreferredInternetCodePageForShiftJis '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"
                             $value = $null
                         }
                         # $value ist jetzt Int oder $null
                    }
                    # --- Aufgeteilte Cases für PowerShellMaxConcurrency und PowerShellMaxCmdletQueueDepth ---
                    "PowerShellMaxConcurrency" {
                         # Wert muss Int > 0 sein oder ein String, der zu Int > 0 konvertiert werden kann, oder leer/null.
                         if ($value -is [string]) {
                            if ([string]::IsNullOrWhiteSpace($value)) {
                                $value = $null # Leerer String -> $null
                            } elseif ([int]::TryParse($value, [ref]$numValue)) {
                                if ($numValue -gt 0) {
                                    $value = $numValue
                                    # Zusätzliche Validierungswarnungen für hohe Werte
                                    if ($value -gt 100) { Write-DebugMessage "Warnung: PowerShellMaxConcurrency ist sehr hoch: $value" -Type "Warning" }
                                } else {
                                    # Wert <= 0 -> $null
                                    Write-DebugMessage "Wert für $paramName muss größer 0 sein ('$value'), wird als \$null behandelt." -Type "Warning"
                                    $value = $null
                                }
                            } else {
                                # Ungültiger Int-String -> $null
                                Write-DebugMessage "Ungültiger numerischer Wert für $paramName '$value' wird als \$null behandelt." -Type "Warning"
                                $value = $null
                            }
                         } elseif ($value -is [int]) {
                             # Wenn es bereits Int ist, prüfe ob > 0
                             if ($value -le 0) {
                                 Write-DebugMessage "Wert für $paramName muss größer 0 sein ('$value'), wird als \$null behandelt." -Type "Warning"
                                 $value = $null
                             }
                             # Ansonsten ist der Int-Wert gültig
                         } elseif ($value -ne $null) {
                             # Unerwarteter Typ -> $null
                             Write-DebugMessage "Unerwarteter Typ für $paramName '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"
                             $value = $null
                         }
                         # $value ist jetzt Int > 0 oder $null
                    }
                    "PowerShellMaxCmdletQueueDepth" {
                         # Wert muss Int > 0 sein oder ein String, der zu Int > 0 konvertiert werden kann, oder leer/null.
                         if ($value -is [string]) {
                            if ([string]::IsNullOrWhiteSpace($value)) {
                                $value = $null # Leerer String -> $null
                            } elseif ([int]::TryParse($value, [ref]$numValue)) {
                                if ($numValue -gt 0) {
                                    $value = $numValue
                                     # Zusätzliche Validierungswarnungen für hohe Werte
                                    if ($value -gt 1000) { Write-DebugMessage "Warnung: PowerShellMaxCmdletQueueDepth ist sehr hoch: $value" -Type "Warning" }
                                } else {
                                    # Wert <= 0 -> $null
                                    Write-DebugMessage "Wert für $paramName muss größer 0 sein ('$value'), wird als \$null behandelt." -Type "Warning"
                                    $value = $null
                                }
                            } else {
                                # Ungültiger Int-String -> $null
                                Write-DebugMessage "Ungültiger numerischer Wert für $paramName '$value' wird als \$null behandelt." -Type "Warning"
                                $value = $null
                            }
                         } elseif ($value -is [int]) {
                             # Wenn es bereits Int ist, prüfe ob > 0
                             if ($value -le 0) {
                                 Write-DebugMessage "Wert für $paramName muss größer 0 sein ('$value'), wird als \$null behandelt." -Type "Warning"
                                 $value = $null
                             }
                             # Ansonsten ist der Int-Wert gültig
                         } elseif ($value -ne $null) {
                             # Unerwarteter Typ -> $null
                             Write-DebugMessage "Unerwarteter Typ für $paramName '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"
                             $value = $null
                         }
                         # $value ist jetzt Int > 0 oder $null
                    }
                    # --- Aufgeteilte Cases für String/Enum Parameter ---
                    "PowerShellMaxCmdletsExecutionDuration" { # String erwartet (TimeSpan Format)
                         # Für String-Parameter: Leerer/Whitespace String -> $null
                         # Für Enum-Parameter: Set-OrganizationConfig behandelt die Validierung. "(Nicht konfiguriert)" wurde bereits zu $null.
                         if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) {
                             $value = $null
                         }
                         # Hier könnten weitere String-Format-Validierungen erfolgen (z.B. für TimeSpan)
                         # Aktuell wird der String (oder $null) direkt übergeben.
                    }
                    "DefaultAuthenticationPolicy" {             # String erwartet
                         if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) {
                             $value = $null
                         }
                    }
                    "HierarchicalAddressBookRoot" {             # String erwartet (DN)
                         if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) {
                             $value = $null
                         }
                    }
                    "ActivityBasedAuthenticationTimeoutInterval" { # String erwartet (TimeSpan Format)
                         if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) {
                             $value = $null
                         }
                    }
                    "InformationBarrierMode" {                  # String Enum erwartet
                         if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) {
                             $value = $null
                         }
                    }
                    "EwsApplicationAccessPolicy" {              # String Enum erwartet
                         if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) {
                             $value = $null
                         }
                    }
                    "SearchQueryLanguage" {                    # String Enum erwartet
                         if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) {
                             $value = $null
                         }
                    }
                    # --- Aufgeteilte Cases für Boolean Parameter ---
                    "ActivityBasedAuthenticationTimeoutEnabled" {
                         # Wert sollte bereits Boolean sein (oder $null, wenn nicht konfiguriert)
                         if ($value -is [string]) {
                             if ($value -eq 'True') { $value = $true }
                             elseif ($value -eq 'False') { $value = $false }
                             else {
                                 Write-DebugMessage "Ungültiger boolescher String für $paramName '$value' wird als \$null behandelt." -Type "Warning"; $value = $null
                             }
                         } elseif ($value -ne $null -and $value -isnot [bool]) {
                             Write-DebugMessage "Unerwarteter Typ für booleschen Parameter $paramName '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"; $value = $null
                         }
                    }
                    "AppsForOfficeEnabled" {
                         if ($value -is [string]) {
                             if ($value -eq 'True') { $value = $true }
                             elseif ($value -eq 'False') { $value = $false }
                             else {
                                 Write-DebugMessage "Ungültiger boolescher String für $paramName '$value' wird als \$null behandelt." -Type "Warning"; $value = $null
                             }
                         } elseif ($value -ne $null -and $value -isnot [bool]) {
                             Write-DebugMessage "Unerwarteter Typ für booleschen Parameter $paramName '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"; $value = $null
                         }
                    }
                    "SendFromAliasEnabled" {
                         if ($value -is [string]) {
                             if ($value -eq 'True') { $value = $true }
                             elseif ($value -eq 'False') { $value = $false }
                             else {
                                 Write-DebugMessage "Ungültiger boolescher String für $paramName '$value' wird als \$null behandelt." -Type "Warning"; $value = $null
                             }
                         } elseif ($value -ne $null -and $value -isnot [bool]) {
                             Write-DebugMessage "Unerwarteter Typ für booleschen Parameter $paramName '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"; $value = $null
                         }
                    }
                    "AuditDisabled" {
                         if ($value -is [string]) {
                             if ($value -eq 'True') { $value = $true }
                             elseif ($value -eq 'False') { $value = $false }
                             else {
                                 Write-DebugMessage "Ungültiger boolescher String für $paramName '$value' wird als \$null behandelt." -Type "Warning"; $value = $null
                             }
                         } elseif ($value -ne $null -and $value -isnot [bool]) {
                             Write-DebugMessage "Unerwarteter Typ für booleschen Parameter $paramName '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"; $value = $null
                         }
                    }
                    "FocusedInboxOn" {
                         if ($value -is [string]) {
                             if ($value -eq 'True') { $value = $true }
                             elseif ($value -eq 'False') { $value = $false }
                             else {
                                 Write-DebugMessage "Ungültiger boolescher String für $paramName '$value' wird als \$null behandelt." -Type "Warning"; $value = $null
                             }
                         } elseif ($value -ne $null -and $value -isnot [bool]) {
                             Write-DebugMessage "Unerwarteter Typ für booleschen Parameter $paramName '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"; $value = $null
                         }
                    }
                    "EwsEnabled" {
                         if ($value -is [string]) {
                             if ($value -eq 'True') { $value = $true }
                             elseif ($value -eq 'False') { $value = $false }
                             else {
                                 Write-DebugMessage "Ungültiger boolescher String für $paramName '$value' wird als \$null behandelt." -Type "Warning"; $value = $null
                             }
                         } elseif ($value -ne $null -and $value -isnot [bool]) {
                             Write-DebugMessage "Unerwarteter Typ für booleschen Parameter $paramName '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"; $value = $null
                         }
                    }
                    "OutlookMobileEnabled" {
                         if ($value -is [string]) {
                             if ($value -eq 'True') { $value = $true }
                             elseif ($value -eq 'False') { $value = $false }
                             else {
                                 Write-DebugMessage "Ungültiger boolescher String für $paramName '$value' wird als \$null behandelt." -Type "Warning"; $value = $null
                             }
                         } elseif ($value -ne $null -and $value -isnot [bool]) {
                             Write-DebugMessage "Unerwarteter Typ für booleschen Parameter $paramName '$($value.GetType().Name)' wird als \$null behandelt." -Type "Warning"; $value = $null
                         }
                    }
                    # Füge hier weitere Cases für andere Parameter mit spezieller Behandlung hinzu, falls nötig
                    default {
                        # Keine spezielle Behandlung definiert, Wert wird wie gelesen verwendet
                        # (nach der "(Nicht konfiguriert)" -> $null Konvertierung)
                    }
                } # Ende switch ($paramName)

                # Füge den (potenziell konvertierten/validierten) Wert zur Parameter-Hashtable hinzu.
                # $null Werte werden explizit übergeben, was Set-OrganizationConfig oft zum Zurücksetzen/Löschen des Werts verwendet.
                $params[$paramName] = $value
                # Write-DebugMessage "Parameter '$paramName' vorbereitet mit Wert: '$value' (Typ: $(if ($null -ne $value) {$value.GetType().Name} else {'null'}))" -Type Debug # Optionales detailliertes Log

            } elseif ($null -ne $script:organizationConfigSettings[$key] -and $script:organizationConfigSettings[$key] -ne "(Nicht konfiguriert)") {
                 # Logge ignorierte Parameter nur, wenn sie ursprünglich einen relevanten Wert hatten
                 # Dies passiert, wenn der Key aus $script:organizationConfigSettings nicht zu einem gültigen Parameter gemappt werden konnte.
                 Write-DebugMessage "Ungültiger oder unbekannter Parameter '$paramName' (aus Key '$key' mit Wert '$($script:organizationConfigSettings[$key])') wird ignoriert." -Type "Warning"
            }
        } # Ende foreach ($key in $script:organizationConfigSettings.Keys)

        # Debug-Log alle Parameter, die tatsächlich an das Cmdlet übergeben werden
        Write-DebugMessage "Folgende Parameter werden an Set-OrganizationConfig übergeben:" -Type "Info"
        if ($params.Count -gt 0) {
            foreach ($key in $params.Keys | Sort-Object) {
                $displayValue = if ($null -eq $params[$key]) { "\$null" } elseif ($params[$key] -is [bool]) { "\$($params[$key])" } else { "'$($params[$key])'" }
                Write-DebugMessage "  - $key = $displayValue" -Type "Info"
            }
        } else {
             Write-DebugMessage "  (Keine gültigen Parameter zum Ändern gefunden)" -Type "Info"
        }


        if ($params.Count -eq 0) {
            Write-DebugMessage "Keine gültigen Parameter zum Speichern gefunden." -Type "Warning"
            Show-MessageBox -Message "Es wurden keine gültigen Änderungen zum Speichern gefunden." -Title "Keine Änderungen" -Type "Info"
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Text = "Keine Änderungen zum Speichern."
            }
            return # Verlässt die Funktion, da nichts zu tun ist
        }

        # Organisationseinstellungen aktualisieren
        # Dieser Aufruf ist immer noch innerhalb des Haupt-Try-Blocks
        Set-OrganizationConfig @params -ErrorAction Stop

        # Erfolgsmeldungen
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organisationseinstellungen erfolgreich gespeichert."
        }
        Write-DebugMessage "Organisationseinstellungen erfolgreich gespeichert" -Type "Success"
        Show-MessageBox -Message "Die Organisationseinstellungen wurden erfolgreich gespeichert." -Title "Erfolg" -Type "Info"

        # Aktuelle Konfiguration neu laden, um Änderungen in der UI zu reflektieren
        Get-CurrentOrganizationConfig

    } # Ende des Haupt-Try-Blocks
    catch {
        # Dieser Catch-Block fängt Fehler aus dem gesamten Try-Block ab
        # (Verbindungsprüfung, Parameter-Vorbereitung, Set-OrganizationConfig)
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Verarbeiten oder Speichern der Organisationseinstellungen."
        if ($null -ne $script:txtStatus) {
            # Kürzere Meldung für Statusleiste, aber spezifischer als nur "Fehler beim Speichern"
            $statusError = if ($_.FullyQualifiedErrorId -like "*SetOrganizationConfig*") {
                               "Fehler beim Speichern: $($_.Exception.Message)"
                           } else {
                               "Fehler bei Verarbeitung: $($_.Exception.Message)"
                           }
            $script:txtStatus.Text = $statusError
        }
        Write-DebugMessage $errorMsg -Type "Error" # Detaillierte Fehlermeldung ins Log
        Show-MessageBox -Message "Fehler beim Verarbeiten oder Speichern der Organisationseinstellungen: $($_.Exception.Message)" -Title "Fehler" -Type "Error" # Benutzerdialog
    }
    # Kein Finally-Block nötig, da keine spezifischen Ressourcen bereinigt werden müssen.
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
        Write-DebugMessage "Debug-Logging deaktiviert" -Type "IWrite-Lognfo"
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
                    if ([string]::IsNullOrWhiteSpace($script:txtExistingGroupName.Text)) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie den Namen der Gruppe an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Funktion zur Gruppenlöschung aufrufen
                    $result = Remove-DistributionGroupAction -GroupName $script:txtExistingGroupName.Text
                    
                    if ($result) {
                        $script:txtStatus.Text = "Gruppe wurde erfolgreich gelöscht."
                        
                        # Felder zurücksetzen
                        $script:txtExistingGroupName.Text = ""
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
                        [System.Windows.MessageBox]::Show("Bitte geben Sie den Namen der Gruppe und den Benutzer an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Funktion zum Hinzufügen eines Benutzers zur Gruppe aufrufen
                    $result = Add-GroupMemberAction -GroupName $script:txtExistingGroupName.Text -User $script:txtGroupUser.Text
                    
                    if ($result) {
                        $script:txtStatus.Text = "Benutzer wurde erfolgreich zur Gruppe hinzugefügt."
                        
                        # Felder zurücksetzen
                        $script:txtGroupUser.Text = ""
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
                        [System.Windows.MessageBox]::Show("Bitte geben Sie den Namen der Gruppe und den Benutzer an.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    # Funktion zum Entfernen eines Benutzers aus der Gruppe aufrufen
                    $result = Remove-GroupMemberAction -GroupName $script:txtExistingGroupName.Text -User $script:txtGroupUser.Text
                    
                    if ($result) {
                        $script:txtStatus.Text = "Benutzer wurde erfolgreich aus der Gruppe entfernt."
                        
                        # Felder zurücksetzen
                        $script:txtGroupUser.Text = ""
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
            Write-DebugMessage "Aktualisiere Shared Mailbox Liste..." -Type Info
            
            $script:sharedMailboxes = Get-EXOSharedMailboxList -ErrorAction Stop
            
            if ($null -ne $script:lstSharedMailboxes) {
                $script:lstSharedMailboxes.Dispatcher.Invoke([Action]{
                    $script:lstSharedMailboxes.ItemsSource = $script:sharedMailboxes
                    $script:lstSharedMailboxes.Items.Refresh()
                }, "Normal")
            }
            
            Write-DebugMessage "Shared Mailbox Liste aktualisiert" -Type Success
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Aktualisieren der Shared Mailbox Liste: $errorMsg" -Type Error
            Show-MessageBox -Message $errorMsg -Title "Aktualisierungsfehler" -Type Error
        }
    }

    function Get-EXOSharedMailboxList {
        [CmdletBinding()]
        param()

        try {
            return Get-Mailbox -ResultSize Unlimited -Filter "RecipientTypeDetails -eq 'SharedMailbox'" | 
                Select-Object DisplayName,PrimarySmtpAddress,@{Name="Size";Expression={$_.ProhibitSendReceiveQuota}}
        }
        catch {
            throw "Fehler beim Abrufen der Shared Mailboxes: $($_.Exception.Message)"
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
            $lstSharedMailboxPermissions = Get-XamlElement -ElementName "lstSharedMailboxPermissions" # Angenommenes Element für Berechtigungsliste
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
            $script:lstSharedMailboxPermissions = $lstSharedMailboxPermissions
            # Stelle sicher, dass das ComboBox-Objekt existiert
            if ($null -ne $script:cmbSharedMailboxDomain) {
                Write-DebugMessage "Fülle cmbSharedMailboxDomain..." -Type Info
                try {
                    # Lade Accepted Domains (falls noch nicht geschehen)
                    # Annahme: Verbindung besteht oder Fehler wird abgefangen
                    $acceptedDomains = Get-AcceptedDomain -ErrorAction Stop | Select-Object -ExpandProperty DomainName
                    
                    $script:cmbSharedMailboxDomain.Dispatcher.Invoke([Action]{
                        $script:cmbSharedMailboxDomain.Items.Clear() # Vorherige Einträge löschen
                        if ($null -ne $acceptedDomains) {
                            foreach ($domain in $acceptedDomains) {
                                [void]$script:cmbSharedMailboxDomain.Items.Add($domain)
                            }
                            if ($script:cmbSharedMailboxDomain.Items.Count -gt 0) {
                                $script:cmbSharedMailboxDomain.SelectedIndex = 0 # Ersten Eintrag auswählen
                            }
                        } else {
                            Write-DebugMessage "Keine akzeptierten Domains gefunden zum Befüllen von cmbSharedMailboxDomain." -Type Warning
                            [void]$script:cmbSharedMailboxDomain.Items.Add("Keine Domains")
                            $script:cmbSharedMailboxDomain.IsEnabled = $false
                        }
                    }) # End Invoke
                } catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Abrufen der Accepted Domains für cmbSharedMailboxDomain: $errorMsg" -Type Error
                    # UI im Fehlerfall aktualisieren (optional, aber empfohlen)
                    $script:cmbSharedMailboxDomain.Dispatcher.Invoke([Action]{
                        $script:cmbSharedMailboxDomain.Items.Clear()
                        [void]$script:cmbSharedMailboxDomain.Items.Add("Fehler/Keine Verbindung")
                        $script:cmbSharedMailboxDomain.IsEnabled = $false
                    }) # End Invoke
                }
            } else {
                Write-DebugMessage "ComboBox cmbSharedMailboxDomain konnte nicht gefunden werden." -Type Error
            }
            $script:txtSharedMailboxPermSource = $txtSharedMailboxPermSource
            $script:txtSharedMailboxPermUser = $txtSharedMailboxPermUser
            $script:cmbSharedMailboxPermType = $cmbSharedMailboxPermType
            $script:cmbSharedMailboxSelect = $cmbSharedMailboxSelect
            # Stelle sicher, dass das ComboBox-Objekt existiert
            if ($null -ne $script:cmbSharedMailboxSelect) {
                Write-DebugMessage "Fülle cmbSharedMailboxSelect..." -Type Info
                try {
                    # Lade Shared Mailboxes (falls noch nicht geschehen)
                    # Annahme: Verbindung besteht oder Fehler wird abgefangen
                    $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop | Select-Object -ExpandProperty PrimarySmtpAddress
                    
                    $script:cmbSharedMailboxSelect.Dispatcher.Invoke([Action]{
                        $script:cmbSharedMailboxSelect.Items.Clear() # Vorherige Einträge löschen
                        if ($null -ne $sharedMailboxes -and $sharedMailboxes.Count -gt 0) {
                            foreach ($mailbox in $sharedMailboxes) {
                                [void]$script:cmbSharedMailboxSelect.Items.Add($mailbox)
                            }
                            if ($script:cmbSharedMailboxSelect.Items.Count -gt 0) {
                                # Optional: Index setzen oder leer lassen
                                # $script:cmbSharedMailboxSelect.SelectedIndex = 0
                                $script:cmbSharedMailboxSelect.Text = "Bitte auswählen..." # Platzhaltertext
                                $script:cmbSharedMailboxSelect.IsEnabled = $true
                            }
                        } else {
                            Write-DebugMessage "Keine Shared Mailboxes gefunden zum Befüllen von cmbSharedMailboxSelect." -Type Warning
                            $script:cmbSharedMailboxSelect.Text = "Keine gefunden"
                            $script:cmbSharedMailboxSelect.IsEnabled = $false
                        }
                    }) # End Invoke
                } catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Abrufen der Shared Mailboxes für cmbSharedMailboxSelect: $errorMsg" -Type Error
                    # UI im Fehlerfall aktualisieren (optional, aber empfohlen)
                    $script:cmbSharedMailboxSelect.Dispatcher.Invoke([Action]{
                        $script:cmbSharedMailboxSelect.Items.Clear()
                        [void]$script:cmbSharedMailboxSelect.Items.Add("Fehler/Keine Verbindung")
                        $script:cmbSharedMailboxSelect.Text = "Fehler/Keine Verbindung"
                        $script:cmbSharedMailboxSelect.IsEnabled = $false
                    }) # End Invoke
                }
            } else {
                Write-DebugMessage "ComboBox cmbSharedMailboxSelect konnte nicht gefunden werden." -Type Error
            }
            $script:chkAutoMapping = $chkAutoMapping
            $script:txtForwardingAddress = $txtForwardingAddress
            
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
                    } elseif ([string]::IsNullOrWhiteSpace($mailboxEmail)) {
                         [System.Windows.MessageBox]::Show("Bitte geben Sie eine E-Mail-Adresse an oder wählen Sie eine Domain aus.", 
                            "Unvollständige Angaben", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return                       
                    }
                    
                    # Funktion zur Erstellung der Shared Mailbox aufrufen
                    $result = New-SharedMailboxAction -Name $script:txtSharedMailboxName.Text -EmailAddress $mailboxEmail
                    
                    if ($result) {
                        $script:txtStatus.Text = "Shared Mailbox wurde erfolgreich erstellt."
                        
                        # Felder zurücksetzen
                        $script:txtSharedMailboxName.Text = ""
                        $script:txtSharedMailboxEmail.Text = ""
                        
                        # Shared Mailbox Liste aktualisieren
                        RefreshSharedMailboxList # Ruft die Funktion auf, die cmbSharedMailboxSelect neu befüllt
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
                        $convertResult = Convert-ToSharedMailboxAction -Identity $script:txtSharedMailboxEmail.Text # Variable umbenannt, um Konflikt zu vermeiden
                        
                        if ($convertResult) {
                            $script:txtStatus.Text = "Postfach wurde erfolgreich in eine Shared Mailbox umgewandelt."
                            
                            # Shared Mailbox Liste aktualisieren
                            RefreshSharedMailboxList
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
                } # Closing brace for try block
                catch {
                    $errorMsg = $_.Exception.Message
                    # Corrected error message context
                    Write-DebugMessage "Fehler beim Entfernen der Berechtigung: $errorMsg" -Type "Error" 
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnRemoveSharedMailboxPermission" # Closing brace for handler script block
            
            # Event-Handler für andere Steuerelemente (Show, Update, Set, Hide, Remove, Help) hier hinzufügen...
            # Beispiel für Hilfe-Link (ähnlich wie im Gruppen-Tab)
            if ($null -ne $helpLinkShared) {
                $helpLinkShared.Add_MouseLeftButtonDown({
                    Show-HelpDialog -Topic "SharedMailboxes"
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

            # Event-Handler für btnShowSharedMailboxes
            Register-EventHandler -Control $btnShowSharedMailboxes -Handler {
                # Ruft die Funktion auf, die die Liste holt und cmbSharedMailboxSelect aktualisiert
                RefreshSharedMailboxList 
            } -ControlName "btnShowSharedMailboxes"

            # Event-Handler für btnShowSharedMailboxPerms (Beispiel, muss implementiert werden)
            Register-EventHandler -Control $btnShowSharedMailboxPerms -Handler {
                try {
                    if (-not $script:isConnected) {
                        Show-MessageBox -Message "Bitte zuerst mit Exchange verbinden." -Title "Nicht verbunden" -Type Info 
                        return
                    }
                    # Prüfen ob eine Mailbox ausgewählt wurde
                    $selectedMailbox = $null
                    if ($script:cmbSharedMailboxSelect.SelectedItem -ne $null) {
                        $selectedMailbox = $script:cmbSharedMailboxSelect.SelectedItem.ToString()
                    } elseif (-not [string]::IsNullOrWhiteSpace($script:cmbSharedMailboxSelect.Text) -and $script:cmbSharedMailboxSelect.Text -ne "Bitte auswählen..." -and $script:cmbSharedMailboxSelect.Text -ne "Keine gefunden" -and $script:cmbSharedMailboxSelect.Text -ne "Fehler/Keine Verbindung") {
                        # Fallback: Textinhalt verwenden, wenn kein Item ausgewählt ist (z.B. bei manueller Eingabe)
                        $selectedMailbox = $script:cmbSharedMailboxSelect.Text
                    }

                    if ([string]::IsNullOrWhiteSpace($selectedMailbox)) {
                         Show-MessageBox -Message "Bitte wählen Sie zuerst eine Shared Mailbox aus der Liste aus oder geben Sie eine gültige Adresse ein." -Title "Keine Auswahl" -Type Info 
                         return
                    }
                    # Hier Logik zum Anzeigen der Berechtigungen implementieren
                    # z.B. Aufruf einer Funktion Show-SharedMailboxPermissionsAction -Mailbox $selectedMailbox
                    $script:txtStatus.Text = "Anzeigen der Berechtigungen für $selectedMailbox..."
                    # Beispiel: Show-SharedMailboxPermissionsAction -Mailbox $selectedMailbox # Diese Funktion muss existieren
                    Write-DebugMessage "Anzeigen der Berechtigungen für $selectedMailbox angefordert." -Type Info
                    Show-MessageBox -Message "Funktion zum Anzeigen der Berechtigungen noch nicht implementiert." -Title "Info" -Type Info 

                } catch {
                     $errorMsg = $_.Exception.Message
                     Write-DebugMessage "Fehler beim Anzeigen der Shared Mailbox Berechtigungen: $errorMsg" -Type "Error"
                     Show-MessageBox -Message "Fehler beim Anzeigen der Berechtigungen: $errorMsg" -Title "Fehler" -Type Error 
                     if ($null -ne $script:txtStatus) {
                         $script:txtStatus.Dispatcher.InvokeAsync({ $script:txtStatus.Text = "Fehler beim Anzeigen der Berechtigungen." }) | Out-Null
                     }
                }
            } -ControlName "btnShowSharedMailboxPerms"

            # Weitere Event-Handler für UpdateAutoMapping, SetForwarding, Hide/Show GAL, RemoveSharedMailbox...

            Write-DebugMessage "Shared Mailbox-Tab erfolgreich initialisiert" -Type "Success"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Initialisieren des Shared Mailbox-Tabs: $errorMsg" -Type "Error"
            # Optional: Statusleiste aktualisieren, wenn Initialisierung fehlschlägt
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Dispatcher.InvokeAsync({ $script:txtStatus.Text = "Fehler beim Initialisieren des Shared Mailbox Tabs." }) | Out-Null
            }
            return $false
        }
    }

#region Resources Tab Functions
# -----------------------------------------------
# Funktionen für den Ressourcen-Tab
# -----------------------------------------------

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
            # Überprüfen, ob Elemente bereits hinzugefügt wurden (z.B. durch XAML)
            $currentItems = @($script:cmbResourceType.Items | ForEach-Object { $_.ToString() })
            $requiredItems = @("Room", "Equipment")
            $needsRefresh = $false
            
            # Prüfen, ob alle benötigten Elemente vorhanden sind
            if ($currentItems.Count -ne $requiredItems.Count) {
            $needsRefresh = $true
            } else {
            foreach ($item in $requiredItems) {
                if ($currentItems -notcontains $item) {
                $needsRefresh = $true
                break
                }
            }
            }
            
            # Nur aktualisieren, wenn nötig
            if ($needsRefresh) {
            Write-DebugMessage "Fülle cmbResourceType mit erforderlichen Optionen." -Type "Info"
            $script:cmbResourceType.Items.Clear()
            foreach ($item in $requiredItems) {
                [void]$script:cmbResourceType.Items.Add($item)
            }
            $script:cmbResourceType.SelectedIndex = 0
            } else {
            Write-DebugMessage "cmbResourceType enthält bereits die erforderlichen Elemente." -Type "Info"
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
             Write-DebugMessage "Registriere Event-Handler für helpLinkResources" -Type Info
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

# Funktion zum Erstellen einer neuen Ressource (Raum oder Ausstattung)
function Create-ResourceAction {
    [CmdletBinding(SupportsShouldProcess=$true)] # ShouldProcess hinzugefügt
    param(
        [Parameter(Mandatory = $true)]
        [string]$ResourceName,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Room", "Equipment")]
        [string]$ResourceType
    )

    Write-DebugMessage "Funktion Create-ResourceAction aufgerufen für '$ResourceName' ($ResourceType)" -Type "Info"

    # Prüfen, ob eine Exchange-Verbindung besteht
    if (-not $script:isConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type "Warning"
        return $false
    }

    # ShouldProcess-Prüfung
    if (-not $pscmdlet.ShouldProcess("Ressource '$ResourceName' ($ResourceType)", "Erstellen")) {
        Write-DebugMessage "Erstellung von '$ResourceName' durch ShouldProcess abgebrochen." -Type "Info"
        return $false
    }

    try {
        $script:txtStatus.Text = "Erstelle Ressource '$ResourceName'..."
        # Eindeutigen UPN generieren (Beispiel: Annahme einer Standarddomäne)
        # $domain = ($script:ConnectedUserPrincipalName -split '@')[1] # Domäne vom verbundenen Benutzer holen
        # if ([string]::IsNullOrWhiteSpace($domain)) {
        #     # Fallback oder Konfiguration für Standarddomäne verwenden
        #     $domain = "contoso.com" # Beispiel - Dies sollte konfigurierbar sein!
        #     Write-DebugMessage "Konnte Domäne nicht automatisch ermitteln, verwende Fallback: $domain" -Type Warning
        # }
        # $userPrincipalName = "$($ResourceName -replace '\s','')@$domain" # Leerzeichen entfernen
        # Write-DebugMessage "Generierter UPN: $userPrincipalName" -Type Verbose

        $params = @{
            Name        = $ResourceName
            DisplayName = $ResourceName # Standardmäßig Name als Anzeigename
            # UserPrincipalName = $userPrincipalName # UPN explizit setzen
        }

        # Passwort wird automatisch generiert, wenn nicht angegeben.
        # Für Ressourcen ist oft kein komplexes Passwortmanagement nötig.

        if ($ResourceType -eq "Room") {
            $params.Add("Room", $true)
            Write-DebugMessage "Erstelle Raum-Postfach mit Parametern: $($params | Out-String)" -Type "Verbose"
            New-Mailbox @params -ErrorAction Stop
            $successMsg = "Raum-Ressource '$ResourceName' erfolgreich erstellt."
        }
        elseif ($ResourceType -eq "Equipment") {
            $params.Add("Equipment", $true)
            Write-DebugMessage "Erstelle Ausstattungs-Postfach mit Parametern: $($params | Out-String)" -Type "Verbose"
            New-Mailbox @params -ErrorAction Stop
            $successMsg = "Ausstattungs-Ressource '$ResourceName' erfolgreich erstellt."
        }
        else {
            # Sollte durch ValidateSet nicht passieren, aber sicher ist sicher
            throw "Ungültiger Ressourcentyp: $ResourceType"
        }

        Write-DebugMessage $successMsg -Type "Success"
        $script:txtStatus.Text = $successMsg
        Log-Action $successMsg

        # Liste aktualisieren
        Refresh-ResourcesAction # Zeigt alle Ressourcen nach Erstellung an

        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Erstellen der Ressource '$ResourceName': $errorMsg" -Type "Error"
        Log-Action "Fehler beim Erstellen der Ressource '$ResourceName': $errorMsg"
        $script:txtStatus.Text = "Fehler: $errorMsg"
        Show-MessageBox -Message "Fehler beim Erstellen der Ressource '$ResourceName': $errorMsg" -Title "Fehler" -Type "Error"
        return $false
    }
}

# Funktion zum Anzeigen von Ressourcen eines bestimmten Typs oder aller Typen
function Show-ResourcesAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Room", "Equipment", "All")]
        [string]$Type,

        [string]$SearchTerm # Optional für die Suche
    )
    Write-DebugMessage "Funktion Show-ResourcesAction aufgerufen (Type: $Type, SearchTerm: '$SearchTerm')" -Type "Info"

    # Prüfen, ob eine Exchange-Verbindung besteht
    if (-not $script:isConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type "Warning"
        return $null
    }

    try {
        $filterParts = @()
        $statusText = ""

        # Filter basierend auf Typ bauen
        switch ($Type) {
            "Room"      { $filterParts += "(RecipientTypeDetails -eq 'RoomMailbox')"; $statusText = "Suche nach Raum-Ressourcen..." }
            "Equipment" { $filterParts += "(RecipientTypeDetails -eq 'EquipmentMailbox')"; $statusText = "Suche nach Ausstattungs-Ressourcen..." }
            "All"       { $filterParts += "((RecipientTypeDetails -eq 'RoomMailbox') -or (RecipientTypeDetails -eq 'EquipmentMailbox'))"; $statusText = "Suche nach allen Ressourcen..." }
        }

        # Filter basierend auf Suchbegriff hinzufügen
        if (-not [string]::IsNullOrWhiteSpace($SearchTerm)) {
            $escapedSearchTerm = $SearchTerm -replace "'","''" # Escape single quotes for OPATH
            $filterParts += "((Name -like '*$escapedSearchTerm*') -or (DisplayName -like '*$escapedSearchTerm*'))"
            $statusText = "Suche nach Ressourcen mit '$SearchTerm'..."
        }

        # Filter kombinieren
        $filter = $filterParts -join " -and "
        $script:txtStatus.Text = $statusText
        Write-DebugMessage "Verwende Filter: $filter" -Type "Verbose"

        # Ressourcen abrufen
        $resources = Get-Mailbox -Filter $filter -ResultSize Unlimited -ErrorAction Stop | Sort-Object Name
        $displayList = $resources | ForEach-Object { Format-ResourceForDisplay -Mailbox $_ }

        Update-ResourceDataGrid -ResourceList $displayList
        $count = $displayList.Count
        $statusResultText = "$count Ressource(n)"
        if (-not [string]::IsNullOrWhiteSpace($SearchTerm)) {
            $statusResultText += " für '$SearchTerm'"
        } elseif ($Type -ne 'All') {
             $statusResultText += " vom Typ '$Type'"
        }
        $statusResultText += " gefunden."

        $script:txtStatus.Text = $statusResultText
        Write-DebugMessage "$count Ressource(n) erfolgreich abgerufen." -Type "Success"
        return $displayList
    }
    catch {
        $errorMsg = $_.Exception.Message
        $actionDescription = "beim Abrufen der Ressourcen"
        if (-not [string]::IsNullOrWhiteSpace($SearchTerm)) { $actionDescription += " für '$SearchTerm'" }
        if ($Type -ne 'All') { $actionDescription += " vom Typ '$Type'" }

        Write-DebugMessage "Fehler $($actionDescription): $errorMsg" -Type "Error"        
        Log-Action "Fehler $($actionDescription): $errorMsg"
                if($script:txtStatus -ne $null) {
            $script:txtStatus.Text = "Fehler: $errorMsg"
        }
        Show-MessageBox -Message "Fehler $($actionDescription): $errorMsg" -Title "Fehler" -Type "Error"
        Update-ResourceDataGrid -ResourceList @() # Leere Liste anzeigen
        return $null
    }
}


# Funktion zum Anzeigen aller Raum-Ressourcen (Wrapper für Show-ResourcesAction)
# Funktion zum Anzeigen von Raum-Ressourcen
function Show-RoomResourcesAction {
    [CmdletBinding()]
    param()

    Write-DebugMessage "Funktion Show-RoomResourcesAction aufgerufen" -Type "Info"
    if (-not $script:isConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type "Warning"
        return
    }

    try {
        $script:txtStatus.Text = "Lade Raum-Ressourcen..."
        $resources = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, ResourceCapacity, Identity
        Write-DebugMessage "Anzahl gefundener Raum-Ressourcen: $($resources.Count)" -Type "Info"

        # !!! KORREKTUR: DataGrid leeren, bevor ItemsSource neu gesetzt wird !!!
        $script:dgResources.ItemsSource = $null
        $script:dgResources.Items.Clear()
        $script:dgResources.ItemsSource = @($resources)

        $script:txtStatus.Text = "$($resources.Count) Raum-Ressourcen geladen."
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Laden der Raum-Ressourcen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Laden der Raum-Ressourcen: $errorMsg"
        $script:txtStatus.Text = "Fehler: $errorMsg"
    }
}

# Funktion zum Anzeigen aller Ausstattungs-Ressourcen (Wrapper für Show-ResourcesAction)
# Funktion zum Anzeigen von Ausstattungs-Ressourcen
function Show-EquipmentResourcesAction {
    [CmdletBinding()]
    param()

    Write-DebugMessage "Funktion Show-EquipmentResourcesAction aufgerufen" -Type "Info"
    if (-not $script:isConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type "Warning"
        return
    }

    try {
        $script:txtStatus.Text = "Lade Ausstattungs-Ressourcen..."
        $resources = Get-Mailbox -RecipientTypeDetails EquipmentMailbox -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, ResourceCapacity, Identity
        Write-DebugMessage "Anzahl gefundener Ausstattungs-Ressourcen: $($resources.Count)" -Type "Info"

        # !!! KORREKTUR: DataGrid leeren, bevor ItemsSource neu gesetzt wird !!!
        $script:dgResources.ItemsSource = $null
        $script:dgResources.Items.Clear()
        $script:dgResources.ItemsSource = @($resources)

        $script:txtStatus.Text = "$($resources.Count) Ausstattungs-Ressourcen geladen."
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Laden der Ausstattungs-Ressourcen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Laden der Ausstattungs-Ressourcen: $errorMsg"
        $script:txtStatus.Text = "Fehler: $errorMsg"
    }
}

# Funktion zum Suchen von Ressourcen
function Search-ResourcesAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm
    )

    Write-DebugMessage "Funktion Search-ResourcesAction aufgerufen mit Suchbegriff: '$SearchTerm'" -Type "Info"
    if (-not $script:isConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type "Warning"
        return
    }

    try {
        $script:txtStatus.Text = "Suche Ressourcen mit '$SearchTerm'..."
        # Suche nach Name oder E-Mail - beide Typen berücksichtigen
        $filter = "(RecipientTypeDetails -eq 'RoomMailbox' -or RecipientTypeDetails -eq 'EquipmentMailbox') -and (DisplayName -like '*$SearchTerm*' -or PrimarySmtpAddress -like '*$SearchTerm*')"
        $resources = Get-Mailbox -Filter $filter -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, ResourceCapacity, Identity
        Write-DebugMessage "Anzahl gefundener Ressourcen für '$SearchTerm': $($resources.Count)" -Type "Info"

        # !!! KORREKTUR: DataGrid leeren, bevor ItemsSource neu gesetzt wird !!!
        $script:dgResources.ItemsSource = $null
        $script:dgResources.Items.Clear()
        $script:dgResources.ItemsSource = @($resources)

        $script:txtStatus.Text = "$($resources.Count) Ressourcen für '$SearchTerm' gefunden."
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Suchen der Ressourcen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Suchen der Ressourcen mit '$SearchTerm': $errorMsg"
        $script:txtStatus.Text = "Fehler: $errorMsg"
    }
}

# Funktion zum Anzeigen ALLER Ressourcen (Aktualisieren)
function Refresh-ResourcesAction {
    [CmdletBinding()]
    param()

    Write-DebugMessage "Funktion Refresh-ResourcesAction aufgerufen" -Type "Info"
    if (-not $script:isConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type "Warning"
        return
    }

    try {
        $script:txtStatus.Text = "Lade alle Ressourcen..."
        $resources = Get-Mailbox -Filter "RecipientTypeDetails -eq 'RoomMailbox' -or RecipientTypeDetails -eq 'EquipmentMailbox'" -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, ResourceCapacity, Identity
        Write-DebugMessage "Anzahl gefundener Ressourcen: $($resources.Count)" -Type "Info"

        # !!! KORREKTUR: DataGrid leeren, bevor ItemsSource neu gesetzt wird !!!
        $script:dgResources.ItemsSource = $null
        $script:dgResources.Items.Clear()
        $script:dgResources.ItemsSource = @($resources)

        # !!! KORREKTUR: ComboBox cmbResourceSelect füllen und leeren !!!
        $script:cmbResourceSelect.ItemsSource = $null
        $script:cmbResourceSelect.Items.Clear()
        # Filtere nur die Namen für die Auswahlbox
        $resourceNames = $resources | Select-Object -ExpandProperty DisplayName
        $script:cmbResourceSelect.ItemsSource = @($resourceNames)
        if ($resourceNames.Count -gt 0) {
            $script:cmbResourceSelect.SelectedIndex = 0 # Erstes Element auswählen
        }


        $script:txtStatus.Text = "$($resources.Count) Ressourcen geladen."
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Laden aller Ressourcen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Laden aller Ressourcen: $errorMsg"
        $script:txtStatus.Text = "Fehler: $errorMsg"
    }
}

# Funktion zum Anzeigen/Bearbeiten der Einstellungen einer ausgewählten Ressource
function Edit-ResourceSettingsAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$SelectedResource
    )
    Write-DebugMessage "Funktion Edit-ResourceSettingsAction aufgerufen für '$($SelectedResource.Name)'" -Type "Info"

    # Prüfen, ob eine Exchange-Verbindung besteht
    if (-not $script:isConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type "Warning"
        return $null
    }

    try {
        $identity = $SelectedResource.Identity
        $script:txtStatus.Text = "Lade Einstellungen für '$($SelectedResource.Name)'..."

        # Kalenderverarbeitungseinstellungen abrufen
        Write-DebugMessage "Rufe Get-CalendarProcessing für '$identity' auf" -Type Verbose
        $calendarSettings = Get-CalendarProcessing -Identity $identity -ErrorAction Stop

        # Mailbox-Einstellungen abrufen (Beispiel: Kontingente)
        Write-DebugMessage "Rufe Get-Mailbox für '$identity' auf (für Kontingente)" -Type Verbose
        $mailboxSettings = Get-Mailbox -Identity $identity | Select-Object ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota

        # TODO: Hier einen dedizierten Dialog implementieren, um die Einstellungen anzuzeigen UND zu bearbeiten.
        # Der Dialog sollte $calendarSettings und $mailboxSettings entgegennehmen und geänderte Werte zurückgeben.
        # Beispiel:
        # $editDialog = New-ResourceSettingsDialog -CalendarSettings $calendarSettings -MailboxSettings $mailboxSettings
        # $dialogResult = $editDialog.ShowDialog()
        # if ($dialogResult -eq $true) {
        #    $updatedCalendarSettings = $editDialog.UpdatedCalendarSettings
        #    $updatedMailboxSettings = $editDialog.UpdatedMailboxSettings
        #    # Hier Set-CalendarProcessing und Set-Mailbox mit den neuen Werten aufrufen
        #    # Update-ResourceSettings -Identity $identity -CalendarSettings $updatedCalendarSettings -MailboxSettings $updatedMailboxSettings
        #    $script:txtStatus.Text = "Einstellungen für '$($SelectedResource.Name)' aktualisiert."
        # } else {
        #    $script:txtStatus.Text = "Bearbeitung der Einstellungen für '$($SelectedResource.Name)' abgebrochen."
        # }

        # Temporäre Anzeige der Einstellungen in einer MessageBox
        $settingsInfo = @"
Ressource: $($SelectedResource.Name) ($($SelectedResource.ResourceType))
E-Mail: $($SelectedResource.PrimarySmtpAddress)

Kalenderverarbeitung (Auszug):
AutomateProcessing: $($calendarSettings.AutomateProcessing)
AllowConflicts: $($calendarSettings.AllowConflicts)
BookingWindowInDays: $($calendarSettings.BookingWindowInDays)
MaximumDurationInMinutes: $($calendarSettings.MaximumDurationInMinutes)
AllowRecurringMeetings: $($calendarSettings.AllowRecurringMeetings)
TentativePendingApproval: $($calendarSettings.TentativePendingApproval)
(Weitere Einstellungen über Get-CalendarProcessing verfügbar)

Postfachkontingente:
Warnung bei: $($mailboxSettings.IssueWarningQuota)
Senden verbieten bei: $($mailboxSettings.ProhibitSendQuota)
Senden/Empfangen verbieten bei: $($mailboxSettings.ProhibitSendReceiveQuota)
"@

        Write-DebugMessage "Einstellungen für '$($SelectedResource.Name)' erfolgreich abgerufen." -Type "Success"
        $script:txtStatus.Text = "Einstellungen für '$($SelectedResource.Name)' geladen. (Nur Anzeige)"
        # Zeige die Informationen in einem einfachen Dialog an
        Show-MessageBox -Message $settingsInfo -Title "Einstellungen für $($SelectedResource.Name) (Nur Anzeige)" -Type "Info"

        return $calendarSettings # Gibt vorerst nur die Kalendereinstellungen zurück
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Ressourceneinstellungen für '$($SelectedResource.Name)': $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Ressourceneinstellungen für '$($SelectedResource.Name)': $errorMsg"
        $script:txtStatus.Text = "Fehler: $errorMsg"
        Show-MessageBox -Message "Fehler beim Abrufen der Einstellungen für '$($SelectedResource.Name)': $errorMsg" -Title "Fehler" -Type "Error"
        return $null
    }
}

# Funktion zum Entfernen einer ausgewählten Ressource
function Remove-ResourceAction {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')] # ShouldProcess und ConfirmImpact hinzugefügt
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$SelectedResource
    )
    Write-DebugMessage "Funktion Remove-ResourceAction aufgerufen für '$($SelectedResource.Name)'" -Type "Info"

    # Prüfen, ob eine Exchange-Verbindung besteht
    if (-not $script:isConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type "Warning"
        return $false
    }

    try {
        $identity = $SelectedResource.Identity
        $resourceName = $SelectedResource.Name

        # ShouldProcess-Prüfung (ersetzt die manuelle MessageBox-Abfrage, wenn -Confirm nicht verwendet wird)
        if (-not $pscmdlet.ShouldProcess("Ressource '$resourceName' (Identity: $identity)", "Entfernen")) {
             Write-DebugMessage "Entfernen von '$resourceName' durch ShouldProcess abgebrochen." -Type "Info"
             $script:txtStatus.Text = "Entfernen abgebrochen."
             return $false
        }

        # Manuelle Bestätigung (zusätzlich oder alternativ zu -Confirm / ShouldProcess, je nach Präferenz)
        # $confirmResult = [System.Windows.MessageBox]::Show(
        #     "Möchten Sie die Ressource '$resourceName' wirklich dauerhaft entfernen?",
        #     "Ressource entfernen bestätigen",
        #     [System.Windows.MessageBoxButton]::YesNo,
        #     [System.Windows.MessageBoxImage]::Warning
        # )
        # if ($confirmResult -ne [System.Windows.MessageBoxResult]::Yes) {
        #     Write-DebugMessage "Entfernen von '$resourceName' durch Benutzer abgebrochen." -Type "Info"
        #     $script:txtStatus.Text = "Entfernen abgebrochen."
        #     return $false
        # }

        $script:txtStatus.Text = "Entferne Ressource '$resourceName'..."
        Write-DebugMessage "Entferne Postfach mit Identity: $identity" -Type "Verbose"
        # -Confirm wird durch SupportsShouldProcess und ConfirmImpact='High' gesteuert
        # Wenn das Skript mit -Confirm:$false aufgerufen wird, erfolgt keine Abfrage.
        # Wenn das Skript normal aufgerufen wird, erfolgt die PowerShell Standard-Confirm-Abfrage.
        Remove-Mailbox -Identity $identity -ErrorAction Stop

        $successMsg = "Ressource '$resourceName' erfolgreich entfernt."
        Write-DebugMessage $successMsg -Type "Success"
        $script:txtStatus.Text = $successMsg
        Log-Action $successMsg

        # Liste aktualisieren
        Refresh-ResourcesAction

        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen der Ressource '$($SelectedResource.Name)': $errorMsg" -Type "Error"
        Log-Action "Fehler beim Entfernen der Ressource '$($SelectedResource.Name)': $errorMsg"
        $script:txtStatus.Text = "Fehler: $errorMsg"
        Show-MessageBox -Message "Fehler beim Entfernen der Ressource '$($SelectedResource.Name)': $errorMsg" -Title "Fehler" -Type "Error"
        return $false
    }
}

# Funktion zum Exportieren der aktuell im DataGrid angezeigten Ressourcenliste
function Export-ResourcesAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IList]$ResourceList # Die aktuelle Liste aus dem DataGrid
    )
    Write-DebugMessage "Funktion Export-ResourcesAction aufgerufen" -Type "Info"

    if ($null -eq $ResourceList -or $ResourceList.Count -eq 0) {
        Show-MessageBox -Message "Es gibt keine Daten zum Exportieren." -Title "Export nicht möglich" -Type "Info"
        return $false
    }

    try {
        # SaveFileDialog initialisieren
        $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveFileDialog.Filter = "CSV (Semikolon-getrennt)|*.csv|Textdatei (Formatierte Liste)|*.txt"
        $saveFileDialog.Title = "Ressourcenliste exportieren"
        $defaultFileName = "ExchangeResources_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $saveFileDialog.FileName = $defaultFileName
        $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments') # Start im Dokumente-Ordner

        # Dialog anzeigen
        if ($saveFileDialog.ShowDialog() -ne $true) {
            Write-DebugMessage "Export durch Benutzer abgebrochen." -Type "Info"
            $script:txtStatus.Text = "Export abgebrochen."
            return $false
        }

        $exportPath = $saveFileDialog.FileName
        $fileExtension = [System.IO.Path]::GetExtension($exportPath).ToLower()

        $script:txtStatus.Text = "Exportiere Ressourcenliste nach '$exportPath'..."

        # Daten für den Export vorbereiten (nur relevante Spalten aus der übergebenen Liste)
        # Die Objekte in ResourceList sollten bereits die benötigten Eigenschaften haben (aus Format-ResourceForDisplay)
        $exportData = $ResourceList | Select-Object Name, DisplayName, PrimarySmtpAddress, ResourceType

        if ($fileExtension -eq ".csv") {
            # Als CSV exportieren
            Write-DebugMessage "Exportiere als CSV nach '$exportPath'" -Type Verbose
            $exportData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8 -Delimiter ";" -ErrorAction Stop
            $exportType = "CSV"
        }
        elseif ($fileExtension -eq ".txt") {
            # Als Text exportieren (Format-List)
            Write-DebugMessage "Exportiere als Text (Format-List) nach '$exportPath'" -Type Verbose
            ($exportData | Format-List | Out-String) | Out-File -FilePath $exportPath -Encoding utf8 -ErrorAction Stop
            $exportType = "Textdatei"
        }
        else {
             # Sollte durch den Filter nicht passieren, aber zur Sicherheit
             throw "Ungültige Dateierweiterung für Export: $fileExtension"
        }


        $successMsg = "$($ResourceList.Count) Ressource(n) erfolgreich als $exportType nach '$exportPath' exportiert."
        Write-DebugMessage $successMsg -Type "Success"
        $script:txtStatus.Text = $successMsg
        Log-Action $successMsg
        Show-MessageBox -Message $successMsg -Title "Export erfolgreich" -Type "Info"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Exportieren der Ressourcenliste: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Exportieren der Ressourcenliste: $errorMsg"
        $script:txtStatus.Text = "Fehler: $errorMsg"
        Show-MessageBox -Message "Fehler beim Exportieren der Ressourcenliste: $errorMsg`nPfad: $exportPath" -Title "Fehler" -Type "Error"
        return $false
    }
}

# Hilfsfunktion zum Formatieren eines Mailbox-Objekts für die Anzeige im DataGrid
function Format-ResourceForDisplay {
    param(
        [Parameter(Mandatory = $true)]
        # Typisierung lockern, da Get-Recipient auch andere Typen liefern könnte, obwohl wir filtern
        [PSObject]$Mailbox # War [Microsoft.Exchange.Data.Directory.Management.Mailbox]
    )
    # Sicherstellen, dass die erwarteten Eigenschaften vorhanden sind
    $name = $Mailbox.PSObject.Properties['Name'].Value
    $displayName = $Mailbox.PSObject.Properties['DisplayName'].Value
    $primarySmtpAddress = $Mailbox.PSObject.Properties['PrimarySmtpAddress'].Value
    $recipientTypeDetails = $Mailbox.PSObject.Properties['RecipientTypeDetails'].Value
    $identity = $Mailbox.PSObject.Properties['Identity'].Value

    return [PSCustomObject]@{
        Name               = if ($null -ne $name) { $name } else { $Mailbox.Alias } # Fallback auf Alias
        DisplayName        = if ($null -ne $displayName) { $displayName } else { $name } # Fallback auf Name
        PrimarySmtpAddress = if ($null -ne $primarySmtpAddress) { $primarySmtpAddress.ToString() } else { "N/A" }
        ResourceType       = if ($null -ne $recipientTypeDetails) { $recipientTypeDetails.ToString() } else { "Unbekannt" }
        Identity           = if ($null -ne $identity) { $identity.ToString() } else { "N/A" }
        # Hier könnten weitere relevante Eigenschaften hinzugefügt werden
    }
}

# Hilfsfunktion zum Aktualisieren des Ressourcen-DataGrids
function Update-ResourceDataGrid {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IList]$ResourceList
    )
    if ($null -ne $script:dgResources) {
        try {
            # Sicherstellen, dass der Zugriff auf das UI-Element vom richtigen Thread erfolgt
            $script:dgResources.Dispatcher.InvokeAsync({
                # ItemsSource auf $null setzen, bevor die neue Liste zugewiesen wird,
                # kann manchmal Darstellungsprobleme bei großen Änderungen verhindern.
                # $script:dgResources.ItemsSource = $null
                $script:dgResources.ItemsSource = $ResourceList
                # Items.Refresh() ist oft nicht nötig, wenn ItemsSource neu gesetzt wird,
                # kann aber bei direkter Manipulation der gebundenen Liste helfen.
                # if ($ResourceList -ne $null -and $ResourceList.Count -gt 0) {
                #     $script:dgResources.Items.Refresh()
                # }
                Write-DebugMessage "Ressourcen-DataGrid aktualisiert mit $($ResourceList.Count) Einträgen." -Type "Info"
            }) | Out-Null # InvokeAsync gibt ein Task-Objekt zurück, das wir hier nicht brauchen
        } catch {
            # Fehler im Dispatcher-Thread abfangen
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Aktualisieren des Ressourcen-DataGrids auf dem Dispatcher-Thread: $errorMsg" -Type "Error"
            # Ggf. Fehler im UI anzeigen, aber vorsichtig sein, um keine Endlosschleife auszulösen
            # $script:txtStatus.Dispatcher.InvokeAsync({ $script:txtStatus.Text = "Fehler beim Aktualisieren der Liste." }) | Out-Null
        }
    } else {
        Write-DebugMessage "Ressourcen-DataGrid (dgResources) ist null oder nicht initialisiert." -Type "Warning"
    }
}

#endregion Resources Tab Functions
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
        $lstReportCategories = Get-XamlElement -ElementName "lstReportCategories" # NEU
        $cmbReportType = Get-XamlElement -ElementName "cmbReportType"
        $dpReportStartDate = Get-XamlElement -ElementName "dpReportStartDate"
        $dpReportEndDate = Get-XamlElement -ElementName "dpReportEndDate"
        $btnGenerateReport = Get-XamlElement -ElementName "btnGenerateReport"
        $lstReportResults = Get-XamlElement -ElementName "lstReportResults"
        $btnExportReport = Get-XamlElement -ElementName "btnExportReport"
        $helpLinkReports = Get-XamlElement -ElementName "helpLinkReports"

        # Globale Variablen setzen
        $script:lstReportCategories = $lstReportCategories # NEU
        $script:cmbReportType = $cmbReportType
        $script:dpReportStartDate = $dpReportStartDate
        $script:dpReportEndDate = $dpReportEndDate
        $script:lstReportResults = $lstReportResults

        # --- Start: Logik für Report-Kategorien und Typen --- NEUER ABSCHNITT ---

        # Definition der Berichte pro Kategorie (als Script-Variable speichern)
        $script:ReportDefinitions = @{
            Postfaecher = @(
                "Alle Postfachgrößen",
                "Letzte Anmeldung (Alle)",
                "Inaktive Postfächer (90 Tage)",
                "Postfächer über Größenlimit"
            )
            Berechtigungen = @(
                "Postfachberechtigungen (Alle)",
                "Kalenderberechtigungen (Alle)",
                "SendAs-Berechtigungen (Alle)",
                "Freigegebene Postfach-Berechtigungen (Alle)"
            )
            Gruppen = @(
                "Alle Gruppenmitglieder",
                "Leere Gruppen",
                "Verwaiste Gruppen (ohne Besitzer)"
            )
            Ressourcen = @(
                "Alle Raum-Postfächer",
                "Alle Geräte-Postfächer"
            )
            # Weitere Kategorien und Berichte hier hinzufügen
        }

        # Kategorien in die ListBox laden
        if ($null -ne $script:lstReportCategories) {
            $script:lstReportCategories.Items.Clear()
            foreach ($categoryKey in $script:ReportDefinitions.Keys) {
                $item = New-Object System.Windows.Controls.ListBoxItem
                $item.Content = $categoryKey # Angezeigter Text
                $item.Tag = $categoryKey     # Interner Schlüssel für Handler
                [void]$script:lstReportCategories.Items.Add($item)
            }
        } else {
             Write-DebugMessage "lstReportCategories konnte nicht gefunden werden." -Type Warning
        }

        # Event-Handler für die Auswahl in der Kategorie-Liste
        Register-EventHandler -Control $script:lstReportCategories -EventName SelectionChanged -Handler {
            param($sender, $e) # Parameter für Event Handler

            try {
                $selectedItem = $script:lstReportCategories.SelectedItem
                if ($null -eq $selectedItem) {
                    if ($null -ne $script:cmbReportType) {
                         $script:cmbReportType.ItemsSource = $null
                    }
                    return
                }

                $categoryKey = $selectedItem.Tag.ToString()

                if ($script:ReportDefinitions.ContainsKey($categoryKey)) {
                    $reportsForCategory = $script:ReportDefinitions[$categoryKey]
                    if ($null -ne $script:cmbReportType) {
                        $script:cmbReportType.Items.Clear()
                        $script:cmbReportType.ItemsSource = $reportsForCategory
                        if ($script:cmbReportType.Items.Count -gt 0) {
                            $script:cmbReportType.SelectedIndex = 0
                        } else {
                             $script:cmbReportType.ItemsSource = $null
                        }
                        Write-DebugMessage "Berichtstypen für Kategorie '$categoryKey' geladen." -Type Info
                    }
                } else {
                     Write-DebugMessage "Keine Berichtsdefinitionen für Kategorie '$categoryKey' gefunden." -Type Warning
                     if ($null -ne $script:cmbReportType) {
                         $script:cmbReportType.ItemsSource = $null
                     }
                }
            } catch {
                $errorMsg = $_.Exception.Message
                Write-DebugMessage "Fehler im SelectionChanged-Handler für lstReportCategories: $errorMsg" -Type Error
                # $script:txtStatus.Text = "Fehler beim Laden der Berichtstypen."
            }
        } -ControlName "lstReportCategories"

        # Initial ersten Eintrag auswählen (nachdem Handler registriert wurde)
        if ($null -ne $script:lstReportCategories -and $script:lstReportCategories.Items.Count -gt 0) {
            $script:lstReportCategories.SelectedIndex = 0 # Löst den Handler aus und füllt die ComboBox
        }

        # --- Ende: Logik für Report-Kategorien und Typen ---

        # Event-Handler für Buttons registrieren (bestehender Code)
        Register-EventHandler -Control $btnGenerateReport -Handler {
            try {
                # Verbindungsprüfung
                if (-not $script:isConnected) {
                    [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her.", "Keine Verbindung",
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }

                # Parameter sammeln
                $reportType = $script:cmbReportType.SelectedItem # Nicht .Text, da ItemsSource verwendet wird
                if ($null -eq $reportType) {
                    [System.Windows.MessageBox]::Show("Bitte wählen Sie einen Berichtstyp aus.",
                        "Keine Auswahl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }

                $reportName = $reportType.ToString() # Der ausgewählte String aus der Liste

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

                # ===== HIER MUSS DIE SWITCH-ANWEISUNG ANGEPASST WERDEN =====
                # Die Logik muss die NEUEN Berichtsnamen aus $script:ReportDefinitions verwenden
                # Beispiel:
                switch -Wildcard ($reportName) {
                    "Alle Postfachgrößen" {
                        # Bestehende Logik für Postfachgrößen
                        $reportData = Get-Mailbox -ResultSize Unlimited |
                            Get-MailboxStatistics |
                            Select-Object DisplayName, TotalItemSize, ItemCount, LastLogonTime, @{
                                Name = "TotalSizeGB"
                                Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1GB), 2) }
                            }
                    }
                    "Letzte Anmeldung (Alle)" {
                         # Bestehende Logik für letzte Anmeldung
                         $reportData = Get-Mailbox -ResultSize Unlimited |
                            Get-MailboxStatistics |
                            Select-Object DisplayName, LastLogonTime, ItemCount
                    }
                    "Inaktive Postfächer (90 Tage)" {
                         # NEUE LOGIK erforderlich
                         $inactiveThreshold = (Get-Date).AddDays(-90)
                         $reportData = Get-Mailbox -ResultSize Unlimited |
                            Get-MailboxStatistics |
                            Where-Object { $_.LastLogonTime -lt $inactiveThreshold -or $null -eq $_.LastLogonTime } |
                            Select-Object DisplayName, LastLogonTime, ItemCount, TotalItemSize
                    }
                    "Postfächer über Größenlimit" {
                         # NEUE LOGIK erforderlich (Beispiel: Limit 50 GB)
                         $sizeLimitBytes = 50GB
                         $reportData = Get-Mailbox -ResultSize Unlimited |
                             Get-MailboxStatistics |
                             Where-Object { $_.TotalItemSize.Value.ToBytes() -gt $sizeLimitBytes } |
                             Select-Object DisplayName, @{Name="SizeGB";Expression={[math]::Round($_.TotalItemSize.Value.ToGB(), 2)}}, ItemCount, LastLogonTime
                    }
                    "Postfachberechtigungen (Alle)" {
                         # Bestehende Logik für Postfachberechtigungen
                        $reportData = @()
                        $mailboxes = Get-Mailbox -ResultSize Unlimited
                        foreach ($mailbox in $mailboxes) {
                            $permissions = Get-MailboxPermission -Identity $mailbox.Identity |
                                Where-Object { $_.User -notlike "NT AUTHORITY\\*" -and $_.IsInherited -eq $false }
                            foreach ($perm in $permissions) {
                                $reportData += [PSCustomObject]@{
                                    Mailbox      = $mailbox.DisplayName
                                    MailboxEmail = $mailbox.PrimarySmtpAddress
                                    User         = $perm.User
                                    AccessRights = ($perm.AccessRights -join ", ")
                                }
                            }
                        }
                    }
                    "Kalenderberechtigungen (Alle)" {
                         # Bestehende Logik für Kalenderberechtigungen
                        $reportData = @()
                        $mailboxes = Get-Mailbox -ResultSize Unlimited
                        foreach ($mailbox in $mailboxes) {
                            try {
                                $calendarFolder = $mailbox.PrimarySmtpAddress.ToString() + ":\\Kalender"
                                $permissions = Get-MailboxFolderPermission -Identity $calendarFolder -ErrorAction SilentlyContinue
                                if ($null -eq $permissions) {
                                    $calendarFolder = $mailbox.PrimarySmtpAddress.ToString() + ":\\Calendar"
                                    $permissions = Get-MailboxFolderPermission -Identity $calendarFolder -ErrorAction SilentlyContinue
                                }
                                if ($null -ne $permissions) {
                                    foreach ($perm in $permissions) {
                                        if ($perm.User.ToString() -ne "Anonymous" -and $perm.User.ToString() -ne "Default") {
                                            $reportData += [PSCustomObject]@{
                                                Mailbox      = $mailbox.DisplayName
                                                MailboxEmail = $mailbox.PrimarySmtpAddress
                                                User         = $perm.User
                                                AccessRights = ($perm.AccessRights -join ", ")
                                            }
                                        }
                                    }
                                }
                            } catch { continue }
                        }
                    }
                    "SendAs-Berechtigungen (Alle)" {
                        # NEUE LOGIK erforderlich
                        $reportData = @()
                        $mailboxes = Get-Mailbox -ResultSize Unlimited
                        foreach ($mbx in $mailboxes) {
                            $sendAsPerms = Get-RecipientPermission -Identity $mbx.Identity | Where-Object {$_.Trustee -ne $null -and $_.Trustee -notlike "NT AUTHORITY\\*"}
                            foreach($perm in $sendAsPerms){
                                $reportData += [PSCustomObject]@{
                                    Mailbox = $mbx.DisplayName
                                    Trustee = $perm.Trustee
                                    IsAllowed = $perm.IsAllowed
                                }
                            }
                        }
                    }
                     "Freigegebene Postfach-Berechtigungen (Alle)" {
                        # Bestehende Logik für Shared Mailboxes (leicht angepasst)
                        $reportData = @()
                        $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
                        foreach ($mailbox in $sharedMailboxes) {
                            # FullAccess
                            $permissions = Get-MailboxPermission -Identity $mailbox.Identity |
                                Where-Object { $_.User -notlike "NT AUTHORITY\\*" -and $_.IsInherited -eq $false -and $_.AccessRights -contains "FullAccess"}
                            foreach ($perm in $permissions) {
                                $reportData += [PSCustomObject]@{
                                    SharedMailbox = $mailbox.DisplayName
                                    User          = $perm.User
                                    PermissionType= "FullAccess"
                                    AccessRights  = ($perm.AccessRights -join ", ")
                                    AutoMapping   = $perm.AutoMapping # AutoMapping hier verfügbar
                                }
                            }
                            # SendAs
                             $sendAsPerms = Get-RecipientPermission -Identity $mailbox.Identity | Where-Object {$_.Trustee -ne $null -and $_.Trustee -notlike "NT AUTHORITY\\*"}
                             foreach($perm in $sendAsPerms){
                                $reportData += [PSCustomObject]@{
                                    SharedMailbox = $mailbox.DisplayName
                                    User          = $perm.Trustee
                                    PermissionType= "SendAs"
                                    AccessRights  = "SendAs"
                                    AutoMapping   = "N/A"
                                }
                            }
                             # SendOnBehalf (falls benötigt)
                             # $sendOnBehalf = Get-Mailbox -Identity $mailbox.Identity | Select-Object -ExpandProperty GrantSendOnBehalfTo
                             # foreach ($delegate in $sendOnBehalf) { ... }
                        }
                    }
                    "Alle Gruppenmitglieder" {
                         # Bestehende Logik für Gruppenmitglieder
                        $reportData = @()
                        $groups = Get-DistributionGroup -ResultSize Unlimited
                        $groups += Get-UnifiedGroup -ResultSize Unlimited # Auch M365 Gruppen berücksichtigen
                        foreach ($group in $groups) {
                            try {
                                $members = Get-DistributionGroupMember -Identity $group.Identity -ErrorAction SilentlyContinue
                                if($null -eq $members -and $group.RecipientTypeDetails -eq "GroupMailbox"){ # M365 Gruppe
                                    $members = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members -ResultSize Unlimited
                                }

                                foreach ($member in $members) {
                                    $reportData += [PSCustomObject]@{
                                        Gruppe        = $group.DisplayName
                                        GruppenEmail  = $group.PrimarySmtpAddress
                                        Mitglied      = $member.Name # Name statt DisplayName für Konsistenz
                                        MitgliedEmail = $member.PrimarySmtpAddress # Nicht immer verfügbar
                                        MitgliedTyp   = if ($member.RecipientTypeDetails) {$member.RecipientTypeDetails} else {$member.RecipientType}
                                    }
                                }
                            } catch {
                                 Write-DebugMessage "Fehler beim Abrufen der Mitglieder für Gruppe $($group.DisplayName): $($_.Exception.Message)" -Type Warning
                                 continue
                            }
                        }
                    }
                    "Leere Gruppen" {
                         # NEUE LOGIK erforderlich
                         $reportData = @()
                         $groups = Get-DistributionGroup -ResultSize Unlimited
                         $groups += Get-UnifiedGroup -ResultSize Unlimited
                         foreach ($group in $groups) {
                             try {
                                 $memberCount = (Get-DistributionGroupMember -Identity $group.Identity -ResultSize 1 -ErrorAction SilentlyContinue).Count
                                 if($group.RecipientTypeDetails -eq "GroupMailbox"){ # M365 Gruppe
                                     $memberCount = (Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members -ResultSize 1).Count
                                 }

                                 if ($memberCount -eq 0) {
                                     $reportData += [PSCustomObject]@{
                                         Gruppe        = $group.DisplayName
                                         GruppenEmail  = $group.PrimarySmtpAddress
                                         GruppenTyp    = $group.RecipientTypeDetails
                                         ErstelltAm    = $group.WhenCreated
                                     }
                                 }
                             } catch { continue }
                         }
                    }
                     "Verwaiste Gruppen (ohne Besitzer)" {
                         # NEUE LOGIK erforderlich (nur für M365 Gruppen relevant)
                         $reportData = @()
                         $groups = Get-UnifiedGroup -ResultSize Unlimited
                         foreach ($group in $groups) {
                            try {
                                $owners = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Owners -ResultSize 1
                                if ($owners.Count -eq 0) {
                                     $reportData += [PSCustomObject]@{
                                         Gruppe        = $group.DisplayName
                                         GruppenEmail  = $group.PrimarySmtpAddress
                                         ErstelltAm    = $group.WhenCreated
                                     }
                                }
                            } catch { continue }
                         }
                     }
                    "Alle Raum-Postfächer" {
                        # NEUE LOGIK erforderlich
                        $reportData = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
                            Select-Object DisplayName, PrimarySmtpAddress, ResourceCapacity, WhenCreated
                    }
                    "Alle Geräte-Postfächer" {
                         # NEUE LOGIK erforderlich
                         $reportData = Get-Mailbox -RecipientTypeDetails EquipmentMailbox -ResultSize Unlimited |
                            Select-Object DisplayName, PrimarySmtpAddress, WhenCreated
                    }
                    # --- Füge hier Cases für weitere Berichte hinzu ---
                    default {
                        [System.Windows.MessageBox]::Show("Der gewählte Berichtstyp '$reportName' ist noch nicht implementiert.",
                            "Nicht implementiert", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                         $script:txtStatus.Text = "Bereit." # Status zurücksetzen
                        return
                    }
                }
                # ===== Ende der Anpassung der SWITCH-Anweisung =====


                # Ergebnisse in der DataGrid anzeigen
                if ($null -ne $script:lstReportResults) {
                    $script:lstReportResults.ItemsSource = $reportData
                }

                # Status aktualisieren
                $count = 0
                if ($null -ne $reportData) { $count = $reportData.Count }
                $script:txtStatus.Text = "Bericht '$reportName' generiert: $count Datensätze gefunden."

            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-DebugMessage "Fehler beim Generieren des Berichts '$reportName': $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
                [System.Windows.MessageBox]::Show("Fehler beim Generieren des Berichts: $errorMsg",
                    "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                 # Optional: Cursor zurücksetzen, falls er auf Warten gesetzt wurde
                 # $script:MainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
                 Write-DebugMessage "Berichtsgenerierung für '$reportName' abgeschlossen." -Type Info
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
                $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*" # Filter erweitert
                $saveFileDialog.Title = "Bericht exportieren"
                # Versuche, den Berichtsnamen im Dateinamen zu verwenden
                $selectedReportName = "UnbekannterBericht"
                if($null -ne $script:cmbReportType.SelectedItem) {
                    $selectedReportName = $script:cmbReportType.SelectedItem.ToString() -replace '[^a-zA-Z0-9_]', '_' # Sonderzeichen ersetzen
                }
                $saveFileDialog.FileName = "easyEXO_Bericht_$($selectedReportName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

                # Dialog anzeigen
                if ($saveFileDialog.ShowDialog() -eq $true) {
                    $filePath = $saveFileDialog.FileName
                    $script:txtStatus.Text = "Exportiere Bericht nach '$filePath'..."
                    # Exportiere die Daten nach CSV
                    $data | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Delimiter ';' # UTF8 und Semikolon für Excel
                    $script:txtStatus.Text = "Bericht erfolgreich nach '$filePath' exportiert."
                    Write-DebugMessage "Bericht erfolgreich nach '$filePath' exportiert." -Type Success

                    # Optional: Datei nach Export öffnen?
                    # if ([System.Windows.MessageBox]::Show("Möchten Sie die exportierte Datei jetzt öffnen?", "Export abgeschlossen", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question) -eq 'Yes') {
                    #    Invoke-Item $filePath
                    # }
                } else {
                    $script:txtStatus.Text = "Export abgebrochen."
                    Write-DebugMessage "Berichtsexport abgebrochen." -Type Info
                }

            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-DebugMessage "Fehler beim Exportieren des Berichts: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler beim Export: $errorMsg"
                [System.Windows.MessageBox]::Show("Fehler beim Exportieren des Berichts: $errorMsg",
                    "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                 Write-DebugMessage "Berichtsexport abgeschlossen." -Type Info
            }
        } -ControlName "btnExportReport"

        # HelpLink Handler (angepasst für MouseLeftButtonDown)
        if ($null -ne $helpLinkReports) {
            # Optional: Sicherstellen, dass der Cursor auf Hand geändert wird, wenn man drüber fährt (falls nicht im XAML)
            # $helpLinkReports.Cursor = [System.Windows.Input.Cursors]::Hand

            Register-EventHandler -Control $helpLinkReports -EventName MouseLeftButtonDown -Handler { # <<< Event geändert
                param($sender, $e)
                try {
                    # --- HIER Logik zum Öffnen des benutzerdefinierten Hilfe-Fensters ---
                    # Annahme: Es gibt eine Funktion wie Show-HelpWindow oder Ähnliches
                    # Der Name des Tabs oder ein anderer Bezeichner könnte übergeben werden
                    $topic = "Reports" # Beispiel-Identifikator für das Thema
                    Write-DebugMessage "Öffne Hilfe-Fenster für Thema: $topic" -Type Info

                    # Platzhalter für den tatsächlichen Aufruf der Funktion, die das Fenster anzeigt
                    # Beispiel: Show-CustomHelpWindow -Topic $topic
                    [System.Windows.MessageBox]::Show("Hier würde das Hilfe-Fenster für '$topic' erscheinen.", "Hilfe (Platzhalter)", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

                    $e.Handled = $true # Wichtig bei Maus-Events, um weitere Verarbeitung zu stoppen

                } catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler beim Öffnen des Hilfe-Fensters für Reports: $errorMsg" -Type Error
                    # Optional: Benutzerfeedback
                    # $script:txtStatus.Text = "Fehler beim Anzeigen der Hilfe."
                }
            } -ControlName "helpLinkReports"
       }

        Write-DebugMessage "Berichte-Tab erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Initialisieren des Berichte-Tabs: $errorMsg" -Type "Error"
        return $false
    }
} # Ende von Initialize-ReportsTab
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
    # --- Header Button Event Handlers ---
    try {
        # Versuche, die Buttons über ihre Namen zu finden
        $btnShowConnectionStatus = $script:Form.FindName("btnShowConnectionStatus")
        $btnSettings = $script:Form.FindName("btnSettings")
        $btnInfo = $script:Form.FindName("btnInfo")
        $btnClose = $script:Form.FindName("btnClose")

        # Handler für Verbindungsstatus-Button
        if ($null -ne $btnShowConnectionStatus) {
            $btnShowConnectionStatus.Add_Click({
                Write-DebugMessage "Button 'ShowConnectionStatus' geklickt." -Type Info
                # Hier später die Funktion Show-ConnectionStatus aufrufen
                if ($script:isConnected) {
                    $userName = $script:ConnectedUser
                    $tenantId = $script:ConnectedTenantId
                    [System.Windows.MessageBox]::Show("Verbunden als '$userName' mit Tenant '$tenantId'.", "Verbindungsstatus", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                } else {
                    [System.Windows.MessageBox]::Show("Sie sind aktuell nicht mit Exchange Online verbunden.", "Verbindungsstatus", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                }
            })
            Write-DebugMessage "Event-Handler für btnShowConnectionStatus hinzugefügt." -Type Info
        } else { Write-DebugMessage "Header Button 'btnShowConnectionStatus' nicht gefunden." -Type Warning }

        # Handler für Settings-Button
        if ($null -ne $btnSettings) {
            $btnSettings.Add_Click({ Show-SettingsWindow })
            Write-DebugMessage "Event-Handler für btnSettings hinzugefügt." -Type Info
        } else {
            Write-DebugMessage "Header Button 'btnSettings' nicht gefunden." -Type Warning
        }

        # Handler für Info-Button
        if ($null -ne $btnInfo) {
            $btnInfo.Add_Click({
                Write-DebugMessage "Button 'Info' geklickt." -Type Info
                # Hier später die Funktion Show-InfoDialog aufrufen
                $version = "Unbekannt"
                $appName = "easyEXO"
                try {
                     if ($null -ne $script:config -and $null -ne $script:config['General']) {
                        if ($script:config['General'].ContainsKey('Version')) { $version = $script:config['General']['Version'] }
                        if ($script:config['General'].ContainsKey('AppName')) { $appName = $script:config['General']['AppName'] }
                     }
                 } catch { Write-DebugMessage "Fehler beim Lesen der Version/AppName aus Config für Info-Dialog." -Type Warning }

                 $infoText = @"
$appName - Version $version

Einfache Verwaltung von Exchange Online Aufgaben.

Autor: PhinIT
(c) $(Get-Date -Format yyyy)
"@
                 [System.Windows.MessageBox]::Show($infoText, "Über $appName", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            })
             Write-DebugMessage "Event-Handler für btnInfo hinzugefügt." -Type Info
        } else { Write-DebugMessage "Header Button 'btnInfo' nicht gefunden." -Type Warning }

        # Handler für Close-Button
        if ($null -ne $btnClose) {
            $btnClose.Add_Click({
                Write-DebugMessage "Button 'Close' geklickt." -Type Info
                # Hier später die Funktion Close-Application aufrufen
                try {
                    if ($script:isConnected) {
                        Write-DebugMessage "Trenne Verbindung vor dem Schließen..." -Type Info
                        # Annahme: Disconnect-ExchangeOnlineSession existiert und funktioniert
                        Disconnect-ExchangeOnlineSession
                    }
                    Write-DebugMessage "Schließe Fenster..." -Type Info
                    $script:Form.Close()
                } catch {
                     Write-DebugMessage "Fehler beim Schließen: $($_.Exception.Message)" -Type Error
                     # Notfall-Schließung
                     try { $script:Form.Close() } catch {}
                }
            })
            Write-DebugMessage "Event-Handler für btnClose hinzugefügt." -Type Info
        } else { Write-DebugMessage "Header Button 'btnClose' nicht gefunden." -Type Warning }

        Write-DebugMessage "Event-Handler für Header-Buttons erfolgreich registriert." -Type Info
    } catch {
        # Fehler beim Finden der Buttons oder Registrieren der Handler
        Write-DebugMessage "Kritischer Fehler beim Registrieren der Header-Button-Handler: $($_.Exception.Message)" -Type Error
    }
    # --- Ende Header Button Event Handlers ---
# Initialisiere alle Tabs
function Initialize-AllTabs {
    [CmdletBinding()]
    param()

    try {
        Write-DebugMessage "Initialisiere alle Tabs" -Type "Info"
        $results = @{
            EXOSettings     = Initialize-EXOSettingsTab
            Calendar        = Initialize-CalendarTab
            Mailbox         = Initialize-MailboxTab
            Audit           = Initialize-AuditTab
            Troubleshooting = Initialize-TroubleshootingTab
            Groups          = Initialize-GroupsTab
            SharedMailbox   = Initialize-SharedMailboxTab
            Contacts        = Initialize-ContactsTab
            Resources       = Initialize-ResourcesTab
            Reports         = Initialize-ReportsTab
        }

        $successCount = ($results.Values | Where-Object { $_ -eq $true }).Count
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
    # Sicherstellen, dass hier ein gültiger Typ verwendet wird (war bereits korrekt)
    Write-DebugMessage "Initialize-AllTabs Ergebnis: $result" -Type "Info"

    # Hilfe-Links initialisieren
    $result = Initialize-HelpLinks
    # Sicherstellen, dass hier ein gültiger Typ verwendet wird (war bereits korrekt)
    Write-DebugMessage "Initialize-HelpLinks Ergebnis: $result" -Type "Info"

    # --- Manueller Zugriff auf EXO Setting Buttons via Struktur ---
    try {
        # Sicherstellen, dass hier ein gültiger Typ verwendet wird (war bereits korrekt)
        Write-DebugMessage "Starte manuellen Zugriff auf EXO Setting Buttons im Loaded Event..." -Type Info

        # 1. Finde das TabItem
        $tabItemEXOSettings = $script:Form.FindName("tabEXOSettings")
        if ($null -eq $tabItemEXOSettings) { throw "TabItem 'tabEXOSettings' nicht gefunden." }
        # Sicherstellen, dass hier ein gültiger Typ verwendet wird (war bereits korrekt)
        Write-DebugMessage "TabItem 'tabEXOSettings' gefunden." -Type Info

        # 2. Finde das Haupt-Grid im TabItem (Annahme: es ist das erste/einzige Kind)
        if ($tabItemEXOSettings.Content -is [System.Windows.Controls.Grid]) {
            $mainGridInTab = $tabItemEXOSettings.Content
            # Sicherstellen, dass hier ein gültiger Typ verwendet wird (war bereits korrekt)
            Write-DebugMessage "Haupt-Grid im TabItem gefunden." -Type Info

            # 3. Finde die erste GroupBox (Organisationseinstellungen) in diesem Grid (Annahme: Grid.Row="1")
            $groupBoxOrgSettings = $null
            foreach($child in $mainGridInTab.Children) {
                if ($child -is [System.Windows.Controls.GroupBox] -and [System.Windows.Controls.Grid]::GetRow($child) -eq 1) {
                    $groupBoxOrgSettings = $child
                    break
                }
            }
            if ($null -eq $groupBoxOrgSettings) { throw "GroupBox 'Organisationseinstellungen' (Grid.Row=1) nicht gefunden." }
             Write-DebugMessage "GroupBox 'Organisationseinstellungen' gefunden." -Type Info

             # 4. Finde das Grid *innerhalb* der GroupBox (Annahme: Content der GroupBox)
             if ($groupBoxOrgSettings.Content -is [System.Windows.Controls.Grid]) {
                 $gridInGroupBox = $groupBoxOrgSettings.Content
                 Write-DebugMessage "Grid in GroupBox gefunden." -Type Info

                 # 5. Finde das Grid mit den Buttons (Annahme: Grid.Row="1" in diesem Grid)
                 $buttonGrid = $null
                 foreach($childInGroupBox in $gridInGroupBox.Children) {
                    if ($childInGroupBox -is [System.Windows.Controls.Grid] -and [System.Windows.Controls.Grid]::GetRow($childInGroupBox) -eq 1) {
                        $buttonGrid = $childInGroupBox
                        break
                    }
                 }
                 if ($null -eq $buttonGrid) { throw "Grid mit Buttons (Grid.Row=1 innerhalb GroupBox-Grid) nicht gefunden." }
                 Write-DebugMessage "Button-Grid gefunden." -Type Info

                 # 6. Greife auf die Buttons über ihren Index in den Children des Button-Grids zu
                 $btnGetOrgCfg_Manual = $null
                 $btnSetOrgCfg_Manual = $null

                 foreach($button in $buttonGrid.Children) {
                    if ($button -is [System.Windows.Controls.Button]) {
                        $col = [System.Windows.Controls.Grid]::GetColumn($button)
                        if ($col -eq 1) { $btnGetOrgCfg_Manual = $button }
                        elseif ($col -eq 2) { $btnSetOrgCfg_Manual = $button }
                    }
                 }

                 if ($null -ne $btnGetOrgCfg_Manual) {
                     $btnGetOrgCfg_Manual.Add_Click({ Get-CurrentOrganizationConfig })
                     Write-DebugMessage ">>> Event-Handler für 'btnGetOrganizationConfig' (manuell) hinzugefügt." -Type Success
                 } else {
                     # Sicherstellen, dass hier ein gültiger Typ verwendet wird (war bereits korrekt)
                     Write-DebugMessage "Button in Grid.Column=1 im ButtonGrid nicht gefunden oder kein Button." -Type Warning
                 }

                 if ($null -ne $btnSetOrgCfg_Manual) {
                     $btnSetOrgCfg_Manual.Add_Click({ Set-CustomOrganizationConfig })
                      Write-DebugMessage ">>> Event-Handler für 'btnSetOrganizationConfig' (manuell) hinzugefügt." -Type Success
                 } else {
                      # Sicherstellen, dass hier ein gültiger Typ verwendet wird (war bereits korrekt)
                      Write-DebugMessage "Button in Grid.Column=2 im ButtonGrid nicht gefunden oder kein Button." -Type Warning
                 }

             } else { throw "Inhalt der GroupBox 'Organisationseinstellungen' ist kein Grid." }
        } else { throw "Inhalt des TabItems 'tabEXOSettings' ist kein Grid." }

        # --- Handler für Export-Button (der funktionierte ja mit FindName) ---
         $btnExport_Global = $script:Form.FindName("btnExportOrgConfig")
         if ($null -ne $btnExport_Global) {
             $btnExport_Global.Add_Click({ Export-OrganizationConfig })
             Write-DebugMessage ">>> Event-Handler für btnExportOrgConfig (via globaler Suche) hinzugefügt." -Type Success
         } else {
              Write-DebugMessage "Button btnExportOrgConfig global nicht gefunden (obwohl es vorher ging?)." -Type Warning
         }

    } catch {
        Write-DebugMessage "Fehler beim manuellen Zugriff/Zuweisung: $($_.Exception.Message)" -Type Error
    }
})

# Funktion zum Speichern der Konfiguration in die INI-Datei
function Save-ConfigToFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$ConfigData,
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )

    try {
        Write-DebugMessage "Speichere Konfiguration nach: $FilePath" -Type Info
        $iniContent = @()
        # Gehe durch die Abschnitte (z.B. General, Defaults)
        foreach ($sectionKey in $ConfigData.Keys) {
            $iniContent += "[$sectionKey]"
            $sectionData = $ConfigData[$sectionKey]
            # Gehe durch die Schlüssel-Wert-Paare im Abschnitt
            foreach ($key in $sectionData.Keys | Sort-Object) { # Sortieren für konsistente Reihenfolge
                $value = $sectionData[$key]
                # Konvertiere boolesche Werte korrekt
                if ($value -is [bool]) {
                    $value = if ($value) { "true" } else { "false" }
                }
                # Füge Zeile hinzu (nur wenn Wert nicht null oder leer ist?) - Hier erstmal alle speichern
                $iniContent += "$key = $value"
            }
            $iniContent += "" # Leerzeile zwischen Abschnitten
        }

        # Stelle sicher, dass das Verzeichnis existiert
        $directory = Split-Path -Path $FilePath -Parent
        if (-not (Test-Path -Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
            Write-DebugMessage "Verzeichnis erstellt: $directory" -Type Info
        }

        # Schreibe Inhalt in die Datei (UTF8 ohne BOM ist oft gut für INIs)
        # Verwende Set-Content für robustere Schreibvorgänge
        Set-Content -Path $FilePath -Value ($iniContent -join [Environment]::NewLine) -Encoding UTF8 -Force
        Write-DebugMessage "Konfiguration erfolgreich gespeichert." -Type Success
        Log-Action "Konfiguration gespeichert nach $FilePath"
        return $true
    }
    catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Speichern der Konfiguration nach '$FilePath'."
        Write-DebugMessage $errorMsg -Type Error
        Log-Action "Fehler beim Speichern der Konfiguration nach '$FilePath': $errorMsg"
        [System.Windows.MessageBox]::Show(
            "Fehler beim Speichern der Einstellungen:`n$errorMsg",
            "Speicherfehler",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return $false
    }
}


# Funktion zum Anzeigen und Verwalten des Einstellungsfensters
function Show-SettingsWindow {
    [CmdletBinding()]
    param()

    try {
        # Sicherstellen, dass die notwendigen Assemblies geladen sind
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $xamlPath = Join-Path -Path $PSScriptRoot -ChildPath "assets\SettingsWindow.xaml"
        if (-not (Test-Path $xamlPath)) {
            $errorMsg = "Fehler: SettingsWindow.xaml nicht gefunden unter $xamlPath"
            Write-DebugMessage $errorMsg -Type Error
            [System.Windows.MessageBox]::Show("Die XAML-Datei für das Einstellungsfenster wurde nicht gefunden.`nPfad: $xamlPath", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Return
        }

        Write-DebugMessage "Lade SettingsWindow.xaml" -Type Info
        # Korrigierte Methode zum Laden von XAML aus einer Datei
        $xamlContent = Get-Content -Path $xamlPath -Raw
        $stringReader = New-Object System.IO.StringReader -ArgumentList $xamlContent
        $xmlReader = [System.Xml.XmlReader]::Create($stringReader)
        $settingsWindow = $null # Initialisieren für den finally-Block
        try {
            $settingsWindow = [System.Windows.Markup.XamlReader]::Load($xmlReader)
        }
        catch {
            # Spezifischer Fehler beim Laden von XAML abfangen
            $loadErrorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Parsen der XAML-Datei '$xamlPath'."
            Write-DebugMessage $loadErrorMsg -Type Error
            Log-Action "Fehler beim Parsen der XAML für SettingsWindow: $loadErrorMsg"
            [System.Windows.MessageBox]::Show(
                "Fehler beim Laden des Einstellungsfensters:`n$loadErrorMsg",
                "XAML Ladefehler",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            Return # Beenden, wenn XAML nicht geladen werden kann
        }
        finally {
            # Sicherstellen, dass die Reader geschlossen werden
            if ($null -ne $xmlReader) { $xmlReader.Close() }
            if ($null -ne $stringReader) { $stringReader.Close() }
        }

        # Füge die Windows Forms Assembly hinzu (für FolderBrowserDialog)
        try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop } catch {
            Write-DebugMessage "Konnte System.Windows.Forms nicht laden. Ordnerauswahl nicht verfügbar." -Type Warning
        }

        # Steuerelemente finden
        $txtDefaultUser = $settingsWindow.FindName("txtDefaultUser")
        $chkEnableDebug = $settingsWindow.FindName("chkEnableDebug")
        $txtLogPath = $settingsWindow.FindName("txtLogPath")
        $btnBrowseLogPath = $settingsWindow.FindName("btnBrowseLogPath")
        $cmbTheme = $settingsWindow.FindName("cmbTheme")
        $btnSaveSettings = $settingsWindow.FindName("btnSaveSettings")
        $btnCancelSettings = $settingsWindow.FindName("btnCancelSettings")

        # Sicherstellen, dass $script:config existiert und Abschnitte ggf. initialisieren
        if ($null -eq $script:config) { $script:config = @{} }
        if (-not $script:config.ContainsKey("Defaults")) { $script:config["Defaults"] = @{} }
        if (-not $script:config.ContainsKey("Logging")) { $script:config["Logging"] = @{} }
        if (-not $script:config.ContainsKey("Appearance")) { $script:config["Appearance"] = @{} }


        # Aktuelle Einstellungen laden und anzeigen (aus $script:config)
        if ($null -ne $txtDefaultUser -and $script:config["Defaults"].ContainsKey("DefaultUser")) {
            $txtDefaultUser.Text = $script:config["Defaults"]["DefaultUser"]
        }
        if ($null -ne $chkEnableDebug -and $script:config["Logging"].ContainsKey("DebugEnabled")) {
             # Sicherstellen, dass der Wert als Boolean interpretiert wird
            $debugEnabledValue = $script:config["Logging"]["DebugEnabled"]
            if ($debugEnabledValue -is [string] -and $debugEnabledValue -match '^(true|false)$') {
                 $chkEnableDebug.IsChecked = [System.Convert]::ToBoolean($debugEnabledValue)
            } elseif ($debugEnabledValue -is [bool]) {
                 $chkEnableDebug.IsChecked = $debugEnabledValue
            } else {
                 $chkEnableDebug.IsChecked = $false # Fallback
                 Write-DebugMessage "Ungültiger Wert für DebugEnabled in Config: '$debugEnabledValue'. Setze auf false." -Type Warning
            }
        }
         if ($null -ne $txtLogPath -and $script:config["Logging"].ContainsKey("LogDirectory")) {
            $txtLogPath.Text = $script:config["Logging"]["LogDirectory"]
        }
         if ($null -ne $cmbTheme -and $script:config["Appearance"].ContainsKey("Theme")) {
             $currentThemeTag = $script:config["Appearance"]["Theme"]
             foreach($item in $cmbTheme.Items) {
                 if ($item.Tag -eq $currentThemeTag) {
                     $cmbTheme.SelectedItem = $item
                     break
                 }
             }
             # Fallback, falls gespeicherter Tag nicht existiert
             if ($null -eq $cmbTheme.SelectedItem) { $cmbTheme.SelectedIndex = 0 }
        } else {
            # Standardwert setzen, falls nichts in Config
            if ($null -ne $cmbTheme) { $cmbTheme.SelectedIndex = 0 }
        }


        # Event Handler für "Durchsuchen..."
        if ($null -ne $btnBrowseLogPath) {
            $btnBrowseLogPath.Add_Click({
                try {
                    # Prüfen ob Windows Forms verfügbar ist
                    if (-not ([System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms"))) {
                        [System.Windows.MessageBox]::Show("Die Komponente für die Ordnerauswahl (System.Windows.Forms) konnte nicht geladen werden.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        Return
                    }

                    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                    $folderBrowser.Description = "Wählen Sie das Verzeichnis für Log-Dateien"
                    $folderBrowser.ShowNewFolderButton = $true
                    if ($txtLogPath.Text -and (Test-Path $txtLogPath.Text -PathType Container)) {
                        $folderBrowser.SelectedPath = $txtLogPath.Text
                    }

                    # WPF Fenster als Parent setzen (optional, für besseres modales Verhalten)
                    $wpfWindowInteropHelper = New-Object System.Windows.Interop.WindowInteropHelper -ArgumentList $settingsWindow
                    $ownerHandle = $wpfWindowInteropHelper.Handle

                    # Erstelle ein Objekt, das IWin32Window implementiert
                    $ownerWindow = New-Object System.Windows.Forms.NativeWindow
                    $ownerWindow.AssignHandle($ownerHandle)

                    try {
                        # Dialog anzeigen und das NativeWindow-Objekt als Owner übergeben
                        if ($folderBrowser.ShowDialog($ownerWindow) -eq [System.Windows.Forms.DialogResult]::OK) {
                            $txtLogPath.Text = $folderBrowser.SelectedPath
                            Write-DebugMessage "Neues Log-Verzeichnis ausgewählt: $($txtLogPath.Text)" -Type Info
                        }
                    }
                    finally {
                        # Handle freigeben, um Ressourcenlecks zu vermeiden
                        $ownerWindow.ReleaseHandle()
                    }
                } catch {
                     $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Öffnen des Ordnerauswahldialogs."
                     Write-DebugMessage $errorMsg -Type Error
                     [System.Windows.MessageBox]::Show("Fehler beim Anzeigen des Ordnerauswahldialogs: $errorMsg", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            })
        }

        # Event Handler für "Speichern"
        if ($null -ne $btnSaveSettings) {
            $btnSaveSettings.Add_Click({
                # Werte auslesen
                $newDefaultUser = if($null -ne $txtDefaultUser) { $txtDefaultUser.Text.Trim() } else { "" } # Standardmäßig leer, nicht null
                $newDebugEnabled = if($null -ne $chkEnableDebug) { $chkEnableDebug.IsChecked } else { $false }
                $newLogPath = if($null -ne $txtLogPath) { $txtLogPath.Text.Trim() } else { "" } # Standardmäßig leer
                $newThemeTag = if($null -ne $cmbTheme.SelectedItem) { $cmbTheme.SelectedItem.Tag } else { "Light" }

                # Validieren (optional, hier einfach)
                if (-not $newLogPath) {
                    [System.Windows.MessageBox]::Show("Bitte geben Sie ein Log-Verzeichnis an.", "Validierung fehlgeschlagen", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                # Prüfen ob Verzeichnis existiert oder erstellt werden kann? Hier nicht, wird in Save-ConfigToFile geprüft/erstellt.

                # $script:config aktualisieren
                $script:config["Defaults"]["DefaultUser"] = $newDefaultUser
                $script:config["Logging"]["DebugEnabled"] = $newDebugEnabled # Wird als bool gespeichert
                $script:config["Logging"]["LogDirectory"] = $newLogPath
                $script:config["Appearance"]["Theme"] = $newThemeTag

                # Konfiguration speichern
                $saveResult = Save-ConfigToFile -ConfigData $script:config -FilePath $script:configFile

                if ($saveResult) {
                    # Laufzeitvariablen aktualisieren
                    $script:DebugPreference = if($newDebugEnabled) { 'Continue' } else { 'SilentlyContinue' }
                    $script:LogDir = $newLogPath
                    # Logging neu initialisieren oder Pfad aktualisieren, falls Funktion existiert
                    Initialize-Logging # Annahme: Diese Funktion existiert und verwendet $script:LogDir und $script:DebugPreference

                    Update-GuiText -TextElement $script:txtStatus -Message "Einstellungen erfolgreich gespeichert."
                    Log-Action "Einstellungen wurden gespeichert und angewendet."
                    $settingsWindow.Close()
                } else {
                    # Fehler wurde bereits in Save-ConfigToFile angezeigt
                    Update-GuiText -TextElement $script:txtStatus -Message "Fehler beim Speichern der Einstellungen."
                }
            })
        }

        # Event Handler für "Abbrechen"
        if ($null -ne $btnCancelSettings) {
            $btnCancelSettings.Add_Click({ $settingsWindow.Close() })
        }

        # Fenstereigentümer setzen, damit es modal zum Hauptfenster ist
        # Prüfen, ob $script:Form existiert und ein Fenster ist
        if ($null -ne $script:Form -and $script:Form -is [System.Windows.Window]) {
            $settingsWindow.Owner = $script:Form
        } else {
            Write-DebugMessage "Hauptfenster (\$script:Form) nicht gefunden oder ungültig. Einstellungsfenster wird nicht modal angezeigt." -Type Warning
        }


        # Fenster anzeigen
        Write-DebugMessage "Zeige Einstellungsfenster an" -Type Info
        [void]$settingsWindow.ShowDialog()
        Write-DebugMessage "Einstellungsfenster geschlossen" -Type Info

    }
    catch {
         # Allgemeiner Fehler im Einstellungsfenster (außerhalb des XAML-Ladevorgangs)
         $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Unerwarteter Fehler im Einstellungsfenster."
         Write-DebugMessage $errorMsg -Type Error
         Log-Action "Fehler im Einstellungsfenster: $errorMsg"
         [System.Windows.MessageBox]::Show(
             "Fehler im Einstellungsfenster:`n$errorMsg",
             "Fensterfehler",
             [System.Windows.MessageBoxButton]::OK,
             [System.Windows.MessageBoxImage]::Error
         )
    }
}

# --- Ende Funktionen für Einstellungsfenster ---
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