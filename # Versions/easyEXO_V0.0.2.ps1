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
=======================================================================
#>

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
AppName = Exchange Berechtigungen Verwaltung
Version = 0.0.1
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
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Mit Exchange verbunden (MFA)"
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
                    # Kalender existiert, aber Berechtigung noch nicht
                    Write-DebugMessage "Keine bestehende Berechtigung gefunden, prüfe Kalender-Existenz" -Type "Info"
                    
                    if (Get-MailboxFolderPermission -Identity $identityDE -ErrorAction SilentlyContinue) {
                        $identity = $identityDE
                        Write-DebugMessage "Deutscher Kalender existiert: $identityDE" -Type "Info"
                    } 
                    else {
                        $identity = $identityEN
                        Write-DebugMessage "Englischer Kalender existiert: $identityEN" -Type "Info"
                    }
                }
            }
        } 
        catch {
            # Bei Fehler wird automatisch Add-MailboxFolderPermission versucht
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler bei der Prüfung bestehender Berechtigungen: $errorMsg" -Type "Warning"
            
            Write-DebugMessage "Versuche Kalender-Existenz zu prüfen" -Type "Info"
            if (Get-MailboxFolderPermission -Identity $identityDE -ErrorAction SilentlyContinue) {
                $identity = $identityDE
                Write-DebugMessage "Deutscher Kalender existiert: $identityDE" -Type "Info"
            } 
            else {
                $identity = $identityEN
                Write-DebugMessage "Englischer Kalender existiert: $identityEN" -Type "Info"
            }
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
                    throw "Kalenderberechtigung konnte nicht entfernt werden: $errorMsg"
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
                    AccessRights = $perm.AccessRights
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
                SendAs = $hasSendAs
                IsInherited = $perm.IsInherited
                Deny = $perm.Deny
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
                    AccessRights = "Keine direkten Mailboxberechtigungen"
                    SendAs = $true
                    IsInherited = $false
                    Deny = $false
                }
                
                Write-DebugMessage "SendAs-Berechtigung verarbeitet: $($sendPerm.Trustee)" -Type "Info"
                $resultCollection += $entry
            }
        }
        
        # Wenn keine Berechtigungen gefunden wurden
        if ($resultCollection.Count -eq 0) {
            # Füge "NT AUTHORITY\SELF" hinzu, der normalerweise vorhanden ist
            $selfPerm = Get-MailboxPermission -Identity $Mailbox | Where-Object { $_.User -like "NT AUTHORITY\SELF" } | Select-Object -First 1
            
            if ($null -ne $selfPerm) {
                $entry = [PSCustomObject]@{
                    Identity = $Mailbox
                    User = "NT AUTHORITY\SELF"
                    AccessRights = $selfPerm.AccessRights -join ", "
                    SendAs = $false
                    IsInherited = $selfPerm.IsInherited
                    Deny = $selfPerm.Deny
                }
                
                $resultCollection += $entry
                Write-DebugMessage "Nur Standard-Postfachberechtigung gefunden: NT AUTHORITY\SELF" -Type "Info"
            }
            else {
                Write-DebugMessage "Keine Postfachberechtigungen gefunden" -Type "Warning"
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
                Write-DebugMessage "Englischer Kalenderpfad verwendet: $identity" -Type "Info"
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
                Write-DebugMessage "Englischer Kalenderpfad verwendet: $identity" -Type "Info"
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
                $email = $mailbox.PrimarySmtpAddress.ToString()
                
                # Status aktualisieren
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Bearbeite $progressIndex von $totalCount ($progressPercentage%): $email"
                }
                
                Write-DebugMessage "Bearbeite Mailbox $progressIndex/$totalCount - $email" -Type "Info"
                
                # Standard-Berechtigung setzen
                Set-DefaultCalendarPermission -MailboxUser $email -AccessRights $AccessRights -ErrorAction Stop
                $successCount++
                
                Write-DebugMessage "Standard-Kalenderberechtigungen für $email erfolgreich gesetzt" -Type "Success"
            }
            catch {
                $errorMsg = $_.Exception.Message
                $errorCount++
                Write-DebugMessage "Fehler bei $email - $errorMsg" -Type "Error"
                Log-Action "Fehler beim Setzen der Standard-Kalenderberechtigungen für $email - $errorMsg"
                # Weitermachen mit nächstem Postfach
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
        Write-DebugMessage "Fehler beim Setzen der Standard-Kalenderberechtigungen für alle: $errorMsg" -Type "Error"
        
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
                $email = $mailbox.PrimarySmtpAddress.ToString()
                
                # Status aktualisieren
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Bearbeite $progressIndex von $totalCount ($progressPercentage%): $email"
                }
                
                Write-DebugMessage "Bearbeite Mailbox $progressIndex/$totalCount - $email" -Type "Info"
                
                # Anonym-Berechtigung setzen
                Set-AnonymousCalendarPermission -MailboxUser $email -AccessRights $AccessRights -ErrorAction Stop
                $successCount++
                
                Write-DebugMessage "Anonym-Kalenderberechtigungen für $email erfolgreich gesetzt" -Type "Success"
            }
            catch {
                $errorMsg = $_.Exception.Message
                $errorCount++
                Write-DebugMessage "Fehler bei $email - $errorMsg" -Type "Error"
                Log-Action "Fehler beim Setzen der Anonym-Kalenderberechtigungen für $email - $errorMsg"
                # Weitermachen mit nächstem Postfach
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
        Write-DebugMessage "Fehler beim Setzen der Anonym-Kalenderberechtigungen für alle: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Setzen der Anonym-Kalenderberechtigungen für alle: $errorMsg"
        return $false
    }
}

function Set-CalendarDefaultPermissionsAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Standard", "Anonym", "Beides")]
        [string]$PermissionType,
        
        [Parameter(Mandatory = $true)]
        [string]$AccessRights,
        
        [Parameter(Mandatory = $false)]
        [switch]$ForAllMailboxes = $false
    )
    
    try {
        Write-DebugMessage "Setze Standardberechtigungen für Kalender: $PermissionType mit $AccessRights" -Type "Info"
        
        if ($ForAllMailboxes) {
            # Frage den Benutzer ob er das wirklich tun möchte (kann lange dauern)
            $confirmResult = [System.Windows.MessageBox]::Show(
                "Möchten Sie wirklich die $PermissionType-Berechtigungen für ALLE Postfächer setzen?",
                "Massenoperation bestätigen",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning)
                
            if ($confirmResult -eq [System.Windows.MessageBoxResult]::No) {
                Write-DebugMessage "Operation vom Benutzer abgebrochen" -Type "Info"
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Operation abgebrochen."
                }
                return $false
            }
            
            Log-Action "Starte Setzen von Standardberechtigungen ($PermissionType) für alle Kalender mit Rechten: $AccessRights"
            
            # Alle Postfächer abrufen
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Rufe alle Postfächer ab..."
            }
            
            Write-DebugMessage "Rufe alle Postfächer ab" -Type "Info"
            $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox"}
            $totalCount = $mailboxes.Count
            $successCount = 0
            $errorCount = 0
            
            Write-DebugMessage "$totalCount Postfächer gefunden" -Type "Info"
            
            # Fortschrittsanzeige vorbereiten
            $progressIndex = 0
            
            foreach ($mailbox in $mailboxes) {
                $progressIndex++
                $progressPercentage = [math]::Round(($progressIndex / $totalCount) * 100)
                $email = $mailbox.PrimarySmtpAddress.ToString()
                
                # Status aktualisieren
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Bearbeite $progressIndex von $totalCount ($progressPercentage%): $email"
                }
                
                try {
                    Write-DebugMessage "Bearbeite Mailbox $progressIndex/$totalCount - $email" -Type "Info"
                    
                    if ($PermissionType -eq "Standard" -or $PermissionType -eq "Beides") {
                        Set-DefaultCalendarPermission -MailboxUser $email -AccessRights $AccessRights
                        Write-DebugMessage "Standard-Kalenderberechtigungen für $email gesetzt" -Type "Success"
                    }
                    
                    if ($PermissionType -eq "Anonym" -or $PermissionType -eq "Beides") {
                        Set-AnonymousCalendarPermission -MailboxUser $email -AccessRights $AccessRights
                        Write-DebugMessage "Anonym-Kalenderberechtigungen für $email gesetzt" -Type "Success"
                    }
                    
                    $successCount++
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $errorCount++
                    Write-DebugMessage "Fehler bei $email - $errorMsg" -Type "Error"
                    Log-Action "Fehler beim Setzen der $PermissionType-Kalenderberechtigungen für $email - $errorMsg"
                    # Weitermachen mit nächstem Postfach
                }
            }
            
            $statusMessage = "$PermissionType-Kalenderberechtigungen für alle Postfächer gesetzt. Erfolgreich: $successCount, Fehler: $errorCount"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message $statusMessage -Color $script:connectedBrush
            }
            
            Write-DebugMessage $statusMessage -Type "Success"
            Log-Action $statusMessage
        }
        else {
            # Hier wird nur ein einzelnes Postfach bearbeitet, welches in einer anderen Funktion
            # spezifiziert werden muss. Diese Funktion wird dann wieder hierher zurückkehren.
            Write-DebugMessage "Kein ForAllMailboxes Flag gesetzt - Funktion wurde falsch aufgerufen" -Type "Warning"
            return $false
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Setzen der Kalenderberechtigungen: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Setzen der Kalenderberechtigungen: $errorMsg"
        return $false
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
        
        Write-DebugMessage "Füge SendAs-Berechtigung hinzu: $SourceUser -> $TargetUser" -Type "Info"
        
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
            Update-GuiText -TextElement $txtStatus -Message "SendAs-Berechtigung entfernt."
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
            try {
                # Eigenschaften manuell in ein neues Objekt kopieren
                $processedPermission = [PSCustomObject]@{
                    Identity    = if ($null -ne $permission.Identity) { $permission.Identity.ToString() } else { "" }
                    Trustee     = if ($null -ne $permission.Trustee) { $permission.Trustee.ToString() } else { "" }
                    AccessRights = if ($null -ne $permission.AccessRights) { ($permission.AccessRights -join ", ") } else { "" }
                    IsInherited = $permission.IsInherited
                }
                
                $processedPermissions += $processedPermission
                Write-DebugMessage "SendAs-Berechtigung verarbeitet: $($processedPermission.Trustee) -> $($processedPermission.AccessRights)" -Type "Info"
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-DebugMessage "Fehler beim Verarbeiten einer SendAs-Berechtigung: $errorMsg" -Type "Warning"
                Log-Action "Fehler beim Verarbeiten einer SendAs-Berechtigung für $($permission.Trustee): $errorMsg"
                # Fehler bei einzelner Berechtigung übersprungen, weitere verarbeiten
            }
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
            Update-GuiText -TextElement $txtStatus -Message "SendOnBehalf-Berechtigung entfernt."
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

# -------------------------------------------------
# Abschnitt: GUI Design (WPF/XAML)
# -------------------------------------------------
function Load-XAML {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$XamlFilePath
    )
    
    try {
        Write-DebugMessage "Lade XAML aus Datei: $XamlFilePath" -Type "Info"
        
        # Prüfen, ob die XAML-Datei existiert
        if (-not (Test-Path -Path $XamlFilePath)) {
            throw "XAML-Datei nicht gefunden: $XamlFilePath"
        }
        
        # XAML-Datei laden
        [xml]$xaml = Get-Content -Path $XamlFilePath -Encoding UTF8
        
        # Reader erstellen und XAML laden
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $Form = [Windows.Markup.XamlReader]::Load($reader)
        
        Write-DebugMessage "XAML-Datei erfolgreich geladen" -Type "Success"
        Log-Action "XAML-GUI erfolgreich aus Datei geladen: $XamlFilePath"
        
        return $Form
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Laden der XAML-Datei: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Laden der XAML-GUI aus Datei: $errorMsg"
        
        # Fehler-MessageBox anzeigen
        try {
            [System.Windows.MessageBox]::Show(
                "Die XAML-Datei konnte nicht geladen werden. Fehler: $errorMsg`n`nBitte wenden Sie sich an den Administrator.",
                "Fehler beim Laden der GUI",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error)
        }
        catch {
            # Fallback-Fehlerausgabe
            Write-Host "KRITISCHER FEHLER: XAML-GUI konnte nicht geladen werden: $errorMsg" -ForegroundColor Red
        }
        
        # Um das Skript zu beenden
        throw "Kritischer Fehler beim Laden der GUI: $errorMsg"
    }
}

# Pfad zur XAML-Datei
$script:xamlFilePath = Join-Path -Path $PSScriptRoot -ChildPath "assets\EXOGUI.xaml"

# -------------------------------------------------
# Abschnitt: GUI Laden
# -------------------------------------------------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

try {
    # GUI aus externer XAML-Datei laden
    $Form = Load-XAML -XamlFilePath $script:xamlFilePath
}
catch {
    Write-Host "Kritischer Fehler: $($_.Exception.Message)" -ForegroundColor Red
    
    # Falls die XAML-Datei nicht geladen werden kann, zeige einen alternativen Fehlerhinweis
    try {
        $noGuiMessage = "Die GUI konnte nicht geladen werden. Bitte stellen Sie sicher, dass die XAML-Datei im Ordner 'assets' vorhanden ist.`n`nPfad: $($script:xamlFilePath)`n`nFehlermeldung: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show($noGuiMessage, "Kritischer Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
    catch {
        Write-Host "GUI konnte nicht geladen werden und auch keine Fehlermeldung angezeigt werden." -ForegroundColor Red
    }
    
    exit
}

# -------------------------------------------------
# Abschnitt: GUI-Elemente referenzieren
# -------------------------------------------------
$btnConnect           = $Form.FindName("btnConnect")
$tabContent           = $Form.FindName("tabContent")
$tabCalendar          = $Form.FindName("tabCalendar")
$tabMailbox           = $Form.FindName("tabMailbox")
$txtStatus            = $Form.FindName("txtStatus")
$txtVersion           = $Form.FindName("txtVersion")
$txtConnectionStatus  = $Form.FindName("txtConnectionStatus")

# Referenzierung der Elemente des "Install & Check" Bereichs
$btnCheckPrerequisites   = $Form.FindName("btnCheckPrerequisites")
$btnInstallPrerequisites = $Form.FindName("btnInstallPrerequisites")
$txtInstallStatus        = $Form.FindName("txtInstallStatus")

# Referenzierung der Navigationselemente
$btnNavCalendar          = $Form.FindName("btnNavCalendar")
$btnNavMailbox           = $Form.FindName("btnNavMailbox")
$btnInfo                 = $Form.FindName("btnInfo")
$btnSettings             = $Form.FindName("btnSettings")
$btnClose                = $Form.FindName("btnClose")

# Referenzierung der Kalender-Tab-Elemente
$txtCalSourceUser        = $Form.FindName("txtCalSourceUser")
$txtCalTargetUser        = $Form.FindName("txtCalTargetUser")
$cmbCalPermission        = $Form.FindName("cmbCalPermission")
$btnCalAddPermission     = $Form.FindName("btnCalAddPermission")
$btnCalRemovePermission  = $Form.FindName("btnCalRemovePermission")
$btnCalGetPermissions    = $Form.FindName("btnCalGetPermissions")
$dgCalPermissions        = $Form.FindName("dgCalPermissions")

# Referenzierung der Mailbox-Tab-Elemente
$txtMbxSourceUser        = $Form.FindName("txtMbxSourceUser")
$txtMbxTargetUser        = $Form.FindName("txtMbxTargetUser")
$btnMbxAddPermission     = $Form.FindName("btnMbxAddPermission")
$btnMbxRemovePermission  = $Form.FindName("btnMbxRemovePermission")
$btnMbxGetPermissions    = $Form.FindName("btnMbxGetPermissions")
$dgMbxPermissions        = $Form.FindName("dgMbxPermissions")

# Referenzierung der neuen Standard/Anonym-Berechtigungs-Elemente
$txtCalSpecialUser       = $Form.FindName("txtCalSpecialUser")
$cmbCalDefaultPermission = $Form.FindName("cmbCalDefaultPermission")
$cmbCalAnonymousPermission = $Form.FindName("cmbCalAnonymousPermission")
$btnCalSetDefaultPermission = $Form.FindName("btnCalSetDefaultPermission")
$btnCalSetAnonymousPermission = $Form.FindName("btnCalSetAnonymousPermission")
$btnCalGetSpecialPermissions = $Form.FindName("btnCalGetSpecialPermissions")

# Referenzierung der neuen Massenaktions-Elemente
$cmbCalDefaultPermissionAll = $Form.FindName("cmbCalDefaultPermissionAll")
$cmbCalAnonymousPermissionAll = $Form.FindName("cmbCalAnonymousPermissionAll")
$btnCalSetDefaultPermissionAll = $Form.FindName("btnCalSetDefaultPermissionAll")
$btnCalSetAnonymousPermissionAll = $Form.FindName("btnCalSetAnonymousPermissionAll")
$btnCalSetBothPermissionsAll = $Form.FindName("btnCalSetBothPermissionsAll")

# Referenzierung der neuen erweiterten Postfachberechtigungs-Elemente
$btnMbxAddSendAsPermission = $Form.FindName("btnMbxAddSendAsPermission")
$btnMbxRemoveSendAsPermission = $Form.FindName("btnMbxRemoveSendAsPermission")
$btnMbxGetSendAsPermissions = $Form.FindName("btnMbxGetSendAsPermissions")
$btnMbxAddSendOnBehalfPermission = $Form.FindName("btnMbxAddSendOnBehalfPermission")
$btnMbxRemoveSendOnBehalfPermission = $Form.FindName("btnMbxRemoveSendOnBehalfPermission")
$btnMbxGetSendOnBehalfPermissions = $Form.FindName("btnMbxGetSendOnBehalfPermissions")

# -------------------------------------------------
# Abschnitt: Event Handler
# -------------------------------------------------

# Verbindung zu Exchange Online herstellen
$btnConnect.Add_Click({
    try {
        if ($btnConnect.Tag -eq "connect") {
            # ModernAuth-Verbindung ohne UPN starten - der Benutzer wird im Browser aufgefordert
            Connect-ExchangeOnlineSession
        } 
        elseif ($btnConnect.Tag -eq "disconnect") {
            Disconnect-ExchangeOnlineSession
        }
    } 
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler: $errorMsg"
        }
        Log-Action "Fehler beim Verbinden/Trennen: $errorMsg"
    }
})

# Event-Handler für "Check" Button im Install & Check Bereich
$btnCheckPrerequisites.Add_Click({
    try {
        if (Test-ModuleInstalled -ModuleName "ExchangeOnlineManagement") {
            $txtInstallStatus.Text = "Status - ExchangeOnlineManagement Modul installiert."
            Log-Action "Modul ExchangeOnlineManagement ist vorhanden."
        } else {
            $txtInstallStatus.Text = "Status - Modul fehlt. Bitte installieren."
            Log-Action "Modul ExchangeOnlineManagement fehlt."
        }
    } catch {
        $errorMsg = $_.Exception.Message
        $txtInstallStatus.Text = "Fehler bei der Prüfung."
        Log-Action "Fehler bei der Prüfung der Module: $errorMsg"
    }
})

# Event-Handler für "Install" Button im Install & Check Bereich
$btnInstallPrerequisites.Add_Click({
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction Stop
        $txtInstallStatus.Text = "Status: Modul erfolgreich installiert."
        Log-Action "Modul ExchangeOnlineManagement wurde erfolgreich installiert."
    } catch {
        $txtInstallStatus.Text = "Status: Fehler bei der Installation."
        Log-Action "Fehler bei der Installation des ExchangeOnlineManagement Moduls: $($_.Exception.Message)"
    }
})

# -------------------------------------------------
# Abschnitt: Event Handler für Navigation
# -------------------------------------------------
$btnNavCalendar.Add_Click({
    try {
        # Verstecke alle TabItems
        foreach ($tab in $tabContent.Items) {
            $tab.Visibility = [System.Windows.Visibility]::Collapsed
        }
        # Zeige nur Kalender TabItem
        $tabCalendar.Visibility = [System.Windows.Visibility]::Visible
        $tabCalendar.IsSelected = $true
        $txtStatus.Text = "Kalenderberechtigungen gewählt"
        Log-Action "Navigation zu Kalenderberechtigungen"
    } catch {
        $errorMsg = $_.Exception.Message
        $txtStatus.Text = "Fehler bei Navigation: $errorMsg"
        Log-Action "Fehler bei Navigation zu Kalenderberechtigungen: $errorMsg"
    }
})

$btnNavMailbox.Add_Click({
    try {
        # Verstecke alle TabItems
        foreach ($tab in $tabContent.Items) {
            $tab.Visibility = [System.Windows.Visibility]::Collapsed
        }
        # Zeige nur Mailbox TabItem
        $tabMailbox.Visibility = [System.Windows.Visibility]::Visible
        $tabMailbox.IsSelected = $true
        $txtStatus.Text = "Postfachberechtigungen gewählt"
        Log-Action "Navigation zu Postfachberechtigungen"
    } catch {
        $errorMsg = $_.Exception.Message
        $txtStatus.Text = "Fehler bei Navigation: $errorMsg"
        Log-Action "Fehler bei Navigation zu Postfachberechtigungen: $errorMsg"
    }
})

$btnClose.Add_Click({
    try {
        if ($script:isConnected) {
            $confirmResult = [System.Windows.MessageBox]::Show(
                "Sie sind noch mit Exchange Online verbunden. Möchten Sie die Verbindung trennen und das Programm beenden?",
                "Programm beenden",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question)
            
            if ($confirmResult -eq [System.Windows.MessageBoxResult]::Yes) {
                Disconnect-ExchangeOnlineSession
                $Form.Close()
            }
        } else {
            $Form.Close()
        }
    } catch {
        $errorMsg = $_.Exception.Message
        Log-Action "Fehler beim Schließen des Programms: $errorMsg"
        $Form.Close()
    }
})

$btnInfo.Add_Click({
    try {
        $infoMessage = @"
Exchange Berechtigungen Verwaltung Tool
Version: 0.0.1

Mit diesem Tool können Sie Kalender- und Postfachberechtigungen in Exchange Online verwalten.

Voraussetzungen:
- PowerShell 5.1 oder höher
- ExchangeOnlineManagement Modul
- Internetverbindung
- Exchange Online Admin-Berechtigungen
"@
        [System.Windows.MessageBox]::Show($infoMessage, "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    } catch {
        $errorMsg = $_.Exception.Message
        Log-Action "Fehler bei Info-Anzeige: $errorMsg"
    }
})

# -------------------------------------------------
# Abschnitt: Event Handler für Kalenderberechtigungen
# -------------------------------------------------
$btnCalAddPermission.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Kalenderberechtigungen hinzufügen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtCalSourceUser.Text.Trim()
        $targetUser = $txtCalTargetUser.Text.Trim()
        $selectedItem = $cmbCalPermission.SelectedItem
        
        Write-DebugMessage "Kalenderberechtigung hinzufügen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser) -or [string]::IsNullOrEmpty($targetUser) -or $null -eq $selectedItem) {
            Write-DebugMessage "Kalenderberechtigung hinzufügen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte alle Felder ausfüllen."
            }
            
            Log-Action "Unvollständige Eingabe beim Hinzufügen von Kalenderberechtigungen."
            return
        }
        
        # Mapping der deutschen UI-Einträge auf die PowerShell-Cmdlet Parameter
        $permissionMapping = @{
            "Keine" = "None"
            "Besitzer" = "Owner"
            "PublishingEditor" = "PublishingEditor"
            "Editor" = "Editor"
            "PublishingAuthor" = "PublishingAuthor"
            "Autor" = "Author"
            "NichtBearbeitenderAutor" = "NonEditingAuthor"
            "Reviewer" = "Reviewer"
            "Mitwirkender" = "Contributor"
            "FreiBusyZeit" = "AvailabilityOnly"
            "FreiBusyZeitBetreffortUndBeschreibung" = "LimitedDetails"
        }
        
        $permission = $permissionMapping[$selectedItem.Content]
        Write-DebugMessage "Kalenderberechtigung hinzufügen: Quellbenutzer=$sourceUser, Zielbenutzer=$targetUser, Berechtigung=$permission" -Type "Info"
        
        # Berechtigungen hinzufügen
        $result = Add-CalendarPermission -SourceUser $sourceUser -TargetUser $targetUser -Permission $permission
        
        if ($result) {
            Write-DebugMessage "Kalenderberechtigung erfolgreich hinzugefügt/aktualisiert. Aktualisiere DataGrid." -Type "Success"
            # Nach erfolgreicher Aktion die Berechtigungen aktualisieren
            $btnCalGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen der Kalenderberechtigung: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Hinzufügen der Kalenderberechtigung: $errorMsg"
    }
})

$btnCalRemovePermission.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Kalenderberechtigungen entfernen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtCalSourceUser.Text.Trim()
        $targetUser = $txtCalTargetUser.Text.Trim()
        
        Write-DebugMessage "Kalenderberechtigung entfernen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser) -or [string]::IsNullOrEmpty($targetUser)) {
            Write-DebugMessage "Kalenderberechtigung entfernen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Quell- und Zielbenutzer angeben."
            }
            
            Log-Action "Unvollständige Eingabe beim Entfernen von Kalenderberechtigungen."
            return
        }
        
        Write-DebugMessage "Kalenderberechtigung entfernen: Quellbenutzer=$sourceUser, Zielbenutzer=$targetUser" -Type "Info"
        $result = Remove-CalendarPermission -SourceUser $sourceUser -TargetUser $targetUser
        
        if ($result) {
            Write-DebugMessage "Kalenderberechtigung erfolgreich entfernt. Aktualisiere DataGrid." -Type "Success"
            # Nach erfolgreicher Aktion die Berechtigungen aktualisieren
            $btnCalGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen der Kalenderberechtigung: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Entfernen der Kalenderberechtigung: $errorMsg"
    }
})

$btnCalGetPermissions.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Kalenderberechtigungen abrufen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtCalSourceUser.Text.Trim()
        
        Write-DebugMessage "Kalenderberechtigungen abrufen: Validiere Benutzereingabe" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser)) {
            Write-DebugMessage "Kalenderberechtigungen abrufen: Quellpostfach fehlt" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Quellpostfach angeben."
            }
            
            Log-Action "Quellpostfach fehlt beim Abrufen von Kalenderberechtigungen."
            return
        }
        
        # Berechtigungen abrufen
        Write-DebugMessage "Kalenderberechtigungen abrufen für: $sourceUser" -Type "Info"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Berechtigungen werden abgerufen..."
        }
        
        $permissions = Get-CalendarPermission -MailboxUser $sourceUser
        
        # DataGrid leeren und neu füllen
        if ($null -ne $dgCalPermissions) {
            $dgCalPermissions.Dispatcher.Invoke([Action]{
                $dgCalPermissions.ItemsSource = $null
                
                if ($permissions -and $permissions.Count -gt 0) {
                    $dgCalPermissions.ItemsSource = $permissions
                    Write-DebugMessage "Kalenderberechtigungen erfolgreich abgerufen: $($permissions.Count) Einträge gefunden" -Type "Success"
                }
                else {
                    Write-DebugMessage "Keine Kalenderberechtigungen gefunden" -Type "Info"
                }
            }, "Normal")
        }
        
        if ($permissions -and $permissions.Count -gt 0) {
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Kalenderberechtigungen für $sourceUser abgerufen: $($permissions.Count) Einträge gefunden." -Color $script:connectedBrush
            }
            
            Log-Action "Kalenderberechtigungen für $sourceUser abgerufen: $($permissions.Count) Einträge gefunden."
        } 
        else {
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine Kalenderberechtigungen gefunden oder Fehler beim Abrufen."
            }
            
            Log-Action "Keine Kalenderberechtigungen für $sourceUser gefunden."
        }
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Kalenderberechtigungen: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Abrufen der Kalenderberechtigungen: $errorMsg"
    }
})

# -------------------------------------------------
# Abschnitt: Event Handler für Postfachberechtigungen
# -------------------------------------------------
$btnMbxAddPermission.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Postfachberechtigungen hinzufügen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtMbxSourceUser.Text.Trim()
        $targetUser = $txtMbxTargetUser.Text.Trim()
        
        Write-DebugMessage "Postfachberechtigung hinzufügen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser) -or [string]::IsNullOrEmpty($targetUser)) {
            Write-DebugMessage "Postfachberechtigung hinzufügen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte alle Felder ausfüllen."
            }
            
            Log-Action "Unvollständige Eingabe beim Hinzufügen von Postfachberechtigungen."
            return
        }
        
        Write-DebugMessage "Postfachberechtigung hinzufügen: Quellbenutzer=$sourceUser, Zielbenutzer=$targetUser" -Type "Info"
        $result = Add-MailboxPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
        
        if ($result) {
            Write-DebugMessage "Postfachberechtigung erfolgreich hinzugefügt. Aktualisiere DataGrid." -Type "Success"
            # Nach erfolgreicher Aktion die Berechtigungen aktualisieren
            $btnMbxGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen der Postfachberechtigung: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Hinzufügen der Postfachberechtigung: $errorMsg"
    }
})

$btnMbxRemovePermission.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Postfachberechtigungen entfernen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtMbxSourceUser.Text.Trim()
        $targetUser = $txtMbxTargetUser.Text.Trim()
        
        Write-DebugMessage "Postfachberechtigung entfernen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser) -or [string]::IsNullOrEmpty($targetUser)) {
            Write-DebugMessage "Postfachberechtigung entfernen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Quell- und Zielbenutzer angeben."
            }
            
            Log-Action "Unvollständige Eingabe beim Entfernen von Postfachberechtigungen."
            return
        }
        
        Write-DebugMessage "Postfachberechtigung entfernen: Quellbenutzer=$sourceUser, Zielbenutzer=$targetUser" -Type "Info"
        $result = Remove-MailboxPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
        
        if ($result) {
            Write-DebugMessage "Postfachberechtigung erfolgreich entfernt. Aktualisiere DataGrid." -Type "Success"
            # Nach erfolgreicher Aktion die Berechtigungen aktualisieren
            $btnMbxGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen der Postfachberechtigung: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Entfernen der Postfachberechtigung: $errorMsg"
    }
})

$btnMbxGetPermissions.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Postfachberechtigungen abrufen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtMbxSourceUser.Text.Trim()
        
        Write-DebugMessage "Postfachberechtigungen abrufen: Validiere Benutzereingabe" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser)) {
            Write-DebugMessage "Postfachberechtigungen abrufen: Quellpostfach fehlt" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Quellpostfach angeben."
            }
            
            Log-Action "Quellpostfach fehlt beim Abrufen von Postfachberechtigungen."
            return
        }
        
        # Berechtigungen abrufen
        Write-DebugMessage "Postfachberechtigungen abrufen für: $sourceUser" -Type "Info"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Berechtigungen werden abgerufen..."
        }
        
        $permissions = Get-MailboxPermissionsAction -MailboxUser $sourceUser
        
        # DataGrid leeren und neu füllen
        if ($null -ne $dgMbxPermissions) {
            $dgMbxPermissions.Dispatcher.Invoke([Action]{
                $dgMbxPermissions.ItemsSource = $null
                
                if ($permissions -and $permissions.Count -gt 0) {
                    $dgMbxPermissions.ItemsSource = $permissions
                    Write-DebugMessage "Postfachberechtigungen erfolgreich abgerufen: $($permissions.Count) Einträge gefunden" -Type "Success"
                }
                else {
                    Write-DebugMessage "Keine Postfachberechtigungen gefunden" -Type "Info"
                }
            }, "Normal")
        }
        
        if ($permissions -and $permissions.Count -gt 0) {
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Postfachberechtigungen für $sourceUser abgerufen: $($permissions.Count) Einträge gefunden." -Color $script:connectedBrush
            }
            
            Log-Action "Postfachberechtigungen für $sourceUser abgerufen: $($permissions.Count) Einträge gefunden."
        }
        else {
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine Postfachberechtigungen gefunden oder Fehler beim Abrufen."
            }
            
            Log-Action "Keine Postfachberechtigungen für $sourceUser gefunden."
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Postfachberechtigungen: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Abrufen der Postfachberechtigungen: $errorMsg"
    }
})

# -------------------------------------------------
# Abschnitt: Event Handler für Standard/Anonym-Kalenderberechtigungen
# -------------------------------------------------
$btnCalSetDefaultPermission.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Standard-Kalenderberechtigungen setzen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $calendarUser = $txtCalSpecialUser.Text.Trim()
        $selectedItem = $cmbCalDefaultPermission.SelectedItem
        
        Write-DebugMessage "Standard-Kalenderberechtigungen setzen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($calendarUser) -or $null -eq $selectedItem) {
            Write-DebugMessage "Standard-Kalenderberechtigungen setzen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte alle Felder ausfüllen."
            }
            
            Log-Action "Unvollständige Eingabe beim Setzen von Standard-Kalenderberechtigungen."
            return
        }
        
        # Mapping der deutschen UI-Einträge auf die PowerShell-Cmdlet Parameter
        $permissionMapping = @{
            "Keine" = "None"
            "Besitzer" = "Owner"
            "PublishingEditor" = "PublishingEditor"
            "Editor" = "Editor"
            "PublishingAuthor" = "PublishingAuthor"
            "Autor" = "Author"
            "NichtBearbeitenderAutor" = "NonEditingAuthor"
            "Reviewer" = "Reviewer"
            "Mitwirkender" = "Contributor"
            "FreiBusyZeit" = "AvailabilityOnly"
            "FreiBusyZeitBetreffortUndBeschreibung" = "LimitedDetails"
        }
        
        $permission = $permissionMapping[$selectedItem.Content]
        Write-DebugMessage "Standard-Kalenderberechtigungen setzen: Postfach=$calendarUser, Berechtigung=$permission" -Type "Info"
        
        # Standard-Berechtigungen setzen
        try {
            Set-DefaultCalendarPermission -MailboxUser $calendarUser -AccessRights $permission
            
            Write-DebugMessage "Standard-Kalenderberechtigungen erfolgreich gesetzt." -Type "Success"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Standard-Kalenderberechtigungen für $calendarUser erfolgreich gesetzt." -Color $script:connectedBrush
            }
            
            # Nach erfolgreicher Aktion die Berechtigungen aktualisieren
            if (-not [string]::IsNullOrEmpty($txtCalSourceUser.Text) -and $txtCalSourceUser.Text -eq $calendarUser) {
                $btnCalGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            }
            
            # Auch im Speziellen Bereich die Berechtigungen neu laden
            $btnCalGetSpecialPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Setzen der Standard-Kalenderberechtigungen: $errorMsg" -Type "Error"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
            }
            
            Log-Action "Fehler beim Setzen der Standard-Kalenderberechtigungen: $errorMsg"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler im Event-Handler für Standard-Kalenderberechtigungen: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler im Event-Handler für Standard-Kalenderberechtigungen: $errorMsg"
    }
})

$btnCalSetAnonymousPermission.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Anonym-Kalenderberechtigungen setzen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $calendarUser = $txtCalSpecialUser.Text.Trim()
        $selectedItem = $cmbCalAnonymousPermission.SelectedItem
        
        Write-DebugMessage "Anonym-Kalenderberechtigungen setzen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($calendarUser) -or $null -eq $selectedItem) {
            Write-DebugMessage "Anonym-Kalenderberechtigungen setzen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte alle Felder ausfüllen."
            }
            
            Log-Action "Unvollständige Eingabe beim Setzen von Anonym-Kalenderberechtigungen."
            return
        }
        
        # Mapping der deutschen UI-Einträge auf die PowerShell-Cmdlet Parameter
        $permissionMapping = @{
            "Keine" = "None"
            "Besitzer" = "Owner"
            "PublishingEditor" = "PublishingEditor"
            "Editor" = "Editor"
            "PublishingAuthor" = "PublishingAuthor"
            "Autor" = "Author"
            "NichtBearbeitenderAutor" = "NonEditingAuthor"
            "Reviewer" = "Reviewer"
            "Mitwirkender" = "Contributor"
            "FreiBusyZeit" = "AvailabilityOnly"
            "FreiBusyZeitBetreffortUndBeschreibung" = "LimitedDetails"
        }
        
        $permission = $permissionMapping[$selectedItem.Content]
        Write-DebugMessage "Anonym-Kalenderberechtigungen setzen: Postfach=$calendarUser, Berechtigung=$permission" -Type "Info"
        
        # Anonym-Berechtigungen setzen
        try {
            Set-AnonymousCalendarPermission -MailboxUser $calendarUser -AccessRights $permission
            
            Write-DebugMessage "Anonym-Kalenderberechtigungen erfolgreich gesetzt." -Type "Success"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Anonym-Kalenderberechtigungen für $calendarUser erfolgreich gesetzt." -Color $script:connectedBrush
            }
            
            # Nach erfolgreicher Aktion die Berechtigungen aktualisieren
            if (-not [string]::IsNullOrEmpty($txtCalSourceUser.Text) -and $txtCalSourceUser.Text -eq $calendarUser) {
                $btnCalGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            }
            
            # Auch im Speziellen Bereich die Berechtigungen neu laden
            $btnCalGetSpecialPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Setzen der Anonym-Kalenderberechtigungen: $errorMsg" -Type "Error"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
            }
            
            Log-Action "Fehler beim Setzen der Anonym-Kalenderberechtigungen: $errorMsg"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler im Event-Handler für Anonym-Kalenderberechtigungen: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler im Event-Handler für Anonym-Kalenderberechtigungen: $errorMsg"
    }
})

$btnCalGetSpecialPermissions.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Spezielle Kalenderberechtigungen abrufen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $calendarUser = $txtCalSpecialUser.Text.Trim()
        
        Write-DebugMessage "Spezielle Kalenderberechtigungen abrufen: Validiere Benutzereingabe" -Type "Info"
        
        if ([string]::IsNullOrEmpty($calendarUser)) {
            Write-DebugMessage "Spezielle Kalenderberechtigungen abrufen: Postfach fehlt" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Kalenderpostfach angeben."
            }
            
            Log-Action "Kalenderpostfach fehlt beim Abrufen von speziellen Kalenderberechtigungen."
            return
        }
        
        # Berechtigungen abrufen
        Write-DebugMessage "Spezielle Kalenderberechtigungen abrufen für: $calendarUser" -Type "Info"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Berechtigungen werden abgerufen..."
        }
        
        # Kopiere den Wert ins Quellfeld des normalen Bereichs für einfacheres Arbeiten
        $txtCalSourceUser.Text = $calendarUser
        
        # Normalen "Berechtigungen abrufen" Knopf auslösen, um das DataGrid zu füllen
        $btnCalGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        
        # Erfolgreiche Meldung anzeigen
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Kalenderberechtigungen für $calendarUser abgerufen." -Color $script:connectedBrush
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der speziellen Kalenderberechtigungen: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Abrufen der speziellen Kalenderberechtigungen: $errorMsg"
    }
})

# -------------------------------------------------
# Abschnitt: Event Handler für Massenaktionen
# -------------------------------------------------
$btnCalSetDefaultPermissionAll.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Standard-Kalenderberechtigungen für alle setzen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $selectedItem = $cmbCalDefaultPermissionAll.SelectedItem
        
        Write-DebugMessage "Standard-Kalenderberechtigungen für alle setzen: Validiere Benutzereingaben" -Type "Info"
        
        if ($null -eq $selectedItem) {
            Write-DebugMessage "Standard-Kalenderberechtigungen für alle setzen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Berechtigungstyp auswählen."
            }
            
            Log-Action "Unvollständige Eingabe beim Setzen von Standard-Kalenderberechtigungen für alle."
            return
        }
        
        # Mapping der deutschen UI-Einträge auf die PowerShell-Cmdlet Parameter
        $permissionMapping = @{
            "Keine" = "None"
            "Besitzer" = "Owner"
            "PublishingEditor" = "PublishingEditor"
            "Editor" = "Editor"
            "PublishingAuthor" = "PublishingAuthor"
            "Autor" = "Author"
            "NichtBearbeitenderAutor" = "NonEditingAuthor"
            "Reviewer" = "Reviewer"
            "Mitwirkender" = "Contributor"
            "FreiBusyZeit" = "AvailabilityOnly"
            "FreiBusyZeitBetreffortUndBeschreibung" = "LimitedDetails"
        }
        
        $permission = $permissionMapping[$selectedItem.Content]
        Write-DebugMessage "Standard-Kalenderberechtigungen für alle setzen: Berechtigung=$permission" -Type "Info"
        
        # Standard-Berechtigungen für alle setzen
        try {
            Set-DefaultCalendarPermissionForAll -AccessRights $permission
            
            Write-DebugMessage "Standard-Kalenderberechtigungen für alle erfolgreich gesetzt." -Type "Success"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Standard-Kalenderberechtigungen für alle erfolgreich gesetzt." -Color $script:connectedBrush
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Setzen der Standard-Kalenderberechtigungen für alle: $errorMsg" -Type "Error"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
            }
            
            Log-Action "Fehler beim Setzen der Standard-Kalenderberechtigungen für alle: $errorMsg"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler im Event-Handler für Standard-Kalenderberechtigungen für alle: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler im Event-Handler für Standard-Kalenderberechtigungen für alle: $errorMsg"
    }
})

$btnCalSetAnonymousPermissionAll.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Anonym-Kalenderberechtigungen für alle setzen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $selectedItem = $cmbCalAnonymousPermissionAll.SelectedItem
        
        Write-DebugMessage "Anonym-Kalenderberechtigungen für alle setzen: Validiere Benutzereingaben" -Type "Info"
        
        if ($null -eq $selectedItem) {
            Write-DebugMessage "Anonym-Kalenderberechtigungen für alle setzen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Berechtigungstyp auswählen."
            }
            
            Log-Action "Unvollständige Eingabe beim Setzen von Anonym-Kalenderberechtigungen für alle."
            return
        }
        
        # Mapping der deutschen UI-Einträge auf die PowerShell-Cmdlet Parameter
        $permissionMapping = @{
            "Keine" = "None"
            "Besitzer" = "Owner"
            "PublishingEditor" = "PublishingEditor"
            "Editor" = "Editor"
            "PublishingAuthor" = "PublishingAuthor"
            "Autor" = "Author"
            "NichtBearbeitenderAutor" = "NonEditingAuthor"
            "Reviewer" = "Reviewer"
            "Mitwirkender" = "Contributor"
            "FreiBusyZeit" = "AvailabilityOnly"
            "FreiBusyZeitBetreffortUndBeschreibung" = "LimitedDetails"
        }
        
        $permission = $permissionMapping[$selectedItem.Content]
        Write-DebugMessage "Anonym-Kalenderberechtigungen für alle setzen: Berechtigung=$permission" -Type "Info"
        
        # Anonym-Berechtigungen für alle setzen
        try {
            Set-AnonymousCalendarPermissionForAll -AccessRights $permission
            
            Write-DebugMessage "Anonym-Kalenderberechtigungen für alle erfolgreich gesetzt." -Type "Success"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Anonym-Kalenderberechtigungen für alle erfolgreich gesetzt." -Color $script:connectedBrush
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Setzen der Anonym-Kalenderberechtigungen für alle: $errorMsg" -Type "Error"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
            }
            
            Log-Action "Fehler beim Setzen der Anonym-Kalenderberechtigungen für alle: $errorMsg"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler im Event-Handler für Anonym-Kalenderberechtigungen für alle: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler im Event-Handler für Anonym-Kalenderberechtigungen für alle: $errorMsg"
    }
})

$btnCalSetBothPermissionsAll.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Standard und Anonym-Kalenderberechtigungen für alle setzen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $selectedDefaultItem = $cmbCalDefaultPermissionAll.SelectedItem
        $selectedAnonymousItem = $cmbCalAnonymousPermissionAll.SelectedItem
        
        Write-DebugMessage "Standard und Anonym-Kalenderberechtigungen für alle setzen: Validiere Benutzereingaben" -Type "Info"
        
        if ($null -eq $selectedDefaultItem -or $null -eq $selectedAnonymousItem) {
            Write-DebugMessage "Standard und Anonym-Kalenderberechtigungen für alle setzen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte beide Berechtigungstypen auswählen."
            }
            
            Log-Action "Unvollständige Eingabe beim Setzen von Standard und Anonym-Kalenderberechtigungen für alle."
            return
        }
        
        # Mapping der deutschen UI-Einträge auf die PowerShell-Cmdlet Parameter
        $permissionMapping = @{
            "Keine" = "None"
            "Besitzer" = "Owner"
            "PublishingEditor" = "PublishingEditor"
            "Editor" = "Editor"
            "PublishingAuthor" = "PublishingAuthor"
            "Autor" = "Author"
            "NichtBearbeitenderAutor" = "NonEditingAuthor"
            "Reviewer" = "Reviewer"
            "Mitwirkender" = "Contributor"
            "FreiBusyZeit" = "AvailabilityOnly"
            "FreiBusyZeitBetreffortUndBeschreibung" = "LimitedDetails"
        }
        
        $defaultPermission = $permissionMapping[$selectedDefaultItem.Content]
        $anonymousPermission = $permissionMapping[$selectedAnonymousItem.Content]
        Write-DebugMessage "Standard und Anonym-Kalenderberechtigungen für alle setzen: Standard=$defaultPermission, Anonym=$anonymousPermission" -Type "Info"
        
        # Frage den Benutzer ob er das wirklich tun möchte (kann lange dauern)
        $confirmResult = [System.Windows.MessageBox]::Show(
            "Möchten Sie wirklich die Standard- UND Anonym-Berechtigungen für ALLE Postfächer setzen?",
            "Massenoperation bestätigen",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
            
        if ($confirmResult -eq [System.Windows.MessageBoxResult]::No) {
            Write-DebugMessage "Operation vom Benutzer abgebrochen" -Type "Info"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Operation abgebrochen."
            }
            return
        }
        
        # Standard-Berechtigungen für alle setzen
        try {
            Set-DefaultCalendarPermissionForAll -AccessRights $defaultPermission
            Write-DebugMessage "Standard-Kalenderberechtigungen für alle erfolgreich gesetzt." -Type "Success"
            
            # Nach einer kurzen Pause die anonymen Berechtigungen setzen
            Start-Sleep -Seconds 2
            
            Set-AnonymousCalendarPermissionForAll -AccessRights $anonymousPermission
            Write-DebugMessage "Anonym-Kalenderberechtigungen für alle erfolgreich gesetzt." -Type "Success"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Standard und Anonym-Kalenderberechtigungen für alle erfolgreich gesetzt." -Color $script:connectedBrush
            }
            
            Log-Action "Standard ($defaultPermission) und Anonym-Kalenderberechtigungen ($anonymousPermission) für alle erfolgreich gesetzt."
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Setzen der Standard und Anonym-Kalenderberechtigungen für alle: $errorMsg" -Type "Error"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
            }
            
            Log-Action "Fehler beim Setzen der Standard und Anonym-Kalenderberechtigungen für alle: $errorMsg"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler im Event-Handler für Standard und Anonym-Kalenderberechtigungen für alle: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler im Event-Handler für Standard und Anonym-Kalenderberechtigungen für alle: $errorMsg"
    }
})

# -------------------------------------------------
# Event Handler für SendAs/SendOnBehalf Berechtigungen
# -------------------------------------------------

# SendAs hinzufügen
$btnMbxAddSendAsPermission.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "SendAs-Berechtigungen hinzufügen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtMbxSourceUser.Text.Trim()
        $targetUser = $txtMbxTargetUser.Text.Trim()
        
        Write-DebugMessage "SendAs-Berechtigung hinzufügen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser) -or [string]::IsNullOrEmpty($targetUser)) {
            Write-DebugMessage "SendAs-Berechtigung hinzufügen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte alle Felder ausfüllen."
            }
            
            Log-Action "Unvollständige Eingabe beim Hinzufügen von SendAs-Berechtigungen."
            return
        }
        
        Write-DebugMessage "SendAs-Berechtigung hinzufügen: Quellbenutzer=$sourceUser, Zielbenutzer=$targetUser" -Type "Info"
        $result = Add-SendAsPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
        
        if ($result) {
            Write-DebugMessage "SendAs-Berechtigung erfolgreich hinzugefügt. Aktualisiere DataGrid." -Type "Success"
            # Nach erfolgreicher Aktion die Berechtigungen aktualisieren
            $btnMbxGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen der SendAs-Berechtigung: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Hinzufügen der SendAs-Berechtigung: $errorMsg"
    }
})

# SendAs entfernen
$btnMbxRemoveSendAsPermission.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "SendAs-Berechtigungen entfernen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtMbxSourceUser.Text.Trim()
        $targetUser = $txtMbxTargetUser.Text.Trim()
        
        Write-DebugMessage "SendAs-Berechtigung entfernen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser) -or [string]::IsNullOrEmpty($targetUser)) {
            Write-DebugMessage "SendAs-Berechtigung entfernen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Quell- und Zielbenutzer angeben."
            }
            
            Log-Action "Unvollständige Eingabe beim Entfernen von SendAs-Berechtigungen."
            return
        }
        
        Write-DebugMessage "SendAs-Berechtigung entfernen: Quellbenutzer=$sourceUser, Zielbenutzer=$targetUser" -Type "Info"
        $result = Remove-SendAsPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
        
        if ($result) {
            Write-DebugMessage "SendAs-Berechtigung erfolgreich entfernt. Aktualisiere DataGrid." -Type "Success"
            # Nach erfolgreicher Aktion die Berechtigungen aktualisieren
            $btnMbxGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen der SendAs-Berechtigung: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Entfernen der SendAs-Berechtigung: $errorMsg"
    }
})

# SendAs Berechtigungen anzeigen
$btnMbxGetSendAsPermissions.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "SendAs-Berechtigungen anzeigen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtMbxSourceUser.Text.Trim()
        
        Write-DebugMessage "SendAs-Berechtigungen anzeigen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser)) {
            Write-DebugMessage "SendAs-Berechtigungen anzeigen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Quellpostfach angeben."
            }
            
            Log-Action "Unvollständige Eingabe beim Anzeigen von SendAs-Berechtigungen."
            return
        }
        
        Write-DebugMessage "SendAs-Berechtigungen anzeigen für: $sourceUser" -Type "Info"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "SendAs-Berechtigungen werden abgerufen..."
        }
        
        # Berechtigungen abrufen
        $permissions = Get-SendAsPermissionAction -MailboxUser $sourceUser
        
        # DataGrid leeren und neu füllen
        if ($null -ne $dgMbxPermissions) {
            $dgMbxPermissions.Dispatcher.Invoke([Action]{
                $dgMbxPermissions.ItemsSource = $null
                
                if ($permissions -and $permissions.Count -gt 0) {
                    $dgMbxPermissions.ItemsSource = $permissions
                    Write-DebugMessage "SendAs-Berechtigungen erfolgreich abgerufen: $($permissions.Count) Einträge gefunden" -Type "Success"
                }
                else {
                    Write-DebugMessage "Keine SendAs-Berechtigungen gefunden" -Type "Info"
                }
            }, "Normal")
        }
        
        if ($permissions -and $permissions.Count -gt 0) {
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "SendAs-Berechtigungen für $sourceUser abgerufen: $($permissions.Count) Einträge gefunden." -Color $script:connectedBrush
            }
            
            Log-Action "SendAs-Berechtigungen für $sourceUser abgerufen: $($permissions.Count) Einträge gefunden."
        }
        else {
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine SendAs-Berechtigungen gefunden oder Fehler beim Abrufen."
            }
            
            Log-Action "Keine SendAs-Berechtigungen für $sourceUser gefunden."
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der SendAs-Berechtigungen: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Abrufen der SendAs-Berechtigungen: $errorMsg"
    }
})

# SendOnBehalf hinzufügen
$btnMbxAddSendOnBehalfPermission.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "SendOnBehalf-Berechtigungen hinzufügen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtMbxSourceUser.Text.Trim()
        $targetUser = $txtMbxTargetUser.Text.Trim()
        
        Write-DebugMessage "SendOnBehalf-Berechtigung hinzufügen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser) -or [string]::IsNullOrEmpty($targetUser)) {
            Write-DebugMessage "SendOnBehalf-Berechtigung hinzufügen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte alle Felder ausfüllen."
            }
            
            Log-Action "Unvollständige Eingabe beim Hinzufügen von SendOnBehalf-Berechtigungen."
            return
        }
        
        Write-DebugMessage "SendOnBehalf-Berechtigung hinzufügen: Quellbenutzer=$sourceUser, Zielbenutzer=$targetUser" -Type "Info"
        $result = Add-SendOnBehalfPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
        
        if ($result) {
            Write-DebugMessage "SendOnBehalf-Berechtigung erfolgreich hinzugefügt. Aktualisiere DataGrid." -Type "Success"
            # Nach erfolgreicher Aktion die Berechtigungen aktualisieren
            $btnMbxGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen der SendOnBehalf-Berechtigung: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Hinzufügen der SendOnBehalf-Berechtigung: $errorMsg"
    }
})

# SendOnBehalf entfernen
$btnMbxRemoveSendOnBehalfPermission.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "SendOnBehalf-Berechtigungen entfernen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtMbxSourceUser.Text.Trim()
        $targetUser = $txtMbxTargetUser.Text.Trim()
        
        Write-DebugMessage "SendOnBehalf-Berechtigung entfernen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser) -or [string]::IsNullOrEmpty($targetUser)) {
            Write-DebugMessage "SendOnBehalf-Berechtigung entfernen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Quell- und Zielbenutzer angeben."
            }
            
            Log-Action "Unvollständige Eingabe beim Entfernen von SendOnBehalf-Berechtigungen."
            return
        }
        
        Write-DebugMessage "SendOnBehalf-Berechtigung entfernen: Quellbenutzer=$sourceUser, Zielbenutzer=$targetUser" -Type "Info"
        $result = Remove-SendOnBehalfPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
        
        if ($result) {
            Write-DebugMessage "SendOnBehalf-Berechtigung erfolgreich entfernt. Aktualisiere DataGrid." -Type "Success"
            # Nach erfolgreicher Aktion die Berechtigungen aktualisieren
            $btnMbxGetPermissions.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Entfernen der SendOnBehalf-Berechtigung: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Entfernen der SendOnBehalf-Berechtigung: $errorMsg"
    }
})

# SendOnBehalf Berechtigungen anzeigen
$btnMbxGetSendOnBehalfPermissions.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "SendOnBehalf-Berechtigungen anzeigen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $sourceUser = $txtMbxSourceUser.Text.Trim()
        
        Write-DebugMessage "SendOnBehalf-Berechtigungen anzeigen: Validiere Benutzereingaben" -Type "Info"
        
        if ([string]::IsNullOrEmpty($sourceUser)) {
            Write-DebugMessage "SendOnBehalf-Berechtigungen anzeigen: Unvollständige Eingabe" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Bitte Quellpostfach angeben."
            }
            
            Log-Action "Unvollständige Eingabe beim Anzeigen von SendOnBehalf-Berechtigungen."
            return
        }
        
        Write-DebugMessage "SendOnBehalf-Berechtigungen anzeigen für: $sourceUser" -Type "Info"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "SendOnBehalf-Berechtigungen werden abgerufen..."
        }
        
        # Berechtigungen abrufen (Mailbox-Objekt holen)
        try {
            $mailbox = Get-Mailbox -Identity $sourceUser -ErrorAction Stop
            
            # Eigene Struktur für DataGrid erstellen
            $permissions = @()
            
            if ($mailbox.GrantSendOnBehalfTo) {
                foreach ($delegate in $mailbox.GrantSendOnBehalfTo) {
                    $permObj = [PSCustomObject]@{
                        Identity = $sourceUser
                        User = $delegate.ToString()
                        AccessRights = "SendOnBehalf"
                        IsInherited = $false
                        Deny = $false
                    }
                    $permissions += $permObj
                }
            }
            
            # DataGrid leeren und neu füllen
            if ($null -ne $dgMbxPermissions) {
                $dgMbxPermissions.Dispatcher.Invoke([Action]{
                    $dgMbxPermissions.ItemsSource = $null
                    
                    if ($permissions -and $permissions.Count -gt 0) {
                        $dgMbxPermissions.ItemsSource = $permissions
                        Write-DebugMessage "SendOnBehalf-Berechtigungen erfolgreich abgerufen: $($permissions.Count) Einträge gefunden" -Type "Success"
                    }
                    else {
                        Write-DebugMessage "Keine SendOnBehalf-Berechtigungen gefunden" -Type "Info"
                    }
                }, "Normal")
            }
            
            if ($permissions -and $permissions.Count -gt 0) {
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "SendOnBehalf-Berechtigungen für $sourceUser abgerufen: $($permissions.Count) Einträge gefunden." -Color $script:connectedBrush
                }
                
                Log-Action "SendOnBehalf-Berechtigungen für $sourceUser abgerufen: $($permissions.Count) Einträge gefunden."
            }
            else {
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Keine SendOnBehalf-Berechtigungen gefunden."
                }
                
                Log-Action "Keine SendOnBehalf-Berechtigungen für $sourceUser gefunden."
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $errorMsg" -Type "Error"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
            }
            
            Log-Action "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $errorMsg"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
        }
        
        Log-Action "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $errorMsg"
    }
})

# -------------------------------------------------
# Fenster anzeigen
# -------------------------------------------------
$null = $Form.ShowDialog()
