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
# Abschnitt: Throttling Policy Funktionen
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
        if ([string]::IsNullOrEmpty($PolicyName)) {
            Write-DebugMessage "Kein Policy-Name angegeben, verwende 'Global'" -Type "Info"
            $PolicyName = "Global"
        }
        
        Write-DebugMessage "Rufe Throttling Policy ab: $PolicyName" -Type "Info"
        
        # Throttling Policy abrufen
        if ($PolicyName -eq "All") {
            $throttlingPolicy = Get-ThrottlingPolicy -ErrorAction Stop
            Write-DebugMessage "Alle Throttling Policies erfolgreich abgerufen" -Type "Success"
        } else {
            $throttlingPolicy = Get-ThrottlingPolicy -Identity $PolicyName -ErrorAction Stop
            Write-DebugMessage "Throttling Policy erfolgreich abgerufen: $PolicyName" -Type "Success"
        }
        
        if ($throttlingPolicy) {
            # Überprüfen, ob wir nur EWS-bezogene Informationen anzeigen sollen
            if ($ShowEWSOnly) {
                $result = "EWS Throttling Policy Informationen für: $PolicyName`n`n"
                
                # Wenn es sich um eine Sammlung handelt, jede Policy einzeln verarbeiten
                if ($throttlingPolicy -is [System.Array]) {
                    foreach ($policy in $throttlingPolicy) {
                        $result += "Policy: $($policy.Name)`n"
                        $result += "EWSMaxConcurrency: $($policy.EWSMaxConcurrency)`n"
                        $result += "EWSMaxSubscriptions: $($policy.EWSMaxSubscriptions)`n"
                        $result += "EWSFastSearchTimeoutInSeconds: $($policy.EWSFastSearchTimeoutInSeconds)`n"
                        $result += "EWSFindCountLimit: $($policy.EWSFindCountLimit)`n"
                        $result += "EWSMaxConnections: $($policy.EWSMaxConnections)`n"
                        $result += "EWSMaxBatchSize: $($policy.EWSMaxBatchSize)`n"
                        $result += "-----------------------------------------`n"
                    }
                } else {
                    $result += "Policy: $($throttlingPolicy.Name)`n"
                    $result += "EWSMaxConcurrency: $($throttlingPolicy.EWSMaxConcurrency)`n"
                    $result += "EWSMaxSubscriptions: $($throttlingPolicy.EWSMaxSubscriptions)`n"
                    $result += "EWSFastSearchTimeoutInSeconds: $($throttlingPolicy.EWSFastSearchTimeoutInSeconds)`n"
                    $result += "EWSFindCountLimit: $($throttlingPolicy.EWSFindCountLimit)`n"
                    $result += "EWSMaxConnections: $($throttlingPolicy.EWSMaxConnections)`n"
                    $result += "EWSMaxBatchSize: $($throttlingPolicy.EWSMaxBatchSize)`n"
                }
                
                return $result
            }
            elseif ($DetailedView) {
                # Detaillierte Ansicht mit Format-List
                return $throttlingPolicy | Format-List | Out-String
            }
            else {
                # Standard-Tabellenansicht mit den wichtigsten Eigenschaften
                $result = "Throttling Policy Übersicht:`n`n"
                
                if ($throttlingPolicy -is [System.Array]) {
                    $result += $throttlingPolicy | Format-Table Name, 
                        EWSMaxConcurrency, 
                        EWSMaxConnections, 
                        PowerShellMaxConcurrency, 
                        RCAMaxConcurrency, 
                        IsServiceAccount -AutoSize | Out-String
                } else {
                    $result += "Policy: $($throttlingPolicy.Name)`n`n"
                    $result += "Exchange Web Services (EWS):`n"
                    $result += "  - MaxConcurrency: $($throttlingPolicy.EWSMaxConcurrency)`n"
                    $result += "  - MaxConnections: $($throttlingPolicy.EWSMaxConnections)`n`n"
                    $result += "PowerShell:`n"
                    $result += "  - MaxConcurrency: $($throttlingPolicy.PowerShellMaxConcurrency)`n`n"
                    $result += "Remote Connectivity Analyzer (RCA):`n"
                    $result += "  - MaxConcurrency: $($throttlingPolicy.RCAMaxConcurrency)`n`n"
                    $result += "Ist Service-Account: $($throttlingPolicy.IsServiceAccount)`n"
                }
                
                return $result
            }
        } else {
            Write-DebugMessage "Keine Throttling Policy gefunden: $PolicyName" -Type "Warning"
            return "Keine Throttling Policy gefunden: $PolicyName"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Abrufen der Throttling Policy: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Throttling Policy '$PolicyName': $errorMsg"
        return "Fehler beim Abrufen der Throttling Policy: $errorMsg"
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
                return Get-ThrottlingPolicyAction -PolicyName "Global" -ShowEWSOnly
            }
            "PowerShell" {
                # Für Remote PowerShell Throttling
                return Get-ThrottlingPolicyAction -PolicyName "Global" | Where-Object { $_ -match "PowerShell|RCA" }
            }
            "All" {
                # Zeige alle Policies in der Übersicht
                return Get-ThrottlingPolicyAction -PolicyName "All"
            }
            default {
                # Standard-Ansicht der globalen Policy
                return Get-ThrottlingPolicyAction -PolicyName "Global" -DetailedView
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
        
        # EWS Policy abrufen
        $ewsPolicy = Get-ThrottlingPolicyAction -PolicyName "Global" -ShowEWSOnly
        
        # Ergebnis formatieren mit Empfehlungen
        $result = $ewsPolicy
        $result += "`n`nEmpfehlungen für Migrationen:`n"
        $result += "- EWSMaxConcurrency: Sollte für Migrationen auf mindestens 20-50 gesetzt sein`n"
        $result += "- EWSMaxConnections: Sollte für Migrationen auf mindestens 10-20 gesetzt sein`n"
        $result += "- EWSMaxBatchSize: Sollte auf 500-1000 für optimale Migrationsbandbreite gesetzt sein`n`n"
        
        $result += "Hinweis: Die tatsächlichen optimalen Werte hängen von Ihrer Umgebung ab.\n"
        $result += "Zu hohe Werte können zu Ressourcenengpässen führen."
        
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

# Initialisiere die Exchange-Diagnose-Array, falls es noch nicht existiert
if ($null -eq $script:exchangeDiagnostics) {
    $script:exchangeDiagnostics = @(
        [PSCustomObject]@{
            Name = "EWS Throttling Policy für Migrationen"
            Description = "Überprüft die EWS Throttling Policy Einstellungen für optimale Migrations-Performance"
            PowerShellCheck = "Test-EWSThrottlingPolicy"
            Category = "Performance"
        }
    )
} else {
    # Integration in die bestehende Troubleshooting-Diagnose wenn das Array bereits existiert
    # Aktualisiere den ersten Diagnoseeintrag für die Migration EWS Throttling Policy
    if ($script:exchangeDiagnostics.Count -gt 0) {
        $script:exchangeDiagnostics[0].PowerShellCheck = "Test-EWSThrottlingPolicy"
    } else {
        # Oder füge einen neuen Eintrag hinzu, wenn das Array leer ist
        $script:exchangeDiagnostics += [PSCustomObject]@{
            Name = "EWS Throttling Policy für Migrationen"
            Description = "Überprüft die EWS Throttling Policy Einstellungen für optimale Migrations-Performance"
            PowerShellCheck = "Test-EWSThrottlingPolicy"
            Category = "Performance"
        }
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
$tabMailboxAudit      = $Form.FindName("tabMailboxAudit")
$tabTroubleshooting   = $Form.FindName("tabTroubleshooting")  # Neue Tab-Referenz
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
$btnNavAudit             = $Form.FindName("btnNavAudit")
$btnNavAudit2            = $Form.FindName("btnNavAudit2")
$btnNavAudit3            = $Form.FindName("btnNavAudit3")
$btnNavTroubleshooting   = $Form.FindName("btnNavTroubleshooting")  # Neuer Navigationsbutton
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

# Referenzierung der Audit-Tab-Elemente
$txtAuditMailbox         = $Form.FindName("txtAuditMailbox")
$cmbAuditType            = $Form.FindName("cmbAuditType")
$cmbAuditCategory        = $Form.FindName("cmbAuditCategory") # Neue Referenz hinzugefügt
$lblAuditMailbox         = $Form.FindName("lblAuditMailbox")
$btnRunAudit             = $Form.FindName("btnRunAudit")
$txtAuditResult          = $Form.FindName("txtAuditResult")

# Referenzierung der Troubleshooting-Tab-Elemente
$lstDiagnostics         = $Form.FindName("lstDiagnostics")
$txtDiagnosticDesc      = $Form.FindName("txtDiagnosticDesc")
$txtDiagnosticUser      = $Form.FindName("txtDiagnosticUser")
$txtDiagnosticUser2     = $Form.FindName("txtDiagnosticUser2")
$txtDiagnosticEmail     = $Form.FindName("txtDiagnosticEmail")
$btnRunDiagnostic       = $Form.FindName("btnRunDiagnostic")
$btnOpenAdminCenter     = $Form.FindName("btnOpenAdminCenter")
$txtDiagnosticResult    = $Form.FindName("txtDiagnosticResult")

# -------------------------------------------------
# Abschnitt: Hilfsfunktionen für UI
# -------------------------------------------------
function Add-SafeClickHandler {
    param (
        [Parameter(Mandatory = $false)]
        [System.Windows.Controls.Button]$Button,
        
        [Parameter(Mandatory = $true)]
        [scriptblock]$Handler,
        
        [Parameter(Mandatory = $false)]
        [string]$ButtonName = "Nicht angegeben"
    )
    
    if ($null -eq $Button) {
        Write-DebugMessage "Button '$ButtonName' wurde nicht gefunden. Event-Handler wird übersprungen." -Type "Warning"
        Log-Action "UI-Element nicht gefunden: Button '$ButtonName' - Event-Handler wurde nicht angehängt"
        return $false
    }
    
    try {
        $Button.Add_Click($Handler)
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Hinzufügen des Event-Handlers für Button '$ButtonName': $errorMsg" -Type "Error"
        Log-Action "Fehler beim Hinzufügen des Event-Handlers für Button '$ButtonName': $errorMsg"
        return $false
    }
}

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

# Neue Event-Handler für Audit-Navigation
$auditHandler = {
    try {
        # Verstecke alle TabItems
        if ($null -ne $tabContent) {
            foreach ($tab in $tabContent.Items) {
                $tab.Visibility = [System.Windows.Visibility]::Collapsed
            }
        }
        
        # Zeige Mailbox Audit TabItem, falls vorhanden
        if ($null -ne $tabMailboxAudit) {
            $tabMailboxAudit.Visibility = [System.Windows.Visibility]::Visible
            $tabMailboxAudit.IsSelected = $true
        }
        
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Audit-Funktionen gewählt"
        }
        
        # Initialisiere die Audit-Dropdowns falls diese existieren
        if ($null -ne $cmbAuditCategory -and $cmbAuditCategory.Items.Count -gt 0) {
            if ($null -eq $cmbAuditType -or $cmbAuditType.Items.Count -eq 0) {
                # Setze die Standardkategorie und initialisiere die Options
                $cmbAuditCategory.SelectedIndex = 0
                if ($null -ne $cmbAuditCategory.SelectedItem) {
                    try {
                        $selectedCategory = $cmbAuditCategory.SelectedItem.Tag.ToString()
                        Update-AuditOptions -Category $selectedCategory
                    } catch {
                        Write-DebugMessage "Fehler beim Initialisieren der Audit-Optionen: $($_.Exception.Message)" -Type "Error"
                    }
                }
            }
        } else {
            Write-DebugMessage "Audit-Kategorie ComboBox nicht gefunden oder leer" -Type "Warning"
        }
        
        if ($null -ne $txtAuditResult) {
            $txtAuditResult.Text = "Wählen Sie eine Kategorie und einen Informationstyp, dann klicken Sie auf 'Ausführen'."
        }
        
        Log-Action "Navigation zu Audit-Funktionen"
    } catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler bei Navigation zu Audit-Funktionen: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Fehler bei Navigation: $errorMsg"
        }
        
        Log-Action "Fehler bei Navigation zu Audit-Funktionen: $errorMsg"
    }
}

# Sichere Anbindung der Event-Handler für Audit-Navigation
Add-SafeClickHandler -Button $btnNavAudit -Handler $auditHandler -ButtonName "btnNavAudit"
Add-SafeClickHandler -Button $btnNavAudit2 -Handler $auditHandler -ButtonName "btnNavAudit2"
Add-SafeClickHandler -Button $btnNavAudit3 -Handler $auditHandler -ButtonName "btnNavAudit3"

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
# Abschnitt: Mailbox Audit Funktionen
# -------------------------------------------------
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
            Update-GuiText -TextElement $txtStatus -Message "Rufe Informationen ab..."
        }
        
        # Bestimme NavigationType anhand des ausgewählten Dropdown-Eintrags wenn nicht explizit übergeben
        if ([string]::IsNullOrEmpty($NavigationType) -and $null -ne $cmbAuditCategory) {
            $NavigationType = $cmbAuditCategory.SelectedValue.ToString()
        }
        
        Write-DebugMessage "Führe Mailbox-Audit aus. NavigationType: $NavigationType, InfoType: $InfoType, Mailbox: $Mailbox" -Type "Info"
        
        switch ($NavigationType) {
            "user" { # User-spezifische Abfragen
                if ([string]::IsNullOrEmpty($Mailbox)) {
                    throw "Für Benutzerpostfach-Abfragen wird eine Mailbox-Adresse benötigt."
                }
                
                # Versuche, die Mailbox zu validieren
                try {
                    $mailboxInfo = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
                    $mailboxType = $mailboxInfo.RecipientTypeDetails
                    Write-DebugMessage "Mailbox gefunden: $Mailbox (Typ: $mailboxType)" -Type "Info"
                } 
                catch {
                    throw "Die angegebene Mailbox konnte nicht gefunden werden: $($_.Exception.Message)"
                }
                
                switch ($InfoType) {
                    0 { # Basis Mailbox-Informationen
                        $result = Get-Mailbox -Identity $Mailbox | 
                                 Select-Object DisplayName, Alias, PrimarySmtpAddress, RecipientTypeDetails, ExchangeGuid, 
                                             WhenMailboxCreated, IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota |
                                 Format-List | Out-String
                        return $result
                    }
                    
                    1 { # Detaillierte Informationen (Format-List)
                        $result = Get-Mailbox -Identity $Mailbox | Format-List | Out-String
                        return $result
                    }
                    
                    2 { # Postfachstatistiken (Größe, Elementanzahl)
                        $result = Get-MailboxStatistics -Identity $Mailbox | 
                                  Select-Object DisplayName, TotalItemSize, ItemCount, LastLogonTime, LastLogoffTime, 
                                              StorageLimitStatus, IsArchiveMailbox, DeletedItemCount |
                                  Format-List | Out-String
                        return $result
                    }
                    
                    3 { # Ordnerstatistiken
                        $result = "Ordnerstatistiken für $Mailbox`n`n"
                        $folderStats = Get-MailboxFolderStatistics -Identity $Mailbox -ErrorAction Stop | 
                                      Select-Object Name, FolderType, ItemsInFolder, FolderSize | 
                                      Sort-Object @{Expression={$_.FolderSize.ToString() -replace '[^\d]', ''}; Descending=$true}
                        
                        # Formatierte Ausgabe
                        $result += $folderStats | Format-Table -AutoSize | Out-String
                        return $result
                    }
                    
                    4 { # Kalenderberechtigungen
                        $result = "Kalenderberechtigungen für $Mailbox`n`n"
                        
                        # Versuche deutsche und englische Kalenderordnernamen
                        $calendarFound = $false
                        try {
                            $permissions = Get-MailboxFolderPermission -Identity "$($Mailbox):\Kalender" -ErrorAction Stop
                            $result += "Kalenderordner: 'Kalender'`n`n"
                            $calendarFound = $true
                        } 
                        catch {
                            try {
                                $permissions = Get-MailboxFolderPermission -Identity "$($Mailbox):\Calendar" -ErrorAction Stop
                                $result += "Kalenderordner: 'Calendar'`n`n"
                                $calendarFound = $true
                            } 
                            catch {
                                throw "Kalenderordner konnte nicht gefunden werden. Weder 'Kalender' noch 'Calendar' sind zugänglich."
                            }
                        }
                        
                        if ($calendarFound) {
                            $result += $permissions | Format-Table -AutoSize | Out-String
                        } 
                        else {
                            $result += "Keine Kalenderberechtigungen gefunden."
                        }
                        
                        return $result
                    }
                    
                    5 { # Postfachberechtigungen
                        $result = "Postfachberechtigungen für $Mailbox`n`n"
                        
                        $mbxPermissions = Get-MailboxPermission -Identity $Mailbox | Where-Object { 
                            $_.User -notlike "NT AUTHORITY\SELF" -and 
                            $_.User -notlike "S-1-5*" -and 
                            $_.IsInherited -eq $false 
                        }
                        
                        if ($mbxPermissions) {
                            $result += "Vollzugriff-Berechtigungen:`n"
                            $result += $mbxPermissions | Select-Object User, AccessRights, Deny | Format-Table -AutoSize | Out-String
                            $result += "`n"
                        } 
                        else {
                            $result += "Keine Vollzugriff-Berechtigungen gefunden.`n`n"
                        }
                        
                        # SendAs
                        $sendAsPermissions = Get-RecipientPermission -Identity $Mailbox | Where-Object {
                            $_.Trustee -notlike "NT AUTHORITY\SELF" -and
                            $_.Trustee -notlike "S-1-5*" -and
                            $_.IsInherited -eq $false
                        }
                        
                        if ($sendAsPermissions) {
                            $result += "SendAs-Berechtigungen:`n"
                            $result += $sendAsPermissions | Select-Object Trustee, AccessRights | Format-Table -AutoSize | Out-String
                            $result += "`n"
                        } 
                        else {
                            $result += "Keine SendAs-Berechtigungen gefunden.`n`n"
                        }
                        
                        # SendOnBehalf
                        $mailboxObj = Get-Mailbox -Identity $Mailbox
                        if ($mailboxObj.GrantSendOnBehalfTo -and $mailboxObj.GrantSendOnBehalfTo.Count -gt 0) {
                            $result += "SendOnBehalf-Berechtigungen:`n"
                            $result += $mailboxObj.GrantSendOnBehalfTo | ForEach-Object {
                                [PSCustomObject]@{
                                    User = $_
                                    Permission = "SendOnBehalf"
                                }
                            } | Format-Table -AutoSize | Out-String
                        } 
                        else {
                            $result += "Keine SendOnBehalf-Berechtigungen gefunden.`n"
                        }
                        
                        return $result
                    }
                    
                    6 { # Clientzugriffseinstellungen (CAS)
                        $result = "Clientzugriffseinstellungen für $Mailbox`n`n"
                        
                        try {
                            $casSettings = Get-CasMailbox -Identity $Mailbox -ErrorAction Stop
                            $result += $casSettings | Format-List | Out-String
                        } 
                        catch {
                            $result = "CAS-Informationen konnten nicht abgerufen werden: $($_.Exception.Message)`n"
                            $result += "Dieser Befehl wird möglicherweise in Ihrer Exchange-Version nicht unterstützt."
                        }
                        
                        return $result
                    }
                    
                    7 { # Mobile Geräteinformationen
                        $result = "Mobile Geräteinformationen für $Mailbox`n`n"
                        
                        try {
                            $mobileDevices = Get-MobileDeviceStatistics -Mailbox $Mailbox -ErrorAction Stop
                            
                            if ($mobileDevices -and $mobileDevices.Count -gt 0) {
                                $result += $mobileDevices | 
                                          Select-Object DeviceId, DeviceModel, DeviceOS, DeviceType,
                                                      FirstSyncTime, LastSuccessSync, Status |
                                          Format-Table -AutoSize | Out-String
                            } 
                            else {
                                $result += "Keine mobilen Geräte für dieses Postfach gefunden."
                            }
                        } 
                        catch {
                            $result = "Mobile Geräteinformationen konnten nicht abgerufen werden: $($_.Exception.Message)`n"
                            $result += "Dieser Befehl wird möglicherweise in Ihrer Exchange-Version nicht unterstützt."
                        }
                        
                        return $result
                    }
                    
                    default {
                        return "Ungültiger Informationstyp für Benutzerpostfächer ausgewählt."
                    }
                }
            }
            
            "general" { # Allgemeine Postfach-Abfragen
                switch ($InfoType) {
                    0 { # Alle Postfächer anzeigen
                        Write-DebugMessage "Rufe alle Postfächer ab..." -Type "Info"
                        $result = "Liste aller Postfächer:`n`n"
                        
                        $mailboxes = Get-Mailbox -ResultSize Unlimited | 
                                    Select-Object DisplayName, Alias, PrimarySmtpAddress, RecipientTypeDetails
                        
                        if ($mailboxes.Count -gt 0) {
                            $result += $mailboxes | Format-Table -AutoSize | Out-String
                            $result += "`nAnzahl gefundener Postfächer: $($mailboxes.Count)"
                        } 
                        else {
                            $result += "Keine Postfächer gefunden."
                        }
                        
                        return $result
                    }
                    
                    1 { # Shared Mailboxen anzeigen
                        Write-DebugMessage "Rufe Shared Mailboxen ab..." -Type "Info"
                        $result = "Liste der Shared Mailboxen:`n`n"
                        
                        $sharedMailboxes = Get-Mailbox -ResultSize Unlimited | 
                                          Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"} |
                                          Select-Object DisplayName, Alias, PrimarySmtpAddress
                        
                        if ($sharedMailboxes.Count -gt 0) {
                            $result += $sharedMailboxes | Format-Table -AutoSize | Out-String
                            $result += "`nAnzahl gefundener Shared Mailboxen: $($sharedMailboxes.Count)"
                        } 
                        else {
                            $result += "Keine Shared Mailboxen gefunden."
                        }
                        
                        return $result
                    }
                    
                    2 { # Raumpostfächer anzeigen
                        Write-DebugMessage "Rufe Raumpostfächer ab..." -Type "Info"
                        $result = "Liste der Raumpostfächer:`n`n"
                        
                        $roomMailboxes = Get-Mailbox -ResultSize Unlimited | 
                                        Where-Object {$_.RecipientTypeDetails -eq "RoomMailbox"} |
                                        Select-Object DisplayName, Alias, PrimarySmtpAddress, ResourceCapacity
                        
                        if ($roomMailboxes.Count -gt 0) {
                            $result += $roomMailboxes | Format-Table -AutoSize | Out-String
                            $result += "`nAnzahl gefundener Raumpostfächer: $($roomMailboxes.Count)"
                        } 
                        else {
                            $result += "Keine Raumpostfächer gefunden."
                        }
                        
                        return $result
                    }
                    
                    3 { # Ressourcenpostfächer anzeigen
                        Write-DebugMessage "Rufe Ressourcenpostfächer ab..." -Type "Info"
                        $result = "Liste der Ressourcenpostfächer (Equipment):`n`n"
                        
                        $equipmentMailboxes = Get-Mailbox -ResultSize Unlimited | 
                                            Where-Object {$_.RecipientTypeDetails -eq "EquipmentMailbox"} |
                                            Select-Object DisplayName, Alias, PrimarySmtpAddress
                        
                        if ($equipmentMailboxes.Count -gt 0) {
                            $result += $equipmentMailboxes | Format-Table -AutoSize | Out-String
                            $result += "`nAnzahl gefundener Ressourcenpostfächer: $($equipmentMailboxes.Count)"
                        } 
                        else {
                            $result += "Keine Ressourcenpostfächer gefunden."
                        }
                        
                        return $result
                    }
                    
                    4 { # Empfänger nach Typ gruppieren
                        Write-DebugMessage "Gruppiere Postfächer nach Typ..." -Type "Info"
                        $result = "Postfächer nach Typ gruppiert:`n`n"
                        
                        $mailboxes = Get-Mailbox -ResultSize Unlimited
                        $grouped = $mailboxes | Group-Object RecipientTypeDetails
                        
                        foreach ($group in $grouped) {
                            $result += "---- $($group.Name) ($($group.Count) Postfächer) ----`n"
                            $result += $group.Group | 
                                     Select-Object DisplayName, Alias, PrimarySmtpAddress | 
                                     Format-Table -AutoSize | Out-String
                            $result += "`n"
                        }
                        
                        return $result
                    }
                    
                    5 { # Postfächer nach Größe sortieren
                        Write-DebugMessage "Analysiere Postfachgrößen..." -Type "Info"
                        
                        if ($null -ne $txtStatus) {
                            Update-GuiText -TextElement $txtStatus -Message "Rufe Mailbox-Statistiken ab (kann einige Minuten dauern)..."
                        }
                        
                        $result = "Postfächer nach Größe sortiert (größte zuerst):`n`n"
                        $mailboxesWithStats = @()
                        
                        # Hole alle Postfächer
                        $mailboxes = Get-Mailbox -ResultSize Unlimited
                        $count = $mailboxes.Count
                        $i = 0
                        
                        foreach ($mailbox in $mailboxes) {
                            $i++
                            if ($i % 10 -eq 0) {
                                if ($null -ne $txtStatus) {
                                    $percent = [math]::Round(($i / $count) * 100)
                                    Update-GuiText -TextElement $txtStatus -Message "Verarbeite Postfächer: $i von $count ($percent%)"
                                }
                            }
                            
                            try {
                                $stats = Get-MailboxStatistics -Identity $mailbox.Identity -ErrorAction SilentlyContinue
                                if ($stats) {
                                    $mbxInfo = [PSCustomObject]@{
                                        DisplayName = $mailbox.DisplayName
                                        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                                        TotalItemSize = $stats.TotalItemSize
                                        TotalItemSizeMB = if ($stats.TotalItemSize) {
                                            [regex]::Match($stats.TotalItemSize.ToString(), "([0-9,.]+)").Groups[1].Value
                                        } else { "0" }
                                        ItemCount = $stats.ItemCount
                                        Type = $mailbox.RecipientTypeDetails
                                        LastLogon = $stats.LastLogonTime
                                    }
                                    $mailboxesWithStats += $mbxInfo
                                }
                            } 
                            catch {
                                # Ignoriere Fehler bei einzelnen Postfächern
                                Write-DebugMessage "Fehler beim Abrufen der Statistiken für $($mailbox.DisplayName): $($_.Exception.Message)" -Type "Warning"
                            }
                        }
                        
                        # Sortiere nach Größe und formatiere Ausgabe
                        $sortedMailboxes = $mailboxesWithStats | Sort-Object -Property @{Expression={[double]$_.TotalItemSizeMB}; Descending=$true}
                        $result += $sortedMailboxes | 
                                 Select-Object DisplayName, PrimarySmtpAddress, TotalItemSize, ItemCount, Type | 
                                 Format-Table -AutoSize | Out-String
                        
                        return $result
                    }
                    
                    6 { # Postfächer nach letzter Anmeldung sortieren
                        Write-DebugMessage "Analysiere letzte Anmeldezeitpunkte..." -Type "Info"
                        
                        if ($null -ne $txtStatus) {
                            Update-GuiText -TextElement $txtStatus -Message "Rufe Anmeldeinformationen ab (kann einige Minuten dauern)..."
                        }

                        $result = "Postfächer nach letzter Anmeldung sortiert (neueste zuerst):`n`n"
                        $mailboxesWithStats = @()
                        
                        # Hole alle Postfächer
                        $mailboxes = Get-Mailbox -ResultSize Unlimited
                        $count = $mailboxes.Count
                        $i = 0
                        
                        foreach ($mailbox in $mailboxes) {
                            $i++
                            if ($i % 10 -eq 0) {
                                if ($null -ne $txtStatus) {
                                    $percent = [math]::Round(($i / $count) * 100)
                                    Update-GuiText -TextElement $txtStatus -Message "Verarbeite Postfächer: $i von $count ($percent%)"
                                }
                            }
                            
                            try {
                                $stats = Get-MailboxStatistics -Identity $mailbox.Identity -ErrorAction SilentlyContinue
                                if ($stats) {
                                    $mbxInfo = [PSCustomObject]@{
                                        DisplayName = $mailbox.DisplayName
                                        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                                        LastLogonTime = $stats.LastLogonTime
                                        ItemCount = $stats.ItemCount
                                        Type = $mailbox.RecipientTypeDetails
                                    }
                                    $mailboxesWithStats += $mbxInfo
                                }
                            } 
                            catch {
                                # Ignoriere Fehler bei einzelnen Postfächern
                                Write-DebugMessage "Fehler beim Abrufen der Statistiken für $($mailbox.DisplayName): $($_.Exception.Message)" -Type "Warning"
                            }
                        }
                        
                        # Sortiere nach letzter Anmeldung und formatiere Ausgabe
                        $sortedMailboxes = $mailboxesWithStats | Sort-Object LastLogonTime -Descending
                        $result += $sortedMailboxes | 
                                 Select-Object DisplayName, PrimarySmtpAddress, LastLogonTime, Type | 
                                 Format-Table -AutoSize | Out-String
                        
                        return $result
                    }
                    
                    7 { # Externe E-Mail-Weiterleitungen anzeigen
                        Write-DebugMessage "Suche nach externen E-Mail-Weiterleitungen..." -Type "Info"
                        $result = "Postfächer mit externen E-Mail-Weiterleitungen:`n`n"
                        
                        # Hole alle Postfächer mit Weiterleitungen
                        $mailboxesWithForwarding = Get-Mailbox -ResultSize Unlimited | 
                            Where-Object {$_.ForwardingAddress -ne $null -or $_.ForwardingSmtpAddress -ne $null}
                        
                        if ($mailboxesWithForwarding.Count -eq 0) {
                            $result += "Keine Postfächer mit externen Weiterleitungen gefunden."
                        } 
                        else {
                            $forwardingList = @()
                            foreach ($mailbox in $mailboxesWithForwarding) {
                                $forwardingInfo = [PSCustomObject]@{
                                    DisplayName = $mailbox.DisplayName
                                    PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                                    ForwardingAddress = $mailbox.ForwardingAddress
                                    ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                                    DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                                }
                                $forwardingList += $forwardingInfo
                            }
                            
                            $result += $forwardingList | Format-Table -AutoSize | Out-String
                        }
                        
                        return $result
                    }
                    
                    8 { # Postfächer mit aktiviertem Litigation Hold
                        Write-DebugMessage "Suche nach Postfächern mit Litigation Hold..." -Type "Info"
                        $result = "Postfächer mit aktiviertem Litigation Hold:`n`n"
                        
                        # Hole alle Postfächer mit Litigation Hold
                        $mailboxesWithLitigationHold = Get-Mailbox -ResultSize Unlimited | 
                            Where-Object {$_.LitigationHoldEnabled -eq $true}
                        
                        if ($mailboxesWithLitigationHold.Count -eq 0) {
                            $result += "Keine Postfächer mit aktiviertem Litigation Hold gefunden."
                        } 
                        else {
                            $result += $mailboxesWithLitigationHold | 
                                      Select-Object DisplayName, PrimarySmtpAddress, LitigationHoldDate, LitigationHoldOwner, LitigationHoldDuration |
                                      Format-Table -AutoSize | Out-String
                        }
                        
                        return $result
                    }
                    
                    9 { # Postfächer über Speicherlimit
                        Write-DebugMessage "Suche nach Postfächern über dem Speicherlimit..." -Type "Info"
                        
                        if ($null -ne $txtStatus) {
                            Update-GuiText -TextElement $txtStatus -Message "Rufe Mailbox-Statistiken ab (kann einige Zeit dauern)..."
                        }
                        
                        $result = "Postfächer nahe oder über dem Speicherlimit:`n`n"
                        $mailboxesOverLimit = @()
                        
                        # Hole alle Postfächer
                        $mailboxes = Get-Mailbox -ResultSize Unlimited
                        $count = $mailboxes.Count
                        $i = 0
                        
                        foreach ($mailbox in $mailboxes) {
                            $i++
                            if ($i % 10 -eq 0) {
                                if ($null -ne $txtStatus) {
                                    $percent = [math]::Round(($i / $count) * 100)
                                    Update-GuiText -TextElement $txtStatus -Message "Verarbeite Postfächer: $i von $count ($percent%)"
                                }
                            }
                            
                            try {
                                $stats = Get-MailboxStatistics -Identity $mailbox.Identity -ErrorAction SilentlyContinue
                                if ($stats) {
                                    $warningQuota = $mailbox.IssueWarningQuota
                                    $prohibitSendQuota = $mailbox.ProhibitSendQuota
                                    $totalSize = $stats.TotalItemSize
                                    $storageStatus = $stats.StorageLimitStatus
                                    
                                    # Prüfe ob der Status problematisch ist oder ob wir uns dem Limit nähern
                                    if ($storageStatus -eq "IssueWarning" -or $storageStatus -eq "ProhibitSend" -or $storageStatus -eq "MailboxDisabled") {
                                        $mailboxesOverLimit += [PSCustomObject]@{
                                            DisplayName = $mailbox.DisplayName
                                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                                            TotalItemSize = $totalSize
                                            StorageLimitStatus = $storageStatus
                                            WarningQuota = $warningQuota
                                            ProhibitSendQuota = $prohibitSendQuota
                                        }
                                    }
                                }
                            } 
                            catch {
                                # Ignoriere Fehler bei einzelnen Postfächern
                                Write-DebugMessage "Fehler beim Abrufen der Statistiken für $($mailbox.DisplayName): $($_.Exception.Message)" -Type "Warning"
                            }
                        }
                        
                        # Sortiere nach Status und formatiere Ausgabe
                        if ($mailboxesOverLimit.Count -eq 0) {
                            $result += "Keine Postfächer nahe oder über dem Speicherlimit gefunden."
                        } 
                        else {
                            $result += $mailboxesOverLimit | Sort-Object StorageLimitStatus -Descending |
                                      Format-Table -AutoSize | Out-String
                        }
                        
                        return $result
                    }
                    
                    default {
                        return "Ungültiger Informationstyp für allgemeine Postfachabfragen ausgewählt."
                    }
                }
            }
            
            "inactive" { # Inaktive Benutzer
                Write-DebugMessage "Starte Inaktive-Benutzer-Analyse..." -Type "Info"
                
                $daysBack = 180  # Standardwert
                
                switch ($InfoType) {
                    0 { $daysBack = 180 } # 180 Tage
                    1 { $daysBack = 90 }  # 90 Tage
                    2 { $daysBack = 30 }  # 30 Tage
                    default { $daysBack = 180 }
                }
                
                $result = "# Analyse inaktiver Benutzer (letzte $daysBack Tage) #`n`n"
                
                try {
                    # Prüfe, ob MSOnline verfügbar ist
                    if (-not (Get-Module -Name MSOnline -ListAvailable)) {
                        $result += "MSOnline-Modul nicht installiert. Bitte installieren Sie das Modul mit:`n"
                        $result += "Install-Module -Name MSOnline -Force -AllowClobber`n`n"
                        $result += "Alternativ können Sie die vereinfachte Analyse verwenden:"
                        
                        # Vereinfachte Analyse basierend auf lokalen Exchange-Daten
                        $result += "`n`nVereinfachte Analyse (nur Exchange-Daten):`n`n"
                        
                        $inactiveUsers = @()
                        $mailboxes = Get-Mailbox -ResultSize Unlimited
                        
                        foreach ($mailbox in $mailboxes) {
                            try {
                                $stats = Get-MailboxStatistics -Identity $mailbox.Identity -ErrorAction SilentlyContinue
                                if ($stats -and $stats.LastLogonTime) {
                                    $lastLogon = $stats.LastLogonTime
                                    $daysSinceLogon = (New-TimeSpan -Start $lastLogon -End (Get-Date)).Days
                                    
                                    if ($daysSinceLogon -gt $daysBack) {
                                        $inactiveUsers += [PSCustomObject]@{
                                            DisplayName = $mailbox.DisplayName
                                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                                            LastLogon = $lastLogon
                                            DaysSinceLogon = $daysSinceLogon
                                        }
                                    }
                                }
                            } 
                            catch {
                                # Ignoriere Fehler
                            }
                        }
                        
                        if ($inactiveUsers.Count -gt 0) {
                            $result += $inactiveUsers | 
                                      Sort-Object DaysSinceLogon -Descending |
                                      Format-Table -AutoSize | Out-String
                            $result += "`nAnzahl inaktiver Benutzer: $($inactiveUsers.Count) von $($mailboxes.Count) Postfächern"
                        } 
                        else {
                            $result += "Keine inaktiven Benutzer gefunden."
                        }
                        
                        return $result
                    }
                    
                    # MSOnline ist verfügbar, volle Analyse durchführen
                    # Verbindung zu MSOnline herstellen, falls nötig
                    try {
                        $null = Get-MsolDomain -ErrorAction Stop
                        Write-DebugMessage "MSOnline-Verbindung bereits aktiv" -Type "Info"
                    } 
                    catch {
                        Write-DebugMessage "Stelle Verbindung zu MSOnline her..." -Type "Info"
                        Connect-MsolService -ErrorAction Stop
                    }
                    
                    # Aktive Benutzer abrufen
                    $allMsolUsers = Get-MsolUser -All -EnabledFilter EnabledOnly | 
                                    Where-Object { $_.UserType -eq "Member" }
                    
                    $totalUsers = $allMsolUsers.Count
                    $result += "Anzahl aktiver Azure AD Benutzer: $totalUsers`n`n"
                    
                    # Anmeldungen in den letzten X Tagen abrufen
                    $startDate = (Get-Date).AddDays(-$daysBack).ToString('MM/dd/yyyy')
                    $endDate = (Get-Date).ToString('MM/dd/yyyy')
                    $result += "Suche nach Anmeldungen zwischen $startDate und $endDate...`n`n"
                    
                    try {
                        # Audit-Log abfragen
                        $loggedOnUsers = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn, PasswordLogonInitialAuthUsingPassword, UserLoginFailed -ResultSize 5000
                        
                        if ($loggedOnUsers) {
                            $loggedOnUserIds = $loggedOnUsers | Select-Object -ExpandProperty UserIds -Unique
                            $result += "Gefundene Anmeldungen: $($loggedOnUserIds.Count)`n`n"
                            
                            # Inaktive Benutzer identifizieren
                            $inactiveUsers = $allMsolUsers | Where-Object { $loggedOnUserIds -notcontains $_.UserPrincipalName }
                            
                            $result += "Inaktive Benutzer (keine Anmeldung in den letzten $daysBack Tagen):`n`n"
                            if ($inactiveUsers.Count -gt 0) {
                                $result += $inactiveUsers | 
                                          Select-Object DisplayName, UserPrincipalName, LastPasswordChangeTimestamp, WhenCreated |
                                          Format-Table -AutoSize | Out-String
                                $result += "`nAnzahl inaktiver Benutzer: $($inactiveUsers.Count) von $totalUsers"
                            } 
                            else {
                                $result += "Keine inaktiven Benutzer gefunden."
                            }
                        } 
                        else {
                            $result += "Keine Anmeldungen im angegebenen Zeitraum gefunden.`nMöglicherweise ist das Audit-Log nicht aktiviert oder hat keine Daten."
                        }
                    } 
                    catch {
                        $result += "Fehler beim Abfragen des Audit-Logs: $($_.Exception.Message)`n`n"
                        $result += "Alternative Berechnung basierend auf LastPasswordChangeTimestamp:`n`n"
                        
                        # Alternative Berechnung basierend auf Passwortänderungsdatum
                        $cutoffDate = (Get-Date).AddDays(-$daysBack)
                        $inactiveUsers = $allMsolUsers | Where-Object { 
                            $_.LastPasswordChangeTimestamp -lt $cutoffDate -and
                            $_.WhenCreated -lt $cutoffDate
                        }
                        
                        if ($inactiveUsers.Count -gt 0) {
                            $result += $inactiveUsers | 
                                      Select-Object DisplayName, UserPrincipalName, LastPasswordChangeTimestamp, WhenCreated |
                                      Format-Table -AutoSize | Out-String
                            $result += "`nAnzahl potentiell inaktiver Benutzer: $($inactiveUsers.Count) von $totalUsers"
                            $result += "`n(Basierend auf letzter Passwortänderung vor $daysBack Tagen)"
                        } 
                        else {
                            $result += "Keine potenziell inaktiven Benutzer gefunden."
                        }
                    }
                    
                    return $result
                } 
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Fehler bei der Inaktiven-Benutzer-Analyse: $errorMsg" -Type "Error"
                    return "Fehler bei der Analyse inaktiver Benutzer: $errorMsg"
                }
            }
            
            default {
                return "Ungültige Navigationsart ausgewählt. Bitte wählen Sie eine der vordefinierten Kategorien aus."
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

# -------------------------------------------------
# Änderungen in den Event Handlern für Audit-Funktionen
# -------------------------------------------------

# In den Navigationselementen wird jetzt ein einzelner Button verwendet, und die Optionen kommen
# aus einem neuen DropDown (cmbAuditCategory), welches die Navigationskategorie festlegt.
# Das alte cmbAuditType enthält dann die spezifischen Optionen für die gewählte Kategorie.

# Event-Handler für das Ändern der Audit-Kategorie
function Update-AuditOptions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Category
    )
    
    try {
        # Leere die bisherigen Optionen
        $cmbAuditType.Items.Clear()
        
        # Befülle das Dropdown je nach Kategorie
        switch ($Category) {
            "user" { # Benutzerpostfach-Abfragen
                $cmbAuditType.Items.Add("Basis Mailbox-Informationen")
                $cmbAuditType.Items.Add("Detaillierte Informationen (Format-List)")
                $cmbAuditType.Items.Add("Postfachstatistiken (Größe, Elementanzahl)")
                $cmbAuditType.Items.Add("Ordnerstatistiken")
                $cmbAuditType.Items.Add("Kalenderberechtigungen")
                $cmbAuditType.Items.Add("Postfachberechtigungen")
                $cmbAuditType.Items.Add("Clientzugriffseinstellungen (CAS)")
                $cmbAuditType.Items.Add("Mobile Geräteinformationen")
                
                # Mailbox-Eingabefeld anzeigen
                if ($null -ne $lblAuditMailbox) {
                    $lblAuditMailbox.Visibility = [System.Windows.Visibility]::Visible
                }
                if ($null -ne $txtAuditMailbox) {
                    $txtAuditMailbox.Visibility = [System.Windows.Visibility]::Visible
                }
                if ($null -ne $txtAuditResult) {
                    $txtAuditResult.Text = "Geben Sie eine Mailbox-Adresse ein und wählen Sie die gewünschten Informationen."
                }
            }
            "general" { # Allgemeine Postfachabfragen
                $cmbAuditType.Items.Add("Alle Postfächer anzeigen")
                $cmbAuditType.Items.Add("Shared Mailboxen anzeigen")
                $cmbAuditType.Items.Add("Raumpostfächer anzeigen")
                $cmbAuditType.Items.Add("Ressourcenpostfächer anzeigen")
                $cmbAuditType.Items.Add("Empfänger nach Typ gruppieren")
                $cmbAuditType.Items.Add("Postfächer nach Größe sortieren")
                $cmbAuditType.Items.Add("Postfächer nach letzter Anmeldung sortieren")
                $cmbAuditType.Items.Add("Externe E-Mail-Weiterleitungen anzeigen")
                $cmbAuditType.Items.Add("Postfächer mit aktiviertem Litigation Hold")
                $cmbAuditType.Items.Add("Postfächer über Speicherlimit")
                
                # Mailbox-Eingabefeld ausblenden
                if ($null -ne $lblAuditMailbox) {
                    $lblAuditMailbox.Visibility = [System.Windows.Visibility]::Collapsed
                }
                if ($null -ne $txtAuditMailbox) {
                    $txtAuditMailbox.Visibility = [System.Windows.Visibility]::Collapsed
                    $txtAuditMailbox.Text = ""  # Feld leeren
                }
                if ($null -ne $txtAuditResult) {
                    $txtAuditResult.Text = "Wählen Sie den gewünschten Informationstyp und klicken Sie auf 'Ausführen'."
                }
            }
            "inactive" { # Inaktive Benutzer
                $cmbAuditType.Items.Add("Inaktive Benutzer (180 Tage)")
                $cmbAuditType.Items.Add("Inaktive Benutzer (90 Tage)")
                $cmbAuditType.Items.Add("Inaktive Benutzer (30 Tage)")
                
                # Mailbox-Eingabefeld ausblenden
                if ($null -ne $lblAuditMailbox) {
                    $lblAuditMailbox.Visibility = [System.Windows.Visibility]::Collapsed
                }
                if ($null -ne $txtAuditMailbox) {
                    $txtAuditMailbox.Visibility = [System.Windows.Visibility]::Collapsed
                    $txtAuditMailbox.Text = ""  # Feld leeren
                }
                if ($null -ne $txtAuditResult) {
                    $txtAuditResult.Text = "Diese Funktion analysiert Benutzeraccounts, die sich seit längerer Zeit nicht angemeldet haben."
                }
            }
        }
        
        # Erste Option auswählen
        if ($cmbAuditType.Items.Count > 0) {
            $cmbAuditType.SelectedIndex = 0
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Aktualisieren der Audit-Optionen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Aktualisieren der Audit-Optionen für Kategorie $Category - $errorMsg"
    }
}

# Event-Handler für Audit Run-Button
if ($null -ne $btnRunAudit) {
    $btnRunAudit.Add_Click({
        try {
            if (-not $script:isConnected) {
                Write-DebugMessage "Mailbox-Audit ausführen: Benutzer ist nicht verbunden." -Type "Warning"
                [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            $mailbox = $txtAuditMailbox.Text.Trim()
            $infoType = $cmbAuditType.SelectedIndex
            $navigationType = $cmbAuditCategory.SelectedValue.ToString()
            
            # Spezielle Validierung für User-Abfragen
            if ($navigationType -eq "user" -and [string]::IsNullOrEmpty($mailbox)) {
                Write-DebugMessage "Mailbox-Audit ausführen: Keine Mailbox angegeben für User-Abfrage" -Type "Warning"
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Bitte Mailbox angeben."
                }
                [System.Windows.MessageBox]::Show("Bitte geben Sie eine Mailbox-ID an.", "Eingabe fehlt", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            # Bei bestimmten allgemeinen Postfachabfragen Warnhinweis anzeigen
            if ($navigationType -eq "general" && ($infoType -eq 0 || $infoType -eq 4 || $infoType -eq 5 || $infoType -eq 6)) {
                $confirmResult = [System.Windows.MessageBox]::Show(
                    "Diese Abfrage verarbeitet viele Postfächer und kann längere Zeit dauern. Fortfahren?",
                    "Massenabfrage bestätigen",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Question)
                    
                if ($confirmResult -eq [System.Windows.MessageBoxResult]::No) {
                    Write-DebugMessage "Massenabfrage vom Benutzer abgebrochen" -Type "Info"
                    if ($null -ne $txtStatus) {
                        Update-GuiText -TextElement $txtStatus -Message "Abfrage abgebrochen."
                    }
                    return
                }
            }
            
            # Spezielle Behandlung für Inaktive Benutzer
            if ($navigationType -eq "inactive") {
                try {
                    # MSOnline-Modul prüfen
                    if (-not (Test-ModuleInstalled -ModuleName "MSOnline")) {
                        $installMsol = [System.Windows.MessageBox]::Show(
                            "Für die vollständige Analyse inaktiver Benutzer wird das MSOnline-Modul empfohlen. Möchten Sie es jetzt installieren?", 
                            "Modul empfohlen", 
                            [System.Windows.MessageBoxButton]::YesNo, 
                            [System.Windows.MessageBoxImage]::Question)
                        
                        if ($installMsol -eq [System.Windows.MessageBoxResult]::Yes) {
                            if ($null -ne $txtStatus) {
                                Update-GuiText -TextElement $txtStatus -Message "Installiere MSOnline Modul..."
                            }
                            Install-Module -Name MSOnline -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                            Import-Module -Name MSOnline -ErrorAction Stop
                        }
                    }
                } catch {
                    # Fehler beim Installieren ignorieren, da die Funktion auch ohne MSOnline arbeitet
                    $errorMsg = $_.Exception.Message
                    Write-DebugMessage "Hinweis: MSOnline-Modul konnte nicht installiert werden: $errorMsg" -Type "Warning"
                    if ($null -ne $txtStatus) {
                        Update-GuiText -TextElement $txtStatus -Message "MSOnline-Modul nicht verfügbar, verwende alternative Analyse."
                    }
                }
            }
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Rufe Informationen ab..."
            }
            
            # Informationsabruf mit Fortschrittsanzeige
            $txtAuditResult.Text = "Informationen werden abgerufen..."
            $result = Get-FormattedMailboxInfo -Mailbox $mailbox -InfoType $infoType -NavigationType $navigationType
            
            # Ergebnis anzeigen
            if ($result) {
                $txtAuditResult.Text = $result
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Informationen erfolgreich abgerufen." -Color $script:connectedBrush
                }
                Log-Action "Informationen für $navigationType (Typ $infoType) erfolgreich abgerufen."
            }
            else {
                $txtAuditResult.Text = "Keine Informationen gefunden oder Fehler aufgetreten."
                if ($null -ne $txtStatus) {
                    Update-GuiText -TextElement $txtStatus -Message "Keine Informationen gefunden."
                }
                Log-Action "Keine Informationen für $navigationType (Typ $infoType) gefunden."
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler beim Ausführen der Abfrage: $errorMsg" -Type "Error"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Fehler: $errorMsg"
            }
            
            $txtAuditResult.Text = "Fehler beim Abrufen der Informationen: $errorMsg"
            Log-Action "Fehler beim Ausführen der Abfrage: $errorMsg"
        }
    })
} else {
    Write-DebugMessage "Audit Run Button nicht gefunden" -Type "Warning"
    Log-Action "UI-Element nicht gefunden: btnRunAudit - Event-Handler wurde nicht angehängt"
}

# Event-Handler für Audit-Kategorie ComboBox
if ($null -ne $cmbAuditCategory) {
    $cmbAuditCategory.Add_SelectionChanged({
        try {
            if ($null -ne $cmbAuditCategory.SelectedItem) {
                $selectedCategory = $cmbAuditCategory.SelectedItem.Tag.ToString()
                Update-AuditOptions -Category $selectedCategory
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage "Fehler bei Audit-Kategorie Änderung: $errorMsg" -Type "Error"
            Log-Action "Fehler bei Audit-Kategorie Änderung: $errorMsg"
        }
    })
} else {
    Write-DebugMessage "Audit Kategorie ComboBox nicht gefunden" -Type "Warning"
    Log-Action "UI-Element nicht gefunden: cmbAuditCategory - Event-Handler wurde nicht angehängt"
}

# Initiale Einrichtung der Audit-Optionen
if ($null -ne $cmbAuditCategory -and $null -ne $cmbAuditType) {
    try {
        $cmbAuditCategory.SelectedIndex = 0
        if ($cmbAuditCategory.SelectedItem -ne $null) {
            $initialCategory = $cmbAuditCategory.SelectedItem.Tag.ToString()
            Update-AuditOptions -Category $initialCategory
        }
    } catch {
        Write-DebugMessage "Fehler bei initialer Einrichtung der Audit-Optionen: $($_.Exception.Message)" -Type "Error"
        Log-Action "Fehler bei initialer Einrichtung der Audit-Optionen: $($_.Exception.Message)"
    }
}

# Script-Initialisierung - Keine zusätzliche Initialisierung mehr nötig
# da die Dropdown-Elemente jetzt dynamisch bei Button-Klick befüllt werden

# -------------------------------------------------
# Abschnitt: Exchange Online Troubleshooting Diagnostics
# -------------------------------------------------

# Diagnostics data structure
$script:exchangeDiagnostics = @(
    @{
        Name = "Migration EWS Throttling Policy"
        Description = "Verify that the EWS throttling policy isn't too restrictive for mailbox data migrations (for third‑party tools)."
        PowerShellCheck = "Get-ThrottlingPolicy –Identity 'EWSPolicy'"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/throttling"
        Tooltip = "Open Throttling Policy settings in the Exchange Admin Center"
    },
    @{
        Name = "Exchange Online Accepted Domain diagnostics"
        Description = "Check if a domain is correctly configured as an accepted domain in Exchange Online."
        PowerShellCheck = "Get-AcceptedDomain"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/accepted-domains"
        Tooltip = "Review accepted domains configuration"
    },
    @{
        Name = "Test a user's Exchange Online RBAC permissions"
        Description = "Verify that a user has the necessary RBAC roles to execute specific Exchange Online cmdlets."
        PowerShellCheck = "Get-ManagementRoleAssignment –RoleAssignee '[USER]'"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/permissions"
        Tooltip = "Open RBAC settings to manage user permissions"
        RequiresUser = $true
    },
    @{
        Name = "Compare EXO RBAC Permissions for Two Users"
        Description = "Compare the RBAC roles of two users to identify discrepancies if one user encounters cmdlet errors."
        PowerShellCheck = "Compare-Object (Get-ManagementRoleAssignment –RoleAssignee '[USER1]') (Get-ManagementRoleAssignment –RoleAssignee '[USER2]')"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/permissions"
        Tooltip = "Review and compare RBAC assignments for troubleshooting"
        RequiresTwoUsers = $true
    },
    @{
        Name = "Recipient failure"
        Description = "Check the health and configuration of an Exchange Online recipient to resolve provisioning or sync issues."
        PowerShellCheck = "Get-EXORecipient –Identity '[USER]'"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/recipients"
        Tooltip = "Review recipient configuration and provisioning status"
        RequiresUser = $true
    },
    @{
        Name = "Exchange Organization Object check"
        Description = "Diagnose issues with the Exchange Online organization object, such as tenant provisioning or RBAC misconfigurations."
        PowerShellCheck = "Get-OrganizationConfig | Format-List"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/organization"
        Tooltip = "Review organization configuration settings"
    },
    @{
        Name = "Mailbox or message size"
        Description = "Check mailbox size and message size (including attachments) to identify storage issues."
        PowerShellCheck = "Get-EXOMailboxStatistics –Identity '[USER]' | Select-Object DisplayName, TotalItemSize, ItemCount"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/mailboxes"
        Tooltip = "Review mailbox size and storage statistics"
        RequiresUser = $true
    },
    @{
        Name = "Deleted mailbox diagnostics"
        Description = "Verify the state of recently deleted (soft-deleted) mailboxes for restoration or cleanup."
        PowerShellCheck = "Get-Mailbox –SoftDeletedMailbox"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/deletedmailboxes"
        Tooltip = "Manage deleted mailboxes"
    },
    @{
        Name = "Exchange Remote PowerShell throttling policy"
        Description = "Assess and update the Remote PowerShell throttling policy settings to minimize connection issues."
        PowerShellCheck = "Get-ThrottlingPolicy | Format-Table Name, RCAMaxConcurrency, EwsMaxConnections –AutoSize"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/remote-powershell"
        Tooltip = "Review Remote PowerShell throttling settings"
    },
    @{
        Name = "Email delivery troubleshooter"
        Description = "Diagnose email delivery issues by tracing message paths and identifying failures."
        PowerShellCheck = "Get-MessageTrace –StartDate (Get-Date).AddDays(-7) –EndDate (Get-Date)"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/mailflow"
        Tooltip = "Review mail flow and delivery issues"
    },
    @{
        Name = "Archive mailbox diagnostics"
        Description = "Check the configuration and status of archive mailboxes to ensure archiving is enabled and functioning."
        PowerShellCheck = "Get-EXOMailbox –Identity '[USER]' | Select-Object DisplayName, ArchiveStatus"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/archivemailboxes"
        Tooltip = "Review archive mailbox configuration"
        RequiresUser = $true
    },
    @{
        Name = "Retention policy diagnostics for a user mailbox"
        Description = "Check retention policy settings (including tags and policies) on a user mailbox to ensure compliance with organizational policies."
        PowerShellCheck = "Get-EXOMailbox –Identity '[USER]' | Select-Object DisplayName, RetentionPolicy; Get-RetentionPolicy"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/compliance"
        Tooltip = "Review retention policy settings"
        RequiresUser = $true
    },
    @{
        Name = "DomainKeys Identified Mail (DKIM) diagnostics"
        Description = "Validate that DKIM signing is correctly configured and that the proper DNS entries have been published."
        PowerShellCheck = "Get-DkimSigningConfig"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/dkim"
        Tooltip = "Manage DKIM settings and review DNS configuration"
    },
    @{
        Name = "Proxy address conflict diagnostics"
        Description = "Identify the Exchange recipient using a specific proxy (email address) that causes conflicts, such as errors during mailbox creation."
        PowerShellCheck = "Get-Recipient –Filter {EmailAddresses -like '[EMAIL]'}"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/recipients"
        Tooltip = "Identify and resolve proxy address conflicts"
        RequiresEmail = $true
    },
    @{
        Name = "Mailbox safe/blocked sender list diagnostics"
        Description = "Verify safe senders and blocked senders/domains in a mailbox's junk email settings to troubleshoot potential delivery issues."
        PowerShellCheck = "Get-MailboxJunkEmailConfiguration –Identity '[USER]' | Select-Object SafeSenders, BlockedSenders"
        AdminCenterLink = "https://admin.exchange.microsoft.com/#/mailboxes"
        Tooltip = "Review safe and blocked sender lists"
        RequiresUser = $true
    }
)

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
        }
        
        if ($diagnostic.RequiresTwoUsers -and -not [string]::IsNullOrEmpty($User) -and -not [string]::IsNullOrEmpty($User2)) {
            $command = $command -replace '\[USER1\]', $User
            $command = $command -replace '\[USER2\]', $User2
        }
        
        if ($diagnostic.RequiresEmail -and -not [string]::IsNullOrEmpty($Email)) {
            $command = $command -replace '\[EMAIL\]', $Email
        }
        
        # Befehl ausführen
        Write-DebugMessage "Führe PowerShell-Befehl aus: $command" -Type "Info"
        
        # Create ScriptBlock from command string and execute
        $scriptBlock = [Scriptblock]::Create($command)
        $result = & $scriptBlock | Out-String
        
        Log-Action "Exchange-Diagnose ausgeführt: $($diagnostic.Name)"
        Write-DebugMessage "Diagnose abgeschlossen: $($diagnostic.Name)" -Type "Success"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Diagnose abgeschlossen: $($diagnostic.Name)" -Color $script:connectedBrush
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
        
        return "Fehler bei der Ausführung der Diagnose: $errorMsg"
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
        $url = $diagnostic.AdminCenterLink
        
        Write-DebugMessage "Öffne Admin Center Link: $url" -Type "Info"
        
        # Öffne URL im Standardbrowser
        Start-Process $url
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Admin Center geöffnet: $($diagnostic.Name)"
        }
        
        Log-Action "Admin Center Link geöffnet: $($diagnostic.Name) - $url"
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Öffnen des Admin Center Links: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler beim Öffnen des Admin Center Links: $errorMsg"
        }
        
        Log-Action "Fehler beim Öffnen des Admin Center Links: $errorMsg"
    }
}

# -------------------------------------------------
# Abschnitt: Event Handler für Navigation
# -------------------------------------------------
# ...existing code...

# Event-Handler für Troubleshooting-Navigation
$btnNavTroubleshooting.Add_Click({
    try {
        # Verstecke alle TabItems
        foreach ($tab in $tabContent.Items) {
            $tab.Visibility = [System.Windows.Visibility]::Collapsed
        }
        # Zeige nur Troubleshooting TabItem
        $tabTroubleshooting.Visibility = [System.Windows.Visibility]::Visible
        $tabTroubleshooting.IsSelected = $true
        $txtStatus.Text = "Exchange Online Troubleshooting gewählt"
        Log-Action "Navigation zu Exchange Online Troubleshooting"
        
        # Fülle die Diagnostics-Liste, falls sie leer ist
        if ($null -ne $lstDiagnostics -and $lstDiagnostics.Items.Count -eq 0) {
            foreach ($diagnostic in $script:exchangeDiagnostics) {
                $item = New-Object System.Windows.Controls.ListBoxItem
                $item.Content = $diagnostic.Name
                $item.ToolTip = $diagnostic.Description
                $item.Tag = $script:exchangeDiagnostics.IndexOf($diagnostic)
                $lstDiagnostics.Items.Add($item)
            }
        }
    } catch {
        $errorMsg = $_.Exception.Message
        $txtStatus.Text = "Fehler bei Navigation: $errorMsg"
        Log-Action "Fehler bei Navigation zu Exchange Online Troubleshooting: $errorMsg"
    }
})

# -------------------------------------------------
# Abschnitt: Event Handler für Troubleshooting
# -------------------------------------------------
$lstDiagnostics.Add_SelectionChanged({
    try {
        if ($lstDiagnostics.SelectedItem -ne $null) {
            $selectedIndex = $lstDiagnostics.SelectedItem.Tag
            $diagnostic = $script:exchangeDiagnostics[$selectedIndex]
            
            # Beschreibung anzeigen
            $txtDiagnosticDesc.Text = $diagnostic.Description
            
            # Benutzereingabefelder je nach Diagnostic-Typ ein-/ausblenden
            if ($diagnostic.RequiresUser -eq $true) {
                $txtDiagnosticUser.IsEnabled = $true
                $txtDiagnosticUser.Visibility = [System.Windows.Visibility]::Visible
            } else {
                $txtDiagnosticUser.IsEnabled = $false
                $txtDiagnosticUser.Visibility = [System.Windows.Visibility]::Collapsed
            }
            
            if ($diagnostic.RequiresTwoUsers -eq $true) {
                $txtDiagnosticUser.IsEnabled = $true
                $txtDiagnosticUser.Visibility = [System.Windows.Visibility]::Visible
                $txtDiagnosticUser2.IsEnabled = $true
                $txtDiagnosticUser2.Visibility = [System.Windows.Visibility]::Visible
            } else {
                $txtDiagnosticUser2.IsEnabled = $false
                $txtDiagnosticUser2.Visibility = [System.Windows.Visibility]::Collapsed
            }
            
            if ($diagnostic.RequiresEmail -eq $true) {
                $txtDiagnosticEmail.IsEnabled = $true
                $txtDiagnosticEmail.Visibility = [System.Windows.Visibility]::Visible
            } else {
                $txtDiagnosticEmail.IsEnabled = $false
                $txtDiagnosticEmail.Visibility = [System.Windows.Visibility]::Collapsed
            }
            
            # AdminCenter-Link als Tooltip für den Button setzen
            $btnOpenAdminCenter.ToolTip = $diagnostic.Tooltip
        }
    } catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler bei Auswahl eines Diagnose-Elements: $errorMsg" -Type "Error"
        Log-Action "Fehler bei Auswahl eines Diagnose-Elements: $errorMsg"
    }
})

$btnRunDiagnostic.Add_Click({
    try {
        if (-not $script:isConnected) {
            Write-DebugMessage "Exchange-Diagnose ausführen: Benutzer ist nicht verbunden." -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte verbinden Sie sich zuerst mit Exchange Online.", "Nicht verbunden", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ($null -eq $lstDiagnostics.SelectedItem) {
            Write-DebugMessage "Exchange-Diagnose ausführen: Keine Diagnose ausgewählt" -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Diagnose aus der Liste.", "Keine Auswahl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $selectedIndex = $lstDiagnostics.SelectedItem.Tag
        $diagnostic = $script:exchangeDiagnostics[$selectedIndex]
        
        # Prüfen, ob benötigte Benutzerparameter vorhanden sind
        if ($diagnostic.RequiresUser -and [string]::IsNullOrEmpty($txtDiagnosticUser.Text)) {
            Write-DebugMessage "Exchange-Diagnose ausführen: Benutzer fehlt" -Type "Warning"
            [System.Windows.MessageBox]::Show("Diese Diagnose erfordert einen Benutzernamen.", "Fehlender Parameter", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ($diagnostic.RequiresTwoUsers -and ([string]::IsNullOrEmpty($txtDiagnosticUser.Text) -or [string]::IsNullOrEmpty($txtDiagnosticUser2.Text))) {
            Write-DebugMessage "Exchange-Diagnose ausführen: Zwei Benutzer benötigt" -Type "Warning"
            [System.Windows.MessageBox]::Show("Diese Diagnose erfordert zwei Benutzernamen.", "Fehlende Parameter", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ($diagnostic.RequiresEmail -and [string]::IsNullOrEmpty($txtDiagnosticEmail.Text)) {
            Write-DebugMessage "Exchange-Diagnose ausführen: E-Mail-Adresse fehlt" -Type "Warning"
            [System.Windows.MessageBox]::Show("Diese Diagnose erfordert eine E-Mail-Adresse.", "Fehlender Parameter", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Status aktualisieren
        $txtDiagnosticResult.Text = "Diagnose wird ausgeführt..."
        
        # Diagnose ausführen
        $result = Run-ExchangeDiagnostic -DiagnosticIndex $selectedIndex -User $txtDiagnosticUser.Text -User2 $txtDiagnosticUser2.Text -Email $txtDiagnosticEmail.Text
        
        # Ergebnis anzeigen
        $txtDiagnosticResult.Text = $result
    } catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler bei Ausführung der Diagnose: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler bei Diagnose: $errorMsg"
        }
        
        $txtDiagnosticResult.Text = "Fehler bei der Ausführung der Diagnose:`n$errorMsg"
        Log-Action "Fehler bei Ausführung der Diagnose: $errorMsg"
    }
})

$btnOpenAdminCenter.Add_Click({
    try {
        if ($null -eq $lstDiagnostics.SelectedItem) {
            Write-DebugMessage "Admin Center öffnen: Keine Diagnose ausgewählt" -Type "Warning"
            [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Diagnose aus der Liste.", "Keine Auswahl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $selectedIndex = $lstDiagnostics.SelectedItem.Tag
        Open-AdminCenterLink -DiagnosticIndex $selectedIndex
    } catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Fehler beim Öffnen des Admin Centers: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler beim Öffnen des Admin Centers: $errorMsg"
        }
        
        Log-Action "Fehler beim Öffnen des Admin Centers: $errorMsg"
    }
})

# -------------------------------------------------
# Fenster anzeigen
# -------------------------------------------------
$null = $Form.ShowDialog()

# Helper-Erweiterung für Größenkonvertierungen
# Diese Erweiterungsmethode wird für die Größenanzeige der Postfächer verwendet
Add-Type -TypeDefinition @"
    using System;
    public static class ByteExtensions
    {
        public static double ToMB(this long bytes)
        {
            return Math.Round((double)bytes / 1024 / 1024, 2);
        }
    }
"@
