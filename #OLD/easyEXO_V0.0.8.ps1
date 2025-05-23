# -------------------------------------------------
# Grundlegende Logging-Funktionen
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
            Write-LogEntry "Using PowerShell version: $($psVersion.ToString())"
            return $true
        }
        
        
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
                        Start-Process -FilePath "winget" -ArgumentList "install Microsoft.PowerShell" -Wait -NoNewWindow
                    }
                    else {
                        # Fall back to the MSI installer
                        $installerUrl = "https://github.com/PowerShell/PowerShell/releases/download/v7.3.6/PowerShell-7.3.6-win-x64.msi"
                        $installerPath = "$env:TEMP\PowerShell-7.3.6-win-x64.msi"
                        
                        # Download the installer
                        Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
                        
                        # Install PowerShell 7
                        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$installerPath`" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1" -Wait
                        
                        # Clean up
                        Remove-Item -Path $installerPath -Force
                    }
                    
                    
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
        return $false
    }
    catch {
        $errorMsg = $_.Exception.Message
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
        Write-Log  -Message "$Title - $Type - $Message" -Type "Info"
        
        # Ergebnis zurückgeben (wichtig für Ja/Nein-Fragen)
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        
        # Fallback-Ausgabe
        
        if ($Type -eq "Question") {
            return [System.Windows.MessageBoxResult]::No
        }
    }
}
# Registry-Pfad für Konfigurationseinstellungen
$script:registryPath = "HKCU:\Software\easyIT\easyEXO"

# Stelle sicher, dass der Registry-Pfad existiert
if (-not (Test-Path -Path $script:registryPath)) {
    try {
        # Erstelle den Pfad hierarchisch
        if (-not (Test-Path -Path "HKCU:\Software\easyIT")) {
            New-Item -Path "HKCU:\Software" -Name "easyIT" -Force | Out-Null
        }
        New-Item -Path "HKCU:\Software\easyIT" -Name "easyEXO" -Force | Out-Null
        
        # Standardwerte setzen
        New-ItemProperty -Path $script:registryPath -Name "Debug" -Value 0 -PropertyType DWORD -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "AppName" -Value "Exchange Online Verwaltung" -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "Version" -Value "0.0.8" -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "ThemeColor" -Value "#0078D7" -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "LogPath" -Value "$PSScriptRoot\Logs" -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "HeaderLogoURL" -Value "https://psscripts.de" -PropertyType String -Force | Out-Null
    }
    catch {
        # Fehler beim Erstellen der Registry-Einträge
    }
}

# Lade Konfiguration aus Registry
function Get-RegistryConfig {
    [CmdletBinding()]
    param()
    
    try {
        $config = @{
            "General" = @{}
            "Paths" = @{}
            "UI" = @{}
        }
        
        # Lese alle Registry-Werte
        $regValues = Get-ItemProperty -Path $script:registryPath -ErrorAction SilentlyContinue
        
        if ($regValues) {
            # Debug-Einstellung
            if ($null -ne $regValues.Debug) {
                $config["General"]["Debug"] = $regValues.Debug.ToString()
            }
            
            # AppName
            if ($null -ne $regValues.AppName) {
                $config["General"]["AppName"] = $regValues.AppName
            }
            
            # Version
            if ($null -ne $regValues.Version) {
                $config["General"]["Version"] = $regValues.Version
            }
            
            # ThemeColor
            if ($null -ne $regValues.ThemeColor) {
                $config["General"]["ThemeColor"] = $regValues.ThemeColor
            }
            
            # LogPath
            if ($null -ne $regValues.LogPath) {
                $config["Paths"]["LogPath"] = $regValues.LogPath
            }
            
            # HeaderLogoURL
            if ($null -ne $regValues.HeaderLogoURL) {
                $config["UI"]["HeaderLogoURL"] = $regValues.HeaderLogoURL
            }
        }
        
        return $config
    }
    catch {
        # Fallback zu Standardwerten bei Fehlern
        return @{
            "General" = @{
                "Debug" = "1"
                "AppName" = "Exchange Online Verwaltung"
                "Version" = "0.0.7"
                "ThemeColor" = "#0078D7"
            }
            "Paths" = @{
                "LogPath" = "$PSScriptRoot\Logs"
            }
            "UI" = @{
                "HeaderLogoURL" = "https://psscripts.de"
            }
        }
    }
}

# Lade Konfiguration
$script:config = Get-RegistryConfig

# Debug-Modus einschalten, wenn in Registry aktiviert
if ($script:config["General"]["Debug"] -eq "1") {
    $script:debugMode = $true
}

# --------------------------------------------------------------
# Verbesserte Debug- und Logging-Funktionen
# --------------------------------------------------------------
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateSet("Info", "Warning", "Error", "Success", "Debug")]
        [string]$Type = "Info",

        [Parameter(Mandatory = $false)]
        [switch]$NoLog,

        [Parameter(Mandatory = $false)]
        [switch]$NoConsole
    )

    try {
        # Farbzuordnung für verschiedene Nachrichtentypen
        $colorMap = @{
            "Info"     = "White"
            "Warning"  = "Yellow"
            "Error"    = "Red"
            "Success"  = "Green"
            "Debug"    = "Cyan"
        }

        # Zeitstempel erzeugen
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Nachricht formatieren
        $formattedMessage = "[$timestamp] [$Type] $Message"

        # Ausgabe in Konsole, wenn nicht unterdrückt
        if (-not $NoConsole) {
            # Prüfen, ob die aktuelle Host-Umgebung Farben unterstützt oder der Parameter vorhanden ist.
            try {
                # Versuch mit Farbe
                Write-Host $formattedMessage -ForegroundColor $colorMap[$Type]
            }
            catch [System.Management.Automation.ParameterBindingException] {
                # Speziell den Parameter-Fehler abfangen
                if ($_.Exception.Message -like "*ForegroundColor*") {
                    # Fallback ohne Farbe, wenn -ForegroundColor nicht unterstützt wird
                    Write-Host $formattedMessage
                } else {
                    # Anderen Parameterfehler weiterwerfen (könnte im äußeren Catch landen)
                    throw
                }
            }
            catch {
                # Anderen Fehler beim Schreiben behandeln -> Fallback ohne Farbe
                 Write-Host $formattedMessage # Sicherer Fallback
            }
        }

        # Logging, wenn nicht unterdrückt und Log-Pfad gesetzt ist
        if (-not $NoLog -and $script:logFilePath) {
             try {
                 # Sicherstellen, dass Write-LogEntry existiert, bevor es aufgerufen wird
                 # Annahme: Write-LogEntry ist eine andere Funktion, die das eigentliche Logging übernimmt
                 if (Get-Command Write-LogEntry -ErrorAction SilentlyContinue) {
                    # Write-LogEntry ist verantwortlich für das Logging und interne Fehlerbehandlung
                    Write-LogEntry "$Type - $Message" # Übergibt die Nachricht an die Log-Funktion
                 } else {
                    # Nur warnen (ohne Farbe!), wenn Write-LogEntry fehlt, aber nicht abbrechen
                    Write-Host "Warnung: Funktion Write-LogEntry nicht gefunden. Logging übersprungen." -ForegroundColor Yellow
                 }
             } catch {
                 # Fehler beim Loggen anzeigen (ohne Farbe!), aber Skript nicht unbedingt abbrechen
                 $logErrorMsg = "Fehler beim Aufruf von Write-LogEntry: $($_.Exception.Message)" -replace '[^\x20-\x7E\r\n]', '?'
                 Write-Host $logErrorMsg -ForegroundColor Red
             }
        }
    }
    catch {
        # Fallback bei Fehlern innerhalb der Write-Log Funktion selbst (z.B. durch 'throw' oben oder andere unerwartete Fehler)
        try {
            $errorDetail = "Unbekannter Fehler in Write-Log"
            if ($_) { # Prüfen ob $_ (Fehlerobjekt) existiert
                 if ($_.Exception) { $errorDetail = $_.Exception.ToString() } # Komplette Exception für mehr Details
                 elseif ($_.Message) { $errorDetail = $_.Message }
                 else { $errorDetail = $_.ToString() }
            }
            # Sanitize für den Fall, dass die Fehlermeldung selbst problematische Zeichen enthält
            $sanitizedErrorDetail = $errorDetail -replace '[^\x20-\x7E\r\n]', '?'

            $errorMessage = "Kritischer Fehler in Write-Log Funktion: $sanitizedErrorDetail"

            # Direkter Fallback zur Ausgabe ohne Farbe auf der Konsole
            Write-Host $errorMessage -ForegroundColor Red

            # Versuch, den Fehler zu loggen, falls möglich und Log-Pfad vorhanden
            if ($script:logFilePath) {
                try {
                    $logErrorMessage = "[$([System.DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] [CRITICAL_ERROR_IN_WRITEHOST] $errorMessage"
                    # Sicherstellen, dass das Log-Verzeichnis existiert
                    $logFolder = Split-Path -Path $script:logFilePath -Parent
                    if (-not (Test-Path $logFolder)) {
                        New-Item -ItemType Directory -Path $logFolder -Force -ErrorAction SilentlyContinue | Out-Null
                    }
                    # Direkter Log-Versuch
                    Add-Content -Path $script:logFilePath -Value $logErrorMessage -Encoding UTF8 -ErrorAction SilentlyContinue
                } catch {
                    # Logging des Fehlers ist fehlgeschlagen, nichts weiter tun
                }
            }
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
            Write-Log  "Logverzeichnis wurde erstellt: $logFolder" -Type "Info"
        }
        
        # Log-Eintrag schreiben
        Add-Content -Path $script:logFilePath -Value "[$timestamp] $sanitizedMessage" -Encoding UTF8
        
        # Bei zu langer Logdatei (>10 MB) rotieren
        $logFile = Get-Item -Path $script:logFilePath -ErrorAction SilentlyContinue
        if ($logFile -and $logFile.Length -gt 10MB) {
            $backupLogPath = "$($script:logFilePath)_$(Get-Date -Format 'yyyyMMdd_HHmmss').bak"
            Move-Item -Path $script:logFilePath -Destination $backupLogPath -Force
            Write-Log  "Logdatei wurde rotiert: $backupLogPath" -Type "Info"
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
            Write-Log  "GUI-Element ist null in Update-GuiText" -Type "Warning"
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
            Write-Log  "Fehler in Update-GuiText: $errorMsg" -Type "Error"
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
        Write-Log  -Message $Message -Type $Type
        
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
        Write-Log  "Fehler in Write-StatusMessage: $errorMsg" -Type "Error"
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
function Connect-ExchangeOnline {
    [CmdletBinding()]
    param()
    
    try {
        Write-LogEntry "Verbindungsversuch zu Exchange Online..." -Type "Info"
        
        # Prüfen, ob das ExchangeOnlineManagement Modul installiert ist
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            $errorMsg = "ExchangeOnlineManagement Modul ist nicht installiert. Bitte installieren Sie das Modul mit 'Install-Module ExchangeOnlineManagement -Force'"
            Write-LogEntry $errorMsg -Type "Error"
            Show-MessageBox -Message $errorMsg -Title "Modul fehlt" -Type "Error"
            return $false
        }
        
        # Modul laden
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        
        # WPF-Fenster für die Benutzereingabe erstellen
        $inputXaml = @"
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                Title="Exchange Online Anmeldung" Height="150" Width="400" WindowStartupLocation="CenterScreen">
            <Grid Margin="10">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Label Grid.Row="0" Content="EXO Administrativer Logins:"/>
                <TextBox Grid.Row="1" Name="txtEmail" Margin="0,5"/>
                <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                    <Button Name="btnOK" Content="OK" Width="75" Margin="0,0,5,0"/>
                    <Button Name="btnCancel" Content="Abbrechen" Width="75"/>
                </StackPanel>
            </Grid>
        </Window>
"@
        
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.LoadXml($inputXaml)
        $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        $txtEmail = $window.FindName("txtEmail")
        $btnOK = $window.FindName("btnOK")
        $btnCancel = $window.FindName("btnCancel")
        
        # Variable für die E-Mail-Adresse im Skript-Bereich definieren
        $script:userPrincipalName = $null
        
        $btnOK.Add_Click({
            if (-not [string]::IsNullOrWhiteSpace($txtEmail.Text)) {
                $script:userPrincipalName = $txtEmail.Text
                $window.DialogResult = $true
                $window.Close()
            }
        })
        
        $btnCancel.Add_Click({
            $window.DialogResult = $false
            $window.Close()
        })
        
        $result = $window.ShowDialog()
        
        if (-not $result) {
            $errorMsg = "Anmeldung abgebrochen."
            Write-LogEntry $errorMsg -Type "Warning"
            Show-MessageBox -Message $errorMsg -Title "Abgebrochen" -Type "Warning"
            return $false
        }
        
        # Überprüfen, ob die E-Mail-Adresse erfolgreich gespeichert wurde
        if ([string]::IsNullOrWhiteSpace($script:userPrincipalName)) {
            $errorMsg = "Keine E-Mail-Adresse eingegeben oder erkannt. Verbindung abgebrochen."
            Write-LogEntry $errorMsg -Type "Warning"
            Show-MessageBox -Message $errorMsg -Title "Abgebrochen" -Type "Warning"
            return $false
        }

        # Verbindungsparameter für V3
        $connectParams = @{
            UserPrincipalName = $script:userPrincipalName
            ErrorAction = "Stop"
        }
        
        # Prüfen, ob der ShowBanner-Parameter unterstützt wird
        $cmdInfo = Get-Command Microsoft.PowerShell.Core\Get-Command -Module ExchangeOnlineManagement -Name Connect-ExchangeOnline -ErrorAction SilentlyContinue
        if ($cmdInfo -and $cmdInfo.Parameters.ContainsKey('ShowBanner')) {
            $connectParams.Add('ShowBanner', $false)
        }
        
        # Verbindung herstellen
        Show-MessageBox -Message "Verbindung wird hergestellt für: $script:userPrincipalName"
        & (Get-Module ExchangeOnlineManagement).ExportedCommands['Connect-ExchangeOnline'] @connectParams
        
        # Verbindung testen
        $null = Get-OrganizationConfig -ErrorAction Stop
        
        # Globale und Skript-Variablen setzen, um den Verbindungsstatus zu speichern
        $Global:IsConnectedToExo = $true
        $script:isConnected = $true
        
        Write-LogEntry "Exchange Online Verbindung erfolgreich hergestellt für $script:userPrincipalName" -Title "Verbindung hergestellt" -Type "Success"
        $script:txtConnectionStatus.Text = "Verbunden mit Exchange Online ($script:userPrincipalName)"
        $script:txtConnectionStatus.Foreground = "#008000"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-LogEntry "Fehler beim Verbinden mit Exchange Online: $errorMsg" -Type "Error"
        $script:txtConnectionStatus.Text = "Nicht verbunden"
        $script:txtConnectionStatus.Foreground = "#d83b01"
        $Global:IsConnectedToExo = $false
        $script:isConnected = $false
        Show-MessageBox -Message "Fehler beim Verbinden mit Exchange Online: $errorMsg" -Title "Verbindungsfehler" -Type "Error"
        return $false
    }
}

# Funktion zum Überprüfen der Exchange Online Verbindung
function Test-ExchangeOnlineConnection {
    [CmdletBinding()]
    param()
    
    try {
        # Prüfe, ob eine aktive Exchange Online Session existiert
        $exoSession = Get-PSSession | Where-Object { 
            $_.ConfigurationName -eq "Microsoft.Exchange" -and 
            $_.State -eq "Opened" -and 
            $_.Availability -eq "Available" 
        }
        
        if ($null -eq $exoSession) {
            Write-Log "Keine aktive Exchange Online Verbindung gefunden. Versuche neu zu verbinden..." -Type "Warning"
            Connect-ExchangeOnline -ShowBanner:$false
            Start-Sleep -Seconds 2
            
            # Prüfe erneut nach dem Verbindungsversuch
            $exoSession = Get-PSSession | Where-Object { 
                $_.ConfigurationName -eq "Microsoft.Exchange" -and 
                $_.State -eq "Opened" -and 
                $_.Availability -eq "Available" 
            }
            
            if ($null -eq $exoSession) {
                Write-Log "Verbindung zu Exchange Online konnte nicht hergestellt werden." -Type "Error"
                return $false
            }
        }
        
        # Teste die Verbindung mit einem einfachen Kommando
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Log "Exchange Online Verbindung erfolgreich bestätigt." -Type "Info"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler bei der Exchange Online Verbindung: $errorMsg" -Type "Error"
        return $false
    }
}

function Disconnect-ExchangeOnlineSession {
    [CmdletBinding()]
    param()
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Log-Action "Exchange Online Verbindung getrennt"
        
        # Setze alle Verbindungsvariablen zurück
        $Global:IsConnectedToExo = $false
        $script:isConnected = $false
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Exchange Verbindung getrennt"
        }
        if ($null -ne $script:txtConnectionStatus) {
            $script:txtConnectionStatus.Text = "Nicht verbunden"
            $script:txtConnectionStatus.Foreground = $script:disconnectedBrush
        }
        
        # Button-Status aktualisieren
        if ($null -ne $script:btnConnect) {
            $script:btnConnect.Content = "Mit Exchange verbinden"
            $script:btnConnect.Tag = "connect"
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler beim Trennen der Verbindung: $errorMsg"
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
            Write-Log "Fehler beim Trennen der Verbindung: $errorMsg"  
        }
        
        return $false
    }
}

# Funktion zur Überprüfung der Exchange Online-Verbindung
function Confirm-ExchangeConnection {
    [CmdletBinding()]
    param()
    
    try {
        # Überprüfen, ob eine der Verbindungsvariablen gesetzt ist
        if ($Global:IsConnectedToExo -eq $true -or $script:isConnected -eq $true) {
            # Verbindung testen durch Abrufen einer Exchange-Information
            try {
                $null = Get-OrganizationConfig -ErrorAction Stop
                # Stelle sicher, dass beide Variablen konsistent sind
                $Global:IsConnectedToExo = $true
                $script:isConnected = $true
                return $true
            }
            catch {
                # Verbindung ist nicht mehr gültig, setze beide Variablen zurück
                $Global:IsConnectedToExo = $false
                $script:isConnected = $false
                Write-Log "Exchange Online Verbindung getrennt: $($_.Exception.Message)" -Type "Warning"
                return $false
            }
        }
        else {
            return $false
        }
    }
    catch {
        $Global:IsConnectedToExo = $false
        $script:isConnected = $false
        Write-Log "Fehler bei der Überprüfung der Exchange Online-Verbindung: $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

function Ensure-ExchangeConnection {
    # Prüfen, ob eine gültige Verbindung besteht
    if (-not (Confirm-ExchangeConnection)) {
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Verbindung zu Exchange Online wird hergestellt..."
        }
        try {
            # Verbindung herstellen
            $result = Connect-ExchangeOnline
            if ($result) {
                if ($null -ne $script:txtStatus) {
                    $script:txtStatus.Text = "Verbindung zu Exchange Online hergestellt"
                }
                return $true
            } else {
                if ($null -ne $script:txtStatus) {
                    $script:txtStatus.Text = "Fehler beim Verbinden mit Exchange Online"
                }
                return $false
            }
        }
        catch {
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Text = "Fehler beim Verbinden mit Exchange Online: $($_.Exception.Message)"
            }
            return $false
        }
    }
    return $true
}
# Funktion zum Überprüfen der Voraussetzungen (Module)
function Check-Prerequisites {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log  "Überprüfe benötigte PowerShell-Module" -Type "Info"
        
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
        Write-Log  "Fehler bei der Überprüfung der Module: $errorMsg" -Type "Error"
        
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
        Write-Log  "Installiere benötigte PowerShell-Module" -Type "Info"
        
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
            Write-Log  "PowerShellGet-Modul ist veraltet oder nicht installiert, versuche zu aktualisieren" -Type "Warning"
            
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
                    Write-Log  "Benutzer hat Administrator-Neustart abgelehnt" -Type "Warning"
                    
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
                    Write-Log  "PowerShellGet erfolgreich aktualisiert" -Type "Success"
                } catch {
                    Write-Log  "Fehler beim Aktualisieren von PowerShellGet: $($_.Exception.Message)" -Type "Error"
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
            
            Write-Log  "Installiere/Aktualisiere Modul: $moduleName" -Type "Info"
            
            try {
                # Prüfe, ob Modul bereits installiert ist
                $module = Get-Module -Name $moduleName -ListAvailable -ErrorAction SilentlyContinue
                
                if ($null -ne $module) {
                    $latestVersion = ($module | Sort-Object Version -Descending | Select-Object -First 1).Version
                    
                    # Prüfe, ob Update notwendig ist
                    if ($null -ne $minVersion -and $latestVersion -lt [Version]$minVersion) {
                        Write-Log  "Aktualisiere Modul $moduleName von $latestVersion auf mindestens $minVersion" -Type "Info"
                        Install-Module -Name $moduleName -Force -AllowClobber -MinimumVersion $minVersion
                        $newVersion = (Get-Module -Name $moduleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version
                        
                        $results += [PSCustomObject]@{
                            Module = $moduleName
                            Status = "Aktualisiert"
                            AlteVersion = $latestVersion
                            NeueVersion = $newVersion
                        }
                    } else {
                        Write-Log  "Modul $moduleName ist bereits in ausreichender Version ($latestVersion) installiert" -Type "Info"
                        
                        $results += [PSCustomObject]@{
                            Module = $moduleName
                            Status = "Bereits aktuell"
                            AlteVersion = $latestVersion
                            NeueVersion = $latestVersion
                        }
                    }
                } else {
                    # Installiere Modul
                    Write-Log  "Installiere Modul $moduleName" -Type "Info"
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
                Write-Log  "Fehler beim Installieren/Aktualisieren von $moduleName - $errorMsg" -Type "Error"
                
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
        Write-Log  "Fehler bei der Modulinstallation: $errorMsg" -Type "Error"
        
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
        
        Write-Log  "Rufe Kalenderberechtigungen ab für: $MailboxUser" -Type "Info"
        
        # Prüfe deutsche und englische Kalenderordner
        $permissions = $null
        try {
            # Versuche mit deutschem Pfad
            $identity = "${MailboxUser}:\Kalender"
            Write-Log  "Versuche deutschen Kalenderpfad: $identity" -Type "Info"
            $permissions = Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop
        } 
        catch {
            try {
                # Versuche mit englischem Pfad
                $identity = "${MailboxUser}:\Calendar"
                Write-Log  "Versuche englischen Kalenderpfad: $identity" -Type "Info"
                $permissions = Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop
            } 
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log  "Beide Kalenderpfade fehlgeschlagen: $errorMsg" -Type "Error"
                throw "Kalenderordner konnte nicht gefunden werden. Weder 'Kalender' noch 'Calendar' sind zugänglich."
            }
        }
        
        Write-Log  "Kalenderberechtigungen abgerufen: $($permissions.Count) Einträge gefunden" -Type "Success"
        Log-Action "Kalenderberechtigungen für $MailboxUser erfolgreich abgerufen: $($permissions.Count) Einträge."
        return $permissions
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der Kalenderberechtigungen: $errorMsg" -Type "Error"
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
        Write-Log  "Fehler beim Anzeigen der Kalenderberechtigungen: $errorMsg" -Type "Error"
        
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
        Write-Log  "Setze Standardberechtigungen für Kalender: $PermissionType mit $AccessRights" -Type "Info"
        
        if ($ForAllMailboxes) {
            # Frage den Benutzer ob er das wirklich tun möchte
            $confirmResult = [System.Windows.MessageBox]::Show(
                "Möchten Sie wirklich die $PermissionType-Berechtigungen für ALLE Postfächer setzen? Diese Aktion kann bei vielen Postfächern lange dauern.",
                "Massenänderung bestätigen",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning)
                
            if ($confirmResult -eq [System.Windows.MessageBoxResult]::No) {
                Write-Log  "Massenänderung vom Benutzer abgebrochen" -Type "Info"
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
        
        Write-Log  "Standardberechtigungen für Kalender erfolgreich gesetzt: $PermissionType mit $AccessRights" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Standardberechtigungen gesetzt: $PermissionType mit $AccessRights" -Color $script:connectedBrush
        }
        Log-Action "Standardberechtigungen für Kalender gesetzt: $PermissionType mit $AccessRights"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Setzen der Standardberechtigungen für Kalender: $errorMsg" -Type "Error"
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
        
        Write-Log  "Füge Kalenderberechtigung hinzu/aktualisiere: $SourceUser -> $TargetUser ($Permission)" -Type "Info"
        
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
                Write-Log  "Bestehende Berechtigung gefunden (DE): $($existingPermDE.AccessRights)" -Type "Info"
            }
            else {
                # Dann den englischen Kalender probieren
                $existingPermEN = Get-MailboxFolderPermission -Identity $identityEN -User $TargetUser -ErrorAction SilentlyContinue
                if ($null -ne $existingPermEN) {
                    $calendarExists = $true
                    $identity = $identityEN
                    Write-Log  "Bestehende Berechtigung gefunden (EN): $($existingPermEN.AccessRights)" -Type "Info"
                }
            }
    }
    catch {
            Write-Log  "Fehler bei der Prüfung bestehender Berechtigungen: $($_.Exception.Message)" -Type "Warning"
        }
        
        # Falls noch kein identifizierter Kalender, versuchen wir die Kalender zu prüfen ohne Benutzerberechtigungen
        if ($null -eq $identity) {
            try {
                # Prüfen, ob der deutsche Kalender existiert
                $deExists = Get-MailboxFolderPermission -Identity $identityDE -ErrorAction SilentlyContinue
                if ($null -ne $deExists) {
                    $identity = $identityDE
                    Write-Log  "Deutscher Kalenderordner gefunden: $identityDE" -Type "Info"
                }
                else {
                    # Prüfen, ob der englische Kalender existiert
                    $enExists = Get-MailboxFolderPermission -Identity $identityEN -ErrorAction SilentlyContinue
                    if ($null -ne $enExists) {
                        $identity = $identityEN
                        Write-Log  "Englischer Kalenderordner gefunden: $identityEN" -Type "Info"
                    }
                }
            }
            catch {
                Write-Log  "Fehler beim Prüfen der Kalenderordner: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Falls immer noch kein Kalender gefunden, über Statistiken suchen
        if ($null -eq $identity) {
            try {
                $folderStats = Get-MailboxFolderStatistics -Identity $SourceUser -FolderScope Calendar -ErrorAction Stop
                foreach ($folder in $folderStats) {
                    if ($folder.FolderType -eq "Calendar" -or $folder.Name -eq "Kalender" -or $folder.Name -eq "Calendar") {
                        $identity = "$SourceUser`:" + $folder.FolderPath.Replace("/", "\")
                        Write-Log  "Kalenderordner über FolderStatistics gefunden: $identity" -Type "Info"
                        break
                    }
                }
            }
            catch {
                Write-Log  "Fehler beim Suchen des Kalenderordners über FolderStatistics: $($_.Exception.Message)" -Type "Warning"
            }
        }
        
        # Wenn immer noch kein Kalender gefunden, Exception werfen
        if ($null -eq $identity) {
            throw "Kein Kalenderordner für $SourceUser gefunden. Bitte stellen Sie sicher, dass das Postfach existiert und Sie Zugriff haben."
        }
        
        # Je nachdem ob Berechtigung existiert, update oder add
        if ($calendarExists) {
            Write-Log  "Aktualisiere bestehende Berechtigung: $identity ($Permission)" -Type "Info"
            Set-MailboxFolderPermission -Identity $identity -User $TargetUser -AccessRights $Permission -ErrorAction Stop
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Kalenderberechtigung aktualisiert." -Color $script:connectedBrush
            }
            
            Write-Log  "Kalenderberechtigung erfolgreich aktualisiert" -Type "Success"
            Log-Action "Kalenderberechtigung aktualisiert: $SourceUser -> $TargetUser mit $Permission"
        }
        else {
            Write-Log  "Füge neue Berechtigung hinzu: $identity ($Permission)" -Type "Info"
            Add-MailboxFolderPermission -Identity $identity -User $TargetUser -AccessRights $Permission -ErrorAction Stop
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Kalenderberechtigung hinzugefügt." -Color $script:connectedBrush
            }
            
            Write-Log  "Kalenderberechtigung erfolgreich hinzugefügt" -Type "Success"
            Log-Action "Kalenderberechtigung hinzugefügt: $SourceUser -> $TargetUser mit $Permission"
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Hinzufügen/Aktualisieren der Kalenderberechtigung: $errorMsg" -Type "Error"
        
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
        
        Write-Log  "Entferne Kalenderberechtigung: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfe deutsche und englische Kalenderordner
        $removed = $false
        
        try {
            $identityDE = "${SourceUser}:\Kalender"
            Write-Log  "Prüfe deutsche Kalenderberechtigungen: $identityDE" -Type "Info"
            
            # Prüfe ob Berechtigung existiert
            $existingPerm = Get-MailboxFolderPermission -Identity $identityDE -User $TargetUser -ErrorAction SilentlyContinue
            
            if ($existingPerm) {
                Write-Log  "Gefundene Berechtigung wird entfernt (DE): $($existingPerm.AccessRights)" -Type "Info"
                Remove-MailboxFolderPermission -Identity $identityDE -User $TargetUser -Confirm:$false -ErrorAction Stop
                $removed = $true
                Write-Log  "Berechtigung erfolgreich entfernt (DE)" -Type "Success"
            }
            else {
                Write-Log  "Keine Berechtigung gefunden für deutschen Kalender" -Type "Info"
            }
        } 
        catch {
            $errorMsg = $_.Exception.Message
            Write-Log  "Fehler beim Entfernen der deutschen Kalenderberechtigungen: $errorMsg" -Type "Warning"
            # Bei Fehler einfach weitermachen und englischen Pfad versuchen
        }
        
        if (-not $removed) {
            try {
                $identityEN = "${SourceUser}:\Calendar"
                Write-Log  "Prüfe englische Kalenderberechtigungen: $identityEN" -Type "Info"
                
                # Prüfe ob Berechtigung existiert
                $existingPerm = Get-MailboxFolderPermission -Identity $identityEN -User $TargetUser -ErrorAction SilentlyContinue
                
                if ($existingPerm) {
                    Write-Log  "Gefundene Berechtigung wird entfernt (EN): $($existingPerm.AccessRights)" -Type "Info"
                    Remove-MailboxFolderPermission -Identity $identityEN -User $TargetUser -Confirm:$false -ErrorAction Stop
                    $removed = $true
                    Write-Log  "Berechtigung erfolgreich entfernt (EN)" -Type "Success"
                }
                else {
                    Write-Log  "Keine Berechtigung gefunden für englischen Kalender" -Type "Info"
                }
            } 
            catch {
                if (-not $removed) {
                    $errorMsg = $_.Exception.Message
                    Write-Log  "Fehler beim Entfernen der englischen Kalenderberechtigungen: $errorMsg" -Type "Error"
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
            Write-Log  "Keine Kalenderberechtigung zum Entfernen gefunden" -Type "Warning"
            
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine Kalenderberechtigung gefunden zum Entfernen."
            }
            
            Log-Action "Keine Kalenderberechtigung gefunden zum Entfernen: $SourceUser -> $TargetUser"
            return $false
        }
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Entfernen der Kalenderberechtigung: $errorMsg" -Type "Error"
        
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
        
        Write-Log  "Füge Postfachberechtigung hinzu: $SourceUser -> $TargetUser (FullAccess)" -Type "Info"
        
        # Prüfen, ob die Berechtigung bereits existiert
        $existingPermissions = Get-MailboxPermission -Identity $SourceUser -User $TargetUser -ErrorAction SilentlyContinue
        $fullAccessExists = $existingPermissions | Where-Object { $_.AccessRights -like "*FullAccess*" }
        
        if ($fullAccessExists) {
            Write-Log  "Berechtigung existiert bereits, keine Änderung notwendig" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Postfachberechtigung bereits vorhanden." -Color $script:connectedBrush
            }
            Log-Action "Postfachberechtigung bereits vorhanden: $SourceUser -> $TargetUser"
            return $true
        }
        
        # Berechtigung hinzufügen
        Add-MailboxPermission -Identity $SourceUser -User $TargetUser -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop
        
        Write-Log  "Postfachberechtigung erfolgreich hinzugefügt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Postfachberechtigung hinzugefügt." -Color $script:connectedBrush
        }
        Log-Action "Postfachberechtigung hinzugefügt: $SourceUser -> $TargetUser (FullAccess)"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Hinzufügen der Postfachberechtigung: $errorMsg" -Type "Error"
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
        
        Write-Log  "Entferne Postfachberechtigung: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfen, ob die Berechtigung existiert
        $existingPermissions = Get-MailboxPermission -Identity $SourceUser -User $TargetUser -ErrorAction SilentlyContinue
        if (-not $existingPermissions) {
            Write-Log  "Keine Berechtigung zum Entfernen gefunden" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine Postfachberechtigung zum Entfernen gefunden."
            }
            Log-Action "Keine Postfachberechtigung zum Entfernen gefunden: $SourceUser -> $TargetUser"
            return $false
        }
        
        # Berechtigung entfernen
        Remove-MailboxPermission -Identity $SourceUser -User $TargetUser -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
        
        Write-Log  "Postfachberechtigung erfolgreich entfernt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Postfachberechtigung entfernt."
        }
        Log-Action "Postfachberechtigung entfernt: $SourceUser -> $TargetUser"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Entfernen der Postfachberechtigung: $errorMsg" -Type "Error"
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
        Write-Log  "Postfachberechtigungen abrufen: Validiere Benutzereingabe" -Type "Info"
        
        if ([string]::IsNullOrEmpty($MailboxUser)) {
            Write-Log  "Keine gültige E-Mail-Adresse angegeben" -Type "Error"
            return $null
        }
        
        Write-Log  "Postfachberechtigungen abrufen für: $MailboxUser" -Type "Info"
        Write-Log  "Rufe Postfachberechtigungen ab für: $MailboxUser" -Type "Info"
        
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
                Write-Log  "Postfachberechtigung verarbeitet: $($perm.User) -> $($perm.AccessRights)" -Type "Info"
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
                Write-Log  "SendAs-Berechtigung verarbeitet: $($perm.User) -> SendAs" -Type "Info"
            }
        }
        
        $count = $allPermissions.Count
        Write-Log  "Postfachberechtigungen abgerufen und verarbeitet: $count Einträge gefunden" -Type "Success"
        
        return $allPermissions
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der Postfachberechtigungen: $errorMsg" -Type "Error"
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
        Write-Log  "Postfachberechtigungen abrufen: Validiere Benutzereingabe" -Type "Info"
        
        # E-Mail-Format überprüfen
        if (-not ($Mailbox -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")) {
            if (-not ($Mailbox -match "^[a-zA-Z0-9\s.-]+$")) {
                throw "Ungültige E-Mail-Adresse oder Benutzername: $Mailbox"
            }
        }
        
        Write-Log  "Postfachberechtigungen abrufen für: $Mailbox" -Type "Info"
        
        # Postfachberechtigungen abrufen
        Write-Log  "Rufe Postfachberechtigungen ab für: $Mailbox" -Type "Info"
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
            
            Write-Log  "Postfachberechtigung verarbeitet: $($perm.User) -> $($perm.AccessRights -join ', ')" -Type "Info"
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
                Write-Log  "Separate SendAs-Berechtigung verarbeitet: $($sendPerm.Trustee)" -Type "Info"
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
                Write-Log  "Keine benutzerdefinierten Berechtigungen gefunden, nur Standardzugriff" -Type "Info"
            }
            else {
                $entry = [PSCustomObject]@{
                    Identity = $Mailbox
                    User = "Keine Berechtigungen gefunden"
                    AccessRights = "Unbekannt"
                }
                $resultCollection += $entry
                Write-Log  "Keine Berechtigungen gefunden" -Type "Warning"
            }
        }
        
        Write-Log  "Postfachberechtigungen abgerufen und verarbeitet: $($resultCollection.Count) Einträge gefunden" -Type "Success"
        
        # Wichtig: Rückgabe als Array für die GUI-Darstellung
        return ,$resultCollection
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der Postfachberechtigungen: $errorMsg" -Type "Error"
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
        
        Write-Log  "Setze Standard-Kalenderberechtigungen für: $MailboxUser auf: $AccessRights" -Type "Info"
        
        # Prüfe deutsche und englische Kalenderordner
        $identityDE = "${MailboxUser}:\Kalender"
        $identityEN = "${MailboxUser}:\Calendar"
        $identity = $null
        
        # Prüfe, welcher Pfad existiert
        try {
            if (Get-MailboxFolderPermission -Identity $identityDE -User Default -ErrorAction SilentlyContinue) {
                $identity = $identityDE
                Write-Log  "Deutscher Kalenderpfad gefunden: $identity" -Type "Info"
            } else {
                $identity = $identityEN
                Write-Log  "Englischer Kalenderpfad wird verwendet: $identity" -Type "Info"
            }
        } catch {
            $identity = $identityEN
            Write-Log  "Fehler beim Prüfen des deutschen Pfads, verwende englischen Pfad: $identity" -Type "Warning"
        }
        
        # Standard-Berechtigungen setzen
        Write-Log  "Aktualisiere Standard-Berechtigungen für: $identity" -Type "Info"
        Set-MailboxFolderPermission -Identity $identity -User Default -AccessRights $AccessRights -ErrorAction Stop
        
        Write-Log  "Standard-Kalenderberechtigungen erfolgreich gesetzt" -Type "Success"
        Log-Action "Standard-Kalenderberechtigungen für $MailboxUser auf $AccessRights gesetzt"
        return $true
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Setzen der Standard-Kalenderberechtigungen: $errorMsg" -Type "Error"
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
        
        Write-Log  "Setze Anonym-Kalenderberechtigungen für: $MailboxUser auf: $AccessRights" -Type "Info"
        
        # Prüfe deutsche und englische Kalenderordner
        $identityDE = "${MailboxUser}:\Kalender"
        $identityEN = "${MailboxUser}:\Calendar"
        $identity = $null
        
        # Prüfe, welcher Pfad existiert
        try {
            if (Get-MailboxFolderPermission -Identity $identityDE -User Anonymous -ErrorAction SilentlyContinue) {
                $identity = $identityDE
                Write-Log  "Deutscher Kalenderpfad gefunden: $identity" -Type "Info"
            } else {
                $identity = $identityEN
                Write-Log  "Englischer Kalenderpfad wird verwendet: $identity" -Type "Info"
            }
        } catch {
            $identity = $identityEN
            Write-Log  "Fehler beim Prüfen des deutschen Pfads, verwende englischen Pfad: $identity" -Type "Warning"
        }
        
        # Anonym-Berechtigungen setzen
        Write-Log  "Aktualisiere Anonymous-Berechtigungen für: $identity" -Type "Info"
        Set-MailboxFolderPermission -Identity $identity -User Anonymous -AccessRights $AccessRights -ErrorAction Stop
        
        Write-Log  "Anonymous-Kalenderberechtigungen erfolgreich gesetzt" -Type "Success"
        Log-Action "Anonymous-Kalenderberechtigungen für $MailboxUser auf $AccessRights gesetzt"
        return $true
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Setzen der Anonymous-Kalenderberechtigungen: $errorMsg" -Type "Error"
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
        Write-Log  "Setze Standard-Kalenderberechtigungen für alle Postfächer auf: $AccessRights" -Type "Info"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Setze Standard-Kalenderberechtigungen für alle Postfächer..."
        }
        
        # Alle Mailboxen abrufen
        Write-Log  "Rufe alle Mailboxen ab" -Type "Info"
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
        $totalCount = $mailboxes.Count
        $successCount = 0
        $errorCount = 0
        
        Write-Log  "$totalCount Mailboxen gefunden" -Type "Info"
        
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
                Write-Log  "Bearbeite Postfach $progressIndex/$totalCount - $mailboxAddress" -Type "Info"
                
                Set-DefaultCalendarPermission -MailboxUser $mailboxAddress -AccessRights $AccessRights
                $successCount++
                Write-Log  "Standard-Kalenderberechtigungen erfolgreich für $mailboxAddress gesetzt" -Type "Success"
            }
            catch {
                $errorCount++
                $errorMsg = $_.Exception.Message
                Write-Log  "Fehler bei Postfach $mailboxAddress - $errorMsg" -Type "Error"
                Log-Action "Fehler beim Setzen der Standard-Kalenderberechtigungen für $mailboxAddress`: $errorMsg"
            }
        }
        
        $statusMessage = "Standard-Kalenderberechtigungen für alle Postfächer gesetzt. Erfolgreich - $successCount, Fehler: $errorCount"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message $statusMessage -Color $script:connectedBrush
        }
        
        Write-Log  $statusMessage -Type "Success"
        Log-Action $statusMessage
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Setzen der Standard-Kalenderberechtigungen für alle - $errorMsg" -Type "Error"
        
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
        Write-Log  "Setze Anonym-Kalenderberechtigungen für alle Postfächer auf: $AccessRights" -Type "Info"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Setze Anonym-Kalenderberechtigungen für alle Postfächer..."
        }
        
        # Alle Mailboxen abrufen
        Write-Log  "Rufe alle Mailboxen ab" -Type "Info"
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
        $totalCount = $mailboxes.Count
        $successCount = 0
        $errorCount = 0
        
        Write-Log  "$totalCount Mailboxen gefunden" -Type "Info"
        
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
                Write-Log  "Bearbeite Postfach $progressIndex/$totalCount - $mailboxAddress" -Type "Info"
                
                Set-AnonymousCalendarPermission -MailboxUser $mailboxAddress -AccessRights $AccessRights
                $successCount++
                Write-Log  "Anonym-Kalenderberechtigungen erfolgreich für $mailboxAddress gesetzt" -Type "Success"
            }
            catch {
                $errorCount++
                $errorMsg = $_.Exception.Message
                Write-Log  "Fehler bei Postfach $mailboxAddress - $errorMsg" -Type "Error"
                Log-Action "Fehler beim Setzen der Anonym-Kalenderberechtigungen für $mailboxAddress`: $errorMsg"
            }
        }
        
        $statusMessage = "Anonym-Kalenderberechtigungen für alle Postfächer gesetzt. Erfolgreich - $successCount, Fehler: $errorCount"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message $statusMessage -Color $script:connectedBrush
        }
        
        Write-Log  $statusMessage -Type "Success"
        Log-Action $statusMessage
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Setzen der Anonym-Kalenderberechtigungen für alle - $errorMsg" -Type "Error"
        
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
        Write-Log  "Rufe Exchange Throttling Informationen ab: $InfoType" -Type "Info"
        
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

        Write-Log  "Exchange Throttling Information erfolgreich erstellt" -Type "Success"
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der Exchange Throttling Informationen: $errorMsg" -Type "Error"
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
        Write-Log  "Rufe alternative Throttling-Informationen ab" -Type "Info"
        
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
        Write-Log  "Fehler beim Abrufen der Throttling-Informationen: $errorMsg" -Type "Error"
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
        Write-Log  "Führe Throttling Policy Troubleshooting aus: $PolicyType" -Type "Info"
        
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
        Write-Log  "Fehler beim Throttling Policy Troubleshooting: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Throttling Policy Troubleshooting: $errorMsg"
        return "Fehler beim Abrufen der Throttling Policy Informationen: $errorMsg"
    }
}

# Erweitere die Diagnostics-Funktionen um einen speziellen Throttling-Test
function Test-EWSThrottlingPolicy {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log  "Prüfe EWS Throttling Policy für Migration" -Type "Info"
        
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
        
        Write-Log  "EWS Throttling Policy Test abgeschlossen" -Type "Success"
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Testen der EWS Throttling Policy: $errorMsg" -Type "Error"
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
        Write-Log  "Führe Diagnose aus: $($diagnostic.Name)" -Type "Info"
        
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
        Write-Log  "Führe PowerShell-Befehl aus: $command" -Type "Info"
        
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
                Write-Log  "Get-ThrottlingPolicy ist nicht verfügbar, verwende alternative Informationsquellen" -Type "Warning"
                $result = Get-ExchangeThrottlingInfo -InfoType $(if ($command -like "*EWS*") { "EWSPolicy" } elseif ($command -like "*PowerShell*") { "PowerShell" } else { "General" })
            }
            elseif ($_.Exception.Message -like "*not recognized as the name of a cmdlet*") {
                Write-Log  "Cmdlet wird nicht erkannt: $($_.Exception.Message)" -Type "Warning"
                
                # Spezifische Behandlung für bekannte alte Cmdlets und deren Ersatz
                if ($command -like "*Get-EXORecipient*") {
                    Write-Log  "Versuche Get-Recipient als Alternative zu Get-EXORecipient" -Type "Info"
                    $alternativeCommand = $command -replace "Get-EXORecipient", "Get-Recipient"
                    try {
                        $scriptBlock = [Scriptblock]::Create($alternativeCommand)
                        $result = & $scriptBlock | Out-String
                    } catch {
                        throw "Fehler beim Ausführen des alternativen Befehls: $($_.Exception.Message)"
                    }
                }
                elseif ($command -like "*Get-EXOMailboxStatistics*") {
                    Write-Log  "Versuche Get-MailboxStatistics als Alternative zu Get-EXOMailboxStatistics" -Type "Info"
                    $alternativeCommand = $command -replace "Get-EXOMailboxStatistics", "Get-MailboxStatistics"
                    try {
                        $scriptBlock = [Scriptblock]::Create($alternativeCommand)
                        $result = & $scriptBlock | Out-String
                    } catch {
                        throw "Fehler beim Ausführen des alternativen Befehls: $($_.Exception.Message)"
                    }
                }
                elseif ($command -like "*Get-EXOMailbox*") {
                    Write-Log  "Versuche Get-Mailbox als Alternative zu Get-EXOMailbox" -Type "Info"
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
        Write-Log  "Diagnose abgeschlossen: $($diagnostic.Name)" -Type "Success"
        
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
        Write-Log  "Fehler bei der Diagnose: $errorMsg" -Type "Error"
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
        Write-Log  "Fehler beim Abrufen der Audit-Konfiguration: $($_.Exception.Message)" -Type "Error"
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
        Write-Log  "Fehler beim Abrufen der Weiterleitungsinformationen: $($_.Exception.Message)" -Type "Error"
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
        
        Write-Log  "Führe Mailbox-Audit aus. NavigationType: $NavigationType, InfoType: $InfoType, Mailbox: $Mailbox" -Type "Info"
        
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
        Write-Log  "Fehler beim Abrufen der Informationen: $errorMsg" -Type "Error"
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
        Write-Log  "Fehler beim Abrufen der Postfachinformationen: $($_.Exception.Message)" -Type "Error"
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
        Write-Log  "Fehler beim Abrufen der Postfach-Statistiken: $($_.Exception.Message)" -Type "Error"
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
        Write-Log  "Fehler beim Abrufen der Berechtigungszusammenfassung: $($_.Exception.Message)" -Type "Error"
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
        
        Write-Log  "SendAs-Berechtigung hinzufügen: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfen, ob die Berechtigung bereits existiert
        $existingPermissions = Get-RecipientPermission -Identity $SourceUser -Trustee $TargetUser -ErrorAction SilentlyContinue
        
        if ($existingPermissions) {
            Write-Log  "SendAs-Berechtigung existiert bereits, keine Änderung notwendig" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "SendAs-Berechtigung bereits vorhanden." -Color $script:connectedBrush
            }
            Log-Action "SendAs-Berechtigung bereits vorhanden: $SourceUser -> $TargetUser"
            return $true
        }
        
        # Berechtigung hinzufügen
        Add-RecipientPermission -Identity $SourceUser -Trustee $TargetUser -AccessRights SendAs -Confirm:$false -ErrorAction Stop
        
        Write-Log  "SendAs-Berechtigung erfolgreich hinzugefügt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "SendAs-Berechtigung hinzugefügt." -Color $script:connectedBrush
        }
        Log-Action "SendAs-Berechtigung hinzugefügt: $SourceUser -> $TargetUser"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Hinzufügen der SendAs-Berechtigung: $errorMsg" -Type "Error"
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
        
        Write-Log  "Entferne SendAs-Berechtigung: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfen, ob die Berechtigung existiert
        $existingPermissions = Get-RecipientPermission -Identity $SourceUser -Trustee $TargetUser -ErrorAction SilentlyContinue
        if (-not $existingPermissions) {
            Write-Log  "Keine SendAs-Berechtigung zum Entfernen gefunden" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine SendAs-Berechtigung zum Entfernen gefunden."
            }
            Log-Action "Keine SendAs-Berechtigung zum Entfernen gefunden: $SourceUser -> $TargetUser"
            return $false
        }
        
        # Berechtigung entfernen
        Remove-RecipientPermission -Identity $SourceUser -Trustee $TargetUser -AccessRights SendAs -Confirm:$false -ErrorAction Stop
        
        Write-Log  "SendAs-Berechtigung erfolgreich entfernt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "SendAs-Berechtigung entfernt." -Color $script:connectedBrush
        }
        Log-Action "SendAs-Berechtigung entfernt: $SourceUser -> $TargetUser"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Entfernen der SendAs-Berechtigung: $errorMsg" -Type "Error"
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
        
        Write-Log  "Rufe SendAs-Berechtigungen ab für: $MailboxUser" -Type "Info"
        
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
            Write-Log  "SendAs-Berechtigung verarbeitet: $($permission.Trustee)" -Type "Info"
        }
        
        Write-Log  "SendAs-Berechtigungen abgerufen und verarbeitet: $($processedPermissions.Count) Einträge gefunden" -Type "Success"
        Log-Action "SendAs-Berechtigungen für $MailboxUser abgerufen: $($processedPermissions.Count) Einträge gefunden"
        return $processedPermissions
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der SendAs-Berechtigungen: $errorMsg" -Type "Error"
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
        
        Write-Log  "Füge SendOnBehalf-Berechtigung hinzu: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfen, ob die Berechtigung bereits existiert
        $mailbox = Get-Mailbox -Identity $SourceUser -ErrorAction Stop
        $currentDelegates = $mailbox.GrantSendOnBehalfTo
        
        if ($currentDelegates -contains $TargetUser) {
            Write-Log  "SendOnBehalf-Berechtigung existiert bereits, keine Änderung notwendig" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "SendOnBehalf-Berechtigung bereits vorhanden." -Color $script:connectedBrush
            }
            Log-Action "SendOnBehalf-Berechtigung bereits vorhanden: $SourceUser -> $TargetUser"
            return $true
        }
        
        # Berechtigung hinzufügen (bestehende Berechtigungen beibehalten)
        $newDelegates = $currentDelegates + $TargetUser
        Set-Mailbox -Identity $SourceUser -GrantSendOnBehalfTo $newDelegates -ErrorAction Stop
        
        Write-Log  "SendOnBehalf-Berechtigung erfolgreich hinzugefügt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "SendOnBehalf-Berechtigung hinzugefügt." -Color $script:connectedBrush
        }
        Log-Action "SendOnBehalf-Berechtigung hinzugefügt: $SourceUser -> $TargetUser"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Hinzufügen der SendOnBehalf-Berechtigung: $errorMsg" -Type "Error"
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
        
        Write-Log  "Entferne SendOnBehalf-Berechtigung: $SourceUser -> $TargetUser" -Type "Info"
        
        # Prüfen, ob die Berechtigung existiert
        $mailbox = Get-Mailbox -Identity $SourceUser -ErrorAction Stop
        $currentDelegates = $mailbox.GrantSendOnBehalfTo
        
        if (-not ($currentDelegates -contains $TargetUser)) {
            Write-Log  "Keine SendOnBehalf-Berechtigung zum Entfernen gefunden" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Keine SendOnBehalf-Berechtigung zum Entfernen gefunden."
            }
            Log-Action "Keine SendOnBehalf-Berechtigung zum Entfernen gefunden: $SourceUser -> $TargetUser"
            return $false
        }
        
        # Berechtigung entfernen
        $newDelegates = $currentDelegates | Where-Object { $_ -ne $TargetUser }
        Set-Mailbox -Identity $SourceUser -GrantSendOnBehalfTo $newDelegates -ErrorAction Stop
        
        Write-Log  "SendOnBehalf-Berechtigung erfolgreich entfernt" -Type "Success"
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "SendOnBehalf-Berechtigung entfernt." -Color $script:connectedBrush
        }
        Log-Action "SendOnBehalf-Berechtigung entfernt: $SourceUser -> $TargetUser"
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Entfernen der SendOnBehalf-Berechtigung: $errorMsg" -Type "Error"
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
        
        Write-Log  "Rufe SendOnBehalf-Berechtigungen ab für: $MailboxUser" -Type "Info"
        
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
                Write-Log  "SendOnBehalf-Berechtigung verarbeitet: $delegate" -Type "Info"
            }
        }
        
        Write-Log  "SendOnBehalf-Berechtigungen abgerufen: $($processedDelegates.Count) Einträge gefunden" -Type "Success"
        Log-Action "SendOnBehalf-Berechtigungen für $MailboxUser abgerufen: $($processedDelegates.Count) Einträge gefunden"
        
        return $processedDelegates
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der SendOnBehalf-Berechtigungen: $errorMsg"
        
        # Bei Fehler ein leeres Array zurückgeben, damit die GUI nicht abstürzt
        return @()
    }
}

# -------------------------------------------------
# Abschnitt: Regionaleinstellungen-Funktionen
# -------------------------------------------------

# Hilfsfunktion zur formatierten Ausgabe von Fehlermeldungen
function Get-FormattedError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory = $false)]
        [string]$DefaultText = "Ein unbekannter Fehler ist aufgetreten."
    )

    try {
        if ($null -ne $ErrorRecord) {
            # Versuche, die aussagekräftigste Fehlermeldung zu extrahieren
            $errorMessage = $ErrorRecord.ToString() # Startet mit der Standard-ToString()-Methode

            # Wenn eine Exception vorhanden ist, nutze deren Nachricht (oft detaillierter)
            if ($null -ne $ErrorRecord.Exception) {
                $errorMessage = $ErrorRecord.Exception.Message

                # Füge Nachrichten von inneren Exceptions hinzu, falls vorhanden
                $innerEx = $ErrorRecord.Exception.InnerException
                $depth = 0 # Begrenzung der Tiefe, um Endlosschleifen zu vermeiden
                while ($null -ne $innerEx -and $depth -lt 5) {
                    if (-not [string]::IsNullOrWhiteSpace($innerEx.Message)) {
                        $errorMessage += " --> $($innerEx.Message)"
                    }
                    $innerEx = $innerEx.InnerException
                    $depth++
                }
            }

            # Füge Informationen zum Skript und zur Zeilennummer hinzu, falls verfügbar
            $scriptInfo = ""
            if ($ErrorRecord.InvocationInfo -ne $null) {
                $scriptName = $ErrorRecord.InvocationInfo.ScriptName
                $lineNumber = $ErrorRecord.InvocationInfo.ScriptLineNumber
                if (-not [string]::IsNullOrWhiteSpace($scriptName) -and $lineNumber -gt 0) {
                    $scriptInfo = " (Skript: '$(Split-Path -Leaf $scriptName)', Zeile: $lineNumber)"
                } elseif ($lineNumber -gt 0) {
                     $scriptInfo = " (Zeile: $lineNumber)"
                }
            }

            # Kombiniere die Meldung und die Skriptinformationen
            $formattedMessage = "$errorMessage$scriptInfo"

            # Bereinige die Nachricht von überflüssigen Zeilenumbrüchen am Anfang/Ende
            return $formattedMessage.Trim()

        } else {
            # Wenn kein ErrorRecord übergeben wurde, gib den Standardtext zurück
            return $DefaultText
        }
    } catch {
        # Fallback, falls beim Formatieren der Fehlermeldung selbst ein Fehler auftritt
        Write-Warning "Kritischer Fehler in Get-FormattedError beim Verarbeiten von `$ErrorRecord: $($_.Exception.Message)"
        # Versuche zumindest die ursprüngliche ToString()-Methode oder den Standardtext zurückzugeben
        if ($null -ne $ErrorRecord) {
            try {
                return $ErrorRecord.ToString()
            } catch {
                return $DefaultText
            }
        } else {
            return $DefaultText
        }
    }
}


# Funktion zur Ausführung der Aktion "Regionaleinstellungen anwenden"
function Invoke-SetRegionSettingsAction {
    [CmdletBinding()]
    param()

    # Prüfen, ob verbunden
    if (-not $script:IsConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Nicht verbunden"
        return
    }

    # Werte aus der GUI holen
    # Verwende die korrekten Namen aus dem XAML via $script:-Variablen, die in Initialize-RegionSettingsTab gesetzt wurden
    $mailboxInput = $script:txtRegionMailbox.Text.Trim()
    $selectedLanguageItem = $script:cmbRegionLanguage.SelectedItem
    $selectedTimezoneItem = $script:cmbRegionTimezone.SelectedItem
    $statusTextBlock = $script:txtStatus # Annahme: txtStatus ist der globale Status TextBlock

    # Postfach-Eingabe validieren
    if ([string]::IsNullOrWhiteSpace($mailboxInput)) {
        Show-MessageBox -Message "Bitte geben Sie mindestens eine Postfach-Identität an oder 'ALL'." -Title "Eingabe fehlt"
        return
    }

    # Parameter für Set-MailboxRegionalConfiguration zusammenstellen
    $params = @{}
    $displayParams = @{} # Für Bestätigungsdialog

    # Sprache prüfen
    if ($null -ne $selectedLanguageItem -and $null -ne $selectedLanguageItem.Tag -and $selectedLanguageItem.Tag -ne "") {
        $params.Add("Language", $selectedLanguageItem.Tag)
        $displayParams.Add("Sprache", "$($selectedLanguageItem.Content) ($($selectedLanguageItem.Tag))")
        Write-Log "Ausgewählte Sprache: $($selectedLanguageItem.Content) ($($selectedLanguageItem.Tag))" -Type Debug
    } else {
        $displayParams.Add("Sprache", "(Keine Änderung)")
        Write-Log "Keine Sprache zur Änderung ausgewählt." -Type Debug
    }

    # Zeitzone prüfen
    if ($null -ne $selectedTimezoneItem -and $null -ne $selectedTimezoneItem.Tag -and $selectedTimezoneItem.Tag -ne "") {
        $params.Add("TimeZone", $selectedTimezoneItem.Tag)
        $displayParams.Add("Zeitzone", "$($selectedTimezoneItem.Content) ($($selectedTimezoneItem.Tag))")
         Write-Log "Ausgewählte Zeitzone: $($selectedTimezoneItem.Content) ($($selectedTimezoneItem.Tag))" -Type Debug
    } else {
        $displayParams.Add("Zeitzone", "(Keine Änderung)")
         Write-Log "Keine Zeitzone zur Änderung ausgewählt." -Type Debug
    }

    # Prüfen, ob mindestens eine Einstellung geändert werden soll
    if ($params.Count -eq 0) {
        Show-MessageBox -Message "Bitte wählen Sie mindestens eine Sprache oder Zeitzone zur Änderung aus." -Title "Keine Auswahl"
        return
    }

    # Zielpostfächer ermitteln
    $targetMailboxes = @()
    if ($mailboxInput -eq 'ALL') {
        # Bestätigung für 'ALL' einholen
        $confirmMessage = "WARNUNG: Sie sind dabei, die regionalen Einstellungen für ALLE Benutzerpostfächer zu ändern.`n"
        $confirmMessage += "Sprache: $($displayParams.Sprache)`n"
        $confirmMessage += "Zeitzone: $($displayParams.Zeitzone)`n`n"
        $confirmMessage += "Dies kann sehr lange dauern und sollte mit Vorsicht verwendet werden.`n`n"
        $confirmMessage += "Möchten Sie wirklich fortfahren?"

        $confirmResult = [System.Windows.MessageBox]::Show(
            $confirmMessage,
            "Bestätigung erforderlich",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($confirmResult -ne 'Yes') {
            Update-GuiText -TextElement $statusTextBlock -Message "Vorgang abgebrochen."
            Log-Action "Anwendung der regionalen Einstellungen für 'ALL' abgebrochen."
            return
        }

        Update-GuiText -TextElement $statusTextBlock -Message "Rufe alle Benutzerpostfächer ab (kann dauern)..."
        Log-Action "Beginne Abruf aller Benutzerpostfächer für Regionaleinstellungen-Änderung."
        try {
            # Hinweis: Get-Mailbox ohne Filter kann sehr lange dauern! Ggf. Filter hinzufügen oder Paging implementieren.
            # Es werden nur UserMailboxes berücksichtigt.
            $targetMailboxes = Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -eq 'UserMailbox'} | Select-Object -ExpandProperty Identity -ErrorAction Stop
            if ($targetMailboxes.Count -eq 0) {
                 Show-MessageBox -Message "Keine Benutzerpostfächer gefunden." -Title "Keine Postfächer"
                 Update-GuiText -TextElement $statusTextBlock -Message "Keine Benutzerpostfächer für 'ALL' gefunden."
                 Log-Action "Keine Benutzerpostfächer für 'ALL' gefunden."
                 return
            }
            Update-GuiText -TextElement $statusTextBlock -Message "$($targetMailboxes.Count) Postfächer gefunden. Starte Änderungen..."
            Log-Action "$($targetMailboxes.Count) Postfächer für 'ALL' gefunden."

        } catch {
             $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Abrufen aller Postfächer."
             Show-MessageBox -Message "Fehler beim Abrufen der Postfächer für 'ALL':`n$errorMsg" -Title "Fehler"
             Update-GuiText -TextElement $statusTextBlock -Message "Fehler beim Abrufen der Postfächer."
             Log-Action "Fehler beim Abrufen aller Postfächer: $errorMsg"
             return
        }

    } else {
        # Mehrere Postfächer durch Semikolon getrennt erlauben und leere Einträge filtern
        $targetMailboxes = $mailboxInput -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }

    # Finale Prüfung, ob Postfächer vorhanden sind
    if ($targetMailboxes.Count -eq 0) {
         Show-MessageBox -Message "Keine gültigen Postfach-Identitäten angegeben oder gefunden." -Title "Keine Postfächer"
         Update-GuiText -TextElement $statusTextBlock -Message "Keine gültigen Postfächer zur Verarbeitung angegeben."
         Log-Action "Keine gültigen Postfächer zur Verarbeitung angegeben für '$mailboxInput'."
         return
    }

    # Lange laufende Operation starten (Referenziert die Hilfsfunktion Set-ExoMailboxRegionalSettings)
    Update-GuiText -TextElement $statusTextBlock -Message "Starte Verarbeitung von $($targetMailboxes.Count) Postfach/Postfächern..."
    Log-Action "Starte Set-ExoMailboxRegionalSettings für $($targetMailboxes.Count) Postfächer mit Parametern: $($params.Keys -join ', ')"

    Start-LongRunningOperation -ScriptBlock {
        param($MailboxesToProcess, $Parameters, $FormRef, $StatusTextBlockRef) # Übergebe Referenzen

        # Rufe die eigentliche Verarbeitungsfunktion auf
        # WICHTIG: Übergib die GUI-Elemente an Set-ExoMailboxRegionalSettings für den Fortschritt
        $results = Set-ExoMailboxRegionalSettings -MailboxesToProcess $MailboxesToProcess -Parameters $Parameters -Form $FormRef -StatusTextBlock $StatusTextBlockRef
        return $results # Gib die Ergebnisse zurück
    } -ArgumentList $targetMailboxes, $params, $script:Form, $statusTextBlock -CompletionScript {
        param($OperationResult)

        # $OperationResult ist das Hashtable von Set-ExoMailboxRegionalSettings (@{Success = [...]; Failed = [...]})
        $successCount = $OperationResult.Success.Count
        $failedCount = $OperationResult.Failed.Count
        $totalCount = $successCount + $failedCount

        $summaryMessage = "Regionaleinstellungen verarbeitet: $successCount erfolgreich, $failedCount fehlgeschlagen (von $totalCount)."
        Update-GuiText -TextElement $statusTextBlock -Message $summaryMessage
        Log-Action $summaryMessage

        if ($failedCount -gt 0) {
            # Fehlerdetails aufbereiten (ggf. kürzen, wenn zu lang)
            $errorDetails = $OperationResult.Failed | ForEach-Object { "  - $($_.Mailbox): $($_.Error)" }
            $maxErrorLines = 10 # Maximale Anzahl anzuzeigender Fehlerzeilen in der MessageBox
            $truncated = $false
            if ($errorDetails.Count -gt $maxErrorLines) {
                $errorDetails = $errorDetails[0..($maxErrorLines - 1)]
                $truncated = $true
            }
            $errorDetailsString = $errorDetails -join "`n"
            if ($truncated) {
                $errorDetailsString += "`n  ... (weitere Fehler im Log)"
            }

            # Ergebnis auch im Result-TextBlock anzeigen, falls vorhanden
            if ($null -ne $script:txtRegionResult) {
                 Update-GuiText -TextElement $script:txtRegionResult -Message "Fehler bei Verarbeitung:`n$errorDetailsString"
            }

            Show-MessageBox -Message "Einige regionale Einstellungen konnten nicht angewendet werden:`n$errorDetailsString" -Title "Fehler bei Verarbeitung"
            # Detaillierte Fehler im Log speichern
            Log-Action "Fehlerdetails Regionaleinstellungen: $($OperationResult.Failed | ConvertTo-Json -Depth 3)"
        } elseif ($successCount -gt 0) {
             # Ergebnis auch im Result-TextBlock anzeigen
             if ($null -ne $script:txtRegionResult) {
                 Update-GuiText -TextElement $script:txtRegionResult -Message "$successCount Postfach/Postfächer erfolgreich aktualisiert."
             }
             Show-MessageBox -Message "Regionaleinstellungen für $successCount Postfach/Postfächer erfolgreich angewendet." -Title "Erfolg"
        } else {
             # Ergebnis auch im Result-TextBlock anzeigen
             if ($null -ne $script:txtRegionResult) {
                  Update-GuiText -TextElement $script:txtRegionResult -Message "Keine Änderungen vorgenommen oder alle Operationen fehlgeschlagen."
             }
             Show-MessageBox -Message "Keine Änderungen vorgenommen oder alle Operationen fehlgeschlagen." -Title "Kein Erfolg"
        }
    }
}
# Ende Invoke-SetRegionSettingsAction

# Funktion zur Ausführung der Aktion "Aktuelle Einstellungen abrufen"
function Invoke-GetRegionSettingsAction {
    [CmdletBinding()]
    param()

    # Prüfen, ob verbunden
    if (-not $script:IsConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Nicht verbunden"
        return
    }

    # Werte aus der GUI holen (verwende die korrekten Namen aus XAML)
    # Nehme an, die Textbox zum Abrufen heißt txtGetRegionMailbox (gemäß XAML)
    $mailboxId = $script:txtGetRegionMailbox.Text.Trim()
    $statusTextBlock = $script:txtStatus # Annahme für Status
    $resultTextBlock = $script:txtRegionResult # Annahme für Ergebnis-Textbox (gemäß XAML)
    $languageComboBox = $script:cmbRegionLanguage # Gemäß XAML
    $timezoneComboBox = $script:cmbRegionTimezone # Gemäß XAML


    # Postfach-Eingabe validieren
    if ([string]::IsNullOrWhiteSpace($mailboxId)) {
        Show-MessageBox -Message "Bitte geben Sie die E-Mail-Adresse des Postfachs an, dessen Einstellungen abgerufen werden sollen." -Title "Eingabe fehlt"
        return
    }
    # Stelle sicher, dass nicht 'ALL' eingegeben wurde, da nur ein Postfach abgefragt werden kann
     if ($mailboxId -eq 'ALL') {
        Show-MessageBox -Message "Bitte geben Sie eine einzelne Postfach-Identität zum Abrufen an (nicht 'ALL')." -Title "Ungültige Eingabe"
        return
    }
     # Optional: E-Mail-Validierung hinzufügen
    # if (-not (Validate-Email -Email $mailboxId)) {
    #    Show-MessageBox -Message "Bitte geben Sie eine gültige E-Mail-Adresse ein." -Title "Ungültige Eingabe" 
    #    return
    # }


    # Lange laufende Operation starten (Referenziert die Hilfsfunktion Get-ExoMailboxRegionalSettings)
    Update-GuiText -TextElement $statusTextBlock -Message "Rufe regionale Einstellungen für '$mailboxId' ab..."
    Log-Action "Starte Get-ExoMailboxRegionalSettings für '$mailboxId'."

    # Ergebnis-Textfeld leeren
    if ($null -ne $resultTextBlock) {
         Update-GuiText -TextElement $resultTextBlock -Message ""
    }

    Start-LongRunningOperation -ScriptBlock {
        param($Identity)
        # Rufe die vorhandene Funktion zum Abrufen der Daten auf
        $result = Get-ExoMailboxRegionalSettings -MailboxIdentity $Identity
        return $result # Gib das Ergebnis zurück (@{Success=$true/$false; Error=...; Result=...})
    } -ArgumentList $mailboxId -CompletionScript {
        param($OperationResult)

        if ($OperationResult.Success) {
            $regionalSettings = $OperationResult.Result # Enthält das Objekt von Get-MailboxRegionalConfiguration
            $retrievedMailboxId = $regionalSettings.Identity # Korrekte ID aus dem Ergebnis holen
            $message = "Regionale Einstellungen für '$retrievedMailboxId' abgerufen: Sprache=$($regionalSettings.Language), Zeitzone=$($regionalSettings.TimeZone)"

            Update-GuiText -TextElement $statusTextBlock -Message "Abruf erfolgreich."
            Log-Action $message

            # Ergebnis im Result-TextBlock anzeigen
            if ($null -ne $resultTextBlock) {
                 $resultOutput = @"
Abgerufene Einstellungen für: $retrievedMailboxId
--------------------------------------------------
Sprache (Language): $($regionalSettings.Language)
Zeitzone (TimeZone): $($regionalSettings.TimeZone)
Datumsformat (DateFormat): $($regionalSettings.DateFormat)
Zeitformat (TimeFormat): $($regionalSettings.TimeFormat)
"@
                 Update-GuiText -TextElement $resultTextBlock -Message $resultOutput
            }


            # --- WICHTIG: GUI ComboBoxen aktualisieren ---
            try {
                # Sprache auswählen
                $langItem = $languageComboBox.Items | Where-Object { $_.Tag -eq $regionalSettings.Language } | Select-Object -First 1
                if ($null -ne $langItem) {
                    $languageComboBox.Dispatcher.InvokeAsync({ $languageComboBox.SelectedItem = $langItem }) | Out-Null
                    Write-Log "Sprache '$($regionalSettings.Language)' in ComboBox ausgewählt." -Type Debug
                } else {
                    # Sprache nicht in Liste? Fallback oder Text setzen (wenn ComboBox editierbar)
                    $warnMsg = "Abgerufene Sprache '$($regionalSettings.Language)' nicht in ComboBox gefunden. Fallback auf '(Keine Änderung)'."
                    Write-Log $warnMsg -Type Warning
                    Log-Action "WARNUNG: $warnMsg"
                    # Optional: $languageComboBox.Text = $regionalSettings.Language (wenn IsEditable=true)
                    $languageComboBox.Dispatcher.InvokeAsync({ $languageComboBox.SelectedIndex = 0 }) | Out-Null # Fallback auf "(Keine Änderung)"
                }

                # Zeitzone auswählen
                $tzItem = $timezoneComboBox.Items | Where-Object { $_.Tag -eq $regionalSettings.TimeZone } | Select-Object -First 1
                if ($null -ne $tzItem) {
                    $timezoneComboBox.Dispatcher.InvokeAsync({ $timezoneComboBox.SelectedItem = $tzItem }) | Out-Null
                     Write-Log "Zeitzone '$($regionalSettings.TimeZone)' in ComboBox ausgewählt." -Type Debug
                } else {
                    # Zeitzone nicht in Liste? Fallback oder Text setzen (ist editierbar laut XAML)
                    $warnMsg = "Abgerufene Zeitzone '$($regionalSettings.TimeZone)' nicht in ComboBox gefunden. Setze Text."
                     Write-Log $warnMsg -Type Warning
                     Log-Action "WARNUNG: $warnMsg"
                     # Da die Zeitzonen-ComboBox editierbar ist, können wir den Text setzen
                     $timezoneComboBox.Dispatcher.InvokeAsync({ $timezoneComboBox.Text = $regionalSettings.TimeZone }) | Out-Null
                     # Alternative: Fallback auf "(Keine Änderung)" -> $timezoneComboBox.Dispatcher.InvokeAsync({ $timezoneComboBox.SelectedIndex = 0 }) | Out-Null
                }
            } catch {
                $errorMsgGui = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Aktualisieren der GUI nach Abruf."
                Update-GuiText -TextElement $statusTextBlock -Message "Fehler bei GUI-Update: $errorMsgGui"
                Log-Action "FEHLER: $errorMsgGui"
                 # Ergebnis im Result-TextBlock anzeigen
                if ($null -ne $resultTextBlock) {
                     Update-GuiText -TextElement $resultTextBlock -Message "Fehler beim Aktualisieren der Dropdowns nach Abruf:`n$errorMsgGui"
                }
                Show-MessageBox -Message "Die Einstellungen wurden abgerufen, aber die Dropdowns konnten nicht aktualisiert werden: $errorMsgGui" -Title "GUI Fehler" 
            }
            # --- Ende GUI ComboBoxen Aktualisierung ---

        } else {
            # Fehler beim Abrufen mit Get-ExoMailboxRegionalSettings
            $errorMsg = $OperationResult.Error
            Update-GuiText -TextElement $statusTextBlock -Message "Fehler beim Abrufen: $errorMsg"
            Log-Action "FEHLER beim Abrufen der regionalen Einstellungen für '$mailboxId': $errorMsg"
             # Ergebnis im Result-TextBlock anzeigen
            if ($null -ne $resultTextBlock) {
                 Update-GuiText -TextElement $resultTextBlock -Message "Fehler beim Abrufen für '$mailboxId':`n$errorMsg"
            }
            Show-MessageBox -Message "Fehler beim Abrufen der regionalen Einstellungen für '$mailboxId':`n$errorMsg" -Title "Fehler" 

             # Setze ComboBoxen auf Standard zurück bei Fehler
             try {
                 $languageComboBox.Dispatcher.InvokeAsync({ $languageComboBox.SelectedIndex = 0 }) | Out-Null
                 $timezoneComboBox.Dispatcher.InvokeAsync({ $timezoneComboBox.SelectedIndex = 0 }) | Out-Null
             } catch {}
        }
    }
}
# Ende Invoke-GetRegionSettingsAction


# Hilfsfunktion zum Befüllen der Sprachen-ComboBox
function Populate-LanguageComboBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ComboBox]$ComboBox
    )
    try {
        Write-Log "Befülle Sprachen-ComboBox..." -Type Debug

        # Aktuelle Auswahl merken (Tag oder Text)
        $currentValue = $null
        if ($ComboBox.SelectedItem -ne $null -and $ComboBox.SelectedItem.Tag -ne $null) {
            $currentValue = $ComboBox.SelectedItem.Tag
        } elseif (-not [string]::IsNullOrEmpty($ComboBox.Text)) {
             $currentValue = $ComboBox.Text # Fallback auf Text
        }

        $ComboBox.Items.Clear()

        # Standard "(Keine Änderung)" hinzufügen
        $noChangeItem = New-Object System.Windows.Controls.ComboBoxItem
        $noChangeItem.Content = "(Keine Änderung)"
        $noChangeItem.Tag = "" # Leerer Tag ist wichtig für die Logik
        $noChangeItem.IsSelected = $true # Standardauswahl
        [void]$ComboBox.Items.Add($noChangeItem)

        # System-Sprachen laden (spezifische Kulturen für Exchange)
        $cultures = [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::SpecificCultures) | Sort-Object -Property DisplayName
        foreach ($culture in $cultures) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $culture.DisplayName # z.B. "Englisch (Vereinigte Staaten)"
            $item.Tag = $culture.Name          # z.B. "en-US"
            [void]$ComboBox.Items.Add($item)
        }

         # Vorherige Auswahl versuchen wiederherzustellen
        if (-not [string]::IsNullOrEmpty($currentValue)) {
            $itemToSelect = $ComboBox.Items | Where-Object { $_.Tag -eq $currentValue -or $_.Content -eq $currentValue } | Select-Object -First 1
            if ($null -ne $itemToSelect) {
                $ComboBox.SelectedItem = $itemToSelect
            } else {
                # Wenn nicht gefunden, Text setzen (falls editierbar und nicht Standard)
                if ($ComboBox.IsEditable -and $currentValue -ne "(Keine Änderung)") {
                     $ComboBox.Text = $currentValue
                } else {
                     $ComboBox.SelectedItem = $noChangeItem # Fallback auf Standard
                }
            }
        } else {
             # Standardmäßig "(Keine Änderung)" auswählen, falls nichts vorher ausgewählt war
             $ComboBox.SelectedItem = $noChangeItem
        }
        Write-Log "Sprachen-ComboBox erfolgreich befüllt." -Type Debug
        return $true
    } catch {
        Write-Log "Fehler beim Befüllen der Sprachen-ComboBox: $($_.Exception.Message)" -Type Warning
        Log-Action "Warnung: Sprachen konnten nicht dynamisch geladen werden. Fallback auf XAML-Definition (falls vorhanden)."
        # Im Fehlerfall bleiben ggf. XAML-definierte Items oder die Liste ist leer bis auf "(Keine Änderung)"
        return $false
    }
}

# Hilfsfunktion zum Befüllen der Zeitzonen-ComboBox
function Populate-TimezoneComboBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ComboBox]$ComboBox
    )
    try {
        Write-Log "Befülle Zeitzonen-ComboBox..." -Type Debug

        # Aktuelle Auswahl merken (Tag oder Text)
        $currentValue = $null
         if ($ComboBox.SelectedItem -ne $null -and $ComboBox.SelectedItem.Tag -ne $null) {
            $currentValue = $ComboBox.SelectedItem.Tag
        } elseif (-not [string]::IsNullOrEmpty($ComboBox.Text)) {
             $currentValue = $ComboBox.Text # Fallback auf Text
        }

        $ComboBox.Items.Clear()

        # Standard "(Keine Änderung)" hinzufügen
        $noChangeItem = New-Object System.Windows.Controls.ComboBoxItem
        $noChangeItem.Content = "(Keine Änderung)"
        $noChangeItem.Tag = "" # Leerer Tag ist wichtig für die Logik
        $noChangeItem.IsSelected = $true # Standardauswahl
        [void]$ComboBox.Items.Add($noChangeItem)

        # System-Zeitzonen laden
        $timezones = [System.TimeZoneInfo]::GetSystemTimeZones() | Sort-Object -Property Id
        foreach ($tz in $timezones) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $tz.DisplayName # Anzeige freundlicher Name (z.B. "(UTC+01:00) Amsterdam, Berlin, Bern, Rom, Stockholm, Wien")
            $item.Tag = $tz.Id             # Interner ID als Wert (z.B. "W. Europe Standard Time")
            [void]$ComboBox.Items.Add($item)
        }

        # Vorherige Auswahl versuchen wiederherzustellen
        if (-not [string]::IsNullOrEmpty($currentValue)) {
            $itemToSelect = $ComboBox.Items | Where-Object { $_.Tag -eq $currentValue -or $_.Content -eq $currentValue } | Select-Object -First 1
            if ($null -ne $itemToSelect) {
                $ComboBox.SelectedItem = $itemToSelect
            } else {
                 # Wenn nicht gefunden, Text setzen (falls editierbar und nicht Standard)
                 if ($ComboBox.IsEditable -and $currentValue -ne "(Keine Änderung)") {
                     $ComboBox.Text = $currentValue
                 } else {
                     $ComboBox.SelectedItem = $noChangeItem # Fallback auf Standard
                 }
            }
        } else {
             # Standardmäßig "(Keine Änderung)" auswählen
             $ComboBox.SelectedItem = $noChangeItem
        }
        Write-Log "Zeitzonen-ComboBox erfolgreich befüllt." -Type Debug
        return $true
    } catch {
        Write-Log "Fehler beim Befüllen der Zeitzonen-ComboBox: $($_.Exception.Message)" -Type Warning
        Log-Action "Warnung: Zeitzonen konnten nicht dynamisch geladen werden. Fallback auf XAML-Definition (falls vorhanden)."
        # Im Fehlerfall bleiben ggf. XAML-definierte Items oder die Liste ist leer bis auf "(Keine Änderung)"
        return $false
    }
}


# Funktion zum Hinzufügen eines Standard-Elements zu einer ComboBox
function Add-ComboBoxDefaultItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.ComboBox]$Control,

        [Parameter(Mandatory = $true)]
        [string]$DefaultItemText,

        [Parameter(Mandatory = $true)]
        $DefaultItemValue
    )

    try {
        # Sicherstellen, dass das Control gültig ist
        if ($null -eq $Control) {
            Write-Log "Add-ComboBoxDefaultItem: Control ist null." -Type Warning
            return
        }

        # Standard-Element erstellen
        $defaultItem = New-Object System.Windows.Controls.ComboBoxItem
        $defaultItem.Content = $DefaultItemText
        $defaultItem.Tag = $DefaultItemValue # Wert im Tag speichern

        # Element zur ComboBox hinzufügen (im UI-Thread)
        $Control.Dispatcher.InvokeAsync({
            param($ItemToAdd)
            # Prüfen, ob das Element bereits existiert (basierend auf Tag)
            $exists = $false
            foreach ($item in $Control.Items) {
                if ($item -is [System.Windows.Controls.ComboBoxItem] -and $item.Tag -eq $ItemToAdd.Tag) {
                    $exists = $true
                    break
                }
            }
            if (-not $exists) {
                $Control.Items.Insert(0, $ItemToAdd) # Am Anfang einfügen
                # Optional: Standardmäßig auswählen, wenn die Box leer war
                if ($Control.Items.Count -eq 1) {
                    $Control.SelectedIndex = 0
                }
            }
        }, "Normal", $defaultItem) | Out-Null

    } catch {
        # Use ${} for clarity
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Hinzufügen des Standard-Elements '${DefaultItemText}' zur ComboBox."
        Write-Log $errorMsg -Type Error
        Log-Action "FEHLER: Add-ComboBoxDefaultItem - ${errorMsg}"
    }
}

# Funktion zum Aktualisieren der Elemente einer ComboBox aus einer Datenquelle
function Update-ComboBoxItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.ComboBox]$Control,

        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$DataSource,

        [Parameter(Mandatory = $true)]
        [string]$DisplayMember, # Eigenschaft für die Anzeige

        [Parameter(Mandatory = $true)]
        [string]$ValueMember,   # Eigenschaft für den Wert (im Tag gespeichert)

        [string]$DefaultItemText = "(Auswählen)", # Optionaler Text für das Standardelement
        $DefaultItemValue = $null                 # Optionaler Wert für das Standardelement
    )

    try {
        # Sicherstellen, dass das Control gültig ist
        if ($null -eq $Control) {
            Write-Log "Update-ComboBoxItems: Control ist null." -Type Warning
            return
        }

        # Ausführung im UI-Thread sicherstellen
        $Control.Dispatcher.InvokeAsync({
            # ComboBox leeren
            $Control.Items.Clear()

            # Standard-Element hinzufügen, falls definiert
            if (-not [string]::IsNullOrEmpty($DefaultItemText)) {
                Add-ComboBoxDefaultItem -Control $Control -DefaultItemText $DefaultItemText -DefaultItemValue $DefaultItemValue
            }

            # Elemente aus der Datenquelle hinzufügen
            if ($null -ne $DataSource) {
                foreach ($item in $DataSource) {
                    try {
                        $display = $item.$DisplayMember
                        $value = $item.$ValueMember

                        $comboBoxItem = New-Object System.Windows.Controls.ComboBoxItem
                        $comboBoxItem.Content = $display
                        $comboBoxItem.Tag = $value # Wert im Tag speichern

                        $Control.Items.Add($comboBoxItem)
                    } catch {
                        # Use ${} for clarity
                        $itemIdentifier = try { $item | Out-String -Stream } catch { "Unbekanntes Element" }
                        Write-Log "Fehler beim Verarbeiten eines Elements für ComboBox: $($_.Exception.Message). Element: ${itemIdentifier}" -Type Warning
                        # Überspringe dieses Element und fahre fort
                    }
                }
            }

            # Standardmäßig das erste Element (meist das DefaultItem) auswählen
            if ($Control.Items.Count -gt 0) {
                $Control.SelectedIndex = 0
            }

        }, "Normal") | Out-Null # InvokeAsync

    } catch {
        # Use ${} for clarity
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Aktualisieren der ComboBox-Elemente."
        Write-Log $errorMsg -Type Error
        Log-Action "FEHLER: Update-ComboBoxItems - ${errorMsg}"
    }
}
# Funktion zum Abrufen der regionalen Einstellungen eines Postfachs
function Get-ExoMailboxRegionalSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxIdentity
    )

    # Prüfen, ob eine Verbindung besteht
    if (-not $script:IsConnected) {
        Write-Log "Get-ExoMailboxRegionalSettings: Keine Verbindung zu Exchange Online." -Type Warning
        # Rückgabe eines Fehlerobjekts oder einer Nachricht
        return [PSCustomObject]@{
            Success = $false
            Error   = "Keine Verbindung zu Exchange Online hergestellt."
            Result  = $null
        }
    }

    # Eingabe validieren
    if ([string]::IsNullOrWhiteSpace($MailboxIdentity)) {
        Write-Log "Get-ExoMailboxRegionalSettings: Keine Postfach-Identität angegeben." -Type Warning
        return [PSCustomObject]@{
            Success = $false
            Error   = "Keine Postfach-Identität angegeben."
            Result  = $null
        }
    }

    try {
        # Use ${} for clarity
        Write-Log "Rufe regionale Einstellungen für ${MailboxIdentity} ab..." -Type Debug
        # Use ${} for clarity
        Log-Action "Versuche, regionale Einstellungen für ${MailboxIdentity} abzurufen."

        # Regionale Einstellungen abrufen
        $regionalSettings = Get-MailboxRegionalConfiguration -Identity $MailboxIdentity -ErrorAction Stop

        # Erfolgreiches Ergebnis zurückgeben
        # Use ${} for clarity
        Write-Log "Regionale Einstellungen für ${MailboxIdentity} erfolgreich abgerufen." -Type Success
        # Use $() for expressions, ${} for clarity
        Log-Action "Regionale Einstellungen für ${MailboxIdentity} erfolgreich abgerufen: Sprache=$($regionalSettings.Language), Zeitzone=$($regionalSettings.TimeZone)"

        return [PSCustomObject]@{
            Success = $true
            Error   = $null
            Result  = $regionalSettings # Das gesamte Objekt zurückgeben für Flexibilität
        }

    } catch {
        # Fehler behandeln und zurückgeben
        # Use ${} for clarity in default text
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Abrufen der regionalen Einstellungen für ${MailboxIdentity}."
        Write-Log $errorMsg -Type Error
        # Use ${} for clarity
        Log-Action "FEHLER: Get-ExoMailboxRegionalSettings - ${errorMsg}"

        return [PSCustomObject]@{
            Success = $false
            Error   = $errorMsg
            Result  = $null
        }
    }
}

# Funktion zum Setzen der regionalen Einstellungen für ein oder mehrere Postfächer
function Set-ExoMailboxRegionalSettings {
    [CmdletBinding(SupportsShouldProcess = $true)] # Hinzugefügt für -WhatIf/-Confirm Potenzial, obwohl -Confirm intern auf $false gesetzt wird
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[string]]$MailboxesToProcess, # Erwartet eine Liste von Postfach-Identitäten

        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters, # Hashtable mit den zu setzenden Parametern (z.B. @{Language="de-DE"; TimeZone="W. Europe Standard Time"})

        # Optionale Parameter für die GUI-Fortschrittsanzeige aus einem Hintergrundthread
        [Parameter(Mandatory = $false)]
        [System.Windows.Window]$Form,

        [Parameter(Mandatory = $false)]
        [System.Windows.Controls.TextBlock]$StatusTextBlock
    )

    # Initialisiert ein Hashtable zur Speicherung der Ergebnisse
    $results = @{
        Success = [System.Collections.Generic.List[string]]::new()
        Failed  = [System.Collections.Generic.List[object]]::new() # Liste für Objekte mit Mailbox und Fehlerdetails
    }
    $total = $MailboxesToProcess.Count
    $processed = 0

    # Verarbeitet jedes Postfach in der übergebenen Liste
    foreach ($mailbox in $MailboxesToProcess) {
        $processed++
        $progress = [int](($processed / $total) * 100)

        # Versucht, den Fortschritt im GUI-Thread zu aktualisieren, falls die entsprechenden Steuerelemente übergeben wurden
        if ($null -ne $Form -and $null -ne $StatusTextBlock -and $null -ne $Form.Dispatcher) {
            try {
                # Sendet die Aktualisierung an den UI-Thread
                $Form.Dispatcher.InvokeAsync({
                    param($Mailbox, $Processed, $Total, $Progress, $CurrentStatusTextBlock)
                    # Stellt sicher, dass das TextBlock-Element noch gültig ist
                    if ($null -ne $CurrentStatusTextBlock) {
                        # Aktualisiert den Text des Status-TextBlocks
                        # Hinweis: Die Funktion Update-GuiText ist hier nicht direkt im Scope, daher direkte Zuweisung.
                        $CurrentStatusTextBlock.Text = "Verarbeite ${Mailbox} (${Processed}/${Total})... [${Progress}%]"
                    }
                }, "Normal", $mailbox, $processed, $total, $progress, $StatusTextBlock) | Out-Null # Out-Null unterdrückt die Ausgabe des Dispatcher-Objekts
            }
            catch {
                # Gibt eine Warnung aus, falls die GUI-Aktualisierung fehlschlägt
                Write-Warning "Fehler beim Aktualisieren des GUI-Status aus Hintergrundthread für (${mailbox}): $($_.Exception.Message)"
                Log-Action "WARNUNG: Fehler beim GUI-Update im Hintergrund für (${mailbox}): $($_.Exception.Message)"
            }
        }
        else {
            # Fallback-Ausgabe in der Konsole, wenn keine GUI-Elemente für den Fortschritt bereitgestellt werden
            Write-Verbose "Verarbeite ${mailbox} (${processed}/${total})... [${progress}%]"
        }

        # Führt die eigentliche Konfigurationsänderung durch
        try {
            # Protokolliert die Aktion und die verwendeten Parameter detailliert
            # Use ${} for clarity and $() for expression
            Write-Log "Setze regionale Einstellungen für ${mailbox} mit Parametern: $($Parameters | Out-String)" -Type Debug
            Log-Action "Versuche Set-MailboxRegionalConfiguration für ${mailbox} mit Parametern: $($Parameters.Keys -join ', ')"

            # Ruft das Exchange Online Cmdlet auf, um die Einstellungen zu ändern
            if ($PSCmdlet.ShouldProcess($mailbox, "Regionale Einstellungen anwenden (Sprache/Zeitzone)")) {
                 Set-MailboxRegionalConfiguration -Identity $mailbox @Parameters -Confirm:$false -ErrorAction Stop
            }

            # Fügt das Postfach zur Liste der erfolgreich verarbeiteten hinzu
            $results.Success.Add($mailbox)
            # Use ${} for clarity
            Log-Action "Regionale Einstellungen für ${mailbox} erfolgreich gesetzt."
        }
        catch {
            # Fängt Fehler ab, die beim Setzen der Einstellungen für ein einzelnes Postfach auftreten
            # Use ${} for clarity in default text
            $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Setzen der Einstellungen für ${mailbox}."
            # Fügt das Postfach und die Fehlermeldung zur Liste der fehlgeschlagenen Versuche hinzu
            $results.Failed.Add(@{ Mailbox = $mailbox; Error = $errorMsg })
            # Use ${} for clarity
            Log-Action "FEHLER beim Setzen der regionalen Einstellungen für ${mailbox}: ${errorMsg}"
            # Optional: Write-Error oder Write-Warning hier ausgeben, wenn die Funktion auch interaktiv genutzt wird
            Write-Warning "Fehler bei Postfach ${mailbox}: ${errorMsg}"
        }
    }

    # Gibt das Ergebnis-Hashtable zurück
    return $results
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
        Write-Log  "Erstelle neue Gruppe: $GroupName ($GroupType)" -Type "Info"
        
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
                    Write-Log  "Mitglied $member zu Gruppe $GroupName hinzugefügt" -Type "Info"
                }
                catch {
                    Write-Log  "Fehler beim Hinzufügen von $member zu Gruppe $GroupName - $($_.Exception.Message)" -Type "Warning"
                }
            }
        }
        
        Write-Log  "Gruppe $GroupName erfolgreich erstellt" -Type "Success"
        Log-Action "Gruppe $GroupName ($GroupType) mit E-Mail $GroupEmail erstellt"
        
        # Status aktualisieren
        if ($null -ne $txtStatus) {
            $txtStatus.Text = "Gruppe $GroupName erfolgreich erstellt."
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Erstellen der Gruppe: $errorMsg" -Type "Error"
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
        Write-Log  "Lösche Gruppe: $GroupName" -Type "Info"
        
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
            Write-Log  "Microsoft 365-Gruppe $GroupName erfolgreich gelöscht" -Type "Success"
        }
        else {
            Remove-DistributionGroup -Identity $GroupName -Confirm:$false -ErrorAction Stop
            Write-Log  "Verteilerliste/Sicherheitsgruppe $GroupName erfolgreich gelöscht" -Type "Success"
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
        Write-Log  "Fehler beim Löschen der Gruppe: $errorMsg" -Type "Error"
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
        Write-Log  "Füge $MemberIdentity zu Gruppe $GroupName hinzu" -Type "Info"
        
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
            Write-Log  "$MemberIdentity erfolgreich zur Microsoft 365-Gruppe $GroupName hinzugefügt" -Type "Success"
        }
        else {
            Add-DistributionGroupMember -Identity $GroupName -Member $MemberIdentity -ErrorAction Stop
            Write-Log  "$MemberIdentity erfolgreich zur Gruppe $GroupName hinzugefügt" -Type "Success"
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
        Write-Log  "Fehler beim Hinzufügen des Benutzers zur Gruppe: $errorMsg" -Type "Error"
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
        Write-Log  "Entferne $MemberIdentity aus Gruppe $GroupName" -Type "Info"
        
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
            Write-Log  "$MemberIdentity erfolgreich aus Microsoft 365-Gruppe $GroupName entfernt" -Type "Success"
        }
        else {
            Remove-DistributionGroupMember -Identity $GroupName -Member $MemberIdentity -Confirm:$false -ErrorAction Stop
            Write-Log  "$MemberIdentity erfolgreich aus Gruppe $GroupName entfernt" -Type "Success"
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
        Write-Log  "Fehler beim Entfernen des Benutzers aus der Gruppe: $errorMsg" -Type "Error"
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
        Write-Log  "Rufe Mitglieder der Gruppe $GroupName ab" -Type "Info"
        
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
        
        Write-Log  "Mitglieder der Gruppe $GroupName erfolgreich abgerufen: $($memberList.Count)" -Type "Success"
        
        return $memberList
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der Gruppenmitglieder: $errorMsg" -Type "Error"
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
        Write-Log  "Rufe Einstellungen der Gruppe $GroupName ab" -Type "Info"
        
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
        Write-Log  "Fehler beim Abrufen der Gruppeneinstellungen: $errorMsg" -Type "Error"
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
        Write-Log  "Aktualisiere Einstellungen für Gruppe $GroupName" -Type "Info"
        
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
                
                Write-Log  "Microsoft 365-Gruppe $GroupName erfolgreich aktualisiert" -Type "Success"
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
            
            Write-Log  "Gruppe $GroupName erfolgreich aktualisiert" -Type "Success"
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
        Write-Log  "Fehler beim Aktualisieren der Gruppeneinstellungen: $errorMsg" -Type "Error"
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
        Write-Log  "Erstelle neue Shared Mailbox: $Name mit Adresse $EmailAddress" -Type "Info"
        New-Mailbox -Name $Name -PrimarySmtpAddress $EmailAddress -Shared -ErrorAction Stop
        Write-Log  "Shared Mailbox $Name erfolgreich erstellt" -Type "Success"
        Log-Action "Shared Mailbox $Name ($EmailAddress) erfolgreich erstellt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Erstellen der Shared Mailbox: $errorMsg" -Type "Error"
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
        Write-Log  "Konvertiere Postfach zu Shared Mailbox: $Identity" -Type "Info"
        Set-Mailbox -Identity $Identity -Type Shared -ErrorAction Stop
        Write-Log  "Postfach $Identity erfolgreich zu Shared Mailbox konvertiert" -Type "Success"
        Log-Action "Postfach $Identity erfolgreich zu Shared Mailbox konvertiert"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Konvertieren des Postfachs: $errorMsg" -Type "Error"
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
        Write-Log  "Füge Shared Mailbox Berechtigung hinzu: $PermissionType für $User auf $Mailbox" -Type "Info"
        
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
        
        Write-Log  "Shared Mailbox Berechtigung erfolgreich hinzugefügt" -Type "Success"
        Log-Action "Shared Mailbox Berechtigung $PermissionType für $User auf $Mailbox hinzugefügt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Hinzufügen der Shared Mailbox Berechtigung: $errorMsg" -Type "Error"
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
        Write-Log  "Entferne Shared Mailbox Berechtigung: $PermissionType für $User auf $Mailbox" -Type "Info"
        
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
        
        Write-Log  "Shared Mailbox Berechtigung erfolgreich entfernt" -Type "Success"
        Log-Action "Shared Mailbox Berechtigung $PermissionType für $User auf $Mailbox entfernt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Entfernen der Shared Mailbox Berechtigung: $errorMsg" -Type "Error"
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
        Write-Log  "Rufe Berechtigungen für Shared Mailbox ab: $Mailbox" -Type "Info"
        
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
        
        Write-Log  "Shared Mailbox Berechtigungen erfolgreich abgerufen: $($permissions.Count) Einträge" -Type "Success"
        Log-Action "Shared Mailbox Berechtigungen für $Mailbox abgerufen: $($permissions.Count) Einträge"
        
        return $permissions
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der Shared Mailbox Berechtigungen: $errorMsg" -Type "Error"
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
        Write-Log  "Aktualisiere AutoMapping für Shared Mailbox $Mailbox auf $AutoMapping" -Type "Info"
        
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
            Write-Log  "AutoMapping für $user auf $Mailbox aktualisiert" -Type "Info"
        }
        
        Write-Log  "AutoMapping für Shared Mailbox erfolgreich aktualisiert" -Type "Success"
        Log-Action "AutoMapping für Shared Mailbox $Mailbox auf $AutoMapping gesetzt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Aktualisieren des AutoMapping: $errorMsg" -Type "Error"
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
        Write-Log  "Setze Weiterleitung für Shared Mailbox $Mailbox auf $ForwardingAddress" -Type "Info"
        
        if ([string]::IsNullOrEmpty($ForwardingAddress)) {
            # Weiterleitung entfernen
            Set-Mailbox -Identity $Mailbox -ForwardingAddress $null -ForwardingSmtpAddress $null -ErrorAction Stop
            Write-Log  "Weiterleitung für Shared Mailbox erfolgreich entfernt" -Type "Success"
        } else {
            # Weiterleitung setzen
            Set-Mailbox -Identity $Mailbox -ForwardingSmtpAddress $ForwardingAddress -DeliverToMailboxAndForward $true -ErrorAction Stop
            Write-Log  "Weiterleitung für Shared Mailbox erfolgreich gesetzt" -Type "Success"
        }
        
        Log-Action "Weiterleitung für Shared Mailbox $Mailbox auf $ForwardingAddress gesetzt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Setzen der Weiterleitung: $errorMsg" -Type "Error"
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
        Write-Log  "Setze GAL-Sichtbarkeit für Shared Mailbox $Mailbox auf HideFromGAL=$HideFromGAL" -Type "Info"
        
        Set-Mailbox -Identity $Mailbox -HiddenFromAddressListsEnabled $HideFromGAL -ErrorAction Stop
        
        $visibilityStatus = if ($HideFromGAL) { "ausgeblendet" } else { "sichtbar" }
        Write-Log  "GAL-Sichtbarkeit für Shared Mailbox erfolgreich gesetzt - $visibilityStatus" -Type "Success"
        Log-Action "Shared Mailbox $Mailbox wurde in GAL $visibilityStatus gesetzt"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Setzen der GAL-Sichtbarkeit: $errorMsg" -Type "Error"
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
        Write-Log  "Lösche Shared Mailbox: $Mailbox" -Type "Info"
        
        Remove-Mailbox -Identity $Mailbox -Confirm:$false -ErrorAction Stop
        
        Write-Log  "Shared Mailbox erfolgreich gelöscht" -Type "Success"
        Log-Action "Shared Mailbox $Mailbox wurde gelöscht"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Löschen der Shared Mailbox: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Löschen der Shared Mailbox: $errorMsg"
        return $false
    }
}

# Neue Funktion zum Aktualisieren der Domain-Liste
function Update-DomainList {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log  "Aktualisiere Domain-Liste für die ComboBox" -Type "Info"
        
        # Prüfen, ob die ComboBox existiert
        if ($null -eq $script:cmbSharedMailboxDomain) {
            $script:cmbSharedMailboxDomain = Get-XamlElement -ElementName "cmbSharedMailboxDomain"
            if ($null -eq $script:cmbSharedMailboxDomain) {
                Write-Log  "Domain-ComboBox nicht gefunden" -Type "Warning"
                return $false
            }
        }
        
        # Prüfen, ob eine Verbindung besteht
        if (-not $script:isConnected) {
            Write-Log  "Keine Exchange-Verbindung für Domain-Abfrage" -Type "Warning"
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
        
        Write-Log  "Domain-Liste erfolgreich aktualisiert: $($domains.Count) Domains geladen" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Aktualisieren der Domain-Liste: $errorMsg" -Type "Error"
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
            Write-Log  "Öffne Admin Center Link: $($diagnostic.AdminCenterLink)" -Type "Info"
            
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
            Write-Log  "Kein Admin Center Link für diese Diagnose vorhanden" -Type "Warning"
            if ($null -ne $txtStatus) {
                Update-GuiText -TextElement $txtStatus -Message "Kein Admin Center Link für diese Diagnose vorhanden."
            }
            
            return $false
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Öffnen des Admin Center Links: $errorMsg" -Type "Error"
        
        if ($null -ne $txtStatus) {
            Update-GuiText -TextElement $txtStatus -Message "Fehler beim Öffnen des Admin Center Links: $errorMsg"
        }
        
        Log-Action "Fehler beim Öffnen des Admin Center Links: $errorMsg"
        return $false
    }
}

function New-ResourceAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $true)]
        [string]$ResourceType,
        
        [Parameter(Mandatory = $false)]
        [string]$Capacity,
        
        [Parameter(Mandatory = $false)]
        [string]$Location
    )
    
    try {
        Write-Log  "Erstelle neue Ressource: $Name (Typ: $ResourceType)" -Type "Info"
        
        # Parameter für die Ressourcenerstellung vorbereiten
        $params = @{
            Name = $Name
            DisplayName = $DisplayName
            ResourceCapacity = if ([string]::IsNullOrEmpty($Capacity)) { $null } else { [int]$Capacity }
        }
        
        # Standort hinzufügen, wenn angegeben
        if (-not [string]::IsNullOrEmpty($Location)) {
            $params.Add("Location", $Location)
        }
        
        # Ressource basierend auf Typ erstellen
        if ($ResourceType -eq "Room") {
            $result = New-Mailbox -Room @params -ErrorAction Stop
            $resourceTypeName = "Raumressource"
        } 
        elseif ($ResourceType -eq "Equipment") {
            $result = New-Mailbox -Equipment @params -ErrorAction Stop
            $resourceTypeName = "Ausstattungsressource"
        }
        else {
            throw "Ungültiger Ressourcentyp: $ResourceType. Erlaubte Werte sind 'Room' oder 'Equipment'."
        }
        
        Write-Log  "$resourceTypeName erfolgreich erstellt: $Name" -Type "Success"
        Log-Action "$resourceTypeName erstellt: $Name"
        
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Erstellen der Ressource: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Erstellen der Ressource: $errorMsg"
        throw $_
    }
}

function Get-RoomResourcesAction {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log  "Rufe alle Raumressourcen ab" -Type "Info"
        
        $rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited | 
                 Select-Object DisplayName, Name, PrimarySmtpAddress, ResourceCapacity, Office
        
        Write-Log  "Erfolgreich $($rooms.Count) Raumressourcen abgerufen" -Type "Success"
        Log-Action "Raumressourcen abgerufen: $($rooms.Count) gefunden"
        
        return $rooms
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der Raumressourcen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Raumressourcen: $errorMsg"
        throw $_
    }
}

function Get-EquipmentResourcesAction {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log  "Rufe alle Ausstattungsressourcen ab" -Type "Info"
        
        $equipment = Get-Mailbox -RecipientTypeDetails EquipmentMailbox -ResultSize Unlimited | 
                     Select-Object DisplayName, Name, PrimarySmtpAddress, ResourceCapacity, Office
        
        Write-Log  "Erfolgreich $($equipment.Count) Ausstattungsressourcen abgerufen" -Type "Success"
        Log-Action "Ausstattungsressourcen abgerufen: $($equipment.Count) gefunden"
        
        return $equipment
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der Ausstattungsressourcen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Ausstattungsressourcen: $errorMsg"
        throw $_
    }
}

function Search-ResourcesAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm
    )
    
    try {
        Write-Log  "Suche nach Ressourcen mit Suchbegriff: $SearchTerm" -Type "Info"
        
        # Alle Ressourcen abrufen und filtern
        $resources = Get-Mailbox -RecipientTypeDetails RoomMailbox,EquipmentMailbox -ResultSize Unlimited | 
                     Where-Object { 
                         $_.DisplayName -like "*$SearchTerm*" -or 
                         $_.Name -like "*$SearchTerm*" -or 
                         $_.PrimarySmtpAddress -like "*$SearchTerm*" -or
                         $_.Office -like "*$SearchTerm*"
                     } | 
                     Select-Object DisplayName, Name, PrimarySmtpAddress, RecipientTypeDetails, ResourceCapacity, Office
        
        Write-Log  "Suchergebnis: $($resources.Count) Ressourcen gefunden" -Type "Success"
        Log-Action "Ressourcensuche für '$SearchTerm': $($resources.Count) Ergebnisse"
        
        return $resources
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler bei der Ressourcensuche: $errorMsg" -Type "Error"
        Log-Action "Fehler bei der Ressourcensuche: $errorMsg"
        throw $_
    }
}

function Get-AllResourcesAction {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log  "Rufe alle Ressourcen (Räume und Ausstattung) ab" -Type "Info"
        
        $resources = Get-Mailbox -RecipientTypeDetails RoomMailbox,EquipmentMailbox -ResultSize Unlimited | 
                     Select-Object DisplayName, Name, PrimarySmtpAddress, RecipientTypeDetails, ResourceCapacity, Office
        
        Write-Log  "Erfolgreich $($resources.Count) Ressourcen abgerufen" -Type "Success"
        Log-Action "Alle Ressourcen abgerufen: $($resources.Count) gefunden"
        
        return $resources
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen aller Ressourcen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen aller Ressourcen: $errorMsg"
        throw $_
    }
}

function Remove-ResourceAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    
    try {
        Write-Log  "Lösche Ressource: $Identity" -Type "Info"
        
        # Ressource abrufen, um den Typ zu bestimmen
        $resource = Get-Mailbox -Identity $Identity -ErrorAction Stop
        $resourceType = if ($resource.RecipientTypeDetails -eq "RoomMailbox") { "Raumressource" } else { "Ausstattungsressource" }
        
        # Ressource löschen
        Remove-Mailbox -Identity $Identity -Confirm:$false -ErrorAction Stop
        
        Write-Log  "$resourceType erfolgreich gelöscht: $Identity" -Type "Success"
        Log-Action "$resourceType gelöscht: $Identity"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Löschen der Ressource: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Löschen der Ressource: $errorMsg"
        throw $_
    }
}

function Export-ResourcesAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Resources,
        
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        Write-Log  "Exportiere $($Resources.Count) Ressourcen nach: $FilePath" -Type "Info"
        
        # Ressourcen in CSV-Datei exportieren
        $Resources | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        
        Write-Log  "Ressourcen erfolgreich exportiert" -Type "Success"
        Log-Action "Ressourcen exportiert nach: $FilePath"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Exportieren der Ressourcen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Exportieren der Ressourcen: $errorMsg"
        throw $_
    }
}

function Get-ResourceSettingsAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    
    try {
        Write-Log  "Rufe Ressourceneinstellungen ab für: $Identity" -Type "Info"
        
        # Ressource abrufen
        $resource = Get-Mailbox -Identity $Identity -ErrorAction Stop
        
        # Kalenderverarbeitungseinstellungen abrufen
        $calendarProcessing = Get-CalendarProcessing -Identity $Identity -ErrorAction Stop
        
        # Ressourceneinstellungen zusammenstellen
        $resourceSettings = [PSCustomObject]@{
            Name = $resource.Name
            DisplayName = $resource.DisplayName
            PrimarySmtpAddress = $resource.PrimarySmtpAddress
            ResourceType = $resource.RecipientTypeDetails
            Capacity = $resource.ResourceCapacity
            Location = $resource.ResourceCustom
            AutoAccept = $calendarProcessing.AutomateProcessing -eq "AutoAccept"
            AllowConflicts = $calendarProcessing.AllowConflicts
            BookingWindowInDays = $calendarProcessing.BookingWindowInDays
            MaximumDurationInMinutes = $calendarProcessing.MaximumDurationInMinutes
            AllowRecurringMeetings = $calendarProcessing.AllowRecurringMeetings
            EnforceSchedulingHorizon = $calendarProcessing.EnforceSchedulingHorizon
            ScheduleOnlyDuringWorkHours = $calendarProcessing.ScheduleOnlyDuringWorkHours
            DeleteComments = $calendarProcessing.DeleteComments
            DeleteSubject = $calendarProcessing.DeleteSubject
            RemovePrivateProperty = $calendarProcessing.RemovePrivateProperty
            AddOrganizerToSubject = $calendarProcessing.AddOrganizerToSubject
        }
        
        Write-Log  "Ressourceneinstellungen erfolgreich abgerufen" -Type "Success"
        Log-Action "Ressourceneinstellungen abgerufen für: $Identity"
        
        return $resourceSettings
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Abrufen der Ressourceneinstellungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Ressourceneinstellungen: $errorMsg"
        throw $_
    }
}

function Update-ResourceSettingsAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $false)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $false)]
        [int]$Capacity,
        
        [Parameter(Mandatory = $false)]
        [string]$Location,
        
        [Parameter(Mandatory = $false)]
        [bool]$AutoAccept,
        
        [Parameter(Mandatory = $false)]
        [bool]$AllowConflicts,
        
        [Parameter(Mandatory = $false)]
        [int]$BookingWindowInDays,
        
        [Parameter(Mandatory = $false)]
        [int]$MaximumDurationInMinutes,
        
        [Parameter(Mandatory = $false)]
        [bool]$AllowRecurringMeetings,
        
        [Parameter(Mandatory = $false)]
        [bool]$ScheduleOnlyDuringWorkHours,
        
        [Parameter(Mandatory = $false)]
        [bool]$DeleteComments,
        
        [Parameter(Mandatory = $false)]
        [bool]$DeleteSubject,
        
        [Parameter(Mandatory = $false)]
        [bool]$RemovePrivateProperty,
        
        [Parameter(Mandatory = $false)]
        [bool]$AddOrganizerToSubject
    )
    
    try {
        Write-Log  "Aktualisiere Ressourceneinstellungen für: $Identity" -Type "Info"
        
        # Ressource abrufen, um den Typ zu bestimmen
        $resource = Get-Mailbox -Identity $Identity -ErrorAction Stop
        $resourceType = if ($resource.RecipientTypeDetails -eq "RoomMailbox") { "Raumressource" } else { "Ausstattungsressource" }
        
        # Mailbox-Eigenschaften aktualisieren
        $mailboxParams = @{
            Identity = $Identity
            ErrorAction = "Stop"
        }
        
        if ($PSBoundParameters.ContainsKey('DisplayName')) {
            $mailboxParams.Add("DisplayName", $DisplayName)
        }
        
        if ($PSBoundParameters.ContainsKey('Capacity')) {
            $mailboxParams.Add("ResourceCapacity", $Capacity)
        }
        
        if ($PSBoundParameters.ContainsKey('Location')) {
            $mailboxParams.Add("ResourceCustom", $Location)
        }
        
        if ($mailboxParams.Count -gt 2) {
            Set-Mailbox @mailboxParams
            Write-Log  "Mailbox-Eigenschaften aktualisiert" -Type "Info"
        }
        
        # Kalenderverarbeitungseinstellungen aktualisieren
        $calendarParams = @{
            Identity = $Identity
            ErrorAction = "Stop"
        }
        
        if ($PSBoundParameters.ContainsKey('AutoAccept')) {
            $calendarParams.Add("AutomateProcessing", $(if ($AutoAccept) { "AutoAccept" } else { "AutoUpdate" }))
        }
        
        if ($PSBoundParameters.ContainsKey('AllowConflicts')) {
            $calendarParams.Add("AllowConflicts", $AllowConflicts)
        }
        
        if ($PSBoundParameters.ContainsKey('BookingWindowInDays')) {
            $calendarParams.Add("BookingWindowInDays", $BookingWindowInDays)
        }
        
        if ($PSBoundParameters.ContainsKey('MaximumDurationInMinutes')) {
            $calendarParams.Add("MaximumDurationInMinutes", $MaximumDurationInMinutes)
        }
        
        if ($PSBoundParameters.ContainsKey('AllowRecurringMeetings')) {
            $calendarParams.Add("AllowRecurringMeetings", $AllowRecurringMeetings)
        }
        
        if ($PSBoundParameters.ContainsKey('ScheduleOnlyDuringWorkHours')) {
            $calendarParams.Add("ScheduleOnlyDuringWorkHours", $ScheduleOnlyDuringWorkHours)
        }
        
        if ($PSBoundParameters.ContainsKey('DeleteComments')) {
            $calendarParams.Add("DeleteComments", $DeleteComments)
        }
        
        if ($PSBoundParameters.ContainsKey('DeleteSubject')) {
            $calendarParams.Add("DeleteSubject", $DeleteSubject)
        }
        
        if ($PSBoundParameters.ContainsKey('RemovePrivateProperty')) {
            $calendarParams.Add("RemovePrivateProperty", $RemovePrivateProperty)
        }
        
        if ($PSBoundParameters.ContainsKey('AddOrganizerToSubject')) {
            $calendarParams.Add("AddOrganizerToSubject", $AddOrganizerToSubject)
        }
        
        if ($calendarParams.Count -gt 2) {
            Set-CalendarProcessing @calendarParams
            Write-Log  "Kalenderverarbeitungseinstellungen aktualisiert" -Type "Info"
        }
        
        Write-Log  "$resourceType-Einstellungen erfolgreich aktualisiert: $Identity" -Type "Success"
        Log-Action "$resourceType-Einstellungen aktualisiert: $Identity"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Aktualisieren der Ressourceneinstellungen: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Aktualisieren der Ressourceneinstellungen: $errorMsg"
        throw $_
    }
}

function Show-ResourceSettingsDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    
    try {
        Write-Log  "Öffne Ressourceneinstellungen-Dialog für: $Identity" -Type "Info"
        
        # Ressourceneinstellungen abrufen
        $resourceSettings = Get-ResourceSettingsAction -Identity $Identity
        
        # XAML für den Dialog erstellen
        $xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Ressourceneinstellungen bearbeiten" Height="550" Width="600" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="Ressourceneinstellungen für: $($resourceSettings.DisplayName)" FontWeight="Bold" Margin="0,0,0,10"/>
        
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <StackPanel>
                <GroupBox Header="Allgemeine Einstellungen" Margin="0,5,0,10">
                    <Grid Margin="5">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Anzeigename:" Margin="0,5,10,5" VerticalAlignment="Center"/>
                        <TextBox Grid.Row="0" Grid.Column="1" Name="txtDisplayName" Margin="0,5,0,5"/>
                        
                        <TextBlock Grid.Row="1" Grid.Column="0" Text="Kapazität:" Margin="0,5,10,5" VerticalAlignment="Center"/>
                        <TextBox Grid.Row="1" Grid.Column="1" Name="txtCapacity" Margin="0,5,0,5"/>
                        
                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Standort:" Margin="0,5,10,5" VerticalAlignment="Center"/>
                        <TextBox Grid.Row="2" Grid.Column="1" Name="txtLocation" Margin="0,5,0,5"/>
                    </Grid>
                </GroupBox>
                
                <GroupBox Header="Buchungseinstellungen" Margin="0,5,0,10">
                    <Grid Margin="5">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <CheckBox Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="2" Name="chkAutoAccept" Content="Buchungen automatisch akzeptieren" Margin="0,5,0,5"/>
                        
                        <CheckBox Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2" Name="chkAllowConflicts" Content="Terminüberschneidungen erlauben" Margin="0,5,0,5"/>
                        
                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Buchungszeitraum (Tage):" Margin="0,5,10,5" VerticalAlignment="Center"/>
                        <TextBox Grid.Row="2" Grid.Column="1" Name="txtBookingWindow" Margin="0,5,0,5"/>
                        
                        <TextBlock Grid.Row="3" Grid.Column="0" Text="Maximale Dauer (Minuten):" Margin="0,5,10,5" VerticalAlignment="Center"/>
                        <TextBox Grid.Row="3" Grid.Column="1" Name="txtMaxDuration" Margin="0,5,0,5"/>
                        
                        <CheckBox Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Name="chkAllowRecurring" Content="Serientermine erlauben" Margin="0,5,0,5"/>
                        
                        <CheckBox Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="2" Name="chkWorkHoursOnly" Content="Nur während Arbeitszeiten buchen" Margin="0,5,0,5"/>
                    </Grid>
                </GroupBox>
                
                <GroupBox Header="Terminverarbeitung" Margin="0,5,0,10">
                    <Grid Margin="5">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <CheckBox Grid.Row="0" Grid.Column="0" Name="chkDeleteComments" Content="Kommentare löschen" Margin="0,5,0,5"/>
                        <CheckBox Grid.Row="0" Grid.Column="1" Name="chkDeleteSubject" Content="Betreff löschen" Margin="0,5,0,5"/>
                        
                        <CheckBox Grid.Row="1" Grid.Column="0" Name="chkRemovePrivate" Content="Private Markierung entfernen" Margin="0,5,0,5"/>
                        <CheckBox Grid.Row="1" Grid.Column="1" Name="chkAddOrganizer" Content="Organisator zum Betreff hinzufügen" Margin="0,5,0,5"/>
                    </Grid>
                </GroupBox>
            </StackPanel>
        </ScrollViewer>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Name="btnSave" Content="Speichern" Width="100" Margin="0,0,10,0"/>
            <Button Name="btnCancel" Content="Abbrechen" Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@
        
        # XAML laden und Fenster erstellen
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
        $window = [System.Windows.Markup.XamlReader]::Load($reader)
        
        # UI-Elemente referenzieren
        $txtDisplayName = $window.FindName("txtDisplayName")
        $txtCapacity = $window.FindName("txtCapacity")
        $txtLocation = $window.FindName("txtLocation")
        $chkAutoAccept = $window.FindName("chkAutoAccept")
        $chkAllowConflicts = $window.FindName("chkAllowConflicts")
        $txtBookingWindow = $window.FindName("txtBookingWindow")
        $txtMaxDuration = $window.FindName("txtMaxDuration")
        $chkAllowRecurring = $window.FindName("chkAllowRecurring")
        $chkWorkHoursOnly = $window.FindName("chkWorkHoursOnly")
        $chkDeleteComments = $window.FindName("chkDeleteComments")
        $chkDeleteSubject = $window.FindName("chkDeleteSubject")
        $chkRemovePrivate = $window.FindName("chkRemovePrivate")
        $chkAddOrganizer = $window.FindName("chkAddOrganizer")
        $btnSave = $window.FindName("btnSave")
        $btnCancel = $window.FindName("btnCancel")
        
        # Werte in die UI-Elemente laden
        $txtDisplayName.Text = $resourceSettings.DisplayName
        $txtCapacity.Text = $resourceSettings.Capacity
        $txtLocation.Text = $resourceSettings.Location
        $chkAutoAccept.IsChecked = $resourceSettings.AutoAccept
        $chkAllowConflicts.IsChecked = $resourceSettings.AllowConflicts
        $txtBookingWindow.Text = $resourceSettings.BookingWindowInDays
        $txtMaxDuration.Text = $resourceSettings.MaximumDurationInMinutes
        $chkAllowRecurring.IsChecked = $resourceSettings.AllowRecurringMeetings
        $chkWorkHoursOnly.IsChecked = $resourceSettings.ScheduleOnlyDuringWorkHours
        $chkDeleteComments.IsChecked = $resourceSettings.DeleteComments
        $chkDeleteSubject.IsChecked = $resourceSettings.DeleteSubject
        $chkRemovePrivate.IsChecked = $resourceSettings.RemovePrivateProperty
        $chkAddOrganizer.IsChecked = $resourceSettings.AddOrganizerToSubject
        
        # Event-Handler für Speichern-Button
        $btnSave.Add_Click({
            try {
                # Parameter sammeln
                $params = @{
                    Identity = $Identity
                    DisplayName = $txtDisplayName.Text
                    Capacity = [int]::Parse($txtCapacity.Text)
                    Location = $txtLocation.Text
                    AutoAccept = $chkAutoAccept.IsChecked
                    AllowConflicts = $chkAllowConflicts.IsChecked
                    BookingWindowInDays = [int]::Parse($txtBookingWindow.Text)
                    MaximumDurationInMinutes = [int]::Parse($txtMaxDuration.Text)
                    AllowRecurringMeetings = $chkAllowRecurring.IsChecked
                    ScheduleOnlyDuringWorkHours = $chkWorkHoursOnly.IsChecked
                    DeleteComments = $chkDeleteComments.IsChecked
                    DeleteSubject = $chkDeleteSubject.IsChecked
                    RemovePrivateProperty = $chkRemovePrivate.IsChecked
                    AddOrganizerToSubject = $chkAddOrganizer.IsChecked
                }
                
                # Ressourceneinstellungen aktualisieren
                $result = Update-ResourceSettingsAction @params
                
                if ($result) {
                    [System.Windows.MessageBox]::Show("Die Ressourceneinstellungen wurden erfolgreich aktualisiert.", 
                        "Erfolg", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    $window.DialogResult = $true
                    $window.Close()
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                [System.Windows.MessageBox]::Show("Fehler beim Aktualisieren der Ressourceneinstellungen: $errorMsg", 
                    "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })
        
        # Event-Handler für Abbrechen-Button
        $btnCancel.Add_Click({
            $window.DialogResult = $false
            $window.Close()
        })
        
        # Dialog anzeigen
        $result = $window.ShowDialog()
        
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Log-Action "Fehler beim Öffnen des Ressourceneinstellungen-Dialogs: $errorMsg"
        throw $_
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

function Initialize-ContactsTab {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log  "Initialisiere Kontakte-Tab" -Type "Info"
        
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
        
        Write-Log  "Event-Handler für btnShowMailUsers registrieren" -Type "Info"

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
        
        return $true
            }
            catch {
                $errorMsg = $_.Exception.Message
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
        Write-Log  "Lade XAML von: $XamlFilePath" -Type "Info"
        
        # Überprüfen, ob die XAML-Datei existiert
        if (-not (Test-Path -Path $XamlFilePath)) {
            throw "XAML-Datei nicht gefunden: $XamlFilePath"
        }
        
        # XAML-Datei laden
        [xml]$xamlContent = Get-Content -Path $XamlFilePath -Encoding UTF8
        
        # Parse XAML
        $reader = New-Object System.Xml.XmlNodeReader $xamlContent
        
        try {
            $window = [System.Windows.Markup.XamlReader]::Load($reader)
        }
        catch {
            Write-Log "Fehler beim Laden des XAML: $($_.Exception.Message)" -Type "Error"
            throw
        }
        
        if ($null -eq $window) {
            throw "XamlReader.Load gab ein null-Window-Objekt zurück"
        }
        
        # TabControl überprüfen
        $tabControl = $window.FindName("tabContent")
        if ($null -eq $tabControl) {
            throw "TabControl 'tabContent' nicht im XAML gefunden"
        }
        
        # WICHTIG: Stelle sicher, dass Items-Sammlung initialisiert ist
        if ($null -eq $tabControl.Items) {
            Write-Log "Warnung: TabControl.Items ist null - initialisiere" -Type "Warning"
            # Stellen Sie sicher, dass TabControl korrekt initialisiert ist
            $tabControl.UpdateLayout()
        }
        
        # Überprüfe jeden Tab
        if ($null -ne $tabControl.Items) {
            Write-Log "TabControl hat $($tabControl.Items.Count) Items" -Type "Info"
            foreach ($item in $tabControl.Items) {
                Write-Log "Tab gefunden: Name=$($item.Name), Header=$($item.Header), Visibility=$($item.Visibility)" -Type "Info"
            }
        } else {
            Write-Log "TabControl.Items ist immer noch null!" -Type "Error"
        }
        
        return $window
    }
    
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler in Load-XAML: $errorMsg" -Type "Error"
        throw
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
}

# Prüfen, ob XAML-Datei gefunden wurde
if (-not (Test-Path -Path $script:xamlFilePath)) {
    Write-Log "KRITISCHER FEHLER: XAML-Datei nicht gefunden an beiden Standardpfaden!"  
    Write-Log "Gesucht wurde in: $PSScriptRoot und $PSScriptRoot\assets"  
    try {
        $tempXamlPath = [System.IO.Path]::GetTempFileName() + ".xaml"
        Set-Content -Path $tempXamlPath -Value $minimalXaml -Encoding UTF8
        
        $script:xamlFilePath = $tempXamlPath
    }
    catch {
        Write-Log "Konnte keine Notfall-GUI erstellen. Das Programm wird beendet."  
        exit
    }
}

try {
    # GUI aus externer XAML-Datei laden
    $script:Form = Load-XAML -XamlFilePath $script:xamlFilePath
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
            
            if ($null -eq $script:Form) {
                throw "Form-Objekt ist nicht initialisiert"
            }
            
            
            $element = $script:Form.FindName($ElementName)
            
            if ($null -eq $element -and $Required.IsPresent) {
                throw "Erforderliches Element nicht gefunden: $ElementName"
            }
            elseif ($null -eq $element) {
            }
            else {
            }
            return $element
        }
        catch {
            $errorMsg = $_.Exception.Message
            if ($Required.IsPresent) { throw }
            return $null
        }
    }
    
    # Hauptelemente referenzieren
    $script:btnConnect          = Get-XamlElement -ElementName "btnConnect" -Required
    $script:tabContent          = Get-XamlElement -ElementName "tabContent" -Required
    $script:tabEXOSettings      = Get-XamlElement -ElementName "tabEXOSettings"
    $script:tabRegion           = Get-XamlElement -ElementName "tabRegion"
    $script:tabCalendar         = Get-XamlElement -ElementName "tabCalendar"
    $script:tabMailbox          = Get-XamlElement -ElementName "tabMailbox"
    $script:tabResources        = Get-XamlElement -ElementName "tabResources"
    $script:tabContacts         = Get-XamlElement -ElementName "tabContacts"
    $script:tabMailboxAudit     = Get-XamlElement -ElementName "tabMailboxAudit"
    $script:tabTroubleshooting  = Get-XamlElement -ElementName "tabTroubleshooting"
    $script:tabGroups           = Get-XamlElement -ElementName "tabGroups" 
    $script:tabSharedMailbox    = Get-XamlElement -ElementName "tabSharedMailbox"
    $script:tabReports          = Get-XamlElement -ElementName "tabReports"
    $script:txtStatus           = Get-XamlElement -ElementName "txtStatus" -Required
    $script:txtVersion          = Get-XamlElement -ElementName "txtVersion"
    $script:txtConnectionStatus = Get-XamlElement -ElementName "txtConnectionStatus" -Required
    
    # Referenzierung der Navigationselemente
    $script:btnNavEXOSettings     = Get-XamlElement -ElementName "btnNavEXOSettings"
    $script:btnNavRegion          = Get-XamlElement -ElementName "btnNavRegion"
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

    # Button-Handler
    $script:btnConnect.Add_Click({ Connect-ExchangeOnline })
    $script:btnClose.Add_Click({ $script:Form.Close() })

    # Hier den Tab-Handler einfügen:
    $script:tabContent.Add_SelectionChanged({
        param($sender, $e)
        $selectedTab = $sender.SelectedItem
        if ($null -ne $selectedTab) {
            Write-LogEntry "Tab gewechselt zu: $($selectedTab.Header)"
        }
    })

    # Navigation Button Handler - Fehlerbehandlung hinzugefügt
    $script:btnNavRegion.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabRegion) {
            $script:tabContent.SelectedItem = $script:tabRegion
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })

    # Navigation Button Handler - Fehlerbehandlung hinzugefügt
    $script:btnNavEXOSettings.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabEXOSettings) {
            $script:tabContent.SelectedItem = $script:tabEXOSettings
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })

    $script:btnNavCalendar.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabCalendar) {
            $script:tabContent.SelectedItem = $script:tabCalendar
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })

    $script:btnNavMailbox.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabMailbox) {
            $script:tabContent.SelectedItem = $script:tabMailbox
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })

    $script:btnNavGroups.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabGroups) {
            $script:tabContent.SelectedItem = $script:tabGroups
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })

    $script:btnNavSharedMailbox.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabSharedMailbox) {
            $script:tabContent.SelectedItem = $script:tabSharedMailbox
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })

    $script:btnNavResources.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabResources) {
            $script:tabContent.SelectedItem = $script:tabResources
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })

    $script:btnNavContacts.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabContacts) {
            $script:tabContent.SelectedItem = $script:tabContacts
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })

    $script:btnNavAudit.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabMailboxAudit) {
            $script:tabContent.SelectedItem = $script:tabMailboxAudit
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })

    $script:btnNavReports.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabReports) {
            $script:tabContent.SelectedItem = $script:tabReports
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })

    $script:btnNavTroubleshooting.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabTroubleshooting) {
            $script:tabContent.SelectedItem = $script:tabTroubleshooting
        } else {
            Write-Log "Fehler: Tab oder TabControl ist null" -Type "Error"
        }
    })
    
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
            Write-Log  "Control nicht gefunden: $ControlName" -Type "Warning"
            return $false
        }
        
        try {
            # Event-Handler hinzufügen
            $event = "Add_$EventName"
            $Control.$event($Handler)
            Write-Log  "Event-Handler für $ControlName.$EventName registriert" -Type "Info"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-Log  "Fehler beim Registrieren des Event-Handlers für $ControlName - $errorMsg" -Type "Error"
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

    Write-Log "Beginne Initialisierung: EXO Settings Tab" -Type "Info"
    [bool]$success = $true

    try {
        # Prüfe Exchange Online Verbindung NICHT hier - das passiert beim Laden der Daten
        Write-Log "EXOSettingsTab: Prüfe Verbindung nicht bei Initialisierung, nur Event Handler setzen." -Type "Debug"

        # Speichere Referenz auf den Tab selbst für spätere Verwendung
        $script:tabEXOSettings = Get-XamlElement -ElementName "tabEXOSettings"
        if ($null -eq $script:tabEXOSettings) {
            Write-Log "EXOSettingsTab: Tab-Element nicht gefunden!" -Type "Warning"
        }

        # Textfeld für Status finden
        if ($null -eq $script:txtStatus) {
            $script:txtStatus = Get-XamlElement -ElementName "txtStatus"
        }

        # Event-Handler für Help-Link
        $helpLinkEXOSettings = Get-XamlElement -ElementName "helpLinkEXOSettings"
        if ($null -ne $helpLinkEXOSettings) {
            $helpLinkEXOSettings.Add_MouseLeftButtonDown({
                try { Start-Process "https://learn.microsoft.com/de-de/powershell/module/exchange/set-organizationconfig?view=exchange-ps" } catch { Write-Log "Fehler beim Öffnen des HelpLinks: $($_.Exception.Message)" -Type "Error" }
            })
            Write-Log "EXOSettingsTab: HelpLink Handler registriert." -Type "Debug"
        } else { Write-Log "EXOSettingsTab: helpLinkEXOSettings nicht gefunden." -Type "Warning"; $success = $false }

        # Event-Handler für "Aktuelle Einstellungen laden" Button
        $btnGetOrganizationConfig = Get-XamlElement -ElementName "btnGetOrganizationConfig"
        if ($null -ne $btnGetOrganizationConfig) {
            $btnGetOrganizationConfig.Add_Click({
                Write-Log "Button 'btnGetOrganizationConfig' geklickt." -Type "Info"
                # Die Funktion Get-CurrentOrganizationConfig prüft die Verbindung intern
                Get-CurrentOrganizationConfig # Keine explizite Prüfung hier nötig
            })
            Write-Log "EXOSettingsTab: btnGetOrganizationConfig Handler registriert." -Type "Debug"
        } else { Write-Log "EXOSettingsTab: btnGetOrganizationConfig nicht gefunden." -Type "Warning"; $success = $false }

        # Event-Handler für "Einstellungen speichern" Button
        $btnSetOrganizationConfig = Get-XamlElement -ElementName "btnSetOrganizationConfig"
        if ($null -ne $btnSetOrganizationConfig) {
            $btnSetOrganizationConfig.Add_Click({
                Write-Log "Button 'btnSetOrganizationConfig' geklickt." -Type "Info"
                # Die Funktion Set-CustomOrganizationConfig prüft die Verbindung intern
                Set-CustomOrganizationConfig # Keine explizite Prüfung hier nötig
            })
            Write-Log "EXOSettingsTab: btnSetOrganizationConfig Handler registriert." -Type "Debug"
        } else { Write-Log "EXOSettingsTab: btnSetOrganizationConfig nicht gefunden." -Type "Warning"; $success = $false }

        # Event-Handler für "Konfiguration exportieren" Button
        $btnExportOrgConfig = Get-XamlElement -ElementName "btnExportOrgConfig"
        if ($null -ne $btnExportOrgConfig) {
            $btnExportOrgConfig.Add_Click({
                Write-Log "Button 'btnExportOrgConfig' geklickt." -Type "Info"
                # Die Funktion Export-OrganizationConfig prüft die Verbindung intern
                Export-OrganizationConfig # Keine explizite Prüfung hier nötig
            })
            Write-Log "EXOSettingsTab: btnExportOrgConfig Handler registriert." -Type "Debug"
        } else { Write-Log "EXOSettingsTab: btnExportOrgConfig nicht gefunden." -Type "Warning"; $success = $false }

        # Stelle sicher, dass alle UI-Elemente sichtbar sind
        foreach ($elementName in $script:knownUIElements) {
            $element = Get-XamlElement -ElementName $elementName
            if ($null -ne $element) {
                # Stelle sicher, dass das Element sichtbar ist
                $element.Visibility = [System.Windows.Visibility]::Visible
                Write-Log "EXOSettingsTab: Element '$elementName' auf sichtbar gesetzt." -Type "Debug"
            } else {
                Write-Log "EXOSettingsTab: Element '$elementName' nicht gefunden." -Type "Warning"
                $success = $false
            }
        }

        # Initialize all UI controls for the OrganizationConfig tab
        Write-Log "EXOSettingsTab: Rufe Initialize-OrganizationConfigControls auf..." -Type "Debug"
        $controlsInitResult = Initialize-OrganizationConfigControls
        Write-Log "EXOSettingsTab: Initialize-OrganizationConfigControls Ergebnis: $controlsInitResult" -Type "Debug"
        if (-not $controlsInitResult) { $success = $false } # Wenn Controls nicht initialisiert werden können, ist der Tab fehlerhaft

        # Stelle sicher, dass der Tab selbst sichtbar ist
        if ($null -ne $script:tabEXOSettings) {
            $script:tabEXOSettings.Visibility = [System.Windows.Visibility]::Visible
            Write-Log "EXOSettingsTab: Tab auf sichtbar gesetzt." -Type "Debug"
        }

        # Stelle sicher, dass alle TabItems innerhalb des TabControls sichtbar sind
        $tabControl = Get-XamlElement -ElementName "mainTabControl"
        if ($null -ne $tabControl) {
            foreach ($tabItem in $tabControl.Items) {
                if ($tabItem.Name -eq "tabEXOSettings") {
                    $tabItem.Visibility = [System.Windows.Visibility]::Visible
                    Write-Log "EXOSettingsTab: TabItem im TabControl auf sichtbar gesetzt." -Type "Debug"
                }
            }
        }

        # Lade initial die aktuellen Einstellungen - ABER NUR WENN SCHON VERBUNDEN
        # Das Laden beim Tab-Wechsel ist oft nicht gewünscht, besser per Buttonklick.
        # Wir rufen es hier nur auf, wenn die Verbindung *bereits* besteht.
        $connectionExists = Confirm-ExchangeConnection
        if ($connectionExists) {
            Write-Log "EXOSettingsTab: Verbindung besteht, lade initiale OrgConfig..." -Type "Info"
            Get-CurrentOrganizationConfig
        } else {
            Write-Log "EXOSettingsTab: Keine Verbindung, lade initiale OrgConfig nicht." -Type "Info"
            # Platzhalter anzeigen
            $txtOrgConfig = Get-XamlElement -ElementName "txtOrganizationConfig"
            if ($null -ne $txtOrgConfig) { 
                $txtOrgConfig.Text = "Bitte mit Exchange Online verbinden und auf 'Aktuelle Einstellungen laden' klicken."
                $txtOrgConfig.Visibility = [System.Windows.Visibility]::Visible
            }
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Initialisieren des EXO Settings Tab: $errorMsg" -Type "Error"
        Write-Log $_.Exception.StackTrace -Type "Error"
        $success = $false
    }

    $messageType = if ($success) { "Success" } else { "Error" }
    Write-Log "Abschluss Initialisierung: EXO Settings Tab (Erfolg: $success)" -Type $messageType
    return $success
}
#endregion EXOSettings Tab Initialization

function Initialize-OrganizationConfigControls {
    [CmdletBinding()]
    param()
    
    try {
        # Initialisiere die Einstellungs-Hashtable, falls noch nicht vorhanden
        if ($null -eq $script:organizationConfigSettings) {
            $script:organizationConfigSettings = @{}
        }
        
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
                        Write-Log "Checkbox $checkBoxName wurde auf $($checkBox.IsChecked) gesetzt" -Type "Info"
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
                            "cmbActivityBasedAuthenticationTimeoutInterval" {
                                # Format: "01:00:00 (1h)" zu "01:00:00"
                                $value = ($selectedContent -split ' ')[0]
                                $script:organizationConfigSettings["ActivityBasedAuthenticationTimeoutInterval"] = $value
                            }
                            "cmbLargeAudienceThreshold" {
                                # Konvertieren zu Integer
                                $value = [int]$selectedContent
                                $script:organizationConfigSettings["MailTipsLargeAudienceThreshold"] = $value
                            }
                            "cmbInformationBarrierMode" {
                                $script:organizationConfigSettings["InformationBarrierMode"] = $selectedContent
                            }
                            "cmbEwsAppAccessPolicy" {
                                $script:organizationConfigSettings["EwsApplicationAccessPolicy"] = $selectedContent
                            }
                            "cmbOfficeFeatures" {
                                $script:organizationConfigSettings["OfficeFeatures"] = $selectedContent
                            }
                            "cmbSearchQueryLanguage" {
                                $script:organizationConfigSettings["SearchQueryLanguage"] = $selectedContent
                            }
                            default {
                                # Generische Behandlung für andere ComboBoxen
                                if ($comboBoxName -like "cmb*" -and $comboBoxName.Length -gt 3) {
                                    $propertyName = $comboBoxName.Substring(3)
                                    $script:organizationConfigSettings[$propertyName] = $selectedContent
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
                            }
                        }
                        "txtPowerShellMaxCmdletQueueDepth" {
                            if ([int]::TryParse($textBox.Text, [ref]$null)) {
                                $script:organizationConfigSettings["PowerShellMaxCmdletQueueDepth"] = [int]$textBox.Text
                            }
                        }
                        "txtPowerShellMaxCmdletsExecutionDuration" {
                            if ([int]::TryParse($textBox.Text, [ref]$null)) {
                                $script:organizationConfigSettings["PowerShellMaxCmdletsExecutionDuration"] = [int]$textBox.Text
                            }
                        }
                        "txtDefaultAuthPolicy" {
                            $script:organizationConfigSettings["DefaultAuthenticationPolicy"] = $textBox.Text
                        }
                        "txtHierAddressBookRoot" {
                            $script:organizationConfigSettings["HierarchicalAddressBookRoot"] = $textBox.Text
                        }
                        "txtPreferredInternetCodePageForShiftJis" {
                            if ([int]::TryParse($textBox.Text, [ref]$null)) {
                                $script:organizationConfigSettings["PreferredInternetCodePageForShiftJis"] = [int]$textBox.Text
                            }
                        }
                        default {
                            # Generische Behandlung für andere TextBoxen
                            if ($textBoxName -like "txt*" -and $textBoxName.Length -gt 3) {
                                $propertyName = $textBoxName.Substring(3)
                                $script:organizationConfigSettings[$propertyName] = $textBox.Text
                            }
                        }
                    }
                }
            }
        }
        
        # Überprüfe, ob der Tab korrekt geladen wurde
        $tabEXOSettings = Get-XamlElement -ElementName "tabEXOSettings"
        if ($null -eq $tabEXOSettings) {
            Write-Log "TabItem 'tabEXOSettings' nicht gefunden" -Type "Error"
            throw "TabItem 'tabEXOSettings' nicht gefunden"
        }
        
        # Überprüfe, ob das TabControl für die Organisationseinstellungen existiert
        $tabOrgSettings = Get-XamlElement -ElementName "tabOrgSettings"
        if ($null -eq $tabOrgSettings) {
            Write-Log "TabControl 'tabOrgSettings' nicht gefunden" -Type "Error"
            throw "TabControl 'tabOrgSettings' nicht gefunden"
        }
        
        # Stelle sicher, dass der Tab sichtbar ist
        $tabEXOSettings.Visibility = [System.Windows.Visibility]::Visible
        
        # Statistiken für die Registrierung
        $registeredControls = @{
            "CheckBox" = 0
            "ComboBox" = 0
            "TextBox" = 0
        }
        
        # Definiere bekannte UI-Elemente basierend auf der XAML-Struktur
        $script:knownUIElements = @(
            # Benutzereinstellungen
            "chkActivityBasedAuthenticationTimeoutEnabled",
            "cmbActivityBasedAuthenticationTimeoutInterval",
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
            
            # Administration & Sicherheit
            "chkAdditionalStorageProvidersBlocked",
            "chkAuditDisabled",
            "chkAutodiscoverPartialDirSync",
            "chkAutoEnableArchiveMailbox",
            "chkAutoExpandingArchive",
            "chkCalendarVersionStoreEnabled",
            "chkCASMailboxHasPermissionsIncludingSubfolders",
            "chkComplianceEnabled",
            
            # Erweitert
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
            
            # Buttons
            "btnGetOrganizationConfig",
            "btnSetOrganizationConfig",
            "btnExportOrgConfig"
        )
        
        # Registriere Event-Handler für alle bekannten Elemente
        foreach ($elementName in $script:knownUIElements) {
            $element = Get-XamlElement -ElementName $elementName
            
            if ($null -ne $element) {
                # Bestimme den Typ des Elements und registriere den entsprechenden Handler
                if ($element -is [System.Windows.Controls.CheckBox]) {
                    Register-EventHandler -Control $element -Handler $controlHandlers["CheckBox"]["Handler"] -ControlName $elementName -EventName $controlHandlers["CheckBox"]["EventName"]
                    $registeredControls["CheckBox"]++
                }
                elseif ($element -is [System.Windows.Controls.ComboBox]) {
                    Register-EventHandler -Control $element -Handler $controlHandlers["ComboBox"]["Handler"] -ControlName $elementName -EventName $controlHandlers["ComboBox"]["EventName"]
                    $registeredControls["ComboBox"]++
                }
                elseif ($element -is [System.Windows.Controls.TextBox]) {
                    Register-EventHandler -Control $element -Handler $controlHandlers["TextBox"]["Handler"] -ControlName $elementName -EventName $controlHandlers["TextBox"]["EventName"]
                    $registeredControls["TextBox"]++
                }
            } else {
                Write-Log "Element '$elementName' nicht gefunden" -Type "Warning"
            }
        }
        
        # Registriere Button-Handler
        $btnGetOrganizationConfig = Get-XamlElement -ElementName "btnGetOrganizationConfig"
        if ($null -ne $btnGetOrganizationConfig) {
            Register-EventHandler -Control $btnGetOrganizationConfig -Handler {
                param($sender, $e)
                Get-CurrentOrganizationConfig
            } -ControlName "btnGetOrganizationConfig" -EventName "Click"
        } else {
            Write-Log "Button 'btnGetOrganizationConfig' nicht gefunden" -Type "Warning"
        }
        
        $btnSetOrganizationConfig = Get-XamlElement -ElementName "btnSetOrganizationConfig"
        if ($null -ne $btnSetOrganizationConfig) {
            Register-EventHandler -Control $btnSetOrganizationConfig -Handler {
                param($sender, $e)
                Update-OrganizationConfig
            } -ControlName "btnSetOrganizationConfig" -EventName "Click"
        } else {
            Write-Log "Button 'btnSetOrganizationConfig' nicht gefunden" -Type "Warning"
        }
        
        $btnExportOrgConfig = Get-XamlElement -ElementName "btnExportOrgConfig"
        if ($null -ne $btnExportOrgConfig) {
            Register-EventHandler -Control $btnExportOrgConfig -Handler {
                param($sender, $e)
                Export-OrganizationConfig
            } -ControlName "btnExportOrgConfig" -EventName "Click"
        } else {
            Write-Log "Button 'btnExportOrgConfig' nicht gefunden" -Type "Warning"
        }
        
        # Hilfe-Link registrieren
        $helpLinkEXOSettings = Get-XamlElement -ElementName "helpLinkEXOSettings"
        if ($null -ne $helpLinkEXOSettings) {
            Register-EventHandler -Control $helpLinkEXOSettings -Handler {
                param($sender, $e)
                Start-Process "https://learn.microsoft.com/de-de/exchange/exchange-online-organization-settings"
            } -ControlName "helpLinkEXOSettings" -EventName "MouseLeftButtonDown"
        } else {
            Write-Log "Hilfe-Link 'helpLinkEXOSettings' nicht gefunden" -Type "Warning"
        }
        
        Write-Log "Registrierte Controls: CheckBoxes: $($registeredControls['CheckBox']), ComboBoxes: $($registeredControls['ComboBox']), TextBoxes: $($registeredControls['TextBox'])" -Type "Info"
        
        # Überprüfe, ob die Ergebnistextbox existiert
        $txtOrganizationConfig = Get-XamlElement -ElementName "txtOrganizationConfig"
        if ($null -eq $txtOrganizationConfig) {
            Write-Log "TextBox 'txtOrganizationConfig' nicht gefunden" -Type "Warning"
        } else {
            $txtOrganizationConfig.Text = "Bereit zum Laden der Organisationseinstellungen. Bitte auf 'Aktuelle Einstellungen laden' klicken."
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler in Initialize-OrganizationConfigControls: $errorMsg" -Type "Error"
        Write-Log $_.Exception.StackTrace -Type "Error"
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
            Show-MessageBox -Message "Bitte verbinden Sie sich zuerst mit Exchange Online, um die Einstellungen zu laden." -Title "Nicht verbunden" -Type "Warning"
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Text = "Nicht verbunden. Einstellungen können nicht geladen werden."
            }
            # Optional: UI-Elemente leeren oder deaktivieren
            $txtOrganizationConfig = Get-XamlElement -ElementName "txtOrganizationConfig"
            if ($null -ne $txtOrganizationConfig) { $txtOrganizationConfig.Text = "Nicht mit Exchange Online verbunden." }
            return
        }

        Write-Log "Rufe aktuelle Organisationseinstellungen ab..." -Type "Info"
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Lade Organisationseinstellungen..."
        }

        # Organisationseinstellungen abrufen
        $script:currentOrganizationConfig = Get-OrganizationConfig -ErrorAction Stop
        if ($null -eq $script:currentOrganizationConfig) {
                throw "Get-OrganizationConfig lieferte keine Daten zurück."
        }
            Write-Log "Organisationseinstellungen erfolgreich abgerufen." -Type "Info"

        $configProperties = $script:currentOrganizationConfig | Get-Member -MemberType Properties |
                            Where-Object { $_.Name -notlike "__*" } |
                            Select-Object -ExpandProperty Name

        # Aktuelle Einstellungs-Hashtable leeren und neu befüllen
        $script:organizationConfigSettings = @{}
            Write-Log "Beginne mit dem Aktualisieren der UI-Elemente..." -Type "Debug"

        # UI-Elemente aktualisieren (nur vorhandene Controls)
        foreach ($elementName in $script:knownUIElements) {
                # Search within the specific tab first, fallback to form
                $element = $null
                if ($null -ne $script:tabEXOSettings) {
                    $element = $script:tabEXOSettings.FindName($elementName) # Suche zuerst im Tab
                }
                if ($null -eq $element) {
                    # Fallback to searching the main form if not found in the tab
                    $element = Get-XamlElement -ElementName $elementName # Suche im Hauptfenster als Fallback
                }

                if ($null -ne $element) {
                    # Elementtyp bestimmen
                    $elementType = if ($element -is [System.Windows.Controls.CheckBox]) { "CheckBox" }
                                  elseif ($element -is [System.Windows.Controls.ComboBox]) { "ComboBox" }
                                  elseif ($element -is [System.Windows.Controls.TextBox]) { "TextBox" }
                                  else { "Unknown" }

                    # Passende Eigenschaft basierend auf Element-Namen finden und Wert setzen
                    try {
                        switch ($elementType) {
                            "CheckBox" {
                                if ($elementName -like "chk*") {
                                    $propertyName = $elementName.Substring(3)
                                    if ($configProperties -contains $propertyName) {
                                        $value = $script:currentOrganizationConfig.$propertyName
                                        # Sicherstellen, dass der Wert ein Boolean ist oder konvertiert werden kann
                                        $boolValue = $false
                                        if ($null -ne $value) {
                                            if ($value -is [bool]) {
                                                $boolValue = $value
                                            } elseif ($value -is [string] -and ($value -eq 'True' -or $value -eq 'False')) {
                                                $boolValue = [bool]::Parse($value)
                                            }
                                            # Weitere Konvertierungen könnten hier nötig sein
                                        }
                                        $element.IsChecked = $boolValue
                                        $script:organizationConfigSettings[$propertyName] = $boolValue
                                        Write-Log "CheckBox '$elementName' gesetzt auf '$boolValue' (Eigenschaft: $propertyName)" -Type "Info"
                                    } else { Write-Log "Warnung: Eigenschaft '$propertyName' nicht in OrgConfig gefunden für CheckBox '$elementName'." -Type "Warning" }
                                }
                            }
                            "ComboBox" {
                                # Spezifische Logik für jede ComboBox
                                $handled = $false
                                switch ($elementName) {
                                    "chkActivityBasedAuthenticationTimeoutInterval" { # Korrekter XAML-Name beibehalten
                                        $propertyName = "ActivityBasedAuthenticationTimeoutInterval"
                                        if ($configProperties -contains $propertyName -and $null -ne $script:currentOrganizationConfig.$propertyName) {
                                            $timeoutValue = $script:currentOrganizationConfig.$propertyName.ToString() # z.B. "06:00:00"
                                            $matchFound = $false
                                            foreach ($item in $element.Items) {
                                                if ($item.Content.ToString().StartsWith($timeoutValue)) {
                                                    $element.SelectedItem = $item
                                                    $script:organizationConfigSettings[$propertyName] = $timeoutValue
                                                    $matchFound = $true; break
                                                }
                                            }
                                            if (-not $matchFound -and $element.Items.Count -gt 0) { $element.SelectedIndex = 1 } # Default 6h
                                                Write-Log "ComboBox '$elementName' gesetzt (Wert: $timeoutValue, Match: $matchFound)" -Type "Info"
                                                $handled = $true
                                        }
                                    }
                                    "cmbLargeAudienceThreshold" {
                                        $propertyName = "MailTipsLargeAudienceThreshold"
                                        if ($configProperties -contains $propertyName -and $null -ne $script:currentOrganizationConfig.$propertyName) {
                                            $value = $script:currentOrganizationConfig.$propertyName.ToString()
                                            $matchFound = $false
                                            foreach ($item in $element.Items) { if ($item.Content.ToString() -eq $value) { $element.SelectedItem = $item; $matchFound = $true; break } }
                                            if (-not $matchFound -and $element.Items.Count -gt 0) { $element.SelectedIndex = 1 } # Default 50
                                            $script:organizationConfigSettings[$propertyName] = [int]$element.SelectedItem.Content
                                                Write-Log "ComboBox '$elementName' gesetzt (Wert: $value, Match: $matchFound)" -Type "Info"
                                                $handled = $true
                                        }
                                    }
                                        # Füge hier die restlichen ComboBoxen hinzu...
                                        "cmbInformationBarrierMode" { $propertyName = "InformationBarrierMode"; $handled = $true }
                                        "cmbEwsAppAccessPolicy" { $propertyName = "EwsApplicationAccessPolicy"; $handled = $true }
                                        "cmbOfficeFeatures" { $propertyName = "OfficeFeatures"; $handled = $true }
                                        "cmbSearchQueryLanguage" { $propertyName = "SearchQueryLanguage"; $handled = $true }
                                }

                                # Generische Behandlung für die anderen ComboBoxen
                                if (-not $handled -and $elementName -like "cmb*" -and $elementName.Length -gt 3) {
                                    $propertyName = $elementName.Substring(3)
                                    $handled = $true
                                }

                                if ($handled -and $configProperties -contains $propertyName -and $null -ne $script:currentOrganizationConfig.$propertyName) {
                                    $value = $script:currentOrganizationConfig.$propertyName.ToString()
                                    $matchFound = $false
                                    foreach ($item in $element.Items) { 
                                        if ($item.Content.ToString() -eq $value) { 
                                            $element.SelectedItem = $item
                                            $matchFound = $true
                                            break 
                                        } 
                                    }
                                    if (-not $matchFound -and $element.Items.Count -gt 0) { $element.SelectedIndex = 0 } # Default erstes Element
                                    if ($null -ne $element.SelectedItem) {
                                        $script:organizationConfigSettings[$propertyName] = $element.SelectedItem.Content.ToString()
                                        Write-Log "ComboBox '$elementName' gesetzt (Wert: $value, Match: $matchFound)" -Type "Info"
                                    } else {
                                        Write-Log "Warnung: Kein Element in ComboBox '$elementName' ausgewählt." -Type "Warning"
                                    }
                                } elseif ($handled) { Write-Log "Warnung: Eigenschaft '$propertyName' nicht in OrgConfig gefunden/null für ComboBox '$elementName'." -Type "Warning" }
                            }
                            "TextBox" {
                                # Spezifische Logik für jede TextBox
                                $handled = $false
                                switch ($elementName) {
                                        "txtPowerShellMaxConcurrency" { $propertyName = "PowerShellMaxConcurrency"; $handled = $true }
                                        "txtPowerShellMaxCmdletQueueDepth" { $propertyName = "PowerShellMaxCmdletQueueDepth"; $handled = $true }
                                        "txtPowerShellMaxCmdletsExecutionDuration" { $propertyName = "PowerShellMaxCmdletsExecutionDuration"; $handled = $true }
                                        "txtDefaultAuthPolicy" { $propertyName = "DefaultAuthenticationPolicy"; $handled = $true }
                                        "txtHierAddressBookRoot" { $propertyName = "HierarchicalAddressBookRoot"; $handled = $true }
                                        "txtPreferredInternetCodePageForShiftJis" { $propertyName = "PreferredInternetCodePageForShiftJis"; $handled = $true }
                                }
                                if (-not $handled -and $elementName -like "txt*" -and $elementName.Length -gt 3) {
                                    $propertyName = $elementName.Substring(3)
                                    $handled = $true
                                }
                                if ($handled -and $configProperties -contains $propertyName) {
                                    $value = $script:currentOrganizationConfig.$propertyName
                                    $element.Text = if ($null -ne $value) { $value.ToString() } else { "" }
                                    $script:organizationConfigSettings[$propertyName] = $value # Store original value type
                                        Write-Log "TextBox '$elementName' gesetzt (Wert: '$($element.Text)')" -Type "Info"
                                } elseif ($handled) { Write-Log "Warnung: Eigenschaft '$propertyName' nicht in OrgConfig gefunden für TextBox '$elementName'." -Type "Warning" }
                            }
                            default {
                                    Write-Log "Element '$elementName' hat unerwarteten Typ '$elementType'" -Type "Debug"
                            }
                        }
                    } catch {
                            Write-Log "Fehler beim Setzen des Werts für UI-Element '$elementName': $($_.Exception.Message)" -Type "Error"
                            # Nicht abbrechen, versuche andere Elemente weiter zu setzen
                    }
                } else {
                        Write-Log "Element '$elementName' nicht in der XAML gefunden." -Type "Debug"
                }
        }

        # Vollständige Konfiguration in TextBox anzeigen (am Ende, nachdem alles andere versucht wurde)
        $txtOrganizationConfig = Get-XamlElement -ElementName "txtOrganizationConfig"
        if ($null -ne $txtOrganizationConfig) {
            try {
                $configText = $script:currentOrganizationConfig | Format-List | Out-String
                $txtOrganizationConfig.Text = $configText
                Write-Log "Gesamte OrgConfig in txtOrganizationConfig angezeigt." -Type "Info"
            } catch {
                    Write-Log "Fehler beim Anzeigen der gesamten OrgConfig: $($_.Exception.Message)" -Type "Error"
                    $txtOrganizationConfig.Text = "Fehler beim Formatieren der OrgConfig: $($_.Exception.Message)"
            }
        }

        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organisationseinstellungen erfolgreich geladen und angezeigt."
        }
        Write-Log "UI-Aktualisierung für Organisationseinstellungen abgeschlossen." -Type "Success"
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Schwerwiegender Fehler in Get-CurrentOrganizationConfig: $errorMsg" -Type "Error"
        Write-Log $_.Exception.StackTrace -Type "Error"
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler beim Laden der Organisationseinstellungen: $errorMsg"
        }
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
        $cmbActivityTimeout = Get-XamlElement -ElementName "chkActivityBasedAuthenticationTimeoutInterval"
        if ($null -ne $cmbActivityTimeout -and $null -ne $cmbActivityTimeout.SelectedItem) {
            $selectedText = $cmbActivityTimeout.SelectedItem.Content.ToString()
            $timeoutValue = ($selectedText -split ' ')[0]
            $script:organizationConfigSettings["ActivityBasedAuthenticationTimeoutInterval"] = $timeoutValue
        }
        
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
                if ($textBoxName -eq "txtPowerShellMaxConcurrency" -or 
                    $textBoxName -eq "txtPowerShellMaxCmdletQueueDepth" -or 
                    $textBoxName -eq "txtPowerShellMaxCmdletsExecutionDuration" -or
                    $textBoxName -eq "txtPreferredInternetCodePageForShiftJis") {
                    if ([int]::TryParse($textValue, [ref]$null)) {
                        $script:organizationConfigSettings[$propertyName] = [int]$textValue
                    } else {
                        throw "Der Wert für $propertyName muss eine ganze Zahl sein."
                    }
                }
                else {
                    $script:organizationConfigSettings[$propertyName] = $textValue
                }
            }
            else {
                # Für leere TextBoxen die Eigenschaft aus den Einstellungen entfernen, falls sie existiert
                if ($script:organizationConfigSettings.ContainsKey($propertyName)) {
                    $script:organizationConfigSettings.Remove($propertyName)
                }
            }
        }
        
        # Parameter für Set-OrganizationConfig vorbereiten
        $params = @{}
        foreach ($key in $script:organizationConfigSettings.Keys) {
            $params[$key] = $script:organizationConfigSettings[$key]
        }
        
        # Debug-Log alle Parameter
        foreach ($key in $params.Keys | Sort-Object) {
            Write-Log "Parameter: $key = $($params[$key])" -Type "Debug"
        }
        
        # Organisationseinstellungen aktualisieren
        Set-OrganizationConfig @params -ErrorAction Stop
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organisationseinstellungen erfolgreich gespeichert."
        }
        Show-MessageBox -Message "Die Organisationseinstellungen wurden erfolgreich gespeichert." -Title "Erfolg" -Type "Info"
        
        # Aktuelle Konfiguration neu laden, um Änderungen zu sehen
        Get-CurrentOrganizationConfig
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Speichern der Organisationseinstellungen: $errorMsg" -Type "Error"
        Write-Log $_.Exception.StackTrace -Type "Error"
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler beim Speichern der Organisationseinstellungen: $errorMsg"
        }
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
        Show-MessageBox -Message "Die Organisationseinstellungen wurden erfolgreich nach '$exportPath' exportiert." -Title "Export erfolgreich" -Type "Info"
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler beim Exportieren der Organisationseinstellungen: $errorMsg"
        }
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
        
    }
    else {
    }
}
#endregion EXOSettings Main Functions
function Initialize-CalendarTab {
    [CmdletBinding()]
    param()
    
    try {
        
        # Referenzieren der UI-Elemente im Kalender-Tab
        $txtCalendarSource = Get-XamlElement -ElementName "txtCalendarSource"
        $txtCalendarTarget = Get-XamlElement -ElementName "txtCalendarTarget"
        $cmbCalendarPermission = Get-XamlElement -ElementName "cmbCalendarPermission"
        $btnAddCalendarPermission = Get-XamlElement -ElementName "btnAddCalendarPermission"
        $btnRemoveCalendarPermission = Get-XamlElement -ElementName "btnRemoveCalendarPermission"
        $txtCalendarMailboxUser = Get-XamlElement -ElementName "txtCalendarMailboxUser"
        $btnShowCalendarPermissions = Get-XamlElement -ElementName "btnShowCalendarPermissions"
        $cmbDefaultPermission = Get-XamlElement -ElementName "cmbDefaultPermission"
        $btnSetDefaultPermission = Get-XamlElement -ElementName "btnSetDefaultPermission"
        $btnSetAnonymousPermission = Get-XamlElement -ElementName "btnSetAnonymousPermission"
        $btnSetAllCalPermission = Get-XamlElement -ElementName "btnSetAllCalPermission"
        $lstCalendarPermissions = Get-XamlElement -ElementName "lstCalendarPermissions"
        $btnExportCalendarPermissions = Get-XamlElement -ElementName "btnExportCalendarPermissions"
        $helpLinkCalendar = Get-XamlElement -ElementName "helpLinkCalendar"
        
        # Globale Variablen für spätere Verwendung setzen
        $script:txtCalendarSource = $txtCalendarSource
        $script:txtCalendarTarget = $txtCalendarTarget
        $script:cmbCalendarPermission = $cmbCalendarPermission
        $script:txtCalendarMailboxUser = $txtCalendarMailboxUser
        $script:lstCalendarPermissions = $lstCalendarPermissions
        $script:cmbDefaultPermission = $cmbDefaultPermission
        
        # Berechtigungsstufen für Kalender hinzufügen
        $calendarPermissions = @(
            "Owner", 
            "PublishingEditor", 
            "Editor", 
            "PublishingAuthor", 
            "Author", 
            "NonEditingAuthor", 
            "Reviewer", 
            "Contributor", 
            "AvailabilityOnly", 
            "LimitedDetails", 
            "None"
        )
        
        foreach ($permission in $calendarPermissions) {
            $cmbCalendarPermission.Items.Add($permission) | Out-Null
            $cmbDefaultPermission.Items.Add($permission) | Out-Null
        }
        
        # Standardwerte setzen
        if ($cmbCalendarPermission.Items.Count -gt 0) {
            $cmbCalendarPermission.SelectedIndex = 6  # Reviewer als Standard
        }
        
        if ($cmbDefaultPermission.Items.Count -gt 0) {
            $cmbDefaultPermission.SelectedIndex = 8  # AvailabilityOnly als Standard
        }
        
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
                        "Bitte geben Sie Quell- und Zielbenutzer an und wählen Sie eine Berechtigungsstufe.",
                        "Unvollständige Eingabe",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                    return
                }
                
                $sourceMailbox = $script:txtCalendarSource.Text
                $targetUser = $script:txtCalendarTarget.Text
                $permission = $script:cmbCalendarPermission.SelectedItem.ToString()
                
                # Kalenderberechtigung hinzufügen
                $result = Add-CalendarPermission -SourceMailbox $sourceMailbox -TargetUser $targetUser -AccessRight $permission
                
                if ($result) {
                    $script:txtStatus.Text = "Kalenderberechtigung erfolgreich hinzugefügt."
                    # Aktualisiere die Berechtigungsliste, wenn der aktuelle Benutzer angezeigt wird
                    if ($script:txtCalendarMailboxUser.Text -eq $sourceMailbox) {
                        Show-CalendarPermissions -Mailbox $sourceMailbox
                    }
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                
                $sourceMailbox = $script:txtCalendarSource.Text
                $targetUser = $script:txtCalendarTarget.Text
                
                # Kalenderberechtigung entfernen
                $result = Remove-CalendarPermission -SourceMailbox $sourceMailbox -TargetUser $targetUser
                
                if ($result) {
                    $script:txtStatus.Text = "Kalenderberechtigung erfolgreich entfernt."
                    # Aktualisiere die Berechtigungsliste, wenn der aktuelle Benutzer angezeigt wird
                    if ($script:txtCalendarMailboxUser.Text -eq $sourceMailbox) {
                        Show-CalendarPermissions -Mailbox $sourceMailbox
                    }
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                
                # Prüfen, ob ein Benutzer angegeben wurde
                if ([string]::IsNullOrWhiteSpace($script:txtCalendarMailboxUser.Text)) {
                    [System.Windows.MessageBox]::Show(
                        "Bitte geben Sie einen Benutzer an, dessen Kalenderberechtigungen angezeigt werden sollen.",
                        "Unvollständige Eingabe",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                    return
                }
                
                $mailbox = $script:txtCalendarMailboxUser.Text
                
                # Kalenderberechtigungen anzeigen
                Show-CalendarPermissions -Mailbox $mailbox
                
                $script:txtStatus.Text = "Kalenderberechtigungen für $mailbox wurden geladen."
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                if ([string]::IsNullOrWhiteSpace($script:txtCalendarSource.Text) -or 
                    $null -eq $script:cmbDefaultPermission.SelectedItem) {
                    [System.Windows.MessageBox]::Show(
                        "Bitte geben Sie einen Quellbenutzer an und wählen Sie eine Standard-Berechtigungsstufe.",
                        "Unvollständige Eingabe",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                    return
                }
                
                $sourceMailbox = $script:txtCalendarSource.Text
                $permission = $script:cmbDefaultPermission.SelectedItem.ToString()
                
                # Standard-Kalenderberechtigung setzen
                $result = Set-DefaultCalendarPermission -Mailbox $sourceMailbox -AccessRight $permission
                
                if ($result) {
                    $script:txtStatus.Text = "Standard-Kalenderberechtigung erfolgreich gesetzt."
                    # Aktualisiere die Berechtigungsliste, wenn der aktuelle Benutzer angezeigt wird
                    if ($script:txtCalendarMailboxUser.Text -eq $sourceMailbox) {
                        Show-CalendarPermissions -Mailbox $sourceMailbox
                    }
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                if ([string]::IsNullOrWhiteSpace($script:txtCalendarSource.Text) -or 
                    $null -eq $script:cmbDefaultPermission.SelectedItem) {
                    [System.Windows.MessageBox]::Show(
                        "Bitte geben Sie einen Quellbenutzer an und wählen Sie eine Berechtigungsstufe für anonyme Benutzer.",
                        "Unvollständige Eingabe",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                    return
                }
                
                $sourceMailbox = $script:txtCalendarSource.Text
                $permission = $script:cmbDefaultPermission.SelectedItem.ToString()
                
                # Anonyme Kalenderberechtigung setzen
                $result = Set-AnonymousCalendarPermission -Mailbox $sourceMailbox -AccessRight $permission
                
                if ($result) {
                    $script:txtStatus.Text = "Anonyme Kalenderberechtigung erfolgreich gesetzt."
                    # Aktualisiere die Berechtigungsliste, wenn der aktuelle Benutzer angezeigt wird
                    if ($script:txtCalendarMailboxUser.Text -eq $sourceMailbox) {
                        Show-CalendarPermissions -Mailbox $sourceMailbox
                    }
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                
                # Sicherheitsabfrage
                $result = [System.Windows.MessageBox]::Show(
                    "Diese Aktion setzt die Standard-Kalenderberechtigungen für ALLE Postfächer in der Organisation. Möchten Sie fortfahren?",
                    "Massenänderung bestätigen",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Warning)
                
                if ($result -eq [System.Windows.MessageBoxResult]::No) {
                    return
                }
                
                # Prüfen, ob eine Berechtigungsstufe ausgewählt wurde
                if ($null -eq $script:cmbDefaultPermission.SelectedItem) {
                    [System.Windows.MessageBox]::Show(
                        "Bitte wählen Sie eine Standard-Berechtigungsstufe.",
                        "Unvollständige Eingabe",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                    return
                }
                
                $permission = $script:cmbDefaultPermission.SelectedItem.ToString()
                
                # Standard-Kalenderberechtigung für alle Postfächer setzen
                $result = Set-AllCalendarPermissions -AccessRight $permission
                
                if ($result) {
                    $script:txtStatus.Text = "Standard-Kalenderberechtigungen für alle Postfächer erfolgreich gesetzt."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnSetAllCalPermission"
        
        Register-EventHandler -Control $btnExportCalendarPermissions -Handler {
            try {
                # Prüfen, ob Daten zum Exportieren vorhanden sind
                if ($null -eq $script:lstCalendarPermissions.Items -or $script:lstCalendarPermissions.Items.Count -eq 0) {
                    [System.Windows.MessageBox]::Show(
                        "Es sind keine Kalenderberechtigungen zum Exportieren vorhanden. Bitte zeigen Sie zuerst die Berechtigungen eines Benutzers an.",
                        "Keine Daten",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                    return
                }
                
                # Speicherort für die CSV-Datei auswählen
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
                $saveDialog.Title = "Kalenderberechtigungen exportieren"
                $saveDialog.FileName = "Kalenderberechtigungen_$($script:txtCalendarMailboxUser.Text)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                
                $result = $saveDialog.ShowDialog()
                
                if ($result -eq $true) {
                    $exportPath = $saveDialog.FileName
                    
                    # Daten aus dem DataGrid extrahieren und als CSV exportieren
                    $script:lstCalendarPermissions.Items | 
                        Select-Object User, AccessRights, IsInherited | 
                        Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
                    
                    $script:txtStatus.Text = "Kalenderberechtigungen wurden nach $exportPath exportiert."
                    
                    # Erfolgsmeldung anzeigen
                    [System.Windows.MessageBox]::Show(
                        "Die Kalenderberechtigungen wurden erfolgreich nach '$exportPath' exportiert.",
                        "Export erfolgreich",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnExportCalendarPermissions"
        
        # Hilfe-Link initialisieren
        if ($null -ne $helpLinkCalendar) {
            $helpLinkCalendar.Add_MouseLeftButtonDown({
                Show-HelpDialog -Topic "Calendar"
            })
            
            $helpLinkCalendar.Add_MouseEnter({
                $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                $this.Cursor = [System.Windows.Input.Cursors]::Hand
            })
            
            $helpLinkCalendar.Add_MouseLeave({
                $this.TextDecorations = $null
                $this.Cursor = [System.Windows.Input.Cursors]::Arrow
            })
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        return $false
    }
}

function Initialize-MailboxTab {
    [CmdletBinding()]
    param()
    
    try {
        
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
                
                $result = Add-MailboxPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                
                if ($result) {
                    $script:txtStatus.Text = "Postfachberechtigung erfolgreich hinzugefügt."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                
                $result = Remove-MailboxPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                
                if ($result) {
                    $script:txtStatus.Text = "Postfachberechtigung erfolgreich entfernt."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                
                $result = Add-SendAsPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                
                if ($result) {
                    $script:txtStatus.Text = "SendAs-Berechtigung erfolgreich hinzugefügt."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                
                $result = Remove-SendAsPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                
                if ($result) {
                    $script:txtStatus.Text = "SendAs-Berechtigung erfolgreich entfernt."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                
                $result = Add-SendOnBehalfPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                
                if ($result) {
                    $script:txtStatus.Text = "SendOnBehalf-Berechtigung erfolgreich hinzugefügt."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                
                $result = Remove-SendOnBehalfPermissionAction -SourceUser $sourceUser -TargetUser $targetUser
                
                if ($result) {
                    $script:txtStatus.Text = "SendOnBehalf-Berechtigung erfolgreich entfernt."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        return $false
    }
}
function Initialize-GroupsTab {
    [CmdletBinding()]
    param()
    
    try {
        
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
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        return $false
    }
}

function Initialize-ResourcesTab {
    [CmdletBinding()]
    param()

    try {

        # UI-Elemente für Ressourcen referenzieren basierend auf EXOGUI.xaml
        $helpLinkResources = Get-XamlElement -ElementName "helpLinkResources"
        
        # Elemente aus der XAML referenzieren
        $cmbResourceType = Get-XamlElement -ElementName "cmbResourceType"
        $btnCreateResource = Get-XamlElement -ElementName "btnCreateResource"
        $cmbResourceSelect = Get-XamlElement -ElementName "cmbResourceSelect"
        $btnEditResourceSettings = Get-XamlElement -ElementName "btnEditResourceSettings"
        $btnRemoveResource = Get-XamlElement -ElementName "btnRemoveResource"
        $dgResources = Get-XamlElement -ElementName "dgResources"
        $btnExportResources = Get-XamlElement -ElementName "btnExportResources"

        # Globale Variablen setzen für die vorhandenen Elemente
        if ($null -ne $cmbResourceType) { $script:cmbResourceType = $cmbResourceType }
        if ($null -ne $cmbResourceSelect) { $script:cmbResourceSelect = $cmbResourceSelect }
        if ($null -ne $dgResources) { $script:dgResources = $dgResources }

        # Event-Handler für "Ressource erstellen"
        if ($null -ne $btnCreateResource) {
            Register-EventHandler -Control $btnCreateResource -Handler {
                try {
                    if (-not $script:isConnected) { Throw "Keine Verbindung zu Exchange Online."; return }
                    $name = $script:txtResourceName.Text
                    $displayName = $script:txtResourceDisplayName.Text
                    $capacity = $script:txtResourceCapacity.Text
                    $location = $script:txtResourceLocation.Text
                    $type = if ($script:cmbResourceType.SelectedItem.Content -eq "Raum") { "Room" } else { "Equipment" }

                    if ([string]::IsNullOrWhiteSpace($name)) {
                        Throw "Name der Ressource darf nicht leer sein."
                    }
                    
                    if ([string]::IsNullOrWhiteSpace($displayName)) {
                        $displayName = $name
                    }

                    $result = New-ResourceAction -Name $name -DisplayName $displayName -ResourceType $type -Capacity $capacity -Location $location
                    
                    if ($result) { 
                        $script:txtStatus.Text = "Ressource '$displayName' erfolgreich erstellt." 
                        # Felder zurücksetzen
                        $script:txtResourceName.Text = ""
                        $script:txtResourceDisplayName.Text = ""
                        $script:txtResourceCapacity.Text = ""
                        $script:txtResourceLocation.Text = ""
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnCreateResource"
        }

        # Event-Handler für "Einstellungen bearbeiten"
        if ($null -ne $btnEditResourceSettings) {
            Register-EventHandler -Control $btnEditResourceSettings -Handler {
                try {
                    if (-not $script:isConnected) { Throw "Keine Verbindung zu Exchange Online."; return }
                    $selectedResource = $script:cmbResourceSelect.SelectedItem
                    if ($null -eq $selectedResource) { Throw "Bitte wählen Sie eine Ressource aus." }
                    
                    
                    # Dialog zum Bearbeiten der Ressourceneinstellungen öffnen
                    $result = Show-ResourceSettingsDialog -Identity $selectedResource
                    
                    if ($result) {
                        # Ressourcenliste aktualisieren, wenn Änderungen vorgenommen wurden
                        $allResources = Get-AllResourcesAction
                        $script:dgResources.ItemsSource = $allResources
                        
                        # ComboBox aktualisieren
                        $script:cmbResourceSelect.Items.Clear()
                        foreach ($res in $allResources) {
                            [void]$script:cmbResourceSelect.Items.Add($res.PrimarySmtpAddress)
                        }
                        
                        $script:txtStatus.Text = "Ressourceneinstellungen für '$selectedResource' erfolgreich aktualisiert."
                    }
                    else {
                        $script:txtStatus.Text = "Bearbeitung der Ressourceneinstellungen abgebrochen."
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnEditResourceSettings"
        }

        # Event-Handler für "Ressource löschen"
        if ($null -ne $btnRemoveResource) {
            Register-EventHandler -Control $btnRemoveResource -Handler {
                try {
                    if (-not $script:isConnected) { Throw "Keine Verbindung zu Exchange Online."; return }
                    $selectedResource = $script:cmbResourceSelect.SelectedItem
                    if ($null -eq $selectedResource) { Throw "Bitte wählen Sie eine Ressource aus." }

                    $confirm = [System.Windows.MessageBox]::Show(
                        "Möchten Sie die Ressource '$selectedResource' wirklich löschen?",
                        "Bestätigung erforderlich",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Warning
                    )

                    if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
                        $result = Remove-ResourceAction -Identity $selectedResource
                        if ($result) { 
                            $script:txtStatus.Text = "Ressource '$selectedResource' erfolgreich gelöscht."
                            # Ressourcenliste aktualisieren
                            $allResources = Get-AllResourcesAction
                            $script:dgResources.ItemsSource = $allResources
                            
                            # ComboBox aktualisieren
                            $script:cmbResourceSelect.Items.Clear()
                            foreach ($res in $allResources) {
                                [void]$script:cmbResourceSelect.Items.Add($res.PrimarySmtpAddress)
                            }
                        }
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnRemoveResource"
        }

        # Event-Handler für "Ressourcenliste exportieren"
        if ($null -ne $btnExportResources) {
            Register-EventHandler -Control $btnExportResources -Handler {
                try {
                    if (-not $script:isConnected) { Throw "Keine Verbindung zu Exchange Online."; return }
                    if ($null -eq $script:dgResources.ItemsSource -or $script:dgResources.Items.Count -eq 0) {
                        Throw "Keine Ressourcen zum Exportieren vorhanden."
                    }
                    
                    
                    # SaveFileDialog erstellen
                    $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
                    $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
                    $saveFileDialog.Title = "Ressourcenliste exportieren"
                    $saveFileDialog.FileName = "Ressourcenliste_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                    
                    # Dialog anzeigen
                    $result = $saveFileDialog.ShowDialog()
                    
                    if ($result -eq $true) {
                        # Daten exportieren
                        $resources = $script:dgResources.ItemsSource
                        Export-ResourcesAction -Resources $resources -FilePath $saveFileDialog.FileName
                        $script:txtStatus.Text = "Ressourcenliste erfolgreich exportiert nach: $($saveFileDialog.FileName)"
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $script:txtStatus.Text = "Fehler: $errorMsg"
                }
            } -ControlName "btnExportResources"
        }

        # Event-Handler für Hilfe-Link
        if ($null -ne $helpLinkResources) {
            $helpLinkResources.Add_MouseLeftButtonDown({
                Show-HelpDialog -Topic "Resources"
            })
            
            $helpLinkResources.Add_MouseEnter({
                $this.TextDecorations = [System.Windows.TextDecorations]::Underline
                $this.Cursor = [System.Windows.Input.Cursors]::Hand
            })
            
            $helpLinkResources.Add_MouseLeave({
                $this.TextDecorations = $null
                $this.Cursor = [System.Windows.Input.Cursors]::Arrow
            })
        }

        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        return $false
    }
}

function Initialize-SharedMailboxTab {
    [CmdletBinding()]
    param()
    
    try {
        
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
        $chkKeepCopy = Get-XamlElement -ElementName "chkKeepCopy"
        $btnSetForwarding = Get-XamlElement -ElementName "btnSetForwarding"
        $btnRemoveForwarding = Get-XamlElement -ElementName "btnRemoveForwarding"
        $btnGetForwardingMailboxes = Get-XamlElement -ElementName "btnGetForwardingMailboxes"
        $btnHideFromGAL = Get-XamlElement -ElementName "btnHideFromGAL"
        $btnShowInGAL = Get-XamlElement -ElementName "btnShowInGAL"
        $btnRemoveSharedMailbox = Get-XamlElement -ElementName "btnRemoveSharedMailbox"
        $helpLinkShared = Get-XamlElement -ElementName "helpLinkShared"
        $lstCurrentPermissions = Get-XamlElement -ElementName "lstCurrentPermissions"
        $btnExportSharedPermissions = Get-XamlElement -ElementName "btnExportSharedPermissions"
        
        # Globale Variablen setzen
        $script:txtSharedMailboxName = $txtSharedMailboxName
        $script:txtSharedMailboxEmail = $txtSharedMailboxEmail
        $script:cmbSharedMailboxDomain = $cmbSharedMailboxDomain
        $script:txtSharedMailboxPermSource = $txtSharedMailboxPermSource
        $script:txtSharedMailboxPermUser = $txtSharedMailboxPermUser
        $script:cmbSharedMailboxPermType = $cmbSharedMailboxPermType
        $script:cmbSharedMailboxSelect = $cmbSharedMailboxSelect
        $script:chkAutoMapping = $chkAutoMapping
        $script:chkKeepCopy = $chkKeepCopy
        $script:txtForwardingAddress = $txtForwardingAddress
        $script:lstCurrentPermissions = $lstCurrentPermissions
        
        # Hilfsfunktion zum Aktualisieren der Shared Mailbox-Liste
        function RefreshSharedMailboxList {
            try {
                if ($null -ne $script:cmbSharedMailboxSelect) {
                    $script:cmbSharedMailboxSelect.Items.Clear()
                    
                    # Prüfen, ob eine Verbindung besteht
                    if (-not $script:isConnected) {
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
                    
                }
                return $true
    }
    catch {
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
                Write-Log  "Fehler beim Laden der Domains: $($_.Exception.Message)" -Type "Warning"
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
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        }
        
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
                
                # Funktion zum Anzeigen der Berechtigungen aufrufen
                $permissions = Get-SharedMailboxPermissionsAction -Mailbox $mailbox
                
                # Berechtigungen in der Benutzeroberfläche anzeigen
                if ($permissions.Count -gt 0) {
                    # DataGrid für die Anzeige der Berechtigungen verwenden
                    if ($null -ne $script:lstCurrentPermissions) {
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
                        
                        $script:lstCurrentPermissions.ItemsSource = $permissionsCollection
                        
                        $script:txtStatus.Text = "$($permissions.Count) Berechtigungen gefunden für $mailbox."
                    } else {
                        
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
                    if ($null -ne $script:lstCurrentPermissions) {
                        $script:lstCurrentPermissions.ItemsSource = $null
                    }
                    
                    $script:txtStatus.Text = "Keine Berechtigungen gefunden für $mailbox."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                
                # Funktion zum Anzeigen aller Weiterleitungen aufrufen
                $forwardingMailboxes = Get-ForwardingMailboxesAction
                
                if ($forwardingMailboxes.Count -gt 0) {
                    # Anzeige der Weiterleitungen in einem neuen Fenster oder Dialog
                    $forwardingText = "Aktive Weiterleitungen:`n`n"
                    foreach ($fwd in $forwardingMailboxes) {
                        $forwardingText += "- $($fwd.Mailbox) → $($fwd.ForwardingAddress) (Kopie behalten: $($fwd.DeliverToMailboxAndForward))`n"
                    }
                    
                    [System.Windows.MessageBox]::Show(
                        $forwardingText,
                        "Mailbox-Weiterleitungen",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information)
                    
                    $script:txtStatus.Text = "$($forwardingMailboxes.Count) Weiterleitungen gefunden."
                } else {
                    [System.Windows.MessageBox]::Show(
                        "Es wurden keine aktiven Weiterleitungen gefunden.",
                        "Mailbox-Weiterleitungen",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information)
                    
                    $script:txtStatus.Text = "Keine Weiterleitungen gefunden."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                $keepCopy = $script:chkKeepCopy.IsChecked
                
                # Funktion zum Einrichten der Weiterleitung aufrufen
                $result = Set-SharedMailboxForwardingAction -Mailbox $mailbox -ForwardingAddress $forwardingAddress -KeepCopy $keepCopy
                
                if ($result) {
                    $script:txtStatus.Text = "Weiterleitung wurde erfolgreich eingerichtet."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnSetForwarding"
        
        Register-EventHandler -Control $btnRemoveForwarding -Handler {
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
                
                # Funktion zum Entfernen der Weiterleitung aufrufen
                $result = Remove-SharedMailboxForwardingAction -Mailbox $mailbox
                
                if ($result) {
                    $script:txtStatus.Text = "Weiterleitung wurde erfolgreich entfernt."
                    $script:txtForwardingAddress.Text = ""
                    $script:chkKeepCopy.IsChecked = $false
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnRemoveForwarding"
        
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
                
                # Funktion zum Einblenden in GAL aufrufen
                $result = Set-SharedMailboxGALVisibilityAction -Mailbox $mailbox -HideFromGAL $false
                
                if ($result) {
                    $script:txtStatus.Text = "Shared Mailbox wurde in der GAL sichtbar gemacht."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
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
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        return $false
    }
}

function Initialize-AuditTab {
    [CmdletBinding()]
    param()
    
    try {
        
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
                
                $result = Get-FormattedMailboxInfo -Mailbox $mailbox -InfoType $infoType -NavigationType $navigationType
                
                if ($null -ne $script:txtAuditResult) {
                    $script:txtAuditResult.Text = $result
                }
                
                $script:txtStatus.Text = "Audit erfolgreich ausgeführt."
            }
            catch {
                $errorMsg = $_.Exception.Message
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
                    Write-Log  "Fehler beim Aktualisieren der Audit-Typen: $errorMsg" -Type "Error"
                }
            })
            
            # Initial die erste Kategorie auswählen, um die Typen zu laden
            if ($cmbAuditCategory.Items.Count -gt 0) {
                $cmbAuditCategory.SelectedIndex = 0
            }
        }
        
        Write-Log  "Audit-Tab erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Initialisieren des Audit-Tabs: $errorMsg" -Type "Error"
        return $false
    }
}
function Initialize-ReportsTab {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log  "Initialisiere Berichte-Tab" -Type "Info"
        
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
                Write-Log  "Fehler beim Generieren des Berichts: $errorMsg" -Type "Error"
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
                Write-Log  "Fehler beim Exportieren des Berichts: $errorMsg" -Type "Error"
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
        
        Write-Log  "Berichte-Tab erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Initialisieren des Berichte-Tabs: $errorMsg" -Type "Error"
        return $false
    }
}
function Initialize-TroubleshootingTab {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log  "Initialisiere Troubleshooting-Tab" -Type "Info"
        
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
                
                Write-Log  "Führe Diagnose aus: Index=$diagnosticIndex, User=$user, User2=$user2, Email=$email" -Type "Info"
                
                $result = Run-ExchangeDiagnostic -DiagnosticIndex $diagnosticIndex -User $user -User2 $user2 -Email $email
                
                if ($null -ne $script:txtDiagnosticResult) {
                    $script:txtDiagnosticResult.Text = $result
                    $script:txtStatus.Text = "Diagnose erfolgreich ausgeführt."
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log  "Fehler bei der Diagnose: $errorMsg" -Type "Error"
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
                
                Write-Log  "Öffne Admin-Center für Diagnose: Index=$diagnosticIndex" -Type "Info"
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
                Write-Log  "Fehler beim Öffnen des Admin-Centers: $errorMsg" -Type "Error"
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
        
        Write-Log  "Troubleshooting-Tab erfolgreich initialisiert" -Type "Success"
        return $true
            }
            catch {
                $errorMsg = $_.Exception.Message
        Write-Log  "Fehler beim Initialisieren des Troubleshooting-Tabs: $errorMsg" -Type "Error"
        return $false
    }
}
# Funktion zum Initialisieren des Regionaleinstellungen-Tabs
function Initialize-RegionSettingsTab {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Initialisiere Regionaleinstellungen-Tab..." -Type Info

        # GUI-Elemente finden - BITTE MIT XAML ABGLEICHEN!
        # Versuche, jedes Element zu finden und logge das Ergebnis
        $script:txtRegionMailbox = $script:Form.FindName("txtRegionMailbox")
        Write-Log ("Steuerelement 'txtRegionMailbox' gefunden: " + ($null -ne $script:txtRegionMailbox)) -Type Debug
        $script:cmbRegionLanguage = $script:Form.FindName("cmbRegionLanguage")
        Write-Log ("Steuerelement 'cmbRegionLanguage' gefunden: " + ($null -ne $script:cmbRegionLanguage)) -Type Debug
        $script:cmbRegionTimezone = $script:Form.FindName("cmbRegionTimezone")
        Write-Log ("Steuerelement 'cmbRegionTimezone' gefunden: " + ($null -ne $script:cmbRegionTimezone)) -Type Debug
        $script:btnSetRegionSettings = $script:Form.FindName("btnSetRegionSettings")
        Write-Log ("Steuerelement 'btnSetRegionSettings' gefunden: " + ($null -ne $script:btnSetRegionSettings)) -Type Debug
        $script:btnGetRegionSettings = $script:Form.FindName("btnGetRegionSettings")
        Write-Log ("Steuerelement 'btnGetRegionSettings' gefunden: " + ($null -ne $script:btnGetRegionSettings)) -Type Debug
        $script:txtRegionResult = $script:Form.FindName("txtRegionResult")
        Write-Log ("Steuerelement 'txtRegionResult' gefunden: " + ($null -ne $script:txtRegionResult)) -Type Debug
        $script:helpLinkRegionSettings = $script:Form.FindName("helpLinkRegionSettings")
        Write-Log ("Steuerelement 'helpLinkRegionSettings' gefunden: " + ($null -ne $script:helpLinkRegionSettings)) -Type Debug


        # Überprüfen, ob die *erforderlichen* Hauptelemente gefunden wurden
        # WICHTIG: Wenn die Namen oben nicht mit dem XAML übereinstimmen, schlägt diese Prüfung fehl!
        # Prüfung um btnGetRegionSettings und txtRegionResult erweitert
        if ($null -eq $script:txtRegionMailbox -or $null -eq $script:cmbRegionLanguage -or $null -eq $script:cmbRegionTimezone -or $null -eq $script:btnSetRegionSettings -or $null -eq $script:btnGetRegionSettings -or $null -eq $script:txtRegionResult) {
            Write-Log "Ein oder mehrere erforderliche Steuerelemente im Regionaleinstellungen-Tab nicht gefunden. Bitte XAML-Namen prüfen: txtRegionMailbox, cmbRegionLanguage, cmbRegionTimezone, btnSetRegionSettings, btnGetRegionSettings, txtRegionResult" -Type Error
            Log-Action "Fehler: Wichtige Steuerelemente für Regionaleinstellungen nicht gefunden. XAML-Namen prüfen!"
            Show-MessageBox -Message "Fehler beim Initialisieren des Tabs: Wichtige Steuerelemente (Textbox, Comboboxen, Buttons, Ergebnis-Textfeld) konnten nicht gefunden werden. Überprüfen Sie die XAML-Definition." -Title "Initialisierungsfehler"
            return $false
        }

        # --- Dynamische Befüllung der ComboBoxen mithilfe der dedizierten Funktionen ---
        # Variablen für Erfolg initialisieren
        $langSuccess = $false
        $tzSuccess = $false

        # Sprachen laden
        try {
            Write-Log "Versuche, Sprachen-ComboBox zu befüllen..." -Type Debug
            $langSuccess = Populate-LanguageComboBox -ComboBox $script:cmbRegionLanguage
            if (-not $langSuccess) {
                # Dies wird erreicht, wenn Populate-LanguageComboBox $false zurückgibt
                Write-Log "Warnung: Populate-LanguageComboBox hat '$false' zurückgegeben. Befüllen möglicherweise fehlgeschlagen oder Liste leer." -Type Warning
                # Optional: ComboBox deaktivieren?
                # $script:cmbRegionLanguage.IsEnabled = $false
            } else {
                 Write-Log "Aufruf von Populate-LanguageComboBox erfolgreich beendet." -Type Debug
            }
        } catch {
            # Fange spezifische Fehler ab, die *während* des Aufrufs von Populate-LanguageComboBox auftreten
            # Dies fängt z.B. den gemeldeten ValidateSet-Fehler ab, falls er in Populate-LanguageComboBox auftritt.
            $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Aufruf von Populate-LanguageComboBox."
            Write-Log "FEHLER beim Befüllen der Sprachen-ComboBox: $errorMsg" -Type Error
            Log-Action "FEHLER beim Befüllen der Sprachen-ComboBox: $errorMsg"
            # Setze Erfolg explizit auf false, da ein Fehler aufgetreten ist
            $langSuccess = $false
            # Zeige dem Benutzer eine Meldung, dass etwas schiefgelaufen ist
            Show-MessageBox -Message "Fehler beim Laden der Sprachenliste: $($_.Exception.Message)" -Title "Fehler Sprachen"
        }

        # Zeitzonen laden
        try {
            Write-Log "Versuche, Zeitzonen-ComboBox zu befüllen..." -Type Debug
            $tzSuccess = Populate-TimezoneComboBox -ComboBox $script:cmbRegionTimezone
             if (-not $tzSuccess) {
                # Dies wird erreicht, wenn Populate-TimezoneComboBox $false zurückgibt
                Write-Log "Warnung: Populate-TimezoneComboBox hat '$false' zurückgegeben. Befüllen möglicherweise fehlgeschlagen oder Liste leer." -Type Warning
                # Optional: ComboBox deaktivieren?
                # $script:cmbRegionTimezone.IsEnabled = $false
            } else {
                 Write-Log "Aufruf von Populate-TimezoneComboBox erfolgreich beendet." -Type Debug
            }
        } catch {
            # Fange spezifische Fehler ab, die *während* des Aufrufs von Populate-TimezoneComboBox auftreten
            # Dies fängt z.B. den gemeldeten ValidateSet-Fehler ab, falls er in Populate-TimezoneComboBox auftritt.
            $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Aufruf von Populate-TimezoneComboBox."
            Write-Log "FEHLER beim Befüllen der Zeitzonen-ComboBox: $errorMsg" -Type Error
            Log-Action "FEHLER beim Befüllen der Zeitzonen-ComboBox: $errorMsg"
            # Setze Erfolg explizit auf false, da ein Fehler aufgetreten ist
            $tzSuccess = $false
             # Zeige dem Benutzer eine Meldung, dass etwas schiefgelaufen ist
            Show-MessageBox -Message "Fehler beim Laden der Zeitzonenliste: $($_.Exception.Message)" -Title "Fehler Zeitzonen"
        }
        # --- Ende Dynamische Befüllung ---


        # Event-Handler für den Button "Aktuelle Einstellungen abrufen"
        # Direkte Zuweisung des Event-Handlers statt benutzerdefinierter Funktion
        if ($null -ne $script:btnGetRegionSettings) {
            $script:btnGetRegionSettings.Add_Click({
                try {
                    # Stelle sicher, dass die aufgerufene Funktion existiert
                    if (Get-Command -Name Invoke-GetRegionSettingsAction -ErrorAction SilentlyContinue) {
                        Invoke-GetRegionSettingsAction
                    } else {
                        Write-Log "Fehler: Funktion 'Invoke-GetRegionSettingsAction' nicht gefunden." -Type Error
                        Show-MessageBox -Message "Aktion 'Einstellungen abrufen' konnte nicht ausgeführt werden (Funktion nicht gefunden)." -Title "Fehler"
                    }
                } catch {
                    $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Abrufen der Regionaleinstellungen."
                    Write-Log $errorMsg -Type Error
                    Log-Action "FEHLER bei Aktion 'Einstellungen abrufen': $errorMsg"
                    Show-MessageBox -Message "Ein Fehler ist beim Abrufen der Einstellungen aufgetreten:`n$($_.Exception.Message)" -Title "Aktionsfehler"
                }
            })
            Write-Log "Event-Handler für 'btnGetRegionSettings' registriert." -Type Debug
        } else {
            # Diese Meldung sollte durch die Prüfung am Anfang eigentlich nicht mehr erreicht werden,
            # aber zur Sicherheit belassen.
            Write-Log "Button 'btnGetRegionSettings' nicht gefunden, Event-Handler wird nicht registriert." -Type Warning
            Log-Action "Warnung: Button 'btnGetRegionSettings' nicht gefunden."
        }

        # Event-Handler für den Button "Regionaleinstellungen anwenden"
        # Direkte Zuweisung des Event-Handlers
        if ($null -ne $script:btnSetRegionSettings) {
            $script:btnSetRegionSettings.Add_Click({
                try {
                     # Stelle sicher, dass die aufgerufene Funktion existiert
                    if (Get-Command -Name Invoke-SetRegionSettingsAction -ErrorAction SilentlyContinue) {
                        Invoke-SetRegionSettingsAction
                    } else {
                        Write-Log "Fehler: Funktion 'Invoke-SetRegionSettingsAction' nicht gefunden." -Type Error
                        Show-MessageBox -Message "Aktion 'Einstellungen anwenden' konnte nicht ausgeführt werden (Funktion nicht gefunden)." -Title "Fehler"
                    }
                } catch {
                    $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Anwenden der Regionaleinstellungen."
                    Write-Log $errorMsg -Type Error
                    Log-Action "FEHLER bei Aktion 'Einstellungen anwenden': $errorMsg"
                    Show-MessageBox -Message "Ein Fehler ist beim Anwenden der Einstellungen aufgetreten:`n$($_.Exception.Message)" -Title "Aktionsfehler"
                }
            })
            Write-Log "Event-Handler für 'btnSetRegionSettings' registriert." -Type Debug
        } else {
             # Diese Meldung sollte durch die Prüfung am Anfang eigentlich nicht mehr erreicht werden.
             Write-Log "Button 'btnSetRegionSettings' nicht gefunden, Event-Handler wird nicht registriert." -Type Warning
             Log-Action "Warnung: Button 'btnSetRegionSettings' nicht gefunden."
        }


        # Event-Handler für Hilfe-Link - Direkte Zuweisung der Event-Handler
        if ($null -ne $script:helpLinkRegionSettings) {
            # Klick-Event
            $script:helpLinkRegionSettings.Add_MouseLeftButtonDown({
                try {
                    # Stelle sicher, dass die aufgerufene Funktion existiert
                    if (Get-Command -Name Show-HelpDialog -ErrorAction SilentlyContinue) {
                        Show-HelpDialog -Topic "RegionSettings"
                    } else {
                         Write-Log "Fehler: Funktion 'Show-HelpDialog' nicht gefunden." -Type Error
                         Show-MessageBox -Message "Hilfe konnte nicht angezeigt werden (Funktion nicht gefunden)." -Title "Fehler"
                    }
                } catch {
                     $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Anzeigen des Hilfe-Dialogs für Regionaleinstellungen."
                     Write-Log $errorMsg -Type Error
                     Log-Action "FEHLER beim Anzeigen des Hilfe-Dialogs (RegionSettings): $errorMsg"
                     Show-MessageBox -Message "Ein Fehler ist beim Anzeigen der Hilfe aufgetreten:`n$($_.Exception.Message)" -Title "Hilfefehler"
                }
            })

            # Maus-Events für Hover-Effekt
            $script:helpLinkRegionSettings.Add_MouseEnter({
                $this.Cursor = [System.Windows.Input.Cursors]::Hand
                $this.TextDecorations = [System.Windows.TextDecorations]::Underline
            })

            $script:helpLinkRegionSettings.Add_MouseLeave({
                $this.TextDecorations = $null
                $this.Cursor = [System.Windows.Input.Cursors]::Arrow
            })
             Write-Log "Event-Handler für 'helpLinkRegionSettings' registriert." -Type Debug
        } else {
            # Dies ist weniger kritisch, daher nur eine Warnung loggen
            Write-Log "Hilfe-Link 'helpLinkRegionSettings' nicht gefunden, Event-Handler werden nicht registriert." -Type Warning
            Log-Action "Warnung: Hilfe-Link 'helpLinkRegionSettings' nicht gefunden."
        }


        Write-Log "Regionaleinstellungen-Tab erfolgreich initialisiert." -Type Success
        Log-Action "Regionaleinstellungen-Tab initialisiert."
        return $true
    }
    catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Schwerwiegender Fehler beim Initialisieren des Regionaleinstellungen-Tabs."
        Write-Log $errorMsg -Type Error
        Log-Action "FEHLER: Schwerwiegender Fehler beim Initialisieren des Regionaleinstellungen-Tabs: ${errorMsg}"
        Show-MessageBox -Message "Ein unerwarteter schwerwiegender Fehler ist beim Initialisieren des Regionaleinstellungen-Tabs aufgetreten: `n$($_.Exception.Message)`nBitte überprüfen Sie die Logs und die XAML-Definition." -Title "Schwerer Initialisierungsfehler"
        return $false
    }
}

# Verbesserte Version der HelpLinks-Initialisierung
function Initialize-HelpLinks {
    [CmdletBinding()]
    param()
    
    try {
        
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
                
            }
        }
        
        if ($foundLinks -eq 0) {
            return $false
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
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
                Write-Log "Button 'ShowConnectionStatus' geklickt." -Type Info
                # Hier später die Funktion Show-ConnectionStatus aufrufen
                if ($script:isConnected) {
                    $userName = $script:ConnectedUser
                    $tenantId = $script:ConnectedTenantId
                    [System.Windows.MessageBox]::Show("Verbunden als '$userName' mit Tenant '$tenantId'.", "Verbindungsstatus", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                } else {
                    [System.Windows.MessageBox]::Show("Sie sind aktuell nicht mit Exchange Online verbunden.", "Verbindungsstatus", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                }
            })
            Write-Log "Event-Handler für btnShowConnectionStatus hinzugefügt." -Type Info
        } else { Write-Log "Header Button 'btnShowConnectionStatus' nicht gefunden." -Type Warning }

        # Handler für Settings-Button
        if ($null -ne $btnSettings) {
            $btnSettings.Add_Click({ Show-SettingsWindow })
            Write-Log "Event-Handler für btnSettings hinzugefügt." -Type Info
        } else {
            Write-Log "Header Button 'btnSettings' nicht gefunden." -Type Warning
        }

        # Handler für Info-Button
        if ($null -ne $btnInfo) {
            $btnInfo.Add_Click({
                Write-Log "Button 'Info' geklickt." -Type Info
                # Hier später die Funktion Show-InfoDialog aufrufen
                $version = "Unbekannt"
                $appName = "easyEXO"
                try {
                     if ($null -ne $script:config -and $null -ne $script:config['General']) {
                        if ($script:config['General'].ContainsKey('Version')) { $version = $script:config['General']['Version'] }
                        if ($script:config['General'].ContainsKey('AppName')) { $appName = $script:config['General']['AppName'] }
                     }
                 } catch { Write-Log "Fehler beim Lesen der Version/AppName aus Config für Info-Dialog." -Type Warning }

                 $infoText = @"
$appName - Version $version

Einfache Verwaltung von Exchange Online Aufgaben.

Autor: PhinIT
(c) $(Get-Date -Format yyyy)
"@
                 [System.Windows.MessageBox]::Show($infoText, "Über $appName", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            })
             Write-Log "Event-Handler für btnInfo hinzugefügt." -Type Info
        } else { Write-Log "Header Button 'btnInfo' nicht gefunden." -Type Warning }

        # Handler für Close-Button
        if ($null -ne $btnClose) {
            $btnClose.Add_Click({
                Write-Log "Button 'Close' geklickt." -Type Info
                # Hier später die Funktion Close-Application aufrufen
                try {
                    if ($script:isConnected) {
                        Write-Log "Trenne Verbindung vor dem Schließen..." -Type Info
                        # Annahme: Disconnect-ExchangeOnlineSession existiert und funktioniert
                        Disconnect-ExchangeOnlineSession
                    }
                    Write-Log "Schließe Fenster..." -Type Info
                    $script:Form.Close()
                } catch {
                     Write-Log "Fehler beim Schließen: $($_.Exception.Message)" -Type Error
                     # Notfall-Schließung
                     try { $script:Form.Close() } catch {}
                }
            })
            Write-Log "Event-Handler für btnClose hinzugefügt." -Type Info
        } else { Write-Log "Header Button 'btnClose' nicht gefunden." -Type Warning }

        Write-Log "Event-Handler für Header-Buttons erfolgreich registriert." -Type Info
    } catch {
        # Fehler beim Finden der Buttons oder Registrieren der Handler
        Write-Log "Kritischer Fehler beim Registrieren der Header-Button-Handler: $($_.Exception.Message)" -Type Error
    }
    # --- Ende Header Button Event Handlers ---


function Initialize-TabNavigation {
    [CmdletBinding()]
    param()
    
    try {
        
        # Tab-Mapping definieren
        $script:tabMapping = @{
            'btnNavCalendar' = @{ TabName = 'tabCalendar'; Index = 0 }
            'btnNavMailbox' = @{ TabName = 'tabMailbox'; Index = 1 }
            'btnNavGroups' = @{ TabName = 'tabGroups'; Index = 2 }
            'btnNavSharedMailbox' = @{ TabName = 'tabSharedMailbox'; Index = 3 }
            'btnNavResources' = @{ TabName = 'tabResources'; Index = 4 }
            'btnNavContacts' = @{ TabName = 'tabContacts'; Index = 5 }
            'btnNavAudit' = @{ TabName = 'tabMailboxAudit'; Index = 6 }
            'btnNavReports' = @{ TabName = 'tabReports'; Index = 7 }
            'btnNavTroubleshooting' = @{ TabName = 'tabTroubleshooting'; Index = 8 }
            'btnNavEXOSettings' = @{ TabName = 'tabEXOSettings'; Index = 9 }
            'btnNavRegionSettings' = @{ TabName = 'tabRegion'; Index = 10 }
        }
        
        # TabControl überprüfen und initialisieren
        $tabContent = $script:Form.FindName("tabContent")
        if ($null -eq $tabContent) {
            throw "TabControl 'tabContent' nicht gefunden"
        }
        
        
        # Sicherstellen, dass Items-Collection initialisiert ist
        if ($null -eq $tabContent.Items) {
            throw "TabControl.Items ist nicht initialisiert"
        }
        
        
        # Alle Tabs überprüfen und ausblenden
        foreach ($mapping in $script:tabMapping.Values) {
            $tab = $script:Form.FindName($mapping.TabName)
            if ($null -ne $tab) {
                try {
                    $tab.Visibility = 'Collapsed'
                }
                catch {
                }
            }
            else {
            }
        }
        
        # Event-Handler für jeden Button registrieren
        foreach ($btnName in $script:tabMapping.Keys) {
            $button = $script:Form.FindName($btnName)
            $tabInfo = $script:tabMapping[$btnName]
            $tab = $script:Form.FindName($tabInfo.TabName)
            
            if ($null -ne $button -and $null -ne $tab) {
                # Event-Handler mit GetNewClosure() für korrekte Variable-Bindung
                $button.Add_Click({
                    param($sender, $e)
                    
                    try {
                        $clickedButton = $sender
                        $buttonName = $clickedButton.Name
                        $targetTabInfo = $script:tabMapping[$buttonName]
                        
                        $tabContent = $script:Form.FindName("tabContent")
                        if ($null -eq $tabContent) {
                            throw "TabControl nicht gefunden"
                        }
                        
                        # Alle Tabs ausblenden
                        foreach ($mapping in $script:tabMapping.Values) {
                            $tab = $script:Form.FindName($mapping.TabName)
                            if ($null -ne $tab) {
                                $tab.Visibility = 'Collapsed'
                            }
                        }
                        
                        # Ziel-Tab einblenden
                        $targetTab = $script:Form.FindName($targetTabInfo.TabName)
                        if ($null -ne $targetTab) {
                            $targetTab.Visibility = 'Visible'
                            $tabContent.SelectedIndex = $targetTabInfo.Index
                        }
                    }
                    catch {
                    }
                }.GetNewClosure())
                
            }
            else {
            }
        }
        
        # Initial den ersten Tab anzeigen
        $firstTabName = $script:tabMapping['btnNavCalendar'].TabName
        $firstTab = $script:Form.FindName($firstTabName)
        if ($null -ne $firstTab) {
            $firstTab.Visibility = 'Visible'
            $tabContent.SelectedIndex = 0
        }
        else {
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        return $false
    }
}

# Initialisiere alle Tabs
function Initialize-AllTabs {
    [CmdletBinding()]
    param()
    
    try {
        
        # Zuerst die Tab-Navigation initialisieren
        $navResult = Initialize-TabNavigation
        if (-not $navResult) {
        }
        
        $results = @{
            Navigation = $navResult
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
            RegionSettings = Initialize-RegionSettingsTab
        }
        
        $successCount = ($results.Values | Where-Object {$_ -eq $true}).Count
        $totalCount = $results.Count
        
        
        foreach ($tab in $results.Keys) {
            $status = if ($results[$tab]) { "erfolgreich" } else { "fehlgeschlagen" }
        }
        
        if ($successCount -eq $totalCount) {
            return $true
        } else {
            return ($successCount -gt 0)
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        return $false
    }
}

# Event-Handler für das Loaded-Event des Formulars
$script:Form.Add_Loaded({
    Log-Action "GUI-Loaded-Event ausgelöst, initialisiere Komponenten"
    
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
            Log-Action "Fehler beim Setzen der Version: $($_.Exception.Message)"
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
    Log-Action "Initialize-AllTabs Ergebnis: $result"
    
    # Hilfe-Links initialisieren
    $result = Initialize-HelpLinks
    Log-Action "Initialize-HelpLinks Ergebnis: $result"
})

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
            Write-Log $errorMsg -Type Error
            [System.Windows.MessageBox]::Show("Die XAML-Datei für das Einstellungsfenster wurde nicht gefunden.`nPfad: $xamlPath", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Return
        }

        Write-Log "Lade SettingsWindow.xaml" -Type Info
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
            Write-Log $loadErrorMsg -Type Error
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
            Write-Log "Konnte System.Windows.Forms nicht laden. Ordnerauswahl nicht verfügbar." -Type Warning
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
                 Write-Log "Ungültiger Wert für DebugEnabled in Config: '$debugEnabledValue'. Setze auf false." -Type Warning
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
                            Write-Log "Neues Log-Verzeichnis ausgewählt: $($txtLogPath.Text)" -Type Info
                        }
                    }
                    finally {
                        # Handle freigeben, um Ressourcenlecks zu vermeiden
                        $ownerWindow.ReleaseHandle()
                    }
                } catch {
                     $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Öffnen des Ordnerauswahldialogs."
                     Write-Log $errorMsg -Type Error
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
            Write-Log "Hauptfenster (\$script:Form) nicht gefunden oder ungültig. Einstellungsfenster wird nicht modal angezeigt." -Type Warning
        }


        # Fenster anzeigen
        Write-Log "Zeige Einstellungsfenster an" -Type Info
        [void]$settingsWindow.ShowDialog()
        Write-Log "Einstellungsfenster geschlossen" -Type Info

    }
    catch {
         # Allgemeiner Fehler im Einstellungsfenster (außerhalb des XAML-Ladevorgangs)
         $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Unerwarteter Fehler im Einstellungsfenster."
         Write-Log $errorMsg -Type Error
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
try {
    Log-Action "Öffne GUI-Fenster"
    
    if ($null -eq $script:Form) {
        throw "Form-Objekt ist null"
    }
    
    $tabContent = $script:Form.FindName("tabContent")
    if ($null -eq $tabContent) {
        throw "TabControl nicht gefunden"
    }
    
    if ($null -eq $tabContent.Items) {
        throw "TabControl.Items ist null"
    }
    
    if ($tabContent.Items.Count -eq 0) {
        throw "TabControl enthält keine Tabs"
    }
    
    Log-Action "TabControl Status vor ShowDialog: Items=$($tabContent.Items.Count)"
    foreach ($item in $tabContent.Items) {
        Log-Action "Tab vorhanden: $($item.Name)"
    }
    
    [void]$script:Form.ShowDialog()
    Log-Action "GUI-Fenster wurde geschlossen"
}
catch {
    $errorMsg = $_.Exception.Message
    Write-Log "Kritischer Fehler beim Laden oder Anzeigen der GUI: $errorMsg"  
    Log-Action "Stack Trace: $($_.Exception.StackTrace)"
    
    try {
        [System.Windows.MessageBox]::Show(
            "Kritischer Fehler beim Laden oder Anzeigen der GUI: $errorMsg", 
            "Fehler", 
            [System.Windows.MessageBoxButton]::OK, 
            [System.Windows.MessageBoxImage]::Error
        )
    }
    catch {
        Write-Log "Konnte keine MessageBox anzeigen. Zusätzlicher Fehler: $($_.Exception.Message)"  
    }
}

# Aufräumen nach Schließen des Fensters
if ($script:isConnected) {
    Log-Action "Trenne Exchange Online-Verbindung..."
    Disconnect-ExchangeOnlineSession
}
}
catch {
$errorMsg = $_.Exception.Message
Write-Log "Kritischer Fehler beim Laden oder Anzeigen der GUI: $errorMsg"  

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
    Write-Log "Konnte keine MessageBox anzeigen. Zusätzlicher Fehler: $($_.Exception.Message)"  
}
}
finally {
# Aufräumarbeiten
if ($null -ne $script:Form) {
    $script:Form.Close()
    $script:Form = $null
}

Log-Action "Aufräumarbeiten abgeschlossen"
}