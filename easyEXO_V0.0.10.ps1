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

function Test-PowerShell7AndAdminRights {
    [CmdletBinding()]
    param()

    try {
        $psVersion = $PSVersionTable.PSVersion
        $isPSCore = $psVersion.Major -ge 7
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $scriptPath = $MyInvocation.MyCommand.Path
        $currentPSEnginePath = (Get-Process -Id $PID).Path # Pfad zur aktuellen powershell.exe oder pwsh.exe

        Write-LogEntry "Aktueller Status: PowerShell Version $($psVersion.ToString()), Läuft als Administrator: $isAdmin"

        # Idealfall: Bereits PS7+ und Admin
        if ($isPSCore -and $isAdmin) {
            Write-LogEntry "Optimale Bedingungen (PowerShell 7+ und Administratorrechte) sind erfüllt."
            return $true
        }

        # PowerShell 7 Pfad suchen (nur wenn nicht bereits PS7+ und $ps7ExecutablePath noch nicht gesetzt)
        $ps7ExecutablePath = $null
        if (-not $isPSCore) {
            $ps7SearchPaths = @(
                Join-Path $env:ProgramFiles "PowerShell\7\pwsh.exe"
                Join-Path $env:ProgramFiles "(x86)\PowerShell\7\pwsh.exe"
                Join-Path $env:LOCALAPPDATA "Programs\PowerShell\7\pwsh.exe"
                # Versuche, pwsh.exe aus dem PATH zu finden, das nicht die Windows PowerShell ist
                (Get-Command pwsh -ErrorAction SilentlyContinue | Where-Object { $_.Source -and $_.Source -notlike "*\System32\WindowsPowerShell\*" -and $_.Source -notlike "*\SysWOW64\WindowsPowerShell\*" } | Select-Object -ExpandProperty Source -First 1)
            )
            $ps7ExecutablePath = $ps7SearchPaths | Where-Object { $_ -ne $null -and (Test-Path $_ -PathType Leaf) } | Select-Object -First 1
            
            if ($ps7ExecutablePath) {
                Write-LogEntry "PowerShell 7 gefunden unter: $ps7ExecutablePath"
            } else {
                Write-LogEntry "PowerShell 7 wurde auf dem System nicht gefunden."
            }
        }

        # Bedingungen für Neustart oder Installation
        $needsAdminPrivileges = (-not $isAdmin)
        $needsPS7Upgrade = (-not $isPSCore -and $null -ne $ps7ExecutablePath) # PS7 ist da, aber wir nutzen es nicht
        $needsPS7Installation = (-not $isPSCore -and $null -eq $ps7ExecutablePath) # PS7 ist nicht da und wir nutzen es nicht

        # Fall 1: PowerShell 7 muss installiert werden
        if ($needsPS7Installation) {
            $installMsg = "PowerShell 7 wird für dieses Skript empfohlen, wurde aber nicht gefunden."
            if ($needsAdminPrivileges) {
                $installMsg += " Zusätzlich sind Administratorrechte für einige Operationen und die Installation erforderlich."
            }
            $installMsg += " Möchten Sie PowerShell 7 jetzt installieren?"
            
            $userChoiceInstall = [System.Windows.MessageBox]::Show($installMsg, "PowerShell 7 Installation", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            
            if ($userChoiceInstall -eq [System.Windows.MessageBoxResult]::No) {
                Write-LogEntry "Benutzer hat die Installation von PowerShell 7 abgelehnt."
                [System.Windows.MessageBox]::Show("Ohne PowerShell 7 und/oder Administratorrechte können einige Funktionen des Skripts eingeschränkt sein oder fehlschlagen.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return $false 
            }

            Write-LogEntry "Benutzer hat der Installation von PowerShell 7 zugestimmt."
            $useWinget = $false
            try {
                $null = winget --version # Einfacher Test, ob winget existiert und funktioniert
                if ($LASTEXITCODE -eq 0) { $useWinget = $true }
            } catch { $useWinget = $false }

            if ($useWinget) {
                Write-LogEntry "Versuche PowerShell 7 Installation via winget."
                # Winget benötigt Admin für systemweite Installation. Start-Process mit RunAs für winget selbst.
                try {
                    Start-Process -FilePath "winget" -ArgumentList "install Microsoft.PowerShell --accept-package-agreements --accept-source-agreements" -Verb RunAs -Wait
                    Write-LogEntry "Winget-Installation von PowerShell 7 abgeschlossen (oder versucht)."
                } catch {
                     Write-LogEntry "Fehler beim Starten der Winget-Installation als Admin: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show("Fehler beim Starten der Winget-Installation für PowerShell 7: $($_.Exception.Message)`nVersuchen Sie, die Installation manuell als Administrator durchzuführen.", "Installationsfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return $false
                }
            } else {
                Write-LogEntry "Winget nicht verfügbar. Versuche PowerShell 7 Installation via MSI."
                $installerUrl = "https://github.com/PowerShell/PowerShell/releases/download/v7.4.2/PowerShell-7.4.2-win-x64.msi" # Aktuelle LTS Version
                $installerPath = Join-Path $env:TEMP "PowerShell-latest-win-x64.msi"
                try {
                    Write-LogEntry "Downloade PowerShell 7 MSI von $installerUrl"
                    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing
                    Write-LogEntry "PowerShell 7 MSI heruntergeladen nach $installerPath. Starte Installation."
                    # MSI Installation benötigt Admin. Start-Process mit RunAs.
                    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$installerPath`" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1" -Verb RunAs -Wait
                    Write-LogEntry "MSI-Installation von PowerShell 7 abgeschlossen (oder versucht)."
                } catch {
                    Write-LogEntry "FEHLER bei MSI Download/Installation: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show("Fehler beim Herunterladen oder Installieren von PowerShell 7 via MSI: $($_.Exception.Message)`nVersuchen Sie, die Installation manuell als Administrator durchzuführen.", "Installationsfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return $false
                } finally {
                    if (Test-Path $installerPath) { Remove-Item -Path $installerPath -Force }
                }
            }
            
            # Pfad nach Installation erneut suchen
            $ps7SearchPathsAfterInstall = @(
                Join-Path $env:ProgramFiles "PowerShell\7\pwsh.exe"
                Join-Path $env:ProgramFiles "(x86)\PowerShell\7\pwsh.exe"
                Join-Path $env:LOCALAPPDATA "Programs\PowerShell\7\pwsh.exe"
                (Get-Command pwsh -ErrorAction SilentlyContinue | Where-Object { $_.Source -and $_.Source -notlike "*\System32\WindowsPowerShell\*" -and $_.Source -notlike "*\SysWOW64\WindowsPowerShell\*" } | Select-Object -ExpandProperty Source -First 1)
            )
            $ps7ExecutablePath = $ps7SearchPathsAfterInstall | Where-Object { $_ -ne $null -and (Test-Path $_ -PathType Leaf) } | Select-Object -First 1

            if (-not $ps7ExecutablePath) {
                Write-LogEntry "PowerShell 7 konnte nach der Installation nicht gefunden werden."
                [System.Windows.MessageBox]::Show("PowerShell 7 wurde installiert, konnte aber nicht automatisch gefunden werden. Bitte starten Sie das Skript manuell mit PowerShell 7 (und ggf. Administratorrechten).", "Installationshinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return $false
            }
            Write-LogEntry "PowerShell 7 nach Installation gefunden unter: $ps7ExecutablePath"
            $needsPS7Upgrade = $true # Da wir gerade installiert haben, wollen wir es auch nutzen.
            $needsPS7Installation = $false # Nicht mehr relevant
        }

        # Fall 2: Neustart erforderlich (für Admin-Rechte und/oder PS7-Upgrade)
        # Die Bedingung ($needsAdminPrivileges -or $needsPS7Upgrade) deckt folgende Fälle ab:
        # 1. PowerShell < 7 ohne Admin-Rechte: $needsAdminPrivileges ist true, $needsPS7Upgrade ist true (falls PS7 gefunden). Neustart mit PS7 und Admin.
        # 2. PowerShell < 7 mit Admin-Rechten: $needsAdminPrivileges ist false, $needsPS7Upgrade ist true (falls PS7 gefunden). Neustart mit PS7 und Admin.
        # 3. PowerShell 7 ohne Admin-Rechte: $needsAdminPrivileges ist true, $needsPS7Upgrade ist false. Neustart mit aktueller PS7 und Admin. (Dies erfüllt die Anforderung)
        if ($needsAdminPrivileges -or $needsPS7Upgrade) {
            $restartMsgParts = @()
            $targetExecutableForRestart = $currentPSEnginePath # Standard: aktuelle PS-Engine

            if ($needsPS7Upgrade) { # PS7 ist verfügbar (oder gerade installiert) und wir sind nicht in PS7
                $restartMsgParts += "PowerShell 7"
                $targetExecutableForRestart = $ps7ExecutablePath
            }
            if ($needsAdminPrivileges) {
                $restartMsgParts += "Administratorrechten"
            }
            
            $reasonForRestart = $restartMsgParts -join " und "
            $currentContextDesc = "Sie verwenden derzeit PowerShell $($psVersion.Major).$($psVersion.Minor)"
            if ($isAdmin) { $currentContextDesc += " mit Administratorrechten." } else { $currentContextDesc += " ohne Administratorrechte."}

            $restartQueryMsg = "$currentContextDesc Für optimale Funktionalität wird ein Neustart mit $reasonForRestart empfohlen."
            if ($needsPS7Upgrade -and $ps7ExecutablePath) {
                 $restartQueryMsg += " PowerShell 7 ist unter '$ps7ExecutablePath' verfügbar."
            }
            $restartQueryMsg += " Möchten Sie das Skript jetzt neu starten?"

            $userChoiceRestart = [System.Windows.MessageBox]::Show($restartQueryMsg, "Neustart empfohlen", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)

            if ($userChoiceRestart -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-LogEntry "Benutzer stimmt Neustart zu. Ziel-Executable: $targetExecutableForRestart, Admin-Rechte werden angefordert."
                $argumentsForRestart = "-File `"$scriptPath`""
                try {
                    Start-Process -FilePath $targetExecutableForRestart -ArgumentList $argumentsForRestart -Verb RunAs
                    Write-LogEntry "Neustart-Prozess wurde initiiert."
                    exit # Aktuelles Skript beenden, da der neue Prozess gestartet wird
                } catch {
                    $errMsg = $_.Exception.Message
                    Write-LogEntry "FEHLER beim Versuch, das Skript neu zu starten: $errMsg"
                    [System.Windows.MessageBox]::Show("Fehler beim Versuch, das Skript neu zu starten: $errMsg", "Neustartfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return $false # Neustart ist fehlgeschlagen
                }
            } else {
                Write-LogEntry "Benutzer hat den empfohlenen Neustart abgelehnt."
                [System.Windows.MessageBox]::Show("Ohne die empfohlenen Einstellungen (PowerShell 7 und/oder Administratorrechte) können einige Funktionen des Skripts eingeschränkt sein oder fehlschlagen.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return $false 
            }
        }

        # Wenn wir hier ankommen, bedeutet das, dass entweder die Bedingungen initial nicht optimal waren
        # und der Benutzer die Korrekturmaßnahmen (Installation/Neustart) abgelehnt hat.
        Write-LogEntry "Die optimalen Ausführungsbedingungen wurden nicht erreicht oder vom Benutzer abgelehnt."
        return $false
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-LogEntry "Ein schwerwiegender Fehler ist in der Funktion Test-PowerShell7AndAdminRights aufgetreten: $errorMsg"
        Write-LogEntry "KRITISCHER FEHLER (Test-PowerShell7AndAdminRights): Ein interner Fehler ist bei der Überprüfung der Ausführungsumgebung aufgetreten: $errorMsg. Das Skript wird möglicherweise nicht korrekt funktionieren."
        return $false # Im Falle eines unerwarteten Fehlers in der Funktion selbst
    }
}

# Check for PowerShell 7 at startup
Test-PowerShell7AndAdminRights
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

# Globale Variable für gültige Datumsformate pro Kultur für Exchange Online
$script:ExchangeValidDateFormats = @{
    "de-DE" = @( # Deutsch (Deutschland)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" },
        @{ Display = "T.M.JJJJ"; Value = "d.M.yyyy" },
        @{ Display = "TT.MM.JJ"; Value = "dd.MM.yy" },
        @{ Display = "T.M.JJ"; Value = "d.M.yy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "de-AT" = @( # Deutsch (Österreich)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" },
        @{ Display = "T.M.JJJJ"; Value = "d.M.yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "de-CH" = @( # Deutsch (Schweiz)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" },
        @{ Display = "T.M.JJJJ"; Value = "d.M.yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "en-US" = @( # Englisch (USA)
        @{ Display = "MM/TT/JJJJ (Standard)"; Value = "MM/dd/yyyy" },
        @{ Display = "M/T/JJJJ"; Value = "M/d/yyyy" },
        @{ Display = "MM/TT/JJ"; Value = "MM/dd/yy" },
        @{ Display = "M/T/JJ"; Value = "M/d/yy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "en-GB" = @( # Englisch (UK) - Gemäß Ihrer Fehlermeldung
        @{ Display = "TT/MM/JJJJ (Standard)"; Value = "dd/MM/yyyy" },
        @{ Display = "TT/MM/JJ"; Value = "dd/MM/yy" },
        @{ Display = "T/M/JJ"; Value = "d/M/yy" }, # Angepasst an Fehlermeldung (war d/M/yyyy)
        @{ Display = "T.M.JJ"; Value = "d.M.yy" },   # Hinzugefügt gemäß Fehlermeldung
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "fr-FR" = @( # Französisch (Frankreich)
        @{ Display = "TT/MM/JJJJ (Standard)"; Value = "dd/MM/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "fr-CA" = @( # Französisch (Kanada) - Oft JJJJ-MM-TT bevorzugt
        @{ Display = "JJJJ-MM-TT (Standard)"; Value = "yyyy-MM-dd" },
        @{ Display = "TT/MM/JJJJ"; Value = "dd/MM/yyyy" }
    );
    "fr-CH" = @( # Französisch (Schweiz)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" }, # Punkte statt Schrägstriche
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "it-IT" = @( # Italienisch (Italien)
        @{ Display = "TT/MM/JJJJ (Standard)"; Value = "dd/MM/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "it-CH" = @( # Italienisch (Schweiz)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" }, # Punkte statt Schrägstriche
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "pl-PL" = @( # Polnisch (Polen)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" }, # Punkte als Trennzeichen
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "es-ES" = @( # Spanisch (Spanien)
        @{ Display = "TT/MM/JJJJ (Standard)"; Value = "dd/MM/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "es-MX" = @( # Spanisch (Mexiko) - Ähnlich wie Spanien
        @{ Display = "TT/MM/JJJJ (Standard)"; Value = "dd/MM/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "nl-NL" = @( # Niederländisch (Niederlande)
        @{ Display = "T-M-JJJJ (Standard)"; Value = "d-M-yyyy" },
        @{ Display = "TT-MM-JJJJ"; Value = "dd-MM-yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "nl-BE" = @( # Niederländisch (Belgien)
        @{ Display = "T/M/JJJJ (Standard)"; Value = "d/M/yyyy" }, # Schrägstriche in Belgien
        @{ Display = "TT/MM/JJJJ"; Value = "dd/MM/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    # Fallback / Standard, falls eine Kultur nicht spezifisch abgedeckt ist
    "DEFAULT" = @(
        @{ Display = "Systemstandard (Keine explizite Auswahl)"; Value = "" }, # Leerer Value für keine Änderung
        @{ Display = "TT.MM.JJJJ"; Value = "dd.MM.yyyy" },
        @{ Display = "MM/TT/JJJJ"; Value = "MM/dd/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    )
}
# Globale Variable für relevante Zeitformate pro Kultur für Exchange Online
$script:RelevantTimeFormatsPerCulture = @{
    "de-DE" = @( # Deutsch (Deutschland)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" }
    );
    "de-AT" = @( # Deutsch (Österreich)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" }
    );
    "de-CH" = @( # Deutsch (Schweiz) - Oft mit Punkt
        @{ Display = "HH.mm (24h, Standard)"; Value = "HH.mm" },
        @{ Display = "H.mm (24h)"; Value = "H.mm" }
    );
    "en-US" = @( # Englisch (USA)
        @{ Display = "h:mm tt (12h AM/PM, Standard)"; Value = "h:mm tt" },
        @{ Display = "hh:mm tt (12h AM/PM)"; Value = "hh:mm tt" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" },
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" }
    );
    "en-GB" = @( # Englisch (UK)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" },
        @{ Display = "h:mm tt (12h AM/PM)"; Value = "h:mm tt" }
    );
    "fr-FR" = @( # Französisch (Frankreich)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" }
    );
    "fr-CA" = @( # Französisch (Kanada)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" }, # Oft 24h bevorzugt
        @{ Display = "h:mm tt (12h AM/PM)"; Value = "h:mm tt" }
    );
    "fr-CH" = @( # Französisch (Schweiz) - Oft mit Punkt
        @{ Display = "HH.mm (24h, Standard)"; Value = "HH.mm" },
        @{ Display = "H.mm (24h)"; Value = "H.mm" }
    );
    "it-IT" = @( # Italienisch (Italien)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" }
    );
    "it-CH" = @( # Italienisch (Schweiz) - Oft mit Punkt
        @{ Display = "HH.mm (24h, Standard)"; Value = "HH.mm" },
        @{ Display = "H.mm (24h)"; Value = "H.mm" }
    );
    "pl-PL" = @( # Polnisch (Polen)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" }
    );
    "es-ES" = @( # Spanisch (Spanien)
        @{ Display = "H:mm (24h, Standard)"; Value = "H:mm" }, # Kann auch HH:mm sein
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" }
    );
    "es-MX" = @( # Spanisch (Mexiko)
        @{ Display = "h:mm tt (12h AM/PM, Standard)"; Value = "h:mm tt" },
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" }
    );
    "nl-NL" = @( # Niederländisch (Niederlande)
        @{ Display = "H:mm (24h, Standard)"; Value = "H:mm" }, # Oder HH:mm
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" }
    );
    "nl-BE" = @( # Niederländisch (Belgien)
        @{ Display = "H:mm (24h, Standard)"; Value = "H:mm" }, # Oder HH:mm
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" }
    );
    "DEFAULT" = @( # Allgemeine Fallbacks, wenn keine spezifische Kultur passt
        @{ Display = "Systemstandard (Keine explizite Auswahl)"; Value = "" },
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" },
        @{ Display = "h:mm tt (12h AM/PM)"; Value = "h:mm tt" }
    )
}
# Globale Variable für relevante Zeitzonen-IDs pro Kultur für Exchange Online
$script:RelevantTimezonesPerCulture = @{
    "de-DE" = @("W. Europe Standard Time", "Central European Standard Time"); # Berlin, Amsterdam, Paris, Rome
    "de-AT" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Vienna
    "de-CH" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Bern, Zurich
    "en-US" = @("Pacific Standard Time", "Mountain Standard Time", "Central Standard Time", "Eastern Standard Time", "Alaskan Standard Time", "Hawaiian Standard Time");
    "en-GB" = @("GMT Standard Time", "Greenwich Standard Time"); # London, Dublin
    "fr-FR" = @("Romance Standard Time", "Central European Standard Time"); # Paris
    "fr-CA" = @("Eastern Standard Time", "Central Standard Time", "Mountain Standard Time", "Pacific Standard Time", "Newfoundland Standard Time", "Atlantic Standard Time"); # Canada
    "fr-CH" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Geneva
    "it-IT" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Rome
    "it-CH" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Italian-speaking Switzerland
    "pl-PL" = @("Central European Standard Time"); # Warsaw
    "es-ES" = @("Romance Standard Time", "Central European Standard Time"); # Madrid
    "es-MX" = @("Central Standard Time (Mexico)", "Mountain Standard Time (Mexico)", "Pacific Standard Time (Mexico)"); # Mexico City, Chihuahua, Tijuana
    "nl-NL" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Amsterdam
    "nl-BE" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Brussels
    # Fallback für nicht explizit gemappte Kulturen - hier könnten alle Zeitzonen geladen werden oder eine Auswahl häufiger
    "DEFAULT_ALL" = $true # Ein Flag, um alle Zeitzonen zu laden, wenn keine spezifische Kultur passt
    # Oder eine kleinere Default-Liste:
    # "DEFAULT" = @("UTC", "GMT Standard Time", "W. Europe Standard Time", "Central European Standard Time", "Eastern Standard Time", "Pacific Standard Time")
}
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
        Write-LogEntry -Message "$Title - $Type - $Message" -Type "Info"
        
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
# Registry-Pfad und Standardwerte für Konfigurationseinstellungen
$script:registryPath = "HKCU:\Software\easyIT\easyEXO"
$currentScriptVersion = "0.0.9" # Aktuelle Version des Skripts

try {
    Write-LogEntry "Prüfe und initialisiere Registry-Konfiguration unter '$($script:registryPath)'."

    # Stelle sicher, dass der Basispfad "HKCU:\Software\easyIT" existiert
    $parentPath = "HKCU:\Software\easyIT"
    if (-not (Test-Path -Path $parentPath)) {
        New-Item -Path "HKCU:\Software" -Name "easyIT" -Force -ErrorAction Stop | Out-Null
        Write-LogEntry "Registry-Basispfad '$parentPath' wurde erstellt."
    }

    # Prüfe, ob der Anwendungspfad existiert.
    $appPathExistedBefore = Test-Path -Path $script:registryPath
    if (-not $appPathExistedBefore) {
        # Wenn der Anwendungspfad nicht existiert, erstelle ihn.
        New-Item -Path $parentPath -Name (Split-Path $script:registryPath -Leaf) -Force -ErrorAction Stop | Out-Null
        Write-LogEntry "Registry-Anwendungspfad '$($script:registryPath)' wurde erstellt."
    }
    
    $performUpdate = $false
    
    if (-not $appPathExistedBefore) {
        # Wenn der Anwendungspfad gerade erst erstellt wurde, ist ein vollständiges Setzen der Standardwerte erforderlich.
        Write-LogEntry "Registry-Anwendungspfad war nicht vorhanden. Standardwerte werden initial geschrieben."
        $performUpdate = $true
    } else {
        # Der Pfad existierte bereits. Prüfe die Version.
        $storedVersion = $null
        try {
            # Versuche, den Versionseintrag zu lesen
            $storedVersionProperty = Get-ItemProperty -Path $script:registryPath -Name "Version" -ErrorAction SilentlyContinue
            if ($null -ne $storedVersionProperty -and $storedVersionProperty.PSObject.Properties["Version"]) {
                $storedVersion = $storedVersionProperty.Version
            }
        }
        catch {
            # Fehler beim Lesen der Version, sicherheitshalber Update durchführen
            Write-LogEntry "WARNUNG: Fehler beim Lesen der Version aus der Registry unter '$($script:registryPath)': $($_.Exception.Message). Standardwerte werden vorsichtshalber aktualisiert."
            $performUpdate = $true # Update erzwingen bei Lesefehler
        }

        if (-not $performUpdate) { # Nur prüfen, wenn nicht schon durch Fehler oben ein Update erzwungen wurde
            if ($null -eq $storedVersion) {
                Write-LogEntry "Kein Versionseintrag in der Registry gefunden unter '$($script:registryPath)' oder Wert ist null. Standardwerte werden geschrieben."
                $performUpdate = $true
            } elseif ($storedVersion -ne $currentScriptVersion) {
                Write-LogEntry "Registry-Version ('$storedVersion') unterscheidet sich von Skript-Version ('$currentScriptVersion'). Standardwerte werden aktualisiert."
                $performUpdate = $true
            } else {
                Write-LogEntry "Registry-Version ('$storedVersion') ist aktuell mit Skript-Version ('$currentScriptVersion'). Keine Aktualisierung der Standardwerte erforderlich."
            }
        }
    }
    
    if ($performUpdate) {
        Write-LogEntry "Setze/Aktualisiere Registry-Standardwerte für '$($script:registryPath)'."
        New-ItemProperty -Path $script:registryPath -Name "Debug" -Value 0 -PropertyType DWORD -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "AppName" -Value easyEXO - Exchange Online Verwaltung" -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "Version" -Value $currentScriptVersion -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "ThemeColor" -Value "#0078D7" -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "LogPath" -Value "$PSScriptRoot\Logs" -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "HeaderLogoURL" -Value "https://psscripts.de" -PropertyType String -Force -ErrorAction Stop | Out-Null
        Write-LogEntry "Registry-Standardwerte für '$($script:registryPath)' erfolgreich gesetzt/aktualisiert."
    }
}
catch {
    $errorMsg = $_.Exception.Message
    Write-LogEntry "FEHLER bei der Registry-Konfigurationsinitialisierung für '$($script:registryPath)': $errorMsg"
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
                "Debug" = "0"
                "AppName" = "Exchange Online Verwaltung"
                "Version" = "0.0.9"
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
            Write-LogEntry  "Logverzeichnis wurde erstellt: $logFolder" -Type "Info" # Konsolenausgabe hiervon wird durch Write-Log gesteuert
        }
        
        # Log-Eintrag schreiben
        Add-Content -Path $script:logFilePath -Value "[$timestamp] $sanitizedMessage" -Encoding UTF8
        
        # Bei zu langer Logdatei (>10 MB) rotieren
        $logFile = Get-Item -Path $script:logFilePath -ErrorAction SilentlyContinue
        if ($logFile -and $logFile.Length -gt 10MB) {
            $backupLogPath = "$($script:logFilePath)_$(Get-Date -Format 'yyyyMMdd_HHmmss').bak"
            Move-Item -Path $script:logFilePath -Destination $backupLogPath -Force
            Write-LogEntry  "Logdatei wurde rotiert: $backupLogPath" -Type "Info" # Konsolenausgabe hiervon wird durch Write-Log gesteuert
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

        # Nachricht formatieren für Konsolenausgabe
        $formattedMessage = "[$timestamp] [$Type] $Message"

        # Ausgabe in Konsole, wenn nicht unterdrückt UND Debug-Modus aktiv ist
        if (-not $NoConsole -and $script:debugMode) {
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

        # Logging mit Log-Action, wenn nicht unterdrückt und Log-Pfad gesetzt ist
        if (-not $NoLog -and $script:logFilePath) {
            try {
                Log-Action -Message "[$Type] $Message"
            }
            catch {
                if ($script:debugMode) {
                    # Sanitize die Fehlermeldung für die Konsolenausgabe
                    $logActionCallError = $($_.Exception.Message) -replace '[^\x20-\x7E\r\n]', '?'
                    Write-Host "Fehler beim Aufruf von Log-Action innerhalb von Write-Log: $logActionCallError" -ForegroundColor Red
                }
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

            # Direkter Fallback zur Ausgabe ohne Farbe auf der Konsole, nur wenn Debug-Modus aktiv ist
            if ($script:debugMode) {
                Write-Host $errorMessage -ForegroundColor Red
            }

            # Versuch, den Fehler mit Log-Action zu loggen, falls möglich und Log-Pfad vorhanden
            if ($script:logFilePath) {
                try {
                    # Log-Action kümmert sich um Zeitstempel und Fehlerbehandlung beim Schreiben.
                    Log-Action -Message $errorMessage
                }
                catch {
                    if ($script:debugMode) {
                        $criticalLogActionError = $($_.Exception.Message) -replace '[^\x20-\x7E\r\n]', '?'
                        Write-Host "Kritischer Fehler: Log-Action konnte den Fehler in Write-Log nicht protokollieren. Fehler beim Aufruf von Log-Action: $criticalLogActionError" -ForegroundColor Red
                    }
                }
            }
        }
        catch {
        }
    }
}

# Funktion zur Initialisierung des Loggings für Log-Action
function Initialize-Logging {
    [CmdletBinding()]
    param()

    try {
        # Standard-Logverzeichnis definieren
        $defaultLogDirectory = Join-Path -Path $PSScriptRoot -ChildPath "Logs"
        $logDirectoryToUse = $defaultLogDirectory # Mit Standardwert beginnen

        # Versuchen, den Log-Pfad aus der Konfiguration zu laden
        if ($null -ne $script:config -and
            $script:config.ContainsKey("Paths") -and
            ($null -ne $script:config["Paths"]) -and # Prüfen, ob "Paths" selbst nicht $null ist
            $script:config["Paths"].ContainsKey("LogPath") -and
            -not [string]::IsNullOrWhiteSpace($script:config["Paths"]["LogPath"])) {

            $configuredLogPathDir = $script:config["Paths"]["LogPath"]

            # Überprüfen, ob der konfigurierte Pfad absolut oder relativ ist
            if ([System.IO.Path]::IsPathRooted($configuredLogPathDir)) {
                $logDirectoryToUse = [System.IO.Path]::GetFullPath($configuredLogPathDir)
            } else {
                # Wenn relativ, relativ zum Skriptverzeichnis auflösen
                $pathForNormalization = Join-Path -Path $PSScriptRoot -ChildPath $configuredLogPathDir
                $logDirectoryToUse = [System.IO.Path]::GetFullPath($pathForNormalization)
            }
        }

        # Log-Dateiname festlegen (dieser Name wird von Log-Action für die Rotation verwendet)
        $logFileName = "easyEXO_activity.log"
        $script:logFilePath = Join-Path -Path $logDirectoryToUse -ChildPath $logFileName

        Log-Action "Log-Action System initialisiert. Logdatei: $($script:logFilePath)"

        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($script:debugMode) {
            Write-Warning "FEHLER bei der Initialisierung des Log-Pfades für Log-Action: $errorMsg. Log-Action wird versuchen, den internen Fallback zu verwenden."
        }
        return $false
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
        Write-Log  -Message $Message -Type $Type # Konsolenausgabe hiervon wird durch Write-Log gesteuert
        
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
        Write-Log  "Fehler in Write-StatusMessage: $errorMsg" -Type "Error" # Konsolenausgabe hiervon wird durch Write-Log gesteuert
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
        Write-Log "Verbindungsversuch zu Exchange Online..." -Type "Info"
        
        # Prüfen, ob das ExchangeOnlineManagement Modul installiert ist
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            $errorMsg = "ExchangeOnlineManagement Modul ist nicht installiert. Bitte installieren Sie das Modul mit 'Install-Module ExchangeOnlineManagement -Force'"
            Write-Log $errorMsg -Type "Error"
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
            Write-Log $errorMsg -Type "Warning"
            Show-MessageBox -Message $errorMsg -Title "Abgebrochen" -Type "Warning"
            return $false
        }
        
        # Überprüfen, ob die E-Mail-Adresse erfolgreich gespeichert wurde
        if ([string]::IsNullOrWhiteSpace($script:userPrincipalName)) {
            $errorMsg = "Keine E-Mail-Adresse eingegeben oder erkannt. Verbindung abgebrochen."
            Write-Log $errorMsg -Type "Warning"
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
        
        Write-Log "Exchange Online Verbindung erfolgreich hergestellt für $script:userPrincipalName" -Type "Success"
        $script:txtConnectionStatus.Text = "Verbunden mit Exchange Online ($script:userPrincipalName)"
        $script:txtConnectionStatus.Foreground = "#008000"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Verbinden mit Exchange Online: $errorMsg" -Type "Error"
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
            
            # Versuche, PowerShellGet zu aktualisieren
            # Die Überprüfung auf Administratorrechte und der Neustart-Mechanismus wurden entfernt.
            # Es wird davon ausgegangen, dass das Skript bei Bedarf mit erhöhten Rechten ausgeführt wird.
            try {
                Write-Log "Versuche PowerShellGet zu aktualisieren/installieren. Administratorrechte könnten erforderlich sein." -Type "Info"
                Install-Module PowerShellGet -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                Write-Log  "PowerShellGet erfolgreich aktualisiert/installiert für den aktuellen Benutzer." -Type "Success"
            } 
            catch {
                Write-Log  "Fehler beim Aktualisieren/Installieren von PowerShellGet für den aktuellen Benutzer: $($_.Exception.Message). Versuche systemweite Installation." -Type "Warning"
                try {
                    Install-Module PowerShellGet -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
                    Write-Log  "PowerShellGet erfolgreich systemweit aktualisiert/installiert." -Type "Success"
                }
                catch {
                    Write-Log  "Fehler beim systemweiten Aktualisieren/Installieren von PowerShellGet: $($_.Exception.Message). Die Modulinstallation könnte fehlschlagen." -Type "Error"
                    Show-MessageBox -Message "Konnte PowerShellGet nicht aktualisieren. Dies kann zu Problemen bei der Installation anderer Module führen. Bitte stellen Sie sicher, dass PowerShellGet aktuell ist und versuchen Sie es ggf. mit Administratorrechten erneut.`nFehler: $($_.Exception.Message)" -Title "PowerShellGet Fehler" -Icon Warning
                    # Fortfahren trotz Fehler, da die Hauptmodule möglicherweise trotzdem installiert werden können, wenn PowerShellGet zumindest vorhanden ist.
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
                        Install-Module -Name $moduleName -Force -AllowClobber -MinimumVersion $minVersion -Scope CurrentUser
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
                    Install-Module -Name $moduleName -Force -AllowClobber -Scope CurrentUser
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
                Write-Log  "Fehler beim Installieren/Aktualisieren von $moduleName - $errorMsg. Administratorrechte könnten erforderlich sein." -Type "Error"
                
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
            $resultText += "Wenn Fehler aufgrund fehlender Berechtigungen aufgetreten sind, starten Sie PowerShell bitte mit Administratorrechten und versuchen Sie es erneut."
            
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

function Show-HelpDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Topic
    )

    try {
        $helpTitle = "Hilfe: $Topic"
        $helpMessage = ""

        switch ($Topic) {
            "Calendar" {
                $helpMessage = "Hier finden Sie Hilfe zum Verwalten von Kalenderberechtigungen.`n`n"
                $helpMessage += "Funktionen:`n"
                $helpMessage += "- Postfach angeben: Geben Sie die E-Mail-Adresse des Postfachs ein, dessen Kalenderberechtigungen Sie verwalten möchten.`n"
                $helpMessage += "- Anzeigen: Zeigt die aktuellen Kalenderberechtigungen für das angegebene Postfach an.`n"
                $helpMessage += "- Benutzer und Zugriffsebene auswählen/eingeben: Wählen Sie einen Benutzer aus der Liste oder geben Sie dessen E-Mail-Adresse ein. Wählen Sie die gewünschte Zugriffsebene (z.B. Editor, Reviewer).`n"
                $helpMessage += "- Hinzufügen: Fügt dem angegebenen Benutzer die ausgewählte Berechtigungsstufe für den Kalender hinzu.`n"
                $helpMessage += "- Ändern: Modifiziert die vorhandene Berechtigungsstufe des ausgewählten Benutzers auf die neu ausgewählte Zugriffsebene.`n"
                $helpMessage += "- Entfernen: Löscht die Kalenderberechtigungen des ausgewählten Benutzers.`n"
                $helpMessage += "- Alle setzen: Ermöglicht das Setzen einer Standardberechtigung (z.B. Verfügbarkeit) für alle Benutzer (außer 'Default' und 'Anonymous'). Bestehende spezifische Berechtigungen bleiben erhalten oder können optional überschrieben werden.`n"
                $helpMessage += "- Exportieren: Exportiert die aktuell angezeigten Kalenderberechtigungen in eine CSV-Datei.`n`n"
                $helpMessage += "Hinweis: 'Default' bezieht sich auf alle authentifizierten Benutzer in Ihrer Organisation. 'Anonymous' bezieht sich auf externe, nicht authentifizierte Benutzer."
            }
            "Mailbox" {
                $helpMessage = "Hier finden Sie Hilfe zum Verwalten von Postfachberechtigungen (Vollzugriff), 'Senden als'-Rechten und 'Senden im Auftrag von'-Rechten.`n`n"
                $helpMessage += "Eingabefelder:`n"
                $helpMessage += "- Ziel-Postfach: Das Postfach, für das Berechtigungen erteilt oder angezeigt werden sollen.`n"
                $helpMessage += "- Benutzer-Postfach: Das Postfach des Benutzers, der die Berechtigungen erhalten oder dem sie entzogen werden sollen.`n`n"
                $helpMessage += "Postfachberechtigungen (Vollzugriff):`n"
                $helpMessage += "- Hinzufügen: Gewährt dem 'Benutzer-Postfach' Vollzugriff auf das 'Ziel-Postfach'.`n"
                $helpMessage += "- Entfernen: Entzieht dem 'Benutzer-Postfach' den Vollzugriff auf das 'Ziel-Postfach'.`n"
                $helpMessage += "- Anzeigen: Listet alle Benutzer auf, die Vollzugriff auf das 'Ziel-Postfach' haben.`n`n"
                $helpMessage += "'Senden als'-Rechte:`n"
                $helpMessage += "- Hinzufügen: Erlaubt dem 'Benutzer-Postfach', E-Mails so zu senden, als kämen sie direkt vom 'Ziel-Postfach'.`n"
                $helpMessage += "- Entfernen: Entzieht dem 'Benutzer-Postfach' die 'Senden als'-Rechte für das 'Ziel-Postfach'.`n"
                $helpMessage += "- Anzeigen: Listet alle Benutzer auf, die 'Senden als'-Rechte für das 'Ziel-Postfach' haben.`n`n"
                $helpMessage += "'Senden im Auftrag von'-Rechte:`n"
                $helpMessage += "- Hinzufügen: Erlaubt dem 'Benutzer-Postfach', E-Mails im Auftrag des 'Ziel-Postfachs' zu senden (Empfänger sehen 'Benutzer A im Auftrag von Benutzer B').`n"
                $helpMessage += "- Entfernen: Entzieht dem 'Benutzer-Postfach' die 'Senden im Auftrag von'-Rechte für das 'Ziel-Postfach'.`n"
                $helpMessage += "- Anzeigen: Listet alle Benutzer auf, die 'Senden im Auftrag von'-Rechte für das 'Ziel-Postfach' haben."
            }
            "Contacts" {
                $helpMessage = "Hier finden Sie Hilfe zum Verwalten von externen Kontakten (MailContacts) und E-Mail-aktivierten Benutzern (MailUsers).`n`n"
                $helpMessage += "Funktionen:`n"
                $helpMessage += "- Neuer Kontakt:`n"
                $helpMessage += "  - Name: Der Anzeigename des Kontakts.`n"
                $helpMessage += "  - E-Mail: Die externe E-Mail-Adresse des Kontakts.`n"
                $helpMessage += "  - Erstellen: Legt einen neuen externen Kontakt (MailContact) an.`n`n"
                $helpMessage += "- Anzeigen (MailContacts): Listet alle externen Kontakte in Ihrer Organisation auf.`n"
                $helpMessage += "- Anzeigen (MailUsers): Listet alle E-Mail-aktivierten Benutzer auf (Benutzer mit Postfächern in Ihrer lokalen Active Directory-Umgebung, die mit Exchange Online synchronisiert werden, aber kein Exchange Online-Postfach haben).`n"
                $helpMessage += "- Ausgewählten Kontakt/MailUser entfernen: Löscht den in der Liste ausgewählten Kontakt oder MailUser. Eine Bestätigung ist erforderlich.`n"
                $helpMessage += "- Kontakte exportieren: Exportiert die aktuell in der Liste angezeigten Kontakte oder MailUser in eine CSV-Datei."
            }
            "Resources" {
                $helpMessage = "Hier finden Sie Hilfe zum Verwalten von Ressourceneinstellungen für Raum- und Gerätepostfächer.`n`n"
                $helpMessage += "Funktionen:`n"
                $helpMessage += "- Raum-Postfächer anzeigen: Listet alle konfigurierten Raumpostfächer auf.`n"
                $helpMessage += "- Geräte-Postfächer anzeigen: Listet alle konfigurierten Gerätepostfächer auf.`n"
                $helpMessage += "- Ausgewählte Ressource bearbeiten: Öffnet einen Dialog zur Anpassung spezifischer Einstellungen für die in der Liste ausgewählte Ressource. Dazu gehören unter anderem:`n"
                $helpMessage += "  - Anzeigename, Kapazität, Standort`n"
                $helpMessage += "  - Automatische Annahme/Ablehnung von Besprechungsanfragen`n"
                $helpMessage += "  - Zulassen von Konflikten und Serienbesprechungen`n"
                $helpMessage += "  - Buchungsfenster (wie weit im Voraus gebucht werden kann)`n"
                $helpMessage += "  - Maximale Besprechungsdauer`n"
                $helpMessage += "  - Verarbeitung von Anfragen außerhalb der Arbeitszeiten`n"
                $helpMessage += "  - Löschen von Kommentaren, Betreffzeilen oder privaten Kennzeichnungen`n"
                $helpMessage += "  - Hinzufügen des Organisators zum Betreff."
            }
            "Groups" {
                $helpMessage = "Hier finden Sie Hilfe zum Verwalten von Verteilergruppen und Microsoft 365-Gruppen.`n`n"
                $helpMessage += "Verteilergruppen:`n"
                $helpMessage += "- Anzeigen: Listet alle Verteilergruppen auf.`n"
                $helpMessage += "- Neu: Erstellt eine neue Verteilergruppe.`n"
                $helpMessage += "  - Name, Alias, Primäre SMTP-Adresse, Typ (Distribution/Security), Beitritts-/Verlassensoptionen.`n"
                $helpMessage += "- Bearbeiten: Ändert Eigenschaften der ausgewählten Verteilergruppe.`n"
                $helpMessage += "- Mitglieder verwalten: Hinzufügen/Entfernen von Mitgliedern.`n"
                $helpMessage += "- Besitzer verwalten: Hinzufügen/Entfernen von Besitzern.`n"
                $helpMessage += "- Löschen: Entfernt die ausgewählte Verteilergruppe.`n`n"
                $helpMessage += "Microsoft 365-Gruppen:`n"
                $helpMessage += "- Anzeigen: Listet alle Microsoft 365-Gruppen auf.`n"
                $helpMessage += "- Neu: Erstellt eine neue Microsoft 365-Gruppe.`n"
                $helpMessage += "  - Name, Alias, Beschreibung, Datenschutz (Öffentlich/Privat), Sprache, Besitzer, Mitglieder.`n"
                $helpMessage += "- Bearbeiten: Ändert Eigenschaften der ausgewählten Microsoft 365-Gruppe.`n"
                $helpMessage += "- Mitglieder verwalten: Hinzufügen/Entfernen von Mitgliedern.`n"
                $helpMessage += "- Besitzer verwalten: Hinzufügen/Entfernen von Besitzern.`n"
                $helpMessage += "- Löschen: Entfernt die ausgewählte Microsoft 365-Gruppe.`n`n"
                $helpMessage += "Exportieren: Exportiert die angezeigte Liste der Gruppen in eine CSV-Datei."
            }
            "General" {
                 $helpTitle = "Allgemeine Hilfe zu easyEXO"
                 $helpMessage = "Willkommen bei easyEXO! Dieses Tool wurde entwickelt, um die Verwaltung gängiger Aufgaben in Exchange Online über eine grafische Benutzeroberfläche zu vereinfachen.`n`n"
                 $helpMessage += "Hauptfunktionen und Tabs:`n"
                 $helpMessage += "- Verbindung: Bevor Sie Aktionen ausführen können, müssen Sie eine Verbindung zu Ihrem Exchange Online Tenant herstellen. Klicken Sie auf 'Verbinden' und geben Sie Ihre Administrator-Anmeldeinformationen ein.`n"
                 $helpMessage += "- Postfächer: Verwalten Sie Vollzugriffsberechtigungen, 'Senden als'-Rechte und 'Senden im Auftrag von'-Rechte für Benutzerpostfächer.`n"
                 $helpMessage += "- Kalender: Verwalten Sie die Freigabeberechtigungen für Benutzerkalender.`n"
                 $helpMessage += "- Kontakte: Erstellen und verwalten Sie externe Kontakte (MailContacts) und E-Mail-aktivierte Benutzer (MailUsers).`n"
                 $helpMessage += "- Ressourcen: Zeigen Sie Raum- und Gerätepostfächer an und bearbeiten Sie deren spezifische Buchungseinstellungen.`n"
                 $helpMessage += "- Gruppen: Verwalten Sie Verteilergruppen und Microsoft 365-Gruppen, deren Mitglieder und Besitzer.`n`n"
                 $helpMessage += "Bedienung:`n"
                 $helpMessage += "- Verwenden Sie die jeweiligen Tabs, um auf die spezifischen Verwaltungsfunktionen zuzugreifen.`n"
                 $helpMessage += "- Statusmeldungen und detaillierte Log-Informationen werden im unteren Bereich der Anwendung angezeigt.`n"
                 $helpMessage += "- Viele Listen können durch Klicken auf die Spaltenüberschriften sortiert werden.`n"
                 $helpMessage += "- Exportfunktionen sind oft verfügbar, um Daten als CSV-Datei zu sichern.`n"
                 $helpMessage += "- Hilfe-Symbole (?) in den Tabs bieten kontextspezifische Unterstützung.`n`n"
                 $helpMessage += "Stellen Sie sicher, dass die erforderlichen PowerShell-Module (insbesondere 'ExchangeOnlineManagement') installiert sind. Das Tool versucht, diese bei Bedarf zu installieren."
            }
            default {
                $helpMessage = "Kein spezifisches Hilfethema für '$Topic' gefunden.`n`n"
                $helpMessage += "Verfügbare Hilfethemen sind: General, Mailbox, Calendar, Contacts, Resources, Groups.`n"
                $helpMessage += "Bitte klicken Sie auf ein Hilfe-Symbol in einem der Tabs, um spezifische Informationen zu erhalten, oder wählen Sie 'General' für einen Überblick."
                $helpTitle = "Hilfe: Unbekanntes Thema"
            }
        }

        [System.Windows.MessageBox]::Show($helpMessage, $helpTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    }
    catch {
        $errorMsg = $_.Exception.Message
        # Versuchen, Write-Log aufzurufen, falls es definiert ist
        try {
            Write-Log -Message "Fehler im Show-HelpDialog für Topic '$Topic': $errorMsg" -Type "Error"
        } catch {}
        [System.Windows.MessageBox]::Show("Fehler beim Anzeigen der Hilfe für '$Topic': $errorMsg", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
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

function Test-PowerShell7AndAdminRights {
    [CmdletBinding()]
    param()

    try {
        $psVersion = $PSVersionTable.PSVersion
        $isPSCore = $psVersion.Major -ge 7
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $scriptPath = $MyInvocation.MyCommand.Path
        $currentPSEnginePath = (Get-Process -Id $PID).Path # Pfad zur aktuellen powershell.exe oder pwsh.exe

        Write-LogEntry "Aktueller Status: PowerShell Version $($psVersion.ToString()), Läuft als Administrator: $isAdmin"

        # Idealfall: Bereits PS7+ und Admin
        if ($isPSCore -and $isAdmin) {
            Write-LogEntry "Optimale Bedingungen (PowerShell 7+ und Administratorrechte) sind erfüllt."
            return $true
        }

        # PowerShell 7 Pfad suchen (nur wenn nicht bereits PS7+ und $ps7ExecutablePath noch nicht gesetzt)
        $ps7ExecutablePath = $null
        if (-not $isPSCore) {
            $ps7SearchPaths = @(
                Join-Path $env:ProgramFiles "PowerShell\7\pwsh.exe"
                Join-Path $env:ProgramFiles "(x86)\PowerShell\7\pwsh.exe"
                Join-Path $env:LOCALAPPDATA "Programs\PowerShell\7\pwsh.exe"
                # Versuche, pwsh.exe aus dem PATH zu finden, das nicht die Windows PowerShell ist
                (Get-Command pwsh -ErrorAction SilentlyContinue | Where-Object { $_.Source -and $_.Source -notlike "*\System32\WindowsPowerShell\*" -and $_.Source -notlike "*\SysWOW64\WindowsPowerShell\*" } | Select-Object -ExpandProperty Source -First 1)
            )
            $ps7ExecutablePath = $ps7SearchPaths | Where-Object { $_ -ne $null -and (Test-Path $_ -PathType Leaf) } | Select-Object -First 1
            
            if ($ps7ExecutablePath) {
                Write-LogEntry "PowerShell 7 gefunden unter: $ps7ExecutablePath"
            } else {
                Write-LogEntry "PowerShell 7 wurde auf dem System nicht gefunden."
            }
        }

        # Bedingungen für Neustart oder Installation
        $needsAdminPrivileges = (-not $isAdmin)
        $needsPS7Upgrade = (-not $isPSCore -and $null -ne $ps7ExecutablePath) # PS7 ist da, aber wir nutzen es nicht
        $needsPS7Installation = (-not $isPSCore -and $null -eq $ps7ExecutablePath) # PS7 ist nicht da und wir nutzen es nicht

        # Fall 1: PowerShell 7 muss installiert werden
        if ($needsPS7Installation) {
            $installMsg = "PowerShell 7 wird für dieses Skript empfohlen, wurde aber nicht gefunden."
            if ($needsAdminPrivileges) {
                $installMsg += " Zusätzlich sind Administratorrechte für einige Operationen und die Installation erforderlich."
            }
            $installMsg += " Möchten Sie PowerShell 7 jetzt installieren?"
            
            $userChoiceInstall = [System.Windows.MessageBox]::Show($installMsg, "PowerShell 7 Installation", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            
            if ($userChoiceInstall -eq [System.Windows.MessageBoxResult]::No) {
                Write-LogEntry "Benutzer hat die Installation von PowerShell 7 abgelehnt."
                [System.Windows.MessageBox]::Show("Ohne PowerShell 7 und/oder Administratorrechte können einige Funktionen des Skripts eingeschränkt sein oder fehlschlagen.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return $false 
            }

            Write-LogEntry "Benutzer hat der Installation von PowerShell 7 zugestimmt."
            $useWinget = $false
            try {
                $null = winget --version # Einfacher Test, ob winget existiert und funktioniert
                if ($LASTEXITCODE -eq 0) { $useWinget = $true }
            } catch { $useWinget = $false }

            if ($useWinget) {
                Write-LogEntry "Versuche PowerShell 7 Installation via winget."
                # Winget benötigt Admin für systemweite Installation. Start-Process mit RunAs für winget selbst.
                try {
                    Start-Process -FilePath "winget" -ArgumentList "install Microsoft.PowerShell --accept-package-agreements --accept-source-agreements" -Verb RunAs -Wait
                    Write-LogEntry "Winget-Installation von PowerShell 7 abgeschlossen (oder versucht)."
                } catch {
                     Write-LogEntry "Fehler beim Starten der Winget-Installation als Admin: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show("Fehler beim Starten der Winget-Installation für PowerShell 7: $($_.Exception.Message)`nVersuchen Sie, die Installation manuell als Administrator durchzuführen.", "Installationsfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return $false
                }
            } else {
                Write-LogEntry "Winget nicht verfügbar. Versuche PowerShell 7 Installation via MSI."
                $installerUrl = "https://github.com/PowerShell/PowerShell/releases/download/v7.4.2/PowerShell-7.4.2-win-x64.msi" # Aktuelle LTS Version
                $installerPath = Join-Path $env:TEMP "PowerShell-latest-win-x64.msi"
                try {
                    Write-LogEntry "Downloade PowerShell 7 MSI von $installerUrl"
                    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing
                    Write-LogEntry "PowerShell 7 MSI heruntergeladen nach $installerPath. Starte Installation."
                    # MSI Installation benötigt Admin. Start-Process mit RunAs.
                    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$installerPath`" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1" -Verb RunAs -Wait
                    Write-LogEntry "MSI-Installation von PowerShell 7 abgeschlossen (oder versucht)."
                } catch {
                    Write-LogEntry "FEHLER bei MSI Download/Installation: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show("Fehler beim Herunterladen oder Installieren von PowerShell 7 via MSI: $($_.Exception.Message)`nVersuchen Sie, die Installation manuell als Administrator durchzuführen.", "Installationsfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return $false
                } finally {
                    if (Test-Path $installerPath) { Remove-Item -Path $installerPath -Force }
                }
            }
            
            # Pfad nach Installation erneut suchen
            $ps7SearchPathsAfterInstall = @(
                Join-Path $env:ProgramFiles "PowerShell\7\pwsh.exe"
                Join-Path $env:ProgramFiles "(x86)\PowerShell\7\pwsh.exe"
                Join-Path $env:LOCALAPPDATA "Programs\PowerShell\7\pwsh.exe"
                (Get-Command pwsh -ErrorAction SilentlyContinue | Where-Object { $_.Source -and $_.Source -notlike "*\System32\WindowsPowerShell\*" -and $_.Source -notlike "*\SysWOW64\WindowsPowerShell\*" } | Select-Object -ExpandProperty Source -First 1)
            )
            $ps7ExecutablePath = $ps7SearchPathsAfterInstall | Where-Object { $_ -ne $null -and (Test-Path $_ -PathType Leaf) } | Select-Object -First 1

            if (-not $ps7ExecutablePath) {
                Write-LogEntry "PowerShell 7 konnte nach der Installation nicht gefunden werden."
                [System.Windows.MessageBox]::Show("PowerShell 7 wurde installiert, konnte aber nicht automatisch gefunden werden. Bitte starten Sie das Skript manuell mit PowerShell 7 (und ggf. Administratorrechten).", "Installationshinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return $false
            }
            Write-LogEntry "PowerShell 7 nach Installation gefunden unter: $ps7ExecutablePath"
            $needsPS7Upgrade = $true # Da wir gerade installiert haben, wollen wir es auch nutzen.
            $needsPS7Installation = $false # Nicht mehr relevant
        }

        # Fall 2: Neustart erforderlich (für Admin-Rechte und/oder PS7-Upgrade)
        # Die Bedingung ($needsAdminPrivileges -or $needsPS7Upgrade) deckt folgende Fälle ab:
        # 1. PowerShell < 7 ohne Admin-Rechte: $needsAdminPrivileges ist true, $needsPS7Upgrade ist true (falls PS7 gefunden). Neustart mit PS7 und Admin.
        # 2. PowerShell < 7 mit Admin-Rechten: $needsAdminPrivileges ist false, $needsPS7Upgrade ist true (falls PS7 gefunden). Neustart mit PS7 und Admin.
        # 3. PowerShell 7 ohne Admin-Rechte: $needsAdminPrivileges ist true, $needsPS7Upgrade ist false. Neustart mit aktueller PS7 und Admin. (Dies erfüllt die Anforderung)
        if ($needsAdminPrivileges -or $needsPS7Upgrade) {
            $restartMsgParts = @()
            $targetExecutableForRestart = $currentPSEnginePath # Standard: aktuelle PS-Engine

            if ($needsPS7Upgrade) { # PS7 ist verfügbar (oder gerade installiert) und wir sind nicht in PS7
                $restartMsgParts += "PowerShell 7"
                $targetExecutableForRestart = $ps7ExecutablePath
            }
            if ($needsAdminPrivileges) {
                $restartMsgParts += "Administratorrechten"
            }
            
            $reasonForRestart = $restartMsgParts -join " und "
            $currentContextDesc = "Sie verwenden derzeit PowerShell $($psVersion.Major).$($psVersion.Minor)"
            if ($isAdmin) { $currentContextDesc += " mit Administratorrechten." } else { $currentContextDesc += " ohne Administratorrechte."}

            $restartQueryMsg = "$currentContextDesc Für optimale Funktionalität wird ein Neustart mit $reasonForRestart empfohlen."
            if ($needsPS7Upgrade -and $ps7ExecutablePath) {
                 $restartQueryMsg += " PowerShell 7 ist unter '$ps7ExecutablePath' verfügbar."
            }
            $restartQueryMsg += " Möchten Sie das Skript jetzt neu starten?"

            $userChoiceRestart = [System.Windows.MessageBox]::Show($restartQueryMsg, "Neustart empfohlen", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)

            if ($userChoiceRestart -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-LogEntry "Benutzer stimmt Neustart zu. Ziel-Executable: $targetExecutableForRestart, Admin-Rechte werden angefordert."
                $argumentsForRestart = "-File `"$scriptPath`""
                try {
                    Start-Process -FilePath $targetExecutableForRestart -ArgumentList $argumentsForRestart -Verb RunAs
                    Write-LogEntry "Neustart-Prozess wurde initiiert."
                    exit # Aktuelles Skript beenden, da der neue Prozess gestartet wird
                } catch {
                    $errMsg = $_.Exception.Message
                    Write-LogEntry "FEHLER beim Versuch, das Skript neu zu starten: $errMsg"
                    [System.Windows.MessageBox]::Show("Fehler beim Versuch, das Skript neu zu starten: $errMsg", "Neustartfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return $false # Neustart ist fehlgeschlagen
                }
            } else {
                Write-LogEntry "Benutzer hat den empfohlenen Neustart abgelehnt."
                [System.Windows.MessageBox]::Show("Ohne die empfohlenen Einstellungen (PowerShell 7 und/oder Administratorrechte) können einige Funktionen des Skripts eingeschränkt sein oder fehlschlagen.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return $false 
            }
        }

        # Wenn wir hier ankommen, bedeutet das, dass entweder die Bedingungen initial nicht optimal waren
        # und der Benutzer die Korrekturmaßnahmen (Installation/Neustart) abgelehnt hat.
        Write-LogEntry "Die optimalen Ausführungsbedingungen wurden nicht erreicht oder vom Benutzer abgelehnt."
        return $false
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-LogEntry "Ein schwerwiegender Fehler ist in der Funktion Test-PowerShell7AndAdminRights aufgetreten: $errorMsg"
        Write-LogEntry "KRITISCHER FEHLER (Test-PowerShell7AndAdminRights): Ein interner Fehler ist bei der Überprüfung der Ausführungsumgebung aufgetreten: $errorMsg. Das Skript wird möglicherweise nicht korrekt funktionieren."
        return $false # Im Falle eines unerwarteten Fehlers in der Funktion selbst
    }
}

# Check for PowerShell 7 at startup
Test-PowerShell7AndAdminRights
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

# Globale Variable für gültige Datumsformate pro Kultur für Exchange Online
$script:ExchangeValidDateFormats = @{
    "de-DE" = @( # Deutsch (Deutschland)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" },
        @{ Display = "T.M.JJJJ"; Value = "d.M.yyyy" },
        @{ Display = "TT.MM.JJ"; Value = "dd.MM.yy" },
        @{ Display = "T.M.JJ"; Value = "d.M.yy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "de-AT" = @( # Deutsch (Österreich)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" },
        @{ Display = "T.M.JJJJ"; Value = "d.M.yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "de-CH" = @( # Deutsch (Schweiz)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" },
        @{ Display = "T.M.JJJJ"; Value = "d.M.yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "en-US" = @( # Englisch (USA)
        @{ Display = "MM/TT/JJJJ (Standard)"; Value = "MM/dd/yyyy" },
        @{ Display = "M/T/JJJJ"; Value = "M/d/yyyy" },
        @{ Display = "MM/TT/JJ"; Value = "MM/dd/yy" },
        @{ Display = "M/T/JJ"; Value = "M/d/yy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "en-GB" = @( # Englisch (UK) - Gemäß Ihrer Fehlermeldung
        @{ Display = "TT/MM/JJJJ (Standard)"; Value = "dd/MM/yyyy" },
        @{ Display = "TT/MM/JJ"; Value = "dd/MM/yy" },
        @{ Display = "T/M/JJ"; Value = "d/M/yy" }, # Angepasst an Fehlermeldung (war d/M/yyyy)
        @{ Display = "T.M.JJ"; Value = "d.M.yy" },   # Hinzugefügt gemäß Fehlermeldung
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "fr-FR" = @( # Französisch (Frankreich)
        @{ Display = "TT/MM/JJJJ (Standard)"; Value = "dd/MM/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "fr-CA" = @( # Französisch (Kanada) - Oft JJJJ-MM-TT bevorzugt
        @{ Display = "JJJJ-MM-TT (Standard)"; Value = "yyyy-MM-dd" },
        @{ Display = "TT/MM/JJJJ"; Value = "dd/MM/yyyy" }
    );
    "fr-CH" = @( # Französisch (Schweiz)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" }, # Punkte statt Schrägstriche
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "it-IT" = @( # Italienisch (Italien)
        @{ Display = "TT/MM/JJJJ (Standard)"; Value = "dd/MM/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "it-CH" = @( # Italienisch (Schweiz)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" }, # Punkte statt Schrägstriche
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "pl-PL" = @( # Polnisch (Polen)
        @{ Display = "TT.MM.JJJJ (Standard)"; Value = "dd.MM.yyyy" }, # Punkte als Trennzeichen
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "es-ES" = @( # Spanisch (Spanien)
        @{ Display = "TT/MM/JJJJ (Standard)"; Value = "dd/MM/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "es-MX" = @( # Spanisch (Mexiko) - Ähnlich wie Spanien
        @{ Display = "TT/MM/JJJJ (Standard)"; Value = "dd/MM/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "nl-NL" = @( # Niederländisch (Niederlande)
        @{ Display = "T-M-JJJJ (Standard)"; Value = "d-M-yyyy" },
        @{ Display = "TT-MM-JJJJ"; Value = "dd-MM-yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    "nl-BE" = @( # Niederländisch (Belgien)
        @{ Display = "T/M/JJJJ (Standard)"; Value = "d/M/yyyy" }, # Schrägstriche in Belgien
        @{ Display = "TT/MM/JJJJ"; Value = "dd/MM/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    );
    # Fallback / Standard, falls eine Kultur nicht spezifisch abgedeckt ist
    "DEFAULT" = @(
        @{ Display = "Systemstandard (Keine explizite Auswahl)"; Value = "" }, # Leerer Value für keine Änderung
        @{ Display = "TT.MM.JJJJ"; Value = "dd.MM.yyyy" },
        @{ Display = "MM/TT/JJJJ"; Value = "MM/dd/yyyy" },
        @{ Display = "JJJJ-MM-TT (ISO)"; Value = "yyyy-MM-dd" }
    )
}
# Globale Variable für relevante Zeitformate pro Kultur für Exchange Online
$script:RelevantTimeFormatsPerCulture = @{
    "de-DE" = @( # Deutsch (Deutschland)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" }
    );
    "de-AT" = @( # Deutsch (Österreich)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" }
    );
    "de-CH" = @( # Deutsch (Schweiz) - Oft mit Punkt
        @{ Display = "HH.mm (24h, Standard)"; Value = "HH.mm" },
        @{ Display = "H.mm (24h)"; Value = "H.mm" }
    );
    "en-US" = @( # Englisch (USA)
        @{ Display = "h:mm tt (12h AM/PM, Standard)"; Value = "h:mm tt" },
        @{ Display = "hh:mm tt (12h AM/PM)"; Value = "hh:mm tt" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" },
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" }
    );
    "en-GB" = @( # Englisch (UK)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" },
        @{ Display = "h:mm tt (12h AM/PM)"; Value = "h:mm tt" }
    );
    "fr-FR" = @( # Französisch (Frankreich)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" }
    );
    "fr-CA" = @( # Französisch (Kanada)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" }, # Oft 24h bevorzugt
        @{ Display = "h:mm tt (12h AM/PM)"; Value = "h:mm tt" }
    );
    "fr-CH" = @( # Französisch (Schweiz) - Oft mit Punkt
        @{ Display = "HH.mm (24h, Standard)"; Value = "HH.mm" },
        @{ Display = "H.mm (24h)"; Value = "H.mm" }
    );
    "it-IT" = @( # Italienisch (Italien)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" }
    );
    "it-CH" = @( # Italienisch (Schweiz) - Oft mit Punkt
        @{ Display = "HH.mm (24h, Standard)"; Value = "HH.mm" },
        @{ Display = "H.mm (24h)"; Value = "H.mm" }
    );
    "pl-PL" = @( # Polnisch (Polen)
        @{ Display = "HH:mm (24h, Standard)"; Value = "HH:mm" },
        @{ Display = "H:mm (24h)"; Value = "H:mm" }
    );
    "es-ES" = @( # Spanisch (Spanien)
        @{ Display = "H:mm (24h, Standard)"; Value = "H:mm" }, # Kann auch HH:mm sein
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" }
    );
    "es-MX" = @( # Spanisch (Mexiko)
        @{ Display = "h:mm tt (12h AM/PM, Standard)"; Value = "h:mm tt" },
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" }
    );
    "nl-NL" = @( # Niederländisch (Niederlande)
        @{ Display = "H:mm (24h, Standard)"; Value = "H:mm" }, # Oder HH:mm
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" }
    );
    "nl-BE" = @( # Niederländisch (Belgien)
        @{ Display = "H:mm (24h, Standard)"; Value = "H:mm" }, # Oder HH:mm
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" }
    );
    "DEFAULT" = @( # Allgemeine Fallbacks, wenn keine spezifische Kultur passt
        @{ Display = "Systemstandard (Keine explizite Auswahl)"; Value = "" },
        @{ Display = "HH:mm (24h)"; Value = "HH:mm" },
        @{ Display = "h:mm tt (12h AM/PM)"; Value = "h:mm tt" }
    )
}
# Globale Variable für relevante Zeitzonen-IDs pro Kultur für Exchange Online
$script:RelevantTimezonesPerCulture = @{
    "de-DE" = @("W. Europe Standard Time", "Central European Standard Time"); # Berlin, Amsterdam, Paris, Rome
    "de-AT" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Vienna
    "de-CH" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Bern, Zurich
    "en-US" = @("Pacific Standard Time", "Mountain Standard Time", "Central Standard Time", "Eastern Standard Time", "Alaskan Standard Time", "Hawaiian Standard Time");
    "en-GB" = @("GMT Standard Time", "Greenwich Standard Time"); # London, Dublin
    "fr-FR" = @("Romance Standard Time", "Central European Standard Time"); # Paris
    "fr-CA" = @("Eastern Standard Time", "Central Standard Time", "Mountain Standard Time", "Pacific Standard Time", "Newfoundland Standard Time", "Atlantic Standard Time"); # Canada
    "fr-CH" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Geneva
    "it-IT" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Rome
    "it-CH" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Italian-speaking Switzerland
    "pl-PL" = @("Central European Standard Time"); # Warsaw
    "es-ES" = @("Romance Standard Time", "Central European Standard Time"); # Madrid
    "es-MX" = @("Central Standard Time (Mexico)", "Mountain Standard Time (Mexico)", "Pacific Standard Time (Mexico)"); # Mexico City, Chihuahua, Tijuana
    "nl-NL" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Amsterdam
    "nl-BE" = @("W. Europe Standard Time", "Central European Standard Time", "Romance Standard Time"); # Brussels
    # Fallback für nicht explizit gemappte Kulturen - hier könnten alle Zeitzonen geladen werden oder eine Auswahl häufiger
    "DEFAULT_ALL" = $true # Ein Flag, um alle Zeitzonen zu laden, wenn keine spezifische Kultur passt
    # Oder eine kleinere Default-Liste:
    # "DEFAULT" = @("UTC", "GMT Standard Time", "W. Europe Standard Time", "Central European Standard Time", "Eastern Standard Time", "Pacific Standard Time")
}
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
        Write-LogEntry -Message "$Title - $Type - $Message" -Type "Info"
        
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
# Registry-Pfad und Standardwerte für Konfigurationseinstellungen
$script:registryPath = "HKCU:\Software\easyIT\easyEXO"
$currentScriptVersion = "0.0.9" # Aktuelle Version des Skripts

try {
    Write-LogEntry "Prüfe und initialisiere Registry-Konfiguration unter '$($script:registryPath)'."

    # Stelle sicher, dass der Basispfad "HKCU:\Software\easyIT" existiert
    $parentPath = "HKCU:\Software\easyIT"
    if (-not (Test-Path -Path $parentPath)) {
        New-Item -Path "HKCU:\Software" -Name "easyIT" -Force -ErrorAction Stop | Out-Null
        Write-LogEntry "Registry-Basispfad '$parentPath' wurde erstellt."
    }

    # Prüfe, ob der Anwendungspfad existiert.
    $appPathExistedBefore = Test-Path -Path $script:registryPath
    if (-not $appPathExistedBefore) {
        # Wenn der Anwendungspfad nicht existiert, erstelle ihn.
        New-Item -Path $parentPath -Name (Split-Path $script:registryPath -Leaf) -Force -ErrorAction Stop | Out-Null
        Write-LogEntry "Registry-Anwendungspfad '$($script:registryPath)' wurde erstellt."
    }
    
    $performUpdate = $false
    
    if (-not $appPathExistedBefore) {
        # Wenn der Anwendungspfad gerade erst erstellt wurde, ist ein vollständiges Setzen der Standardwerte erforderlich.
        Write-LogEntry "Registry-Anwendungspfad war nicht vorhanden. Standardwerte werden initial geschrieben."
        $performUpdate = $true
    } else {
        # Der Pfad existierte bereits. Prüfe die Version.
        $storedVersion = $null
        try {
            # Versuche, den Versionseintrag zu lesen
            $storedVersionProperty = Get-ItemProperty -Path $script:registryPath -Name "Version" -ErrorAction SilentlyContinue
            if ($null -ne $storedVersionProperty -and $storedVersionProperty.PSObject.Properties["Version"]) {
                $storedVersion = $storedVersionProperty.Version
            }
        }
        catch {
            # Fehler beim Lesen der Version, sicherheitshalber Update durchführen
            Write-LogEntry "WARNUNG: Fehler beim Lesen der Version aus der Registry unter '$($script:registryPath)': $($_.Exception.Message). Standardwerte werden vorsichtshalber aktualisiert."
            $performUpdate = $true # Update erzwingen bei Lesefehler
        }

        if (-not $performUpdate) { # Nur prüfen, wenn nicht schon durch Fehler oben ein Update erzwungen wurde
            if ($null -eq $storedVersion) {
                Write-LogEntry "Kein Versionseintrag in der Registry gefunden unter '$($script:registryPath)' oder Wert ist null. Standardwerte werden geschrieben."
                $performUpdate = $true
            } elseif ($storedVersion -ne $currentScriptVersion) {
                Write-LogEntry "Registry-Version ('$storedVersion') unterscheidet sich von Skript-Version ('$currentScriptVersion'). Standardwerte werden aktualisiert."
                $performUpdate = $true
            } else {
                Write-LogEntry "Registry-Version ('$storedVersion') ist aktuell mit Skript-Version ('$currentScriptVersion'). Keine Aktualisierung der Standardwerte erforderlich."
            }
        }
    }
    
    if ($performUpdate) {
        Write-LogEntry "Setze/Aktualisiere Registry-Standardwerte für '$($script:registryPath)'."
        New-ItemProperty -Path $script:registryPath -Name "Debug" -Value 0 -PropertyType DWORD -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "AppName" -Value "Exchange Online Verwaltung" -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "Version" -Value $currentScriptVersion -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "ThemeColor" -Value "#0078D7" -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "LogPath" -Value "$PSScriptRoot\Logs" -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "HeaderLogoURL" -Value "https://psscripts.de" -PropertyType String -Force -ErrorAction Stop | Out-Null
        Write-LogEntry "Registry-Standardwerte für '$($script:registryPath)' erfolgreich gesetzt/aktualisiert."
    }
}
catch {
    $errorMsg = $_.Exception.Message
    Write-LogEntry "FEHLER bei der Registry-Konfigurationsinitialisierung für '$($script:registryPath)': $errorMsg"
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
                "Debug" = "0"
                "AppName" = "Exchange Online Verwaltung"
                "Version" = "0.0.9"
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
            Write-LogEntry  "Logverzeichnis wurde erstellt: $logFolder" -Type "Info" # Konsolenausgabe hiervon wird durch Write-Log gesteuert
        }
        
        # Log-Eintrag schreiben
        Add-Content -Path $script:logFilePath -Value "[$timestamp] $sanitizedMessage" -Encoding UTF8
        
        # Bei zu langer Logdatei (>10 MB) rotieren
        $logFile = Get-Item -Path $script:logFilePath -ErrorAction SilentlyContinue
        if ($logFile -and $logFile.Length -gt 10MB) {
            $backupLogPath = "$($script:logFilePath)_$(Get-Date -Format 'yyyyMMdd_HHmmss').bak"
            Move-Item -Path $script:logFilePath -Destination $backupLogPath -Force
            Write-LogEntry  "Logdatei wurde rotiert: $backupLogPath" -Type "Info" # Konsolenausgabe hiervon wird durch Write-Log gesteuert
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

        # Nachricht formatieren für Konsolenausgabe
        $formattedMessage = "[$timestamp] [$Type] $Message"

        # Ausgabe in Konsole, wenn nicht unterdrückt UND Debug-Modus aktiv ist
        if (-not $NoConsole -and $script:debugMode) {
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

        # Logging mit Log-Action, wenn nicht unterdrückt und Log-Pfad gesetzt ist
        if (-not $NoLog -and $script:logFilePath) {
            try {
                Log-Action -Message "[$Type] $Message"
            }
            catch {
                if ($script:debugMode) {
                    # Sanitize die Fehlermeldung für die Konsolenausgabe
                    $logActionCallError = $($_.Exception.Message) -replace '[^\x20-\x7E\r\n]', '?'
                    Write-Host "Fehler beim Aufruf von Log-Action innerhalb von Write-Log: $logActionCallError" -ForegroundColor Red
                }
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

            # Direkter Fallback zur Ausgabe ohne Farbe auf der Konsole, nur wenn Debug-Modus aktiv ist
            if ($script:debugMode) {
                Write-Host $errorMessage -ForegroundColor Red
            }

            # Versuch, den Fehler mit Log-Action zu loggen, falls möglich und Log-Pfad vorhanden
            if ($script:logFilePath) {
                try {
                    # Log-Action kümmert sich um Zeitstempel und Fehlerbehandlung beim Schreiben.
                    Log-Action -Message $errorMessage
                }
                catch {
                    if ($script:debugMode) {
                        $criticalLogActionError = $($_.Exception.Message) -replace '[^\x20-\x7E\r\n]', '?'
                        Write-Host "Kritischer Fehler: Log-Action konnte den Fehler in Write-Log nicht protokollieren. Fehler beim Aufruf von Log-Action: $criticalLogActionError" -ForegroundColor Red
                    }
                }
            }
        }
        catch {
        }
    }
}

# Funktion zur Initialisierung des Loggings für Log-Action
function Initialize-Logging {
    [CmdletBinding()]
    param()

    try {
        # Standard-Logverzeichnis definieren
        $defaultLogDirectory = Join-Path -Path $PSScriptRoot -ChildPath "Logs"
        $logDirectoryToUse = $defaultLogDirectory # Mit Standardwert beginnen

        # Versuchen, den Log-Pfad aus der Konfiguration zu laden
        if ($null -ne $script:config -and
            $script:config.ContainsKey("Paths") -and
            ($null -ne $script:config["Paths"]) -and # Prüfen, ob "Paths" selbst nicht $null ist
            $script:config["Paths"].ContainsKey("LogPath") -and
            -not [string]::IsNullOrWhiteSpace($script:config["Paths"]["LogPath"])) {

            $configuredLogPathDir = $script:config["Paths"]["LogPath"]

            # Überprüfen, ob der konfigurierte Pfad absolut oder relativ ist
            if ([System.IO.Path]::IsPathRooted($configuredLogPathDir)) {
                $logDirectoryToUse = [System.IO.Path]::GetFullPath($configuredLogPathDir)
            } else {
                # Wenn relativ, relativ zum Skriptverzeichnis auflösen
                $pathForNormalization = Join-Path -Path $PSScriptRoot -ChildPath $configuredLogPathDir
                $logDirectoryToUse = [System.IO.Path]::GetFullPath($pathForNormalization)
            }
        }

        # Log-Dateiname festlegen (dieser Name wird von Log-Action für die Rotation verwendet)
        $logFileName = "easyEXO_activity.log"
        $script:logFilePath = Join-Path -Path $logDirectoryToUse -ChildPath $logFileName

        Log-Action "Log-Action System initialisiert. Logdatei: $($script:logFilePath)"

        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($script:debugMode) {
            Write-Warning "FEHLER bei der Initialisierung des Log-Pfades für Log-Action: $errorMsg. Log-Action wird versuchen, den internen Fallback zu verwenden."
        }
        return $false
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
        Write-Log  -Message $Message -Type $Type # Konsolenausgabe hiervon wird durch Write-Log gesteuert
        
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
        Write-Log  "Fehler in Write-StatusMessage: $errorMsg" -Type "Error" # Konsolenausgabe hiervon wird durch Write-Log gesteuert
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
        Write-Log "Verbindungsversuch zu Exchange Online..." -Type "Info"
        
        # Prüfen, ob das ExchangeOnlineManagement Modul installiert ist
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            $errorMsg = "ExchangeOnlineManagement Modul ist nicht installiert. Bitte installieren Sie das Modul mit 'Install-Module ExchangeOnlineManagement -Force'"
            Write-Log $errorMsg -Type "Error"
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
            Write-Log $errorMsg -Type "Warning"
            Show-MessageBox -Message $errorMsg -Title "Abgebrochen" -Type "Warning"
            return $false
        }
        
        # Überprüfen, ob die E-Mail-Adresse erfolgreich gespeichert wurde
        if ([string]::IsNullOrWhiteSpace($script:userPrincipalName)) {
            $errorMsg = "Keine E-Mail-Adresse eingegeben oder erkannt. Verbindung abgebrochen."
            Write-Log $errorMsg -Type "Warning"
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
        
        Write-Log "Exchange Online Verbindung erfolgreich hergestellt für $script:userPrincipalName" -Type "Success"
        $script:txtConnectionStatus.Text = "Verbunden mit Exchange Online ($script:userPrincipalName)"
        $script:txtConnectionStatus.Foreground = "#008000"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Verbinden mit Exchange Online: $errorMsg" -Type "Error"
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
            
            # Versuche, PowerShellGet zu aktualisieren
            # Die Überprüfung auf Administratorrechte und der Neustart-Mechanismus wurden entfernt.
            # Es wird davon ausgegangen, dass das Skript bei Bedarf mit erhöhten Rechten ausgeführt wird.
            try {
                Write-Log "Versuche PowerShellGet zu aktualisieren/installieren. Administratorrechte könnten erforderlich sein." -Type "Info"
                Install-Module PowerShellGet -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                Write-Log  "PowerShellGet erfolgreich aktualisiert/installiert für den aktuellen Benutzer." -Type "Success"
            } 
            catch {
                Write-Log  "Fehler beim Aktualisieren/Installieren von PowerShellGet für den aktuellen Benutzer: $($_.Exception.Message). Versuche systemweite Installation." -Type "Warning"
                try {
                    Install-Module PowerShellGet -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
                    Write-Log  "PowerShellGet erfolgreich systemweit aktualisiert/installiert." -Type "Success"
                }
                catch {
                    Write-Log  "Fehler beim systemweiten Aktualisieren/Installieren von PowerShellGet: $($_.Exception.Message). Die Modulinstallation könnte fehlschlagen." -Type "Error"
                    Show-MessageBox -Message "Konnte PowerShellGet nicht aktualisieren. Dies kann zu Problemen bei der Installation anderer Module führen. Bitte stellen Sie sicher, dass PowerShellGet aktuell ist und versuchen Sie es ggf. mit Administratorrechten erneut.`nFehler: $($_.Exception.Message)" -Title "PowerShellGet Fehler" -Icon Warning
                    # Fortfahren trotz Fehler, da die Hauptmodule möglicherweise trotzdem installiert werden können, wenn PowerShellGet zumindest vorhanden ist.
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
                        Install-Module -Name $moduleName -Force -AllowClobber -MinimumVersion $minVersion -Scope CurrentUser
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
                    Install-Module -Name $moduleName -Force -AllowClobber -Scope CurrentUser
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
                Write-Log  "Fehler beim Installieren/Aktualisieren von $moduleName - $errorMsg. Administratorrechte könnten erforderlich sein." -Type "Error"
                
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
            $resultText += "Wenn Fehler aufgrund fehlender Berechtigungen aufgetreten sind, starten Sie PowerShell bitte mit Administratorrechten und versuchen Sie es erneut."
            
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

function Show-HelpDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Topic
    )

    try {
        $helpTitle = "Hilfe: $Topic"
        $helpMessage = ""

        switch ($Topic) {
            "Calendar" {
                $helpMessage = "Hier finden Sie Hilfe zum Verwalten von Kalenderberechtigungen.`n`n"
                $helpMessage += "Funktionen:`n"
                $helpMessage += "- Postfach angeben: Geben Sie die E-Mail-Adresse des Postfachs ein, dessen Kalenderberechtigungen Sie verwalten möchten.`n"
                $helpMessage += "- Anzeigen: Zeigt die aktuellen Kalenderberechtigungen für das angegebene Postfach an.`n"
                $helpMessage += "- Benutzer und Zugriffsebene auswählen/eingeben: Wählen Sie einen Benutzer aus der Liste oder geben Sie dessen E-Mail-Adresse ein. Wählen Sie die gewünschte Zugriffsebene (z.B. Editor, Reviewer).`n"
                $helpMessage += "- Hinzufügen: Fügt dem angegebenen Benutzer die ausgewählte Berechtigungsstufe für den Kalender hinzu.`n"
                $helpMessage += "- Ändern: Modifiziert die vorhandene Berechtigungsstufe des ausgewählten Benutzers auf die neu ausgewählte Zugriffsebene.`n"
                $helpMessage += "- Entfernen: Löscht die Kalenderberechtigungen des ausgewählten Benutzers.`n"
                $helpMessage += "- Alle setzen: Ermöglicht das Setzen einer Standardberechtigung (z.B. Verfügbarkeit) für alle Benutzer (außer 'Default' und 'Anonymous'). Bestehende spezifische Berechtigungen bleiben erhalten oder können optional überschrieben werden.`n"
                $helpMessage += "- Exportieren: Exportiert die aktuell angezeigten Kalenderberechtigungen in eine CSV-Datei.`n`n"
                $helpMessage += "Hinweis: 'Default' bezieht sich auf alle authentifizierten Benutzer in Ihrer Organisation. 'Anonymous' bezieht sich auf externe, nicht authentifizierte Benutzer."
            }
            "Mailbox" {
                $helpMessage = "Hier finden Sie Hilfe zum Verwalten von Postfachberechtigungen (Vollzugriff), 'Senden als'-Rechten und 'Senden im Auftrag von'-Rechten.`n`n"
                $helpMessage += "Eingabefelder:`n"
                $helpMessage += "- Ziel-Postfach: Das Postfach, für das Berechtigungen erteilt oder angezeigt werden sollen.`n"
                $helpMessage += "- Benutzer-Postfach: Das Postfach des Benutzers, der die Berechtigungen erhalten oder dem sie entzogen werden sollen.`n`n"
                $helpMessage += "Postfachberechtigungen (Vollzugriff):`n"
                $helpMessage += "- Hinzufügen: Gewährt dem 'Benutzer-Postfach' Vollzugriff auf das 'Ziel-Postfach'.`n"
                $helpMessage += "- Entfernen: Entzieht dem 'Benutzer-Postfach' den Vollzugriff auf das 'Ziel-Postfach'.`n"
                $helpMessage += "- Anzeigen: Listet alle Benutzer auf, die Vollzugriff auf das 'Ziel-Postfach' haben.`n`n"
                $helpMessage += "'Senden als'-Rechte:`n"
                $helpMessage += "- Hinzufügen: Erlaubt dem 'Benutzer-Postfach', E-Mails so zu senden, als kämen sie direkt vom 'Ziel-Postfach'.`n"
                $helpMessage += "- Entfernen: Entzieht dem 'Benutzer-Postfach' die 'Senden als'-Rechte für das 'Ziel-Postfach'.`n"
                $helpMessage += "- Anzeigen: Listet alle Benutzer auf, die 'Senden als'-Rechte für das 'Ziel-Postfach' haben.`n`n"
                $helpMessage += "'Senden im Auftrag von'-Rechte:`n"
                $helpMessage += "- Hinzufügen: Erlaubt dem 'Benutzer-Postfach', E-Mails im Auftrag des 'Ziel-Postfachs' zu senden (Empfänger sehen 'Benutzer A im Auftrag von Benutzer B').`n"
                $helpMessage += "- Entfernen: Entzieht dem 'Benutzer-Postfach' die 'Senden im Auftrag von'-Rechte für das 'Ziel-Postfach'.`n"
                $helpMessage += "- Anzeigen: Listet alle Benutzer auf, die 'Senden im Auftrag von'-Rechte für das 'Ziel-Postfach' haben."
            }
            "Contacts" {
                $helpMessage = "Hier finden Sie Hilfe zum Verwalten von externen Kontakten (MailContacts) und E-Mail-aktivierten Benutzern (MailUsers).`n`n"
                $helpMessage += "Funktionen:`n"
                $helpMessage += "- Neuer Kontakt:`n"
                $helpMessage += "  - Name: Der Anzeigename des Kontakts.`n"
                $helpMessage += "  - E-Mail: Die externe E-Mail-Adresse des Kontakts.`n"
                $helpMessage += "  - Erstellen: Legt einen neuen externen Kontakt (MailContact) an.`n`n"
                $helpMessage += "- Anzeigen (MailContacts): Listet alle externen Kontakte in Ihrer Organisation auf.`n"
                $helpMessage += "- Anzeigen (MailUsers): Listet alle E-Mail-aktivierten Benutzer auf (Benutzer mit Postfächern in Ihrer lokalen Active Directory-Umgebung, die mit Exchange Online synchronisiert werden, aber kein Exchange Online-Postfach haben).`n"
                $helpMessage += "- Ausgewählten Kontakt/MailUser entfernen: Löscht den in der Liste ausgewählten Kontakt oder MailUser. Eine Bestätigung ist erforderlich.`n"
                $helpMessage += "- Kontakte exportieren: Exportiert die aktuell in der Liste angezeigten Kontakte oder MailUser in eine CSV-Datei."
            }
            "Resources" {
                $helpMessage = "Hier finden Sie Hilfe zum Verwalten von Ressourceneinstellungen für Raum- und Gerätepostfächer.`n`n"
                $helpMessage += "Funktionen:`n"
                $helpMessage += "- Raum-Postfächer anzeigen: Listet alle konfigurierten Raumpostfächer auf.`n"
                $helpMessage += "- Geräte-Postfächer anzeigen: Listet alle konfigurierten Gerätepostfächer auf.`n"
                $helpMessage += "- Ausgewählte Ressource bearbeiten: Öffnet einen Dialog zur Anpassung spezifischer Einstellungen für die in der Liste ausgewählte Ressource. Dazu gehören unter anderem:`n"
                $helpMessage += "  - Anzeigename, Kapazität, Standort`n"
                $helpMessage += "  - Automatische Annahme/Ablehnung von Besprechungsanfragen`n"
                $helpMessage += "  - Zulassen von Konflikten und Serienbesprechungen`n"
                $helpMessage += "  - Buchungsfenster (wie weit im Voraus gebucht werden kann)`n"
                $helpMessage += "  - Maximale Besprechungsdauer`n"
                $helpMessage += "  - Verarbeitung von Anfragen außerhalb der Arbeitszeiten`n"
                $helpMessage += "  - Löschen von Kommentaren, Betreffzeilen oder privaten Kennzeichnungen`n"
                $helpMessage += "  - Hinzufügen des Organisators zum Betreff."
            }
            "Groups" {
                $helpMessage = "Hier finden Sie Hilfe zum Verwalten von Verteilergruppen und Microsoft 365-Gruppen.`n`n"
                $helpMessage += "Verteilergruppen:`n"
                $helpMessage += "- Anzeigen: Listet alle Verteilergruppen auf.`n"
                $helpMessage += "- Neu: Erstellt eine neue Verteilergruppe.`n"
                $helpMessage += "  - Name, Alias, Primäre SMTP-Adresse, Typ (Distribution/Security), Beitritts-/Verlassensoptionen.`n"
                $helpMessage += "- Bearbeiten: Ändert Eigenschaften der ausgewählten Verteilergruppe.`n"
                $helpMessage += "- Mitglieder verwalten: Hinzufügen/Entfernen von Mitgliedern.`n"
                $helpMessage += "- Besitzer verwalten: Hinzufügen/Entfernen von Besitzern.`n"
                $helpMessage += "- Löschen: Entfernt die ausgewählte Verteilergruppe.`n`n"
                $helpMessage += "Microsoft 365-Gruppen:`n"
                $helpMessage += "- Anzeigen: Listet alle Microsoft 365-Gruppen auf.`n"
                $helpMessage += "- Neu: Erstellt eine neue Microsoft 365-Gruppe.`n"
                $helpMessage += "  - Name, Alias, Beschreibung, Datenschutz (Öffentlich/Privat), Sprache, Besitzer, Mitglieder.`n"
                $helpMessage += "- Bearbeiten: Ändert Eigenschaften der ausgewählten Microsoft 365-Gruppe.`n"
                $helpMessage += "- Mitglieder verwalten: Hinzufügen/Entfernen von Mitgliedern.`n"
                $helpMessage += "- Besitzer verwalten: Hinzufügen/Entfernen von Besitzern.`n"
                $helpMessage += "- Löschen: Entfernt die ausgewählte Microsoft 365-Gruppe.`n`n"
                $helpMessage += "Exportieren: Exportiert die angezeigte Liste der Gruppen in eine CSV-Datei."
            }
            "General" {
                 $helpTitle = "Allgemeine Hilfe zu easyEXO"
                 $helpMessage = "Willkommen bei easyEXO! Dieses Tool wurde entwickelt, um die Verwaltung gängiger Aufgaben in Exchange Online über eine grafische Benutzeroberfläche zu vereinfachen.`n`n"
                 $helpMessage += "Hauptfunktionen und Tabs:`n"
                 $helpMessage += "- Verbindung: Bevor Sie Aktionen ausführen können, müssen Sie eine Verbindung zu Ihrem Exchange Online Tenant herstellen. Klicken Sie auf 'Verbinden' und geben Sie Ihre Administrator-Anmeldeinformationen ein.`n"
                 $helpMessage += "- Postfächer: Verwalten Sie Vollzugriffsberechtigungen, 'Senden als'-Rechte und 'Senden im Auftrag von'-Rechte für Benutzerpostfächer.`n"
                 $helpMessage += "- Kalender: Verwalten Sie die Freigabeberechtigungen für Benutzerkalender.`n"
                 $helpMessage += "- Kontakte: Erstellen und verwalten Sie externe Kontakte (MailContacts) und E-Mail-aktivierte Benutzer (MailUsers).`n"
                 $helpMessage += "- Ressourcen: Zeigen Sie Raum- und Gerätepostfächer an und bearbeiten Sie deren spezifische Buchungseinstellungen.`n"
                 $helpMessage += "- Gruppen: Verwalten Sie Verteilergruppen und Microsoft 365-Gruppen, deren Mitglieder und Besitzer.`n`n"
                 $helpMessage += "Bedienung:`n"
                 $helpMessage += "- Verwenden Sie die jeweiligen Tabs, um auf die spezifischen Verwaltungsfunktionen zuzugreifen.`n"
                 $helpMessage += "- Statusmeldungen und detaillierte Log-Informationen werden im unteren Bereich der Anwendung angezeigt.`n"
                 $helpMessage += "- Viele Listen können durch Klicken auf die Spaltenüberschriften sortiert werden.`n"
                 $helpMessage += "- Exportfunktionen sind oft verfügbar, um Daten als CSV-Datei zu sichern.`n"
                 $helpMessage += "- Hilfe-Symbole (?) in den Tabs bieten kontextspezifische Unterstützung.`n`n"
                 $helpMessage += "Stellen Sie sicher, dass die erforderlichen PowerShell-Module (insbesondere 'ExchangeOnlineManagement') installiert sind. Das Tool versucht, diese bei Bedarf zu installieren."
            }
            default {
                $helpMessage = "Kein spezifisches Hilfethema für '$Topic' gefunden.`n`n"
                $helpMessage += "Verfügbare Hilfethemen sind: General, Mailbox, Calendar, Contacts, Resources, Groups.`n"
                $helpMessage += "Bitte klicken Sie auf ein Hilfe-Symbol in einem der Tabs, um spezifische Informationen zu erhalten, oder wählen Sie 'General' für einen Überblick."
                $helpTitle = "Hilfe: Unbekanntes Thema"
            }
        }

        [System.Windows.MessageBox]::Show($helpMessage, $helpTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    }
    catch {
        $errorMsg = $_.Exception.Message
        # Versuchen, Write-Log aufzurufen, falls es definiert ist
        try {
            Write-Log -Message "Fehler im Show-HelpDialog für Topic '$Topic': $errorMsg" -Type "Error"
        } catch {}
        [System.Windows.MessageBox]::Show("Fehler beim Anzeigen der Hilfe für '$Topic': $errorMsg", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
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
    $mailboxInput = $script:txtRegionMailbox.Text.Trim()
    $selectedLanguageItem = $script:cmbRegionLanguage.SelectedItem
    $selectedTimezoneItem = $script:cmbRegionTimezone.SelectedItem
    $selectedDateFormatItem = $script:cmbRegionDateFormat.SelectedItem   # NEU
    $selectedTimeFormatItem = $script:cmbRegionTimeFormat.SelectedItem   # NEU
    $localizeFoldersState = $script:chkRegionDefaultFolderNameMatchingUserLanguage.IsChecked # NEU ($true, $false, oder $null)
    $statusTextBlock = $script:txtStatus

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
    # Wichtig: Wenn cmbRegionTimezone IsEditable="True", kann SelectedItem null sein, auch wenn Text vorhanden ist.
    # Wir verlassen uns hier auf den Tag, wenn ein Item ausgewählt ist.
    # Wenn der Benutzer Text eingibt, der keinem Item entspricht, wird dieser nicht automatisch übernommen.
    # Eine robustere Lösung müsste den Text der ComboBox auswerten, wenn SelectedItem null ist.
    # Fürs Erste gehen wir davon aus, dass der Benutzer ein Item auswählt oder "Keine Änderung" belässt.
    if ($null -ne $selectedTimezoneItem -and $null -ne $selectedTimezoneItem.Tag -and $selectedTimezoneItem.Tag -ne "") {
        $params.Add("TimeZone", $selectedTimezoneItem.Tag)
        $displayParams.Add("Zeitzone", "$($selectedTimezoneItem.Content) ($($selectedTimezoneItem.Tag))")
        Write-Log "Ausgewählte Zeitzone: $($selectedTimezoneItem.Content) ($($selectedTimezoneItem.Tag))" -Type Debug
    } elseif (($null -eq $selectedTimezoneItem) -and (-not [string]::IsNullOrWhiteSpace($script:cmbRegionTimezone.Text)) -and ($script:cmbRegionTimezone.Text -ne "(Keine Änderung)")) {
        # Fall: Editierbare ComboBox mit Texteingabe, die keinem Item entspricht
        # Versuche, die Eingabe direkt als Zeitzonen-ID zu verwenden.
        # Get-TimeZone -ListAvailable könnte hier zur Validierung verwendet werden, ist aber aufwändig.
        # Wir übergeben es erstmal so, Set-MailboxRegionalConfiguration wird ggf. einen Fehler werfen.
        $timeZoneText = $script:cmbRegionTimezone.Text.Trim()
        if ($timeZoneText) { # Sicherstellen, dass es nicht nur Whitespace ist
            $params.Add("TimeZone", $timeZoneText)
            $displayParams.Add("Zeitzone", "$timeZoneText (Manuelle Eingabe)")
            Write-Log "Ausgewählte Zeitzone (manuelle Eingabe): $timeZoneText" -Type Debug
        } else {
            $displayParams.Add("Zeitzone", "(Keine Änderung)")
            Write-Log "Keine Zeitzone zur Änderung ausgewählt (manuelle Eingabe war leer oder nur Leerzeichen)." -Type Debug
        }
    } else {
        $displayParams.Add("Zeitzone", "(Keine Änderung)")
        Write-Log "Keine Zeitzone zur Änderung ausgewählt." -Type Debug
    }
    
    # Datumsformat prüfen
    if ($null -ne $selectedDateFormatItem -and $null -ne $selectedDateFormatItem.Tag -and $selectedDateFormatItem.Tag -ne "") {
        $params.Add("DateFormat", $selectedDateFormatItem.Tag)
        $displayParams.Add("Datumsformat", "$($selectedDateFormatItem.Content) ($($selectedDateFormatItem.Tag))")
        Write-Log "Ausgewähltes Datumsformat: $($selectedDateFormatItem.Content) ($($selectedDateFormatItem.Tag))" -Type Debug
    } else {
        $displayParams.Add("Datumsformat", "(Keine Änderung)")
        Write-Log "Kein Datumsformat zur Änderung ausgewählt." -Type Debug
    }

    # Zeitformat prüfen
    if ($null -ne $selectedTimeFormatItem -and $null -ne $selectedTimeFormatItem.Tag -and $selectedTimeFormatItem.Tag -ne "") {
        $params.Add("TimeFormat", $selectedTimeFormatItem.Tag)
        $displayParams.Add("Zeitformat", "$($selectedTimeFormatItem.Content) ($($selectedTimeFormatItem.Tag))")
        Write-Log "Ausgewähltes Zeitformat: $($selectedTimeFormatItem.Content) ($($selectedTimeFormatItem.Tag))" -Type Debug
    } else {
        $displayParams.Add("Zeitformat", "(Keine Änderung)")
        Write-Log "Kein Zeitformat zur Änderung ausgewählt." -Type Debug
    }

    # Standardordnernamen anpassen (LocalizeDefaultFolderName)
    if ($null -ne $localizeFoldersState) { # Nur wenn CheckBox nicht im unbestimmten Zustand ist
        $params.Add("LocalizeDefaultFolderName", $localizeFoldersState) # $true oder $false
        $localizeDisplayValue = if ($localizeFoldersState) { "Ja" } else { "Nein" }
        $displayParams.Add("Ordnernamen anpassen", $localizeDisplayValue)
        Write-Log "Ausgewählte Option für Ordnernamen anpassen: $localizeDisplayValue (Wert: $localizeFoldersState)" -Type Debug
    } else {
        $displayParams.Add("Ordnernamen anpassen", "(Keine Änderung)")
        Write-Log "Keine Änderung für 'Ordnernamen anpassen' ausgewählt." -Type Debug
    }

    # Prüfen, ob mindestens eine Einstellung geändert werden soll
    if ($params.Count -eq 0) {
        Show-MessageBox -Message "Bitte wählen Sie mindestens eine Einstellung (Sprache, Zeitzone, Datumsformat, Zeitformat, Ordnernamen anpassen) zur Änderung aus." -Title "Keine Auswahl"
        return
    }

    # Zielpostfächer ermitteln
    $targetMailboxes = @()
    if ($mailboxInput -eq 'ALL') {
        # Bestätigung für 'ALL' einholen
        $confirmMessage = "WARNUNG: Sie sind dabei, die regionalen Einstellungen für ALLE Benutzerpostfächer zu ändern.`n"
        $confirmMessage += "Sprache: $($displayParams.Sprache)`n"
        $confirmMessage += "Zeitzone: $($displayParams.Zeitzone)`n"
        $confirmMessage += "Datumsformat: $($displayParams.Datumsformat)`n"
        $confirmMessage += "Zeitformat: $($displayParams.Zeitformat)`n"
        $confirmMessage += "Ordnernamen anpassen: $($displayParams.'Ordnernamen anpassen')`n"
        $confirmMessage += "`nDies kann sehr lange dauern und sollte mit Vorsicht verwendet werden.`n`n"
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
        $targetMailboxes = $mailboxInput -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }

    if ($targetMailboxes.Count -eq 0) {
         Show-MessageBox -Message "Keine gültigen Postfach-Identitäten angegeben oder gefunden." -Title "Keine Postfächer"
         Update-GuiText -TextElement $statusTextBlock -Message "Keine gültigen Postfächer zur Verarbeitung angegeben."
         Log-Action "Keine gültigen Postfächer zur Verarbeitung angegeben für '$mailboxInput'."
         return
    }

    Update-GuiText -TextElement $statusTextBlock -Message "Starte Verarbeitung von $($targetMailboxes.Count) Postfach/Postfächern..."
    Log-Action "Starte Set-ExoMailboxRegionalSettings für $($targetMailboxes.Count) Postfächer mit Parametern: $($params.Keys -join ', ')"

    try {
        # Direkter Aufruf der Funktion
        $OperationResult = Set-ExoMailboxRegionalSettings -MailboxesToProcess $targetMailboxes -Parameters $params -Form $script:Form -StatusTextBlock $statusTextBlock
        
        $successCount = $OperationResult.Success.Count
        $failedCount = $OperationResult.Failed.Count
        $totalCount = $successCount + $failedCount

        $summaryMessage = "Regionaleinstellungen verarbeitet: $successCount erfolgreich, $failedCount fehlgeschlagen (von $totalCount)."
        Update-GuiText -TextElement $statusTextBlock -Message $summaryMessage
        Log-Action $summaryMessage

        if ($failedCount -gt 0) {
            $errorDetails = $OperationResult.Failed | ForEach-Object { "  - $($_.Mailbox): $($_.Error)" }
            $maxErrorLines = 10 
            $truncated = $false
            if ($errorDetails.Count -gt $maxErrorLines) {
                $errorDetails = $errorDetails[0..($maxErrorLines - 1)]
                $truncated = $true
            }
            $errorDetailsString = $errorDetails -join "`n"
            if ($truncated) {
                $errorDetailsString += "`n  ... (weitere Fehler im Log)"
            }

            if ($null -ne $script:txtRegionResult) {
                 $script:txtRegionResult.Dispatcher.Invoke({ $script:txtRegionResult.Text = "Fehler bei Verarbeitung:`n$errorDetailsString" }) | Out-Null
            }

            Show-MessageBox -Message "Einige regionale Einstellungen konnten nicht angewendet werden:`n$errorDetailsString" -Title "Fehler bei Verarbeitung"
            Log-Action "Fehlerdetails Regionaleinstellungen: $($OperationResult.Failed | ConvertTo-Json -Depth 3)"
        } elseif ($successCount -gt 0) {
             if ($null -ne $script:txtRegionResult) {
                 $script:txtRegionResult.Dispatcher.Invoke({ $script:txtRegionResult.Text = "$successCount Postfach/Postfächer erfolgreich aktualisiert." }) | Out-Null
             }
             Show-MessageBox -Message "Regionaleinstellungen für $successCount Postfach/Postfächer erfolgreich angewendet." -Title "Erfolg"
        } else { # Weder Fehler noch Erfolg (z.B. wenn keine Änderungen vorgenommen werden sollten oder alle scheiterten und nicht in Failed landeten)
             if ($null -ne $script:txtRegionResult) {
                  $script:txtRegionResult.Dispatcher.Invoke({ $script:txtRegionResult.Text = "Keine Änderungen vorgenommen oder alle Operationen fehlgeschlagen/übersprungen." }) | Out-Null
             }
             Show-MessageBox -Message "Keine Änderungen vorgenommen oder alle Operationen fehlgeschlagen/übersprungen." -Title "Information"
        }
    } catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Schwerwiegender Fehler beim direkten Anwenden der regionalen Einstellungen."
        Update-GuiText -TextElement $statusTextBlock -Message "Fehler: $errorMsg"
        Log-Action "FEHLER beim direkten Anwenden der regionalen Einstellungen: $errorMsg"
        if ($null -ne $script:txtRegionResult) {
            $script:txtRegionResult.Dispatcher.Invoke({ $script:txtRegionResult.Text = "Fehler: $errorMsg" }) | Out-Null
        }
        Show-MessageBox -Message "Ein schwerwiegender Fehler ist aufgetreten:`n$errorMsg" -Title "Fehler"
    }
}
# Ende Invoke-SetRegionSettingsAction

# Funktion zur Ausführung der Aktion "Aktuelle Einstellungen abrufen"
function Invoke-GetRegionSettingsAction {
    [CmdletBinding()]
    param()

    if (-not $script:IsConnected) {
        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Nicht verbunden"
        return
    }

    $mailboxId = $script:txtGetRegionMailbox.Text.Trim()
    $statusTextBlock = $script:txtStatus
    $resultTextBlock = $script:txtRegionResult
    
    $languageComboBox = $script:cmbRegionLanguage
    $timezoneComboBox = $script:cmbRegionTimezone
    $dateFormatComboBox = $script:cmbRegionDateFormat
    $timeFormatComboBox = $script:cmbRegionTimeFormat
    $localizeFoldersCheckBox = $script:chkRegionDefaultFolderNameMatchingUserLanguage

    # ... (Validierung von $mailboxId bleibt gleich) ...
    if ([string]::IsNullOrWhiteSpace($mailboxId)) {
        Show-MessageBox -Message "Bitte geben Sie die E-Mail-Adresse des Postfachs an, dessen Einstellungen abgerufen werden sollen." -Title "Eingabe fehlt"
        return
    }
    if ($mailboxId -eq 'ALL') {
        Show-MessageBox -Message "Bitte geben Sie eine einzelne Postfach-Identität zum Abrufen an (nicht 'ALL')." -Title "Ungültige Eingabe"
        return
    }

    Update-GuiText -TextElement $statusTextBlock -Message "Rufe regionale Einstellungen für '$mailboxId' ab..."
    Log-Action "Starte Get-ExoMailboxRegionalSettings für '$mailboxId'."

    if ($null -ne $resultTextBlock) {
        $resultTextBlock.Dispatcher.Invoke({ $resultTextBlock.Text = "" }) | Out-Null
    }

    try {
        $OperationResult = Get-ExoMailboxRegionalSettings -MailboxIdentity $mailboxId
        
        Write-Log ("Get-ExoMailboxRegionalSettings OperationResult für '$mailboxId': " + ($OperationResult | ConvertTo-Json -Depth 4)) -Type Debug

        if ($OperationResult.Success) {
            $regionalSettings = $OperationResult.Result 
            
            if ($null -eq $regionalSettings) {
                # ... (Logik für Null-Ergebnis und UI-Reset für Sprache, Datum, Zeitzone) ...
                $noDataMessage = "Abruf für '$mailboxId' erfolgreich, aber keine Einstellungsdaten (regionalSettings ist null) zurückgegeben."
                Update-GuiText -TextElement $statusTextBlock -Message $noDataMessage
                Log-Action $noDataMessage -Type Warning
                if ($null -ne $resultTextBlock) {
                    $resultTextBlock.Dispatcher.Invoke({ $resultTextBlock.Text = "Keine Einstellungsdaten für '$mailboxId' empfangen." }) | Out-Null
                }
                try {
                    if ($null -ne $languageComboBox) { $languageComboBox.Dispatcher.Invoke({ $languageComboBox.SelectedItem = ($languageComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null }
                    Populate-DateFormatComboBox -ComboBox $dateFormatComboBox -CultureName "" 
                    Populate-TimeFormatComboBox -ComboBox $timeFormatComboBox -CultureName "" # NEU: Reset TimeFormat
                    Populate-TimezoneComboBox -ComboBox $timezoneComboBox -CultureName "DEFAULT_ALL" 
                    if ($null -ne $localizeFoldersCheckBox) { $localizeFoldersCheckBox.Dispatcher.Invoke({ $localizeFoldersCheckBox.IsChecked = $null }) | Out-Null }
                } catch { Write-Log "Fehler beim Zurücksetzen der UI nach Null-Ergebnis (GetRegion): $($_.Exception.Message)" -Type Warning }

            } else {
                $retrievedIdentity = $mailboxId 
                if ($regionalSettings.PSObject.Properties['Identity'] -and $null -ne $regionalSettings.Identity) {
                    $retrievedIdentity = $regionalSettings.Identity.ToString() 
                } elseif ($regionalSettings.PSObject.Properties['DistinguishedName'] -and $null -ne $regionalSettings.DistinguishedName) {
                    $retrievedIdentity = $regionalSettings.DistinguishedName.ToString()
                }
                
                $statusMsg = "Regionale Einstellungen für '$retrievedIdentity' erfolgreich abgerufen."
                Update-GuiText -TextElement $statusTextBlock -Message "Abruf erfolgreich. Aktualisiere UI..." 
                Log-Action $statusMsg

                $currentCultureForUI = ""

                if ($null -ne $regionalSettings.Language) {
                    $langCodeToSelect = $regionalSettings.Language.ToString()
                    $langItem = $languageComboBox.Items | Where-Object { $_.Tag -eq $langCodeToSelect } | Select-Object -First 1
                    if ($null -ne $langItem) {
                        $languageComboBox.Dispatcher.Invoke({ $languageComboBox.SelectedItem = $langItem }) | Out-Null
                        $currentCultureForUI = $langCodeToSelect 
                    } else {
                        $languageComboBox.Dispatcher.Invoke({ $languageComboBox.SelectedItem = ($languageComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null
                    }
                } else {
                     $languageComboBox.Dispatcher.Invoke({ $languageComboBox.SelectedItem = ($languageComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null
                }
                
                # Formate und Zeitzonen basierend auf der gesetzten Sprache laden
                Populate-DateFormatComboBox -ComboBox $dateFormatComboBox -CultureName $currentCultureForUI
                Populate-TimeFormatComboBox -ComboBox $timeFormatComboBox -CultureName $currentCultureForUI # NEU
                Populate-TimezoneComboBox -ComboBox $timezoneComboBox -CultureName $currentCultureForUI 
                
                # Datumsformat setzen
                if ($null -ne $regionalSettings.DateFormat) {
                    $dateFormatToSelect = $regionalSettings.DateFormat.ToString()
                    $dateItem = $dateFormatComboBox.Items | Where-Object { $_.Tag -eq $dateFormatToSelect } | Select-Object -First 1
                    if ($null -ne $dateItem) {
                        $dateFormatComboBox.Dispatcher.Invoke({ $dateFormatComboBox.SelectedItem = $dateItem }) | Out-Null
                    } else {
                        $dateFormatComboBox.Dispatcher.Invoke({ $dateFormatComboBox.SelectedItem = ($dateFormatComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null
                    }
                } else {
                     $dateFormatComboBox.Dispatcher.Invoke({ $dateFormatComboBox.SelectedItem = ($dateFormatComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null
                }
                
                # Zeitformat setzen (NEU)
                if ($null -ne $regionalSettings.TimeFormat) {
                    $timeFormatToSelect = $regionalSettings.TimeFormat.ToString()
                    $timeItem = $timeFormatComboBox.Items | Where-Object { $_.Tag -eq $timeFormatToSelect } | Select-Object -First 1
                    if ($null -ne $timeItem) {
                        $timeFormatComboBox.Dispatcher.Invoke({ $timeFormatComboBox.SelectedItem = $timeItem }) | Out-Null
                        Write-Log "Zeitformat '$timeFormatToSelect' in UI ausgewählt." -Type Debug
                    } else {
                        Write-Log "Abgerufenes Zeitformat '$timeFormatToSelect' nicht in UI-Liste für Kultur '$currentCultureForUI'. (Keine Änderung) bleibt." -Type Warning
                        $timeFormatComboBox.Dispatcher.Invoke({ $timeFormatComboBox.SelectedItem = ($timeFormatComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null
                    }
                } else {
                     Write-Log "Kein Zeitformat von Exchange empfangen. UI bleibt auf (Keine Änderung)." -Type Debug
                     $timeFormatComboBox.Dispatcher.Invoke({ $timeFormatComboBox.SelectedItem = ($timeFormatComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null
                }

                # Zeitzone setzen
                if ($null -ne $regionalSettings.TimeZone) {
                    $timeZoneToSelect = $regionalSettings.TimeZone.ToString()
                    $tzItem = $timezoneComboBox.Items | Where-Object { $_.Tag -eq $timeZoneToSelect } | Select-Object -First 1
                    if ($null -ne $tzItem) {
                        $timezoneComboBox.Dispatcher.Invoke({ $timezoneComboBox.SelectedItem = $tzItem }) | Out-Null
                    } else {
                         $timezoneComboBox.Dispatcher.Invoke({ $timezoneComboBox.SelectedItem = ($timezoneComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null
                    }
                } else {
                     $timezoneComboBox.Dispatcher.Invoke({ $timezoneComboBox.SelectedItem = ($timezoneComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null
                }
                
                # LocalizeDefaultFolderName setzen
                if ($regionalSettings.PSObject.Properties.ContainsKey("LocalizeDefaultFolderName") -and `
                    $null -ne $regionalSettings.LocalizeDefaultFolderName) {
                    $localizeValue = $regionalSettings.LocalizeDefaultFolderName
                    if ($localizeValue -is [bool]) {
                        $localizeFoldersCheckBox.Dispatcher.Invoke({ $localizeFoldersCheckBox.IsChecked = $localizeValue }) | Out-Null
                    } else {
                        $localizeFoldersCheckBox.Dispatcher.Invoke({ $localizeFoldersCheckBox.IsChecked = $null }) | Out-Null 
                    }
                } else {
                    $localizeFoldersCheckBox.Dispatcher.Invoke({ $localizeFoldersCheckBox.IsChecked = $null }) | Out-Null
                }

                # ... (restliche Logik zur Anzeige in $resultTextBlock) ...
                if ($null -ne $resultTextBlock) {
                    $formattedProperties = $regionalSettings | Format-List * | Out-String 
                    $resultOutput = "Abgerufene Einstellungen für: $retrievedIdentity`n" + ("-" * 50) + "`n$($formattedProperties.Trim())"
                    $resultTextBlock.Dispatcher.Invoke({ $resultTextBlock.Text = $resultOutput }) | Out-Null
                }
                Update-GuiText -TextElement $statusTextBlock -Message "UI mit abgerufenen Einstellungen aktualisiert." 
            }
        } else { 
            # ... (bestehende Fehlerbehandlung im OperationResult.Success -eq $false Fall, inkl. UI-Reset) ...
            $errorMsg = $OperationResult.Error
            $failMessage = "Fehler beim Abrufen der regionalen Einstellungen für '$mailboxId': $errorMsg"
            Update-GuiText -TextElement $statusTextBlock -Message "Fehler beim Abrufen." 
            Log-Action $failMessage -Type Error
            if ($null -ne $resultTextBlock) {
                $resultTextBlock.Dispatcher.Invoke({ $resultTextBlock.Text = $failMessage }) | Out-Null
            }
            Show-MessageBox -Message $failMessage -Title "Fehler" 
            try {
                 if ($null -ne $languageComboBox) { $languageComboBox.Dispatcher.Invoke({ $languageComboBox.SelectedItem = ($languageComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null }
                 Populate-DateFormatComboBox -ComboBox $dateFormatComboBox -CultureName "" 
                 Populate-TimeFormatComboBox -ComboBox $timeFormatComboBox -CultureName "" # NEU: Reset TimeFormat
                 Populate-TimezoneComboBox -ComboBox $timezoneComboBox -CultureName "DEFAULT_ALL"
                 if ($null -ne $localizeFoldersCheckBox) { $localizeFoldersCheckBox.Dispatcher.Invoke({ $localizeFoldersCheckBox.IsChecked = $null }) | Out-Null}
            } catch { Write-Log "Fehler beim Zurücksetzen der UI (GetRegion OperationResult Fail): $($_.Exception.Message)" -Type Warning }
        }
    } catch { 
        # ... (bestehender äußerer catch-Block, inkl. UI-Reset) ...
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Schwerwiegender Fehler beim Abrufen der regionalen Einstellungen für '$mailboxId'."
        Update-GuiText -TextElement $statusTextBlock -Message "Schwerwiegender Fehler."
        Log-Action "FEHLER (Invoke-GetRegionSettingsAction): $errorMsg" -Type Error
        if ($null -ne $resultTextBlock) {
            $resultTextBlock.Dispatcher.Invoke({ $resultTextBlock.Text = "Schwerwiegender Fehler: $errorMsg" }) | Out-Null
        }
        Show-MessageBox -Message "Ein schwerwiegender Fehler ist aufgetreten:`n$errorMsg" -Title "Fehler"
        try {
             if ($null -ne $languageComboBox) { $languageComboBox.Dispatcher.Invoke({ $languageComboBox.SelectedItem = ($languageComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) }) | Out-Null }
             Populate-DateFormatComboBox -ComboBox $dateFormatComboBox -CultureName "" 
             Populate-TimeFormatComboBox -ComboBox $timeFormatComboBox -CultureName "" # NEU: Reset TimeFormat
             Populate-TimezoneComboBox -ComboBox $timezoneComboBox -CultureName "DEFAULT_ALL"
             if ($null -ne $localizeFoldersCheckBox) { $localizeFoldersCheckBox.Dispatcher.Invoke({ $localizeFoldersCheckBox.IsChecked = $null }) | Out-Null}
        } catch { Write-Log "Fehler beim Zurücksetzen der UI (GetRegion äußerer Catch): $($_.Exception.Message)" -Type Warning }
    }
}
# Ende Invoke-GetRegionSettingsAction

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
        $Control.Dispatcher.Invoke({
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
        $Control.Dispatcher.Invoke({
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

        }, "Normal") | Out-Null # Invoke

    } catch {
        # Use ${} for clarity
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Aktualisieren der ComboBox-Elemente."
        Write-Log $errorMsg -Type Error
        Log-Action "FEHLER: Update-ComboBoxItems - ${errorMsg}"
    }
}

# Hilfsfunktion zum Befüllen der Zeitformat-ComboBox basierend auf Kultur
function Populate-TimeFormatComboBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ComboBox]$ComboBox,
        [Parameter(Mandatory=$false)]
        [string]$CultureName # z.B. "de-DE"
    )
    try {
        Write-Log "Befülle Zeitformat-ComboBox für Kultur '$CultureName'..." -Type Debug
        
        $currentValue = $null
        if ($ComboBox.SelectedItem -ne $null -and $ComboBox.SelectedItem.Tag -ne $null) {
            $currentValue = $ComboBox.SelectedItem.Tag
        } elseif (-not [string]::IsNullOrEmpty($ComboBox.Text) -and $ComboBox.Text -ne "(Keine Änderung)") {
             $currentValue = $ComboBox.Text
        }

        $ComboBox.Items.Clear()

        $noChangeItem = New-Object System.Windows.Controls.ComboBoxItem
        $noChangeItem.Content = "(Keine Änderung)"
        $noChangeItem.Tag = "" 
        [void]$ComboBox.Items.Add($noChangeItem)

        $formatsToUse = $null
        if (-not [string]::IsNullOrEmpty($CultureName) -and $script:RelevantTimeFormatsPerCulture.ContainsKey($CultureName)) {
            $formatsToUse = $script:RelevantTimeFormatsPerCulture[$CultureName]
        } else {
            Write-Log "Kultur '$CultureName' nicht in RelevantTimeFormatsPerCulture gefunden oder leer, verwende DEFAULT Zeitformate." -Type Debug
            $formatsToUse = $script:RelevantTimeFormatsPerCulture["DEFAULT"]
        }

        if ($null -ne $formatsToUse) {
            foreach ($formatInfo in $formatsToUse) {
                $item = New-Object System.Windows.Controls.ComboBoxItem
                $item.Content = $formatInfo.Display 
                $item.Tag = $formatInfo.Value     
                [void]$ComboBox.Items.Add($item)
            }
        }
        
        $itemToSelectAfterPopulation = $ComboBox.Items | Where-Object { $_.Tag -eq $currentValue } | Select-Object -First 1
        if ($null -ne $itemToSelectAfterPopulation) {
            $ComboBox.SelectedItem = $itemToSelectAfterPopulation
        } else {
            $ComboBox.SelectedItem = $noChangeItem # Auf "(Keine Änderung)" zurückfallen
        }

        Write-Log "Zeitformat-ComboBox erfolgreich befüllt. Ausgewählt: '$($ComboBox.SelectedItem.Content)' (Tag: '$($ComboBox.SelectedItem.Tag)')" -Type Debug
        return $true
    } catch {
        Write-Log "Fehler beim Befüllen der Zeitformat-ComboBox: $($_.Exception.Message)" -Type Error
        Log-Action "Fehler beim Befüllen der Zeitformat-ComboBox: $($_.Exception.Message)"
        if ($ComboBox.Items.Count > 0) { $ComboBox.SelectedIndex = 0 } # Auf "(Keine Änderung)" zurückfallen
        return $false
    }
}

# Hilfsfunktion zum Befüllen der Sprachen-ComboBox
function Populate-LanguageComboBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ComboBox]$ComboBox
    )
    try {
        Write-Log "Befülle Sprachen-ComboBox mit gefilterter Sprachliste..." -Type Debug

        $currentValue = $null
        if ($ComboBox.SelectedItem -ne $null -and $ComboBox.SelectedItem.Tag -ne $null) {
            $currentValue = $ComboBox.SelectedItem.Tag
        } elseif (-not [string]::IsNullOrEmpty($ComboBox.Text)) {
             $currentValue = $ComboBox.Text
        }

        $ComboBox.Items.Clear()

        $noChangeItem = New-Object System.Windows.Controls.ComboBoxItem
        $noChangeItem.Content = "(Keine Änderung)"
        $noChangeItem.Tag = "" # Leerer Tag für "Keine Änderung"
        $noChangeItem.IsSelected = $true
        [void]$ComboBox.Items.Add($noChangeItem)

        # Gewünschte spezifische Kultur-Codes
        $desiredSpecificCultures = @(
            "de-DE", "de-AT", "de-CH", # Deutsch
            "en-US", "en-GB",         # Englisch
            "fr-FR", "fr-CA", "fr-CH", "fr-BE", "fr-LU", # Französisch
            "it-IT", "it-CH",         # Italienisch
            "pl-PL",                  # Polnisch
            "es-ES", "es-MX", # Spanisch (Spanien, Mexiko)
            "nl-NL", "nl-BE"          # Niederländisch
        )

        $allCultures = [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::SpecificCultures)
        $filteredCultures = [System.Collections.Generic.List[System.Globalization.CultureInfo]]::new()

        foreach ($culture in $allCultures) {
            if ($desiredSpecificCultures -contains $culture.Name) {
                $filteredCultures.Add($culture)
            }
        }
        
        # Sortieren der gefilterten Kulturen nach Anzeigenamen
        $sortedFilteredCultures = $filteredCultures | Sort-Object -Property DisplayName

        foreach ($culture in $sortedFilteredCultures) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $culture.DisplayName # z.B. "Deutsch (Deutschland)"
            $item.Tag = $culture.Name          # z.B. "de-DE"
            [void]$ComboBox.Items.Add($item)
        }

        if (-not [string]::IsNullOrEmpty($currentValue)) {
            $itemToSelect = $ComboBox.Items | Where-Object { $_.Tag -eq $currentValue -or $_.Content -eq $currentValue } | Select-Object -First 1
            if ($null -ne $itemToSelect) {
                $ComboBox.SelectedItem = $itemToSelect
            } elseif ($ComboBox.IsEditable -and $currentValue -ne "(Keine Änderung)") {
                 $ComboBox.Text = $currentValue
            } else {
                 $ComboBox.SelectedItem = $noChangeItem
            }
        } else {
             $ComboBox.SelectedItem = $noChangeItem
        }
        Write-Log "Sprachen-ComboBox erfolgreich mit gefilterter Liste befüllt." -Type Debug
        return $true
    } catch {
        Write-Log "Fehler beim Befüllen der Sprachen-ComboBox (gefiltert): $($_.Exception.Message)" -Type Warning
        Log-Action "Warnung: Sprachen konnten nicht dynamisch gefiltert geladen werden."
        return $false
    }
}

# Hilfsfunktion zum Befüllen der Datumsformat-ComboBox basierend auf Kultur
function Populate-DateFormatComboBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ComboBox]$ComboBox,
        [Parameter(Mandatory=$false)]
        [string]$CultureName # z.B. "de-DE"
    )
    try {
        Write-Log "Befülle Datumsformat-ComboBox für Kultur '$CultureName'..." -Type Debug
        
        # Aktuelle Auswahl merken (Tag oder Text)
        $currentValue = $null
        if ($ComboBox.SelectedItem -ne $null -and $ComboBox.SelectedItem.Tag -ne $null) {
            $currentValue = $ComboBox.SelectedItem.Tag
        } elseif (-not [string]::IsNullOrEmpty($ComboBox.Text) -and $ComboBox.Text -ne "(Keine Änderung)") {
             $currentValue = $ComboBox.Text # Fallback auf Text, wenn manuell editiert
        }

        $ComboBox.Items.Clear()

        $noChangeItem = New-Object System.Windows.Controls.ComboBoxItem
        $noChangeItem.Content = "(Keine Änderung)"
        $noChangeItem.Tag = "" # Leerer Tag für "Keine Änderung"
        [void]$ComboBox.Items.Add($noChangeItem)

        $formatsToUse = $null
        if (-not [string]::IsNullOrEmpty($CultureName) -and $script:ExchangeValidDateFormats.ContainsKey($CultureName)) {
            $formatsToUse = $script:ExchangeValidDateFormats[$CultureName]
        } else {
            Write-Log "Kultur '$CultureName' nicht in ExchangeValidDateFormats gefunden oder leer, verwende DEFAULT Datumsformate." -Type Debug
            $formatsToUse = $script:ExchangeValidDateFormats["DEFAULT"]
        }

        if ($null -ne $formatsToUse) {
            foreach ($formatInfo in $formatsToUse) {
                $item = New-Object System.Windows.Controls.ComboBoxItem
                $item.Content = $formatInfo.Display # z.B. "TT.MM.JJJJ (Standard)"
                $item.Tag = $formatInfo.Value     # z.B. "dd.MM.yyyy"
                [void]$ComboBox.Items.Add($item)
            }
        }
        
        # Versuche vorherige Auswahl wiederherzustellen oder "(Keine Änderung)" zu wählen
        $itemToSelectAfterPopulation = $ComboBox.Items | Where-Object { $_.Tag -eq $currentValue } | Select-Object -First 1
        if ($null -ne $itemToSelectAfterPopulation) {
            $ComboBox.SelectedItem = $itemToSelectAfterPopulation
        } else {
            # Wenn der alte Wert nicht mehr gültig ist (oder keiner war), setze auf "(Keine Änderung)"
            $ComboBox.SelectedItem = $noChangeItem
        }

        Write-Log "Datumsformat-ComboBox erfolgreich befüllt. Ausgewählt: '$($ComboBox.SelectedItem.Content)' (Tag: '$($ComboBox.SelectedItem.Tag)')" -Type Debug
        return $true
    } catch {
        Write-Log "Fehler beim Befüllen der Datumsformat-ComboBox: $($_.Exception.Message)" -Type Error
        Log-Action "Fehler beim Befüllen der Datumsformat-ComboBox: $($_.Exception.Message)"
        # Im Fehlerfall "(Keine Änderung)" auswählen
        if ($ComboBox.Items.Count > 0) { $ComboBox.SelectedIndex = 0 }
        return $false
    }
}

# Hilfsfunktion zum Befüllen der Zeitzonen-ComboBox
function Populate-TimezoneComboBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ComboBox]$ComboBox,
        [Parameter(Mandatory=$false)]
        [string]$CultureName # z.B. "de-DE", leer für alle/default
    )
    try {
        Write-Log "Befülle Zeitzonen-ComboBox für Kultur '$CultureName'..." -Type Debug

        $currentValue = $null
        if ($ComboBox.SelectedItem -ne $null -and $ComboBox.SelectedItem.Tag -ne $null) {
            $currentValue = $ComboBox.SelectedItem.Tag
        } elseif (-not [string]::IsNullOrEmpty($ComboBox.Text) -and $ComboBox.Text -ne "(Keine Änderung)") {
            $currentValue = $ComboBox.Text
        }

        $ComboBox.Items.Clear()

        $noChangeItem = New-Object System.Windows.Controls.ComboBoxItem
        $noChangeItem.Content = "(Keine Änderung)"
        $noChangeItem.Tag = ""
        [void]$ComboBox.Items.Add($noChangeItem)

        $allSystemTimezones = [System.TimeZoneInfo]::GetSystemTimeZones()
        $timezonesToDisplay = [System.Collections.Generic.List[System.TimeZoneInfo]]::new()

        $specificTimezoneIDs = $null
        if (-not [string]::IsNullOrEmpty($CultureName) -and $script:RelevantTimezonesPerCulture.ContainsKey($CultureName)) {
            $mappedValue = $script:RelevantTimezonesPerCulture[$CultureName]
            if ($mappedValue -is [array]) {
                $specificTimezoneIDs = $mappedValue
                Write-Log "Kultur '$CultureName' gefunden. Lade spezifische Zeitzonen-IDs: $($specificTimezoneIDs -join ', ')" -Type Debug
            } elseif ($mappedValue -is [bool] -and $mappedValue -eq $true -and $CultureName -eq "DEFAULT_ALL") {
                # Spezialfall für DEFAULT_ALL, lade alle
                $specificTimezoneIDs = $null # Signalisiert, alle zu nehmen
                Write-Log "Kultur '$CultureName' ist DEFAULT_ALL. Lade alle Systemzeitzonen." -Type Debug
            }
        } elseif ($script:RelevantTimezonesPerCulture.ContainsKey("DEFAULT_ALL") -and $script:RelevantTimezonesPerCulture["DEFAULT_ALL"] -eq $true) {
             Write-Log "Keine spezifische Kultur '$CultureName' oder kein Mapping. Fallback auf DEFAULT_ALL: Lade alle Systemzeitzonen." -Type Debug
             $specificTimezoneIDs = $null # Signalisiert, alle zu nehmen
        }


        if ($null -ne $specificTimezoneIDs) {
            foreach ($tzID in $specificTimezoneIDs) {
                $foundTz = $allSystemTimezones | Where-Object { $_.Id -eq $tzID } | Select-Object -First 1
                if ($null -ne $foundTz -and -not $timezonesToDisplay.Contains($foundTz)) {
                    $timezonesToDisplay.Add($foundTz)
                } else {
                    Write-Log "Zeitzonen-ID '$tzID' für Kultur '$CultureName' nicht im System gefunden oder bereits hinzugefügt." -Type Warning
                }
            }
        } else { # Lade alle Systemzeitzonen
            $timezonesToDisplay.AddRange($allSystemTimezones)
        }
        
        # Sortiere die anzuzeigenden Zeitzonen
        $sortedTimezonesToDisplay = $timezonesToDisplay | Sort-Object @{Expression={$_.BaseUtcOffset.TotalHours}}, @{Expression={$_.DisplayName}}
        
        foreach ($tz in $sortedTimezonesToDisplay) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $offsetHours = $tz.BaseUtcOffset.TotalHours
            $offsetString = ""
            if ($offsetHours -lt 0) {
                $offsetString = "UTC$($tz.BaseUtcOffset.ToString('hh\:mm'))"
            } elseif ($offsetHours -gt 0) {
                $offsetString = "UTC+$($tz.BaseUtcOffset.ToString('hh\:mm'))"
            } else {
                $offsetString = "UTC"
            }
            $item.Content = "($offsetString) $($tz.DisplayName)"
            $item.Tag = $tz.Id
            [void]$ComboBox.Items.Add($item)
        }

        $itemToSelectAfterPopulation = $ComboBox.Items | Where-Object { $_.Tag -eq $currentValue } | Select-Object -First 1
        if ($null -ne $itemToSelectAfterPopulation) {
            $ComboBox.SelectedItem = $itemToSelectAfterPopulation
        } else {
            $ComboBox.SelectedItem = $noChangeItem
        }
        Write-Log "Zeitzonen-ComboBox erfolgreich befüllt für Kultur '$CultureName'. Ausgewählt: '$($ComboBox.SelectedItem.Content)' (Tag: '$($ComboBox.SelectedItem.Tag)')" -Type Debug
        return $true
    } catch {
        Write-Log "Fehler beim Befüllen der Zeitzonen-ComboBox für Kultur '$CultureName': $($_.Exception.Message)" -Type Error
        Log-Action "Fehler beim Befüllen der Zeitzonen-ComboBox: $($_.Exception.Message)"
        if ($ComboBox.Items.Count > 0) { $ComboBox.SelectedIndex = 0 }
        return $false
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
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[string]]$MailboxesToProcess,

        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters, # Enthält Schlüssel wie Language, TimeZone, DateFormat, TimeFormat, LocalizeDefaultFolderName

        [Parameter(Mandatory = $false)]
        [System.Windows.Window]$Form,

        [Parameter(Mandatory = $false)]
        [System.Windows.Controls.TextBlock]$StatusTextBlock
    )

    $results = @{
        Success = [System.Collections.Generic.List[string]]::new()
        Failed  = [System.Collections.Generic.List[object]]::new()
    }
    $total = $MailboxesToProcess.Count
    $processed = 0

    foreach ($mailbox in $MailboxesToProcess) {
        $processed++
        $progress = [int](($processed / $total) * 100)

        if ($null -ne $Form -and $null -ne $StatusTextBlock -and $null -ne $Form.Dispatcher) {
            try {
                $updateUiAction = {
                    param(
                        [string]$CurrentMailbox,
                        [int]$NumProcessed,
                        [int]$TotalMailboxes,
                        [int]$PercentageComplete,
                        [System.Windows.Controls.TextBlock]$TbStatus
                    )
                    if ($null -ne $TbStatus) {
                        $TbStatus.Text = "Verarbeite ${CurrentMailbox} (${NumProcessed}/${TotalMailboxes})... [${PercentageComplete}%]"
                    }
                }
                # KORRIGIERTER AUFRUF
                $Form.Dispatcher.Invoke($updateUiAction, [System.Windows.Threading.DispatcherPriority]::Normal, $mailbox, $processed, $total, $progress, $StatusTextBlock) | Out-Null
            }
            catch {
                Write-Warning "Fehler beim Aktualisieren des GUI-Status aus Hintergrundthread für ($($mailbox)): $($_.Exception.Message)"
                Log-Action "WARNUNG: Fehler beim GUI-Update im Hintergrund für ($($mailbox)): $($_.Exception.Message)"
            }
        }
        else {
            Write-Verbose "Verarbeite ${mailbox} (${processed}/${total})... [${progress}%]"
        }

        try {
            $cmdletParams = $Parameters.Clone()

            if ($cmdletParams.ContainsKey("LocalizeDefaultFolderName")) {
                $localizeValue = $cmdletParams["LocalizeDefaultFolderName"]
                $cmdletParams.Remove("LocalizeDefaultFolderName")
                if ($localizeValue -eq $true) {
                    $cmdletParams.Add("LocalizeDefaultFolderName", $true)
                } elseif ($localizeValue -eq $false) {
                    $cmdletParams.Add("LocalizeDefaultFolderName", $false)
                }
            }

            Write-Log "Setze regionale Einstellungen für ${mailbox} mit cmdletParams: $($cmdletParams | Out-String)" -Type Debug
            Log-Action "Versuche Set-MailboxRegionalConfiguration für ${mailbox} mit Parametern: $($cmdletParams.Keys -join ', ')"

            $warningMessages = New-Object System.Collections.Generic.List[string]
            $originalWarningPreference = $WarningPreference
            $WarningPreference = "SilentlyContinue" # Um Warnungen im Skript abzufangen

            try {
                if ($PSCmdlet.ShouldProcess($mailbox, "Regionale Einstellungen anwenden (Sprache/Zeitzone/Datumsformate/Ordnernamen)")) {
                    Set-MailboxRegionalConfiguration -Identity $mailbox @cmdletParams -Confirm:$false -ErrorAction Stop -WarningVariable +tempWarnings
                    
                    if ($null -ne $tempWarnings -and $tempWarnings.Count -gt 0) {
                        foreach($warn in $tempWarnings){
                            $formattedWarning = $warn.Message
                            Write-Log "Exchange WARNUNG für Postfach ${mailbox}: ${formattedWarning}" -Type Warning
                            Log-Action "Exchange WARNUNG für Postfach ${mailbox}: ${formattedWarning}"
                            $warningMessages.Add($formattedWarning)
                        }
                    }
                    $tempWarnings = $null 
                }
            }
            finally {
                $WarningPreference = $originalWarningPreference
            }

            if ($warningMessages.Count -gt 0) {
                $results.Failed.Add(@{ Mailbox = $mailbox; Error = ($warningMessages -join "`n") })
            } else {
                 $results.Success.Add($mailbox)
                 Log-Action "Regionale Einstellungen für ${mailbox} erfolgreich gesetzt (oder ohne Fehler/Warnungen abgeschlossen)."
            }
        }
        catch {
            $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Setzen der Einstellungen für ${mailbox}."
            $results.Failed.Add(@{ Mailbox = $mailbox; Error = $errorMsg })
            Log-Action "FEHLER beim Setzen der regionalen Einstellungen für ${mailbox}: ${errorMsg}"
            Write-Warning "Fehler bei Postfach ${mailbox}: ${errorMsg}"
        }
    }
    return $results
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

# Funktion zum Abrufen aller relevanten Exchange-Gruppentypen
function Get-AllGroupTypesAction {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Rufe alle relevanten Gruppen ab (Verteiler, E-Mail-aktivierte Sicherheitsgruppen, Microsoft 365-Gruppen)" -Type "Info"
        Update-StatusBar -Message "Rufe alle Gruppentypen ab..." -Type Info
        $allGroupsData = @()

        # Verteilergruppen und E-Mail-aktivierte Sicherheitsgruppen
        # Get-DistributionGroup liefert sowohl Verteilergruppen als auch E-Mail-aktivierte Sicherheitsgruppen
        # Der Parameter -ResultSize Unlimited ist für die EXO V2/V3 Cmdlets nicht gültig/nötig.
        $distributionAndSecurityGroups = Get-DistributionGroup -ErrorAction SilentlyContinue
        if ($null -ne $distributionAndSecurityGroups) {
            $allGroupsData += $distributionAndSecurityGroups | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName        = $_.DisplayName
                    PrimarySmtpAddress = $_.PrimarySmtpAddress
                    DistinguishedName  = $_.DistinguishedName # Eindeutige ID für Get-Unique
                    GroupType          = $_.GroupType # z.B. Distribution, Security
                    RecipientTypeDetails = $_.RecipientTypeDetails # z.B. MailUniversalDistributionGroup, MailUniversalSecurityGroup
                    # Speichere das ursprüngliche Objekt für vollen Zugriff, falls später benötigt
                    OriginalObject     = $_
                }
            }
        } else {
            Write-Log "Keine Verteilergruppen oder E-Mail-aktivierte Sicherheitsgruppen gefunden oder Fehler beim Abruf." -Type Info # Kann auch Warning sein, je nach Erwartung
        }

        # Microsoft 365-Gruppen
        # Der Parameter -ResultSize Unlimited ist für die EXO V2/V3 Cmdlets nicht gültig/nötig.
        $unifiedGroups = Get-UnifiedGroup -ErrorAction SilentlyContinue
        if ($null -ne $unifiedGroups) {
            $allGroupsData += $unifiedGroups | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName        = $_.DisplayName
                    PrimarySmtpAddress = $_.PrimarySmtpAddress
                    DistinguishedName  = $_.ExternalDirectoryObjectId # Eindeutige ID für Get-Unique bei M365 Gruppen
                    GroupType          = "UnifiedGroup" # Explizit für M365-Gruppen setzen
                    RecipientTypeDetails = $_.RecipientTypeDetails # z.B. GroupMailbox
                    OriginalObject     = $_
                }
            }
        } else {
            Write-Log "Keine Microsoft 365-Gruppen gefunden oder Fehler beim Abruf." -Type Info
        }
        
        if ($allGroupsData.Count -eq 0) {
            Write-Log "Keine Gruppen (weder Verteiler/Sicherheit noch M365) gefunden." -Type Warning
            Update-StatusBar -Message "Keine Gruppen gefunden." -Type Warning
            return @()
        }

        # Sortieren nach Anzeigenamen. Get-Unique mit -AsString auf DistinguishedName, um Duplikate zu vermeiden.
        # Es ist sicherer, zuerst auf Eindeutigkeit zu prüfen und dann zu sortieren.
        # DistinguishedName sollte für DG/SecGroup eindeutig sein, ExternalDirectoryObjectId für UnifiedGroup.
        # Da wir beide in derselben Eigenschaft 'DistinguishedName' im PSCustomObject speichern, sollte dies funktionieren.
        $uniqueGroupsByDistinguishedName = $allGroupsData | Sort-Object -Property DistinguishedName -Unique
        $sortedGroups = $uniqueGroupsByDistinguishedName | Sort-Object DisplayName
        
        $statusMsg = "Abruf von $($sortedGroups.Count) Gruppen abgeschlossen."
        Write-Log $statusMsg -Type Success
        Update-StatusBar -Message $statusMsg -Type Success
        return $sortedGroups
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Abrufen der Gruppen: $errorMsg" -Type "Error"
        Update-StatusBar -Message "Fehler beim Abrufen der Gruppen." -Type Error
        # Die alte Statuszeilenaktualisierung wird durch Update-StatusBar ersetzt
        # if ($null -ne $script:txtStatus) {
        #     $script:txtStatus.Text = "Fehler beim Abrufen der Gruppen: $errorMsg"
        # }
        Show-MessageBox -Message "Fehler beim Abrufen der Gruppen: $errorMsg" -Title "Fehler Gruppenabruf" -Type Error
        return @() # Leeres Array bei Fehler zurückgeben
    }
}

# Funktion zum Aktualisieren der ComboBox cmbSelectExistingGroup
function Refresh-ExistingGroupsDropdown {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Aktualisiere Gruppenliste in ComboBox 'cmbSelectExistingGroup'" -Type "Info"
        Update-StatusBar -Message "Aktualisiere Gruppen-Dropdown..." -Type Info
        
        # Sicherstellen, dass das ComboBox-Element über den $script-Scope verfügbar ist.
        # Dies setzt voraus, dass $script:cmbSelectExistingGroup in Initialize-GroupsTab gefüllt wird.
        if ($null -eq $script:cmbSelectExistingGroup) {
            # Versuch, es abzurufen, falls nicht bereits im Skript-Scope, obwohl es dort sein sollte.
            $script:cmbSelectExistingGroup = Get-XamlElement -ElementName "cmbSelectExistingGroup"
            if ($null -eq $script:cmbSelectExistingGroup) {
                $errorMsg = "ComboBox 'cmbSelectExistingGroup' konnte nicht gefunden werden. Stellen Sie sicher, dass sie in Initialize-GroupsTab referenziert wird."
                Write-Log $errorMsg -Type "Error"
                Update-StatusBar -Message "Fehler: Dropdown nicht gefunden." -Type Error
                Show-MessageBox -Message $errorMsg -Title "UI Fehler" -Type Error
                return
            }
        }

        $script:cmbSelectExistingGroup.ItemsSource = $null # Vorherige Bindung / Elemente löschen
        $script:cmbSelectExistingGroup.Items.Clear()   # Sicherstellen, dass sie wirklich leer ist
        
        $groups = Get-AllGroupTypesAction # Diese Funktion hat bereits Logging und Status-Updates

        if ($null -ne $groups -and $groups.Count -gt 0) {
            foreach ($group in $groups) {
                $item = New-Object System.Windows.Controls.ComboBoxItem
                $item.Content = $group.DisplayName
                # Speichere das benutzerdefinierte Objekt (das OriginalObject enthält) in der Tag-Eigenschaft
                $item.Tag = $group 
                $item.ToolTip = "E-Mail: $($group.PrimarySmtpAddress)`nTyp: $($group.RecipientTypeDetails)"
                [void]$script:cmbSelectExistingGroup.Items.Add($item)
            }

            if ($script:cmbSelectExistingGroup.Items.Count -gt 0) {
                 # $script:cmbSelectExistingGroup.SelectedIndex = 0 # Ersten Eintrag auswählen
                 # Es ist oft besser, keine Vorauswahl zu treffen oder einen Platzhalter hinzuzufügen.
                 $script:cmbSelectExistingGroup.SelectedIndex = -1 # Keine Auswahl
            }
            
            $statusMsg = "Gruppenliste aktualisiert. $($groups.Count) Gruppen geladen."
            Write-Log $statusMsg -Type "Success"
            Update-StatusBar -Message $statusMsg -Type Success
            # Die alte Statuszeilenaktualisierung wird durch Update-StatusBar ersetzt
            # if ($null -ne $script:txtStatus) {
            #    $script:txtStatus.Text = $statusMsg
            # }
        }
        else {
            $statusMsg = "Keine Gruppen zum Anzeigen gefunden oder Fehler beim Laden."
            Write-Log $statusMsg -Type "Warning"
            Update-StatusBar -Message $statusMsg -Type Warning
            # Die alte Statuszeilenaktualisierung wird durch Update-StatusBar ersetzt
            # if ($null -ne $script:txtStatus) {
            #    $script:txtStatus.Text = $statusMsg
            # }
            # Optional: Platzhalter-Item hinzufügen, wenn keine Gruppen gefunden wurden
            $placeholderItem = New-Object System.Windows.Controls.ComboBoxItem
            $placeholderItem.Content = "-- Keine Gruppen gefunden --"
            $placeholderItem.IsEnabled = $false # Nicht auswählbar machen
            [void]$script:cmbSelectExistingGroup.Items.Add($placeholderItem)
            $script:cmbSelectExistingGroup.SelectedIndex = 0
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        $logErrorMsg = "Fehler beim Aktualisieren der Gruppenliste in 'cmbSelectExistingGroup': $errorMsg"
        Write-Log $logErrorMsg -Type "Error"
        Update-StatusBar -Message "Fehler beim Aktualisieren der Gruppenliste." -Type Error
        # Die alte Statuszeilenaktualisierung wird durch Update-StatusBar ersetzt
        # if ($null -ne $script:txtStatus) {
        #    $script:txtStatus.Text = $logErrorMsg
        # }
        Show-MessageBox -Message $logErrorMsg -Title "Fehler Dropdown-Aktualisierung" -Type Error
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
function RefreshSharedMailboxList {
    [CmdletBinding()]
    param()

    try {
        # Sicherstellen, dass das ComboBox-Element initialisiert ist
        if ($null -eq $script:cmbSharedMailboxSelect) {
            Write-Log "RefreshSharedMailboxList: Das Steuerelement 'cmbSharedMailboxSelect' ist nicht initialisiert." -Type "Warning"
            return $false
        }

        $script:cmbSharedMailboxSelect.Items.Clear()

        # Prüfen, ob eine Verbindung besteht
        if (-not $script:isConnected) {
            Write-Log "RefreshSharedMailboxList: Keine Verbindung zu Exchange Online. Aktualisierung abgebrochen." -Type "Warning"
            return $false
        }

        Write-Log "Beginne mit der Aktualisierung der Shared Mailbox-Liste." -Type "Info"
        
        # Abrufen aller Shared Mailboxen
        # Das -ErrorAction Stop ist wichtig, damit Fehler im Catch-Block landen
        $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop | 
                            Select-Object -ExpandProperty PrimarySmtpAddress
        
        if ($null -eq $sharedMailboxes -or $sharedMailboxes.Count -eq 0) {
            Write-Log "Keine Shared Mailboxen gefunden oder Fehler beim Abrufen." -Type "Info"
        } else {
            foreach ($mailboxAddress in $sharedMailboxes) {
                [void]$script:cmbSharedMailboxSelect.Items.Add($mailboxAddress)
            }
            Write-Log "$($script:cmbSharedMailboxSelect.Items.Count) Shared Mailboxen zur Liste hinzugefügt." -Type "Info"
        }
        
        if ($script:cmbSharedMailboxSelect.Items.Count -gt 0) {
            $script:cmbSharedMailboxSelect.SelectedIndex = 0
        }
        
        Write-Log "Shared Mailbox-Liste erfolgreich aktualisiert." -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Aktualisieren der Shared Mailbox-Liste: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Aktualisieren der Shared Mailbox-Liste: $errorMsg" # Für detaillierteres Logging
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
            Write-Log "Tab gewechselt zu: $($selectedTab.Header)"
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
    #   Outlook - Kernfunktionen (GroupBox Header="Outlook - Kernfunktionen")
    "chkAppsForOfficeEnabled",
    "chkFocusedInboxOn",
    "chkReadTrackingEnabled",
    "chkSendFromAliasEnabled",
    "chkLinkPreviewEnabled",
    "chkMessageRecallEnabled",
    "chkRecallReadMessagesEnabled",
    #   Kalenderoptionen (GroupBox Header="Kalenderoptionen")
    "cmbShortenEventScopeDefault",
    "txtDefaultMinutesToReduceShortEventsBy",
    "txtDefaultMinutesToReduceLongEventsBy",
    #   Outlook Web App (OWA) - Verhalten (GroupBox Header="Outlook Web App (OWA) - Verhalten")
    "chkAsyncSendEnabled",
    "chkOutlookGifPickerDisabled",
    "chkWebPushNotificationsDisabled",
    "chkWebSuggestedRepliesDisabled",
    #   Weitere Outlook & OWA Features (GroupBox Header="Weitere Outlook & OWA Features")
    "chkMessageRemindersEnabled",
    "chkOutlookPayEnabled",
    "chkOutlookTextPredictionDisabled",
    "chkEnableOutlookEvents", # War vorher unter Admin & Sicherheit -> Grundlegende Sicherheit
    #   Sitzungs-Timeout (OWA) (GroupBox Header="Sitzungs-Timeout (OWA)")
    "chkActivityBasedAuthenticationTimeoutEnabled",
    "txtActivityBasedAuthenticationTimeoutInterval", # War vorher cmbActivityBasedAuthenticationTimeoutInterval, angepasst an knownUIElements
    "chkActivityBasedAuthenticationTimeoutWithSingleSignOnEnabled",
    "chkOwaRedirectToOD4BThisUserEnabled", # War vorher unter Admin & Sicherheit -> OWA & Speicher
    "chkPublicComputersDetectionEnabled", # War vorher unter Admin & Sicherheit -> OWA & Speicher

    # Bookings - Grundeinstellungen Tab (TabItem Header="Bookings - Grundeinstellungen")
    #   Bookings - Allgemein
    "chkBookingsEnabled",
    "chkBookingsAddressEntryRestricted",
    "chkBookingsAuthEnabled",
    "chkBookingsCreationOfCustomQuestionsRestricted",
    "chkBookingsExposureOfStaffDetailsRestricted",
    "chkBookingsMembershipApprovalRequired",
    "chkBookingsNotesEntryRestricted",
    "chkBookingsPaymentsEnabled",
    "chkBookingsPhoneNumberEntryRestricted",
    "chkBookingsSocialSharingRestricted",
    "chkBookingsSmsMicrosoftEnabled",
    #   Bookings - Richtlinien & Kontrolle
    "chkBookingsSchedulingPolicyEnabled",
    "txtBookingsSchedulingPolicy",
    "chkBookingsStaffControlEnabled",
    "chkBookingsAvailabilityAndPricingEnabled",
    "chkBookingsCustomizationEnabled",
    #   Bookings - Daten & Freigabe
    "chkBookingsInternalSharingEnabled",
    "chkBookingsExternalSharingEnabled",
    "chkBookingsOnlinePaymentEnabled",
    "chkBookingsBusinessHoursEnabled",
    "chkBookingsCustomerDataUsageEnabled",
    "chkBookingsDataExportEnabled",
    "chkBookingsSearchEngineIndexDisabled",
    "chkBookingsPersonalDataCollectionAndUseConsentRequired",

    # Admin & Sicherheit Tab (TabItem Header="Admin & Sicherheit")
    #   Grundlegende Sicherheit & Compliance
    "chkAuditDisabled",
    "chkAutoEnableArchiveMailbox",
    "chkAutoExpandingArchive",
    "chkComplianceEnabled",
    "chkCustomerLockboxEnabled",
    "chkElcProcessingDisabled",
    #   MailTips
    "chkMailTipsExternalRecipientsTipsEnabled",
    "chkMailTipsGroupMetricsEnabled",
    "chkMailTipsLargeAudienceThreshold", # Bezieht sich auf die Aktivierung der Funktion, nicht den Wert selbst
    "txtMailTipsLargeAudienceThreshold", # Der Zahlenwert für den Threshold
    "chkMailTipsMailboxSourcedTipsEnabled",
    #   Konnektoren & Integrationen
    "chkConnectorsEnabled",
    "chkConnectorsActionableMessagesEnabled",
    "chkConnectorsEnabledForOutlook",
    "chkConnectorsEnabledForSharepoint",
    "chkConnectorsEnabledForTeams",
    "chkConnectorsEnabledForYammer",
    "chkSmtpActionableMessagesEnabled",
    #   Zugriffskontrolle & Weiterleitung (GroupBox Header="Zugriffskontrolle & Weiterleitung")
    "chkPublicFolderShowClientControl",
    "chkAutodiscoverPartialDirSync",
    "chkWorkspaceTenantEnabled",
    "chkAdditionalStorageProvidersBlocked", # War vorher unter Admin & Sicherheit -> OWA & Speicher
    #   Public Folder Quotas
    "txtDefaultPublicFolderMaxItemSize",
    "txtDefaultPublicFolderIssueWarningQuota",
    "txtDefaultPublicFolderProhibitPostQuota",
    #   Public Folder Settings (additional)
    "txtDefaultPublicFolderAgeLimit", 
    "txtDefaultPublicFolderDeletedItemRetention", 
    "txtDefaultPublicFolderMovedItemRetention", 
    "txtRemotePublicFolderMailboxes", 


    # Erweitert Tab (TabItem Header="Erweitert")
    #   Protokolle & Leistung (GroupBox Header="Protokolle & Leistung")
    "chkSIPEnabled",
    "chkRemotePublicFolderBlobsEnabled",
    "chkMapiHttpEnabled",
    "chkCalendarVersionStoreEnabled", 
    "chkEcRequiresTls", 
    #   Internationalisierung & Suche (GroupBox Header="Internationalisierung & Suche")
    "chkPreferredInternetCodePageForShiftJis", # Bezieht sich auf die Aktivierung der Funktion
    "txtPreferredInternetCodePageForShiftJis", # Der Zahlenwert
    "chkSearchQueryLanguage", # Bezieht sich auf die Aktivierung der Funktion
    "cmbSearchQueryLanguage", # Die Auswahl der Sprache
    #   Diverse Features (GroupBox Header="Diverse Features")
    "chkVisibilityEnabled",
    "chkOnlineMeetingsByDefaultEnabled",
    "chkDirectReportsGroupAutoCreationEnabled",
    "chkUnblockUnsafeSenderPromptEnabled",
    "chkExecutiveAttestation",
    "chkPDPLocationEnabled",
    #   Ratelimiting (GroupBox Header="Ratelimiting")
    "txtPowerShellMaxConcurrency",
    "txtPowerShellMaxCmdletQueueDepth",
    "txtPowerShellMaxCmdletsExecutionDuration"
)
#endregion EXOSettings Global Variables

#region EXOSettings Main Functions
# -----------------------------------------------
# EXOSettings Main Functions
# -----------------------------------------------
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
            "CheckBox" = @{
                "EventName" = "Click"
                "Handler" = {
                    param($sender, $e)
                    $checkBox = $sender
                    $checkBoxName = $checkBox.Name
                    if ($checkBoxName -like "chk*" -and $checkBoxName.Length -gt 3) {
                        $propertyName = $checkBoxName.Substring(3)
                        switch ($checkBoxName) {
                            default {}
                        }
                        $script:organizationConfigSettings[$propertyName] = $checkBox.IsChecked
                        Write-Log "Checkbox $checkBoxName wurde auf $($checkBox.IsChecked) gesetzt (Prop: $propertyName)" -Type "Info"
                    }
                }
            }
"ComboBox" = @{
                    "EventName" = "SelectionChanged"
                    "Handler" = {
                        param($sender, $e)
                        $comboBox = $sender
                        $comboBoxName = $comboBox.Name
                        $propertyName = $null

                        switch ($comboBoxName) {
                            "cmbShortenEventScopeDefault" { $propertyName = "ShortenEventScopeDefault" }
                            "cmbSearchQueryLanguage" { $propertyName = "SearchQueryLanguage" }
                            default {
                                if ($comboBoxName -like "cmb*" -and $comboBoxName.Length -gt 3) {
                                    $propertyName = $comboBoxName.Substring(3)
                                }
                            }
                        }

                        if (-not $propertyName) {
                            Write-Log "Kein PropertyName-Mapping für ComboBox $comboBoxName in Handler gefunden. Auswahländerung wird nicht verarbeitet." -Type "Warning"
                            return
                        }

                        if ($null -eq $comboBox.SelectedItem) {
                            $script:organizationConfigSettings.Remove($propertyName)
                            Write-Log "ComboBox $comboBoxName (Prop: $propertyName): Keine Auswahl. Wert entfernt." -Type "Info"
                            return
                        }

                        $selectedItem = $comboBox.SelectedItem
                        $selectedContent = $null # Wird für die Speicherung verwendet
                        $displayContent = $null  # Wird für das Logging verwendet

                        if ($selectedItem -is [System.Windows.Controls.ComboBoxItem]) {
                            $displayContent = $selectedItem.Content.ToString()
                            # Versuche, den Tag zu verwenden, wenn vorhanden und nicht leer, sonst Content
                            if ($null -ne $selectedItem.Tag -and $selectedItem.Tag.ToString() -ne "") {
                                $selectedContent = $selectedItem.Tag.ToString()
                            } else {
                                $selectedContent = $displayContent
                            }
                        } elseif ($selectedItem -is [string]) {
                            $displayContent = $selectedItem
                            $selectedContent = $selectedItem
                        } else {
                            if ($null -ne $selectedItem) { 
                                $displayContent = $selectedItem.ToString()
                                $selectedContent = $selectedItem.ToString() 
                            }
                        }

                        if ($null -eq $selectedContent) {
                            Write-Log "ComboBox $comboBoxName (Prop: $propertyName) Auswahl geändert, aber SelectedItem.Content/Tag ist null. Wert nicht geändert." -Type "Warning"
                            return
                        }

                        if ($displayContent -eq "KEINE DATEN" -or $selectedContent -eq "KEINE_DATEN_TAG_WERT") { # Annahme: KEINE_DATEN_TAG_WERT für den Tag des "KEINE DATEN" Items
                            $script:organizationConfigSettings.Remove($propertyName)
                            Write-Log "ComboBox $comboBoxName (Prop: $propertyName) Auswahl 'KEINE DATEN'. Wert entfernt." -Type "Info"
                            return
                        }

                        $valueToStore = $selectedContent # Standardmäßig den ausgewählten Inhalt/Tag speichern
                        switch ($comboBoxName) {
                            "cmbShortenEventScopeDefault" {
                                # Hier prüfen, ob der Wert numerisch ist, bevor konvertiert wird.
                                # Exchange erwartet für ShortenEventScopeDefault typischerweise 'None', 0, 15, 30 etc.
                                # Wenn der Tag direkt den Wert enthält (z.B. "0", "15", "None")
                                if ($selectedContent -match "^\d+$") { # Wenn es eine reine Zahl ist
                                    try { $valueToStore = [int]$selectedContent }
                                    catch { Write-Log "Fehler Konvertierung '$selectedContent' zu Int für $comboBoxName. Wert '$selectedContent' (string) gespeichert." -Type "Warning"; $valueToStore = $selectedContent }
                                } else {
                                    # Wenn es Text ist (z.B. "None"), als String speichern
                                    $valueToStore = $selectedContent
                                }
                            }
                            "cmbSearchQueryLanguage" {
                                # Der Tag sollte hier den Sprachcode enthalten, z.B. "en-US"
                                $valueToStore = $selectedContent 
                            }
                            # Weitere cmb-spezifische Logik hier...
                        }
                        
                        $script:organizationConfigSettings[$propertyName] = $valueToStore
                        Write-Log "ComboBox $comboBoxName (Prop: $propertyName) Auswahl geändert auf '$displayContent'. Gespeichert als '$valueToStore'." -Type "Info"
                    }
                }
                
                "TextBox" = @{
                "EventName" = "TextChanged"
                "Handler" = {
                    param($sender, $e)
                    $textBox = $sender
                    $textBoxName = $textBox.Name
                    $currentText = $textBox.Text
                    $propertyName = $null
                    $valueToStore = $currentText

                    switch ($textBoxName) {
                        "txtPowerShellMaxConcurrency" {
                            $propertyName = "PowerShellMaxConcurrency"
                            if ($currentText -eq "" -or $currentText -eq "KEINE DATEN") { $script:organizationConfigSettings.Remove($propertyName); Write-Log "$textBoxName geleert/KEINE DATEN, $propertyName entfernt." -Type "Info"; return }
                            if ([int]::TryParse($currentText, [ref]$null)) { $valueToStore = [int]$currentText } else { Write-Log "Ungültiger Int für ${textBoxName}: '$currentText'" -Type "Warning"; return }
                        }
                        "txtPowerShellMaxCmdletQueueDepth" {
                            $propertyName = "PowerShellMaxCmdletQueueDepth"
                            if ($currentText -eq "" -or $currentText -eq "KEINE DATEN") { $script:organizationConfigSettings.Remove($propertyName); Write-Log "$textBoxName geleert/KEINE DATEN, $propertyName entfernt." -Type "Info"; return }
                            if ([int]::TryParse($currentText, [ref]$null)) { $valueToStore = [int]$currentText } else { Write-Log "Ungültiger Int für ${textBoxName}: '$currentText'" -Type "Warning"; return }
                        }
                        "txtPowerShellMaxCmdletsExecutionDuration" {
                            $propertyName = "PowerShellMaxCmdletsExecutionDuration"
                            if ($currentText -eq "" -or $currentText -eq "KEINE DATEN") { $script:organizationConfigSettings.Remove($propertyName); Write-Log "$textBoxName geleert/KEINE DATEN, $propertyName entfernt." -Type "Info"; return }
                            if ([int]::TryParse($currentText, [ref]$null)) { $valueToStore = [int]$currentText } else { Write-Log "Ungültiger Int für ${textBoxName}: '$currentText'" -Type "Warning"; return }
                        }
                        "txtDefaultMinutesToReduceShortEventsBy" { 
                            $propertyName = "DefaultMinutesToReduceShortEventsBy"
                            if ($currentText -eq "" -or $currentText -eq "KEINE DATEN") { $script:organizationConfigSettings.Remove($propertyName); Write-Log "$textBoxName geleert/KEINE DATEN, $propertyName entfernt." -Type "Info"; return }
                            if ([int]::TryParse($currentText, [ref]$null)) { $valueToStore = [int]$currentText } else { Write-Log "Ungültiger Int für ${textBoxName}: '$currentText'" -Type "Warning"; return }
                        }
                        "txtDefaultMinutesToReduceLongEventsBy" { 
                            $propertyName = "DefaultMinutesToReduceLongEventsBy"
                            if ($currentText -eq "" -or $currentText -eq "KEINE DATEN") { $script:organizationConfigSettings.Remove($propertyName); Write-Log "$textBoxName geleert/KEINE DATEN, $propertyName entfernt." -Type "Info"; return }
                            if ([int]::TryParse($currentText, [ref]$null)) { $valueToStore = [int]$currentText } else { Write-Log "Ungültiger Int für ${textBoxName}: '$currentText'" -Type "Warning"; return }
                        }
                        "txtActivityBasedAuthenticationTimeoutInterval" { 
                            $propertyName = "ActivityBasedAuthenticationTimeoutInterval"
                            if ($currentText -eq "" -or $currentText -eq "KEINE DATEN") { $script:organizationConfigSettings.Remove($propertyName); Write-Log "$textBoxName geleert/KEINE DATEN, $propertyName entfernt." -Type "Info"; return }
                        }
                        "txtMailTipsLargeAudienceThreshold" { 
                            $propertyName = "MailTipsLargeAudienceThreshold"
                            if ($currentText -eq "" -or $currentText -eq "KEINE DATEN") { $script:organizationConfigSettings.Remove($propertyName); Write-Log "$textBoxName geleert/KEINE DATEN, $propertyName entfernt." -Type "Info"; return }
                            if ([int]::TryParse($currentText, [ref]$null)) { $valueToStore = [int]$currentText } else { Write-Log "Ungültiger Int für ${textBoxName}: '$currentText'" -Type "Warning"; return }
                        }
                        "txtPreferredInternetCodePageForShiftJis" {
                            $propertyName = "PreferredInternetCodePageForShiftJis"
                            if ($currentText -eq "" -or $currentText -eq "KEINE DATEN") { $script:organizationConfigSettings.Remove($propertyName); Write-Log "$textBoxName geleert/KEINE DATEN, $propertyName entfernt." -Type "Info"; return }
                            if ([int]::TryParse($currentText, [ref]$null)) { $valueToStore = [int]$currentText } else { Write-Log "Ungültiger Int für ${textBoxName}: '$currentText'" -Type "Warning"; return }
                        }
                        default {
                            if ($textBoxName -like "txt*" -and $textBoxName.Length -gt 3) {
                                $propertyName = $textBoxName.Substring(3)
                                if ($currentText -eq "" -or $currentText -eq "KEINE DATEN") {
                                    if ($script:organizationConfigSettings.ContainsKey($propertyName)) {
                                        $script:organizationConfigSettings.Remove($propertyName)
                                        Write-Log "Generische TextBox $textBoxName geleert/KEINE DATEN, $propertyName entfernt." -Type "Info"
                                    }
                                    return 
                                }
                            }
                        }
                    }

                    if ($propertyName) {
                        $script:organizationConfigSettings[$propertyName] = $valueToStore
                        Write-Log "TextBox $textBoxName Text geändert zu '$currentText' (Prop: $propertyName, Wert: $valueToStore)" -Type "Info"
                    } elseif ($textBoxName -like "txt*" -and $textBoxName.Length -gt 3) {} else {}
                }
            }
        }

        $tabEXOSettings = Get-XamlElement -ElementName "tabEXOSettings"
        if ($null -eq $tabEXOSettings) { throw "TabItem 'tabEXOSettings' nicht gefunden" }
        $tabOrgSettings = Get-XamlElement -ElementName "tabOrgSettings"
        if ($null -eq $tabOrgSettings) { throw "Inneres TabControl 'tabOrgSettings' nicht gefunden" }
        $tabEXOSettings.Visibility = [System.Windows.Visibility]::Visible
        $registeredControls = @{ "CheckBox" = 0; "ComboBox" = 0; "TextBox" = 0 }

        foreach ($elementName in $script:knownUIElements) {
            $element = Get-XamlElement -ElementName $elementName
            if ($null -eq $element -and $null -ne $tabEXOSettings) { try { $element = $tabEXOSettings.FindName($elementName) } catch {} }
            if ($null -eq $element -and $null -ne $tabOrgSettings) { try { $element = $tabOrgSettings.FindName($elementName) } catch {} }

            if ($null -ne $element) {
                $controlType = $null; $propertyNameForEnablingLogic = $null
                if ($element -is [System.Windows.Controls.CheckBox]) { $controlType = "CheckBox"; if ($elementName -like "chk*" -and $elementName.Length -gt 3) { $propertyNameForEnablingLogic = $elementName.Substring(3) } }
                elseif ($element -is [System.Windows.Controls.ComboBox]) { $controlType = "ComboBox"; if ($elementName -like "cmb*" -and $elementName.Length -gt 3) { $propertyNameForEnablingLogic = $elementName.Substring(3) } }
                elseif ($element -is [System.Windows.Controls.TextBox]) { $controlType = "TextBox"; if ($elementName -like "txt*" -and $elementName.Length -gt 3) { $propertyNameForEnablingLogic = $elementName.Substring(3) } }

                if ($controlType -and $controlHandlers.ContainsKey($controlType)) {
                    $eventHandlerInfo = $controlHandlers[$controlType]
                    Register-EventHandler -Control $element -Handler $eventHandlerInfo["Handler"] -ControlName $elementName -EventName $eventHandlerInfo["EventName"]
                    $registeredControls[$controlType]++
                    switch ($elementName) {
                        "txtMailTipsLargeAudienceThreshold" { $propertyNameForEnablingLogic = "MailTipsLargeAudienceThreshold" }
                        "txtActivityBasedAuthenticationTimeoutInterval" { $propertyNameForEnablingLogic = "ActivityBasedAuthenticationTimeoutInterval" }
                        "cmbSearchQueryLanguage" { $propertyNameForEnablingLogic = "SearchQueryLanguage" }
                        "cmbShortenEventScopeDefault" { $propertyNameForEnablingLogic = "ShortenEventScopeDefault" }
                    }
                    if ($propertyNameForEnablingLogic) {
                        if ($null -ne $script:currentOrganizationConfig) {
                            if ($script:currentOrganizationConfig.PSObject.Properties.Name -contains $propertyNameForEnablingLogic) { $element.IsEnabled = $true }
                            else { $element.IsEnabled = $false; Write-Log "$controlType '$elementName' deaktiviert (Eigenschaft '$propertyNameForEnablingLogic' nicht in OrgConfig)." -Type "Info" }
                        } else { $element.IsEnabled = $true; Write-Log "$controlType '$elementName' aktiviert (OrgConfig noch nicht geladen)." -Type "Debug" }
                    } else { $element.IsEnabled = $true; Write-Log "Für $controlType '$elementName' kein PropertyName für Aktivierungslogik. 'IsEnabled'=$true." -Type "Debug" }
                } else { Write-Log "Kein Handler für '$($element.GetType().Name)' oder '$elementName' definiert." -Type "Debug" }
            } else { Write-Log "Element '$elementName' nicht in XAML gefunden (Handler Registrierung)." -Type "Warning" }
        }
        Write-Log "Registrierte Controls: CheckBoxes: $($registeredControls['CheckBox']), ComboBoxes: $($registeredControls['ComboBox']), TextBoxes: $($registeredControls['TextBox'])" -Type "Info"
        return $true
    } catch { Write-Log "Fehler in Initialize-OrganizationConfigControls: $($_.Exception.Message)" -Type "Error"; Write-Log $_.Exception.StackTrace -Type "Error"; return $false }
}

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

        # Event-Handler für "Einstellungen exportieren (html)" Button
        $btnExportOrganizationConfig = Get-XamlElement -ElementName "btnExportOrganizationConfig"
        if ($null -ne $btnExportOrganizationConfig) {
            $btnExportOrganizationConfig.Add_Click({
                Write-Log "Button 'btnExportOrganizationConfig' geklickt." -Type "Info"
                # TODO: Implementiere die Funktion Export-OrganizationConfigurationAsHtml
                # Export-OrganizationConfigurationAsHtml
                Show-NotImplementedDialog -FeatureName "Export der OrganizationConfig als HTML"
            })
            Write-Log "EXOSettingsTab: btnExportOrganizationConfig Handler registriert." -Type "Debug"
        } else { Write-Log "EXOSettingsTab: btnExportOrganizationConfig nicht gefunden." -Type "Warning"; $success = $false }


        foreach ($elementName in $script:knownUIElements) {
            $element = Get-XamlElement -ElementName $elementName
            if ($null -ne $element) {
                # Stelle sicher, dass das Element sichtbar ist
                $element.Visibility = [System.Windows.Visibility]::Visible
                # Write-Log "EXOSettingsTab: Element '$elementName' auf sichtbar gesetzt." -Type "Debug" # Kann sehr verbose sein
            } else {
                Write-Log "EXOSettingsTab: Element '$elementName' aus knownUIElements nicht gefunden in XAML." -Type "Warning"
                # $success = $false # Nicht unbedingt ein Fehler für den Tab-Init, wenn ein Control fehlt,
                                     # aber könnte auf Inkonsistenz zwischen XAML und knownUIElements hinweisen.
            }
        }
        Write-Log "EXOSettingsTab: Sichtbarkeit für Elemente in knownUIElements überprüft." -Type "Debug"


        # Initialize all UI controls for the OrganizationConfig tab (Checkboxen, Textboxen etc. innerhalb von tabOrgSettings)
        Write-Log "EXOSettingsTab: Rufe Initialize-OrganizationConfigControls auf..." -Type "Debug"
        $controlsInitResult = Initialize-OrganizationConfigControls
        Write-Log "EXOSettingsTab: Initialize-OrganizationConfigControls Ergebnis: $controlsInitResult" -Type "Debug"
        if (-not $controlsInitResult) { $success = $false } # Wenn Controls nicht initialisiert werden können, ist der Tab fehlerhaft

        # Stelle sicher, dass der Tab selbst sichtbar ist
        if ($null -ne $script:tabEXOSettings) {
            $script:tabEXOSettings.Visibility = [System.Windows.Visibility]::Visible
            Write-Log "EXOSettingsTab: Tab auf sichtbar gesetzt." -Type "Debug"
        }

        # Stelle sicher, dass das TabItem "tabEXOSettings" im Haupt-TabControl sichtbar ist
        $mainTabControl = Get-XamlElement -ElementName "mainTabControl"
        if ($null -ne $mainTabControl) {
            $exoSettingsTabItem = $mainTabControl.Items | Where-Object { $_.Name -eq "tabEXOSettings" }
            if ($null -ne $exoSettingsTabItem) {
                $exoSettingsTabItem.Visibility = [System.Windows.Visibility]::Visible
                Write-Log "EXOSettingsTab: TabItem 'tabEXOSettings' im mainTabControl auf sichtbar gesetzt." -Type "Debug"
            } else {
                Write-Log "EXOSettingsTab: TabItem 'tabEXOSettings' nicht im mainTabControl gefunden." -Type "Warning"
            }
        } else {
            Write-Log "EXOSettingsTab: Haupt-TabControl 'mainTabControl' nicht gefunden." -Type "Warning"
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

#region EXOSettings Organization Config Management
function Get-CurrentOrganizationConfig {
    [CmdletBinding()]
    param()
    try {
        if (-not (Confirm-ExchangeConnection)) {
            Show-MessageBox -Message "Bitte verbinden Sie sich zuerst mit Exchange Online." -Title "Nicht verbunden" -Type "Warning"
            if ($null -ne $script:txtStatus) { $script:txtStatus.Text = "Nicht verbunden." }
            return
        }
        Write-Log "Rufe aktuelle Organisationseinstellungen ab..." -Type "Info"
        if ($null -ne $script:txtStatus) { $script:txtStatus.Text = "Lade Organisationseinstellungen..." }
        $script:currentOrganizationConfig = Get-OrganizationConfig -ErrorAction Stop
        if ($null -eq $script:currentOrganizationConfig) { throw "Get-OrganizationConfig lieferte keine Daten." }
        Write-Log "Organisationseinstellungen erfolgreich abgerufen." -Type "Info"
        $configProperties = $script:currentOrganizationConfig.PSObject.Properties | ForEach-Object { $_.Name }
        $script:organizationConfigSettings = @{}
        Write-Log "Aktualisiere UI-Elemente und Aktivierungsstatus..." -Type "Debug"

        foreach ($elementName in $script:knownUIElements) {
            $element = $null
            if ($null -ne $script:tabEXOSettings) { $element = $script:tabEXOSettings.FindName($elementName) }
            if ($null -eq $element) { $element = Get-XamlElement -ElementName $elementName }

            if ($null -ne $element) {
                $elementType = "Unknown"
                if ($element -is [System.Windows.Controls.CheckBox]) { $elementType = "CheckBox" }
                elseif ($element -is [System.Windows.Controls.ComboBox]) { $elementType = "ComboBox" }
                elseif ($element -is [System.Windows.Controls.TextBox]) { $elementType = "TextBox" }

                if ($elementType -eq "Unknown") { Write-Log "Element '$elementName' unbekannter Typ." -Type "Debug"; continue }

                $propertyName = $null; $derivedPropertyName = $null
                if ($elementName.StartsWith("chk") -and $elementName.Length -gt 3) { $derivedPropertyName = $elementName.Substring(3) }
                elseif ($elementName.StartsWith("cmb") -and $elementName.Length -gt 3) { $derivedPropertyName = $elementName.Substring(3) }
                elseif ($elementName.StartsWith("txt") -and $elementName.Length -gt 3) { $derivedPropertyName = $elementName.Substring(3) }

                switch ($elementName) {
                    "txtActivityBasedAuthenticationTimeoutInterval" { $propertyName = "ActivityBasedAuthenticationTimeoutInterval"; break }
                    "txtMailTipsLargeAudienceThreshold" { $propertyName = "MailTipsLargeAudienceThreshold"; break }
                    "cmbSearchQueryLanguage" { $propertyName = "SearchQueryLanguage"; break }
                    default { $propertyName = $derivedPropertyName }
                }
                if ($elementType -eq "CheckBox") {
                    switch ($elementName) {
                        "chkMailTipsLargeAudienceThreshold" { $propertyName = "MailTipsLargeAudienceThreshold"; break }
                        "chkPreferredInternetCodePageForShiftJis" { $propertyName = "PreferredInternetCodePageForShiftJis"; break }
                        "chkSearchQueryLanguage" { $propertyName = "SearchQueryLanguage"; break }
                    }
                }

                if (-not $propertyName) { Write-Log "Kein OrgConfig PropertyName für UI '$elementName' ($elementType). Überspringe." -Type "Warning"; $element.IsEnabled = $true; continue }

                if ($configProperties -contains $propertyName) {
                    $element.IsEnabled = $true
                    $valueFromConfig = $script:currentOrganizationConfig.$propertyName
                    Write-Log "Eig. '$propertyName' für UI '$elementName' in OrgConfig gefunden. Wert: '$valueFromConfig'" -Type "Debug"
                    try {
                        switch ($elementType) {
                            "CheckBox" {
                                $boolValue = $false
                                if ($null -ne $valueFromConfig) {
                                    if ($valueFromConfig -is [bool]) { $boolValue = $valueFromConfig }
                                    elseif ($valueFromConfig -is [string] -and ($valueFromConfig.ToLower() -eq 'true' -or $valueFromConfig.ToLower() -eq 'false')) { $boolValue = [bool]::Parse($valueFromConfig) }
                                    else {
                                        if ($elementName -eq "chkMailTipsLargeAudienceThreshold" -or $elementName -eq "chkPreferredInternetCodePageForShiftJis") { try { if ([int]$valueFromConfig -ne 0) { $boolValue = $true } else { $boolValue = $false } } catch { $boolValue = $false } }
                                        elseif ($elementName -eq "chkSearchQueryLanguage") { if (-not [string]::IsNullOrEmpty($valueFromConfig.ToString())) { $boolValue = $true } else { $boolValue = $false } }
                                        else { try { $boolValue = [System.Convert]::ToBoolean($valueFromConfig) } catch { Write-Log "Warn: Konnte '$valueFromConfig' für CheckBox '$elementName' nicht in Boolean konvertieren. $false." -Type "Warning" } }
                                    }
                                }
                                $element.IsChecked = $boolValue
                                if ($null -ne $valueFromConfig) { $script:organizationConfigSettings[$propertyName] = $boolValue } else { $script:organizationConfigSettings.Remove($propertyName) | Out-Null }
                                Write-Log "CheckBox '$elementName' (Prop: $propertyName) gesetzt auf '$boolValue'." -Type "Info"
                            }
                            "ComboBox" {
                                $currentConfigValue = $script:currentOrganizationConfig.$propertyName
                                $configValueAsString = if ($null -ne $currentConfigValue) { $currentConfigValue.ToString() } else { $null }
                                $itemFound = $false; $uiSelectedItem = $null
                                if ($null -ne $configValueAsString) {
                                    foreach ($itemObject in $element.Items) {
                                        $itemContentString = if ($itemObject -is [System.Windows.Controls.ComboBoxItem]) { $itemObject.Content.ToString() } else { $itemObject.ToString() }
                                        if ($itemContentString.Equals($configValueAsString, [System.StringComparison]::OrdinalIgnoreCase)) { $uiSelectedItem = $itemObject; $itemFound = $true; break }
                                    }
                                }
                                if ($itemFound) {
                                    $element.SelectedItem = $uiSelectedItem
                                    $selectedContentString = if ($element.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) { $element.SelectedItem.Content.ToString() } else { $element.SelectedItem.ToString() }
                                    $script:organizationConfigSettings[$propertyName] = $selectedContentString
                                } else {
                                    $keineDatenItem = $null; foreach ($item in $element.Items) { if (($item -is [System.Windows.Controls.ComboBoxItem] -and $item.Content.ToString() -eq "KEINE DATEN") -or ($item.ToString() -eq "KEINE DATEN")) { $keineDatenItem = $item; break } }
                                    if ($null -ne $keineDatenItem) { $element.SelectedItem = $keineDatenItem; Write-Log "ComboBox '$elementName': Wert '$configValueAsString' nicht gefunden/null. 'KEINE DATEN' ausgewählt." -Type "Info" }
                                    else { if ($element.Items.Count -gt 0) { $element.SelectedIndex = 0 } else { $element.SelectedItem = $null }; Write-Log "ComboBox '$elementName': Wert '$configValueAsString' nicht gefunden/null & 'KEINE DATEN' Item nicht da. Fallback." -Type "Debug" }
                                    $script:organizationConfigSettings.Remove($propertyName) | Out-Null
                                }
                                Write-Log "ComboBox '$elementName' (Prop: $propertyName) UI auf '$($element.Text)' (OrgConfig: '$configValueAsString')." -Type "Info"
                            }
                            "TextBox" {
                                if ($null -ne $valueFromConfig) {
                                    if ($valueFromConfig -is [array]) { $element.Text = $valueFromConfig -join ", " }
                                    else { $element.Text = $valueFromConfig.ToString() }
                                    $script:organizationConfigSettings[$propertyName] = $valueFromConfig
                                } else {
                                    $element.Text = "KEINE DATEN"
                                    $script:organizationConfigSettings.Remove($propertyName) | Out-Null
                                }
                                Write-Log "TextBox '$elementName' (Prop: $propertyName) gesetzt auf '$($element.Text)'." -Type "Info"
                            }
                        }
                    } catch { Write-Log "Fehler Setzen UI '$elementName' (Prop: $propertyName): $($_.Exception.Message)" -Type "Error" }
                } else {
                    $element.IsEnabled = $false
                    Write-Log "Eig. '$propertyName' für UI '$elementName' NICHT in OrgConfig. Deaktiviert." -Type "Info"
                    switch ($elementType) {
                        "CheckBox" { $element.IsChecked = $false }
                        "ComboBox" { 
                            $keineDatenItem = $null; foreach ($item in $element.Items) { if (($item -is [System.Windows.Controls.ComboBoxItem] -and $item.Content.ToString() -eq "KEINE DATEN") -or ($item.ToString() -eq "KEINE DATEN")) { $keineDatenItem = $item; break } }
                            if ($null -ne $keineDatenItem) { $element.SelectedItem = $keineDatenItem } else { if ($element.Items.Count -gt 0) { $element.SelectedIndex = 0 } else { $element.SelectedItem = $null } }
                        }
                        "TextBox"  { $element.Text = "KEINE DATEN" }
                    }
                    $script:organizationConfigSettings.Remove($propertyName) | Out-Null
                }
            } else { Write-Log "Element '$elementName' nicht in XAML (UI Update)." -Type "Debug" }
        }
        if ($null -ne $script:txtStatus) { $script:txtStatus.Text = "Organisationseinstellungen erfolgreich geladen." }
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Schwerer Fehler Get-CurrentOrganizationConfig: $errorMsg" -Type "Error"; Write-Log $_.Exception.StackTrace -Type "Error"
        if ($null -ne $script:txtStatus) { $script:txtStatus.Text = "Fehler Laden Organisationseinstellungen: $errorMsg" }
        Show-MessageBox -Message "Fehler Laden Organisationseinstellungen: $errorMsg" -Title "Fehler" -Type "Error"
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

        $paramsToSet = @{}

        # Behandle CheckBoxen
        # $script:knownUIElements sollte alle relevanten UI-Element-Namen enthalten
        foreach ($elementName in $script:knownUIElements) {
            $element = Get-XamlElement -ElementName $elementName
            if ($null -ne $element -and $element -is [System.Windows.Controls.CheckBox] -and $elementName -like "chk*") {
                $propertyName = $elementName.Substring(3)
                # Spezielle Checkboxen, die andere Controls steuern, werden später explizit behandelt.
                # Hier werden alle booleschen Werte erfasst.
                $paramsToSet[$propertyName] = $element.IsChecked
            }
        }

        # Behandle ComboBoxen
        # Spezialfall: ActivityBasedAuthenticationTimeoutInterval (gesteuert von chkActivityBasedAuthenticationTimeoutEnabled)
        if ($paramsToSet.ContainsKey("ActivityBasedAuthenticationTimeoutEnabled") -and $paramsToSet["ActivityBasedAuthenticationTimeoutEnabled"] -eq $true) {
            $cmbActivityTimeout = Get-XamlElement -ElementName "cmbActivityBasedAuthenticationTimeoutInterval"
            if ($null -ne $cmbActivityTimeout -and $null -ne $cmbActivityTimeout.SelectedItem) {
                $selectedText = $cmbActivityTimeout.SelectedItem.Content.ToString()
                $timeoutValue = ($selectedText -split ' ')[0] # Extrahiert "HH:MM:SS"
                if ($timeoutValue -match "^\d{2}:\d{2}:\d{2}$") {
                    $paramsToSet["ActivityBasedAuthenticationTimeoutInterval"] = $timeoutValue
                } else {
                    Write-Log "Ungültiges Format für ActivityBasedAuthenticationTimeoutInterval: '$selectedText'. Parameter wird nicht gesetzt." -Type "Warning"
                    # Entfernen, falls es durch eine vorherige Logik (falls vorhanden) gesetzt wurde
                    $paramsToSet.Remove("ActivityBasedAuthenticationTimeoutInterval") | Out-Null
                }
            } else {
                # Wenn Checkbox aktiviert ist, aber kein Intervall ausgewählt, ist das ein ungültiger Zustand für Set-OrganizationConfig
                Show-MessageBox -Message "Wenn 'Inaktive Sitzungs-Timeout aktivieren' ausgewählt ist, muss ein Timeout-Intervall angegeben werden." -Title "Fehlende Eingabe" -Type "Error"
                if ($null -ne $script:txtStatus) {
                    $script:txtStatus.Text = "Fehler: Timeout-Intervall fehlt."
                }
                return
            }
        } else {
            # Wenn ActivityBasedAuthenticationTimeoutEnabled false ist, den Intervall-Parameter entfernen
            $paramsToSet.Remove("ActivityBasedAuthenticationTimeoutInterval") | Out-Null
        }

        # Allgemeine ComboBoxen
        $comboBoxMappings = @{
            "cmbShortenEventScopeDefault" = "ShortenEventScopeDefault"
            "cmbLargeAudienceThreshold" = "MailTipsLargeAudienceThreshold"
            "cmbInformationBarrierMode" = "InformationBarrierMode"
            "cmbEwsAppAccessPolicy" = "EwsApplicationAccessPolicy"
            "cmbOfficeFeatures" = "OfficeFeatures"
            # cmbSearchQueryLanguage wird speziell behandelt
        }

        foreach ($comboBoxName in $comboBoxMappings.Keys) {
            $comboBox = Get-XamlElement -ElementName $comboBoxName
            $propertyName = $comboBoxMappings[$comboBoxName]

            if ($null -ne $comboBox -and $null -ne $comboBox.SelectedItem) {
                $selectedValue = $comboBox.SelectedItem.Content.ToString()
                if ($comboBoxName -eq "cmbLargeAudienceThreshold") {
                    if ([int]::TryParse($selectedValue, [ref]$null)) {
                        $paramsToSet[$propertyName] = [int]$selectedValue
                    } else {
                        throw "Der Wert für $propertyName ('$selectedValue') muss eine ganze Zahl sein."
                    }
                } else {
                    $paramsToSet[$propertyName] = $selectedValue
                }
            } else {
                # Wenn nichts ausgewählt ist, Parameter nicht senden (implizit durch Nicht-Hinzufügen zu paramsToSet)
                # oder explizit entfernen, falls er anderweitig gesetzt wurde (hier nicht der Fall)
                $paramsToSet.Remove($propertyName) | Out-Null
            }
        }

        # Spezialfall: cmbSearchQueryLanguage (gesteuert von chkSearchQueryLanguage)
        $chkSearchQueryLanguage = Get-XamlElement -ElementName "chkSearchQueryLanguage"
        if ($null -ne $chkSearchQueryLanguage -and $chkSearchQueryLanguage.IsChecked -eq $true) {
            $cmbSearchQueryLanguage = Get-XamlElement -ElementName "cmbSearchQueryLanguage"
            if ($null -ne $cmbSearchQueryLanguage -and $null -ne $cmbSearchQueryLanguage.SelectedItem) {
                $paramsToSet["SearchQueryLanguage"] = $cmbSearchQueryLanguage.SelectedItem.Content.ToString()
            } else {
                # Checkbox ist an, aber nichts ausgewählt -> Parameter nicht senden
                $paramsToSet.Remove("SearchQueryLanguage") | Out-Null
            }
        } else {
            # Checkbox ist aus, Parameter nicht senden (ggf. von generischer Checkbox-Logik gesetzten booleschen Wert entfernen)
            $paramsToSet.Remove("SearchQueryLanguage") | Out-Null
        }


        # Behandle TextBoxen
        $textBoxMappings = @{
            "txtDefaultMinutesToReduceShortEventsBy" = "DefaultMinutesToReduceShortEventsBy"
            "txtDefaultMinutesToReduceLongEventsBy" = "DefaultMinutesToReduceLongEventsBy"
            "txtPowerShellMaxConcurrency" = "PowerShellMaxConcurrency"
            "txtPowerShellMaxCmdletQueueDepth" = "PowerShellMaxCmdletQueueDepth"
            "txtPowerShellMaxCmdletsExecutionDuration" = "PowerShellMaxCmdletsExecutionDuration"
            "txtDefaultAuthPolicy" = "DefaultAuthenticationPolicy"
            "txtHierAddressBookRoot" = "HierarchicalAddressBookRoot"
            # txtPreferredInternetCodePageForShiftJis wird speziell behandelt
        }

        $numericTextBoxesForOrgConfig = @(
            "txtDefaultMinutesToReduceShortEventsBy",
            "txtDefaultMinutesToReduceLongEventsBy",
            "txtPowerShellMaxConcurrency",
            "txtPowerShellMaxCmdletQueueDepth",
            "txtPowerShellMaxCmdletsExecutionDuration"
            # txtPreferredInternetCodePageForShiftJis ist auch numerisch, aber speziell behandelt
        )

        foreach ($textBoxName in $textBoxMappings.Keys) {
            $textBox = Get-XamlElement -ElementName $textBoxName
            $propertyName = $textBoxMappings[$textBoxName]

            if ($null -ne $textBox) {
                if (-not [string]::IsNullOrWhiteSpace($textBox.Text)) {
                    $textValue = $textBox.Text.Trim()
                    if ($numericTextBoxesForOrgConfig -contains $textBoxName) {
                        if ([int]::TryParse($textValue, [ref]$null)) {
                            $paramsToSet[$propertyName] = [int]$textValue
                        } else {
                            throw "Der Wert für $propertyName ('$textValue') muss eine ganze Zahl sein."
                        }
                    } else {
                        $paramsToSet[$propertyName] = $textValue
                    }
                } else {
                    # Leere Textbox -> Parameter nicht senden (implizit durch Nicht-Hinzufügen)
                    # oder explizit entfernen, falls er anderweitig gesetzt wurde
                    $paramsToSet.Remove($propertyName) | Out-Null
                }
            }
        }

        # Spezialfall: txtPreferredInternetCodePageForShiftJis (gesteuert von chkPreferredInternetCodePageForShiftJis)
        $chkPreferredInternetCodePageForShiftJis = Get-XamlElement -ElementName "chkPreferredInternetCodePageForShiftJis"
        if ($null -ne $chkPreferredInternetCodePageForShiftJis -and $chkPreferredInternetCodePageForShiftJis.IsChecked -eq $true) {
            $txtBox = Get-XamlElement -ElementName "txtPreferredInternetCodePageForShiftJis"
            if ($null -ne $txtBox -and -not [string]::IsNullOrWhiteSpace($txtBox.Text)) {
                $textValue = $txtBox.Text.Trim()
                if ([int]::TryParse($textValue, [ref]$null)) {
                    $paramsToSet["PreferredInternetCodePageForShiftJis"] = [int]$textValue
                } else {
                    throw "Der Wert für PreferredInternetCodePageForShiftJis ('$textValue') muss eine ganze Zahl sein."
                }
            } else {
                # Checkbox ist an, aber Textbox leer -> Parameter nicht senden
                $paramsToSet.Remove("PreferredInternetCodePageForShiftJis") | Out-Null
            }
        } else {
            # Checkbox ist aus, Parameter nicht senden (ggf. von generischer Checkbox-Logik gesetzten booleschen Wert entfernen)
            $paramsToSet.Remove("PreferredInternetCodePageForShiftJis") | Out-Null
        }

        # Parameter für Set-OrganizationConfig vorbereiten - nur die, die auch Werte haben
        # $paramsToSet enthält bereits die korrekten Parameter

        if ($paramsToSet.Count -eq 0) {
            Show-MessageBox -Message "Es wurden keine Änderungen zum Speichern gefunden." -Title "Keine Änderungen" -Type "Info"
            if ($null -ne $script:txtStatus) {
                $script:txtStatus.Text = "Keine Änderungen zum Speichern."
            }
            return
        }
        
        # Debug-Log alle Parameter, die gesendet werden
        Write-Log "Folgende Parameter werden an Set-OrganizationConfig übergeben:" -Type "Debug"
        foreach ($key in $paramsToSet.Keys | Sort-Object) {
            Write-Log "Parameter: $key = '$($paramsToSet[$key])' (Typ: $($paramsToSet[$key].GetType().Name))" -Type "Debug"
        }

        # Organisationseinstellungen aktualisieren
        Set-OrganizationConfig @paramsToSet -ErrorAction Stop

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
        
        # SaveFileDialog anzeigen, nur für HTML-Export
        $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveFileDialog.Filter = "HTML-Dateien (*.html)|*.html"
        $saveFileDialog.FileName = "ExchangeOnline_OrgConfig_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        $saveFileDialog.DefaultExt = ".html"
        $saveFileDialog.AddExtension = $true
        
        $result = $saveFileDialog.ShowDialog()
        if ($result -ne $true) {
            return # Benutzer hat Abbrechen gewählt
        }
        
        $exportPath = $saveFileDialog.FileName
        
        # Objekt für den Export vorbereiten, unerwünschte Eigenschaften entfernen
        $objectToExport = $script:currentOrganizationConfig | 
            Select-Object * -ExcludeProperty RunspaceId, PSComputerName, PSShowComputerName, PSSourceJobInstanceId, PSObject, PSCmdlet, PSPath, PSParentPath, PSChildName, PSDrive, PSProvider, PSTypeNames
        
        # HTML-Header für besseres Styling
        $htmlHeader = @"
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #f4f4f4; color: #333; }
    h1 { color: #005a9e; border-bottom: 2px solid #005a9e; padding-bottom: 10px; }
    table { 
        border-collapse: collapse; 
        width: 90%; 
        margin-top: 20px; 
        box-shadow: 0 2px 15px rgba(0,0,0,0.1); 
        background-color: #fff;
    }
    th, td { 
        text-align: left; 
        padding: 12px 15px; 
        border: 1px solid #ddd; 
    }
    th { 
        background-color: #0078d4; 
        color: white; 
        font-weight: bold; 
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    td:first-child { 
        font-weight: bold; 
        background-color: #f8f8f8; 
        width: 35%; /* Breite der Eigenschaftsspalte */
    }
    tr:nth-child(even) td:not(:first-child) { 
        background-color: #eef6fc; 
    }
    tr:hover td:not(:first-child) { 
        background-color: #dcf0ff; 
    }
</style>
"@
        $reportDate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
        $title = "Exchange Online Organisationseinstellungen - Exportiert am $reportDate"

        # Als HTML exportieren (Listenansicht ist für ein einzelnes Konfigurationsobjekt besser geeignet)
        $objectToExport | ConvertTo-Html -As List -Property * -Head $htmlHeader -Title $title -Body "<h1>$title</h1>" | Out-File -FilePath $exportPath -Encoding utf8
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organisationseinstellungen exportiert nach $exportPath"
        }
        Show-MessageBox -Message "Die Organisationseinstellungen wurden erfolgreich nach '$exportPath' exportiert." -Title "Export erfolgreich" -Type "Info"
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Exportieren der Organisationseinstellungen: $errorMsg" -Type "Error"
        Write-Log $_.Exception.StackTrace -Type "Error"
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
#region Initialize-GroupsTab

function Update-StatusBar {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateSet("Info", "Success", "Warning", "Error", "Debug")]
        [string]$Type = "Info",

        [Parameter(Mandatory = $false)]
        [System.Windows.Media.Brush]$OverrideForegroundColor
    )

    if ($null -eq $script:txtStatus) {
        Write-Log -Message "Update-StatusBar: Status-Textfeld (script:txtStatus) wurde nicht gefunden oder ist nicht initialisiert. Meldung: '$Message'" -Type Warning
        return
    }

    try {
        # Zugriff auf das UI-Element muss im UI-Thread erfolgen.
        # Wenn diese Funktion aus einem anderen Thread aufgerufen wird, $script:Form.Dispatcher.Invoke verwenden.
        # Für die meisten Anwendungsfälle in diesem Skript wird sie wahrscheinlich aus dem UI-Thread aufgerufen.
        
        $script:txtStatus.Text = $Message
        
        # Standardfarben basierend auf dem Typ definieren
        $foregroundColor = $null
        if ($PSBoundParameters.ContainsKey('OverrideForegroundColor') -and $null -ne $OverrideForegroundColor) {
            $foregroundColor = $OverrideForegroundColor
        } else {
            switch ($Type) {
                "Success" { $foregroundColor = [System.Windows.Media.Brushes]::Green }
                "Warning" { $foregroundColor = [System.Windows.Media.Brushes]::OrangeRed } # Oder DarkOrange
                "Error"   { $foregroundColor = [System.Windows.Media.Brushes]::Red }
                "Debug"   { $foregroundColor = [System.Windows.Media.Brushes]::BlueViolet } 
                "Info"    { 
                            # Für "Info" die Standard-Vordergrundfarbe verwenden (oft Schwarz)
                            # Oder, falls das Control eine spezifische Standardfarbe hat, diese beibehalten.
                            # Der Einfachheit halber setzen wir es auf Schwarz, wenn keine Override-Farbe angegeben ist.
                            $foregroundColor = [System.Windows.Media.Brushes]::Black 
                          }
                default   { $foregroundColor = [System.Windows.Media.Brushes]::Black }
            }
        }
        
        $script:txtStatus.Foreground = $foregroundColor
        
        # Loggen der Statusänderung (Typ Debug, um Logs nicht zu überfluten)
        Write-Log -Message "StatusBar aktualisiert: [$Type] $Message" -Type Debug 
        
    } catch {
        $errorDetail = $_.Exception.Message
        Write-Log -Message "Fehler in Update-StatusBar beim Setzen von Text/Farbe: $errorDetail. Ursprüngliche Meldung: '$Message'" -Type Error
        # Versuch, den Fehler selbst in der Statusleiste anzuzeigen (als Fallback)
        try {
            $script:txtStatus.Text = "Fehler in StatusBar: $errorDetail"
            $script:txtStatus.Foreground = [System.Windows.Media.Brushes]::Red
        } catch {
            # Wenn selbst das Setzen der Fehlermeldung fehlschlägt, ist das UI-Element möglicherweise nicht mehr gültig.
            Write-Log -Message "Konnte Fehler nicht in StatusBar anzeigen. txtStatus möglicherweise ungültig." -Type Error
        }
    }
}
function Get-ExoDistributionGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    Write-Log "Get-ExoDistributionGroup: Aktion gestartet für Identity '$Identity'." -Type Info
    Update-StatusBar -Message "Rufe Verteilergruppe '$Identity' ab..." -Type Info

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return $null
        }

        $group = Get-DistributionGroup -Identity $Identity -ErrorAction Stop
        
        if ($group) {
            Write-Log "Get-ExoDistributionGroup: Verteilergruppe '$($group.DisplayName)' (Identity: '$Identity') erfolgreich abgerufen." -Type Success
            Update-StatusBar -Message "Verteilergruppe '$($group.DisplayName)' erfolgreich abgerufen." -Type Success
            return $group
        } else {
            # Dieser Fall sollte durch ErrorAction Stop eigentlich nicht eintreten, aber zur Sicherheit
            Write-Log "Get-ExoDistributionGroup: Verteilergruppe mit Identity '$Identity' nicht gefunden (nach Get-DistributionGroup Aufruf)." -Type Warning
            Update-StatusBar -Message "Verteilergruppe '$Identity' nicht gefunden." -Type Warning
            return $null
        }

    } catch {
        $errMsg = $_.Exception.Message
        Write-Log "Get-ExoDistributionGroup: Fehler beim Abrufen der Verteilergruppe '$Identity': $errMsg" -Type Error
        Update-StatusBar -Message "Fehler beim Abrufen der Verteilergruppe '$Identity'." -Type Error
        return $null
    }
}

function Get-ExoDynamicDistributionGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    Write-Log "Get-ExoDynamicDistributionGroup: Aktion gestartet für Identity '$Identity'." -Type Info
    Update-StatusBar -Message "Rufe dynamische Verteilergruppe '$Identity' ab..." -Type Info

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return $null
        }

        $group = Get-DynamicDistributionGroup -Identity $Identity -ErrorAction Stop
        
        if ($group) {
            Write-Log "Get-ExoDynamicDistributionGroup: Dynamische Verteilergruppe '$($group.DisplayName)' (Identity: '$Identity') erfolgreich abgerufen." -Type Success
            Update-StatusBar -Message "Dynamische Verteilergruppe '$($group.DisplayName)' erfolgreich abgerufen." -Type Success
            return $group
        } else {
            # Dieser Fall sollte durch ErrorAction Stop eigentlich nicht eintreten, aber zur Sicherheit
            Write-Log "Get-ExoDynamicDistributionGroup: Dynamische Verteilergruppe mit Identity '$Identity' nicht gefunden (nach Get-DynamicDistributionGroup Aufruf)." -Type Warning
            Update-StatusBar -Message "Dynamische Verteilergruppe '$Identity' nicht gefunden." -Type Warning
            return $null
        }

    } catch {
        $errMsg = $_.Exception.Message
        Write-Log "Get-ExoDynamicDistributionGroup: Fehler beim Abrufen der dynamischen Verteilergruppe '$Identity': $errMsg" -Type Error
        Update-StatusBar -Message "Fehler beim Abrufen der dynamischen Verteilergruppe '$Identity'." -Type Error
        return $null
    }
}

function Get-UnifiedGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    Write-Log "Get-UnifiedGroup: Aktion gestartet für Identity '$Identity'." -Type Info # Korrigierter Log Funktionsname
    Update-StatusBar -Message "Rufe Microsoft 365-Gruppe '$Identity' ab..." -Type Info

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return $null
        }

        $group = Get-UnifiedGroup -Identity $Identity -ErrorAction Stop # EXO Cmdlet ist Get-UnifiedGroup
        
        if ($group) {
            Write-Log "Get-UnifiedGroup: Microsoft 365-Gruppe '$($group.DisplayName)' (Identity: '$Identity') erfolgreich abgerufen." -Type Success
            Update-StatusBar -Message "Microsoft 365-Gruppe '$($group.DisplayName)' erfolgreich abgerufen." -Type Success
            return $group
        } else {
            # Dieser Fall sollte durch ErrorAction Stop eigentlich nicht eintreten, aber zur Sicherheit
            Write-Log "Get-UnifiedGroup: Microsoft 365-Gruppe mit Identity '$Identity' nicht gefunden (nach Get-UnifiedGroup Aufruf)." -Type Warning
            Update-StatusBar -Message "Microsoft 365-Gruppe '$Identity' nicht gefunden." -Type Warning
            return $null
        }

    } catch {
        $errMsg = $_.Exception.Message
        Write-Log "Get-UnifiedGroup: Fehler beim Abrufen der Microsoft 365-Gruppe '$Identity': $errMsg" -Type Error
        Update-StatusBar -Message "Fehler beim Abrufen der Microsoft 365-Gruppe '$Identity'." -Type Error
        return $null
    }
}

function Get-ExoGroupSettingsAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity # Erwartet wird $groupObject.Identity oder eine ähnliche eindeutige ID
    )

    Write-Log "Get-ExoGroupSettingsAction: Aktion gestartet für Identity '$Identity'." -Type Info
    Update-StatusBar -Message "Lade Gruppeneinstellungen für '$Identity'..." -Type Info 

    $groupObjectForDetails = $null # Um Verwechslung mit $groupObject im Catch-Block zu vermeiden, falls es aus UI kommt

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return $null
        }

        # Das Gruppenobjekt basierend auf der übergebenen Identity abrufen, um RecipientTypeDetails und DisplayName zu bestimmen
        # Get-ExoGroupRecipientAction ist eine angenommene Hilfsfunktion, die verschiedene Gruppentypen handhaben kann
        $groupObjectForDetails = Get-ExoGroupRecipientAction -Identity $Identity
        if ($null -eq $groupObjectForDetails) {
            Write-Log "Get-ExoGroupSettingsAction: Gruppe mit Identity '$Identity' nicht gefunden oder Typ konnte nicht ermittelt werden." -Type Warning
            Show-MessageBox -Message "Gruppe '$Identity' nicht gefunden. Einstellungen können nicht geladen werden." -Title "Aktion fehlgeschlagen" -Type Warning
            Update-StatusBar -Message "Gruppe '$Identity' nicht gefunden." -Type Warning
            return $null
        }
        
        $groupDisplayName = $groupObjectForDetails.DisplayName
        $recipientType = $groupObjectForDetails.RecipientTypeDetails
        
        Write-Log "Get-ExoGroupSettingsAction: Rufe Einstellungen für Gruppe '$groupDisplayName' (Identity: '$Identity', Typ: '$recipientType') ab." -Type Info
        Update-StatusBar -Message "Lade Einstellungen für Gruppe '$groupDisplayName'..." -Type Info

        $detailsFromExo = $null
        # Die übergebene $Identity (oder die aufgelöste Identity des groupObjectForDetails) für den EXO-Aufruf verwenden
        $paramsForExoCall = @{ Identity = $groupObjectForDetails.Identity } 

        if ($recipientType -eq "GroupMailbox") { # Microsoft 365 Gruppe
            $rawExoObject = Get-UnifiedGroup @paramsForExoCall -ErrorAction SilentlyContinue # Fehler hier abfangen
            if ($rawExoObject) {
                $detailsFromExo = [PSCustomObject]@{
                    DisplayName                        = $rawExoObject.DisplayName
                    HiddenFromAddressListsEnabled      = $rawExoObject.HiddenFromExchangeAddressListsEnabled
                    RequireSenderAuthenticationEnabled = $rawExoObject.RequireSenderAuthenticationEnabled
                }
            }
        } elseif ($recipientType -in ("MailUniversalDistributionGroup", "MailNonUniversalGroup", "MailUniversalSecurityGroup")) { # Verteilerliste oder E-Mail-aktivierte Sicherheitsgruppe
            $rawExoObject = Get-DistributionGroup @paramsForExoCall -ErrorAction SilentlyContinue # Fehler hier abfangen
            if ($rawExoObject) {
                $detailsFromExo = [PSCustomObject]@{
                    DisplayName                        = $rawExoObject.DisplayName
                    HiddenFromAddressListsEnabled      = $rawExoObject.HiddenFromAddressListsEnabled
                    RequireSenderAuthenticationEnabled = $rawExoObject.RequireSenderAuthenticationEnabled
                }
            }
        } elseif ($recipientType -eq "DynamicDistributionGroup") {
             $rawExoObject = Get-DynamicDistributionGroup @paramsForExoCall -ErrorAction SilentlyContinue
             if ($rawExoObject) {
                $detailsFromExo = [PSCustomObject]@{
                    DisplayName                        = $rawExoObject.DisplayName
                    HiddenFromAddressListsEnabled      = $rawExoObject.HiddenFromAddressListsEnabled
                    RequireSenderAuthenticationEnabled = $rawExoObject.RequireSenderAuthenticationEnabled # Dynamische Gruppen haben dies auch
                }
            }
        } else {
            Write-Log "Get-ExoGroupSettingsAction: Nicht unterstützter Gruppentyp '$recipientType' für Gruppe '$groupDisplayName' (Identity '$Identity')." -Type Error
            Show-MessageBox -Message "Der Gruppentyp '$recipientType' für '$groupDisplayName' wird für das Abrufen von Einstellungen nicht unterstützt." -Title "Nicht unterstützter Typ" -Type Warning
            Update-StatusBar -Message "Nicht unterstützter Gruppentyp: $recipientType" -Type Warning
            return $null
        }

        if ($null -eq $detailsFromExo) {
            Write-Log "Get-ExoGroupSettingsAction: Konnte keine Detailinformationen für Gruppe '$groupDisplayName' (Identity: '$Identity', Typ: '$recipientType') abrufen. EXO-Cmdlet hat nichts zurückgegeben oder ein Fehler ist aufgetreten." -Type Error
            Show-MessageBox -Message "Keine Detailinformationen für Gruppe '$groupDisplayName' gefunden." -Title "Fehler" -Type Error
            Update-StatusBar -Message "Keine Details für '$groupDisplayName' gefunden." -Type Error
            return $null
        }

        $settingsResult = [PSCustomObject]@{
            HiddenFromAddressListsEnabled      = [bool]$detailsFromExo.HiddenFromAddressListsEnabled
            RequireSenderAuthenticationEnabled = [bool]$detailsFromExo.RequireSenderAuthenticationEnabled
            AllowExternalSenders               = -not ([bool]$detailsFromExo.RequireSenderAuthenticationEnabled) 
            DisplayName                        = $detailsFromExo.DisplayName
        }

        Write-Log "Get-ExoGroupSettingsAction: Einstellungen für '$($detailsFromExo.DisplayName)' erfolgreich abgerufen." -Type Success
        Update-StatusBar -Message "Einstellungen für '$($detailsFromExo.DisplayName)' geladen." -Type Success
        return $settingsResult

    } catch {
        $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Unbekannter Fehler beim Abrufen der Gruppeneinstellungen."
        $groupNameForErrorMsg = if ($groupObjectForDetails -and $groupObjectForDetails.DisplayName) { $groupObjectForDetails.DisplayName } else { $Identity }
        
        Write-Log "Fehler in Get-ExoGroupSettingsAction für Gruppe '$groupNameForErrorMsg' (Identity: '$Identity'): $errMsg" -Type Error
        if (Test-Path Function:\Log-Action) { Log-Action "Fehler Get-ExoGroupSettingsAction '$groupNameForErrorMsg': $errMsg" }
        Show-MessageBox -Message "Fehler beim Abrufen der Einstellungen für Gruppe '$groupNameForErrorMsg':`n$($_.Exception.Message)" -Title "Fehler bei Einstellungen" -Type Error
        Update-StatusBar -Message "Fehler beim Laden der Einstellungen für '$groupNameForErrorMsg'." -Type Error
        return $null
    }
}
function Update-ExoGroupSettingsAction {
    [CmdletBinding(SupportsShouldProcess = $true)] # Hinzugefügt SupportsShouldProcess
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [Parameter(Mandatory = $false)]
        [AllowNull()] 
        [bool]$HiddenFromAddressListsEnabled,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [bool]$RequireSenderAuthenticationEnabled,
        
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [bool]$AllowExternalSenders 
    )

    Write-Log "Update-ExoGroupSettingsAction: Aktion gestartet für Gruppe '$Identity'." -Type Info
    Update-StatusBar -Message "Aktualisiere Einstellungen für Gruppe '$Identity'..." -Type Info

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return $false
        }

        $groupObject = Get-ExoGroupRecipientAction -Identity $Identity
        if (-not $groupObject) {
            Update-StatusBar -Message "Gruppe '$Identity' nicht gefunden oder Typ konnte nicht ermittelt werden." -Type Error
            return $false
        }
        $groupDisplayName = $groupObject.DisplayName # Für $ShouldProcess Nachricht
        $recipientType = $groupObject.RecipientTypeDetails

        $params = @{}
        if ($PSBoundParameters.ContainsKey('HiddenFromAddressListsEnabled')) {
            $params.HiddenFromAddressListsEnabled = $HiddenFromAddressListsEnabled
        }
        
        # Logik für RequireSenderAuthenticationEnabled und AllowExternalSenders
        # AllowExternalSenders ist oft eine Umkehrung von RequireSenderAuthenticationEnabled
        # Wenn beide gesetzt sind, muss entschieden werden, welcher Vorrang hat oder wie sie kombiniert werden.
        # Annahme: Wenn AllowExternalSenders gesetzt ist, leitet es RequireSenderAuthenticationEnabled ab.
        # Wenn RequireSenderAuthenticationEnabled explizit gesetzt ist und AllowExternalSenders nicht, wird ersteres verwendet.

        if ($PSBoundParameters.ContainsKey('AllowExternalSenders')) {
            $params.RequireSenderAuthenticationEnabled = -not $AllowExternalSenders
            # Wenn auch RequireSenderAuthenticationEnabled gesetzt wurde, könnte dies überschrieben werden.
            # Um dies zu vermeiden, könnte man prüfen, ob RequireSenderAuthenticationEnabled auch gebunden ist und dann eine Logik anwenden.
            # Für Einfachheit: AllowExternalSenders überschreibt, wenn gesetzt.
            if ($PSBoundParameters.ContainsKey('RequireSenderAuthenticationEnabled') -and ($params.RequireSenderAuthenticationEnabled -ne $RequireSenderAuthenticationEnabled) ) {
                 Write-Log "Update-ExoGroupSettingsAction: Konflikt zwischen AllowExternalSenders und RequireSenderAuthenticationEnabled. AllowExternalSenders ($AllowExternalSenders) hat Vorrang." -Type Warning
            }
        } elseif ($PSBoundParameters.ContainsKey('RequireSenderAuthenticationEnabled')) {
            $params.RequireSenderAuthenticationEnabled = $RequireSenderAuthenticationEnabled
        }


        if ($params.Count -eq 0) {
            Write-Log "Update-ExoGroupSettingsAction: Keine Parameter zum Aktualisieren für '$Identity' angegeben." -Type Info
            Update-StatusBar -Message "Keine Änderungen für '$Identity' angegeben." -Type Info
            return $true 
        }
        
        $actionDescription = "Gruppe '$groupDisplayName' (Identity: $Identity) mit folgenden Einstellungen aktualisieren:"
        foreach($key in $params.Keys) {
            $actionDescription += "`n- $key = $($params[$key])"
        }

        if ($PSCmdlet.ShouldProcess($groupDisplayName, $actionDescription)) {
            switch ($recipientType) {
                "GroupMailbox" { 
                    Set-UnifiedGroup -Identity $Identity @params -ErrorAction Stop
                }
                "MailUniversalDistributionGroup" { 
                    Set-DistributionGroup -Identity $Identity @params -ErrorAction Stop
                }
                "MailUniversalSecurityGroup" { 
                    Set-DistributionGroup -Identity $Identity @params -ErrorAction Stop 
                }
                "DynamicDistributionGroup" {
                    Set-DynamicDistributionGroup -Identity $Identity @params -ErrorAction Stop
                }
                default {
                    $errMsg = "Update-ExoGroupSettingsAction: Nicht unterstützter Gruppentyp '$recipientType' für Gruppe '$Identity'."
                    Write-Log $errMsg -Type Error
                    Show-MessageBox -Message $errMsg -Title "Nicht unterstützter Typ" -Type Warning
                    Update-StatusBar -Message "Nicht unterstützter Gruppentyp: $recipientType" -Type Warning
                    return $false
                }
            }
            $successMsg = "Einstellungen für Gruppe '$groupDisplayName' erfolgreich aktualisiert."
            Write-Log $successMsg -Type Success
            Update-StatusBar -Message $successMsg -Type Success
            if (Test-Path Function:\Log-Action) { Log-Action "Einstellungen für Gruppe '$groupDisplayName' (Identity: $Identity) aktualisiert." }
            return $true
        } else {
            Write-Log "Update-ExoGroupSettingsAction: Aktualisierung für '$groupDisplayName' vom Benutzer abgebrochen." -Type Info
            Update-StatusBar -Message "Aktualisierung für '$groupDisplayName' abgebrochen." -Type Info
            return $false
        }

    } catch {
        $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Unbekannter Fehler beim Aktualisieren der Gruppeneinstellungen."
        $groupNameForError = if ($groupDisplayName) { $groupDisplayName } else { $Identity }
        Write-Log "Fehler in Update-ExoGroupSettingsAction für Gruppe '$groupNameForError': $errMsg" -Type Error
        if (Test-Path Function:\Log-Action) { Log-Action "Fehler Update-ExoGroupSettingsAction '$groupNameForError': $errMsg" }
        Show-MessageBox -Message "Fehler beim Aktualisieren der Einstellungen für Gruppe '$groupNameForError':`n$($_.Exception.Message)" -Title "Fehler bei Aktualisierung" -Type Error
        Update-StatusBar -Message "Fehler beim Aktualisieren der Einstellungen für '$groupNameForError'." -Type Error
        return $false
    }
}

function New-ExoGroupAction {
    [CmdletBinding(SupportsShouldProcess = $true)] # Hinzugefügt SupportsShouldProcess
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,

        [Parameter(Mandatory = $true)]
        [string]$GroupEmail, 

        [Parameter(Mandatory = $true)]
        [ValidateSet("Distribution", "Security", "Microsoft365")] 
        [string]$GroupType,

        [Parameter(Mandatory = $false)]
        [string[]]$Members,

        [Parameter(Mandatory = $false)]
        [string]$Description,

        [Parameter(Mandatory = $false)]
        [string]$Owner 
    )

    Write-Log "New-ExoGroupAction: Aktion gestartet für neue Gruppe '$GroupName' (Typ: $GroupType)." -Type Info
    Update-StatusBar -Message "Erstelle Gruppe '$GroupName'..." -Type Info

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return $null
        }
        
        $alias = if ($GroupEmail.Contains("@")) { $GroupEmail.Split('@')[0] } else { $GroupEmail }
        if ([string]::IsNullOrWhiteSpace($alias)) {
            Write-Log "New-ExoGroupAction: Ungültige E-Mail '$GroupEmail' zur Alias-Generierung." -Type Error
            Show-MessageBox -Message "Die angegebene E-Mail-Adresse '$GroupEmail' ist ungültig für die Alias-Generierung." -Title "Ungültige Eingabe" -Type Error
            Update-StatusBar -Message "Ungültige E-Mail für Alias." -Type Error
            return $null
        }


        $commonParams = @{
            DisplayName = $GroupName
            Alias = $alias
            ErrorAction = 'Stop'
        }
        
        $newGroup = $null
        $actionDescription = "Neue Gruppe '$GroupName' (E-Mail: '$GroupEmail', Typ: '$GroupType') erstellen"

        if ($PSCmdlet.ShouldProcess($GroupName, $actionDescription)) {
            switch ($GroupType) {
                "Distribution" {
                    $params = @{
                        Name = $GroupName # Oder Alias, je nach Konvention
                        Type = "Distribution" 
                        PrimarySmtpAddress = $GroupEmail
                    } + $commonParams
                    if (-not [string]::IsNullOrWhiteSpace($Description)) { $params.Notes = $Description }
                    if ($PSBoundParameters.ContainsKey('Owner') -and -not [string]::IsNullOrWhiteSpace($Owner)) {
                        $params.ManagedBy = $Owner
                    }
                    $newGroup = New-DistributionGroup @params
                }
                "Security" {
                     $params = @{
                        Name = $GroupName # Oder Alias
                        Type = "Security" 
                        PrimarySmtpAddress = $GroupEmail
                    } + $commonParams
                    if (-not [string]::IsNullOrWhiteSpace($Description)) { $params.Notes = $Description }
                    if ($PSBoundParameters.ContainsKey('Owner') -and -not [string]::IsNullOrWhiteSpace($Owner)) {
                        $params.ManagedBy = $Owner
                    }
                    $newGroup = New-DistributionGroup @params 
                }
                "Microsoft365" {
                    $params = @{
                        AccessType = "Private" 
                        PrimarySmtpAddress = $GroupEmail 
                    } + $commonParams
                    # Entferne Notes, da es für UnifiedGroup nicht gültig ist, Description wird separat behandelt
                    $params.Remove("Notes") | Out-Null 

                    if (-not [string]::IsNullOrWhiteSpace($Description)) { $params.Description = $Description }
                    
                    if ($PSBoundParameters.ContainsKey('Owner') -and -not [string]::IsNullOrWhiteSpace($Owner)) {
                        $params.Owners = $Owner 
                    } else {
                        try {
                            # Sicherstellen, dass $script:lblConnectedUser existiert und Content hat
                            if ($null -ne $script:lblConnectedUser -and -not [string]::IsNullOrWhiteSpace($script:lblConnectedUser.Content)) {
                                $currentUserPrincipalName = Get-AzureADUser -ObjectId (Get-MsolUser -UserPrincipalName ($script:lblConnectedUser.Content) -ErrorAction Stop).ObjectId -ErrorAction Stop | Select-Object -ExpandProperty UserPrincipalName
                                if(-not [string]::IsNullOrWhiteSpace($currentUserPrincipalName)){
                                    $params.Owners = $currentUserPrincipalName
                                    Write-Log "New-ExoGroupAction: Aktueller Benutzer '$currentUserPrincipalName' wird als Owner für M365-Gruppe '$GroupName' gesetzt." -Type Info
                                } else {
                                    Write-Log "New-ExoGroupAction: Owner für M365-Gruppe '$GroupName' konnte nicht aus aktuellem Benutzer '$($script:lblConnectedUser.Content)' ermittelt werden." -Type Warning
                                }
                            } else {
                                Write-Log "New-ExoGroupAction: Kein Owner für M365-Gruppe '$GroupName' angegeben und aktueller Benutzer nicht in UI verfügbar." -Type Warning
                            }
                        } catch {
                             Write-Log "New-ExoGroupAction: Fehler beim Ermitteln des aktuellen Benutzers als Owner für M365-Gruppe '$GroupName': $($_.Exception.Message)" -Type Warning
                        }
                    }
                    $newGroup = New-UnifiedGroup @params
                }
                default { # Sollte durch ValidateSet nicht erreicht werden
                    $errMsg = "New-ExoGroupAction: Nicht unterstützter Gruppentyp '$GroupType'."
                    Write-Log $errMsg -Type Error
                    Show-MessageBox -Message $errMsg -Title "Nicht unterstützter Typ" -Type Warning
                    Update-StatusBar -Message "Nicht unterstützter Gruppentyp: $GroupType" -Type Warning
                    return $null
                }
            }

            if ($newGroup) {
                $successMsg = "Gruppe '$($newGroup.DisplayName)' (Typ: $GroupType) erfolgreich erstellt."
                Write-Log $successMsg -Type Success
                Update-StatusBar -Message $successMsg -Type Success
                if (Test-Path Function:\Log-Action) { Log-Action "Gruppe '$($newGroup.DisplayName)' (Identity: $($newGroup.Identity)) erstellt." }

                if ($Members -and $Members.Count -gt 0) {
                    Write-Log "Füge Mitglieder zur neuen Gruppe '$($newGroup.DisplayName)' hinzu..." -Type Info
                    foreach ($member in $Members) {
                        if (-not [string]::IsNullOrWhiteSpace($member)) {
                            # Korrekte Parameter für Add-ExoGroupMemberAction
                            Add-ExoGroupMemberAction -Identity $newGroup.Identity -MemberIdentity $member -ErrorAction SilentlyContinue 
                        }
                    }
                }
                return $newGroup 
            } else {
                $errMsg = "Gruppe '$GroupName' konnte nicht erstellt werden. Überprüfen Sie die Exchange Online-Protokolle oder vorherige Fehlermeldungen."
                Write-Log $errMsg -Type Error
                Update-StatusBar -Message $errMsg -Type Error
                return $null
            }
        } else {
            Write-Log "New-ExoGroupAction: Erstellung der Gruppe '$GroupName' vom Benutzer abgebrochen." -Type Info
            Update-StatusBar -Message "Erstellung der Gruppe '$GroupName' abgebrochen." -Type Info
            return $null
        }

    } catch {
        $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Unbekannter Fehler beim Erstellen der Gruppe."
        Write-Log "Fehler in New-ExoGroupAction für Gruppe '$GroupName': $errMsg" -Type Error
        if (Test-Path Function:\Log-Action) { Log-Action "Fehler New-ExoGroupAction '$GroupName': $errMsg" }
        Show-MessageBox -Message "Fehler beim Erstellen der Gruppe '$GroupName':`n$($_.Exception.Message)" -Title "Fehler bei Erstellung" -Type Error
        Update-StatusBar -Message "Fehler beim Erstellen der Gruppe '$GroupName'." -Type Error
        return $null
    }
}

function Remove-ExoGroupAction {
    [CmdletBinding(SupportsShouldProcess = $true)] # Hinzugefügt SupportsShouldProcess
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    
    $groupDisplayName = $Identity # Fallback, falls Objekt nicht geladen werden kann

    Write-Log "Remove-ExoGroupAction: Aktion gestartet für Gruppe '$Identity'." -Type Info
    Update-StatusBar -Message "Lösche Gruppe '$Identity'..." -Type Info

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return $false
        }

        $groupObject = Get-ExoGroupRecipientAction -Identity $Identity
        if (-not $groupObject) {
            # Gruppe existiert möglicherweise nicht mehr, was beim Löschen kein Fehler sein muss.
            Write-Log "Remove-ExoGroupAction: Gruppe '$Identity' nicht gefunden. Möglicherweise bereits gelöscht." -Type Warning
            Update-StatusBar -Message "Gruppe '$Identity' nicht gefunden (evtl. bereits gelöscht)." -Type Warning
            return $true # Betrachte "nicht gefunden" als Erfolg für einen Löschvorgang
        }
        $recipientType = $groupObject.RecipientTypeDetails
        $groupDisplayName = $groupObject.DisplayName # Aktualisiere mit dem korrekten DisplayName

        # Zeichenfolgen für ShouldProcess vorbereiten, um Parser-Probleme zu umgehen
        $shouldProcessTarget = "Gruppe '{0}' (Identity: {1}, Typ: {2})" -f $groupDisplayName, $groupObject.Identity, $recipientType
        $shouldProcessAction = "Endgültig löschen"

        if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction)) {
            switch ($recipientType) {
                "GroupMailbox" { 
                    Remove-UnifiedGroup -Identity $groupObject.Identity -Confirm:$false -ErrorAction Stop
                }
                "MailUniversalDistributionGroup" { 
                    Remove-DistributionGroup -Identity $groupObject.Identity -Confirm:$false -ErrorAction Stop
                }
                "MailUniversalSecurityGroup" { 
                    Remove-DistributionGroup -Identity $groupObject.Identity -Confirm:$false -ErrorAction Stop
                }
                "DynamicDistributionGroup" {
                    Remove-DynamicDistributionGroup -Identity $groupObject.Identity -Confirm:$false -ErrorAction Stop
                }
                default {
                    $errMsg = "Remove-ExoGroupAction: Nicht unterstützter Gruppentyp '$recipientType' für Gruppe '$groupDisplayName' zum Löschen."
                    Write-Log $errMsg -Type Error
                    Show-MessageBox -Message $errMsg -Title "Nicht unterstützter Typ" -Type Warning
                    Update-StatusBar -Message "Nicht unterstützter Gruppentyp: $recipientType" -Type Warning
                    return $false
                }
            }
            $successMsg = "Gruppe '$groupDisplayName' erfolgreich gelöscht."
            Write-Log $successMsg -Type Success
            Update-StatusBar -Message $successMsg -Type Success
            if (Test-Path Function:\Log-Action) { Log-Action "Gruppe '$groupDisplayName' (Identity: $($groupObject.Identity)) gelöscht." }
            return $true
        } else {
            Write-Log "Remove-ExoGroupAction: Löschen der Gruppe '$groupDisplayName' vom Benutzer abgebrochen." -Type Info
            Update-StatusBar -Message "Löschen der Gruppe '$groupDisplayName' abgebrochen." -Type Info
            return $false
        }

    } catch {
        $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Unbekannter Fehler beim Löschen der Gruppe."
        Write-Log "Fehler in Remove-ExoGroupAction für Gruppe '$groupDisplayName': $errMsg" -Type Error
        if (Test-Path Function:\Log-Action) { Log-Action "Fehler Remove-ExoGroupAction '$groupDisplayName': $errMsg" }
        Show-MessageBox -Message "Fehler beim Löschen der Gruppe '$groupDisplayName':`n$($_.Exception.Message)" -Title "Fehler bei Löschung" -Type Error
        Update-StatusBar -Message "Fehler beim Löschen der Gruppe '$groupDisplayName'." -Type Error
        return $false
    }
}

function Add-ExoGroupMemberAction {
    [CmdletBinding(SupportsShouldProcess = $true)] # Hinzugefügt SupportsShouldProcess
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity, # Gruppen-Identität

        [Parameter(Mandatory = $true)]
        [string]$MemberIdentity # Benutzer/Mitglied-Identität
    )

    Write-Log "Add-ExoGroupMemberAction: Füge Mitglied '$MemberIdentity' zu Gruppe '$Identity' hinzu." -Type Info
    Update-StatusBar -Message "Füge '$MemberIdentity' zu '$Identity' hinzu..." -Type Info
    
    $groupDisplayName = $Identity # Fallback

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return $false
        }

        $group = Get-ExoGroupRecipientAction -Identity $Identity
        if (-not $group) {
            Update-StatusBar -Message "Gruppe '$Identity' nicht gefunden." -Type Error
            return $false
        }
        # Optional: $member = Get-Recipient -Identity $MemberIdentity -ErrorAction SilentlyContinue ...
        
        $recipientType = $group.RecipientTypeDetails
        $groupDisplayName = $group.DisplayName

        # Zeichenfolgen für ShouldProcess vorbereiten
        $shouldProcessTarget = "Mitglied '$MemberIdentity' zu Gruppe '$groupDisplayName' (Typ: $recipientType)"
        $shouldProcessAction = "Hinzufügen"

        if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction)) {
            switch ($recipientType) {
                "GroupMailbox" { 
                    Add-UnifiedGroupLinks -Identity $Identity -LinkType Members -Links $MemberIdentity -Confirm:$false -ErrorAction Stop
                }
                "MailUniversalDistributionGroup" { 
                    Add-DistributionGroupMember -Identity $Identity -Member $MemberIdentity -Confirm:$false -ErrorAction Stop
                }
                "MailUniversalSecurityGroup" { 
                    Add-DistributionGroupMember -Identity $Identity -Member $MemberIdentity -Confirm:$false -ErrorAction Stop
                }
                "DynamicDistributionGroup" {
                     $errMsg = "Add-ExoGroupMemberAction: Mitglieder können nicht manuell zu dynamischen Verteilergruppen hinzugefügt werden (Gruppe '$groupDisplayName')."
                    Write-Log $errMsg -Type Warning
                    Show-MessageBox -Message $errMsg -Title "Aktion nicht unterstützt" -Type Warning
                    Update-StatusBar -Message "Mitglieder können nicht zu dynamischen Gruppen hinzugefügt werden." -Type Warning
                    return $false
                }
                default {
                    $errMsg = "Add-ExoGroupMemberAction: Mitglieder können nicht manuell zu Gruppentyp '$recipientType' (Gruppe '$groupDisplayName') hinzugefügt werden."
                    Write-Log $errMsg -Type Error
                    Show-MessageBox -Message $errMsg -Title "Nicht unterstützter Typ" -Type Warning
                    Update-StatusBar -Message "Mitglieder können nicht zu '$recipientType' hinzugefügt werden." -Type Warning
                    return $false
                }
            }
            $successMsg = "Mitglied '$MemberIdentity' erfolgreich zu Gruppe '$groupDisplayName' hinzugefügt."
            Write-Log $successMsg -Type Success
            Update-StatusBar -Message $successMsg -Type Success
            if (Test-Path Function:\Log-Action) { Log-Action "Mitglied '$MemberIdentity' zu Gruppe '$groupDisplayName' (Identity: $Identity) hinzugefügt." }
            return $true
        } else {
            Write-Log "Add-ExoGroupMemberAction: Hinzufügen von '$MemberIdentity' zu '$groupDisplayName' vom Benutzer abgebrochen." -Type Info
            Update-StatusBar -Message "Hinzufügen zu '$groupDisplayName' abgebrochen." -Type Info
            return $false
        }

    } catch {
        $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Unbekannter Fehler beim Hinzufügen des Gruppenmitglieds."
        Write-Log "Fehler in Add-ExoGroupMemberAction für Gruppe '$groupDisplayName', Mitglied '$MemberIdentity': $errMsg" -Type Error
        if (Test-Path Function:\Log-Action) { Log-Action "Fehler Add-ExoGroupMemberAction '$groupDisplayName'/'$MemberIdentity': $errMsg" }
        Show-MessageBox -Message "Fehler beim Hinzufügen von '$MemberIdentity' zu Gruppe '$groupDisplayName':`n$($_.Exception.Message)" -Title "Fehler bei Mitgliedschaft" -Type Error
        Update-StatusBar -Message "Fehler beim Hinzufügen zu '$groupDisplayName'." -Type Error
        return $false
    }
}

function Remove-ExoGroupMemberAction {
    [CmdletBinding(SupportsShouldProcess = $true)] # Hinzugefügt SupportsShouldProcess
    param(
        # Beibehaltung der ursprünglichen Parameternamen dieser Funktion
        [Parameter(Mandatory = $true)]
        [string]$GroupIdentity, 

        [Parameter(Mandatory = $true)]
        [string]$UserIdentity # Mitglied-Identität
    )

    Write-Log "Remove-ExoGroupMemberAction: Entferne Mitglied '$UserIdentity' aus Gruppe '$GroupIdentity'." -Type Info
    Update-StatusBar -Message "Entferne '$UserIdentity' aus '$GroupIdentity'..." -Type Info
    
    $groupDisplayName = $GroupIdentity # Fallback

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return $false
        }

        $group = Get-ExoGroupRecipientAction -Identity $GroupIdentity
        if (-not $group) {
            Update-StatusBar -Message "Gruppe '$GroupIdentity' nicht gefunden." -Type Error
            return $false
        }
        
        $recipientType = $group.RecipientTypeDetails
        $groupDisplayName = $group.DisplayName

        # Korrektur: $PSCmdlet.ShouldProcess verwenden und Ziel/Aktion trennen
        $shouldProcessTarget = "Mitglied '$UserIdentity' aus Gruppe '$groupDisplayName' (Typ: $recipientType)"
        $shouldProcessAction = "Entfernen"

        if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction)) {
            switch ($recipientType) {
                "GroupMailbox" { 
                    Remove-UnifiedGroupLinks -Identity $GroupIdentity -LinkType Members -Links $UserIdentity -Confirm:$false -ErrorAction Stop
                }
                "MailUniversalDistributionGroup" { 
                    Remove-DistributionGroupMember -Identity $GroupIdentity -Member $UserIdentity -Confirm:$false -ErrorAction Stop
                }
                "MailUniversalSecurityGroup" { 
                    Remove-DistributionGroupMember -Identity $GroupIdentity -Member $UserIdentity -Confirm:$false -ErrorAction Stop
                }
                 "DynamicDistributionGroup" {
                     $errMsg = "Remove-ExoGroupMemberAction: Mitglieder können nicht manuell aus dynamischen Verteilergruppen entfernt werden (Gruppe '$groupDisplayName')."
                    Write-Log $errMsg -Type Warning
                    Show-MessageBox -Message $errMsg -Title "Aktion nicht unterstützt" -Type Warning
                    Update-StatusBar -Message "Mitglieder können nicht aus dynamischen Gruppen entfernt werden." -Type Warning
                    return $false
                }
                default {
                    $errMsg = "Remove-ExoGroupMemberAction: Mitglieder können nicht manuell aus Gruppentyp '$recipientType' (Gruppe '$groupDisplayName') entfernt werden."
                    Write-Log $errMsg -Type Error
                    Show-MessageBox -Message $errMsg -Title "Nicht unterstützter Typ" -Type Warning
                    Update-StatusBar -Message "Mitglieder können nicht aus '$recipientType' entfernt werden." -Type Warning
                    return $false
                }
            }
            $successMsg = "Mitglied '$UserIdentity' erfolgreich aus Gruppe '$groupDisplayName' entfernt."
            Write-Log $successMsg -Type Success
            Update-StatusBar -Message $successMsg -Type Success
            if (Test-Path Function:\Log-Action) { Log-Action "Mitglied '$UserIdentity' aus Gruppe '$groupDisplayName' (Identity: $GroupIdentity) entfernt." }
            return $true
        } else {
            Write-Log "Remove-ExoGroupMemberAction: Entfernen von '$UserIdentity' aus '$groupDisplayName' vom Benutzer abgebrochen." -Type Info
            Update-StatusBar -Message "Entfernen aus '$groupDisplayName' abgebrochen." -Type Info
            return $false
        }

    } catch {
        $errMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Unbekannter Fehler beim Entfernen des Gruppenmitglieds."
        Write-Log "Fehler in Remove-ExoGroupMemberAction für Gruppe '$groupDisplayName', Mitglied '$UserIdentity': $errMsg" -Type Error
        if (Test-Path Function:\Log-Action) { Log-Action "Fehler Remove-ExoGroupMemberAction '$groupDisplayName'/'$UserIdentity': $errMsg" }
        Show-MessageBox -Message "Fehler beim Entfernen von '$UserIdentity' aus Gruppe '$groupDisplayName':`n$($_.Exception.Message)" -Title "Fehler bei Mitgliedschaft" -Type Error
        Update-StatusBar -Message "Fehler beim Entfernen aus '$groupDisplayName'." -Type Error
        return $false
    }
}

function Get-ExoGroupMembersAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity # Gruppenidentität
    )

    Write-Log "Get-ExoGroupMembersAction: Aktion gestartet für Gruppe '$Identity'." -Type Info
    Update-StatusBar -Message "Lade Gruppenmitglieder und -einstellungen für '$Identity'..." -Type Info

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return
        }

        if ($null -eq $script:lstGroupMembers -or $null -eq $script:chkHiddenFromGAL -or $null -eq $script:chkRequireSenderAuth -or $null -eq $script:chkAllowExternalSenders) {
            Write-Log "Get-ExoGroupMembersAction: Kritische UI-Elemente nicht initialisiert." -Type Error
            Show-MessageBox -Message "Einige UI-Elemente für die Gruppenverwaltung sind nicht verfügbar." -Title "UI Fehler" -Type Error
            return
        }
        
        # Gruppe basierend auf der übergebenen Identity abrufen
        $groupObject = Get-ExoGroupRecipientAction -Identity $Identity
        if ($null -eq $groupObject) {
            Write-Log "Get-ExoGroupMembersAction: Gruppe mit Identity '$Identity' nicht gefunden." -Type Error
            Show-MessageBox -Message "Gruppe mit Identity '$Identity' nicht gefunden." -Title "Fehler" -Type Error
            Update-StatusBar -Message "Gruppe '$Identity' nicht gefunden." -Type Error
            # UI zurücksetzen
            $script:lstGroupMembers.Items.Clear(); $script:lstGroupMembers.Tag = $null
            $script:chkHiddenFromGAL.IsChecked = $false; $script:chkHiddenFromGAL.IsEnabled = $false
            $script:chkRequireSenderAuth.IsChecked = $false; $script:chkRequireSenderAuth.IsEnabled = $false
            $script:chkAllowExternalSenders.IsChecked = $false; $script:chkAllowExternalSenders.IsEnabled = $false
            return
        }
        
        $groupDisplayName = $groupObject.DisplayName
        $effectiveIdentityForExo = $groupObject.Identity # Die von Get-ExoGroupRecipientAction aufgelöste Identity verwenden
        $recipientType = $groupObject.RecipientTypeDetails

        Write-Log "Get-ExoGroupMembersAction: Verarbeite Gruppe '$groupDisplayName' (Typ: '$recipientType', ID für EXO: '$effectiveIdentityForExo')" -Type Info
        Update-StatusBar -Message "Lade Mitglieder für '$groupDisplayName'..." -Type Info

        $script:lstGroupMembers.Items.Clear()
        $script:lstGroupMembers.Tag = $null 

        $members = @()
        $membersSuccessfullyRetrieved = $false
        try {
            Write-Log "Get-ExoGroupMembersAction: Rufe Mitglieder für '$groupDisplayName' ab..." -Type Debug
            if ($recipientType -eq "GroupMailbox") { # Korrigiert von "UnifiedGroup"
                $memberLinks = Get-UnifiedGroupLinks -Identity $effectiveIdentityForExo -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                if ($null -ne $memberLinks) { $members = $memberLinks }
            } elseif ($recipientType -in ("MailUniversalDistributionGroup", "MailUniversalSecurityGroup", "MailNonUniversalGroup", "DynamicDistributionGroup")) {
                $members = Get-DistributionGroupMember -Identity $effectiveIdentityForExo -ResultSize Unlimited -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            } else {
                Write-Log "Get-ExoGroupMembersAction: Unbekannter oder nicht unterstützter Gruppentyp '$recipientType' für Mitgliederabruf." -Type Warning
                Show-MessageBox -Message "Der Gruppentyp '$recipientType' wird für den Mitgliederabruf aktuell nicht vollständig unterstützt." -Title "Hinweis" -Type Info
            }
            
            if ($null -ne $members) { # $members kann leer sein @(), aber nicht $null, wenn Cmdlet erfolgreich war
                $membersSuccessfullyRetrieved = $true
                Write-Log "Get-ExoGroupMembersAction: $($members.Count) Mitglieder-Objekte für '$groupDisplayName' abgerufen." -Type Debug
            } else { # Fall, dass Cmdlet $null zurückgibt (sollte bei ErrorAction SilentlyContinue nicht passieren, eher leeres Array)
                 Write-Log "Get-ExoGroupMembersAction: Keine Mitglieder-Objekte für '$groupDisplayName' zurückgegeben (Cmdlet lieferte null oder Fehler)." -Type Debug
                 $members = @() 
            }
        } catch {
            $errorMsg = $_.Exception.Message
            Write-Log "Get-ExoGroupMembersAction: Fehler beim Abrufen der Mitglieder für '$groupDisplayName': $errorMsg" -Type Error
            Show-MessageBox -Message "Fehler beim Abrufen der Gruppenmitglieder für '$groupDisplayName': $errorMsg" -Title "Fehler Mitgliederabruf" -Type Error
        }

        if ($membersSuccessfullyRetrieved -and $members.Count -gt 0) {
            foreach ($member in $members) {
                $displayMemberName = $member.DisplayName
                if ([string]::IsNullOrWhiteSpace($displayMemberName)) { $displayMemberName = $member.Name }
                if ([string]::IsNullOrWhiteSpace($displayMemberName) -and $member.PSObject.Properties["PrimarySmtpAddress"]) { $displayMemberName = $member.PrimarySmtpAddress.ToString() }
                if ([string]::IsNullOrWhiteSpace($displayMemberName) -and $member.PSObject.Properties["Alias"]) { $displayMemberName = $member.Alias.ToString() }
                if ([string]::IsNullOrWhiteSpace($displayMemberName)) { $displayMemberName = if ($member.Identity) {$member.Identity.ToString()} else {"Unbekanntes Mitglied"} }
                
                $item = New-Object System.Windows.Controls.ListViewItem
                $item.Content = $displayMemberName
                $item.Tag = $member 
                $script:lstGroupMembers.Items.Add($item) | Out-Null
            }
            $script:lstGroupMembers.Tag = $members 
        } else {
            $item = New-Object System.Windows.Controls.ListViewItem
            $item.Content = if (-not $membersSuccessfullyRetrieved) { "Mitglieder konnten nicht geladen werden." } else { "Keine Mitglieder gefunden oder Gruppe ist leer." }
            $item.IsEnabled = $false 
            $script:lstGroupMembers.Items.Add($item) | Out-Null
        }
        Write-Log "Get-ExoGroupMembersAction: Mitgliederliste für '$groupDisplayName' aktualisiert. $($members.Count) Mitglieder angezeigt." -Type Info
        Update-StatusBar -Message "$($members.Count) Mitglied(er) für '$groupDisplayName' geladen." -Type Info


        Write-Log "Get-ExoGroupMembersAction: Lade Einstellungen für '$groupDisplayName'..." -Type Debug
        Update-StatusBar -Message "Lade Einstellungen für '$groupDisplayName'..." -Type Info
        try {
            $script:chkHiddenFromGAL.IsEnabled = $true; $script:chkHiddenFromGAL.IsChecked = $false
            $script:chkRequireSenderAuth.IsEnabled = $true; $script:chkRequireSenderAuth.IsChecked = $false
            $script:chkAllowExternalSenders.IsEnabled = $true; $script:chkAllowExternalSenders.IsChecked = $false

            $groupDetailsForSettings = $null
            if ($recipientType -eq "GroupMailbox") { # Korrigiert von "UnifiedGroup"
                $groupDetailsForSettings = Get-UnifiedGroup -Identity $effectiveIdentityForExo -ErrorAction Stop
                if ($groupDetailsForSettings) {
                    $script:chkHiddenFromGAL.IsChecked = $groupDetailsForSettings.HiddenFromExchangeAddressListsEnabled
                    $script:chkRequireSenderAuth.IsChecked = $groupDetailsForSettings.RequireSenderAuthenticationEnabled
                    $script:chkAllowExternalSenders.IsChecked = (-not $groupDetailsForSettings.RequireSenderAuthenticationEnabled)
                }
            } elseif ($recipientType -in ("MailUniversalDistributionGroup", "MailUniversalSecurityGroup", "MailNonUniversalGroup")) {
                $groupDetailsForSettings = Get-DistributionGroup -Identity $effectiveIdentityForExo -ErrorAction Stop
                 if ($groupDetailsForSettings) {
                    $script:chkHiddenFromGAL.IsChecked = $groupDetailsForSettings.HiddenFromAddressListsEnabled
                    $script:chkRequireSenderAuth.IsChecked = $groupDetailsForSettings.RequireSenderAuthenticationEnabled
                    $script:chkAllowExternalSenders.IsChecked = (-not $groupDetailsForSettings.RequireSenderAuthenticationEnabled)
                }
            } elseif ($recipientType -eq "DynamicDistributionGroup") {
                $groupDetailsForSettings = Get-DynamicDistributionGroup -Identity $effectiveIdentityForExo -ErrorAction Stop
                 if ($groupDetailsForSettings) {
                    $script:chkHiddenFromGAL.IsChecked = $groupDetailsForSettings.HiddenFromAddressListsEnabled
                    # RequireSenderAuthenticationEnabled ist für DynamicDistributionGroup nicht direkt so relevant wie für andere
                    # Setze es basierend auf dem Objekt, falls vorhanden, ansonsten deaktiviere die Checkbox
                    if ($groupDetailsForSettings.PSObject.Properties["RequireSenderAuthenticationEnabled"]) {
                        $script:chkRequireSenderAuth.IsChecked = $groupDetailsForSettings.RequireSenderAuthenticationEnabled
                        $script:chkAllowExternalSenders.IsChecked = (-not $groupDetailsForSettings.RequireSenderAuthenticationEnabled)
                    } else {
                        $script:chkRequireSenderAuth.IsEnabled = $false
                        $script:chkAllowExternalSenders.IsEnabled = $false
                    }
                }
            } else {
                Write-Log "Get-ExoGroupMembersAction: Gruppentyp '$recipientType' hat keine unterstützten Einstellungen für die UI." -Type Info
                $script:chkHiddenFromGAL.IsEnabled = $false
                $script:chkRequireSenderAuth.IsEnabled = $false
                $script:chkAllowExternalSenders.IsEnabled = $false
            }
            
            if ($null -ne $groupDetailsForSettings) {
                 Write-Log "Get-ExoGroupMembersAction: Einstellungen für '$groupDisplayName' geladen und UI aktualisiert." -Type Info
            } else {
                 Write-Log "Get-ExoGroupMembersAction: Keine Einstellungsdetails für '$groupDisplayName' (Typ '$recipientType') abrufbar oder Typ nicht unterstützt für Einstellungen." -Type Warning
                 $script:chkHiddenFromGAL.IsEnabled = $false
                 $script:chkRequireSenderAuth.IsEnabled = $false
                 $script:chkAllowExternalSenders.IsEnabled = $false
            }

        } catch {
            $errorMsg = $_.Exception.Message
            Write-Log "Get-ExoGroupMembersAction: Fehler beim Laden der Einstellungen für '$groupDisplayName': $errorMsg" -Type Error
            Show-MessageBox -Message "Fehler beim Laden der Gruppeneinstellungen für '$groupDisplayName': $errorMsg" -Title "Fehler Einstellungen" -Type Error
            $script:chkHiddenFromGAL.IsChecked = $false; $script:chkHiddenFromGAL.IsEnabled = $false
            $script:chkRequireSenderAuth.IsChecked = $false; $script:chkRequireSenderAuth.IsEnabled = $false
            $script:chkAllowExternalSenders.IsChecked = $false; $script:chkAllowExternalSenders.IsEnabled = $false
        }

        Update-StatusBar -Message "Mitglieder und Einstellungen für '$groupDisplayName' geladen." -Type Success
        Write-Log "Get-ExoGroupMembersAction: Aktion erfolgreich abgeschlossen für '$groupDisplayName'." -Type Success

    } catch {
        $errorMsg = $_.Exception.Message
        $fullError = $_.ToString()
        Write-Log "Get-ExoGroupMembersAction: Schwerwiegender Fehler für Gruppe '$Identity': $errorMsg `n$fullError" -Type Error
        Show-MessageBox -Message "Ein schwerwiegender Fehler ist aufgetreten beim Laden der Gruppeninformationen für '$Identity': $errorMsg" -Title "Schwerer Fehler" -Type Error
        Update-StatusBar -Message "Fehler beim Laden der Gruppeninformationen für '$Identity'." -Type Error

        try {
            $script:chkHiddenFromGAL.IsChecked = $false; $script:chkHiddenFromGAL.IsEnabled = $false
            $script:chkRequireSenderAuth.IsChecked = $false; $script:chkRequireSenderAuth.IsEnabled = $false
            $script:chkAllowExternalSenders.IsChecked = $false; $script:chkAllowExternalSenders.IsEnabled = $false
            $script:lstGroupMembers.Items.Clear(); $script:lstGroupMembers.Tag = $null
            $item = New-Object System.Windows.Controls.ListViewItem
            $item.Content = "Fehler beim Laden der Gruppeninformationen."
            $item.IsEnabled = $false
            $script:lstGroupMembers.Items.Add($item) | Out-Null
        } catch {
            Write-Log "Get-ExoGroupMembersAction: Konnte UI nach schwerem Fehler nicht zurücksetzen." -Type Error
        }
    }
}

function Initialize-GroupsTab {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Initialisiere Gruppen-Tab..." -Type "Info"

        $script:cmbGroupType = Get-XamlElement -ElementName "cmbGroupType"
        $script:txtGroupName = Get-XamlElement -ElementName "txtGroupName"
        $script:txtGroupEmail = Get-XamlElement -ElementName "txtGroupEmail"
        $script:txtGroupMembers = Get-XamlElement -ElementName "txtGroupMembers"
        $script:txtGroupDescription = Get-XamlElement -ElementName "txtGroupDescription"
        $script:btnCreateGroup = Get-XamlElement -ElementName "btnCreateGroup"
        $script:btnDeleteGroup = Get-XamlElement -ElementName "btnDeleteGroup"

        $script:cmbSelectExistingGroup = Get-XamlElement -ElementName "cmbSelectExistingGroup" 
        $script:btnRefreshExistingGroups = Get-XamlElement -ElementName "btnRefreshExistingGroups" 
        $script:txtGroupUser = Get-XamlElement -ElementName "txtGroupUser"
        $script:btnAddUserToGroup = Get-XamlElement -ElementName "btnAddUserToGroup"
        $script:btnRemoveUserFromGroup = Get-XamlElement -ElementName "btnRemoveUserFromGroup"
        $script:chkHiddenFromGAL = Get-XamlElement -ElementName "chkHiddenFromGAL"
        $script:chkRequireSenderAuth = Get-XamlElement -ElementName "chkRequireSenderAuth"
        $script:chkAllowExternalSenders = Get-XamlElement -ElementName "chkAllowExternalSenders"
        $script:btnShowGroupMembers = Get-XamlElement -ElementName "btnShowGroupMembers"
        $script:btnUpdateGroupSettings = Get-XamlElement -ElementName "btnUpdateGroupSettings"

        $script:lstGroupMembers = Get-XamlElement -ElementName "lstGroupMembers"
        $script:btnExportGroupMembers = Get-XamlElement -ElementName "btnExportGroupMembers" 

        $script:helpLinkGroups = Get-XamlElement -ElementName "helpLinkGroups"

        if ($null -ne $script:cmbGroupType) {
            # UI-Anzeigenamen. Die Umsetzung zu API-Werten erfolgt im Handler.
            $groupTypesForUI = @("Verteilergruppe", "E-Mail-aktivierte Sicherheitsgruppe", "Microsoft 365-Gruppe") 
            $script:cmbGroupType.Items.Clear()
            foreach ($type in $groupTypesForUI) {
                $item = New-Object System.Windows.Controls.ComboBoxItem
                $item.Content = $type
                [void]$script:cmbGroupType.Items.Add($item)
            }
            if ($script:cmbGroupType.Items.Count -gt 0) {
                $script:cmbGroupType.SelectedIndex = 0 
            }
        }

        if ($null -ne $script:btnRefreshExistingGroups) {
            Register-EventHandler -Control $script:btnRefreshExistingGroups -Handler {
                try {
                    Write-Log "Button 'Gruppenliste aktualisieren' (btnRefreshExistingGroups) geklickt." -Type "Info"
                    Refresh-ExistingGroupsDropdown 
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $logErrorMsg = "Fehler beim Klick auf 'Gruppenliste aktualisieren': $errorMsg"
                    Write-Log $logErrorMsg -Type "Error"
                    Update-StatusBar -Message $logErrorMsg -Type "Error"
                }
            } -ControlName "btnRefreshExistingGroups"
        }

        if ($null -ne $script:cmbSelectExistingGroup) {
            $script:cmbSelectExistingGroup.Add_SelectionChanged({
                param($sender, $e)
                try {
                    if ($null -ne $script:cmbSelectExistingGroup.SelectedItem) {
                        $selectedGroupItem = $script:cmbSelectExistingGroup.SelectedItem
                        $groupObject = $selectedGroupItem.Tag

                        if ($null -eq $groupObject -or $null -eq $groupObject.Identity) {
                            Write-Log "SelectionChanged: Ungültiges Gruppenobjekt oder fehlende Identity im Tag des ausgewählten Elements." -Type Warning
                            if ($null -ne $script:lstGroupMembers) { $script:lstGroupMembers.Items.Clear(); $script:lstGroupMembers.Tag = $null }
                            if ($null -ne $script:chkHiddenFromGAL) { $script:chkHiddenFromGAL.IsChecked = $false; $script:chkHiddenFromGAL.IsEnabled = $false }
                            if ($null -ne $script:chkRequireSenderAuth) { $script:chkRequireSenderAuth.IsChecked = $false; $script:chkRequireSenderAuth.IsEnabled = $false }
                            if ($null -ne $script:chkAllowExternalSenders) { $script:chkAllowExternalSenders.IsChecked = $false; $script:chkAllowExternalSenders.IsEnabled = $false }
                            if ($null -ne $script:txtGroupUser) { $script:txtGroupUser.Text = "" }
                            Update-StatusBar -Message "Ungültige Gruppenauswahl." -Type Warning
                            return
                        }
                    
                        Write-Log "Gruppe '$($selectedGroupItem.Content)' (Identity: $($groupObject.Identity)) in ComboBox ausgewählt. Lade Details..." -Type "Info"
                        Get-ExoGroupMembersAction -Identity $groupObject.Identity 

                        if ($null -ne $script:txtGroupUser) { $script:txtGroupUser.Text = "" } 
                    } else {
                        Write-Log "ComboBox-Auswahl gelöscht oder kein Element ausgewählt (cmbSelectExistingGroup)." -Type Info
                        if ($null -ne $script:lstGroupMembers) { $script:lstGroupMembers.Items.Clear(); $script:lstGroupMembers.Tag = $null }
                        if ($null -ne $script:chkHiddenFromGAL) { $script:chkHiddenFromGAL.IsChecked = $false; $script:chkHiddenFromGAL.IsEnabled = $false }
                        if ($null -ne $script:chkRequireSenderAuth) { $script:chkRequireSenderAuth.IsChecked = $false; $script:chkRequireSenderAuth.IsEnabled = $false }
                        if ($null -ne $script:chkAllowExternalSenders) { $script:chkAllowExternalSenders.IsChecked = $false; $script:chkAllowExternalSenders.IsEnabled = $false }
                        if ($null -ne $script:txtGroupUser) { $script:txtGroupUser.Text = "" }
                        Update-StatusBar -Message "Keine Gruppe ausgewählt." -Type Info
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-Log "Fehler bei SelectionChanged für cmbSelectExistingGroup: $errorMsg" -Type "Error"
                    Update-StatusBar -Message "Fehler bei Gruppenauswahl: $errorMsg" -Type "Error"
                }
            })
        }

        Refresh-ExistingGroupsDropdown

        if ($null -ne $script:btnCreateGroup) {
            Register-EventHandler -Control $script:btnCreateGroup -Handler {
                try {
                    Write-Log "Button 'Gruppe erstellen' geklickt." -Type "Info"
                    if (-not $script:isConnected) {
                        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                        return
                    }

                    if ($null -eq $script:cmbGroupType.SelectedItem) {
                        Show-MessageBox -Message "Bitte wählen Sie einen Gruppentyp aus." -Title "Unvollständige Angaben" -Type Warning
                        return
                    }
                    if ([string]::IsNullOrWhiteSpace($script:txtGroupName.Text)) {
                        Show-MessageBox -Message "Bitte geben Sie einen Gruppennamen an." -Title "Unvollständige Angaben" -Type Warning
                        return
                    }
                     if ([string]::IsNullOrWhiteSpace($script:txtGroupEmail.Text)) {
                        Show-MessageBox -Message "Bitte geben Sie eine E-Mail-Adresse für die Gruppe an." -Title "Unvollständige Angaben" -Type Warning
                        return
                    }

                    $uiGroupType = $script:cmbGroupType.SelectedItem.Content
                    $apiGroupType = ""
                    switch ($uiGroupType) {
                        "Verteilergruppe" { $apiGroupType = "Distribution" }
                        "E-Mail-aktivierte Sicherheitsgruppe" { $apiGroupType = "Security" }
                        "Microsoft 365-Gruppe" { $apiGroupType = "Microsoft365" }
                        default {
                            Show-MessageBox -Message "Unbekannter oder nicht unterstützter Gruppentyp im UI: '$uiGroupType'." -Title "Fehler" -Type Error
                            return
                        }
                    }
                    
                    $membersInput = $script:txtGroupMembers.Text
                    $membersArray = @()
                    if (-not [string]::IsNullOrWhiteSpace($membersInput)) {
                        $membersArray = $membersInput -split '[;,`n`r]' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    }

                    $createParams = @{
                        GroupName = $script:txtGroupName.Text
                        GroupEmail = $script:txtGroupEmail.Text
                        GroupType = $apiGroupType
                        Members = $membersArray
                        Description = $script:txtGroupDescription.Text
                        # Owner wird in New-ExoGroupAction behandelt (aktueller Benutzer als Fallback für M365)
                    }
                    
                    $createResult = New-ExoGroupAction @createParams 
                    
                    if ($createResult) { # New-ExoGroupAction gibt das Gruppenobjekt oder $null zurück
                        Update-StatusBar -Message "Gruppe '$($createResult.DisplayName)' erfolgreich erstellt." -Type Success
                        Write-Log "Gruppe '$($createResult.DisplayName)' erfolgreich erstellt." -Type "Success"
                        $script:txtGroupName.Text = ""
                        $script:txtGroupEmail.Text = ""
                        $script:txtGroupMembers.Text = ""
                        $script:txtGroupDescription.Text = ""
                        Refresh-ExistingGroupsDropdown 
                    } # Fehlerfall wird in New-ExoGroupAction behandelt (Statusbar, Log)

                } catch {
                    $errMsg = "Fehler beim Erstellen der Gruppe: $($_.Exception.Message)"
                    Write-Log $errMsg -Type "Error"
                    Update-StatusBar -Message $errMsg -Type "Error"
                }
            } -ControlName "btnCreateGroup"
        }

        if ($null -ne $script:btnDeleteGroup) {
            Register-EventHandler -Control $script:btnDeleteGroup -Handler {
                try {
                    Write-Log "Button 'Gruppe löschen' geklickt." -Type "Info"

                    if (-not $script:isConnected) {
                        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                        return
                    }

                    $groupToDeleteObject = $null
                    if ($null -ne $script:cmbSelectExistingGroup.SelectedItem) {
                        $groupToDeleteObject = $script:cmbSelectExistingGroup.SelectedItem.Tag 
                    }

                    if ($null -eq $groupToDeleteObject -and -not [string]::IsNullOrWhiteSpace($script:txtGroupName.Text) ) {
                         Write-Log "Versuche Gruppe zum Löschen über txtGroupName '$($script:txtGroupName.Text)' zu finden..." -Type Info
                         $groupToDeleteObject = Get-ExoGroupRecipientAction -Identity $script:txtGroupName.Text # Annahme: Get-ExoGroupRecipientAction kann das
                    }

                    if ($null -eq $groupToDeleteObject) {
                        Show-MessageBox -Message "Bitte wählen Sie eine zu löschende Gruppe aus der Liste 'Gruppe auswählen' aus oder geben Sie einen gültigen Gruppennamen im Feld 'Gruppenname' ein." -Title "Keine Gruppe ausgewählt/gefunden" -Type Warning
                        return
                    }
                    
                    $groupDisplayNameForDelete = $groupToDeleteObject.DisplayName
                    $groupIdentityForDelete = $groupToDeleteObject.Identity

                    $confirmResult = Show-MessageBox -Message "Sind Sie sicher, dass Sie die Gruppe '$groupDisplayNameForDelete' (Identity: $groupIdentityForDelete) löschen möchten? Diese Aktion kann nicht rückgängig gemacht werden." -Title "Gruppe löschen bestätigen" -Type Question -Buttons YesNo
                    if ($confirmResult -ne "Yes") {
                        Write-Log "Löschen der Gruppe '$groupDisplayNameForDelete' abgebrochen." -Type "Info"
                        Update-StatusBar -Message "Löschen der Gruppe '$groupDisplayNameForDelete' abgebrochen." -Type Info
                        return
                    }
                    
                    $deleteResult = Remove-ExoGroupAction -Identity $groupIdentityForDelete 
                                        
                    if ($deleteResult) { # Remove-ExoGroupAction gibt $true/$false zurück
                        Update-StatusBar -Message "Gruppe '$groupDisplayNameForDelete' erfolgreich gelöscht." -Type Success
                        Write-Log "Gruppe '$groupDisplayNameForDelete' erfolgreich gelöscht." -Type "Success"
                        if ($script:txtGroupName.Text -eq $groupDisplayNameForDelete) { $script:txtGroupName.Text = "" }
                        Refresh-ExistingGroupsDropdown 
                    } # Fehlerfall wird in Remove-ExoGroupAction behandelt

                } catch {
                    $errMsg = "Fehler beim Löschen der Gruppe: $($_.Exception.Message)"
                    Write-Log $errMsg -Type "Error"
                    Update-StatusBar -Message $errMsg -Type "Error"
                }
            } -ControlName "btnDeleteGroup"
        }

        if ($null -ne $script:btnAddUserToGroup) {
            Register-EventHandler -Control $script:btnAddUserToGroup -Handler {
                try {
                    Write-Log "Button 'Benutzer hinzufügen' geklickt." -Type "Info"
                    if (-not $script:isConnected) {
                        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                        return
                    }
                    if ($null -eq $script:cmbSelectExistingGroup.SelectedItem) {
                        Show-MessageBox -Message "Bitte wählen Sie zuerst eine Gruppe aus der Liste aus." -Title "Keine Gruppe ausgewählt" -Type Warning
                        return
                    }
                    if ([string]::IsNullOrWhiteSpace($script:txtGroupUser.Text)) {
                        Show-MessageBox -Message "Bitte geben Sie den Benutzer (E-Mail-Adresse oder Alias) an, der hinzugefügt werden soll." -Title "Kein Benutzer angegeben" -Type Warning
                        return
                    }

                    $selectedGroupItem = $script:cmbSelectExistingGroup.SelectedItem
                    $groupObject = $selectedGroupItem.Tag 
                    $userEmail = $script:txtGroupUser.Text

                    # Parameter für Add-ExoGroupMemberAction sind -Identity (für Gruppe) und -MemberIdentity (für Benutzer)
                    $addResult = Add-ExoGroupMemberAction -Identity $groupObject.Identity -MemberIdentity $userEmail 

                    if ($addResult) {
                        Update-StatusBar -Message "Benutzer '$userEmail' zu Gruppe '$($groupObject.DisplayName)' hinzugefügt." -Type Success
                        Write-Log "Benutzer '$userEmail' zu Gruppe '$($groupObject.DisplayName)' hinzugefügt." -Type "Success"
                        $script:txtGroupUser.Text = "" 
                        # Mitgliederliste aktualisieren durch erneuten Aufruf von Get-ExoGroupMembersAction
                        Get-ExoGroupMembersAction -Identity $groupObject.Identity
                    }
                } catch {
                    $errMsg = "Fehler beim Hinzufügen des Benutzers zur Gruppe: $($_.Exception.Message)"
                    Write-Log $errMsg -Type "Error"
                    Update-StatusBar -Message $errMsg -Type "Error"
                }
            } -ControlName "btnAddUserToGroup"
        }

        if ($null -ne $script:btnRemoveUserFromGroup) {
            Register-EventHandler -Control $script:btnRemoveUserFromGroup -Handler {
                try {
                    Write-Log "Button 'Benutzer entfernen' geklickt." -Type "Info"
                     if (-not $script:isConnected) {
                        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                        return
                    }
                    if ($null -eq $script:cmbSelectExistingGroup.SelectedItem) {
                        Show-MessageBox -Message "Bitte wählen Sie zuerst eine Gruppe aus der Liste aus." -Title "Keine Gruppe ausgewählt" -Type Warning
                        return
                    }
                    if ([string]::IsNullOrWhiteSpace($script:txtGroupUser.Text)) {
                        Show-MessageBox -Message "Bitte geben Sie den Benutzer (E-Mail-Adresse oder Alias) an, der entfernt werden soll." -Title "Kein Benutzer angegeben" -Type Warning
                        return
                    }

                    $selectedGroupItem = $script:cmbSelectExistingGroup.SelectedItem
                    $groupObject = $selectedGroupItem.Tag
                    $userEmail = $script:txtGroupUser.Text
                    
                    $confirmResult = Show-MessageBox -Message "Sind Sie sicher, dass Sie den Benutzer '$userEmail' aus der Gruppe '$($groupObject.DisplayName)' entfernen möchten?" -Title "Benutzer entfernen bestätigen" -Type Question -Buttons YesNo
                    if ($confirmResult -ne "Yes") {
                        Write-Log "Entfernen des Benutzers '$userEmail' aus Gruppe '$($groupObject.DisplayName)' abgebrochen." -Type "Info"
                        Update-StatusBar -Message "Entfernen von '$userEmail' abgebrochen." -Type Info
                        return
                    }

                    # Parameter für Remove-ExoGroupMemberAction sind -GroupIdentity und -UserIdentity
                    $removeResult = Remove-ExoGroupMemberAction -GroupIdentity $groupObject.Identity -UserIdentity $userEmail 

                    if ($removeResult) {
                        Update-StatusBar -Message "Benutzer '$userEmail' aus Gruppe '$($groupObject.DisplayName)' entfernt." -Type Success
                        Write-Log "Benutzer '$userEmail' aus Gruppe '$($groupObject.DisplayName)' entfernt." -Type "Success"
                        $script:txtGroupUser.Text = "" 
                        Get-ExoGroupMembersAction -Identity $groupObject.Identity
                    }
                } catch {
                    $errMsg = "Fehler beim Entfernen des Benutzers aus der Gruppe: $($_.Exception.Message)"
                    Write-Log $errMsg -Type "Error"
                    Update-StatusBar -Message $errMsg -Type "Error"
                }
            } -ControlName "btnRemoveUserFromGroup"
        }

        if ($null -ne $script:btnShowGroupMembers) {
            Register-EventHandler -Control $script:btnShowGroupMembers -Handler {
                try {
                    Write-Log "Button 'Mitglieder und Einstellungen anzeigen' geklickt." -Type "Info"
                    if (-not $script:isConnected) {
                        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                        return
                    }
                    if ($null -eq $script:cmbSelectExistingGroup.SelectedItem) {
                        Show-MessageBox -Message "Bitte wählen Sie zuerst eine Gruppe aus der Liste aus." -Title "Keine Gruppe ausgewählt" -Type Warning
                        if ($null -ne $script:lstGroupMembers) { $script:lstGroupMembers.Items.Clear(); $script:lstGroupMembers.Tag = $null }
                        return
                    }

                    $selectedGroupItem = $script:cmbSelectExistingGroup.SelectedItem
                    $groupObject = $selectedGroupItem.Tag
                    
                    Get-ExoGroupMembersAction -Identity $groupObject.Identity 
                    # Get-ExoGroupMembersAction aktualisiert die Mitgliederliste, die Checkboxen und die Statusleiste.
                    # Keine weiteren Aktionen hier notwendig.
                } catch {
                    $errMsg = "Fehler beim Anzeigen der Gruppenmitglieder/Einstellungen: $($_.Exception.Message)"
                    Write-Log $errMsg -Type "Error"
                    Update-StatusBar -Message $errMsg -Type "Error"
                    if ($null -ne $script:lstGroupMembers) { $script:lstGroupMembers.Items.Clear(); $script:lstGroupMembers.Tag = $null }
                }
            } -ControlName "btnShowGroupMembers"
        }

        if ($null -ne $script:btnUpdateGroupSettings) {
            Register-EventHandler -Control $script:btnUpdateGroupSettings -Handler {
                try {
                    Write-Log "Button 'Einstellungen aktualisieren' geklickt." -Type "Info"
                    if (-not $script:isConnected) {
                        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                        return
                    }
                    if ($null -eq $script:cmbSelectExistingGroup.SelectedItem) {
                        Show-MessageBox -Message "Bitte wählen Sie zuerst eine Gruppe aus der Liste aus." -Title "Keine Gruppe ausgewählt" -Type Warning
                        return
                    }
                    
                    $selectedGroupItem = $script:cmbSelectExistingGroup.SelectedItem
                    $groupObject = $selectedGroupItem.Tag

                    $settingsParams = @{
                        Identity = $groupObject.Identity
                    }
                    # Nur gebundene Parameter übergeben, wenn Checkboxen aktiv sind
                    if ($script:chkHiddenFromGAL.IsEnabled)       { $settingsParams.HiddenFromAddressListsEnabled = [bool]$script:chkHiddenFromGAL.IsChecked }
                    # RequireSenderAuthenticationEnabled und AllowExternalSenders sind oft aneinander gekoppelt.
                    # Update-ExoGroupSettingsAction sollte die Logik dafür haben.
                    if ($script:chkRequireSenderAuth.IsEnabled)   { $settingsParams.RequireSenderAuthenticationEnabled = [bool]$script:chkRequireSenderAuth.IsChecked }
                    if ($script:chkAllowExternalSenders.IsEnabled){ $settingsParams.AllowExternalSenders = [bool]$script:chkAllowExternalSenders.IsChecked }


                    $updateResult = Update-ExoGroupSettingsAction @settingsParams 

                    if ($updateResult) {
                        # Status-Update erfolgt in Update-ExoGroupSettingsAction
                        Write-Log "Aufruf von Update-ExoGroupSettingsAction für '$($groupObject.DisplayName)' war erfolgreich (laut Rückgabewert)." -Type "Info"
                        # Optional: Einstellungen neu laden, um UI zu bestätigen
                        Get-ExoGroupMembersAction -Identity $groupObject.Identity
                    }

                } catch {
                    $errMsg = "Fehler beim Aktualisieren der Gruppeneinstellungen: $($_.Exception.Message)"
                    Write-Log $errMsg -Type "Error"
                    Update-StatusBar -Message $errMsg -Type "Error"
                }
            } -ControlName "btnUpdateGroupSettings"
        }
        
        if ($null -ne $script:btnExportGroupMembers) {
            Register-EventHandler -Control $script:btnExportGroupMembers -Handler {
                try {
                    Write-Log "Button 'Mitgliederliste exportieren' geklickt." -Type "Info"
                    if (-not $script:isConnected) {
                        Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                        return
                    }
                    # Verwende $script:lstGroupMembers.Tag, da dort die rohen Mitgliederobjekte gespeichert werden
                    if ($null -eq $script:lstGroupMembers.Tag -or ($script:lstGroupMembers.Tag -is [array] -and $script:lstGroupMembers.Tag.Count -eq 0)) {
                         Show-MessageBox -Message "Es sind keine Mitglieder zum Exportieren vorhanden. Bitte laden Sie zuerst die Mitglieder einer Gruppe." -Title "Keine Daten für Export" -Type Info
                        return
                    }
                    
                    $groupNameForExport = "UnbekannteGruppe"
                    if ($null -ne $script:cmbSelectExistingGroup.SelectedItem) {
                         $groupNameForExport = $script:cmbSelectExistingGroup.SelectedItem.Content -replace '[^a-zA-Z0-9_.-]', '_' 
                    }

                    $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
                    $saveFileDialog.Filter = "CSV-Datei (*.csv)|*.csv|Textdatei (*.txt)|*.txt"
                    $saveFileDialog.FileName = "Mitglieder_$($groupNameForExport)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                    $saveFileDialog.Title = "Mitgliederliste exportieren als"

                    if ($saveFileDialog.ShowDialog() -eq $true) {
                        $filePath = $saveFileDialog.FileName
                        $membersToExport = $script:lstGroupMembers.Tag # Dies ist das Array der Mitgliederobjekte
                        
                        # Wähle relevante Eigenschaften für den Export aus
                        $membersToExport | Select-Object DisplayName, PrimarySmtpAddress, RecipientType, Name, Alias | 
                            Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Delimiter ";" 
                        
                        Update-StatusBar -Message "Mitgliederliste erfolgreich nach '$filePath' exportiert." -Type Success
                        Write-Log "Mitgliederliste erfolgreich nach '$filePath' exportiert." -Type "Success"
                        Show-MessageBox -Message "Die Mitgliederliste wurde erfolgreich exportiert nach:`n$filePath" -Title "Export erfolgreich" -Type Info
                    } else {
                        Write-Log "Export der Mitgliederliste vom Benutzer abgebrochen." -Type "Info"
                        Update-StatusBar -Message "Export abgebrochen." -Type Info
                    }

                } catch {
                    $errMsg = "Fehler beim Exportieren der Mitgliederliste: $($_.Exception.Message)"
                    Write-Log $errMsg -Type "Error"
                    Update-StatusBar -Message "Fehler beim Export: $errMsg" -Type "Error"
                    Show-MessageBox -Message "Fehler beim Exportieren der Mitgliederliste:`n$errMsg" -Title "Exportfehler" -Type Error
                }
            } -ControlName "btnExportGroupMembers"
        }

        if ($null -ne $script:helpLinkGroups) {
            $script:helpLinkGroups.Add_MouseLeftButtonDown({
                try {
                    if (Test-Path Function:\Show-HelpDialog) { Show-HelpDialog -Topic "Groups" }
                    else { Write-Log "Hilfefunktion Show-HelpDialog nicht gefunden." -Type Warning }
                } catch { Write-Log "Fehler beim Öffnen des Hilfe-Dialogs für Gruppen: $($_.Exception.Message)" -Type Error }
            })
            $script:helpLinkGroups.Add_MouseEnter({ try { $this.Cursor = [System.Windows.Input.Cursors]::Hand } catch {} })
            $script:helpLinkGroups.Add_MouseLeave({ try { $this.Cursor = [System.Windows.Input.Cursors]::Arrow } catch {} })
        }

        Write-Log "Initialize-GroupsTab erfolgreich abgeschlossen." -Type "Info"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "FATALER Fehler in Initialize-GroupsTab: $errorMsg `n$($_.ScriptStackTrace)" -Type "Error"
        Update-StatusBar -Message "Schwerer Fehler bei Initialisierung des Gruppen-Tabs: $errorMsg" -Type "Error"
        return $false
    }
}
#endregion Initialize-GroupsTab

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

        # GUI-Elemente finden gemäß XAML-Definition
        $script:txtRegionMailbox = $script:Form.FindName("txtRegionMailbox")
        $script:cmbRegionLanguage = $script:Form.FindName("cmbRegionLanguage")
        $script:cmbRegionTimezone = $script:Form.FindName("cmbRegionTimezone")
        $script:btnSetRegionSettings = $script:Form.FindName("btnSetRegionSettings")
        $script:txtGetRegionMailbox = $script:Form.FindName("txtGetRegionMailbox")
        $script:btnGetRegionSettings = $script:Form.FindName("btnGetRegionSettings")
        $script:txtRegionResult = $script:Form.FindName("txtRegionResult")
        $script:helpLinkRegionSettings = $script:Form.FindName("helpLinkRegionSettings")
        $script:cmbRegionDateFormat = $script:Form.FindName("cmbRegionDateFormat")
        $script:cmbRegionTimeFormat = $script:Form.FindName("cmbRegionTimeFormat")
        $script:chkRegionDefaultFolderNameMatchingUserLanguage = $script:Form.FindName("chkRegionDefaultFolderNameMatchingUserLanguage")
        
        # ... (Control-Validierung bleibt gleich) ...
        $requiredControlsMap = @{
            "txtRegionMailbox"     = $script:txtRegionMailbox
            "cmbRegionLanguage"    = $script:cmbRegionLanguage
            "cmbRegionTimezone"    = $script:cmbRegionTimezone
            "btnSetRegionSettings" = $script:btnSetRegionSettings
            "txtGetRegionMailbox"  = $script:txtGetRegionMailbox
            "btnGetRegionSettings" = $script:btnGetRegionSettings
            "txtRegionResult"      = $script:txtRegionResult
            "cmbRegionDateFormat" = $script:cmbRegionDateFormat
            "cmbRegionTimeFormat" = $script:cmbRegionTimeFormat
            "chkRegionDefaultFolderNameMatchingUserLanguage" = $script:chkRegionDefaultFolderNameMatchingUserLanguage
        }

        $missingControls = [System.Collections.Generic.List[string]]::new()
        foreach ($controlNameInMap in $requiredControlsMap.Keys) {
            if ($null -eq $requiredControlsMap[$controlNameInMap]) {
                $missingControls.Add($controlNameInMap)
            }
        }

        if ($missingControls.Count -gt 0) {
            $missingControlsListString = $missingControls -join ", "
            $logMessage = "Ein oder mehrere erforderliche Steuerelemente im Regionaleinstellungen-Tab nicht gefunden: $missingControlsListString. Bitte XAML 'x:Name' Attribute prüfen."
            Write-Log $logMessage -Type Error
            Log-Action "Fehler: Wichtige Steuerelemente für Regionaleinstellungen ($missingControlsListString) nicht gefunden. XAML-Namen prüfen!"
            Show-MessageBox -Message $logMessage -Title "Initialisierungsfehler"
            return $false
        }

        try {
            Populate-LanguageComboBox -ComboBox $script:cmbRegionLanguage
        } catch {
            Write-Log "FEHLER beim initialen Befüllen der Sprachen-ComboBox: $($_.Exception.Message)" -Type Error
        }

        try {
            Populate-TimezoneComboBox -ComboBox $script:cmbRegionTimezone -CultureName "DEFAULT_ALL"
        } catch {
            Write-Log "FEHLER beim initialen Befüllen der Zeitzonen-ComboBox: $($_.Exception.Message)" -Type Error
        }
        
        try {
            Populate-DateFormatComboBox -ComboBox $script:cmbRegionDateFormat -CultureName "" 
        } catch {
            Write-Log "FEHLER beim initialen Befüllen der Datumsformat-ComboBox: $($_.Exception.Message)" -Type Error
        }
        
        # Zeitformat-ComboBox initial mit DEFAULT befüllen
        try {
            Populate-TimeFormatComboBox -ComboBox $script:cmbRegionTimeFormat -CultureName ""
        } catch {
            Write-Log "FEHLER beim initialen Befüllen der Zeitformat-ComboBox: $($_.Exception.Message)" -Type Error
        }
        
        $script:chkRegionDefaultFolderNameMatchingUserLanguage.IsChecked = $null 
        Write-Log "chkRegionDefaultFolderNameMatchingUserLanguage.IsChecked auf \$null (unbestimmt) gesetzt." -Type Debug
        
        # Event-Handler für Sprachauswahl-Änderung ERWEITERN
        $script:cmbRegionLanguage.Add_SelectionChanged({
            param($sender, $e)
            try {
                $selectedLanguageItem = $sender.SelectedItem
                $cultureNameForFormats = "" 
                $cultureNameForTimezones = "DEFAULT_ALL" 

                if ($null -ne $selectedLanguageItem -and $null -ne $selectedLanguageItem.Tag -and $selectedLanguageItem.Tag -ne "") {
                    $selectedCultureTag = $selectedLanguageItem.Tag.ToString()
                    $cultureNameForFormats = $selectedCultureTag
                    $cultureNameForTimezones = $selectedCultureTag 
                }
                
                Write-Log "Sprachauswahl geändert. Kultur für Formate: '$cultureNameForFormats', Kultur für Zeitzonen: '$cultureNameForTimezones'. Aktualisiere abhängige ComboBoxen." -Type Debug
                Populate-DateFormatComboBox -ComboBox $script:cmbRegionDateFormat -CultureName $cultureNameForFormats
                Populate-TimeFormatComboBox -ComboBox $script:cmbRegionTimeFormat -CultureName $cultureNameForFormats # NEU
                Populate-TimezoneComboBox -ComboBox $script:cmbRegionTimezone -CultureName $cultureNameForTimezones
            } catch {
                $errorMsgSelChange = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler im SelectionChanged Event der Sprach-ComboBox."
                Write-Log "FEHLER im cmbRegionLanguage.SelectionChanged: $errorMsgSelChange" -Type Error
            }
        })
        Write-Log "Event-Handler für cmbRegionLanguage.SelectionChanged registriert (inkl. Datums-, Zeitformat- und Zeitzonen-Update)." -Type Debug
        
        # ... (restliche Event-Handler für Buttons und HelpLink bleiben gleich) ...
        if ($null -ne $script:btnGetRegionSettings) {
            $script:btnGetRegionSettings.Add_Click({
                try {
                    Invoke-GetRegionSettingsAction
                } catch {
                    $errorMsgAction = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Abrufen der Regionaleinstellungen."
                    Write-Log $errorMsgAction -Type Error; Log-Action "FEHLER bei Aktion 'Einstellungen abrufen' ($($script:btnGetRegionSettings.Name)): $errorMsgAction"
                    Show-MessageBox -Message "Ein Fehler ist beim Abrufen der Einstellungen aufgetreten: $($_.Exception.Message)" -Title "Aktionsfehler"
                }
            })
            Write-Log "Event-Handler für '$($script:btnGetRegionSettings.Name)' registriert." -Type Debug
        }

        if ($null -ne $script:btnSetRegionSettings) {
            $script:btnSetRegionSettings.Add_Click({
                try {
                    Invoke-SetRegionSettingsAction
                } catch {
                    $errorMsgAction = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Anwenden der Regionaleinstellungen."
                    Write-Log $errorMsgAction -Type Error; Log-Action "FEHLER bei Aktion 'Einstellungen anwenden' ($($script:btnSetRegionSettings.Name)): $errorMsgAction"
                    Show-MessageBox -Message "Ein Fehler ist beim Anwenden der Einstellungen aufgetreten: $($_.Exception.Message)" -Title "Aktionsfehler"
                }
            })
            Write-Log "Event-Handler für '$($script:btnSetRegionSettings.Name)' registriert." -Type Debug
        }

        if ($null -ne $script:helpLinkRegionSettings) {
            $script:helpLinkRegionSettings.Add_MouseLeftButtonDown({ Show-HelpDialog -Topic "RegionSettings" }) 
            $script:helpLinkRegionSettings.Add_MouseEnter({ $this.Cursor = [System.Windows.Input.Cursors]::Hand; $this.TextDecorations = [System.Windows.TextDecorations]::Underline })
            $script:helpLinkRegionSettings.Add_MouseLeave({ $this.TextDecorations = $null; $this.Cursor = [System.Windows.Input.Cursors]::Arrow })
            Write-Log "Event-Handler für '$($script:helpLinkRegionSettings.Name)' registriert." -Type Debug
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
                $currentTopicForEvent = $topic # Den aktuellen Wert von $topic für diesen spezifischen Event-Handler speichern
                $link.Add_MouseLeftButtonDown({
                    param($sender, $e) # Explizite Parameter für den Event-Handler sind gute Praxis

                    # Den für diesen Klick relevanten Topic-Wert verwenden
                    $resolvedTopicOnClick = $currentTopicForEvent 

                    # Sicherstellen, dass $resolvedTopicOnClick nicht leer ist
                    if ([string]::IsNullOrWhiteSpace($resolvedTopicOnClick)) {
                    }
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

Dieses Tool vereinfacht die Verwaltung gängiger Exchange Online Aufgaben über eine benutzerfreundliche grafische Oberfläche.
Es ermöglicht Administratoren und Helpdesk-Mitarbeitern, schnell und effizient Routineaufgaben durchzuführen, ohne komplexe PowerShell-Skripte manuell ausführen zu müssen.

(c) WWW.PHINIT.DE | Andreas Hepp | $(Get-Date -Format yyyy)
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
                Log-Action "Fehler beim Setzen der Version: $($_.Exception.Message)"
            }
        } catch {
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

        # Sicherstellen, dass $script:config und die benötigten Hauptschlüssel existieren
        # Get-RegistryConfig sollte General, Paths, UI bereits initialisiert haben, falls Werte vorhanden sind.
        # Hier stellen wir sicher, dass sie existieren, auch wenn sie leer sind, für die Logik unten.
        if ($null -eq $script:config) { $script:config = @{} }
        if (-not $script:config.ContainsKey("General")) { $script:config["General"] = @{} }
        if (-not $script:config.ContainsKey("Paths")) { $script:config["Paths"] = @{} }
        # Für Fallback-Logik von alten Konfigurationsstrukturen oder wenn Get-RegistryConfig sie nicht erstellt
        if (-not $script:config.ContainsKey("Defaults")) { $script:config["Defaults"] = @{} }
        if (-not $script:config.ContainsKey("Logging")) { $script:config["Logging"] = @{} }
        if (-not $script:config.ContainsKey("Appearance")) { $script:config["Appearance"] = @{} }


        # Aktuelle Einstellungen laden und anzeigen (aus $script:config)
        # Priorität: Neue Struktur (General, Paths), dann alte Struktur (Defaults, Logging, Appearance) als Fallback

        # DefaultUser
        if ($null -ne $txtDefaultUser) {
            if ($script:config["General"].ContainsKey("DefaultUser")) {
                $txtDefaultUser.Text = $script:config["General"]["DefaultUser"]
            } elseif ($script:config["Defaults"].ContainsKey("DefaultUser")) { # Fallback
                $txtDefaultUser.Text = $script:config["Defaults"]["DefaultUser"]
            }
        }

        # DebugEnabled
        if ($null -ne $chkEnableDebug) {
            if ($script:config["General"].ContainsKey("Debug")) {
                $chkEnableDebug.IsChecked = ($script:config["General"]["Debug"] -eq "1")
            } elseif ($script:config["Logging"].ContainsKey("DebugEnabled")) { # Fallback
                $debugEnabledValue = $script:config["Logging"]["DebugEnabled"]
                if ($debugEnabledValue -is [string] -and $debugEnabledValue -match '^(true|false)$') {
                    $chkEnableDebug.IsChecked = [System.Convert]::ToBoolean($debugEnabledValue)
                } elseif ($debugEnabledValue -is [bool]) {
                    $chkEnableDebug.IsChecked = $debugEnabledValue
                } else {
                    $chkEnableDebug.IsChecked = $false # Fallback bei ungültigem Wert
                    Write-Log "Ungültiger Wert für DebugEnabled in Config (Logging): '$debugEnabledValue'. Setze auf false." -Type Warning
                }
            } else {
                $chkEnableDebug.IsChecked = $false # Standard-Fallback
            }
        }

        # LogPath / LogDirectory
        if ($null -ne $txtLogPath) {
            if ($script:config["Paths"].ContainsKey("LogPath")) {
                $txtLogPath.Text = $script:config["Paths"]["LogPath"]
            } elseif ($script:config["Logging"].ContainsKey("LogDirectory")) { # Fallback
                $txtLogPath.Text = $script:config["Logging"]["LogDirectory"]
            }
        }
        
        # Theme
        if ($null -ne $cmbTheme) {
            $themeToLoad = $null
            if ($script:config["General"].ContainsKey("Theme")) { # Neuer Schlüssel für Theme-Tag
                $themeToLoad = $script:config["General"]["Theme"]
            } elseif ($script:config["Appearance"].ContainsKey("Theme")) { # Fallback auf alte Struktur
                $themeToLoad = $script:config["Appearance"]["Theme"]
            } elseif ($script:config["General"].ContainsKey("ThemeColor")) { # Fallback auf ThemeColor, falls "Theme" nicht existiert
                 # Hier müsste eine Logik stehen, die ThemeColor (z.B. #0078D7) auf einen Tag (z.B. "Light") mappt.
                 # Fürs Erste wird dies ignoriert und der Standard-Theme verwendet, wenn nur ThemeColor vorhanden ist.
                 Write-Log "ThemeColor ('$($script:config["General"]["ThemeColor"])') gefunden, aber 'Theme' (Tag) wird bevorzugt. Mapping nicht implementiert." -Type Debug
            }

            if ($null -ne $themeToLoad) {
                foreach($item in $cmbTheme.Items) {
                    if ($item.Tag -eq $themeToLoad) {
                        $cmbTheme.SelectedItem = $item
                        break
                    }
                }
                if ($null -eq $cmbTheme.SelectedItem) { $cmbTheme.SelectedIndex = 0 } # Fallback, wenn Tag nicht in ComboBox
            } else {
                $cmbTheme.SelectedIndex = 0 # Standardwert setzen, falls nichts in Config
            }
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
                $newDefaultUser = if($null -ne $txtDefaultUser) { $txtDefaultUser.Text.Trim() } else { "" }
                $newDebugEnabled = if($null -ne $chkEnableDebug) { $chkEnableDebug.IsChecked } else { $false }
                $newLogPath = if($null -ne $txtLogPath) { $txtLogPath.Text.Trim() } else { "" }
                $newThemeTag = if($null -ne $cmbTheme.SelectedItem) { $cmbTheme.SelectedItem.Tag } else { "Light" } # Standard-Theme-Tag

                # Validieren
                if (-not $newLogPath) {
                    [System.Windows.MessageBox]::Show("Bitte geben Sie ein Log-Verzeichnis an.", "Validierung fehlgeschlagen", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                if (-not (Test-Path (Split-Path -Path $newLogPath -Parent) -PathType Container)) {
                    [System.Windows.MessageBox]::Show("Das übergeordnete Verzeichnis des Log-Pfads existiert nicht. Bitte wählen Sie einen gültigen Pfad.", "Validierung fehlgeschlagen", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }


                # $script:config mit den neuen Werten aktualisieren (konsistent mit Get-RegistryConfig Struktur)
                $script:config["General"]["DefaultUser"] = $newDefaultUser
                $script:config["General"]["Debug"] = if($newDebugEnabled) {"1"} else {"0"}
                $script:config["Paths"]["LogPath"] = $newLogPath
                $script:config["General"]["Theme"] = $newThemeTag # Speichere den Tag, nicht die Farbe

                # Konfiguration in Registry speichern
                $saveResult = $false
                try {
                    if (-not (Test-Path $script:registryPath)) {
                        Write-Log "Registry-Pfad '$($script:registryPath)' nicht gefunden, wird erstellt." -Type Info
                        New-Item -Path $script:registryPath -Force | Out-Null
                    }
                    Set-ItemProperty -Path $script:registryPath -Name "DefaultUser" -Value $newDefaultUser -Type String -ErrorAction Stop
                    Set-ItemProperty -Path $script:registryPath -Name "Debug" -Value $script:config["General"]["Debug"] -Type String -ErrorAction Stop
                    Set-ItemProperty -Path $script:registryPath -Name "LogPath" -Value $newLogPath -Type String -ErrorAction Stop
                    Set-ItemProperty -Path $script:registryPath -Name "Theme" -Value $newThemeTag -Type String -ErrorAction Stop
                    
                    Write-Log "Einstellungen erfolgreich in die Registry geschrieben." -Type Success
                    $saveResult = $true
                }
                catch {
                    $errorDetail = Get-FormattedError -ErrorRecord $_ -DefaultText "Unbekannter Fehler beim Schreiben in die Registry."
                    Write-Log "Fehler beim Speichern der Einstellungen in die Registry: $errorDetail" -Type Error
                    Log-Action "Registry Speicherfehler: $errorDetail"
                    [System.Windows.MessageBox]::Show("Fehler beim Speichern der Einstellungen in die Registry:`n$($_.Exception.Message)", "Speicherfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    $saveResult = $false
                }

                if ($saveResult) {
                    # Laufzeitvariablen aktualisieren
                    $script:debugMode = $newDebugEnabled # Direkt die boolesche Variable für den aktuellen Lauf setzen
                    $script:DebugPreference = if($newDebugEnabled) { 'Continue' } else { 'SilentlyContinue' } # Für Write-Debug etc.
                    $script:LogDir = $newLogPath
                    
                    # Logging neu initialisieren oder Pfad aktualisieren
                    # Annahme: Initialize-Logging verwendet $script:LogDir und $script:debugMode (oder $script:DebugPreference)
                    Initialize-Logging 
                    
                    # Theme zur Laufzeit anwenden (falls Funktion dafür existiert)
                    # Apply-Theme $newThemeTag # Beispiel, Funktion müsste existieren

                    Update-GuiText -TextElement $script:txtStatus -Message "Einstellungen erfolgreich gespeichert und angewendet."
                    Log-Action "Einstellungen wurden gespeichert und angewendet."
                    $settingsWindow.Close()
                } else {
                    Update-GuiText -TextElement $script:txtStatus -Message "Fehler beim Speichern der Einstellungen."
                }
            })
        }

        # Event Handler für "Abbrechen"
        if ($null -ne $btnCancelSettings) {
            $btnCancelSettings.Add_Click({ $settingsWindow.Close() })
        }

        # Fenstereigentümer setzen, damit es modal zum Hauptfenster ist
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