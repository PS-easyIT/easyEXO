# # Globale Variablen für asynchrone Tab-Initialisierung initialisieren
$script:tabInitRunspace = $null
$script:tabInitAsync = $null
$script:tabInitTimer = $null
$script:tabInitStartTime = $null

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
        Add-Content -Path $script:logFilePath -Value "[$timestamp] $sanitizedMessage"
        
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
function Connect-OwnExchangeOnline {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Verbindungsversuch zu Exchange Online..." -Type "Info"
        
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
        Write-Log "Versuche, Verbindung zu Exchange Online herzustellen für $script:userPrincipalName" -Type "Info"
        ExchangeOnlineManagement\Connect-ExchangeOnline @connectParams
        
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
                $txtStatus.Text = "Alle Module erfolgreich installiert. Bitte mit Exchange Online verbinden ."
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
# Abschnitt: StatusBar Aktualisierung
# -------------------------------------------------

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
# Abschnitt: Exchange Health Check-Funktionen
# -------------------------------------------------
function Get-OwnRetentionCompliancePolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $true)]
        [string]$Identity,
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$DistributionDetail,
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ErrorPolicyOnly,
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ExcludeTeamsPolicy,
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IncludeTestModeResults,
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PriorityCleanup,
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$RetentionRuleTypes,
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$TeamsPolicyOnly
    )
    try {
        Write-Log "Rufe Aufbewahrungsrichtlinien (Retention Compliance Policies) ab..." -Type "Debug"

        # Überprüfen, ob das Cmdlet verfügbar ist
        $cmdletAvailable = Get-Command Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue
        if (-not $cmdletAvailable) {
            throw "Das Cmdlet 'Get-RetentionCompliancePolicy' ist nicht verfügbar."
        }

        # Parameter für das eigentliche Cmdlet vorbereiten
        $params = @{
            ErrorAction = 'Stop'
        }
        if ($PSBoundParameters.ContainsKey('Identity')) {
            $params.Identity = $Identity
        }
        if ($PSBoundParameters.ContainsKey('DistributionDetail')) {
            $params.DistributionDetail = $true
        }
        if ($PSBoundParameters.ContainsKey('ErrorPolicyOnly')) {
            $params.ErrorPolicyOnly = $true
        }
        if ($PSBoundParameters.ContainsKey('ExcludeTeamsPolicy')) {
            $params.ExcludeTeamsPolicy = $true
        }
        if ($PSBoundParameters.ContainsKey('IncludeTestModeResults')) {
            $params.IncludeTestModeResults = $true
        }
        if ($PSBoundParameters.ContainsKey('PriorityCleanup')) {
            $params.PriorityCleanup = $true
        }
        if ($PSBoundParameters.ContainsKey('RetentionRuleTypes')) {
            $params.RetentionRuleTypes = $true
        }
        if ($PSBoundParameters.ContainsKey('TeamsPolicyOnly')) {
            $params.TeamsPolicyOnly = $true
        }

        # Das eigentliche Exchange Online Cmdlet aufrufen
        $policies = Get-RetentionCompliancePolicy @params

        # Extrahieren und Anzeigen der spezifischen Eigenschaften für jede Richtlinie
        $results = foreach ($policy in $policies) {
            [PSCustomObject]@{
                Name    = $policy.Name
                Workload = "All" # Platzhalter, da die Beschreibung erwähnt, dass alle Workloads angezeigt werden
                Enabled = $policy.Enabled
                Mode    = $policy.Mode
            }
        }

        Write-Log "Aufbewahrungsrichtlinien erfolgreich abgerufen." -Type "Debug"
        return $results
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "access denied|permissions") {
            $permissionLink = "https://learn.microsoft.com/en-us/purview/purview-permissions"
            $details = "Fehler: Berechtigungsproblem. Stellen Sie sicher, dass Sie die erforderlichen Berechtigungen haben. Weitere Informationen finden Sie hier: $permissionLink"
        } else {
            $details = "Fehler beim Abrufen der Aufbewahrungsrichtlinien: $errorMsg"
        }
        Write-Log $details -Type "Error"
        Log-Action $details
        throw $details
    }
}

function Get-OwnDlpCompliancePolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $true)]
        [string]$Identity,
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$DistributionDetail,
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ErrorPolicyOnly
    )
    try {
        Write-Log "Rufe DLP-Richtlinien (DLP Compliance Policies) ab..." -Type "Debug"

        # Überprüfen, ob das Cmdlet verfügbar ist
        $cmdletAvailable = Get-Command Get-DlpCompliancePolicy -ErrorAction SilentlyContinue
        if (-not $cmdletAvailable) {
            throw "Das Cmdlet 'Get-DlpCompliancePolicy' ist nicht verfügbar. Dies kann auf fehlende Lizenzen (z.B. Microsoft 365 E3/E5) oder Berechtigungen hindeuten."
        }

        # Parameter für das eigentliche Cmdlet vorbereiten
        $params = @{
            ErrorAction = 'Stop'
        }
        if ($PSBoundParameters.ContainsKey('Identity')) {
            $params.Identity = $Identity
        }
        if ($PSBoundParameters.ContainsKey('DistributionDetail')) {
            $params.DistributionDetail = $true
        }
        if ($PSBoundParameters.ContainsKey('ErrorPolicyOnly')) {
            $params.ErrorPolicyOnly = $true
        }

        # Das eigentliche Exchange Online Cmdlet aufrufen
        $policies = Get-DlpCompliancePolicy @params

        # Extrahieren und Anzeigen der spezifischen Eigenschaften für jede Richtlinie
        $results = foreach ($policy in $policies) {
            [PSCustomObject]@{
                Name    = $policy.Name
                Workload = "All" # Platzhalter
                Enabled = $policy.Enabled
                Mode    = $policy.Mode
            }
        }

        Write-Log "DLP-Richtlinien erfolgreich abgerufen." -Type "Debug"
        return $results
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "access denied|permissions|nicht verfügbar") {
            $permissionLink = "https://learn.microsoft.com/en-us/purview/purview-permissions"
            $details = "Fehler: Berechtigungsproblem oder fehlende Lizenz. Stellen Sie sicher, dass Sie die erforderlichen Berechtigungen und Lizenzen (z.B. M365 E3/E5) haben. Weitere Informationen finden Sie hier: $permissionLink"
        } else {
            $details = "Fehler beim Abrufen der DLP-Richtlinien: $errorMsg"
        }
        Write-Log $details -Type "Error"
        Log-Action $details
        throw $details
    }
}

function Start-ExchangeHealthCheck {
    [CmdletBinding()]
    param()

    try {
        # GUI-Updates vor dem Start
        Update-StatusBar -Message "Starte Exchange Health Check..." -Type "Info"
        $script:pbHealthCheck.IsIndeterminate = $true
        $script:pbHealthCheck.Visibility = [System.Windows.Visibility]::Visible
        $script:lvHealthCheckResults.Items.Clear()

        # Erzwinge ein Neuzeichnen der GUI, damit die Änderungen sichtbar werden
        $script:lvHealthCheckResults.Dispatcher.Invoke([Action]{}, "Background") | Out-Null

        # Definition der auszuführenden Checks
        $checks = @(
            { Test-ExchangeConnection },
            { Test-MailboxService },
            { Test-TransportService },
            { Test-ComplianceService },
            { Test-SecurityServices },
            { Test-LicenseStatus },
            { Test-TenantConfiguration }
        )

        $allResults = [System.Collections.Generic.List[object]]::new()

        # Führe jeden Check nacheinander aus
        foreach ($check in $checks) {
            $checkName = ($check.ToString() -split ' ')[1].TrimEnd('}') # Extrahiert den Funktionsnamen
            Update-StatusBar -Message "Prüfe: $checkName..." -Type "Info"
            Write-Log "Prüfe: $checkName..." -Type "Info"
            
            # Führe den Check aus und füge das Ergebnis der Liste hinzu
            $result = & $check
            if ($null -ne $result) {
                $script:lvHealthCheckResults.Items.Add($result)
                $allResults.Add($result)
                # Erzwinge ein Neuzeichnen der GUI, um das neue Element anzuzeigen
                $script:lvHealthCheckResults.Dispatcher.Invoke([Action]{}, "Background") | Out-Null
            }
        }

        # GUI-Updates nach Abschluss aller Checks
        $script:pbHealthCheck.IsIndeterminate = $false
        $script:pbHealthCheck.Visibility = [System.Windows.Visibility]::Hidden

        # Ergebnisse zusammenfassen
        $okCount = ($allResults | Where-Object { $_.Status -eq "OK" }).Count
        $warningCount = ($allResults | Where-Object { $_.Status -eq "Warnung" }).Count
        $errorCount = ($allResults | Where-Object { $_.Status -eq "Fehler" }).Count

        $summaryMessage = "Health Check abgeschlossen: $okCount OK, $warningCount Warnungen, $errorCount Fehler"
        Update-StatusBar -Message $summaryMessage -Type "Success"

        Write-Log $summaryMessage -Type "Info"
        Log-Action $summaryMessage
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Health Check: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Health Check: $errorMsg"
        
        # GUI im Fehlerfall zurücksetzen
        $script:pbHealthCheck.IsIndeterminate = $false
        $script:pbHealthCheck.Visibility = [System.Windows.Visibility]::Hidden
        Update-StatusBar -Message "Fehler beim Health Check: $errorMsg" -Type "Error"
    }
}

function Test-ExchangeConnection {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Teste Exchange Online Verbindung (erweiterte Prüfung)..." -Type "Debug"
        
        $details = [System.Collections.Generic.List[string]]::new()
        $overallStatus = "OK" # Start with OK, downgrade on any failure

        # 1. Grundlegende Internetverbindung prüfen
        Write-Log "Prüfe grundlegende Internetverbindung..." -Type "Debug"
        if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet) {
            $details.Add("✅ Grundlegende Internetverbindung ist vorhanden.")
        } else {
            $details.Add("❌ FEHLER: Keine grundlegende Internetverbindung (Ping an 8.8.8.8 fehlgeschlagen).")
            $overallStatus = "Fehler"
        }

        # 2. DNS-Auflösung für M365-Endpunkte prüfen
        Write-Log "Prüfe DNS-Auflösung für M365-Endpunkte..." -Type "Debug"
        $dnsTest = Resolve-DnsName -Name "outlook.office365.com" -ErrorAction SilentlyContinue
        if ($dnsTest) {
            $details.Add("✅ DNS-Auflösung für 'outlook.office365.com' erfolgreich.")
        } else {
            $details.Add("❌ FEHLER: DNS-Auflösung für 'outlook.office365.com' fehlgeschlagen. Prüfen Sie Ihre DNS-Einstellungen.")
            $overallStatus = "Fehler"
        }

        # 3. Port-Konnektivität prüfen (HTTPS/443)
        Write-Log "Prüfe Port-Konnektivität (TCP 443)..." -Type "Debug"
        $portTest = Test-NetConnection -ComputerName "outlook.office365.com" -Port 443 -InformationLevel Quiet -ErrorAction SilentlyContinue
        if ($portTest) {
            $details.Add("✅ Verbindung zu 'outlook.office365.com' auf Port 443 (HTTPS) erfolgreich.")
        } else {
            $details.Add("❌ FEHLER: Verbindung zu 'outlook.office365.com' auf Port 443 (HTTPS) fehlgeschlagen. Prüfen Sie Ihre Firewall-Regeln.")
            $overallStatus = "Fehler"
        }

        # Wenn Netzwerk-Checks fehlschlagen, hier abbrechen und Ergebnis zurückgeben
        if ($overallStatus -eq "Fehler") {
            $details.Add("➡️ Aufgrund von Netzwerkproblemen wurde der Test der PowerShell-Befehle übersprungen.")
            return [PSCustomObject]@{
                CheckName = "Exchange Online Verbindung"
                Status    = "Fehler"
                Details   = $details -join "`n"
            }
        }

        # 4. Test der PowerShell-Verbindung mit einem EXO-Befehl
        Write-Log "Prüfe PowerShell-Befehlsausführung (Get-OrganizationConfig)..." -Type "Debug"
        $testResult = Get-OrganizationConfig -ErrorAction SilentlyContinue
        
        if ($null -ne $testResult) {
            $details.Add("✅ PowerShell-Befehl (Get-OrganizationConfig) erfolgreich ausgeführt.")
        } else {
            $details.Add("❌ FEHLER: PowerShell-Befehl konnte nicht ausgeführt werden. Mögliche Ursachen: Berechtigungen, abgelaufene Sitzung oder ein temporäres Microsoft-Dienstproblem.")
            $details.Add("➡️ Service Health Dashboard: https://admin.microsoft.com/Adminportal/Home#/servicehealth")
            $overallStatus = "Fehler"
        }

        return [PSCustomObject]@{
            CheckName = "Exchange Online Verbindung"
            Status    = $overallStatus
            Details   = $details -join "`n"
        }
    }
    catch {
        return [PSCustomObject]@{
            CheckName = "Exchange Online Verbindung"
            Status    = "Fehler"
            Details   = "Unerwarteter Fehler bei der Verbindungsprüfung: $($_.Exception.Message)"
        }
    }
}

function Test-MailboxService {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Teste Mailbox-Service..." -Type "Debug"
        $details = [System.Collections.Generic.List[string]]::new()
        $overallStatus = "OK"

        # 1. Test: Abrufen von Postfach-Objekten (Directory-Teil)
        Write-Log "Prüfe Abruf von Postfach-Objekten (Get-Mailbox)..." -Type "Debug"
        $testMailboxes = Get-Mailbox -ResultSize 5 -ErrorAction SilentlyContinue
        
        if ($null -ne $testMailboxes -and $testMailboxes.Count -gt 0) {
            $details.Add("✅ Postfach-Objekte können abgerufen werden (Get-Mailbox erfolgreich).")
            
            # 2. Test: Abrufen von Postfach-Statistiken (Store-Teil)
            $firstMailbox = $testMailboxes[0]
            Write-Log "Prüfe Abruf von Postfach-Statistiken für '$($firstMailbox.PrimarySmtpAddress)'..." -Type "Debug"
            $stats = Get-MailboxStatistics -Identity $firstMailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue
            
            if ($null -ne $stats) {
                $details.Add("✅ Postfach-Statistiken können abgerufen werden (Get-MailboxStatistics erfolgreich).")
                $details.Add("➡️ Der Postfach-Speicher (Store) scheint erreichbar zu sein.")
            } else {
                $details.Add("⚠️ WARNUNG: Postfach-Statistiken konnten nicht abgerufen werden. Der Postfach-Speicher ist möglicherweise langsam oder nicht erreichbar.")
                if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
            }
        } else {
            $details.Add("❌ FEHLER: Postfach-Objekte konnten nicht abgerufen werden. Der Verzeichnisdienst für Postfächer antwortet nicht.")
            $overallStatus = "Fehler"
        }

        return [PSCustomObject]@{
            CheckName = "Mailbox-Service"
            Status    = $overallStatus
            Details   = $details -join "`n"
        }
    }
    catch {
        return [PSCustomObject]@{
            CheckName = "Mailbox-Service"
            Status    = "Fehler"
            Details   = "Unerwarteter Fehler beim Test des Mailbox-Service: $($_.Exception.Message)"
        }
    }
}

function Test-TransportService {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Teste Transport-Service..." -Type "Debug"
        $details = [System.Collections.Generic.List[string]]::new()
        $overallStatus = "OK"

        # 1. Test: Abrufen der Transport-Konfiguration
        Write-Log "Prüfe Abruf der Transport-Konfiguration (Get-TransportConfig)..." -Type "Debug"
        $transportConfig = Get-TransportConfig -ErrorAction SilentlyContinue
        if ($null -ne $transportConfig) {
            $details.Add("✅ Transport-Konfiguration kann abgerufen werden.")
        } else {
            $details.Add("❌ FEHLER: Transport-Konfiguration konnte nicht abgerufen werden. Dies deutet auf ein grundlegendes Problem hin.")
            $overallStatus = "Fehler"
        }

        # 2. Test: Abrufen von Transport-Regeln
        Write-Log "Prüfe Abruf von Transport-Regeln (Get-TransportRule)..." -Type "Debug"
        $transportRules = Get-TransportRule -ResultSize 1 -ErrorAction SilentlyContinue
        if ($null -ne $transportRules) {
            $details.Add("✅ Transport-Regeln können abgerufen werden.")
        } else {
            # Wenn keine Regeln vorhanden sind, ist das kein Fehler, aber das Cmdlet sollte nicht fehlschlagen.
            if ($null -eq $Error[0] -or $Error[0].CategoryInfo.Category -ne 'OperationStopped') {
                 $details.Add("✅ Transport-Regeln-Dienst ist erreichbar (keine Regeln gefunden).")
            } else {
                $details.Add("⚠️ WARNUNG: Transport-Regeln konnten nicht abgerufen werden. Der Regel-Dienst ist möglicherweise beeinträchtigt.")
                if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
            }
            $Error.Clear()
        }

        # 3. Test: Nachrichtenverfolgungs-Dienst (Message Trace)
        Write-Log "Prüfe Nachrichtenverfolgungs-Dienst (Get-MessageTrace)..." -Type "Debug"
        # Wir führen eine sehr kurze Abfrage durch, nur um die Dienstverfügbarkeit zu testen.
        $traceTest = Get-MessageTrace -StartDate (Get-Date).AddMinutes(-1) -EndDate (Get-Date) -PageSize 1 -ErrorAction SilentlyContinue
        if ($null -ne $traceTest -or ($null -eq $Error[0])) {
            $details.Add("✅ Nachrichtenverfolgungs-Dienst ist erreichbar.")
        } else {
            $details.Add("❌ FEHLER: Der Nachrichtenverfolgungs-Dienst antwortet nicht. Die Analyse des E-Mail-Flusses ist beeinträchtigt.")
            $overallStatus = "Fehler"
        }
        $Error.Clear()

        return [PSCustomObject]@{
            CheckName = "Transport-Service"
            Status    = $overallStatus
            Details   = $details -join "`n"
        }
    }
    catch {
        return [PSCustomObject]@{
            CheckName = "Transport-Service"
            Status    = "Fehler"
            Details   = "Unerwarteter Fehler beim Test des Transport-Service: $($_.Exception.Message)"
        }
    }
}

function Test-ComplianceService {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Teste Compliance-Service..." -Type "Debug"
        $details = [System.Collections.Generic.List[string]]::new()
        $overallStatus = "OK"

        # 1. Test: Audit-Status der Organisation prüfen
        Write-Log "Prüfe Audit-Status der Organisation (Get-OrganizationConfig)..." -Type "Debug"
        $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
        if ($null -ne $orgConfig) {
            if ($orgConfig.AuditDisabled -eq $false) {
                $details.Add("✅ Audit-Protokollierung ist auf Organisationsebene aktiviert.")
            } else {
                $details.Add("⚠️ WARNUNG: Die Audit-Protokollierung ist auf Organisationsebene deaktiviert. Dies ist eine Sicherheitslücke.")
                if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
            }
        } else {
            $details.Add("❌ FEHLER: Die Organisationskonfiguration konnte nicht abgerufen werden. Grundlegende Compliance-Prüfungen sind nicht möglich.")
            $overallStatus = "Fehler"
        }

        # 2. Test: Abrufen von Aufbewahrungsrichtlinien (Retention Policies)
        # Dieses Cmdlet ist Teil des Security & Compliance PowerShell und erfordert entsprechende Berechtigungen.
        Write-Log "Prüfe Abruf von Aufbewahrungsrichtlinien (Get-RetentionCompliancePolicy)..." -Type "Debug"
        try {
            # Wir rufen nur eine Richtlinie ab, um die Dienstverfügbarkeit zu testen.
            $retentionPolicies = Get-OwnRetentionCompliancePolicy -ErrorAction Stop
            $details.Add("✅ Dienst für Aufbewahrungsrichtlinien ist erreichbar.")
        }
        catch {
            $errorMsg = Get-FormattedError -ErrorRecord $_
            $details.Add("⚠️ WARNUNG: Der Dienst für Aufbewahrungsrichtlinien antwortet nicht oder es fehlen Berechtigungen. Details: $errorMsg")
            if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
        }

        # 3. Test: Abrufen von DLP-Richtlinien (Data Loss Prevention)
        # Dieses Cmdlet ist ebenfalls Teil des Security & Compliance PowerShell.
        Write-Log "Prüfe Abruf von DLP-Richtlinien (Get-DlpCompliancePolicy)..." -Type "Debug"
        try {
            $dlpPolicies = Get-OwnDlpCompliancePolicy -ErrorAction Stop
            $details.Add("✅ Dienst für DLP-Richtlinien ist erreichbar.")
        }
        catch {
            $errorMsg = Get-FormattedError -ErrorRecord $_
            $details.Add("⚠️ WARNUNG: Der Dienst für DLP-Richtlinien antwortet nicht oder es fehlen Berechtigungen. Details: $errorMsg")
            if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
        }

        # Wenn grundlegende Checks fehlschlagen, hier abbrechen
        if ($overallStatus -eq "Fehler") {
            return [PSCustomObject]@{
                CheckName = "Compliance-Service"
                Status    = "Fehler"
                Details   = $details -join "`n"
            }
        }

        return [PSCustomObject]@{
            CheckName = "Compliance-Service"
            Status    = $overallStatus
            Details   = $details -join "`n"
        }
    }
    catch {
        return [PSCustomObject]@{
            CheckName = "Compliance-Service"
            Status    = "Fehler"
            Details   = "Unerwarteter Fehler beim Test des Compliance-Service: $($_.Exception.Message)"
        }
    }
}

function Test-SecurityServices {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Teste Security-Services..." -Type "Debug"
        $details = [System.Collections.Generic.List[string]]::new()
        $overallStatus = "OK"

        # 1. Anti-Spam-Richtlinien (Eingehend)
        Write-Log "Prüfe Anti-Spam-Richtlinien (eingehend)..." -Type "Debug"
        $spamPolicies = Get-HostedContentFilterPolicy -ErrorAction SilentlyContinue
        if ($null -ne $spamPolicies) {
            $details.Add("✅ Anti-Spam-Richtlinien (eingehend) sind konfiguriert.")
        } else {
            $details.Add("❌ FEHLER: Anti-Spam-Dienst (eingehend) nicht erreichbar oder keine Richtlinien gefunden.")
            $overallStatus = "Fehler"
        }

        # 2. Anti-Malware-Richtlinien
        Write-Log "Prüfe Anti-Malware-Richtlinien..." -Type "Debug"
        $malwarePolicies = Get-MalwareFilterPolicy -ErrorAction SilentlyContinue
        if ($null -ne $malwarePolicies) {
            $details.Add("✅ Anti-Malware-Richtlinien sind konfiguriert.")
        } else {
            $details.Add("❌ FEHLER: Anti-Malware-Dienst nicht erreichbar oder keine Richtlinien gefunden.")
            $overallStatus = "Fehler"
        }

        # 3. DKIM-Konfiguration
        Write-Log "Prüfe DKIM-Konfiguration..." -Type "Debug"
        $dkimConfigs = Get-DkimSigningConfig -ErrorAction SilentlyContinue
        if ($null -ne $dkimConfigs) {
            if ($dkimConfigs | Where-Object { $_.Enabled }) {
                $details.Add("✅ DKIM ist für mindestens eine Domain aktiviert.")
            } else {
                $details.Add("⚠️ WARNUNG: DKIM ist für keine Domain aktiviert. Dies wird für die E-Mail-Authentizität dringend empfohlen.")
                if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
            }
        } else {
            $details.Add("⚠️ WARNUNG: DKIM-Konfiguration konnte nicht abgerufen werden.")
            if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
        }

        # 4. Defender for Office 365-Richtlinien (optional, da lizenzabhängig)
        Write-Log "Prüfe Safe Links-Richtlinien (Defender for O365)..." -Type "Debug"
        $safeLinksPolicies = Get-SafeLinksPolicy -ErrorAction SilentlyContinue
        if ($null -ne $safeLinksPolicies) {
            $details.Add("ℹ️ Info: Safe Links-Richtlinien (Defender for O365) sind vorhanden.")
        } else {
            $details.Add("ℹ️ Info: Keine Safe Links-Richtlinien gefunden (möglicherweise nicht lizenziert).")
        }

        Write-Log "Prüfe Safe Attachment-Richtlinien (Defender for O365)..." -Type "Debug"
        $safeAttachmentPolicies = Get-SafeAttachmentPolicy -ErrorAction SilentlyContinue
        if ($null -ne $safeAttachmentPolicies) {
            $details.Add("ℹ️ Info: Safe Attachment-Richtlinien (Defender for O365) sind vorhanden.")
        } else {
            $details.Add("ℹ️ Info: Keine Safe Attachment-Richtlinien gefunden (möglicherweise nicht lizenziert).")
        }

        return [PSCustomObject]@{
            CheckName = "Security-Services"
            Status    = $overallStatus
            Details   = $details -join "`n"
        }
    }
    catch {
        return [PSCustomObject]@{
            CheckName = "Security-Services"
            Status    = "Fehler"
            Details   = "Unerwarteter Fehler beim Test der Security-Services: $($_.Exception.Message)"
        }
    }
}

function Test-LicenseStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Teste Lizenz-Status..." -Type "Debug"
        $details = [System.Collections.Generic.List[string]]::new()
        $overallStatus = "OK"

        # Abrufen der Organisationskonfiguration, um den Servicetyp zu ermitteln
        $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
        
        if ($null -ne $orgConfig) {
            $details.Add("✅ Organisationskonfiguration erfolgreich abgerufen.")
            $details.Add("➡️ Organisationstyp: $($orgConfig.OrganizationType)")
            
            # Überprüfen, ob es sich um einen dedizierten oder einen Standard-Mandanten handelt
            if ($orgConfig.IsDehydrated) {
                $details.Add("ℹ️ Info: Dies ist ein 'dehydrierter' Mandant (typisch für Hybrid-Szenarien). Lizenzdetails sind hier oft eingeschränkt.")
            }

            # Überprüfen der Lizenz-SKUs (erfordert MSOnline oder Microsoft.Graph Modul, was hier nicht vorausgesetzt wird)
            # Daher beschränken wir uns auf EXO-interne Checks.
            # Ein guter Indikator ist, ob Postfächer erstellt werden können.
            $testMailbox = Get-Mailbox -ResultSize 1 -ErrorAction SilentlyContinue
            if ($null -ne $testMailbox) {
                $details.Add("✅ Mindestens ein Postfach gefunden. Exchange Online-Lizenzen scheinen zugewiesen zu sein.")
            } else {
                $details.Add("⚠️ WARNUNG: Es konnten keine Postfächer gefunden werden. Überprüfen Sie, ob Lizenzen zugewiesen sind und die Synchronisierung korrekt ist.")
                if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
            }

            # Prüfung auf Premium-Features als Indikator für höhere Lizenzen
            $litigationHoldMailboxes = Get-Mailbox -Filter "LitigationHoldEnabled -eq `$true" -ResultSize 1 -ErrorAction SilentlyContinue
            if ($null -ne $litigationHoldMailboxes) {
                $details.Add("ℹ️ Info: Mindestens ein Postfach mit 'Litigation Hold' gefunden. Dies deutet auf Lizenzen wie E3/E5 oder Add-ons hin.")
            }

            $archiveMailboxes = Get-Mailbox -Filter "ArchiveStatus -ne 'None'" -ResultSize 1 -ErrorAction SilentlyContinue
            if ($null -ne $archiveMailboxes) {
                $details.Add("ℹ️ Info: Mindestens ein Postfach mit aktivem Archiv gefunden. Dies deutet auf Lizenzen wie Exchange Online Plan 2 oder höher hin.")
            }

        }
        else {
            $details.Add("❌ FEHLER: Die Organisationskonfiguration konnte nicht abgerufen werden. Eine grundlegende Lizenzprüfung ist nicht möglich.")
            $overallStatus = "Fehler"
        }

        return [PSCustomObject]@{
            CheckName = "Lizenz-Status"
            Status    = $overallStatus
            Details   = $details -join "`n"
        }
    }
    catch {
        return [PSCustomObject]@{
            CheckName = "Lizenz-Status"
            Status    = "Fehler"
            Details   = "Unerwarteter Fehler beim Test des Lizenz-Status: $($_.Exception.Message)"
        }
    }
}
function Test-TenantConfiguration {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Teste Tenant-Konfiguration..." -Type "Debug"
        $details = [System.Collections.Generic.List[string]]::new()
        $warnings = [System.Collections.Generic.List[string]]::new()
        $overallStatus = "OK"

        # 1. Test: Allgemeine Organisationskonfiguration abrufen
        Write-Log "Prüfe allgemeine Organisationskonfiguration (Get-OrganizationConfig)..." -Type "Debug"
        $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
        if ($null -ne $orgConfig) {
            $details.Add("✅ Allgemeine Organisationskonfiguration erfolgreich abgerufen.")
            
            # Kritische Einstellungen prüfen
            if ($orgConfig.AuditDisabled -eq $true) {
                $warnings.Add("Die Audit-Protokollierung ist auf Organisationsebene deaktiviert. Dies ist eine Sicherheitslücke.")
                if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
            } else {
                $details.Add("✅ Audit-Protokollierung ist aktiviert.")
            }

            if ($orgConfig.ModernAuthEnabled -eq $false) {
                $warnings.Add("Modern Authentication (OAuth2) ist deaktiviert. Dies ist veraltet und unsicher.")
                if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
            } else {
                $details.Add("✅ Modern Authentication ist aktiviert.")
            }
            
            if ($orgConfig.SmtpClientAuthenticationDisabled -eq $false) {
                $warnings.Add("SMTP AUTH ist mandantenweit aktiviert. Dies stellt ein erhebliches Sicherheitsrisiko dar und sollte deaktiviert werden, es sei denn, es wird explizit benötigt.")
                if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
            } else {
                $details.Add("✅ SMTP AUTH ist mandantenweit deaktiviert (Best Practice).")
            }

            $details.Add("➡️ Öffentliche Ordner: $($orgConfig.PublicFoldersEnabled)")

        } else {
            $warnings.Add("Die allgemeine Organisationskonfiguration konnte nicht abgerufen werden.")
            $overallStatus = "Fehler"
        }

        # 2. Test: Standard-Remote-Domain-Einstellungen prüfen
        Write-Log "Prüfe Standard-Remote-Domain-Einstellungen (Get-RemoteDomain)..." -Type "Debug"
        $remoteDomain = Get-RemoteDomain -Identity "*" -ErrorAction SilentlyContinue
        if ($null -ne $remoteDomain) {
            $details.Add("✅ Standard-Remote-Domain-Einstellungen ('*') erfolgreich abgerufen.")
            if ($remoteDomain.AutoForwardEnabled -eq $true) {
                $warnings.Add("Automatische externe Weiterleitungen sind für die gesamte Organisation erlaubt. Dies stellt ein Sicherheitsrisiko dar.")
                if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
            } else {
                $details.Add("✅ Automatische externe Weiterleitungen sind standardmäßig blockiert (Best Practice).")
            }
            if ($remoteDomain.AutoReplyEnabled -eq $false) {
                $details.Add("ℹ️ Info: Automatische Abwesenheitsnotizen an externe Empfänger sind standardmäßig deaktiviert.")
            }
        } else {
            $warnings.Add("Die Standard-Remote-Domain-Einstellungen konnten nicht abgerufen werden.")
            if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
        }

        # 3. Test: Akzeptierte Domains prüfen
        Write-Log "Prüfe akzeptierte Domains (Get-AcceptedDomain)..." -Type "Debug"
        $acceptedDomains = Get-AcceptedDomain -ErrorAction SilentlyContinue
        if ($null -ne $acceptedDomains) {
            $details.Add("✅ Akzeptierte Domains erfolgreich abgerufen ($($acceptedDomains.Count) gefunden).")
            if (-not ($acceptedDomains | Where-Object { $_.IsDefault })) {
                $warnings.Add("Es ist keine Standard-Domain ('IsDefault' = $true) konfiguriert. Dies kann zu Problemen bei der E-Mail-Zustellung führen.")
                if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
            } else {
                $details.Add("✅ Eine Standard-Domain ist konfiguriert.")
            }
        } else {
            $warnings.Add("Die akzeptierten Domains konnten nicht abgerufen werden.")
            if ($overallStatus -ne "Fehler") { $overallStatus = "Warnung" }
        }

        # Ergebnis zusammenstellen
        $finalDetails = ""
        if ($warnings.Count -gt 0) {
            $finalDetails += "⚠️ Warnungen:`n"
            $warnings.ForEach({ $finalDetails += "- $_\n" })
            $finalDetails += "`n"
        }
        $finalDetails += "ℹ️ Details:`n"
        $details.ForEach({ $finalDetails += "- $_\n" })

        return [PSCustomObject]@{
            CheckName = "Tenant-Konfiguration"
            Status    = $overallStatus
            Details   = $finalDetails.Trim()
        }
    }
    catch {
        return [PSCustomObject]@{
            CheckName = "Tenant-Konfiguration"
            Status    = "Fehler"
            Details   = "Tenant-Konfiguration nicht abrufbar: $($_.Exception.Message)"
        }
    }
}

# -------------------------------------------------
# Abschnitt: MailFlow-Funktionen
# -------------------------------------------------

function New-MailFlowRuleAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RuleName,
        
        [Parameter(Mandatory = $true)]
        [string]$ConditionType,
        
        [Parameter(Mandatory = $false)]
        [string]$ConditionValue,
        
        [Parameter(Mandatory = $true)]
        [string]$ActionType,
        
        [Parameter(Mandatory = $false)]
        [string]$ActionValue,
        
        [Parameter(Mandatory = $false)]
        [string]$Mode = "Test"
    )
    
    try {
        Write-Log "Erstelle neue Transport Rule: $RuleName" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Parameter für die Regel aufbauen
        $ruleParams = @{
            Name = $RuleName
            Mode = $Mode
            ErrorAction = "Stop"
        }
        
        # Bedingungen setzen
        switch ($ConditionType) {
            "FromAddressContainsWords" {
                if (-not [string]::IsNullOrWhiteSpace($ConditionValue)) {
                    $ruleParams.Add("FromAddressContainsWords", $ConditionValue.Split(',').Trim())
                }
            }
            "SubjectContainsWords" {
                if (-not [string]::IsNullOrWhiteSpace($ConditionValue)) {
                    $ruleParams.Add("SubjectContainsWords", $ConditionValue.Split(',').Trim())
                }
            }
            "RecipientDomainIs" {
                if (-not [string]::IsNullOrWhiteSpace($ConditionValue)) {
                    $ruleParams.Add("RecipientDomainIs", $ConditionValue.Split(',').Trim())
                }
            }
            "SenderDomainIs" {
                if (-not [string]::IsNullOrWhiteSpace($ConditionValue)) {
                    $ruleParams.Add("SenderDomainIs", $ConditionValue.Split(',').Trim())
                }
            }
            "AttachmentNameMatchesPatterns" {
                if (-not [string]::IsNullOrWhiteSpace($ConditionValue)) {
                    $ruleParams.Add("AttachmentNameMatchesPatterns", $ConditionValue.Split(',').Trim())
                }
            }
            "MessageSizeOver" {
                if (-not [string]::IsNullOrWhiteSpace($ConditionValue)) {
                    $ruleParams.Add("MessageSizeOver", "$ConditionValue KB")
                }
            }
        }
        
        # Aktionen setzen
        switch ($ActionType) {
            "RedirectMessageTo" {
                if (-not [string]::IsNullOrWhiteSpace($ActionValue)) {
                    $ruleParams.Add("RedirectMessageTo", $ActionValue.Split(',').Trim())
                }
            }
            "BlindCopyTo" {
                if (-not [string]::IsNullOrWhiteSpace($ActionValue)) {
                    $ruleParams.Add("BlindCopyTo", $ActionValue.Split(',').Trim())
                }
            }
            "RejectMessageReasonText" {
                if (-not [string]::IsNullOrWhiteSpace($ActionValue)) {
                    $ruleParams.Add("RejectMessageReasonText", $ActionValue)
                }
            }
            "DeleteMessage" {
                $ruleParams.Add("DeleteMessage", $true)
            }
            "ModerateMessageByUser" {
                if (-not [string]::IsNullOrWhiteSpace($ActionValue)) {
                    $ruleParams.Add("ModerateMessageByUser", $ActionValue.Split(',').Trim())
                }
            }
            "SetHeaderName" {
                if (-not [string]::IsNullOrWhiteSpace($ActionValue)) {
                    $headerParts = $ActionValue.Split(':', 2)
                    if ($headerParts.Count -eq 2) {
                        $ruleParams.Add("SetHeaderName", $headerParts[0].Trim())
                        $ruleParams.Add("SetHeaderValue", $headerParts[1].Trim())
                    }
                }
            }
            "ApplyClassification" {
                if (-not [string]::IsNullOrWhiteSpace($ActionValue)) {
                    $ruleParams.Add("ApplyClassification", $ActionValue)
                }
            }
            "SetSCL" {
                if (-not [string]::IsNullOrWhiteSpace($ActionValue)) {
                    $ruleParams.Add("SetSCL", [int]$ActionValue)
                }
            }
            "Quarantine" {
                $ruleParams.Add("Quarantine", $true)
            }
        }
        
        # Transport Rule erstellen
        $newRule = New-TransportRule @ruleParams
        
        Write-Log "Transport Rule '$RuleName' erfolgreich erstellt" -Type "Success"
        Log-Action "Transport Rule erstellt: $RuleName (Modus: $Mode)"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Erstellen der Transport Rule: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Erstellen der Transport Rule $RuleName`: $errorMsg"
        return $false
    }
}

function Test-MailFlowRuleAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RuleName,
        
        [Parameter(Mandatory = $true)]
        [string]$ConditionType,
        
        [Parameter(Mandatory = $false)]
        [string]$ConditionValue
    )
    
    try {
        Write-Log "Teste Transport Rule Konfiguration: $RuleName" -Type "Info"
        
        $testResults = @()
        $testResults += "=== Transport Rule Test Ergebnisse ==="
        $testResults += "Regel Name: $RuleName"
        $testResults += "Bedingung: $ConditionType"
        $testResults += "Wert: $ConditionValue"
        $testResults += ""
        
        # Validierung der Bedingungen
        switch ($ConditionType) {
            "FromAddressContainsWords" {
                if ([string]::IsNullOrWhiteSpace($ConditionValue)) {
                    $testResults += "❌ FEHLER: Keine Wörter für FromAddressContainsWords angegeben"
                } else {
                    $words = $ConditionValue.Split(',').Trim()
                    $testResults += "✅ OK: $($words.Count) Wort(e) für Absender-Filter konfiguriert"
                    foreach ($word in $words) {
                        $testResults += "   - '$word'"
                    }
                }
            }
            "SubjectContainsWords" {
                if ([string]::IsNullOrWhiteSpace($ConditionValue)) {
                    $testResults += "❌ FEHLER: Keine Wörter für SubjectContainsWords angegeben"
                } else {
                    $words = $ConditionValue.Split(',').Trim()
                    $testResults += "✅ OK: $($words.Count) Wort(e) für Betreff-Filter konfiguriert"
                    foreach ($word in $words) {
                        $testResults += "   - '$word'"
                    }
                }
            }
            "RecipientDomainIs" {
                if ([string]::IsNullOrWhiteSpace($ConditionValue)) {
                    $testResults += "❌ FEHLER: Keine Domain für RecipientDomainIs angegeben"
                } else {
                    $domains = $ConditionValue.Split(',').Trim()
                    $testResults += "✅ OK: $($domains.Count) Domain(s) für Empfänger-Filter konfiguriert"
                    foreach ($domain in $domains) {
                        if ($domain -match '^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                            $testResults += "   - ✅ '$domain' (gültiges Format)"
                        } else {
                            $testResults += "   - ❌ '$domain' (ungültiges Domain-Format)"
                        }
                    }
                }
            }
            "MessageSizeOver" {
                if ([string]::IsNullOrWhiteSpace($ConditionValue)) {
                    $testResults += "❌ FEHLER: Keine Größe für MessageSizeOver angegeben"
                } else {
                    if ($ConditionValue -match '^\d+$') {
                        $sizeKB = [int]$ConditionValue
                        $testResults += "✅ OK: Nachrichten über $sizeKB KB werden gefiltert"
                        $testResults += "   (Das entspricht ca. $([math]::Round($sizeKB/1024, 2)) MB)"
                    } else {
                        $testResults += "❌ FEHLER: '$ConditionValue' ist keine gültige Größenangabe (nur Zahlen erlaubt)"
                    }
                }
            }
            default {
                $testResults += "⚠️ WARNUNG: Unbekannter Bedingungstyp '$ConditionType'"
            }
        }
        
        $testResults += ""
        $testResults += "=== Empfehlungen ==="
        $testResults += "• Testen Sie die Regel zunächst im 'Test'-Modus"
        $testResults += "• Überwachen Sie die Regel-Performance nach der Aktivierung"
        $testResults += "• Dokumentieren Sie Änderungen für das Audit"
        
        $result = $testResults -join "`n"
        Write-Log "Transport Rule Test abgeschlossen für: $RuleName" -Type "Success"
        Log-Action "Transport Rule Test durchgeführt für: $RuleName"
        
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Testen der Transport Rule: $errorMsg" -Type "Error"
        return "Fehler beim Testen der Regel: $errorMsg"
    }
}

function Get-MailFlowRulesAction {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Rufe Transport Rules ab..." -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Status aktualisieren
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Lade Transport Rules..."
        }
        
        # Transport Rules abrufen
        $transportRules = Get-TransportRule | Select-Object Name, State, Mode, Priority, Description, 
            FromAddressContainsWords, SubjectContainsWords, RecipientDomainIs, SenderDomainIs,
            RedirectMessageTo, BlindCopyTo, RejectMessageReasonText, DeleteMessage
        
        # Daten für die Anzeige aufbereiten
        $rulesForGrid = @()
        foreach ($rule in $transportRules) {
            $ruleObj = [PSCustomObject]@{
                Name = $rule.Name
                State = $rule.State
                Mode = $rule.Mode
                Priority = $rule.Priority
                Description = $rule.Description
                Conditions = @()
                Actions = @()
            }
            
            # Bedingungen sammeln
            if ($rule.FromAddressContainsWords) {
                $ruleObj.Conditions += "Absender enthält: $($rule.FromAddressContainsWords -join ', ')"
            }
            if ($rule.SubjectContainsWords) {
                $ruleObj.Conditions += "Betreff enthält: $($rule.SubjectContainsWords -join ', ')"
            }
            if ($rule.RecipientDomainIs) {
                $ruleObj.Conditions += "Empfänger-Domain: $($rule.RecipientDomainIs -join ', ')"
            }
            if ($rule.SenderDomainIs) {
                $ruleObj.Conditions += "Absender-Domain: $($rule.SenderDomainIs -join ', ')"
            }
            
            # Aktionen sammeln
            if ($rule.RedirectMessageTo) {
                $ruleObj.Actions += "Weiterleiten an: $($rule.RedirectMessageTo -join ', ')"
            }
            if ($rule.BlindCopyTo) {
                $ruleObj.Actions += "BCC an: $($rule.BlindCopyTo -join ', ')"
            }
            if ($rule.RejectMessageReasonText) {
                $ruleObj.Actions += "Ablehnen: $($rule.RejectMessageReasonText)"
            }
            if ($rule.DeleteMessage) {
                $ruleObj.Actions += "Nachricht löschen"
            }
            
            # Arrays zu Strings konvertieren für Anzeige
            $ruleObj | Add-Member -NotePropertyName "ConditionsDisplay" -NotePropertyValue ($ruleObj.Conditions -join '; ')
            $ruleObj | Add-Member -NotePropertyName "ActionsDisplay" -NotePropertyValue ($ruleObj.Actions -join '; ')
            
            $rulesForGrid += $ruleObj
        }
        
        # Daten in das DataGrid laden
        if ($null -ne $script:dgMailFlowRules) {
            $script:dgMailFlowRules.Dispatcher.Invoke([action]{
                $script:dgMailFlowRules.ItemsSource = $rulesForGrid
            }, "Normal")
        }
        
        Write-Log "Transport Rules erfolgreich geladen: $($rulesForGrid.Count) Regeln" -Type "Success"
        
        # Status aktualisieren
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Transport Rules geladen: $($rulesForGrid.Count) Regeln gefunden."
        }
        
        Log-Action "Transport Rules abgerufen: $($rulesForGrid.Count) Regeln"
        return $rulesForGrid
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Abrufen der Transport Rules: $errorMsg" -Type "Error"
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Fehler beim Laden der Transport Rules: $errorMsg"
        }
        
        Log-Action "Fehler beim Abrufen der Transport Rules: $errorMsg"
        return @()
    }
}

function Set-MailFlowRuleStateAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RuleName,
        
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )
    
    try {
        $state = if ($Enabled) { "Enabled" } else { "Disabled" }
        Write-Log "Setze Transport Rule '$RuleName' auf Status: $state" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Transport Rule Status ändern
        Set-TransportRule -Identity $RuleName -State $state -ErrorAction Stop
        
        $statusMsg = if ($Enabled) { "aktiviert" } else { "deaktiviert" }
        Write-Log "Transport Rule '$RuleName' erfolgreich $statusMsg" -Type "Success"
        Log-Action "Transport Rule '$RuleName' wurde $statusMsg"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        $statusMsg = if ($Enabled) { "Aktivieren" } else { "Deaktivieren" }
        Write-Log "Fehler beim $statusMsg der Transport Rule '$RuleName': $errorMsg" -Type "Error"
        Log-Action "Fehler beim $statusMsg der Transport Rule '$RuleName': $errorMsg"
        return $false
    }
}

function Remove-MailFlowRuleAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RuleName
    )
    
    try {
        Write-Log "Lösche Transport Rule: $RuleName" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Transport Rule löschen
        Remove-TransportRule -Identity $RuleName -Confirm:$false -ErrorAction Stop
        
        Write-Log "Transport Rule '$RuleName' erfolgreich gelöscht" -Type "Success"
        Log-Action "Transport Rule '$RuleName' wurde gelöscht"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Löschen der Transport Rule '$RuleName': $errorMsg" -Type "Error"
        Log-Action "Fehler beim Löschen der Transport Rule '$RuleName': $errorMsg"
        return $false
    }
}

function Export-MailFlowRulesAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        Write-Log "Exportiere Transport Rules nach: $FilePath" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Transport Rules abrufen
        $transportRules = Get-TransportRule | Select-Object Name, State, Mode, Priority, Description,
            FromAddressContainsWords, SubjectContainsWords, RecipientDomainIs, SenderDomainIs,
            RedirectMessageTo, BlindCopyTo, RejectMessageReasonText, DeleteMessage,
            SetHeaderName, SetHeaderValue, ApplyClassification, SetSCL, Quarantine
        
        # Bestimme Exportformat basierend auf Dateiendung
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        
        switch ($extension) {
            ".csv" {
                $transportRules | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            }
            ".xml" {
                $transportRules | Export-Clixml -Path $FilePath -Encoding UTF8
            }
            default {
                # Fallback zu CSV
                $transportRules | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            }
        }
        
        Write-Log "Transport Rules erfolgreich exportiert: $($transportRules.Count) Regeln" -Type "Success"
        Log-Action "Transport Rules exportiert nach: $FilePath ($($transportRules.Count) Regeln)"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Exportieren der Transport Rules: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Exportieren der Transport Rules: $errorMsg"
        return $false
    }
}

function Import-MailFlowRulesAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        Write-Log "Importiere Transport Rules von: $FilePath" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        if (-not (Test-Path -Path $FilePath)) {
            throw "Datei nicht gefunden: $FilePath"
        }
        
        # Bestimme Importformat basierend auf Dateiendung
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        $importedRules = $null
        
        switch ($extension) {
            ".csv" {
                $importedRules = Import-Csv -Path $FilePath -Delimiter ";" -Encoding UTF8
            }
            ".xml" {
                $importedRules = Import-Clixml -Path $FilePath
            }
            default {
                # Fallback zu CSV
                $importedRules = Import-Csv -Path $FilePath -Delimiter ";" -Encoding UTF8
            }
        }
        
        if ($null -eq $importedRules -or $importedRules.Count -eq 0) {
            throw "Keine gültigen Transport Rules in der Datei gefunden."
        }
        
        $successCount = 0
        $errorCount = 0
        $errors = @()
        
        foreach ($rule in $importedRules) {
            try {
                # Überprüfe, ob Regel bereits existiert
                $existingRule = Get-TransportRule -Identity $rule.Name -ErrorAction SilentlyContinue
                if ($existingRule) {
                    Write-Log "Transport Rule '$($rule.Name)' existiert bereits - überspringe" -Type "Warning"
                    continue
                }
                
                # Parameter für Import vorbereiten
                $ruleParams = @{
                    Name = $rule.Name
                    State = $rule.State
                    Mode = $rule.Mode
                    Priority = if ($rule.Priority) { [int]$rule.Priority } else { 0 }
                    ErrorAction = "Stop"
                }
                
                if ($rule.Description) { $ruleParams.Add("Description", $rule.Description) }
                if ($rule.FromAddressContainsWords) { $ruleParams.Add("FromAddressContainsWords", $rule.FromAddressContainsWords.Split(',').Trim()) }
                if ($rule.SubjectContainsWords) { $ruleParams.Add("SubjectContainsWords", $rule.SubjectContainsWords.Split(',').Trim()) }
                if ($rule.RecipientDomainIs) { $ruleParams.Add("RecipientDomainIs", $rule.RecipientDomainIs.Split(',').Trim()) }
                if ($rule.SenderDomainIs) { $ruleParams.Add("SenderDomainIs", $rule.SenderDomainIs.Split(',').Trim()) }
                if ($rule.RedirectMessageTo) { $ruleParams.Add("RedirectMessageTo", $rule.RedirectMessageTo.Split(',').Trim()) }
                if ($rule.BlindCopyTo) { $ruleParams.Add("BlindCopyTo", $rule.BlindCopyTo.Split(',').Trim()) }
                if ($rule.RejectMessageReasonText) { $ruleParams.Add("RejectMessageReasonText", $rule.RejectMessageReasonText) }
                if ($rule.DeleteMessage -eq "True") { $ruleParams.Add("DeleteMessage", $true) }
                if ($rule.SetHeaderName) { $ruleParams.Add("SetHeaderName", $rule.SetHeaderName) }
                if ($rule.SetHeaderValue) { $ruleParams.Add("SetHeaderValue", $rule.SetHeaderValue) }
                if ($rule.ApplyClassification) { $ruleParams.Add("ApplyClassification", $rule.ApplyClassification) }
                if ($rule.SetSCL) { $ruleParams.Add("SetSCL", [int]$rule.SetSCL) }
                if ($rule.Quarantine -eq "True") { $ruleParams.Add("Quarantine", $true) }
                
                # Transport Rule erstellen
                New-TransportRule @ruleParams
                $successCount++
                
                Write-Log "Transport Rule '$($rule.Name)' erfolgreich importiert" -Type "Success"
            }
            catch {
                $errorCount++
                $errorMsg = $_.Exception.Message
                $errors += "Regel '$($rule.Name)': $errorMsg"
                Write-Log "Fehler beim Importieren der Transport Rule '$($rule.Name)': $errorMsg" -Type "Error"
            }
        }
        
        $resultMsg = "Import abgeschlossen: $successCount erfolgreich, $errorCount Fehler"
        Write-Log $resultMsg -Type "Info"
        
        if ($errors.Count -gt 0) {
            $errorSummary = "Fehler beim Import:`n" + ($errors -join "`n")
            Log-Action "Transport Rules Import: $resultMsg. Fehler: $errorSummary"
        } else {
            Log-Action "Transport Rules erfolgreich importiert: $successCount Regeln"
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Importieren der Transport Rules: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Importieren der Transport Rules: $errorMsg"
        return $false
    }
}

# Fehlende Hilfsfunktionen für die ComboBox-Initialisierung
function Initialize-MailFlowRuleConditionsComboBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.ComboBox]$ComboBox
    )
    
    try {
        $ComboBox.Items.Clear()
        
        # Verfügbare Bedingungen definieren
        $conditions = @(
            @{ Display = "Absender-Adresse enthält Wörter"; Value = "FromAddressContainsWords" },
            @{ Display = "Betreff enthält Wörter"; Value = "SubjectContainsWords" },
            @{ Display = "Empfänger-Domain ist"; Value = "RecipientDomainIs" },
            @{ Display = "Absender-Domain ist"; Value = "SenderDomainIs" },
            @{ Display = "Dateianhang-Name entspricht Muster"; Value = "AttachmentNameMatchesPatterns" },
            @{ Display = "Nachrichtengröße größer als (KB)"; Value = "MessageSizeOver" }
        )
        
        foreach ($condition in $conditions) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $condition.Display
            $item.Tag = $condition.Value
            $ComboBox.Items.Add($item)
        }
        
        if ($ComboBox.Items.Count -gt 0) {
            $ComboBox.SelectedIndex = 0
        }
        
        Write-Log "Mail Flow Bedingungen ComboBox initialisiert" -Type "Debug"
    }
    catch {
        Write-Log "Fehler beim Initialisieren der Bedingungen ComboBox: $($_.Exception.Message)" -Type "Error"
    }
}

function Initialize-MailFlowRuleActionsComboBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.ComboBox]$ComboBox
    )
    
    try {
        $ComboBox.Items.Clear()
        
        # Verfügbare Aktionen definieren
        $actions = @(
            @{ Display = "Nachricht weiterleiten an"; Value = "RedirectMessageTo" },
            @{ Display = "Blindkopie (BCC) an"; Value = "BlindCopyTo" },
            @{ Display = "Nachricht ablehnen mit Text"; Value = "RejectMessageReasonText" },
            @{ Display = "Nachricht löschen"; Value = "DeleteMessage" },
            @{ Display = "Moderation durch Benutzer"; Value = "ModerateMessageByUser" },
            @{ Display = "Header setzen (Name:Wert)"; Value = "SetHeaderName" },
            @{ Display = "Klassifizierung anwenden"; Value = "ApplyClassification" },
            @{ Display = "SCL-Wert setzen"; Value = "SetSCL" },
            @{ Display = "In Quarantäne verschieben"; Value = "Quarantine" }
        )
        
        foreach ($action in $actions) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $action.Display
            $item.Tag = $action.Value
            $ComboBox.Items.Add($item)
        }
        
        if ($ComboBox.Items.Count -gt 0) {
            $ComboBox.SelectedIndex = 0
        }
        
        Write-Log "Mail Flow Aktionen ComboBox initialisiert" -Type "Debug"
    }
    catch {
        Write-Log "Fehler beim Initialisieren der Aktionen ComboBox: $($_.Exception.Message)" -Type "Error"
    }
}

function Initialize-MailFlowRuleModeComboBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.ComboBox]$ComboBox
    )
    
    try {
        $ComboBox.Items.Clear()
        
        # Verfügbare Modi definieren
        $modes = @(
            @{ Display = "Test (ohne Aktion ausführen)"; Value = "Test" },
            @{ Display = "Test mit Policy Tips"; Value = "TestWithPolicyTips" },
            @{ Display = "Enforce (Regel aktiv)"; Value = "Enforce" }
        )
        
        foreach ($mode in $modes) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $mode.Display
            $item.Tag = $mode.Value
            $ComboBox.Items.Add($item)
        }
        
        # Standardmäßig "Test with Policy Tips" auswählen
        if ($ComboBox.Items.Count -gt 1) {
            $ComboBox.SelectedIndex = 1
        }
        
        Write-Log "Mail Flow Modi ComboBox initialisiert" -Type "Debug"
    }
    catch {
        Write-Log "Fehler beim Initialisieren der Modi ComboBox: $($_.Exception.Message)" -Type "Error"
    }
}

# -------------------------------------------------
# Abschnitt: Inbox Rules-Funktionen
# -------------------------------------------------

function Get-InboxRulesAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdentity
    )
    
    try {
        Write-Log "Rufe Inbox Rules ab für Benutzer: $UserIdentity" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Status aktualisieren
        if ($null -ne $script:txtStatus) {
            Update-GuiText -TextElement $script:txtStatus -Message "Lade Inbox Rules für $UserIdentity..."
        }
        
        # Inbox Rules abrufen
        $inboxRules = Get-InboxRule -Mailbox $UserIdentity -ErrorAction Stop
        
        # Daten für die Anzeige aufbereiten
        $rulesForGrid = @()
        foreach ($rule in $inboxRules) {
            $ruleObj = [PSCustomObject]@{
                Identity = $rule.Identity
                Name = $rule.Name
                Description = $rule.Description
                Enabled = $rule.Enabled
                Priority = $rule.Priority
                Conditions = @()
                Actions = @()
                ConditionsDisplay = ""
                ActionsDisplay = ""
            }
            
            # Bedingungen sammeln
            if ($rule.From) {
                $ruleObj.Conditions += "Von: $($rule.From -join ', ')"
            }
            if ($rule.FromAddressContainsWords) {
                $ruleObj.Conditions += "Absender enthält: $($rule.FromAddressContainsWords -join ', ')"
            }
            if ($rule.SubjectContainsWords) {
                $ruleObj.Conditions += "Betreff enthält: $($rule.SubjectContainsWords -join ', ')"
            }
            if ($rule.BodyContainsWords) {
                $ruleObj.Conditions += "Text enthält: $($rule.BodyContainsWords -join ', ')"
            }
            if ($rule.ReceivedAfterDate) {
                $ruleObj.Conditions += "Empfangen nach: $($rule.ReceivedAfterDate)"
            }
            if ($rule.ReceivedBeforeDate) {
                $ruleObj.Conditions += "Empfangen vor: $($rule.ReceivedBeforeDate)"
            }
            if ($rule.WithImportance) {
                $ruleObj.Conditions += "Wichtigkeit: $($rule.WithImportance)"
            }
            if ($rule.HasAttachment) {
                $ruleObj.Conditions += "Hat Anhang: $($rule.HasAttachment)"
            }
            
            # Aktionen sammeln
            if ($rule.MoveToFolder) {
                $ruleObj.Actions += "Verschieben in Ordner: $($rule.MoveToFolder)"
            }
            if ($rule.CopyToFolder) {
                $ruleObj.Actions += "Kopieren in Ordner: $($rule.CopyToFolder)"
            }
            if ($rule.ForwardTo) {
                $ruleObj.Actions += "Weiterleiten an: $($rule.ForwardTo -join ', ')"
            }
            if ($rule.ForwardAsAttachmentTo) {
                $ruleObj.Actions += "Als Anhang weiterleiten an: $($rule.ForwardAsAttachmentTo -join ', ')"
            }
            if ($rule.RedirectTo) {
                $ruleObj.Actions += "Umleiten an: $($rule.RedirectTo -join ', ')"
            }
            if ($rule.MarkAsRead) {
                $ruleObj.Actions += "Als gelesen markieren: $($rule.MarkAsRead)"
            }
            if ($rule.MarkImportance) {
                $ruleObj.Actions += "Wichtigkeit setzen: $($rule.MarkImportance)"
            }
            if ($rule.DeleteMessage) {
                $ruleObj.Actions += "Nachricht löschen: $($rule.DeleteMessage)"
            }
            if ($rule.StopProcessingRules) {
                $ruleObj.Actions += "Weitere Regeln stoppen: $($rule.StopProcessingRules)"
            }
            
            # Arrays zu Strings konvertieren für Anzeige
            $ruleObj.ConditionsDisplay = $ruleObj.Conditions -join '; '
            $ruleObj.ActionsDisplay = $ruleObj.Actions -join '; '
            
            $rulesForGrid += $ruleObj
        }
        
        # Daten in das DataGrid laden
        if ($null -ne $script:dgInboxRules) {
            $script:dgInboxRules.Dispatcher.Invoke([action]{
                $script:dgInboxRules.ItemsSource = $rulesForGrid
            }, "Normal")
        }
        
        Write-Log "Inbox Rules erfolgreich geladen: $($rulesForGrid.Count) Regeln" -Type "Success"
        
        # Status aktualisieren
        if ($null -ne $script:txtStatus) {
            Update-GuiText -TextElement $script:txtStatus -Message "Inbox Rules geladen: $($rulesForGrid.Count) Regeln für $UserIdentity gefunden."
        }
        
        Log-Action "Inbox Rules abgerufen für $UserIdentity - $($rulesForGrid.Count) Regeln"
        return $rulesForGrid
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Abrufen der Inbox Rules: $errorMsg" -Type "Error"
        
        if ($null -ne $script:txtStatus) {
            Update-GuiText -TextElement $script:txtStatus -Message "Fehler beim Laden der Inbox Rules: $errorMsg"
        }
        
        Log-Action "Fehler beim Abrufen der Inbox Rules für $UserIdentity - $errorMsg"
        return @()
    }
}

function New-InboxRuleAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdentity,
        
        [Parameter(Mandatory = $true)]
        [string]$RuleName,
        
        [Parameter(Mandatory = $true)]
        [string]$FromAddress,
        
        [Parameter(Mandatory = $false)]
        [string]$TargetFolder = "Inbox",
        
        [Parameter(Mandatory = $false)]
        [bool]$MarkAsRead = $false
    )
    
    try {
        Write-Log "Erstelle neue Inbox Rule: $RuleName für $UserIdentity" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Parameter für die Regel aufbauen
        $ruleParams = @{
            Mailbox = $UserIdentity
            Name = $RuleName
            From = $FromAddress
            ErrorAction = "Stop"
        }
        
        # Zielordner setzen (falls nicht Inbox)
        if ($TargetFolder -ne "Inbox") {
            $fullFolderPath = "$($UserIdentity):\$($TargetFolder)"
            $ruleParams.Add("MoveToFolder", $fullFolderPath)
        }
        
        # Als gelesen markieren
        if ($MarkAsRead) {
            $ruleParams.Add("MarkAsRead", $true)
        }
        
        # Inbox Rule erstellen
        $newRule = New-InboxRule @ruleParams
        
        Write-Log "Inbox Rule '$RuleName' erfolgreich erstellt für $UserIdentity" -Type "Success"
        Log-Action "Inbox Rule erstellt: $RuleName für $UserIdentity"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Erstellen der Inbox Rule: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Erstellen der Inbox Rule $RuleName für $UserIdentity`: $errorMsg"
        throw "Fehler beim Erstellen der Regel: $errorMsg"
    }
}

function Set-InboxRuleStateAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdentity,
        
        [Parameter(Mandatory = $true)]
        [string]$RuleIdentity,
        
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )
    
    try {
        $state = if ($Enabled) { "aktiviert" } else { "deaktiviert" }
        Write-Log "Setze Inbox Rule auf Status $state`: $RuleIdentity für $UserIdentity" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Inbox Rule Status ändern
        if ($Enabled) {
            Enable-InboxRule -Mailbox $UserIdentity -Identity $RuleIdentity -ErrorAction Stop
        }
        else {
            Disable-InboxRule -Mailbox $UserIdentity -Identity $RuleIdentity -ErrorAction Stop
        }
        
        Write-Log "Inbox Rule '$RuleIdentity' erfolgreich $state für $UserIdentity" -Type "Success"
        Log-Action "Inbox Rule '$RuleIdentity' wurde $state für $UserIdentity"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        $statusMsg = if ($Enabled) { "Aktivieren" } else { "Deaktivieren" }
        Write-Log "Fehler beim $statusMsg der Inbox Rule '$RuleIdentity': $errorMsg" -Type "Error"
        Log-Action "Fehler beim $statusMsg der Inbox Rule '$RuleIdentity' für $UserIdentity`: $errorMsg"
        throw "Fehler beim $statusMsg der Regel: $errorMsg"
    }
}

function Remove-InboxRuleAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdentity,
        
        [Parameter(Mandatory = $true)]
        [string]$RuleIdentity
    )
    
    try {
        Write-Log "Lösche Inbox Rule: $RuleIdentity für $UserIdentity" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Inbox Rule löschen
        Remove-InboxRule -Mailbox $UserIdentity -Identity $RuleIdentity -Confirm:$false -ErrorAction Stop
        
        Write-Log "Inbox Rule '$RuleIdentity' erfolgreich gelöscht für $UserIdentity" -Type "Success"
        Log-Action "Inbox Rule '$RuleIdentity' wurde gelöscht für $UserIdentity"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Löschen der Inbox Rule '$RuleIdentity': $errorMsg" -Type "Error"
        Log-Action "Fehler beim Löschen der Inbox Rule '$RuleIdentity' für $UserIdentity`: $errorMsg"
        throw "Fehler beim Löschen der Regel: $errorMsg"
    }
}

function Move-InboxRuleAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdentity,
        
        [Parameter(Mandatory = $true)]
        [string]$RuleIdentity,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("Up", "Down")]
        [string]$Direction
    )
    
    try {
        $directionText = if ($Direction -eq "Up") { "nach oben" } else { "nach unten" }
        Write-Log "Verschiebe Inbox Rule $directionText`: $RuleIdentity für $UserIdentity" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Aktuelle Regeln abrufen, um die Priorität zu bestimmen
        $currentRules = Get-InboxRule -Mailbox $UserIdentity | Sort-Object Priority
        $currentRule = $currentRules | Where-Object { $_.Identity -eq $RuleIdentity }
        
        if (-not $currentRule) {
            throw "Regel nicht gefunden: $RuleIdentity"
        }
        
        $currentPriority = $currentRule.Priority
        $newPriority = $currentPriority
        
        if ($Direction -eq "Up") {
            # Die niedrigste Priorität ist 0. Set-InboxRule erwartet 1 als niedrigsten Wert.
            # Get-InboxRule gibt Priorität 0 zurück, aber Set-InboxRule -Priority 0 schlägt fehl.
            # Wir gehen davon aus, dass die Prioritäten bei 1 beginnen.
            if ($currentPriority -gt 1) {
                $newPriority = $currentPriority - 1
            }
        }
        elseif ($Direction -eq "Down") {
            $maxPriority = ($currentRules | Measure-Object Priority -Maximum).Maximum
            if ($currentPriority -lt $maxPriority) {
                $newPriority = $currentPriority + 1
            }
        }
        
        if ($newPriority -ne $currentPriority) {
            # Inbox Rule Priorität ändern
            Set-InboxRule -Mailbox $UserIdentity -Identity $RuleIdentity -Priority $newPriority -ErrorAction Stop
            
            Write-Log "Inbox Rule '$RuleIdentity' erfolgreich $directionText verschoben für $UserIdentity" -Type "Success"
            Log-Action "Inbox Rule '$RuleIdentity' wurde $directionText verschoben für $UserIdentity (Priorität: $currentPriority -> $newPriority)"
            
            return $true
        }
        else {
            Write-Log "Inbox Rule kann nicht weiter $directionText verschoben werden" -Type "Warning"
            Show-MessageBox -Message "Die Regel befindet sich bereits ganz oben oder ganz unten in der Liste." -Title "Verschieben nicht möglich" -Type Info
            return $false
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Verschieben der Inbox Rule '$RuleIdentity': $errorMsg" -Type "Error"
        Log-Action "Fehler beim Verschieben der Inbox Rule '$RuleIdentity' für $UserIdentity`: $errorMsg"
        throw "Fehler beim Verschieben der Regel: $errorMsg"
    }
}

function Export-InboxRulesAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdentity,
        
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        Write-Log "Exportiere Inbox Rules für $UserIdentity nach: $FilePath" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Inbox Rules abrufen
        $inboxRules = Get-InboxRule -Mailbox $UserIdentity | Select-Object Name, Description, Enabled, Priority,
            From, FromAddressContainsWords, SubjectContainsWords, BodyContainsWords,
            MoveToFolder, CopyToFolder, ForwardTo, RedirectTo, MarkAsRead, DeleteMessage
        
        # Bestimme Exportformat basierend auf Dateiendung
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        
        switch ($extension) {
            ".csv" {
                $inboxRules | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            }
            ".xml" {
                $inboxRules | Export-Clixml -Path $FilePath -Encoding UTF8
            }
            default {
                # Fallback zu CSV
                $inboxRules | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            }
        }
        
        Write-Log "Inbox Rules erfolgreich exportiert: $($inboxRules.Count) Regeln" -Type "Success"
        Log-Action "Inbox Rules exportiert für $UserIdentity nach: $FilePath ($($inboxRules.Count) Regeln)"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Exportieren der Inbox Rules: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Exportieren der Inbox Rules für $UserIdentity`: $errorMsg"
        throw "Fehler beim Exportieren der Regeln: $errorMsg"
    }
}

function Get-MailboxFoldersAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdentity
    )
    
    try {
        Write-Log "Rufe Postfach-Ordner ab für: $UserIdentity" -Type "Info"
        
        if (-not $script:isConnected) {
            throw "Nicht mit Exchange Online verbunden."
        }
        
        # Postfach-Ordnerstatistiken abrufen
        $folders = Get-MailboxFolderStatistics -Identity $UserIdentity | 
                   Where-Object { $_.FolderType -ne "Root" } |
                   Select-Object @{Name="DisplayName"; Expression={$_.FolderPath.Replace("\", " \ ") -replace "/", " / "}}, @{Name="Identity"; Expression={$_.FolderPath}} |
                   Sort-Object DisplayName
        
        Write-Log "Postfach-Ordner erfolgreich abgerufen: $($folders.Count) Ordner" -Type "Success"
        Log-Action "Postfach-Ordner abgerufen für $UserIdentity`: $($folders.Count) Ordner"
        
        return $folders
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Abrufen der Postfach-Ordner: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Abrufen der Postfach-Ordner für $UserIdentity`: $errorMsg"
        return @()
    }
}

function Refresh-InboxRuleUserDropdown {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Aktualisiere Benutzerliste für Inbox Rules" -Type "Info"
        
        # Sicherstellen, dass das ComboBox-Element über den $script-Scope verfügbar ist
        if ($null -eq $script:cmbInboxRuleUser) {
            $script:cmbInboxRuleUser = Get-XamlElement -ElementName "cmbInboxRuleUser"
            if ($null -eq $script:cmbInboxRuleUser) {
                $errorMsg = "ComboBox 'cmbInboxRuleUser' konnte nicht gefunden werden."
                Write-Log $errorMsg -Type "Error"
                Show-MessageBox -Message $errorMsg -Title "UI Fehler" -Type Error
                return
            }
        }

        $script:cmbInboxRuleUser.Dispatcher.Invoke({
            $script:cmbInboxRuleUser.ItemsSource = $null
            $script:cmbInboxRuleUser.Items.Clear()
        }, "Normal") | Out-Null
        
        if (-not $script:isConnected) {
            Write-Log "Keine Exchange-Verbindung für Benutzerabfrage" -Type "Warning"
            return $false
        }
        
        # Alle Postfächer abrufen
        $mailboxes = Get-Mailbox -ResultSize Unlimited | 
                     Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" } |
                     Select-Object DisplayName, PrimarySmtpAddress |
                     Sort-Object DisplayName

        $script:cmbInboxRuleUser.Dispatcher.Invoke({
            if ($null -ne $mailboxes -and $mailboxes.Count -gt 0) {
                foreach ($mailbox in $mailboxes) {
                    $item = New-Object System.Windows.Controls.ComboBoxItem
                    $item.Content = "$($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
                    $item.Tag = $mailbox.PrimarySmtpAddress
                    $item.ToolTip = "E-Mail: $($mailbox.PrimarySmtpAddress)"
                    [void]$script:cmbInboxRuleUser.Items.Add($item)
                }

                if ($script:cmbInboxRuleUser.Items.Count -gt 0) {
                    $script:cmbInboxRuleUser.SelectedIndex = -1 # Keine Auswahl
                }
                
                $statusMsg = "Benutzerliste aktualisiert: $($mailboxes.Count) Benutzer geladen."
                Write-Log $statusMsg -Type "Success"
                Update-GuiText -TextElement $script:txtStatus -Message $statusMsg
            }
            else {
                $statusMsg = "Keine Benutzerpostfächer gefunden."
                Write-Log $statusMsg -Type "Warning"
                Update-GuiText -TextElement $script:txtStatus -Message $statusMsg
                
                $placeholderItem = New-Object System.Windows.Controls.ComboBoxItem
                $placeholderItem.Content = "-- Keine Postfächer gefunden --"
                $placeholderItem.IsEnabled = $false
                [void]$script:cmbInboxRuleUser.Items.Add($placeholderItem)
            }
        }, "Normal") | Out-Null
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Aktualisieren der Benutzerliste für Inbox Rules: $errorMsg" -Type "Error"
        Show-MessageBox -Message "Fehler beim Laden der Benutzer: $errorMsg" -Title "Fehler" -Type Error
        return $false
    }
}

# -------------------------------------------------
# Abschnitt: MessageTrace-Funktionen
# -------------------------------------------------

function Start-MessageTraceAction {
    [CmdletBinding()]
    param()

    if (-not (Confirm-ExchangeConnection)) { return }

    Update-StatusBar -Message "Nachrichtensuche wird gestartet..." -Type "Info"

    $params = @{
        ErrorAction = 'Stop'
    }

    if (-not [string]::IsNullOrWhiteSpace($script:txtTraceSender.Text)) { $params.SenderAddress = $script:txtTraceSender.Text.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($script:txtTraceRecipient.Text)) { $params.RecipientAddress = $script:txtTraceRecipient.Text.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($script:txtTraceSubject.Text)) { $params.Subject = "*$($script:txtTraceSubject.Text.Trim())*" }
    if (-not [string]::IsNullOrWhiteSpace($script:txtTraceMessageId.Text)) { $params.MessageId = $script:txtTraceMessageId.Text.Trim() }
    if ($null -ne $script:dpTraceStart.SelectedDate) { $params.StartDate = $script:dpTraceStart.SelectedDate }
    if ($null -ne $script:dpTraceEnd.SelectedDate) { $params.EndDate = $script:dpTraceEnd.SelectedDate }
    
    $selectedStatus = $script:cmbTraceStatus.SelectedItem.Content
    if ($selectedStatus -ne "All") {
        $params.Status = $selectedStatus
    }

    $traceResults = Get-MessageTrace @params
    
    $script:dgMessageTrace.ItemsSource = $traceResults
    $resultCount = if ($null -ne $traceResults) { $traceResults.Count } else { 0 }
    $script:txtTraceCount.Text = "Results: $resultCount"
    Update-StatusBar -Message "Nachrichtensuche abgeschlossen. $resultCount Ergebnis(se) gefunden." -Type "Success"
}

function Export-MessageTraceAction {
    [CmdletBinding()]
    param()

    $dataToExport = $script:dgMessageTrace.ItemsSource
    if ($null -eq $dataToExport -or $dataToExport.Count -eq 0) {
        Show-MessageBox -Message "Es sind keine Daten zum Exportieren vorhanden." -Title "Keine Daten" -Type "Warning"
        return
    }

    $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv"
    $saveFileDialog.FileName = "MessageTrace_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveFileDialog.ShowDialog() -eq $true) {
        $exportPath = $saveFileDialog.FileName
        $dataToExport | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        Update-StatusBar -Message "Ergebnisse erfolgreich nach '$exportPath' exportiert." -Type "Success"
        Show-MessageBox -Message "Die Ergebnisse wurden erfolgreich exportiert." -Title "Export erfolgreich" -Type "Info"
    }
}

function Show-DetailedMessageTraceAction {
    [CmdletBinding()]
    param()

    if (-not (Confirm-ExchangeConnection)) { return }

    $selectedTrace = $script:dgMessageTrace.SelectedItem
    if ($null -eq $selectedTrace) {
        Show-MessageBox -Message "Bitte wählen Sie einen Eintrag aus der Ergebnisliste aus." -Title "Keine Auswahl" -Type "Warning"
        return
    }

    Update-StatusBar -Message "Lade detaillierte Informationen für die ausgewählte Nachricht..." -Type "Info"
    
    $details = Get-MessageTraceDetail -MessageTraceId $selectedTrace.MessageTraceId -RecipientAddress $selectedTrace.RecipientAddress
    
    $detailString = "Message Trace Details for Message from $($selectedTrace.SenderAddress) to $($selectedTrace.RecipientAddress):`n"
    $detailString += "Subject: $($selectedTrace.Subject)`n"
    $detailString += "--------------------------------------------------`n`n"
    
    foreach ($event in $details) {
        $detailString += "Date: $($event.Date)`n"
        $detailString += "Event: $($event.Event)`n"
        $detailString += "Action: $($event.Action)`n"
        $detailString += "Detail: $($event.Detail)`n"
        $detailString += "---------------------------------`n"
    }

    # Erstelle ein einfaches Text-Fenster für die Anzeige
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Detailed Message Trace" Height="600" Width="800" WindowStartupLocation="CenterScreen">
    <ScrollViewer VerticalScrollBarVisibility="Auto">
        <TextBox TextWrapping="Wrap" IsReadOnly="True" Margin="10" Name="txtDetails" FontFamily="Consolas" />
    </ScrollViewer>
</Window>
"@
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    $txtDetails = $window.FindName("txtDetails")
    $txtDetails.Text = $detailString
    [void]$window.ShowDialog()

    Update-StatusBar -Message "Detaillierte Informationen angezeigt." -Type "Success"
}

# -------------------------------------------------
# Abschnitt: AutoReply-Funktionen
# -------------------------------------------------

function Get-AutoReplyStatusAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$UserIdentity
    )

    if (-not (Confirm-ExchangeConnection)) { return }
    Update-StatusBar -Message "Lade Status der Abwesenheitsnotizen..." -Type Info

    try {
        $mailboxes = @()
        if (-not [string]::IsNullOrWhiteSpace($UserIdentity)) {
            $mailboxes = Get-Mailbox -Identity $UserIdentity -ErrorAction Stop
        } else {
            $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
        }

        $statusList = [System.Collections.Generic.List[object]]::new()
        foreach ($mailbox in $mailboxes) {
            try {
                $config = Get-MailboxAutoReplyConfiguration -Identity $mailbox.UserPrincipalName -ErrorAction Stop
                
                # Bestimme die Anzeige für StartTime und EndTime basierend auf den tatsächlichen Werten
                $startTimeDisplay = "-"
                $endTimeDisplay = "-"
                
                if ($config.StartTime -and $config.StartTime -ne [DateTime]::MinValue) {
                    $startTimeDisplay = Get-Date $config.StartTime -Format "g"
                }
                
                if ($config.EndTime -and $config.EndTime -ne [DateTime]::MinValue) {
                    $endTimeDisplay = Get-Date $config.EndTime -Format "g"
                }
                
                # Prüfe, ob tatsächlich eine Abwesenheitsnotiz konfiguriert ist
                $hasAutoReply = $config.AutoReplyState -ne "Disabled"
                $internalMessage = if ($hasAutoReply -and -not [string]::IsNullOrWhiteSpace($config.InternalMessage)) { 
                    $config.InternalMessage 
                } else { 
                    "Keine Abwesenheitsnotiz konfiguriert" 
                }
                
                $statusList.Add([PSCustomObject]@{
                    DisplayName       = $mailbox.DisplayName
                    UserPrincipalName = $mailbox.UserPrincipalName
                    AutoReplyState    = $hasAutoReply
                    StartTime         = $startTimeDisplay
                    EndTime           = $endTimeDisplay
                    InternalMessage   = $internalMessage
                })
            }
            catch {
                # Einzelnen Benutzer-Fehler behandeln, aber fortfahren
                $statusList.Add([PSCustomObject]@{
                    DisplayName       = $mailbox.DisplayName
                    UserPrincipalName = $mailbox.UserPrincipalName
                    AutoReplyState    = $false
                    StartTime         = "-"
                    EndTime           = "-"
                    InternalMessage   = "Fehler beim Abrufen der Konfiguration"
                })
                Write-Log "Fehler beim Abrufen der AutoReply-Konfiguration für $($mailbox.UserPrincipalName): $($_.Exception.Message)" -Type Warning
            }
        }

        if ($null -ne $script:dgAutoReplyStatus) {
            $script:dgAutoReplyStatus.ItemsSource = $statusList
        }
        
        $activeCount = ($statusList | Where-Object { $_.AutoReplyState -eq $true }).Count
        Update-StatusBar -Message "$($statusList.Count) Statusinformationen geladen ($activeCount aktive Abwesenheitsnotizen)." -Type Success
    } catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_
        Write-Log "Fehler in Get-AutoReplyStatusAction: $errorMsg" -Type Error
        Update-StatusBar -Message "Fehler beim Laden der Statusinformationen." -Type Error
    }
}

function Set-AutoReplyAction {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$UserIdentity,
        [Parameter(Mandatory = $true)][datetime]$StartDate,
        [Parameter(Mandatory = $true)][datetime]$EndDate,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not (Confirm-ExchangeConnection)) { return }
    Update-StatusBar -Message "Aktiviere Abwesenheitsnotiz für $UserIdentity..." -Type Info

    try {
        $actionDescription = "Abwesenheitsnotiz für '$UserIdentity' von $StartDate bis $EndDate aktivieren."
        if ($PSCmdlet.ShouldProcess($UserIdentity, $actionDescription)) {
            Set-MailboxAutoReplyConfiguration -Identity $UserIdentity `
                -AutoReplyState Scheduled `
                -StartTime $StartDate `
                -EndTime $EndDate `
                -InternalMessage $Message `
                -ExternalMessage $Message `
                -ExternalAudience All `
                -ErrorAction Stop
            
            Update-StatusBar -Message "Abwesenheitsnotiz für $UserIdentity erfolgreich aktiviert." -Type Success
            Get-AutoReplyStatusAction -UserIdentity $UserIdentity # Status aktualisieren
        }
    } catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_
        Write-Log "Fehler in Set-AutoReplyAction für '$UserIdentity': $errorMsg" -Type Error
        Update-StatusBar -Message "Fehler beim Aktivieren der Abwesenheitsnotiz für $UserIdentity." -Type Error
    }
}

function Disable-AutoReplyAction {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$UserIdentity
    )

    if (-not (Confirm-ExchangeConnection)) { return }
    Update-StatusBar -Message "Deaktiviere Abwesenheitsnotiz für $UserIdentity..." -Type Info

    try {
        $actionDescription = "Abwesenheitsnotiz für '$UserIdentity' deaktivieren."
        if ($PSCmdlet.ShouldProcess($UserIdentity, $actionDescription)) {
            Set-MailboxAutoReplyConfiguration -Identity $UserIdentity -AutoReplyState Disabled -ErrorAction Stop
            Update-StatusBar -Message "Abwesenheitsnotiz für $UserIdentity erfolgreich deaktiviert." -Type Success
            Get-AutoReplyStatusAction -UserIdentity $UserIdentity # Status aktualisieren
        }
    } catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_
        Write-Log "Fehler in Disable-AutoReplyAction für '$UserIdentity': $errorMsg" -Type Error
        Update-StatusBar -Message "Fehler beim Deaktivieren der Abwesenheitsnotiz für $UserIdentity." -Type Error
    }
}

function Export-AutoReplyStatusAction {
    [CmdletBinding()]
    param()

    try {
        $data = $script:dgAutoReplyStatus.ItemsSource
        if ($null -eq $data -or $data.Count -eq 0) {
            Show-MessageBox -Message "Es sind keine Daten zum Exportieren vorhanden." -Title "Keine Daten" -Type Info
            return
        }

        $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
        $saveFileDialog.FileName = "AutoReplyStatus_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveFileDialog.ShowDialog() -eq $true) {
            $filePath = $saveFileDialog.FileName
            $data | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
            Update-StatusBar -Message "Daten erfolgreich nach '$filePath' exportiert." -Type Success
            Show-MessageBox -Message "Die Daten wurden erfolgreich exportiert." -Title "Export erfolgreich" -Type Info
        }
    } catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_
        Write-Log "Fehler in Export-AutoReplyStatusAction: $errorMsg" -Type Error
        Update-StatusBar -Message "Fehler beim Exportieren der Daten." -Type Error
    }
}

##############TABSFUNKTIONEN####################
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

    # KORREKTUR: Richtige Referenz auf das XAML-Element txtGetRegionMailbox
    $mailboxId = $script:txtGetRegionMailbox.Text.Trim()
    $statusTextBlock = $script:txtStatus
    $resultTextBlock = $script:txtRegionResult
    
    $languageComboBox = $script:cmbRegionLanguage
    $timezoneComboBox = $script:cmbRegionTimezone
    $dateFormatComboBox = $script:cmbRegionDateFormat
    $timeFormatComboBox = $script:cmbRegionTimeFormat
    $localizeFoldersCheckBox = $script:chkRegionDefaultFolderNameMatchingUserLanguage

    # Validierung der Postfach-Eingabe
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
                $noDataMessage = "Abruf für '$mailboxId' erfolgreich, aber keine Einstellungsdaten (regionalSettings ist null) zurückgegeben."
                Update-GuiText -TextElement $statusTextBlock -Message $noDataMessage
                Log-Action $noDataMessage
                if ($null -ne $resultTextBlock) {
                    $resultTextBlock.Dispatcher.Invoke({ $resultTextBlock.Text = "Keine Einstellungsdaten für '$mailboxId' empfangen." }) | Out-Null
                }
                
                # UI-Reset bei Null-Ergebnis
                try {
                    if ($null -ne $languageComboBox) { 
                        $languageComboBox.Dispatcher.Invoke({ 
                            $languageComboBox.SelectedItem = ($languageComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) 
                        }) | Out-Null 
                    }
                    Populate-DateFormatComboBox -ComboBox $dateFormatComboBox -CultureName "" 
                    Populate-TimeFormatComboBox -ComboBox $timeFormatComboBox -CultureName ""
                    Populate-TimezoneComboBox -ComboBox $timezoneComboBox -CultureName "DEFAULT_ALL" 
                    if ($null -ne $localizeFoldersCheckBox) { 
                        $localizeFoldersCheckBox.Dispatcher.Invoke({ $localizeFoldersCheckBox.IsChecked = $null }) | Out-Null 
                    }
                } catch { 
                    Write-Log "Fehler beim Zurücksetzen der UI nach Null-Ergebnis (GetRegion): $($_.Exception.Message)" -Type Warning 
                }

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

                # Sprache setzen
                if ($null -ne $regionalSettings.Language) {
                    $langCodeToSelect = $regionalSettings.Language.ToString()
                    $langItem = $languageComboBox.Items | Where-Object { $_.Tag -eq $langCodeToSelect } | Select-Object -First 1
                    if ($null -ne $langItem) {
                        $languageComboBox.Dispatcher.Invoke({ $languageComboBox.SelectedItem = $langItem }) | Out-Null
                        $currentCultureForUI = $langCodeToSelect 
                    } else {
                        $languageComboBox.Dispatcher.Invoke({ 
                            $languageComboBox.SelectedItem = ($languageComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) 
                        }) | Out-Null
                    }
                } else {
                     $languageComboBox.Dispatcher.Invoke({ 
                         $languageComboBox.SelectedItem = ($languageComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) 
                     }) | Out-Null
                }
                
                # Formate und Zeitzonen basierend auf der gesetzten Sprache laden
                Populate-DateFormatComboBox -ComboBox $dateFormatComboBox -CultureName $currentCultureForUI
                Populate-TimeFormatComboBox -ComboBox $timeFormatComboBox -CultureName $currentCultureForUI
                Populate-TimezoneComboBox -ComboBox $timezoneComboBox -CultureName $currentCultureForUI 
                
                # Datumsformat setzen
                if ($null -ne $regionalSettings.DateFormat) {
                    $dateFormatToSelect = $regionalSettings.DateFormat.ToString()
                    $dateItem = $dateFormatComboBox.Items | Where-Object { $_.Tag -eq $dateFormatToSelect } | Select-Object -First 1
                    if ($null -ne $dateItem) {
                        $dateFormatComboBox.Dispatcher.Invoke({ $dateFormatComboBox.SelectedItem = $dateItem }) | Out-Null
                    } else {
                        $dateFormatComboBox.Dispatcher.Invoke({ 
                            $dateFormatComboBox.SelectedItem = ($dateFormatComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) 
                        }) | Out-Null
                    }
                } else {
                     $dateFormatComboBox.Dispatcher.Invoke({ 
                         $dateFormatComboBox.SelectedItem = ($dateFormatComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) 
                     }) | Out-Null
                }
                
                # Zeitformat setzen
                if ($null -ne $regionalSettings.TimeFormat) {
                    $timeFormatToSelect = $regionalSettings.TimeFormat.ToString()
                    $timeItem = $timeFormatComboBox.Items | Where-Object { $_.Tag -eq $timeFormatToSelect } | Select-Object -First 1
                    if ($null -ne $timeItem) {
                        $timeFormatComboBox.Dispatcher.Invoke({ $timeFormatComboBox.SelectedItem = $timeItem }) | Out-Null
                        Write-Log "Zeitformat '$timeFormatToSelect' in UI ausgewählt." -Type Debug
                    } else {
                        Write-Log "Abgerufenes Zeitformat '$timeFormatToSelect' nicht in UI-Liste für Kultur '$currentCultureForUI'. (Keine Änderung) bleibt." -Type Warning
                        $timeFormatComboBox.Dispatcher.Invoke({ 
                            $timeFormatComboBox.SelectedItem = ($timeFormatComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) 
                        }) | Out-Null
                    }
                } else {
                     Write-Log "Kein Zeitformat von Exchange empfangen. UI bleibt auf (Keine Änderung)." -Type Debug
                     $timeFormatComboBox.Dispatcher.Invoke({ 
                         $timeFormatComboBox.SelectedItem = ($timeFormatComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) 
                     }) | Out-Null
                }

                # Zeitzone setzen
                if ($null -ne $regionalSettings.TimeZone) {
                    $timeZoneToSelect = $regionalSettings.TimeZone.ToString()
                    $tzItem = $timezoneComboBox.Items | Where-Object { $_.Tag -eq $timeZoneToSelect } | Select-Object -First 1
                    if ($null -ne $tzItem) {
                        $timezoneComboBox.Dispatcher.Invoke({ $timezoneComboBox.SelectedItem = $tzItem }) | Out-Null
                    } else {
                         $timezoneComboBox.Dispatcher.Invoke({ 
                             $timezoneComboBox.SelectedItem = ($timezoneComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) 
                         }) | Out-Null
                    }
                } else {
                     $timezoneComboBox.Dispatcher.Invoke({ 
                         $timezoneComboBox.SelectedItem = ($timezoneComboBox.Items | Where-Object {$_.Tag -eq ""} | Select-Object -First 1) 
                     }) | Out-Null
                }
                
                # LocalizeDefaultFolderName setzen
                if (($regionalSettings.PSObject.Properties.Name -contains "LocalizeDefaultFolderName") -and 
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

                # Verbesserte Ergebnis-Anzeige in der ResultBox
                if ($null -ne $resultTextBlock) {
                    $resultOutput = "Aktuelle regionale Einstellungen für: $retrievedIdentity`n"
                    $resultOutput += "=" * 60 + "`n`n"
                    
                    # Strukturierte Darstellung der Einstellungen
                    $resultOutput += "SPRACH- UND REGIONALE EINSTELLUNGEN:`n"
                    $resultOutput += "-" * 40 + "`n"
                    
                    # Sprache
                    if ($null -ne $regionalSettings.Language) {
                        $languageDisplay = $regionalSettings.Language.ToString()
                        # Versuche, den Anzeigenamen der Sprache zu finden
                        try {
                            $culture = [System.Globalization.CultureInfo]::new($regionalSettings.Language.ToString())
                            $languageDisplay = "$($culture.DisplayName) ($($regionalSettings.Language))"
                        } catch {
                            $languageDisplay = $regionalSettings.Language.ToString()
                        }
                        $resultOutput += "Sprache:        $languageDisplay`n"
                    } else {
                        $resultOutput += "Sprache:        (Nicht festgelegt)`n"
                    }
                    
                    # Zeitzone
                    if ($null -ne $regionalSettings.TimeZone) {
                        $timezoneDisplay = $regionalSettings.TimeZone.ToString()
                        # Versuche, zusätzliche Timezone-Info zu bekommen
                        try {
                            $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($regionalSettings.TimeZone.ToString())
                            $timezoneDisplay = "$($tz.DisplayName) ($($regionalSettings.TimeZone))"
                        } catch {
                            $timezoneDisplay = $regionalSettings.TimeZone.ToString()
                        }
                        $resultOutput += "Zeitzone:       $timezoneDisplay`n"
                    } else {
                        $resultOutput += "Zeitzone:       (Nicht festgelegt)`n"
                    }
                    
                    # Datumsformat
                    if ($null -ne $regionalSettings.DateFormat) {
                        $resultOutput += "Datumsformat:   $($regionalSettings.DateFormat)`n"
                    } else {
                        $resultOutput += "Datumsformat:   (Nicht festgelegt)`n"
                    }
                    
                    # Zeitformat
                    if ($null -ne $regionalSettings.TimeFormat) {
                        $resultOutput += "Zeitformat:     $($regionalSettings.TimeFormat)`n"
                    } else {
                        $resultOutput += "Zeitformat:     (Nicht festgelegt)`n"
                    }
                    
                    # Ordnernamen lokalisieren
                    $localizeDisplay = "Unbekannt"
                    if (($regionalSettings.PSObject.Properties.Name -contains "LocalizeDefaultFolderName") -and 
                        $null -ne $regionalSettings.LocalizeDefaultFolderName) {
                        $localizeDisplay = if ($regionalSettings.LocalizeDefaultFolderName -eq $true) { "Ja" } else { "Nein" }
                    } else {
                        $localizeDisplay = "(Nicht festgelegt)"
                    }
                    $resultOutput += "Ordnernamen lokalisieren: $localizeDisplay`n"
                    
                    $resultOutput += "`n" + "=" * 60 + "`n"
                    $resultOutput += "TECHNISCHE DETAILS:`n"
                    $resultOutput += "-" * 20 + "`n"
                    
                    # Weitere technische Eigenschaften anzeigen
                    $technicalProps = @("Identity", "DistinguishedName", "Guid", "ObjectCategory", "WhenCreated", "WhenChanged")
                    foreach ($prop in $technicalProps) {
                        if ($regionalSettings.PSObject.Properties.Name -contains $prop -and $null -ne $regionalSettings.$prop) {
                            $resultOutput += "${prop}: $($regionalSettings.$prop)`n"
                        }
                    }
                    
                    # Den zusammengestellten Text in die TextBox schreiben
                    $resultTextBlock.Dispatcher.Invoke({ $resultTextBlock.Text = $resultOutput }) | Out-Null
                }
                Update-GuiText -TextElement $statusTextBlock -Message $statusMsg
            }
        } else {
            $errorMsg = $OperationResult.Error
            Update-GuiText -TextElement $statusTextBlock -Message "Fehler beim Abrufen der Einstellungen."
            Log-Action "Fehler bei Get-ExoMailboxRegionalSettings für '$mailboxId': $errorMsg"
            if ($null -ne $resultTextBlock) {
                $resultTextBlock.Dispatcher.Invoke({ $resultTextBlock.Text = "Fehler: $errorMsg" }) | Out-Null
            }
            Show-MessageBox -Message "Fehler beim Abrufen der regionalen Einstellungen:`n$errorMsg" -Title "Fehler"
        }
    } catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Schwerwiegender Fehler beim Abrufen der regionalen Einstellungen."
        Update-GuiText -TextElement $statusTextBlock -Message "Fehler: $errorMsg"
        Log-Action "FEHLER bei Get-ExoMailboxRegionalSettings für '$mailboxId': $errorMsg"
        if ($null -ne $resultTextBlock) {
            $resultTextBlock.Dispatcher.Invoke({ $resultTextBlock.Text = "Fehler: $errorMsg" }) | Out-Null
        }
        Show-MessageBox -Message "Ein schwerwiegender Fehler ist aufgetreten:`n$errorMsg" -Title "Fehler"
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
            $group = Get-ExoUnifiedGroup-Identity $GroupName -ErrorAction SilentlyContinue
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
            $group = Get-ExoUnifiedGroup-Identity $GroupName -ErrorAction SilentlyContinue
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
            $group = Get-ExoUnifiedGroup-Identity $GroupName -ErrorAction SilentlyContinue
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

function Create-DistributionGroupAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $false)]
        [string]$Alias,
        
        [Parameter(Mandatory = $true)]
        [string]$Type, # 'Distribution' or 'Security'
        
        [Parameter(Mandatory = $false)]
        [string]$Notes,

        [Parameter(Mandatory = $false)]
        [switch]$RequireSenderAuthentication
    )
    
    try {
        Write-Log "Erstelle neue Gruppe: $Name (Typ: $Type)" -Type "Info"
        
        # Parameter für die Gruppenerstellung vorbereiten
        $params = @{
            Name = $Name
            Type = $Type
            RequireSenderAuthenticationEnabled = $RequireSenderAuthentication
            ErrorAction = "Stop"
        }
        
        if (-not [string]::IsNullOrEmpty($Alias)) {
            $params.Add("Alias", $Alias)
        }
        
        if (-not [string]::IsNullOrEmpty($Notes)) {
            $params.Add("Notes", $Notes)
        }
        
        # Gruppe erstellen
        $newGroup = New-DistributionGroup @params
        
        Write-Log "Gruppe '$($newGroup.DisplayName)' erfolgreich erstellt." -Type "Success"
        Log-Action "Gruppe '$($newGroup.DisplayName)' (Typ: $Type) wurde erstellt."
        
        # Status aktualisieren
        if ($null -ne $script:txtStatus) {
            Update-GuiText -TextElement $script:txtStatus -Message "Gruppe '$($newGroup.DisplayName)' erfolgreich erstellt."
        }
        
        return $newGroup
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Erstellen der Gruppe: $errorMsg" -Type "Error"
        Log-Action "Fehler beim Erstellen der Gruppe '$Name': $errorMsg"
        
        # Status aktualisieren
        if ($null -ne $script:txtStatus) {
            Update-GuiText -TextElement $script:txtStatus -Message "Fehler beim Erstellen der Gruppe: $errorMsg"
        }
        
        # Zeige Fehlermeldung an den Benutzer
        Show-MessageBox -Message "Fehler beim Erstellen der Gruppe: $errorMsg" -Title "Fehler" -Type "Error"
        
        return $null
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
        $unifiedGroups = Get-ExoUnifiedGroup -ErrorAction SilentlyContinue
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
                 $script:cmbSelectExistingGroup.SelectedIndex = -1 # Keine Auswahl
            }
            
            $statusMsg = "Gruppenliste aktualisiert. $($groups.Count) Gruppen geladen."
            Write-Log $statusMsg -Type "Success"
            Update-StatusBar -Message $statusMsg -Type Success
        }
        else {
            $statusMsg = "Keine Gruppen zum Anzeigen gefunden oder Fehler beim Laden."
            Write-Log $statusMsg -Type "Warning"
            Update-StatusBar -Message $statusMsg -Type Warning
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
            $group = Get-ExoUnifiedGroup-Identity $GroupName -ErrorAction SilentlyContinue
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
            $group = Get-ExoUnifiedGroup-Identity $GroupName -ErrorAction SilentlyContinue
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
            $group = Get-ExoUnifiedGroup-Identity $GroupName -ErrorAction SilentlyContinue
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

function Get-ForwardingMailboxesAction {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Rufe alle Postfächer mit Weiterleitung ab" -Type "Info"
        Update-StatusBar -Message "Rufe Postfächer mit Weiterleitung ab..." -Type Info

        # Alle Mailboxen abrufen und filtern
        $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
            (-not [string]::IsNullOrEmpty($_.ForwardingAddress)) -or (-not [string]::IsNullOrEmpty($_.ForwardingSmtpAddress))
        } | Select-Object DisplayName, PrimarySmtpAddress, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward

        if ($null -ne $mailboxes -and $mailboxes.Count -gt 0) {
            Write-Log "$($mailboxes.Count) Postfächer mit Weiterleitung gefunden" -Type "Info"
            Update-StatusBar -Message "$($mailboxes.Count) Postfächer mit Weiterleitung gefunden" -Type Info
        } else {
            Write-Log "Keine Postfächer mit Weiterleitung gefunden" -Type "Info"
            Update-StatusBar -Message "Keine Postfächer mit Weiterleitung gefunden" -Type Info
        }

        return $mailboxes
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Abrufen der Postfächer mit Weiterleitung: $errorMsg" -Type "Error"
        Update-StatusBar -Message "Fehler beim Abrufen der Postfächer mit Weiterleitung: $errorMsg" -Type Error
        return $null
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
        
        $rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox |
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
        
        $equipment = Get-Mailbox -RecipientTypeDetails EquipmentMailbox | 
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
        
        # Array-Syntax für mehrere RecipientTypeDetails verwenden
        $recipientTypes = @("RoomMailbox", "EquipmentMailbox")
        $resources = Get-Mailbox -RecipientTypeDetails $recipientTypes -ResultSize Unlimited
        
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
$script:txtStatus           = Get-XamlElement -ElementName "txtStatus" -Required
$script:txtVersion          = Get-XamlElement -ElementName "txtVersion"
$script:txtConnectionStatus = Get-XamlElement -ElementName "txtConnectionStatus" -Required
$script:btnClose            = Get-XamlElement -ElementName "btnClose" -Required

# Dashboard
$script:btnDashboard        = Get-XamlElement -ElementName "btnDashboard"
$script:tabDashboard        = Get-XamlElement -ElementName "tabDashboard"

# Grundlegende Verwaltung - Tabs
$script:tabCalendar         = Get-XamlElement -ElementName "tabCalendar"
$script:tabMailbox          = Get-XamlElement -ElementName "tabMailbox"
$script:tabSharedMailbox    = Get-XamlElement -ElementName "tabSharedMailbox"
$script:tabGroups           = Get-XamlElement -ElementName "tabGroups"
$script:tabResources        = Get-XamlElement -ElementName "tabResources"
$script:tabContacts         = Get-XamlElement -ElementName "tabContacts"

# Grundlegende Verwaltung - Navigation Buttons
$script:btnNavCalendar      = Get-XamlElement -ElementName "btnNavCalendar"
$script:btnNavMailbox       = Get-XamlElement -ElementName "btnNavMailbox"
$script:btnNavSharedMailbox = Get-XamlElement -ElementName "btnNavSharedMailbox"
$script:btnNavGroups        = Get-XamlElement -ElementName "btnNavGroups"
$script:btnNavResources     = Get-XamlElement -ElementName "btnNavResources"
$script:btnNavContacts      = Get-XamlElement -ElementName "btnNavContacts"

# Mail Flow - Tabs
$script:tabMailFlowRules    = Get-XamlElement -ElementName "tabMailFlowRules"
$script:tabInboxRules       = Get-XamlElement -ElementName "tabInboxRules"
$script:tabMessageTrace     = Get-XamlElement -ElementName "tabMessageTrace"
$script:tabAutoReply        = Get-XamlElement -ElementName "tabAutoReply"

# Mail Flow - Navigation Buttons
$script:btnNavMailFlowRules = Get-XamlElement -ElementName "btnNavMailFlowRules"
$script:btnNavInboxRules    = Get-XamlElement -ElementName "btnNavInboxRules"
$script:btnNavMessageTrace  = Get-XamlElement -ElementName "btnNavMessageTrace"
$script:btnNavAutoReply     = Get-XamlElement -ElementName "btnNavAutoReply"

# Sicherheit & Compliance - Tabs
$script:tab_ATP             = Get-XamlElement -ElementName "tab_ATP"
$script:tab_DLP             = Get-XamlElement -ElementName "tab_DLP"
$script:tab_eDiscovery      = Get-XamlElement -ElementName "tab_eDiscovery"
$script:tab_MDM             = Get-XamlElement -ElementName "tab_MDM"

# Sicherheit & Compliance - Navigation Buttons
$script:btnNavATP           = Get-XamlElement -ElementName "btnNavATP"
$script:btnNavDLP           = Get-XamlElement -ElementName "btnNavDLP"
$script:btnNavEDiscovery    = Get-XamlElement -ElementName "btnNavEDiscovery"
$script:btnNavMDM           = Get-XamlElement -ElementName "btnNavMDM"

# Systemkonfiguration - Tabs
$script:tabEXOSettings      = Get-XamlElement -ElementName "tabEXOSettings"
$script:tabRegion           = Get-XamlElement -ElementName "tabRegion"
$script:tab_MailRouting     = Get-XamlElement -ElementName "tab_MailRouting"

# Systemkonfiguration - Navigation Buttons
$script:btnNavEXOSettings   = Get-XamlElement -ElementName "btnNavEXOSettings"
$script:btnNavRegion        = Get-XamlElement -ElementName "btnNavRegion"
$script:btnNavCrossPremises = Get-XamlElement -ElementName "btnNavCrossPremises"

# Erweiterte Verwaltung - Tabs
$script:tab_HybridExchange  = Get-XamlElement -ElementName "tab_HybridExchange"
$script:tab_MultiForest     = Get-XamlElement -ElementName "tab_MultiForest"

# Erweiterte Verwaltung - Navigation Buttons
$script:btnNavHybridExchange = Get-XamlElement -ElementName "btnNavHybridExchange"
$script:btnNavMultiForest   = Get-XamlElement -ElementName "btnNavMultiForest"

# Monitoring & Support - Tabs
$script:tab_HealthCheck     = Get-XamlElement -ElementName "tab_HealthCheck"
$script:tabMailboxAudit     = Get-XamlElement -ElementName "tabMailboxAudit"
$script:tabReports          = Get-XamlElement -ElementName "tabReports"
$script:tabTroubleshooting  = Get-XamlElement -ElementName "tabTroubleshooting"

# Monitoring & Support - Navigation Buttons
$script:btnNavHealthCheck   = Get-XamlElement -ElementName "btnNavHealthCheck"
$script:btnNavAudit         = Get-XamlElement -ElementName "btnNavAudit"
$script:btnNavReports       = Get-XamlElement -ElementName "btnNavReports"
$script:btnNavTroubleshooting = Get-XamlElement -ElementName "btnNavTroubleshooting"

# Legacy/Fallback Referenzen (falls noch verwendet)
$script:tabOrgSettings      = Get-XamlElement -ElementName "tabOrgSettings"
$script:btnNavOrgSettings   = Get-XamlElement -ElementName "btnNavOrgSettings"
$script:btnInfo             = Get-XamlElement -ElementName "btnInfo"
$script:btnSettings         = Get-XamlElement -ElementName "btnSettings"

# Button-Handler
$script:btnConnect.Add_Click({
    if (-not (Test-PowerShell7AndAdminRights)) {
        return
    }
    # Prüfen, ob die erforderlichen Module vorhanden sind, bevor die Verbindung versucht wird.
    if (-not (Check-Prerequisites)) {
        # Die Funktion Check-Prerequisites sollte den Benutzer bereits informiert haben.
        return
    }
    Connect-OwnExchangeOnline
})

$script:btnClose.Add_Click({ $script:Form.Close() })

# Hier den Tab-Handler einfügen:
$script:tabContent.Add_SelectionChanged({
    param($sender, $e)
    $selectedTab = $sender.SelectedItem
    if ($null -ne $selectedTab) {
        Write-Log "Tab gewechselt zu: $($selectedTab.Header)"
    }
})

# Dashboard
if ($null -ne $script:btnDashboard) {
    $script:btnDashboard.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabDashboard) {
            $script:tabContent.SelectedItem = $script:tabDashboard
        } else {
            Write-Log "Fehler: Dashboard Tab oder TabControl ist null" -Type "Error"
        }
    })
}

# Grundlegende Verwaltung
if ($null -ne $script:btnNavCalendar) {
    $script:btnNavCalendar.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabCalendar) {
            $script:tabContent.SelectedItem = $script:tabCalendar
        } else {
            Write-Log "Fehler: Calendar Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavMailbox) {
    $script:btnNavMailbox.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabMailbox) {
            $script:tabContent.SelectedItem = $script:tabMailbox
        } else {
            Write-Log "Fehler: Mailbox Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavSharedMailbox) {
    $script:btnNavSharedMailbox.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabSharedMailbox) {
            $script:tabContent.SelectedItem = $script:tabSharedMailbox
        } else {
            Write-Log "Fehler: SharedMailbox Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavGroups) {
    $script:btnNavGroups.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabGroups) {
            $script:tabContent.SelectedItem = $script:tabGroups
        } else {
            Write-Log "Fehler: Groups Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavResources) {
    $script:btnNavResources.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabResources) {
            $script:tabContent.SelectedItem = $script:tabResources
        } else {
            Write-Log "Fehler: Resources Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavContacts) {
    $script:btnNavContacts.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabContacts) {
            $script:tabContent.SelectedItem = $script:tabContacts
        } else {
            Write-Log "Fehler: Contacts Tab oder TabControl ist null" -Type "Error"
        }
    })
}

# Mail Flow
if ($null -ne $script:btnNavMailFlowRules) {
    $script:btnNavMailFlowRules.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabMailFlowRules) {
            $script:tabContent.SelectedItem = $script:tabMailFlowRules
        } else {
            Write-Log "Fehler: MailFlowRules Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavInboxRules) {
    $script:btnNavInboxRules.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabInboxRules) {
            $script:tabContent.SelectedItem = $script:tabInboxRules
        } else {
            Write-Log "Fehler: InboxRules Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavMessageTrace) {
    $script:btnNavMessageTrace.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabMessageTrace) {
            $script:tabContent.SelectedItem = $script:tabMessageTrace
        } else {
            Write-Log "Fehler: MessageTrace Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavAutoReply) {
    $script:btnNavAutoReply.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabAutoReply) {
            $script:tabContent.SelectedItem = $script:tabAutoReply
        } else {
            Write-Log "Fehler: AutoReply Tab oder TabControl ist null" -Type "Error"
        }
    })
}

# Sicherheit & Compliance
if ($null -ne $script:btnNavATP) {
    $script:btnNavATP.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tab_ATP) {
            $script:tabContent.SelectedItem = $script:tab_ATP
        } else {
            Write-Log "Fehler: ATP Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavDLP) {
    $script:btnNavDLP.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tab_DLP) {
            $script:tabContent.SelectedItem = $script:tab_DLP
        } else {
            Write-Log "Fehler: DLP Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavEDiscovery) {
    $script:btnNavEDiscovery.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tab_eDiscovery) {
            $script:tabContent.SelectedItem = $script:tab_eDiscovery
        } else {
            Write-Log "Fehler: eDiscovery Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavMDM) {
    $script:btnNavMDM.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tab_MDM) {
            $script:tabContent.SelectedItem = $script:tab_MDM
        } else {
            Write-Log "Fehler: MDM Tab oder TabControl ist null" -Type "Error"
        }
    })
}

# Systemkonfiguration
if ($null -ne $script:btnNavEXOSettings) {
    $script:btnNavEXOSettings.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabEXOSettings) {
            $script:tabContent.SelectedItem = $script:tabEXOSettings
        } else {
            Write-Log "Fehler: EXOSettings Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavRegion) {
    $script:btnNavRegion.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabRegion) {
            $script:tabContent.SelectedItem = $script:tabRegion
        } else {
            Write-Log "Fehler: Region Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavCrossPremises) {
    $script:btnNavCrossPremises.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tab_MailRouting) {
            $script:tabContent.SelectedItem = $script:tab_MailRouting
        } else {
            Write-Log "Fehler: MailRouting Tab oder TabControl ist null" -Type "Error"
        }
    })
}

# Erweiterte Verwaltung
if ($null -ne $script:btnNavHybridExchange) {
    $script:btnNavHybridExchange.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tab_HybridExchange) {
            $script:tabContent.SelectedItem = $script:tab_HybridExchange
        } else {
            Write-Log "Fehler: HybridExchange Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavMultiForest) {
    $script:btnNavMultiForest.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tab_MultiForest) {
            $script:tabContent.SelectedItem = $script:tab_MultiForest
        } else {
            Write-Log "Fehler: MultiForest Tab oder TabControl ist null" -Type "Error"
        }
    })
}

# Monitoring & Support
if ($null -ne $script:btnNavHealthCheck) {
    $script:btnNavHealthCheck.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tab_HealthCheck) {
            $script:tabContent.SelectedItem = $script:tab_HealthCheck
        } else {
            Write-Log "Fehler: HealthCheck Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavAudit) {
    $script:btnNavAudit.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabMailboxAudit) {
            $script:tabContent.SelectedItem = $script:tabMailboxAudit
        } else {
            Write-Log "Fehler: MailboxAudit Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavReports) {
    $script:btnNavReports.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabReports) {
            $script:tabContent.SelectedItem = $script:tabReports
        } else {
            Write-Log "Fehler: Reports Tab oder TabControl ist null" -Type "Error"
        }
    })
}

if ($null -ne $script:btnNavTroubleshooting) {
    $script:btnNavTroubleshooting.Add_Click({
        if ($null -ne $script:tabContent -and $null -ne $script:tabTroubleshooting) {
            $script:tabContent.SelectedItem = $script:tabTroubleshooting
        } else {
            Write-Log "Fehler: Troubleshooting Tab oder TabControl ist null" -Type "Error"
        }
    })
}

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

        # Registriere Event-Handler für Speichern und Exportieren nach erfolgreichem Laden
        $btnSetOrganizationConfig = Get-XamlElement -ElementName "btnSetOrganizationConfig"
        $btnExportOrganizationConfig = Get-XamlElement -ElementName "btnExportOrganizationConfig"
        
        if ($null -ne $btnSetOrganizationConfig -and -not $script:EXOSettingsHandlersRegistered) {
            $btnSetOrganizationConfig.Add_Click({ Set-CustomOrganizationConfig })
            Write-Log "Event-Handler für btnSetOrganizationConfig registriert." -Type "Debug"
        }
        
        if ($null -ne $btnExportOrganizationConfig -and -not $script:EXOSettingsHandlersRegistered) {
            $btnExportOrganizationConfig.Add_Click({ Export-OrganizationConfig })
            Write-Log "Event-Handler für btnExportOrganizationConfig registriert." -Type "Debug"
        }
        
        $script:EXOSettingsHandlersRegistered = $true

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

                # Spezifische Zuordnungen für Elemente mit abweichenden PropertyNames
                switch ($elementName) {
                    "txtActivityBasedAuthenticationTimeoutInterval" { $propertyName = "ActivityBasedAuthenticationTimeoutInterval"; break }
                    "txtMailTipsLargeAudienceThreshold" { $propertyName = "MailTipsLargeAudienceThreshold"; break }
                    "cmbSearchQueryLanguage" { $propertyName = "SearchQueryLanguage"; break }
                    "cmbShortenEventScopeDefault" { $propertyName = "ShortenEventScopeDefault"; break }
                    "cmbLargeAudienceThreshold" { $propertyName = "MailTipsLargeAudienceThreshold"; break }
                    "cmbEwsAppAccessPolicy" { $propertyName = "EwsApplicationAccessPolicy"; break }
                    "cmbPublicFoldersEnabled" { $propertyName = "PublicFoldersEnabled"; break }
                    "txtDefaultAuthPolicy" { $propertyName = "DefaultAuthenticationPolicy"; break }
                    "txtWACDiscoveryEndpoint" { $propertyName = "WACDiscoveryEndpoint"; break }
                    "txtEwsAllowList" { $propertyName = "EwsAllowList"; break }
                    "txtEwsBlockList" { $propertyName = "EwsBlockList"; break }
                    "txtIPListBlocked" { $propertyName = "IPListBlocked"; break }
                    # Fehlende Bookings-Zuordnungen hinzufügen
                    "txtBookingsNamingPolicyPrefix" { $propertyName = "BookingsNamingPolicyPrefix"; break }
                    "txtBookingsNamingPolicySuffix" { $propertyName = "BookingsNamingPolicySuffix"; break }
                    "txtBookingsSchedulingPolicy" { $propertyName = "BookingsSchedulingPolicy"; break }
                    # Fehlende CheckBox-Zuordnungen für spezielle Fälle
                    "chkBookingsBlockedWordsEnabled" { $propertyName = "BookingsBlockedWordsEnabled"; break }
                    "chkBookingsNamingPolicyEnabled" { $propertyName = "BookingsNamingPolicyEnabled"; break }
                    "chkBookingsNamingPolicyPrefixEnabled" { $propertyName = "BookingsNamingPolicyPrefixEnabled"; break }
                    "chkBookingsNamingPolicySuffixEnabled" { $propertyName = "BookingsNamingPolicySuffixEnabled"; break }
                    default { $propertyName = $derivedPropertyName }
                }
                
                # Spezielle CheckBox-Behandlung für Steuerelemente, die andere UI-Elemente aktivieren/deaktivieren
                if ($elementType -eq "CheckBox") {
                    switch ($elementName) {
                        "chkMailTipsLargeAudienceThreshold" { $propertyName = "MailTipsLargeAudienceThreshold"; break }
                        "chkPreferredInternetCodePageForShiftJis" { $propertyName = "PreferredInternetCodePageForShiftJis"; break }
                        "chkSearchQueryLanguage" { $propertyName = "SearchQueryLanguage"; break }
                        "chkMailTipsAllTipsEnabled" { $propertyName = "MailTipsAllTipsEnabled"; break }
                        # Fehlende MailTips-CheckBoxen
                        "chkMailTipsExternalRecipientsTipsEnabled" { $propertyName = "MailTipsExternalRecipientsTipsEnabled"; break }
                        "chkMailTipsGroupMetricsEnabled" { $propertyName = "MailTipsGroupMetricsEnabled"; break }
                        "chkMailTipsMailboxSourcedTipsEnabled" { $propertyName = "MailTipsMailboxSourcedTipsEnabled"; break }
                        # Fehlende Sicherheits-CheckBoxen
                        "chkMaskClientIpInReceivedHeadersEnabled" { $propertyName = "MaskClientIpInReceivedHeadersEnabled"; break }
                        "chkComplianceMLBgdCrawlEnabled" { $propertyName = "ComplianceMLBgdCrawlEnabled"; break }
                        # Fehlende OAuth/Auth-CheckBoxen
                        "chkOAuth2ClientProfileEnabled" { $propertyName = "OAuth2ClientProfileEnabled"; break }
                        "chkPerTenantSwitchToESTSEnabled" { $propertyName = "PerTenantSwitchToESTSEnabled"; break }
                        "chkRefreshSessionEnabled" { $propertyName = "RefreshSessionEnabled"; break }
                        # Fehlende EWS-CheckBoxen
                        "chkEwsEnabled" { $propertyName = "EwsEnabled"; break }
                        "chkEwsAllowEntourage" { $propertyName = "EwsAllowEntourage"; break }
                        "chkEwsAllowMacOutlook" { $propertyName = "EwsAllowMacOutlook"; break }
                        "chkEwsAllowOutlook" { $propertyName = "EwsAllowOutlook"; break }
                        # Fehlende Mobile-CheckBoxen
                        "chkMobileAppEducationEnabled" { $propertyName = "MobileAppEducationEnabled"; break }
                        "chkOutlookMobileGCCRestrictionsEnabled" { $propertyName = "OutlookMobileGCCRestrictionsEnabled"; break }
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
                                        $itemTagString = if ($itemObject -is [System.Windows.Controls.ComboBoxItem] -and $null -ne $itemObject.Tag) { $itemObject.Tag.ToString() } else { $null }
                                        
                                        # Prüfe sowohl Content als auch Tag auf Übereinstimmung
                                        if ($itemContentString.Equals($configValueAsString, [System.StringComparison]::OrdinalIgnoreCase) -or 
                                            ($null -ne $itemTagString -and $itemTagString.Equals($configValueAsString, [System.StringComparison]::OrdinalIgnoreCase))) { 
                                            $uiSelectedItem = $itemObject; $itemFound = $true; break 
                                        }
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

function Get-ExoUnifiedGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Identity
    )

    $logIdentity = if ([string]::IsNullOrWhiteSpace($Identity)) { "*" } else { $Identity }
    Write-Log "Get-ExoUnifiedGroup: Aktion gestartet für Identity '$logIdentity'." -Type Info
    Update-StatusBar -Message "Rufe Microsoft 365-Gruppe(n) '$logIdentity' ab..." -Type Info

    try {
        if (-not $script:isConnected) {
            Show-MessageBox -Message "Sie sind nicht mit Exchange Online verbunden." -Title "Nicht verbunden" -Type Warning
            Update-StatusBar -Message "Nicht verbunden." -Type Warning
            return $null
        }

        $params = @{ ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($Identity)) {
            $params['Identity'] = $Identity
        }

        # Der eigentliche EXO-Cmdlet-Aufruf
        $groups = Get-UnifiedGroup @params
        
        if ($groups) {
            $count = if ($groups.Count) { $groups.Count } else { 1 }
            Write-Log "Get-ExoUnifiedGroup: $count Microsoft 365-Gruppe(n) für Identity '$logIdentity' erfolgreich abgerufen." -Type Success
            Update-StatusBar -Message "$count Microsoft 365-Gruppe(n) erfolgreich abgerufen." -Type Success
            return $groups
        } else {
            Write-Log "Get-ExoUnifiedGroup: Keine Microsoft 365-Gruppe(n) für Identity '$logIdentity' gefunden." -Type Warning
            Update-StatusBar -Message "Keine Microsoft 365-Gruppe(n) für '$logIdentity' gefunden." -Type Warning
            return @() # Leeres Array zurückgeben, wenn nichts gefunden wird
        }

    } catch {
        $errMsg = $_.Exception.Message
        Write-Log "Get-ExoUnifiedGroup: Fehler beim Abrufen der Microsoft 365-Gruppe(n) '$logIdentity': $errMsg" -Type Error
        Update-StatusBar -Message "Fehler beim Abrufen der Microsoft 365-Gruppe(n) '$logIdentity'." -Type Error
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
            $rawExoObject = Get-ExoUnifiedGroup@paramsForExoCall -ErrorAction SilentlyContinue # Fehler hier abfangen
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
                        ErrorAction = 'Stop'
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
                        ErrorAction = 'Stop'
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
                    # Notes ist für UnifiedGroup nicht gültig
                    if ($params.ContainsKey("Notes")) { $params.Remove("Notes") | Out-Null }

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
                    # New-UnifiedGroup erlaubt unter bestimmten Berechtigungen keinen -ErrorAction Parameter.
                    # Daher wird der Aufruf gekapselt, um Fehler explizit abzufangen.
                    try {
                        $newGroup = New-UnifiedGroup @params -ErrorAction Stop
                    }
                    catch {
                        # Wirft den Fehler erneut, damit er vom äußeren Catch-Block behandelt wird.
                        throw $_
                    }
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
    param() # Parameter wird nicht mehr benötigt, da wir direkt auf die ComboBox zugreifen

    # Direkter Zugriff auf das ausgewählte Element der ComboBox
    $selectedItem = $null
    if ($null -ne $script:cmbSelectExistingGroup) {
        $selectedItem = $script:cmbSelectExistingGroup.SelectedItem
    }

    # Debugging: Log the selected item details
    if ($null -ne $selectedItem) {
        $tagType = if ($selectedItem.Tag) { $selectedItem.Tag.GetType().Name } else { 'null' }
        Write-Log "Get-ExoGroupMembersAction: SelectedItem gefunden. Content: '$($selectedItem.Content)', Tag Type: $tagType" -Type Debug
        if ($null -ne $selectedItem.Tag) {
            Write-Log "Get-ExoGroupMembersAction: Tag Properties: $($selectedItem.Tag.PSObject.Properties.Name -join ', ')" -Type Debug
            if ($selectedItem.Tag.PSObject.Properties["Identity"]) {
                Write-Log "Get-ExoGroupMembersAction: Identity in Tag: '$($selectedItem.Tag.Identity)'" -Type Debug
            }
            if ($selectedItem.Tag.PSObject.Properties["OriginalObject"]) {
                Write-Log "Get-ExoGroupMembersAction: OriginalObject vorhanden, Type: $($selectedItem.Tag.OriginalObject.GetType().Name)" -Type Debug
                if ($selectedItem.Tag.OriginalObject.PSObject.Properties["Identity"]) {
                    Write-Log "Get-ExoGroupMembersAction: OriginalObject.Identity: '$($selectedItem.Tag.OriginalObject.Identity)'" -Type Debug
                }
            }
        }
    } else {
        Write-Log "Get-ExoGroupMembersAction: Kein SelectedItem gefunden." -Type Warning
        Update-StatusBar -Message "Keine Gruppe ausgewählt." -Type Warning
        return
    }

    # Überprüfen, ob das ausgewählte Element und sein Tag gültig sind
    # Angepasste Logik basierend auf der Refresh-ExistingGroupsDropdown Funktion
    $groupObject = $null
    $isValidSelection = $false

    if ($null -ne $selectedItem -and $null -ne $selectedItem.Tag) {
        # Das Tag sollte das benutzerdefinierte Objekt aus Get-AllGroupTypesAction enthalten
        $taggedObject = $selectedItem.Tag
        
        # Prüfen ob es ein gültiges Gruppenobjekt ist
        if ($taggedObject.PSObject.Properties["Identity"] -and 
            $taggedObject.PSObject.Properties["DisplayName"] -and 
            $taggedObject.PSObject.Properties["RecipientTypeDetails"]) {
            
            $groupObject = $taggedObject
            $isValidSelection = $true
            Write-Log "Get-ExoGroupMembersAction: Gültiges Gruppenobjekt im Tag gefunden." -Type Debug
        }
        # Fallback: Prüfen ob OriginalObject vorhanden ist (falls die Struktur anders ist)
        elseif ($taggedObject.PSObject.Properties["OriginalObject"]) {
            $originalObject = $taggedObject.OriginalObject
            if ($originalObject.PSObject.Properties["Identity"] -and 
                $originalObject.PSObject.Properties["DisplayName"] -and 
                $originalObject.PSObject.Properties["RecipientTypeDetails"]) {
                
                $groupObject = $originalObject
                $isValidSelection = $true
                Write-Log "Get-ExoGroupMembersAction: Gültiges Gruppenobjekt im OriginalObject gefunden." -Type Debug
            }
        }
    }

    if (-not $isValidSelection -or $null -eq $groupObject) {
        Write-Log "Get-ExoGroupMembersAction: Ungültiges Gruppenobjekt oder fehlende Identity im Tag des ausgewählten Elements. Aktion wird abgebrochen." -Type Warning
        Update-StatusBar -Message "Ungültige Gruppenauswahl." -Type Warning
        return
    }

    # Variablen für die Hauptlogik definieren
    $groupDisplayName = $groupObject.DisplayName
    $recipientType = $groupObject.RecipientTypeDetails
    # Verwende die DistinguishedName oder andere verfügbare Identifier
    $effectiveIdentityForExo = $null
    if ($groupObject.PSObject.Properties["DistinguishedName"] -and ![string]::IsNullOrEmpty($groupObject.DistinguishedName)) {
        $effectiveIdentityForExo = $groupObject.DistinguishedName
    } elseif ($groupObject.PSObject.Properties["Identity"] -and ![string]::IsNullOrEmpty($groupObject.Identity)) {
        $effectiveIdentityForExo = $groupObject.Identity
    } elseif ($groupObject.PSObject.Properties["PrimarySmtpAddress"] -and ![string]::IsNullOrEmpty($groupObject.PrimarySmtpAddress)) {
        $effectiveIdentityForExo = $groupObject.PrimarySmtpAddress
    } else {
        Write-Log "Get-ExoGroupMembersAction: Keine gültige Identity für EXO-Cmdlets gefunden." -Type Error
        Update-StatusBar -Message "Keine gültige Gruppen-Identity gefunden." -Type Error
        return
    }
    
    $members = @()
    $membersSuccessfullyRetrieved = $false

    Write-Log "Get-ExoGroupMembersAction: Verarbeite Gruppe '$groupDisplayName' (Identity: '$effectiveIdentityForExo', Typ: '$recipientType')" -Type Info
    Update-StatusBar -Message "Lade Mitglieder für '$groupDisplayName'..." -Type Info

    try {
        # Prüfen, ob wir mit Exchange Online verbunden sind
        if (-not $script:isConnected) {
            Write-Log "Get-ExoGroupMembersAction: Keine Verbindung zu Exchange Online." -Type Warning
            Update-StatusBar -Message "Keine Verbindung zu Exchange Online." -Type Warning
            return
        }

        # Gruppenmitglieder laden basierend auf dem Gruppentyp
        try {
            if ($recipientType -eq "GroupMailbox") {
                # Für Microsoft 365 Gruppen - verwende Get-UnifiedGroupLinks
                try {
                    $members = Get-UnifiedGroupLinks -Identity $effectiveIdentityForExo -LinkType Members -ErrorAction Stop
                    Write-Log "Get-ExoGroupMembersAction: Verwende Get-UnifiedGroupLinks für Microsoft 365 Gruppe." -Type Debug
                } catch {
                    # Fallback: Verwende alternative Identity falls erste fehlschlägt
                    if ($groupObject.PSObject.Properties["PrimarySmtpAddress"] -and $effectiveIdentityForExo -ne $groupObject.PrimarySmtpAddress) {
                        Write-Log "Get-ExoGroupMembersAction: Fallback zu PrimarySmtpAddress für UnifiedGroupLinks." -Type Debug
                        $members = Get-UnifiedGroupLinks -Identity $groupObject.PrimarySmtpAddress -LinkType Members -ErrorAction Stop
                    } else {
                        throw $_
                    }
                }
            } elseif ($recipientType -in ("MailUniversalDistributionGroup", "MailUniversalSecurityGroup", "MailNonUniversalGroup")) {
                # Für Verteilergruppen - verwende Get-DistributionGroupMember
                try {
                    $members = Get-DistributionGroupMember -Identity $effectiveIdentityForExo -ErrorAction Stop
                    Write-Log "Get-ExoGroupMembersAction: Verwende Get-DistributionGroupMember für Verteilergruppe." -Type Debug
                } catch {
                    # Fallback: Verwende alternative Identity falls erste fehlschlägt
                    if ($groupObject.PSObject.Properties["PrimarySmtpAddress"] -and $effectiveIdentityForExo -ne $groupObject.PrimarySmtpAddress) {
                        Write-Log "Get-ExoGroupMembersAction: Fallback zu PrimarySmtpAddress für DistributionGroupMember." -Type Debug
                        $members = Get-DistributionGroupMember -Identity $groupObject.PrimarySmtpAddress -ErrorAction Stop
                    } else {
                        throw $_
                    }
                }
            } elseif ($recipientType -eq "DynamicDistributionGroup") {
                # Dynamische Verteilergruppen haben keine festen Mitglieder
                $members = @()
                Write-Log "Get-ExoGroupMembersAction: Dynamische Verteilergruppe - keine festen Mitglieder." -Type Info
            } else {
                Write-Log "Get-ExoGroupMembersAction: Unbekannter oder nicht unterstützter Gruppentyp '$recipientType' für Mitgliederabruf." -Type Warning
                Show-MessageBox -Message "Der Gruppentyp '$recipientType' wird für den Mitgliederabruf aktuell nicht vollständig unterstützt." -Title "Hinweis" -Type Info
            }
            
            if ($null -ne $members) { 
                $membersSuccessfullyRetrieved = $true
                Write-Log "Get-ExoGroupMembersAction: $($members.Count) Mitglieder-Objekte für '$groupDisplayName' abgerufen." -Type Debug
            } else {
                Write-Log "Get-ExoGroupMembersAction: Keine Mitglieder-Objekte für '$groupDisplayName' zurückgegeben (Cmdlet lieferte null oder Fehler)." -Type Debug
                $members = @() 
            }
        } catch {
            $errorMsg = $_.Exception.Message
            Write-Log "Get-ExoGroupMembersAction: Fehler beim Abrufen der Mitglieder für '$groupDisplayName': $errorMsg" -Type Error
            Show-MessageBox -Message "Fehler beim Abrufen der Gruppenmitglieder für '$groupDisplayName': $errorMsg" -Title "Fehler Mitgliederabruf" -Type Error
        }

        # UI aktualisieren - Mitgliederliste
        if ($null -ne $script:lstGroupMembers) {
            $script:lstGroupMembers.Items.Clear()
            $script:lstGroupMembers.Tag = $groupObject # Store the group object for other actions

            if ($membersSuccessfullyRetrieved -and $members.Count -gt 0) {
                foreach ($member in $members) {
                    $item = New-Object System.Windows.Controls.ListViewItem
                    $item.Content = $member.DisplayName
                    $item.Tag = $member # Store the full member object
                    $script:lstGroupMembers.Items.Add($item) | Out-Null
                }
            } elseif ($recipientType -eq "DynamicDistributionGroup") {
                $item = New-Object System.Windows.Controls.ListViewItem
                $item.Content = "Dynamische Gruppen haben keine festen Mitglieder."
                $item.IsEnabled = $false
                $script:lstGroupMembers.Items.Add($item) | Out-Null
            } else {
                $item = New-Object System.Windows.Controls.ListViewItem
                $item.Content = "Keine Mitglieder gefunden."
                $item.IsEnabled = $false
                $script:lstGroupMembers.Items.Add($item) | Out-Null
            }
        }

        # UI aktualisieren - Gruppeneinstellungen (Checkboxes)
        try {
            # Enable all controls first
            $script:chkHiddenFromGAL.IsEnabled = $true
            $script:chkRequireSenderAuth.IsEnabled = $true
            $script:chkAllowExternalSenders.IsEnabled = $true
            
            $groupDetailsForSettings = $null
            if ($recipientType -eq "GroupMailbox") {
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
        Write-Log "Get-ExoGroupMembersAction: Schwerwiegender Fehler für Gruppe '$groupDisplayName': $errorMsg `n$fullError" -Type Error
        Show-MessageBox -Message "Ein schwerwiegender Fehler ist aufgetreten beim Laden der Gruppeninformationen für '$groupDisplayName': $errorMsg" -Title "Schwerer Fehler" -Type Error
        Update-StatusBar -Message "Fehler beim Laden der Gruppeninformationen für '$groupDisplayName'." -Type Error

        try {
            $script:chkHiddenFromGAL.IsChecked = $false; $script:chkHiddenFromGAL.IsEnabled = $false
            $script:chkRequireSenderAuth.IsChecked = $false; $script:chkRequireSenderAuth.IsEnabled = $false
            $script:chkAllowExternalSenders.IsChecked = $false; $script:chkAllowExternalSenders.IsEnabled = $false
            if ($null -ne $script:lstGroupMembers) {
                $script:lstGroupMembers.Items.Clear()
                $script:lstGroupMembers.Tag = $null
                $item = New-Object System.Windows.Controls.ListViewItem
                $item.Content = "Fehler beim Laden der Gruppeninformationen."
                $item.IsEnabled = $false
                $script:lstGroupMembers.Items.Add($item) | Out-Null
            }
        } catch {
            Write-Log "Get-ExoGroupMembersAction: Konnte UI nach schwerem Fehler nicht zurücksetzen." -Type Error
        }
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

##############TABSFUNKTIONEN####################

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

    function Set-ActiveNavigationButton {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ButtonName
    )
    
    try {
        Write-Log "Setze aktiven Navigation-Button: $ButtonName" -Type "Debug"
        
        # Liste aller Navigation-Buttons
        $navigationButtons = @(
            'btnDashboard',
            'btnNavCalendar', 'btnNavMailbox', 'btnNavSharedMailbox', 'btnNavGroups', 'btnNavResources', 'btnNavContacts',
            'btnNavMailFlowRules', 'btnNavInboxRules', 'btnNavMessageTrace', 'btnNavAutoReply',
            'btnNavATP', 'btnNavDLP', 'btnNavEDiscovery', 'btnNavMDM',
            'btnNavEXOSettings', 'btnNavRegion', 'btnNavCrossPremises',
            'btnNavHybridExchange', 'btnNavMultiForest',
            'btnNavHealthCheck', 'btnNavAudit', 'btnNavReports', 'btnNavTroubleshooting'
        )
        
        # Alle Navigation-Buttons auf Standard-Style zurücksetzen
        foreach ($btnName in $navigationButtons) {
            $button = $script:Form.FindName($btnName)
            if ($null -ne $button) {
                try {
                    # Standard-Button-Style anwenden
                    $button.Dispatcher.Invoke([Action]{
                        $button.Background = [System.Windows.Media.Brushes]::LightGray
                        $button.Foreground = [System.Windows.Media.Brushes]::Black
                        $button.FontWeight = "Normal"
                    }, "Normal")
                }
                catch {
                    Write-Log "Fehler beim Zurücksetzen von Button $btnName - $($_.Exception.Message)" -Type "Warning"
                }
            }
        }
        
        # Aktiven Button hervorheben
        $activeButton = $script:Form.FindName($ButtonName)
        if ($null -ne $activeButton) {
            try {
                $activeButton.Dispatcher.Invoke([Action]{
                    $activeButton.Background = [System.Windows.Media.Brushes]::DodgerBlue
                    $activeButton.Foreground = [System.Windows.Media.Brushes]::White
                    $activeButton.FontWeight = "Bold"
                }, "Normal")
                
                Write-Log "Navigation-Button '$ButtonName' als aktiv markiert" -Type "Debug"
            }
            catch {
                Write-Log "Fehler beim Hervorheben von Button $ButtonName - $($_.Exception.Message)" -Type "Warning"
            }
        }
        else {
            Write-Log "Aktiver Navigation-Button '$ButtonName' nicht gefunden" -Type "Warning"
        }
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler in Set-ActiveNavigationButton: $errorMsg" -Type "Error"
        return $false
    }
}

#region EXOSettings Tab Initialization
function Initialize-EXOSettingsTab {
    [CmdletBinding()]
    param()

    Write-Log "Beginne Initialisierung: EXO Settings Tab" -Type "Info"
    [bool]$success = $true

    try {
        # Initialisiere den Status-Flag für die Event-Handler
        $script:EXOSettingsHandlersRegistered = $false

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
                # Die Funktion Get-CurrentOrganizationConfig prüft die Verbindung intern und registriert die weiteren Handler
                Get-CurrentOrganizationConfig
            })
            Write-Log "EXOSettingsTab: btnGetOrganizationConfig Handler registriert." -Type "Debug"
        } else { Write-Log "EXOSettingsTab: btnGetOrganizationConfig nicht gefunden." -Type "Warning"; $success = $false }

        # Die Event-Handler für Speichern und Exportieren werden jetzt in Get-CurrentOrganizationConfig registriert,
        # nachdem die Daten erfolgreich geladen wurden.

        foreach ($elementName in $script:knownUIElements) {
            $element = Get-XamlElement -ElementName $elementName
            if ($null -ne $element) {
                $element.Visibility = [System.Windows.Visibility]::Visible
            } else {
                Write-Log "EXOSettingsTab: Element '$elementName' aus knownUIElements nicht gefunden in XAML." -Type "Warning"
            }
        }
        Write-Log "EXOSettingsTab: Sichtbarkeit für Elemente in knownUIElements überprüft." -Type "Debug"

        # Stelle sicher, dass der Tab selbst sichtbar ist
        if ($null -ne $script:tabEXOSettings) {
            $script:tabEXOSettings.Visibility = [System.Windows.Visibility]::Visible
            Write-Log "EXOSettingsTab: Tab auf sichtbar gesetzt." -Type "Debug"
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

                # Debug-Ausgabe für Fehlerbehebung
                Write-Log "SelectionChanged: SelectedItem Content: '$($selectedGroupItem.Content)'" -Type Debug
                if ($null -ne $selectedGroupItem.Tag) {
                    Write-Log "SelectionChanged: Tag Type: $($selectedGroupItem.Tag.GetType().Name)" -Type Debug
                    Write-Log "SelectionChanged: Tag Properties: $($selectedGroupItem.Tag.PSObject.Properties.Name -join ', ')" -Type Debug
                } else {
                    Write-Log "SelectionChanged: Tag ist null" -Type Debug
                }

                # Korrigierte Validierung: DistinguishedName statt Identity verwenden
                if ($null -eq $groupObject -or 
                    -not $groupObject.PSObject.Properties["DistinguishedName"] -or 
                    [string]::IsNullOrWhiteSpace($groupObject.DistinguishedName)) {
                    Write-Log "SelectionChanged: Ungültiges Gruppenobjekt oder fehlende DistinguishedName im Tag des ausgewählten Elements." -Type Warning
                    if ($null -ne $script:lstGroupMembers) { $script:lstGroupMembers.Items.Clear(); $script:lstGroupMembers.Tag = $null }
                    if ($null -ne $script:chkHiddenFromGAL) { $script:chkHiddenFromGAL.IsChecked = $false; $script:chkHiddenFromGAL.IsEnabled = $false }
                    if ($null -ne $script:chkRequireSenderAuth) { $script:chkRequireSenderAuth.IsChecked = $false; $script:chkRequireSenderAuth.IsEnabled = $false }
                    if ($null -ne $script:chkAllowExternalSenders) { $script:chkAllowExternalSenders.IsChecked = $false; $script:chkAllowExternalSenders.IsEnabled = $false }
                    if ($null -ne $script:txtGroupUser) { $script:txtGroupUser.Text = "" }
                    Update-StatusBar -Message "Ungültige Gruppenauswahl." -Type Warning
                    return
                }
                
                # DistinguishedName als Identity verwenden
                $effectiveIdentity = $groupObject.DistinguishedName
                Write-Log "Gruppe '$($selectedGroupItem.Content)' (DistinguishedName: $effectiveIdentity) in ComboBox ausgewählt. Lade Details..." -Type "Info"
                Get-ExoGroupMembersAction

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
        Write-Log "Initialisiere Ressourcen-Tab..." -Type "Info"

        # UI-Elemente für Ressourcen referenzieren basierend auf dem neuen XAML
        $helpLinkResources = Get-XamlElement -ElementName "helpLinkResources"
        $cmbResourceType = Get-XamlElement -ElementName "cmbResourceType"
        $txtResourceName = Get-XamlElement -ElementName "txtResourceName"
        $btnCreateResource = Get-XamlElement -ElementName "btnCreateResource"
        $txtResourceSearch = Get-XamlElement -ElementName "txtResourceSearch"
        $btnSearchResources = Get-XamlElement -ElementName "btnSearchResources"
        $btnRefreshResources = Get-XamlElement -ElementName "btnRefreshResources"
        $btnShowRoomResources = Get-XamlElement -ElementName "btnShowRoomResources"
        $btnShowEquipmentResources = Get-XamlElement -ElementName "btnShowEquipmentResources"
        $cmbResourceSelect = Get-XamlElement -ElementName "cmbResourceSelect"
        $btnRefreshResourceList = Get-XamlElement -ElementName "btnRefreshResourceList"
        $btnEditResourceSettings = Get-XamlElement -ElementName "btnEditResourceSettings"
        $btnRemoveResource = Get-XamlElement -ElementName "btnRemoveResource"
        $dgResources = Get-XamlElement -ElementName "dgResources"
        $btnExportResources = Get-XamlElement -ElementName "btnExportResources"

        # Globale Script-Variablen für die UI-Elemente setzen
        $script:cmbResourceType = $cmbResourceType
        $script:txtResourceName = $txtResourceName
        $script:txtResourceSearch = $txtResourceSearch
        $script:cmbResourceSelect = $cmbResourceSelect
        $script:dgResources = $dgResources

        # Event-Handler für "Alle anzeigen"
        Register-EventHandler -Control $btnRefreshResources -Handler {
            try {
            if (-not $script:isConnected) { 
                Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                Update-StatusBar -Message "Keine Verbindung zu Exchange Online." -Type Warning
                return
            }
            
            Update-StatusBar -Message "Lade alle Ressourcen..." -Type Info
            $allResources = Get-AllResourcesAction
            
            if ($null -eq $allResources) {
                Write-Log "Get-AllResourcesAction gab null zurück." -Type Warning
                $allResources = @()
            }
            
            # Thread-sichere UI-Aktualisierung
            $script:dgResources.Dispatcher.Invoke([Action]{
                $script:dgResources.ItemsSource = $allResources
            }, "Normal")
            
            $script:cmbResourceSelect.Dispatcher.Invoke([Action]{
                $script:cmbResourceSelect.Items.Clear()
                if ($allResources.Count -gt 0) {
                foreach ($resource in $allResources) {
                    if ($null -ne $resource.PrimarySmtpAddress) {
                    [void]$script:cmbResourceSelect.Items.Add($resource.PrimarySmtpAddress)
                    }
                }
                if ($script:cmbResourceSelect.Items.Count -gt 0) {
                    $script:cmbResourceSelect.SelectedIndex = 0
                }
                }
            }, "Normal")
            
            Update-StatusBar -Message "$($allResources.Count) Ressource(n) geladen." -Type Success
            Write-Log "Alle Ressourcen erfolgreich geladen: $($allResources.Count) Einträge" -Type Success
            
            } catch {
            $errorMsg = $_.Exception.Message
            Write-Log "Fehler beim Laden aller Ressourcen: $errorMsg" -Type Error
            Update-StatusBar -Message "Fehler beim Laden der Ressourcen: $errorMsg" -Type Error
            Show-MessageBox -Message "Fehler beim Laden der Ressourcen:`n$errorMsg" -Title "Fehler" -Type Error
            }
        } -ControlName "btnRefreshResources"

        # Event-Handler für "Nur Räume"
        Register-EventHandler -Control $btnShowRoomResources -Handler {
            try {
            if (-not $script:isConnected) { 
                Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                Update-StatusBar -Message "Keine Verbindung zu Exchange Online." -Type Warning
                return
            }
            
            Update-StatusBar -Message "Lade Raumressourcen..." -Type Info
            $rooms = Get-RoomResourcesAction
            
            if ($null -eq $rooms) {
                Write-Log "Get-RoomResourcesAction gab null zurück." -Type Warning
                $rooms = @()
            }
            
            # Thread-sichere UI-Aktualisierung
            $script:dgResources.Dispatcher.Invoke([Action]{
                $script:dgResources.ItemsSource = $rooms
            }, "Normal")
            
            $script:cmbResourceSelect.Dispatcher.Invoke([Action]{
                $script:cmbResourceSelect.Items.Clear()
                if ($rooms.Count -gt 0) {
                foreach ($room in $rooms) {
                    if ($null -ne $room.PrimarySmtpAddress) {
                    [void]$script:cmbResourceSelect.Items.Add($room.PrimarySmtpAddress)
                    }
                }
                if ($script:cmbResourceSelect.Items.Count -gt 0) {
                    $script:cmbResourceSelect.SelectedIndex = 0
                }
                }
            }, "Normal")
            
            Update-StatusBar -Message "$($rooms.Count) Raumressource(n) geladen." -Type Success
            Write-Log "Raumressourcen erfolgreich geladen: $($rooms.Count) Einträge" -Type Success
            
            } catch {
            $errorMsg = $_.Exception.Message
            Write-Log "Fehler beim Laden der Raumressourcen: $errorMsg" -Type Error
            Update-StatusBar -Message "Fehler beim Laden der Raumressourcen: $errorMsg" -Type Error
            Show-MessageBox -Message "Fehler beim Laden der Raumressourcen:`n$errorMsg" -Title "Fehler" -Type Error
            }
        } -ControlName "btnShowRoomResources"

        # Event-Handler für "Nur Ausstattung"
        Register-EventHandler -Control $btnShowEquipmentResources -Handler {
            try {
            if (-not $script:isConnected) { 
                Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                Update-StatusBar -Message "Keine Verbindung zu Exchange Online." -Type Warning
                return
            }
            
            Update-StatusBar -Message "Lade Ausstattungsressourcen..." -Type Info
            $equipment = Get-EquipmentResourcesAction
            
            if ($null -eq $equipment) {
                Write-Log "Get-EquipmentResourcesAction gab null zurück." -Type Warning
                $equipment = @()
            }
            
            # Sicherstellen, dass $equipment immer ein Array ist
            if ($equipment -isnot [array]) {
                $equipment = @($equipment)
            }
            
            # Thread-sichere UI-Aktualisierung
            $script:dgResources.Dispatcher.Invoke([Action]{
                $script:dgResources.ItemsSource = $equipment
            }, "Normal")
            
            $script:cmbResourceSelect.Dispatcher.Invoke([Action]{
                $script:cmbResourceSelect.Items.Clear()
                if ($equipment.Count -gt 0) {
                foreach ($item in $equipment) {
                    if ($null -ne $item.PrimarySmtpAddress) {
                    [void]$script:cmbResourceSelect.Items.Add($item.PrimarySmtpAddress)
                    }
                }
                if ($script:cmbResourceSelect.Items.Count -gt 0) {
                    $script:cmbResourceSelect.SelectedIndex = 0
                }
                }
            }, "Normal")
            
            Update-StatusBar -Message "$($equipment.Count) Ausstattungsressource(n) geladen." -Type Success
            Write-Log "Ausstattungsressourcen erfolgreich geladen: $($equipment.Count) Einträge" -Type Success
            
            } catch {
            $errorMsg = $_.Exception.Message
            Write-Log "Fehler beim Laden der Ausstattungsressourcen: $errorMsg" -Type Error
            Update-StatusBar -Message "Fehler beim Laden der Ausstattungsressourcen: $errorMsg" -Type Error
            Show-MessageBox -Message "Fehler beim Laden der Ausstattungsressourcen:`n$errorMsg" -Title "Fehler" -Type Error
            }
        } -ControlName "btnShowEquipmentResources"

        # Event-Handler für "Suchen"
        Register-EventHandler -Control $btnSearchResources -Handler {
            try {
            if (-not $script:isConnected) { 
                Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                Update-StatusBar -Message "Keine Verbindung zu Exchange Online." -Type Warning
                return
            }
            
            $searchTerm = $script:txtResourceSearch.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($searchTerm)) { 
                Show-MessageBox -Message "Bitte geben Sie einen Suchbegriff ein." -Title "Suchbegriff fehlt" -Type Warning
                Update-StatusBar -Message "Kein Suchbegriff eingegeben." -Type Warning
                return
            }
            
            # Sicherheitsprüfung: Mindestlänge des Suchbegriffs
            if ($searchTerm.Length -lt 2) {
                Show-MessageBox -Message "Der Suchbegriff muss mindestens 2 Zeichen lang sein." -Title "Suchbegriff zu kurz" -Type Warning
                Update-StatusBar -Message "Suchbegriff zu kurz (mindestens 2 Zeichen)." -Type Warning
                return
            }
            
            Update-StatusBar -Message "Suche nach Ressourcen mit '$searchTerm'..." -Type Info
            $foundResources = Search-ResourcesAction -SearchTerm $searchTerm
            
            if ($null -eq $foundResources) {
                Write-Log "Search-ResourcesAction gab null zurück." -Type Warning
                $foundResources = @()
            }
            
            # Sicherstellen, dass $foundResources immer ein Array ist
            if ($foundResources -isnot [array]) {
                $foundResources = @($foundResources)
            }
            
            # Thread-sichere UI-Aktualisierung
            $script:dgResources.Dispatcher.Invoke([Action]{
                $script:dgResources.ItemsSource = $foundResources
            }, "Normal")
            
            $script:cmbResourceSelect.Dispatcher.Invoke([Action]{
                $script:cmbResourceSelect.Items.Clear()
                if ($foundResources.Count -gt 0) {
                foreach ($resource in $foundResources) {
                    if ($null -ne $resource.PrimarySmtpAddress) {
                    [void]$script:cmbResourceSelect.Items.Add($resource.PrimarySmtpAddress)
                    }
                }
                if ($script:cmbResourceSelect.Items.Count -gt 0) {
                    $script:cmbResourceSelect.SelectedIndex = 0
                }
                }
            }, "Normal")
            
            $resultMessage = if ($foundResources.Count -eq 0) {
                "Keine Ressourcen mit '$searchTerm' gefunden."
            } else {
                "$($foundResources.Count) Ressource(n) mit '$searchTerm' gefunden."
            }
            
            Update-StatusBar -Message $resultMessage -Type Success
            Write-Log "Ressourcensuche abgeschlossen: $resultMessage" -Type Success
            
            # Zusätzliche Information bei keinen Ergebnissen
            if ($foundResources.Count -eq 0) {
                Show-MessageBox -Message "Es wurden keine Ressourcen gefunden, die '$searchTerm' enthalten. Versuchen Sie einen anderen Suchbegriff." -Title "Keine Ergebnisse" -Type Info
            }
            
            } catch {
            $errorMsg = $_.Exception.Message
            Write-Log "Fehler bei der Ressourcensuche: $errorMsg" -Type Error
            Update-StatusBar -Message "Fehler bei der Ressourcensuche: $errorMsg" -Type Error
            Show-MessageBox -Message "Fehler bei der Ressourcensuche:`n$errorMsg" -Title "Suchfehler" -Type Error
            }
        } -ControlName "btnSearchResources"


        # Event-Handler für "Ressourcenliste aktualisieren" (ComboBox)
        Register-EventHandler -Control $btnRefreshResourceList -Handler {
            try {
                if (-not $script:isConnected) { Throw "Keine Verbindung zu Exchange Online." }
                Update-StatusBar -Message "Aktualisiere Ressourcenliste..."
                $allResources = Get-AllResourcesAction
                $script:cmbResourceSelect.ItemsSource = $allResources | ForEach-Object { $_.PrimarySmtpAddress }
                if ($script:cmbResourceSelect.Items.Count -gt 0) {
                    $script:cmbResourceSelect.SelectedIndex = 0
                }
                Update-StatusBar -Message "Ressourcenliste aktualisiert." -Type "Success"
            } catch { Update-StatusBar -Message "Fehler: $($_.Exception.Message)" -Type "Error" }
        } -ControlName "btnRefreshResourceList"

        # Event-Handler für "Ressource erstellen"
        Register-EventHandler -Control $btnCreateResource -Handler {
            try {
                if (-not $script:isConnected) { Throw "Keine Verbindung zu Exchange Online." }
                $name = $script:txtResourceName.Text
                $type = if ($script:cmbResourceType.SelectedItem.Content -eq "Raum") { "Room" } else { "Equipment" }

                if ([string]::IsNullOrWhiteSpace($name)) {
                    Throw "Name der Ressource darf nicht leer sein."
                }
                
                $result = New-ResourceAction -Name $name -DisplayName $name -ResourceType $type
                
                if ($result) { 
                    Update-StatusBar -Message "Ressource '$name' erfolgreich erstellt." -Type "Success"
                    $script:txtResourceName.Text = ""
                    # Optional: Ressourcenliste aktualisieren
                    $btnRefreshResources.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                }
            } catch {
                Update-StatusBar -Message "Fehler: $($_.Exception.Message)" -Type "Error"
            }
        } -ControlName "btnCreateResource"

        # Event-Handler für "Einstellungen bearbeiten"
        Register-EventHandler -Control $btnEditResourceSettings -Handler {
            try {
                if (-not $script:isConnected) { Throw "Keine Verbindung zu Exchange Online." }
                $selectedResource = $script:cmbResourceSelect.SelectedItem
                if ($null -eq $selectedResource) { Throw "Bitte wählen Sie eine Ressource aus." }
                
                $result = Show-ResourceSettingsDialog -Identity $selectedResource
                
                if ($result) {
                    Update-StatusBar -Message "Einstellungen für '$selectedResource' aktualisiert." -Type "Success"
                    # Ressourcenliste aktualisieren
                    $btnRefreshResources.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                } else {
                    Update-StatusBar -Message "Bearbeitung abgebrochen." -Type "Info"
                }
            } catch {
                Update-StatusBar -Message "Fehler: $($_.Exception.Message)" -Type "Error"
            }
        } -ControlName "btnEditResourceSettings"

        # Event-Handler für "Ressource löschen"
        Register-EventHandler -Control $btnRemoveResource -Handler {
            try {
                if (-not $script:isConnected) { Throw "Keine Verbindung zu Exchange Online." }
                $selectedResource = $script:cmbResourceSelect.SelectedItem
                if ($null -eq $selectedResource) { Throw "Bitte wählen Sie eine Ressource aus." }

                $confirm = [System.Windows.MessageBox]::Show("Möchten Sie die Ressource '$selectedResource' wirklich löschen?", "Bestätigung", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)

                if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
                    $result = Remove-ResourceAction -Identity $selectedResource
                    if ($result) { 
                        Update-StatusBar -Message "Ressource '$selectedResource' gelöscht." -Type "Success"
                        # Ressourcenliste aktualisieren
                        $btnRefreshResources.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                    }
                }
            } catch {
                Update-StatusBar -Message "Fehler: $($_.Exception.Message)" -Type "Error"
            }
        } -ControlName "btnRemoveResource"

        # Event-Handler für "Ressourcenliste exportieren"
        Register-EventHandler -Control $btnExportResources -Handler {
            try {
                if ($null -eq $script:dgResources.ItemsSource -or $script:dgResources.Items.Count -eq 0) {
                    Throw "Keine Ressourcen zum Exportieren vorhanden."
                }
                
                $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
                $saveFileDialog.Title = "Ressourcenliste exportieren"
                $saveFileDialog.FileName = "Ressourcenliste_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                
                if ($saveFileDialog.ShowDialog() -eq $true) {
                    Export-ResourcesAction -Resources $script:dgResources.ItemsSource -FilePath $saveFileDialog.FileName
                    Update-StatusBar -Message "Ressourcen erfolgreich exportiert." -Type "Success"
                }
            } catch {
                Update-StatusBar -Message "Fehler: $($_.Exception.Message)" -Type "Error"
            }
        } -ControlName "btnExportResources"

        # Event-Handler für Hilfe-Link
        if ($null -ne $helpLinkResources) {
            $helpLinkResources.Add_MouseLeftButtonDown({
                try { Show-HelpDialog -Topic "Resources" } catch {}
            })
        }

        Write-Log "Ressourcen-Tab erfolgreich initialisiert." -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Initialisieren des Ressourcen-Tabs: $errorMsg" -Type "Error"
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
        # HINZUGEFÜGT: Export-Button aus XAML
        $script:btnExportRegionSettings = $script:Form.FindName("btnExportRegionSettings")
        
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
            "btnExportRegionSettings" = $script:btnExportRegionSettings
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

        # ComboBoxen befüllen
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
        
        # Drei-Zustand-CheckBox auf unbestimmt setzen
        $script:chkRegionDefaultFolderNameMatchingUserLanguage.IsChecked = $null 
        Write-Log "chkRegionDefaultFolderNameMatchingUserLanguage.IsChecked auf \$null (unbestimmt) gesetzt." -Type Debug
        
        # Event-Handler für Sprachauswahl-Änderung
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
                Populate-TimeFormatComboBox -ComboBox $script:cmbRegionTimeFormat -CultureName $cultureNameForFormats
                Populate-TimezoneComboBox -ComboBox $script:cmbRegionTimezone -CultureName $cultureNameForTimezones
            } catch {
                $errorMsgSelChange = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler im SelectionChanged Event der Sprach-ComboBox."
                Write-Log "FEHLER im cmbRegionLanguage.SelectionChanged: $errorMsgSelChange" -Type Error
            }
        })
        Write-Log "Event-Handler für cmbRegionLanguage.SelectionChanged registriert (inkl. Datums-, Zeitformat- und Zeitzonen-Update)." -Type Debug
        
        # Event-Handler für "Einstellungen abrufen" Button
        if ($null -ne $script:btnGetRegionSettings) {
            $script:btnGetRegionSettings.Add_Click({
                try {
                    Invoke-GetRegionSettingsAction
                } catch {
                    $errorMsgAction = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Abrufen der Regionaleinstellungen."
                    Write-Log $errorMsgAction -Type Error
                    Log-Action "FEHLER bei Aktion 'Einstellungen abrufen' ($($script:btnGetRegionSettings.Name)): $errorMsgAction"
                    Show-MessageBox -Message "Ein Fehler ist beim Abrufen der Einstellungen aufgetreten: $($_.Exception.Message)" -Title "Aktionsfehler"
                }
            })
            Write-Log "Event-Handler für '$($script:btnGetRegionSettings.Name)' registriert." -Type Debug
        }

        # Event-Handler für "Einstellungen anwenden" Button
        if ($null -ne $script:btnSetRegionSettings) {
            $script:btnSetRegionSettings.Add_Click({
                try {
                    Invoke-SetRegionSettingsAction
                } catch {
                    $errorMsgAction = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Anwenden der Regionaleinstellungen."
                    Write-Log $errorMsgAction -Type Error
                    Log-Action "FEHLER bei Aktion 'Einstellungen anwenden' ($($script:btnSetRegionSettings.Name)): $errorMsgAction"
                    Show-MessageBox -Message "Ein Fehler ist beim Anwenden der Einstellungen aufgetreten: $($_.Exception.Message)" -Title "Aktionsfehler"
                }
            })
            Write-Log "Event-Handler für '$($script:btnSetRegionSettings.Name)' registriert." -Type Debug
        }

        # Event-Handler für Export-Button
        if ($null -ne $script:btnExportRegionSettings) {
            $script:btnExportRegionSettings.Add_Click({
                try {
                    # Prüfen ob Daten zum Exportieren vorhanden sind
                    if ([string]::IsNullOrWhiteSpace($script:txtRegionResult.Text)) {
                        Show-MessageBox -Message "Keine Daten zum Exportieren vorhanden. Bitte rufen Sie zuerst die Einstellungen ab." -Title "Keine Daten" -Type Warning
                        return
                    }

                    # SaveFileDialog für CSV-Export
                    $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
                    $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Textdateien (*.txt)|*.txt"
                    $saveFileDialog.Title = "Regionaleinstellungen exportieren"
                    $saveFileDialog.FileName = "Regionaleinstellungen_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                    
                    if ($saveFileDialog.ShowDialog() -eq $true) {
                        # Daten exportieren
                        $script:txtRegionResult.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
                        
                        Show-MessageBox -Message "Regionaleinstellungen wurden erfolgreich exportiert nach:`n$($saveFileDialog.FileName)" -Title "Export erfolgreich" -Type Info
                        Write-Log "Regionaleinstellungen exportiert nach: $($saveFileDialog.FileName)" -Type Success
                    }
                } catch {
                    $errorMsgExport = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Exportieren der Regionaleinstellungen."
                    Write-Log $errorMsgExport -Type Error
                    Show-MessageBox -Message "Fehler beim Exportieren: $($_.Exception.Message)" -Title "Exportfehler" -Type Error
                }
            })
            Write-Log "Event-Handler für '$($script:btnExportRegionSettings.Name)' registriert." -Type Debug
        }

        # Event-Handler für Hilfe-Link
        if ($null -ne $script:helpLinkRegionSettings) {
            $script:helpLinkRegionSettings.Add_MouseLeftButtonDown({ 
                try {
                    if (Test-Path Function:\Show-HelpDialog) {
                        Show-HelpDialog -Topic "RegionSettings"
                    } else {
                        Write-Log "Hilfefunktion Show-HelpDialog nicht gefunden." -Type Warning
                    }
                } catch {
                    Write-Log "Fehler beim Öffnen des Hilfe-Dialogs: $($_.Exception.Message)" -Type Error
                }
            }) 
            $script:helpLinkRegionSettings.Add_MouseEnter({ 
                $this.Cursor = [System.Windows.Input.Cursors]::Hand
                $this.TextDecorations = [System.Windows.TextDecorations]::Underline 
            })
            $script:helpLinkRegionSettings.Add_MouseLeave({ 
                $this.TextDecorations = $null
                $this.Cursor = [System.Windows.Input.Cursors]::Arrow 
            })
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
        Show-MessageBox -Message "Ein unerwarteter schwerwiegender Fehler ist beim Initialisieren des Regionaleinstellungen-Tabs aufgetreten: `n$($_.Exception.Message)`nBitte überprüfen Sie die Logs und die XAML-Definition." -Title "Schwerer Initialisierungsfehler" -Type Error
        return $false
    }
}

function Initialize-HealthCheckTab {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Initialisiere Health Check-Tab..." -Type "Info"
        
        # UI-Elemente referenzieren
        $btnRunHealthCheck = Get-XamlElement -ElementName "btnRunHealthCheck"
        $pbHealthCheck = Get-XamlElement -ElementName "pbHealthCheck"
        $lvHealthCheckResults = Get-XamlElement -ElementName "lvHealthCheckResults"
        
        # Globale Variablen setzen
        $script:btnRunHealthCheck = $btnRunHealthCheck
        $script:pbHealthCheck = $pbHealthCheck
        $script:lvHealthCheckResults = $lvHealthCheckResults
        
        # Event-Handler für Health Check Button registrieren
        Register-EventHandler -Control $btnRunHealthCheck -Handler {
            try {
                # Prüfen, ob eine Exchange-Verbindung besteht
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type "Warning"
                    return
                }
                
                Write-Log "Health Check wird gestartet..." -Type "Info"
                Update-StatusBar -Message "Health Check wird gestartet..." -Type "Info"
                Start-ExchangeHealthCheck
            }
            catch {
                $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler bei der Ausführung des Health Checks."
                Write-Log "Fehler bei der Ausführung des Health Checks: $errorMsg" -Type "Error"
                Update-StatusBar -Message "Fehler beim Health Check: $($_.Exception.Message)" -Type "Error"
                Show-MessageBox -Message "Fehler beim Ausführen des Health Checks: `n$($_.Exception.Message)" -Title "Fehler" -Type "Error"
            }
        } -ControlName "btnRunHealthCheck"
        
        Write-Log "Health Check-Tab erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Initialisieren des Health Check-Tabs: $errorMsg" -Type "Error"
        return $false
    }
}

function Initialize-MailFlowRulesTab {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Initialisiere Mail Flow Rules-Tab..." -Type "Info"

        # UI-Elemente referenzieren
        $txtRuleName = Get-XamlElement -ElementName "txtRuleName"
        $cmbCondition = Get-XamlElement -ElementName "cmbCondition"
        $txtConditionValue = Get-XamlElement -ElementName "txtConditionValue"
        $cmbAction = Get-XamlElement -ElementName "cmbAction"
        $txtActionValue = Get-XamlElement -ElementName "txtActionValue"
        $cmbRuleMode = Get-XamlElement -ElementName "cmbRuleMode"
        $btnCreateRule = Get-XamlElement -ElementName "btnCreateRule"
        $btnTestRule = Get-XamlElement -ElementName "btnTestRule"
        $dgMailFlowRules = Get-XamlElement -ElementName "dgMailFlowRules"
        $btnRefreshRules = Get-XamlElement -ElementName "btnRefreshRules"
        $btnEnableRule = Get-XamlElement -ElementName "btnEnableRule"
        $btnDisableRule = Get-XamlElement -ElementName "btnDisableRule"
        $btnDeleteRule = Get-XamlElement -ElementName "btnDeleteRule"
        $btnExportRules = Get-XamlElement -ElementName "btnExportRules"
        $btnImportRules = Get-XamlElement -ElementName "btnImportRules"

        # Globale Variablen setzen
        $script:txtRuleName = $txtRuleName
        $script:cmbCondition = $cmbCondition
        $script:txtConditionValue = $txtConditionValue
        $script:cmbAction = $cmbAction
        $script:txtActionValue = $txtActionValue
        $script:cmbRuleMode = $cmbRuleMode
        $script:dgMailFlowRules = $dgMailFlowRules

        # ComboBoxen initialisieren
        if ($null -ne $cmbCondition) {
            Initialize-MailFlowRuleConditionsComboBox -ComboBox $cmbCondition
        }
        
        if ($null -ne $cmbAction) {
            Initialize-MailFlowRuleActionsComboBox -ComboBox $cmbAction
        }
        
        if ($null -ne $cmbRuleMode) {
            Initialize-MailFlowRuleModeComboBox -ComboBox $cmbRuleMode
        }

        # Event-Handler für "Create Rule" Button
        Register-EventHandler -Control $btnCreateRule -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                # Eingabe validieren
                if ([string]::IsNullOrWhiteSpace($script:txtRuleName.Text)) {
                    Show-MessageBox -Message "Bitte geben Sie einen Namen für die Regel ein." -Title "Fehlende Angabe" -Type Warning
                    return
                }

                if ($null -eq $script:cmbCondition.SelectedItem) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Bedingung aus." -Title "Fehlende Angabe" -Type Warning
                    return
                }

                if ($null -eq $script:cmbAction.SelectedItem) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Aktion aus." -Title "Fehlende Angabe" -Type Warning
                    return
                }

                # Parameter sammeln
                $ruleName = $script:txtRuleName.Text.Trim()
                $conditionType = $script:cmbCondition.SelectedItem.Tag
                $conditionValue = $script:txtConditionValue.Text.Trim()
                $actionType = $script:cmbAction.SelectedItem.Tag
                $actionValue = $script:txtActionValue.Text.Trim()
                $ruleMode = if ($null -ne $script:cmbRuleMode.SelectedItem) { $script:cmbRuleMode.SelectedItem.Tag } else { "Test" }

                # Regel erstellen mit der MailFlow-Funktion
                $result = New-MailFlowRuleAction -RuleName $ruleName -ConditionType $conditionType -ConditionValue $conditionValue -ActionType $actionType -ActionValue $actionValue -Mode $ruleMode

                if ($result) {
                    $script:txtStatus.Text = "Transport Rule '$ruleName' erfolgreich erstellt."
                    # Felder zurücksetzen
                    $script:txtRuleName.Text = ""
                    $script:txtConditionValue.Text = ""
                    $script:txtActionValue.Text = ""
                    # Liste aktualisieren
                    Get-MailFlowRulesAction
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Erstellen der Transport Rule: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnCreateRule"

        # Event-Handler für "Test Rule" Button
        Register-EventHandler -Control $btnTestRule -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                # Eingabe validieren
                if ([string]::IsNullOrWhiteSpace($script:txtRuleName.Text)) {
                    Show-MessageBox -Message "Bitte geben Sie einen Namen für die Regel ein." -Title "Fehlende Angabe" -Type Warning
                    return
                }

                if ($null -eq $script:cmbCondition.SelectedItem) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Bedingung aus." -Title "Fehlende Angabe" -Type Warning
                    return
                }

                # Test-Regel mit der MailFlow-Funktion
                $result = Test-MailFlowRuleAction -RuleName $script:txtRuleName.Text.Trim() -ConditionType $script:cmbCondition.SelectedItem.Tag -ConditionValue $script:txtConditionValue.Text.Trim()

                Show-MessageBox -Message $result -Title "Regel-Test Ergebnis" -Type Info
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Testen der Transport Rule: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnTestRule"

        # Event-Handler für "Refresh Rules" Button
        Register-EventHandler -Control $btnRefreshRules -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                Get-MailFlowRulesAction
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Aktualisieren der Transport Rules: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnRefreshRules"

        # Event-Handler für "Enable Rule" Button
        Register-EventHandler -Control $btnEnableRule -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $selectedRule = $script:dgMailFlowRules.SelectedItem
                if ($null -eq $selectedRule) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Regel aus der Liste aus." -Title "Keine Auswahl" -Type Warning
                    return
                }

                $result = Set-MailFlowRuleStateAction -RuleName $selectedRule.Name -Enabled $true

                if ($result) {
                    $script:txtStatus.Text = "Transport Rule '$($selectedRule.Name)' wurde aktiviert."
                    Get-MailFlowRulesAction
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Aktivieren der Transport Rule: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnEnableRule"

        # Event-Handler für "Disable Rule" Button
        Register-EventHandler -Control $btnDisableRule -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $selectedRule = $script:dgMailFlowRules.SelectedItem
                if ($null -eq $selectedRule) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Regel aus der Liste aus." -Title "Keine Auswahl" -Type Warning
                    return
                }

                $result = Set-MailFlowRuleStateAction -RuleName $selectedRule.Name -Enabled $false

                if ($result) {
                    $script:txtStatus.Text = "Transport Rule '$($selectedRule.Name)' wurde deaktiviert."
                    Get-MailFlowRulesAction
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Deaktivieren der Transport Rule: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnDisableRule"

        # Event-Handler für "Delete Rule" Button
        Register-EventHandler -Control $btnDeleteRule -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $selectedRule = $script:dgMailFlowRules.SelectedItem
                if ($null -eq $selectedRule) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Regel aus der Liste aus." -Title "Keine Auswahl" -Type Warning
                    return
                }

                $confirmResult = Show-MessageBox -Message "Sind Sie sicher, dass Sie die Transport Rule '$($selectedRule.Name)' löschen möchten?" -Title "Löschen bestätigen" -Type Question

                if ($confirmResult -eq "Yes") {
                    $result = Remove-MailFlowRuleAction -RuleName $selectedRule.Name

                    if ($result) {
                        $script:txtStatus.Text = "Transport Rule '$($selectedRule.Name)' wurde gelöscht."
                        Get-MailFlowRulesAction
                    }
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Löschen der Transport Rule: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnDeleteRule"

        # Event-Handler für "Export Rules" Button
        Register-EventHandler -Control $btnExportRules -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|XML-Dateien (*.xml)|*.xml"
                $saveFileDialog.Title = "Transport Rules exportieren"
                $saveFileDialog.FileName = "TransportRules_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

                if ($saveFileDialog.ShowDialog() -eq $true) {
                    $result = Export-MailFlowRulesAction -FilePath $saveFileDialog.FileName

                    if ($result) {
                        $script:txtStatus.Text = "Transport Rules erfolgreich exportiert nach: $($saveFileDialog.FileName)"
                        Show-MessageBox -Message "Transport Rules wurden erfolgreich exportiert." -Title "Export erfolgreich" -Type Info
                    }
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Exportieren der Transport Rules: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnExportRules"

        # Event-Handler für "Import Rules" Button
        Register-EventHandler -Control $btnImportRules -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
                $openFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|XML-Dateien (*.xml)|*.xml"
                $openFileDialog.Title = "Transport Rules importieren"

                if ($openFileDialog.ShowDialog() -eq $true) {
                    $confirmResult = Show-MessageBox -Message "Sind Sie sicher, dass Sie Transport Rules aus der Datei '$($openFileDialog.FileName)' importieren möchten?" -Title "Import bestätigen" -Type Question

                    if ($confirmResult -eq "Yes") {
                        $result = Import-MailFlowRulesAction -FilePath $openFileDialog.FileName

                        if ($result) {
                            $script:txtStatus.Text = "Transport Rules erfolgreich importiert."
                            Get-MailFlowRulesAction
                            Show-MessageBox -Message "Transport Rules wurden erfolgreich importiert." -Title "Import erfolgreich" -Type Info
                        }
                    }
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Importieren der Transport Rules: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnImportRules"

        Write-Log "Mail Flow Rules-Tab erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Initialisieren des Mail Flow Rules-Tabs: $errorMsg" -Type "Error"
        return $false
    }
}

function Initialize-InboxRulesTab {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Initialisiere Inbox Rules-Tab..." -Type "Info"

        # UI-Elemente referenzieren
        $cmbInboxRuleUser = Get-XamlElement -ElementName "cmbInboxRuleUser"
        $btnLoadInboxRules = Get-XamlElement -ElementName "btnLoadInboxRules"
        $txtInboxRuleName = Get-XamlElement -ElementName "txtInboxRuleName"
        $txtFromAddress = Get-XamlElement -ElementName "txtFromAddress"
        $cmbTargetFolder = Get-XamlElement -ElementName "cmbTargetFolder"
        $chkMarkAsRead = Get-XamlElement -ElementName "chkMarkAsRead"
        $btnCreateInboxRule = Get-XamlElement -ElementName "btnCreateInboxRule"
        $dgInboxRules = Get-XamlElement -ElementName "dgInboxRules"
        $btnRefreshInboxRules = Get-XamlElement -ElementName "btnRefreshInboxRules"
        $btnEnableInboxRule = Get-XamlElement -ElementName "btnEnableInboxRule"
        $btnDisableInboxRule = Get-XamlElement -ElementName "btnDisableInboxRule"
        $btnDeleteInboxRule = Get-XamlElement -ElementName "btnDeleteInboxRule"
        $btnMoveRuleUp = Get-XamlElement -ElementName "btnMoveRuleUp"
        $btnMoveRuleDown = Get-XamlElement -ElementName "btnMoveRuleDown"
        $btnExportInboxRules = Get-XamlElement -ElementName "btnExportInboxRules"

        # Globale Variablen setzen
        $script:cmbInboxRuleUser = $cmbInboxRuleUser
        $script:txtInboxRuleName = $txtInboxRuleName
        $script:txtFromAddress = $txtFromAddress
        $script:cmbTargetFolder = $cmbTargetFolder
        $script:chkMarkAsRead = $chkMarkAsRead
        $script:dgInboxRules = $dgInboxRules

        # Event-Handler für "Load Rules" Button (🔄)
        Register-EventHandler -Control $btnLoadInboxRules -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $userIdentity = $script:cmbInboxRuleUser.Text
                if ([string]::IsNullOrWhiteSpace($userIdentity)) {
                    Show-MessageBox -Message "Bitte geben Sie einen Benutzer an." -Title "Keine Auswahl" -Type Warning
                    return
                }
                
                Get-InboxRulesAction -UserIdentity $userIdentity
                
                $folders = Get-MailboxFoldersAction -UserIdentity $userIdentity
                if ($null -ne $script:cmbTargetFolder) {
                    $script:cmbTargetFolder.Items.Clear()
                    
                    $inboxItem = New-Object System.Windows.Controls.ComboBoxItem
                    $inboxItem.Content = "Inbox"
                    $inboxItem.Tag = "Inbox"
                    $script:cmbTargetFolder.Items.Add($inboxItem)
                    
                    foreach ($folder in $folders) {
                        $item = New-Object System.Windows.Controls.ComboBoxItem
                        $item.Content = $folder.DisplayName
                        $item.Tag = $folder.Identity
                        $script:cmbTargetFolder.Items.Add($item)
                    }
                    
                    $script:cmbTargetFolder.SelectedIndex = 0
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Laden der Inbox Rules: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnLoadInboxRules"

        # Event-Handler für "Create Rule" Button
        Register-EventHandler -Control $btnCreateInboxRule -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $userIdentity = $script:cmbInboxRuleUser.Text
                if ([string]::IsNullOrWhiteSpace($userIdentity)) {
                    Show-MessageBox -Message "Bitte geben Sie einen Benutzer an." -Title "Keine Auswahl" -Type Warning
                    return
                }

                if ([string]::IsNullOrWhiteSpace($script:txtInboxRuleName.Text)) {
                    Show-MessageBox -Message "Bitte geben Sie einen Namen für die Regel ein." -Title "Fehlende Angabe" -Type Warning
                    return
                }

                if ([string]::IsNullOrWhiteSpace($script:txtFromAddress.Text)) {
                    Show-MessageBox -Message "Bitte geben Sie eine Absenderadresse ein." -Title "Fehlende Angabe" -Type Warning
                    return
                }

                $ruleName = $script:txtInboxRuleName.Text.Trim()
                $fromAddress = $script:txtFromAddress.Text.Trim()
                $targetFolder = if ($null -ne $script:cmbTargetFolder.SelectedItem) { $script:cmbTargetFolder.SelectedItem.Content.ToString() } else { "Inbox" }
                $markAsRead = $script:chkMarkAsRead.IsChecked

                $result = New-InboxRuleAction -UserIdentity $userIdentity -RuleName $ruleName -FromAddress $fromAddress -TargetFolder $targetFolder -MarkAsRead $markAsRead

                if ($result) {
                    $script:txtStatus.Text = "Inbox Rule '$ruleName' erfolgreich erstellt."
                    $script:txtInboxRuleName.Text = ""
                    $script:txtFromAddress.Text = ""
                    $script:chkMarkAsRead.IsChecked = $false
                    Get-InboxRulesAction -UserIdentity $userIdentity
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Erstellen der Inbox Rule: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnCreateInboxRule"

        # Event-Handler für "Refresh" Button
        Register-EventHandler -Control $btnRefreshInboxRules -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $userIdentity = $script:cmbInboxRuleUser.Text
                if ([string]::IsNullOrWhiteSpace($userIdentity)) {
                    Show-MessageBox -Message "Bitte geben Sie einen Benutzer an." -Title "Keine Auswahl" -Type Warning
                    return
                }

                Get-InboxRulesAction -UserIdentity $userIdentity
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Aktualisieren der Inbox Rules: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnRefreshInboxRules"

        # Event-Handler für "Enable" Button
        Register-EventHandler -Control $btnEnableInboxRule -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $selectedRule = $script:dgInboxRules.SelectedItem
                if ($null -eq $selectedRule) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Regel aus der Liste aus." -Title "Keine Auswahl" -Type Warning
                    return
                }

                $userIdentity = $script:cmbInboxRuleUser.Text
                if ([string]::IsNullOrWhiteSpace($userIdentity)) {
                    Show-MessageBox -Message "Kein Benutzer angegeben, für den die Regel aktiviert werden soll." -Title "Fehler" -Type Error
                    return
                }

                $result = Set-InboxRuleStateAction -UserIdentity $userIdentity -RuleIdentity $selectedRule.Identity -Enabled $true

                if ($result) {
                    $script:txtStatus.Text = "Inbox Rule '$($selectedRule.Name)' wurde aktiviert."
                    Get-InboxRulesAction -UserIdentity $userIdentity
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Aktivieren der Inbox Rule: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnEnableInboxRule"

        # Event-Handler für "Disable" Button
        Register-EventHandler -Control $btnDisableInboxRule -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $selectedRule = $script:dgInboxRules.SelectedItem
                if ($null -eq $selectedRule) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Regel aus der Liste aus." -Title "Keine Auswahl" -Type Warning
                    return
                }

                $userIdentity = $script:cmbInboxRuleUser.Text
                if ([string]::IsNullOrWhiteSpace($userIdentity)) {
                    Show-MessageBox -Message "Kein Benutzer angegeben, für den die Regel deaktiviert werden soll." -Title "Fehler" -Type Error
                    return
                }

                $result = Set-InboxRuleStateAction -UserIdentity $userIdentity -RuleIdentity $selectedRule.Identity -Enabled $false

                if ($result) {
                    $script:txtStatus.Text = "Inbox Rule '$($selectedRule.Name)' wurde deaktiviert."
                    Get-InboxRulesAction -UserIdentity $userIdentity
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Deaktivieren der Inbox Rule: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnDisableInboxRule"

        # Event-Handler für "Delete" Button
        Register-EventHandler -Control $btnDeleteInboxRule -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $selectedRule = $script:dgInboxRules.SelectedItem
                if ($null -eq $selectedRule) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Regel aus der Liste aus." -Title "Keine Auswahl" -Type Warning
                    return
                }

                $confirmResult = Show-MessageBox -Message "Sind Sie sicher, dass Sie die Inbox Rule '$($selectedRule.Name)' löschen möchten?" -Title "Löschen bestätigen" -Type Question

                if ($confirmResult -eq "Yes") {
                    $userIdentity = $script:cmbInboxRuleUser.Text
                    if ([string]::IsNullOrWhiteSpace($userIdentity)) {
                        Show-MessageBox -Message "Kein Benutzer angegeben, für den die Regel gelöscht werden soll." -Title "Fehler" -Type Error
                        return
                    }

                    $result = Remove-InboxRuleAction -UserIdentity $userIdentity -RuleIdentity $selectedRule.Identity

                    if ($result) {
                        $script:txtStatus.Text = "Inbox Rule '$($selectedRule.Name)' wurde gelöscht."
                        Get-InboxRulesAction -UserIdentity $userIdentity
                    }
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Löschen der Inbox Rule - $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnDeleteInboxRule"

        # Event-Handler für "Move Up" Button
        Register-EventHandler -Control $btnMoveRuleUp -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $selectedRule = $script:dgInboxRules.SelectedItem
                if ($null -eq $selectedRule) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Regel aus der Liste aus." -Title "Keine Auswahl" -Type Warning
                    return
                }

                $userIdentity = $script:cmbInboxRuleUser.Text
                if ([string]::IsNullOrWhiteSpace($userIdentity)) {
                    Show-MessageBox -Message "Kein Benutzer angegeben, für den die Regel verschoben werden soll." -Title "Fehler" -Type Error
                    return
                }

                $result = Move-InboxRuleAction -UserIdentity $userIdentity -RuleIdentity $selectedRule.Identity -Direction "Up"

                if ($result) {
                    $script:txtStatus.Text = "Inbox Rule '$($selectedRule.Name)' wurde nach oben verschoben."
                    Get-InboxRulesAction -UserIdentity $userIdentity
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Verschieben der Inbox Rule: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnMoveRuleUp"

        # Event-Handler für "Move Down" Button
        Register-EventHandler -Control $btnMoveRuleDown -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $selectedRule = $script:dgInboxRules.SelectedItem
                if ($null -eq $selectedRule) {
                    Show-MessageBox -Message "Bitte wählen Sie eine Regel aus der Liste aus." -Title "Keine Auswahl" -Type Warning
                    return
                }

                $userIdentity = $script:cmbInboxRuleUser.Text
                if ([string]::IsNullOrWhiteSpace($userIdentity)) {
                    Show-MessageBox -Message "Kein Benutzer angegeben, für den die Regel verschoben werden soll." -Title "Fehler" -Type Error
                    return
                }

                $result = Move-InboxRuleAction -UserIdentity $userIdentity -RuleIdentity $selectedRule.Identity -Direction "Down"

                if ($result) {
                    $script:txtStatus.Text = "Inbox Rule '$($selectedRule.Name)' wurde nach unten verschoben."
                    Get-InboxRulesAction -UserIdentity $userIdentity
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Verschieben der Inbox Rule: $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnMoveRuleDown"

         Register-EventHandler -Control $btnExportInboxRules -Handler {
            try {
                if (-not $script:isConnected) {
                    Show-MessageBox -Message "Bitte stellen Sie zuerst eine Verbindung zu Exchange Online her." -Title "Keine Verbindung" -Type Warning
                    return
                }

                $userIdentity = $script:cmbInboxRuleUser.Text
                if ([string]::IsNullOrWhiteSpace($userIdentity)) {
                    Show-MessageBox -Message "Bitte geben Sie einen Benutzer an." -Title "Keine Auswahl" -Type Warning
                    return
                }

                # Datei-Dialog für Export
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Filter = "CSV-Dateien (*.csv)|*.csv|XML-Dateien (*.xml)|*.xml|Alle Dateien (*.*)|*.*"
                $saveDialog.Title = "Inbox Rules exportieren"
                $saveDialog.FileName = "InboxRules_$($userIdentity.Replace('@', '_'))_$(Get-Date -Format 'yyyyMMdd')"

                if ($saveDialog.ShowDialog() -eq $true) {
                    $result = Export-InboxRulesAction -UserIdentity $userIdentity -FilePath $saveDialog.FileName

                    if ($result) {
                        $script:txtStatus.Text = "Inbox Rules erfolgreich exportiert nach - $($saveDialog.FileName)"
                        Show-MessageBox -Message "Inbox Rules wurden erfolgreich exportiert." -Title "Export erfolgreich" -Type Info
                    }
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Fehler beim Exportieren der Inbox Rules - $errorMsg" -Type "Error"
                $script:txtStatus.Text = "Fehler: $errorMsg"
            }
        } -ControlName "btnExportInboxRules"

        Write-Log "Inbox Rules-Tab erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Fehler beim Initialisieren des Inbox Rules-Tabs - $errorMsg" -Type "Error"
        return $false
    }
}

function Initialize-MessageTraceTab {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Initialisiere Message Trace-Tab..." -Type "Info"

        # UI-Elemente referenzieren
        $script:txtTraceSender = Get-XamlElement -ElementName "txtTraceSender"
        $script:txtTraceRecipient = Get-XamlElement -ElementName "txtTraceRecipient"
        $script:txtTraceSubject = Get-XamlElement -ElementName "txtTraceSubject"
        $script:txtTraceMessageId = Get-XamlElement -ElementName "txtTraceMessageId"
        $script:dpTraceStart = Get-XamlElement -ElementName "dpTraceStart"
        $script:dpTraceEnd = Get-XamlElement -ElementName "dpTraceEnd"
        $script:cmbTraceStatus = Get-XamlElement -ElementName "cmbTraceStatus"
        $script:btnStartTrace = Get-XamlElement -ElementName "btnStartTrace"
        $script:dgMessageTrace = Get-XamlElement -ElementName "dgMessageTrace"
        $script:btnExportTrace = Get-XamlElement -ElementName "btnExportTrace"
        $script:btnDetailedTrace = Get-XamlElement -ElementName "btnDetailedTrace"
        $script:btnClearTrace = Get-XamlElement -ElementName "btnClearTrace"
        $script:txtTraceCount = Get-XamlElement -ElementName "txtTraceCount"

        # Standardwerte setzen
        if ($null -ne $script:dpTraceStart) { $script:dpTraceStart.SelectedDate = (Get-Date).AddDays(-2) }
        if ($null -ne $script:dpTraceEnd) { $script:dpTraceEnd.SelectedDate = (Get-Date) }

        # Event-Handler registrieren
        Register-EventHandler -Control $script:btnStartTrace -Handler {
            try {
                Start-MessageTraceAction
            } catch {
                $errorMsg = Get-FormattedError -ErrorRecord $_
                Write-Log "Fehler beim Starten der Nachrichtensuche: $errorMsg" -Type "Error"
                Update-StatusBar -Message "Fehler bei Nachrichtensuche: $($_.Exception.Message)" -Type Error
            }
        } -ControlName "btnStartTrace"

        Register-EventHandler -Control $script:btnExportTrace -Handler {
            try {
                Export-MessageTraceAction
            } catch {
                $errorMsg = Get-FormattedError -ErrorRecord $_
                Write-Log "Fehler beim Exportieren der Nachrichtensuche: $errorMsg" -Type "Error"
                Update-StatusBar -Message "Fehler beim Export: $($_.Exception.Message)" -Type Error
            }
        } -ControlName "btnExportTrace"

        Register-EventHandler -Control $script:btnDetailedTrace -Handler {
            try {
                Show-DetailedMessageTraceAction
            } catch {
                $errorMsg = Get-FormattedError -ErrorRecord $_
                Write-Log "Fehler beim Anzeigen der Detailansicht: $errorMsg" -Type "Error"
                Update-StatusBar -Message "Fehler bei Detailansicht: $($_.Exception.Message)" -Type Error
            }
        } -ControlName "btnDetailedTrace"

        Register-EventHandler -Control $script:btnClearTrace -Handler {
            try {
                if ($null -ne $script:dgMessageTrace) { $script:dgMessageTrace.ItemsSource = $null }
                if ($null -ne $script:txtTraceCount) { $script:txtTraceCount.Text = "Results: 0" }
                Update-StatusBar -Message "Ergebnisse der Nachrichtensuche gelöscht." -Type Info
            } catch {
                $errorMsg = Get-FormattedError -ErrorRecord $_
                Write-Log "Fehler beim Löschen der Ergebnisse: $errorMsg" -Type "Error"
            }
        } -ControlName "btnClearTrace"

        Write-Log "Message Trace-Tab erfolgreich initialisiert." -Type "Success"
        return $true
    }
    catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_
        Write-Log "Fehler beim Initialisieren des Message Trace-Tabs: $errorMsg" -Type "Error"
        return $false
    }
}

function Initialize-AutoReplyTab {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Initialisiere Auto Reply-Tab..." -Type "Info"

        # UI-Elemente referenzieren
        $script:txtAutoReplyUser = Get-XamlElement -ElementName "txtAutoReplyUser"
        $script:dpAutoReplyStart = Get-XamlElement -ElementName "dpAutoReplyStart"
        $script:dpAutoReplyEnd = Get-XamlElement -ElementName "dpAutoReplyEnd"
        $script:txtAutoReplyMessage = Get-XamlElement -ElementName "txtAutoReplyMessage"
        $script:btnSetAutoReply = Get-XamlElement -ElementName "btnSetAutoReply"
        $script:btnDisableAutoReply = Get-XamlElement -ElementName "btnDisableAutoReply"
        $script:dgAutoReplyStatus = Get-XamlElement -ElementName "dgAutoReplyStatus"
        $script:btnRefreshAutoReply = Get-XamlElement -ElementName "btnRefreshAutoReply"
        $script:btnBulkAutoReply = Get-XamlElement -ElementName "btnBulkAutoReply"
        $script:btnExportAutoReply = Get-XamlElement -ElementName "btnExportAutoReply"

        # Standardwerte setzen
        if ($null -ne $script:dpAutoReplyStart) { $script:dpAutoReplyStart.SelectedDate = (Get-Date) }
        if ($null -ne $script:dpAutoReplyEnd) { $script:dpAutoReplyEnd.SelectedDate = (Get-Date).AddDays(7) }

        # Event-Handler registrieren
        Register-EventHandler -Control $script:btnSetAutoReply -Handler {
            try {
                $user = $script:txtAutoReplyUser.Text
                if ([string]::IsNullOrWhiteSpace($user)) {
                    Show-MessageBox -Message "Bitte wählen Sie einen Benutzer aus oder geben Sie eine E-Mail-Adresse ein." -Title "Fehlende Eingabe" -Type Warning
                    return
                }
                Set-AutoReplyAction -UserIdentity $user `
                    -StartDate $script:dpAutoReplyStart.SelectedDate `
                    -EndDate $script:dpAutoReplyEnd.SelectedDate `
                    -Message $script:txtAutoReplyMessage.Text
            } catch {
                $errorMsg = Get-FormattedError -ErrorRecord $_
                Write-Log "Fehler im btnSetAutoReply Handler: $errorMsg" -Type Error
                Update-StatusBar -Message "Fehler: $($_.Exception.Message)" -Type Error
            }
        } -ControlName "btnSetAutoReply"

        Register-EventHandler -Control $script:btnDisableAutoReply -Handler {
            try {
                $user = $script:txtAutoReplyUser.Text
                if ([string]::IsNullOrWhiteSpace($user)) {
                    Show-MessageBox -Message "Bitte wählen Sie einen Benutzer aus oder geben Sie eine E-Mail-Adresse ein." -Title "Fehlende Eingabe" -Type Warning
                    return
                }
                Disable-AutoReplyAction -UserIdentity $user
            } catch {
                $errorMsg = Get-FormattedError -ErrorRecord $_
                Write-Log "Fehler im btnDisableAutoReply Handler: $errorMsg" -Type Error
                Update-StatusBar -Message "Fehler: $($_.Exception.Message)" -Type Error
            }
        } -ControlName "btnDisableAutoReply"

        Register-EventHandler -Control $script:btnRefreshAutoReply -Handler {
            try {
                # Prüfen ob ein spezifischer Benutzer angegeben wurde
                $user = $script:txtAutoReplyUser.Text.Trim()
                if ([string]::IsNullOrWhiteSpace($user)) {
                    # Wenn kein Benutzer angegeben wurde, alle AutoReply-Status abrufen
                    Get-AutoReplyStatusAction
                } else {
                    # Wenn ein Benutzer angegeben wurde, nur dessen Status abrufen
                    Get-AutoReplyStatusAction -UserIdentity $user
                }
            } catch {
                $errorMsg = Get-FormattedError -ErrorRecord $_
                Write-Log "Fehler im btnRefreshAutoReply Handler: $errorMsg" -Type Error
                Update-StatusBar -Message "Fehler beim Aktualisieren: $($_.Exception.Message)" -Type Error
            }
        } -ControlName "btnRefreshAutoReply"

        Register-EventHandler -Control $script:btnExportAutoReply -Handler {
            try {
                Export-AutoReplyStatusAction
            } catch {
                $errorMsg = Get-FormattedError -ErrorRecord $_
                Write-Log "Fehler im btnExportAutoReply Handler: $errorMsg" -Type Error
                Update-StatusBar -Message "Fehler beim Export: $($_.Exception.Message)" -Type Error
            }
        } -ControlName "btnExportAutoReply"
        
        Register-EventHandler -Control $script:btnBulkAutoReply -Handler {
            Show-MessageBox -Message "Die Funktion für die Massenbearbeitung ist noch nicht implementiert." -Title "In Entwicklung" -Type Info
        } -ControlName "btnBulkAutoReply"


        Write-Log "Auto Reply-Tab erfolgreich initialisiert." -Type "Success"
        return $true
    }
    catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_
        Write-Log "Fehler beim Initialisieren des Auto Reply-Tabs: $errorMsg" -Type "Error"
        return $false
    }
}

function Initialize-TabNavigation {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Initialisiere Tab-Navigation..." -Type "Info"
        
        $tabControl = $script:Form.FindName("tabContent")
        if ($null -eq $tabControl) {
            Write-Log "TabControl 'tabContent' nicht gefunden" -Type "Error"
            return $false
        }
        
        # WICHTIG: Verstecke alle Tab-Header
        $tabControl.Dispatcher.Invoke([Action]{
            foreach ($tabItem in $tabControl.Items) {
                if ($tabItem -is [System.Windows.Controls.TabItem]) {
                    # Setze Header auf null oder leer
                    $tabItem.Header = $null
                    # Oder: $tabItem.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }
        }, "Normal")
        
        # Tab-Mapping für Navigation-Buttons zu TabItem-Namen
        $script:tabMapping = @{
            # Dashboard
            'btnDashboard' = 'tabDashboard'
            
            # Grundlegende Verwaltung
            'btnNavCalendar' = 'tabCalendar'
            'btnNavMailbox' = 'tabMailbox'
            'btnNavSharedMailbox' = 'tabSharedMailbox'
            'btnNavGroups' = 'tabGroups'
            'btnNavResources' = 'tabResources'
            'btnNavContacts' = 'tabContacts'
            
            # Mail Flow
            'btnNavMailFlowRules' = 'tabMailFlowRules'
            'btnNavInboxRules' = 'tabInboxRules'
            'btnNavMessageTrace' = 'tabMessageTrace'
            'btnNavAutoReply' = 'tabAutoReply'
            
            # Sicherheit & Compliance
            'btnNavATP' = 'tab_ATP'
            'btnNavDLP' = 'tab_DLP'
            'btnNavEDiscovery' = 'tab_eDiscovery'
            'btnNavMDM' = 'tab_MDM'
            
            # Systemkonfiguration
            'btnNavEXOSettings' = 'tabEXOSettings'
            'btnNavRegion' = 'tabRegion'
            'btnNavCrossPremises' = 'tab_MailRouting'
            
            # Erweiterte Verwaltung
            'btnNavHybridExchange' = 'tab_HybridExchange'
            'btnNavMultiForest' = 'tab_MultiForest'
            
            # Monitoring & Support
            'btnNavHealthCheck' = 'tab_HealthCheck'
            'btnNavAudit' = 'tabMailboxAudit'
            'btnNavReports' = 'tabReports'
            'btnNavTroubleshooting' = 'tabTroubleshooting'
        }
        
  
        # Event-Handler registrieren...
        foreach ($btnName in $script:tabMapping.Keys) {
            $button = $script:Form.FindName($btnName)
            $targetTabName = $script:tabMapping[$btnName]
            
            if ($null -ne $button) {
                $targetTabItem = $null
                foreach ($tabItem in $tabControl.Items) {
                    if ($tabItem.Name -eq $targetTabName) {
                        $targetTabItem = $tabItem
                        break
                    }
                }
                
                if ($null -ne $targetTabItem) {
                    $button.Add_Click({
                        param($sender, $e)
                        try {
                            $clickedButton = $sender
                            $buttonName = $clickedButton.Name
                            $targetTabName = $script:tabMapping[$buttonName]
                            
                            $tc = $script:Form.FindName("tabContent")
                            if ($null -ne $tc) {
                                foreach ($item in $tc.Items) {
                                    if ($item.Name -eq $targetTabName) {
                                        $tc.SelectedItem = $item
                                        Set-ActiveNavigationButton -ButtonName $buttonName
                                        break
                                    }
                                }
                            }
                        }
                        catch {
                            Write-Log "Fehler bei Tab-Navigation: $($_.Exception.Message)" -Type "Error"
                        }
                    })
                }
            }
        }
        
        # Initial ersten Tab auswählen
        if ($tabControl.Items.Count -gt 0) {
            $tabControl.SelectedIndex = 0
        }
        
        Write-Log "Tab-Navigation erfolgreich initialisiert" -Type "Success"
        return $true
    }
    catch {
        Write-Log "Fehler bei Tab-Navigation Initialisierung: $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

# Initialisiert die wesentliche UI und startet die asynchrone Initialisierung der restlichen Tabs
function Initialize-ApplicationUI {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Initialisiere wesentliche UI-Komponenten..." -Type "Info"

        # 1. Tab-Navigation initialisieren, damit die Buttons funktionieren
        $navResult = Initialize-TabNavigation
        if (-not $navResult) {
            Write-Log "Kritisch: Tab-Navigation konnte nicht initialisiert werden. Abbruch." -Type "Error"
            return $false
        }
        Write-Log "Tab-Navigation erfolgreich initialisiert." -Type "Success"

        # 2. Dashboard-Tab als Start-Tab festlegen
        $tabControl = $script:Form.FindName("tabContent")
        $dashboardTab = $script:Form.FindName("tabDashboard")
        if ($null -ne $tabControl -and $null -ne $dashboardTab) {
            $tabControl.SelectedItem = $dashboardTab
            Set-ActiveNavigationButton -ButtonName 'btnDashboard'
            Write-Log "Dashboard-Tab als Startansicht festgelegt." -Type "Info"
        } else {
            Write-Log "Dashboard-Tab oder Haupt-TabControl nicht gefunden. Kann Startansicht nicht festlegen." -Type "Warning"
        }

        # 3. Asynchrone Initialisierung der restlichen Tabs über Background-Job starten
        Start-AsyncTabInitialization
        
        return $true
    }
    catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler bei der Initialisierung der Anwendungs-UI."
        Write-Log $errorMsg -Type "Error"
        return $false
    }
}

# Initialisiert die restlichen Tabs im Hintergrund über einen separaten Thread
function Initialize-RemainingTabs {
    [CmdletBinding()]
    param()

    Write-Log "Beginne mit der Initialisierung der verbleibenden Tabs." -Type "Info"
    
    # Status über Dispatcher aktualisieren
    $script:Form.Dispatcher.Invoke([Action]{
        Update-StatusBar -Message "Lade Module im Hintergrund..."
    }, "Background")

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
        RegionSettings = Initialize-RegionSettingsTab
        MailFlowRules = Initialize-MailFlowRulesTab
        InboxRules = Initialize-InboxRulesTab
        MessageTrace = Initialize-MessageTraceTab
        AutoReply = Initialize-AutoReplyTab
        HealthCheck = Initialize-HealthCheckTab
        # Tabs welche keine Funktion haben, aber im UI vorhanden sind
        #ATP = Initialize-ATPTab
        #DLP = Initialize-DLPTab
        #eDiscovery = Initialize-eDiscoveryTab
        #MDM = Initialize-MDMTab
        #HybridExchange = Initialize-HybridExchangeTab
        #MultiForest = Initialize-MultiForestTab
        #CrossPremises = Initialize-CrossPremisesTab
    }

    $successCount = ($results.Values | Where-Object { $_ -eq $true }).Count
    $totalCount = $results.Count
    $failedTabs = ($results.Keys | Where-Object { $results[$_] -eq $false }) -join ", "

    Write-Log "Zusammenfassung der Tab-Initialisierung:" -Type "Info"
    foreach ($tab in $results.Keys | Sort-Object) {
        $status = if ($results[$tab]) { "Erfolgreich" } else { "Fehlgeschlagen" }
        Write-Log "- $tab`: $status" -Type "Debug"
    }

    # Status über Dispatcher aktualisieren
    $script:Form.Dispatcher.Invoke([Action]{
        if ($successCount -eq $totalCount) {
            Update-StatusBar -Message "Alle Module erfolgreich geladen - Bereit - Bitte mit Exchange Online verbinden." -Type "Success"
            Write-Log "Alle Tabs wurden erfolgreich initialisiert." -Type "Success"
        } else {
            Update-StatusBar -Message "Einige Module konnten nicht geladen werden. Details im Log." -Type "Warning"
            Write-Log "Initialisierung für folgende Tabs fehlgeschlagen: $failedTabs" -Type "Warning"
        }
    }, "Background")
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
    
    # WICHTIG: UI initialisieren (startet den asynchronen Ladevorgang)
    $result = Initialize-ApplicationUI
    Log-Action "Initialize-ApplicationUI Ergebnis: $result"
    
    # Hilfe-Links initialisieren
    $result = Initialize-HelpLinks
    Log-Action "Initialize-HelpLinks Ergebnis: $result"
})

function Start-AsyncTabInitialization {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Starte asynchrone Tab-Initialisierung..." -Type "Info"
        
        # Statt Runspace verwenden wir die echte Initialisierung direkt
        # aber in einem Background-Thread über Dispatcher
        $script:Form.Dispatcher.BeginInvoke([Action]{
            try {
                Write-Log "Beginne synchrone Tab-Initialisierung im Hintergrund..." -Type "Debug"
                
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
                    RegionSettings = Initialize-RegionSettingsTab
                    MailFlowRules = Initialize-MailFlowRulesTab
                    InboxRules = Initialize-InboxRulesTab
                    MessageTrace = Initialize-MessageTraceTab
                    AutoReply = Initialize-AutoReplyTab
                    HealthCheck = Initialize-HealthCheckTab
                }
                
                $successCount = ($results.Values | Where-Object { $_ -eq $true }).Count
                $totalCount = $results.Count
                $failedTabs = ($results.Keys | Where-Object { $results[$_] -eq $false }) -join ", "
                
                Write-Log "Tab-Initialisierung abgeschlossen: $successCount/$totalCount erfolgreich" -Type "Info"
                
                if ($successCount -eq $totalCount) {
                    Update-StatusBar -Message "Alle Module erfolgreich geladen - Bereit - Bitte mit Exchange Online verbinden." -Type "Success"
                    Write-Log "Alle Tabs wurden erfolgreich initialisiert." -Type "Success"
                } else {
                    Update-StatusBar -Message "Einige Module konnten nicht geladen werden. Details im Log." -Type "Warning"
                    Write-Log "Initialisierung für folgende Tabs fehlgeschlagen: $failedTabs" -Type "Warning"
                }
                
                # Detailliertes Logging
                foreach ($tab in $results.Keys | Sort-Object) {
                    $status = if ($results[$tab]) { "Erfolgreich" } else { "Fehlgeschlagen" }
                    Write-Log "- $tab`: $status" -Type "Debug"
                }
                
            } catch {
                $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler bei der Tab-Initialisierung."
                Write-Log $errorMsg -Type "Error"
                Update-StatusBar -Message "Fehler beim Laden der Module." -Type "Error"
            }
        }, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
        
        Write-Log "Asynchrone Tab-Initialisierung gestartet" -Type "Success"
        return $true
        
    } catch {
        $errorMsg = Get-FormattedError -ErrorRecord $_ -DefaultText "Fehler beim Starten der asynchronen Tab-Initialisierung."
        Write-Log $errorMsg -Type "Error"
        
        # Fallback: Direkte synchrone Initialisierung
        try {
            Write-Log "Fallback: Starte direkte Tab-Initialisierung..." -Type "Warning"
            Update-StatusBar -Message "Lade Module direkt..." -Type "Warning"
            Initialize-RemainingTabs
            return $true
        } catch {
            $fallbackError = Get-FormattedError -ErrorRecord $_ -DefaultText "Auch direkte Tab-Initialisierung fehlgeschlagen."
            Write-Log $fallbackError -Type "Error"
            Update-StatusBar -Message "Fehler beim Laden der Module." -Type "Error"
            return $false
        }
    }
}

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

    # Aufräumen nach Schließen des Fensters
    if ($script:isConnected) {
        Log-Action "Trenne Exchange Online-Verbindung..."
        Disconnect-ExchangeOnlineSession
    }
}
catch {
    $errorMsg = $_.Exception.Message
    Write-Log "Kritischer Fehler beim Laden oder Anzeigen der GUI: $errorMsg"  
    Log-Action "Stack Trace: $($_.Exception.StackTrace)"
    
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
        try {
            $script:Form.Close()
            $script:Form = $null
        } catch {
            Write-Log "Fehler beim Schließen des Formulars im finally-Block: $($_.Exception.Message)" -Type Warning
        }
    }
    Log-Action "Aufräumarbeiten abgeschlossen"
}



# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCPD8v38OKoXzTE
# /suWmTGRIIcOon+rqz1UjgenALvXRaCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBq4wggSWoAMCAQICEAc2
# N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAw
# MFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFE
# FUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoi
# GN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YA
# e9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O
# 9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI
# 1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7m
# O1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPK
# qpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8F
# nGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMD
# iP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4Jduyr
# XUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91
# jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW
# 2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H
# +oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4os
# equFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p
# /yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnf
# xI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36T
# U6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0
# cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf
# +yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa6
# 3VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1d
# wvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9E
# FUrnEw4d2zc4GqEr9u3WfPwwgga8MIIEpKADAgECAhALrma8Wrp/lYfG+ekE4zME
# MA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUxMTI1MjM1
# OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAeBgNVBAMT
# F0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/QowIEMSvgjE
# dEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7yijvoQ7u
# jm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHjes4fduks
# THulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhNf1F41nyE
# g5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9HlfqSBePejlY
# eEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPxRNUNK6lY
# k2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhzXomJ2Ple
# I9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I78JpwGpT
# RHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ33c1HG93V
# p6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rtvVcIH7Wv
# G9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkCAwEAAaOC
# AYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAf
# BgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUn1csA3cO
# KBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFt
# cGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4hBJH2UOR
# 9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2uVYFvQe+p
# PTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51sMLMXNTL
# fhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QUAvVSu4kq
# VOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSbdakHJe2B
# VDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRUAYSyyEmY
# tsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CWT/xrW7tw
# ipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZaA0VhqAsM
# HOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULkftARjsyE
# pHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHYSAR16gc0
# dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzLP8lx4Q1z
# ZKDyHcp4VQJLu2kWTsKsOqQxggUKMIIFBgIBATA0MCAxHjAcBgNVBAMMFVBoaW5J
# VC1QU3NjcmlwdHNfU2lnbgIQd487Ml/QoIxIvrAQtqwTEzANBglghkgBZQMEAgEF
# AKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3
# DQEJBDEiBCBk4/CujhAFKh78Xj9Ju86LsIH6rPKBK7ggTreD0nwP/DANBgkqhkiG
# 9w0BAQEFAASCAQCgruHqzZA4G86djwWiU6Cj5nB0YoM/SY9VraXoBWHWvs4U8Edb
# dD4XRCAvOtka8rMdo9o5xuTj+EyZ787KwuacOav1IjMUhUnKePV5zAE5PzNqzWYf
# kvheu8F++pkOyWoQ+D851r8s7UUpC3hmEbaRQ1D9zrnkGI0BwgQUkpSnshtGE3aC
# efxE06b+zgArEJrkp1dOoF4rDUezNAWdwsI+FfRgwoCyA6RhyUTYd69Ks+jxehhm
# b/MmZsthm3tpkHf3EIaYeL4lwIkNYowL53v/nAMtjHWWd9p6qBdKHNTCQcTOsxtn
# o9wQPSeUFPwjkgfX2GJ8XwwtDDxkfykb03WvoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNjE1NDY1OFowLwYJKoZIhvcNAQkEMSIEIPb9jMXu0tNGvSroo/M4UUJKwcAd
# Ca4pvvBlo+Rj4IQBMA0GCSqGSIb3DQEBAQUABIICAJpZKmSGb1AC1S/BYKI/Y+HC
# bQxZHLCELymmbea5tTBr+LhHaPPt+92LroLh4spodfeO/urojku1+OHAkjEwoEWj
# dA4YC+ijyvHrJWzz1chbZI8hdhMyg8wbUI61nqsykaV64UnP9+PqmVw8vbN7wqQf
# Hx+kwzFRLFXUnAowEQfrfBqmqqh8lwOdk7682P/x4qhHA/o/U91zhrrWle5B6QWS
# W9s5DlBNPyhky8SwxLgvbchjSRcVswKlrzFNdD5nOIf7Q048QpGWnr1xDJBXT0f0
# XUT+UJQZnsYmUrvS/nOE+NV/R6IHxbZrc179pbIWMfQPmdtrrgcjts+jUESyrCiO
# lKw0uM6EigwsBCRg9/BeTYtwBcjDWVNh3frh1a3vLPiAr28XyaHJsDGJxrw8LNUD
# eTe8wsjxRSby3MZDDgzT0ApK5h2QeJDb21ecmSVDX5jm6a2uUy8dqsscBmU18V5b
# U6oStbWNNZscukMcxcHKuhBmbwgmJShQVsFqmWFu03pzEyPslVEs+ybVJr55rskj
# OQYnK//RW+BOrv8IXtshR195JUmCEZGqeWWBrEKUCsql6z1glGEgUSvotl6uBdD7
# JdhBs8Rnnso8Y08u27aq2sWin9bXtetdmOF5GWJZgq4AbIvmLQAfJiHwyIE2Iifk
# vddqo1P2eNMlDueFfRjO
# SIG # End signature block
