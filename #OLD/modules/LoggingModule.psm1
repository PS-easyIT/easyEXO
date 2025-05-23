#
# LoggingModule.psm1
# Modul für zentralisiertes Logging in PowerShell-Anwendungen
#

# Variablen für Logging-Konfiguration
$script:LogPath = Join-Path -Path $PSScriptRoot -ChildPath "..\logs"
$script:LogFile = "easyEXO.log"
$script:MaxLogSize = 10MB
$script:MaxLogHistory = 5
$script:LogLevel = "Info" # Mögliche Werte: Debug, Info, Warning, Error

# Stellt sicher, dass der Log-Ordner existiert
if (-not (Test-Path -Path $script:LogPath)) {
    try {
        New-Item -Path $script:LogPath -ItemType Directory -Force | Out-Null
    } catch {
        Write-Error "Konnte Log-Verzeichnis nicht erstellen: $($_.Exception.Message)"
    }
}

# Funktion zum Rotieren des Logfiles
function Rotate-LogFile {
    try {
        $logFilePath = Join-Path -Path $script:LogPath -ChildPath $script:LogFile
        
        if ((Test-Path -Path $logFilePath) -and ((Get-Item -Path $logFilePath).Length -gt $script:MaxLogSize)) {
            # Bestehende Rotationen verschieben
            for ($i = $script:MaxLogHistory; $i -gt 0; $i--) {
                $oldFile = Join-Path -Path $script:LogPath -ChildPath "$($script:LogFile).$($i-1)"
                $newFile = Join-Path -Path $script:LogPath -ChildPath "$($script:LogFile).$i"
                
                if ($i -eq 1) {
                    $oldFile = Join-Path -Path $script:LogPath -ChildPath $script:LogFile
                }
                
                if (Test-Path -Path $oldFile) {
                    if (Test-Path -Path $newFile) {
                        Remove-Item -Path $newFile -Force
                    }
                    
                    Move-Item -Path $oldFile -Destination $newFile -Force
                }
            }
            
            # Neue leere Logdatei erstellen
            $null = New-Item -Path $logFilePath -ItemType File -Force
        }
    } catch {
        # Hier keine Fehlermeldung werfen, um rekursive Logging-Fehler zu vermeiden
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $errorMessage = "[$timestamp] FEHLER bei Log-Rotation: $($_.Exception.Message)"
        try {
            $fallbackLogPath = Join-Path -Path $script:LogPath -ChildPath "fallback_log.txt"
            Add-Content -Path $fallbackLogPath -Value $errorMessage -Encoding UTF8
        } catch {
            # Letzte Möglichkeit: Ausgabe auf der Konsole
            Write-Host $errorMessage -ForegroundColor Red
        }
    }
}

# Soll Log-Level geloggt werden?
function Test-LogLevel {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Debug", "Info", "Warning", "Error")]
        [string]$Level
    )
    
    $logLevelValue = @{
        "Debug" = 0
        "Info" = 1
        "Warning" = 2
        "Error" = 3
    }
    
    return $logLevelValue[$Level] -ge $logLevelValue[$script:LogLevel]
}

# Hauptfunktion für das Logging
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [AllowEmptyString()]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Debug", "Info", "Warning", "Error")]
        [string]$Level = "Info",
        
        [Parameter(Mandatory=$false)]
        [switch]$NoConsole,
        
        [Parameter(Mandatory=$false)]
        [switch]$NoTimestamp
    )
    
    try {
        # Prüfen, ob Level geloggt werden soll
        if (-not (Test-LogLevel -Level $Level)) {
            return
        }
        
        # Log-Rotation prüfen
        Rotate-LogFile
        
        # Timestamp erstellen
        $timestamp = if (-not $NoTimestamp) { Get-Date -Format "yyyy-MM-dd HH:mm:ss" } else { "" }
        
        # Nachricht filtern (nur druckbare ASCII-Zeichen)
        $filteredMessage = $Message -replace '[^\x20-\x7E]', '?'
        
        # Nachricht formatieren
        $logEntry = if (-not $NoTimestamp) { "[$timestamp] [$Level] $filteredMessage" } else { "[$Level] $filteredMessage" }
        
        # In Datei schreiben
        $logFilePath = Join-Path -Path $script:LogPath -ChildPath $script:LogFile
        Add-Content -Path $logFilePath -Value $logEntry -Encoding UTF8
        
        # Auf Konsole ausgeben, falls gewünscht
        if (-not $NoConsole) {
            $consoleColors = @{
                "Debug" = "Gray"
                "Info" = "White"
                "Warning" = "Yellow"
                "Error" = "Red"
            }
            
            Write-Host $logEntry -ForegroundColor $consoleColors[$Level]
        }
    } catch {
        # Bei Fehlern im Logging einen Fallback verwenden
        try {
            $fallbackLogPath = Join-Path -Path $script:LogPath -ChildPath "fallback_log.txt"
            $errorMessage = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] FEHLER im Logging-Modul: $($_.Exception.Message)"
            Add-Content -Path $fallbackLogPath -Value $errorMessage -Encoding UTF8
            Add-Content -Path $fallbackLogPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Original-Nachricht: $Message" -Encoding UTF8
            
            # Auch auf Konsole ausgeben
            Write-Host $errorMessage -ForegroundColor Red
        } catch {
            # Letzte Möglichkeit: Ausgabe auf der Konsole
            Write-Host "KRITISCHER FEHLER im Logging-Modul: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Debug-Funktionen
function Write-Debug {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [AllowEmptyString()]
        [string]$Message
    )
    
    Write-Log -Message $Message -Level "Debug"
}

function Write-Info {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [AllowEmptyString()]
        [string]$Message
    )
    
    Write-Log -Message $Message -Level "Info"
}

function Write-Warning {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [AllowEmptyString()]
        [string]$Message
    )
    
    Write-Log -Message $Message -Level "Warning"
}

function Write-Error {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [AllowEmptyString()]
        [string]$Message
    )
    
    Write-Log -Message $Message -Level "Error"
}

# Funktionen zur Konfiguration des Loggers
function Set-LogLevel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Debug", "Info", "Warning", "Error")]
        [string]$Level
    )
    
    $script:LogLevel = $Level
    Write-Log "Log-Level auf '$Level' gesetzt." -Level "Info"
}

function Set-LogPath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        $newPath = [System.IO.Path]::GetFullPath($Path)
        
        if (-not (Test-Path -Path $newPath)) {
            New-Item -Path $newPath -ItemType Directory -Force | Out-Null
        }
        
        $script:LogPath = $newPath
        Write-Log "Log-Pfad auf '$newPath' gesetzt." -Level "Info"
    } catch {
        Write-Error "Fehler beim Setzen des Log-Pfads: $($_.Exception.Message)"
    }
}

function Set-LogFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$FileName
    )
    
    $script:LogFile = $FileName
    Write-Log "Log-Datei auf '$FileName' gesetzt." -Level "Info"
}

function Set-LogRotationSize {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [long]$SizeInBytes
    )
    
    $script:MaxLogSize = $SizeInBytes
    Write-Log "Log-Rotationsgröße auf $($SizeInBytes/1MB) MB gesetzt." -Level "Info"
}

# Begrenzen der Textfeldlänge für GUI-Ausgaben
function Limit-TextFieldContent {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.TextBox]$TextBox,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxLength = 10000
    )
    
    try {
        if ($null -eq $TextBox) {
            return
        }
        
        if ($TextBox.Text.Length -gt $MaxLength) {
            # Nur die letzten MaxLength Zeichen behalten
            $TextBox.Text = $TextBox.Text.Substring($TextBox.Text.Length - $MaxLength)
            
            # Zur letzten Zeile scrollen
            $TextBox.ScrollToEnd()
        }
    } catch {
        # Fehler beim Begrenzen der Textfeldlänge
        try {
            Write-Log "FEHLER beim Begrenzen der Textfeldlänge: $($_.Exception.Message)" -Level "Error"
        } catch {
            # Nichts tun - stille Fehlerbehandlung
        }
    }
}

# Safe UI Update Funktion
function Update-UITextSafe {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [object]$Control,
        
        [Parameter(Mandatory=$true)]
        [string]$PropertyName,
        
        [Parameter(Mandatory=$true)]
        [string]$Value
    )
    
    try {
        if ($null -eq $Control) {
            return
        }
        
        # Nur druckbare ASCII-Zeichen filtern
        $filteredValue = $Value -replace '[^\x20-\x7E]', '?'
        
        # PowerShell-Runspace sicher aktualisieren
        if ($Control.Dispatcher -and $Control.Dispatcher.CheckAccess()) {
            # Wir sind im UI-Thread, direkt aktualisieren
            $Control.$PropertyName = $filteredValue
        } else {
            # Wir müssen zum UI-Thread wechseln
            $Control.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Normal, [System.Action]{
                $Control.$PropertyName = $filteredValue
            })
        }
        
        # Bei TextBox Längenbegrenzung durchführen
        if ($Control -is [System.Windows.Controls.TextBox]) {
            Limit-TextFieldContent -TextBox $Control
        }
    } catch {
        # Fehler beim Update der UI
        try {
            Write-Log "FEHLER beim Aktualisieren von UI-Element: $($_.Exception.Message)" -Level "Error"
        } catch {
            # Nichts tun - stille Fehlerbehandlung
        }
    }
}

# Exportiere die Module-Funktionen
Export-ModuleMember -Function Write-Log, Write-Debug, Write-Info, Write-Warning, Write-Error, 
                             Set-LogLevel, Set-LogPath, Set-LogFile, Set-LogRotationSize,
                             Limit-TextFieldContent, Update-UITextSafe