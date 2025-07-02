<#
.SYNOPSIS
    Logging-Modul für easyEXO.
.DESCRIPTION
    Stellt Funktionen für Debugging und Logging bereit.
.NOTES
    Version: 1.0
#>

# Modul-Variablen
$script:debugMode = $false
$script:logFilePath = $null

# Initialisierung des Logging-Moduls
function Initialize-EXOLogging {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,
        
        [Parameter(Mandatory = $false)]
        [bool]$DebugMode = $false
    )
    
    try {
        $script:logFilePath = $LogFilePath
        $script:debugMode = $DebugMode
        
        # Logverzeichnis erstellen, falls nicht vorhanden
        $logFolder = Split-Path -Path $script:logFilePath -Parent
        if (-not (Test-Path $logFolder)) {
            New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
        }
        
        Write-Debug "Logging initialisiert: $LogFilePath, DebugMode: $DebugMode"
        Write-Log -Message "Logging-System initialisiert" -Level "Info"
        
        return $true
    }
    catch {
        # Fallback für kritische Initialisierungsfehler
        Write-Error "Fehler bei der Initialisierung des Logging-Systems: $($_.Exception.Message)"
        
        # Versuche trotzdem in eine lokale Datei zu loggen
        try {
            $fallbackLogPath = Join-Path -Path $env:TEMP -ChildPath "easyEXO_fallback.log"
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -Path $fallbackLogPath -Value "[$timestamp] KRITISCH: Logging-Initialisierung fehlgeschlagen: $($_.Exception.Message)" -Encoding UTF8
        }
        catch {
            # Ignorieren - letzte Rückfallebene
        }
        
        return $false
    }
}

# Hauptfunktion für Debug-Ausgabe
function Write-Debug {
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
            Write-Log -Message "DEBUG: $Type - $sanitizedMessage" -Level $Type
        }
    }
    catch {
        # Fallback für Fehler in der Debug-Funktion - schreibe direkt ins Log
        try {
            $errorMsg = $_.Exception.Message -replace '[^\x20-\x7E]', '?'
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $fallbackLogPath = Join-Path -Path $env:TEMP -ChildPath "easyEXO_debug_fallback.log"
            Add-Content -Path $fallbackLogPath -Value "[$timestamp] Fehler in Write-Debug: $errorMsg" -Encoding UTF8
        }
        catch {
            # Absoluter Fallback - ignoriere Fehler um Programmablauf nicht zu stören
        }
    }
}

# Hauptfunktion für Log-Ausgabe
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info",
        
        [Parameter(Mandatory = $false)]
        [switch]$NoRotation
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
        }
        
        # Log-Eintrag mit Level schreiben
        Add-Content -Path $script:logFilePath -Value "[$timestamp] [$Level] $sanitizedMessage" -Encoding UTF8
        
        # Bei zu langer Logdatei rotieren, es sei denn NoRotation ist angegeben
        if (-not $NoRotation) {
            $logFile = Get-Item -Path $script:logFilePath -ErrorAction SilentlyContinue
            if ($logFile -and $logFile.Length -gt 10MB) {
                $backupLogPath = "$($script:logFilePath)_$(Get-Date -Format 'yyyyMMdd_HHmmss').bak"
                Move-Item -Path $script:logFilePath -Destination $backupLogPath -Force
                Write-Debug -Message "Logdatei wurde rotiert: $backupLogPath" -Type "Info"
            }
        }
    }
    catch {
        # Fallback für Fehler in der Log-Funktion
        try {
            $errorMsg = $_.Exception.Message -replace '[^\x20-\x7E]', '?'
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $fallbackLogPath = Join-Path -Path $env:TEMP -ChildPath "easyEXO_log_fallback.log"
            
            Add-Content -Path $fallbackLogPath -Value "[$timestamp] Fehler in Write-Log: $errorMsg" -Encoding UTF8
            Add-Content -Path $fallbackLogPath -Value "[$timestamp] Ursprüngliche Nachricht: $sanitizedMessage" -Encoding UTF8
        }
        catch {
            # Absoluter Fallback - ignoriere Fehler um Programmablauf nicht zu stören
        }
    }
}

# Exportiere Funktionen
Export-ModuleMember -Function Initialize-EXOLogging, Write-Debug, Write-Log
