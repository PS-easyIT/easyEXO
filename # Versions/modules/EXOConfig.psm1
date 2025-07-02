<#
.SYNOPSIS
    Konfigurationsmodul für easyEXO.
.DESCRIPTION
    Stellt Funktionen für das Laden, Speichern und Verwalten von Konfigurationseinstellungen bereit.
.NOTES
    Version: 1.0
#>

# Modul-Variablen
$script:configFilePath = $null
$script:config = $null
$script:configWatcher = $null
$script:onConfigChangedAction = $null

# Modul initialisieren
function Initialize-EXOConfig {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ConfigFilePath,
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$ConfigChangedAction = $null
    )
    
    try {
        $script:configFilePath = $ConfigFilePath
        $script:onConfigChangedAction = $ConfigChangedAction
        
        $configFolder = Split-Path -Path $ConfigFilePath -Parent
        if (-not (Test-Path -Path $configFolder)) {
            New-Item -ItemType Directory -Path $configFolder -Force | Out-Null
        }
        
        # Wenn die Konfigurationsdatei nicht existiert, erstelle sie mit Standardwerten
        if (-not (Test-Path -Path $ConfigFilePath)) {
            New-DefaultConfig -FilePath $ConfigFilePath
        }
        
        # Konfiguration laden
        $script:config = Get-IniContent -FilePath $ConfigFilePath
        
        # FileSystemWatcher für Hot-Reload einrichten
        $script:configWatcher = Setup-ConfigFileWatcher -FilePath $ConfigFilePath
        
        return $true
    }
    catch {
        Write-Error "Fehler bei der Initialisierung des Konfigurations-Moduls: $($_.Exception.Message)"
        return $false
    }
}

function Get-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        $ini = @{}
        switch -regex -file $FilePath {
            "^\[(.+)\]" {
                $section = $matches[1]
                $ini[$section] = @{}
                continue
            }
            "^\s*([^#].+?)\s*=\s*(.*)" {
                $name, $value = $matches[1..2]
                if ($name -and $section) {
                    $ini[$section][$name] = $value.Trim()
                }
                continue
            }
        }
        return $ini
    }
    catch {
        Write-Error "Fehler beim Lesen der INI-Datei: $($_.Exception.Message)"
        return @{}
    }
}

function Set-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$IniData
    )
    
    try {
        $content = @()
        
        foreach ($section in $IniData.Keys) {
            $content += "[$section]"
            
            foreach ($key in $IniData[$section].Keys) {
                $content += "$key = $($IniData[$section][$key])"
            }
            
            $content += ""  # Leere Zeile zwischen Abschnitten
        }
        
        $content | Out-File -FilePath $FilePath -Encoding UTF8 -Force
        return $true
    }
    catch {
        Write-Error "Fehler beim Schreiben der INI-Datei: $($_.Exception.Message)"
        return $false
    }
}

function New-DefaultConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        $scriptRoot = Split-Path -Path $FilePath -Parent | Split-Path -Parent
        $logPath = Join-Path -Path $scriptRoot -ChildPath "Logs"
        
        $defaultConfig = @"
[General]
Debug = 1
AppName = Exchange Berechtigungen Verwaltung
Version = 0.0.7
ThemeColor = #0078D7
DarkMode = 0

[Paths]
LogPath = $logPath

[UI]
HeaderLogoURL = https://www.microsoft.com/de-de/microsoft-365/exchange/email
"@
        
        Set-Content -Path $FilePath -Value $defaultConfig -Encoding UTF8
        return $true
    }
    catch {
        Write-Error "Fehler beim Erstellen der Standard-Konfigurationsdatei: $($_.Exception.Message)"
        return $false
    }
}

function Setup-ConfigFileWatcher {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        $folder = Split-Path -Path $FilePath -Parent
        $filename = Split-Path -Path $FilePath -Leaf
        
        $watcher = New-Object System.IO.FileSystemWatcher
        $watcher.Path = $folder
        $watcher.Filter = $filename
        $watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::FileName
        
        # Handler für Änderungen definieren
        $onChange = Register-ObjectEvent -InputObject $watcher -EventName Changed -Action {
            try {
                $Global:configUpdated = $true
                
                # Kurze Verzögerung, um sicherzustellen, dass die Datei vollständig geschrieben wurde
                Start-Sleep -Milliseconds 500
                
                # Konfiguration neu laden
                $script:config = Get-IniContent -FilePath $script:configFilePath
                
                # Wenn ein Action-Handler definiert ist, ausführen
                if ($null -ne $script:onConfigChangedAction) {
                    & $script:onConfigChangedAction
                }
            }
            catch {
                # Fehler bei der Aktualisierung der Konfiguration ignorieren
            }
        }
        
        $watcher.EnableRaisingEvents = $true
        
        return $watcher
    }
    catch {
        Write-Error "Fehler beim Einrichten des Konfigurations-File-Watchers: $($_.Exception.Message)"
        return $null
    }
}

function Get-ConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Section,
        
        [Parameter(Mandatory = $true)]
        [string]$Key,
        
        [Parameter(Mandatory = $false)]
        [object]$DefaultValue = $null
    )
    
    try {
        if ($script:config -and $script:config.ContainsKey($Section) -and $script:config[$Section].ContainsKey($Key)) {
            return $script:config[$Section][$Key]
        }
        
        return $DefaultValue
    }
    catch {
        Write-Error "Fehler beim Abrufen des Konfigurationswerts [$Section].$Key: $($_.Exception.Message)"
        return $DefaultValue
    }
}

function Set-ConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Section,
        
        [Parameter(Mandatory = $true)]
        [string]$Key,
        
        [Parameter(Mandatory = $true)]
        [string]$Value
    )
    
    try {
        # Sicherstellen, dass die Sektion existiert
        if (-not $script:config.ContainsKey($Section)) {
            $script:config[$Section] = @{}
        }
        
        # Wert setzen
        $script:config[$Section][$Key] = $Value
        
        # In Datei schreiben
        Set-IniContent -FilePath $script:configFilePath -IniData $script:config
        
        return $true
    }
    catch {
        Write-Error "Fehler beim Setzen des Konfigurationswerts [$Section].$Key: $($_.Exception.Message)"
        return $false
    }
}

function Get-CurrentConfig {
    [CmdletBinding()]
    param()
    
    try {
        return $script:config
    }
    catch {
        Write-Error "Fehler beim Abrufen der aktuellen Konfiguration: $($_.Exception.Message)"
        return @{}
    }
}

function Refresh-ConfigFromFile {
    [CmdletBinding()]
    param()
    
    try {
        $script:config = Get-IniContent -FilePath $script:configFilePath
        return $true
    }
    catch {
        Write-Error "Fehler beim Aktualisieren der Konfiguration aus der Datei: $($_.Exception.Message)"
        return $false
    }
}

function Stop-ConfigWatcher {
    [CmdletBinding()]
    param()
    
    try {
        if ($script:configWatcher -ne $null) {
            $script:configWatcher.EnableRaisingEvents = $false
            $script:configWatcher.Dispose()
            $script:configWatcher = $null
        }
        return $true
    }
    catch {
        Write-Error "Fehler beim Beenden des Konfigurations-File-Watchers: $($_.Exception.Message)"
        return $false
    }
}

# Exportiere Modul-Funktionen
Export-ModuleMember -Function Initialize-EXOConfig, Get-ConfigValue, Set-ConfigValue, Get-CurrentConfig, Refresh-ConfigFromFile, Stop-ConfigWatcher
