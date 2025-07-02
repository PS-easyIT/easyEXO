<#
.SYNOPSIS
    GUI-Modul für easyEXO.
.DESCRIPTION
    Stellt Funktionen für die graphische Benutzeroberfläche bereit.
.NOTES
    Version: 1.0
#>

# Notwendige Assemblies für WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Modul-Variablen
$script:connectedBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Green)
$script:disconnectedBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Red)
$script:warningBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Orange)
$script:infoBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Blue)

# Initialize-GUI Funktion
function Initialize-EXOGUI {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$XamlFilePath,
        
        [Parameter(Mandatory = $false)]
        [string]$AppName = "Exchange Berechtigungen Verwaltung Tool",
        
        [Parameter(Mandatory = $false)]
        [string]$Version = "1.0.0"
    )
    
    try {
        Write-Debug -Message "Initialisiere GUI aus XAML-Datei: $XamlFilePath" -Type "Info"
        
        # Prüfen, ob die XAML-Datei existiert
        if (-not (Test-Path -Path $XamlFilePath)) {
            throw "XAML-Datei nicht gefunden: $XamlFilePath"
        }
        
        # XAML-Datei laden
        [xml]$xaml = Get-Content -Path $XamlFilePath -Encoding UTF8
        
        # Reader erstellen und XAML laden
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        # Versionsinformation setzen
        $txtVersion = $window.FindName("txtVersion")
        if ($null -ne $txtVersion) {
            $txtVersion.Text = "Version: $Version"
        }
        
        # Fenstertitel setzen
        $window.Title = $AppName
        
        Write-Debug -Message "GUI erfolgreich initialisiert" -Type "Success"
        Write-Log -Message "GUI erfolgreich initialisiert: $AppName, Version $Version" -Level "Info"
        
        return $window
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Debug -Message "Fehler bei der GUI-Initialisierung: $errorMsg" -Type "Error"
        Write-Log -Message "Fehler bei der GUI-Initialisierung: $errorMsg" -Level "Error"
        
        # Fallback-Fehlerausgabe
        try {
            [System.Windows.MessageBox]::Show(
                "Die GUI konnte nicht geladen werden. Bitte stellen Sie sicher, dass die XAML-Datei korrekt ist.`n`nFehler: $errorMsg",
                "Kritischer Fehler",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
        catch {
            # Absoluter Fallback
            Write-Error "KRITISCHER FEHLER: GUI konnte nicht initialisiert werden: $errorMsg"
        }
        
        throw $errorMsg
    }
}

# Funktion zum Aktualisieren der GUI-Text mit Fehlerbehandlung
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
            Write-Debug -Message "GUI-Element ist null in Update-GuiText" -Type "Warning"
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
            Write-Debug -Message "Fehler in Update-GuiText: $errorMsg" -Type "Error"
            Write-Log -Message "GUI-Ausgabefehler: $errorMsg" -Level "Error"
        }
        catch {
            # Ignoriere Fehler in der Fehlerbehandlung
        }
    }
}

# Funktion zum Anzeigen einer Nachricht
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
        Write-Log -Message "MessageBox angezeigt: $Title - $Type - $Message" -Level "Info"
        
        # Ergebnis zurückgeben (wichtig für Ja/Nein-Fragen)
        return $result
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Debug -Message "Fehler beim Anzeigen der MessageBox: $errorMsg" -Type "Error"
        Write-Log -Message "Fehler beim Anzeigen der MessageBox: $errorMsg" -Level "Error"
        
        # Fallback-Ausgabe
        Write-Host "Meldung ($Type): $Title - $Message" -ForegroundColor Red
        
        if ($Type -eq "Question") {
            return [System.Windows.MessageBoxResult]::No
        }
    }
}

# Funktion zum Aktualisieren des Connection Status
function Update-ConnectionStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.TextBlock]$StatusElement,
        
        [Parameter(Mandatory = $true)]
        [bool]$IsConnected
    )
    
    try {
        if ($null -eq $StatusElement) {
            Write-Debug -Message "Status-Element ist null in Update-ConnectionStatus" -Type "Warning"
            return
        }
        
        $StatusElement.Dispatcher.Invoke([Action]{
            if ($IsConnected) {
                $StatusElement.Text = "Verbunden"
                $StatusElement.Foreground = $script:connectedBrush
            } else {
                $StatusElement.Text = "Nicht verbunden"
                $StatusElement.Foreground = $script:disconnectedBrush
            }
        }, "Normal")
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Debug -Message "Fehler in Update-ConnectionStatus: $errorMsg" -Type "Error"
        Write-Log -Message "Fehler beim Aktualisieren des Verbindungsstatus: $errorMsg" -Level "Error"
    }
}

# Funktion zum Laden einer bestimmten Tab-Seite
function Select-TabPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.TabControl]$TabControl,
        
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.TabItem]$TargetTab
    )
    
    try {
        if ($null -eq $TabControl -or $null -eq $TargetTab) {
            Write-Debug -Message "Tab-Elemente sind null in Select-TabPage" -Type "Warning"
            return $false
        }
        
        # Alle TabItems verstecken
        $TabControl.Dispatcher.Invoke([Action]{
            foreach ($tab in $TabControl.Items) {
                if ($tab -is [System.Windows.Controls.TabItem]) {
                    $tab.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }
            
            # Zieltab anzeigen und auswählen
            $TargetTab.Visibility = [System.Windows.Visibility]::Visible
            $TargetTab.IsSelected = $true
        }, "Normal")
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Debug -Message "Fehler in Select-TabPage: $errorMsg" -Type "Error"
        Write-Log -Message "Fehler beim Wechseln der Tab-Seite: $errorMsg" -Level "Error"
        return $false
    }
}

# Exportiere Modul-Funktionen
Export-ModuleMember -Function Initialize-EXOGUI, Update-GuiText, Show-MessageBox, 
                             Update-ConnectionStatus, Select-TabPage -Variable connectedBrush, disconnectedBrush, 
                             warningBrush, infoBrush
