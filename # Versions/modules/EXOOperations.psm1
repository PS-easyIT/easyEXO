<#
.SYNOPSIS
    Operations-Modul für easyEXO.
.DESCRIPTION
    Stellt Funktionen für Exchange Online-Operationen bereit.
.NOTES
    Version: 1.0
#>

# Modul-Variablen
$script:isConnected = $false

# Prüft, ob das Exchange Online Management-Modul installiert ist
function Test-ExchangeOnlineModuleInstalled {
    [CmdletBinding()]
    param()
    
    try {
        if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
            return $true
        } else {
            return $false
        }
    } 
    catch {
        Write-LogEntry -Message "Fehler bei der Prüfung des ExchangeOnlineManagement-Moduls: $($_.Exception.Message)" -Type Error
        return $false
    }
}

# Installiert das Exchange Online Management-Modul
function Install-ExchangeOnlineModule {
    [CmdletBinding()]
    param()
    
    try {
        if (-not (Test-ExchangeOnlineModuleInstalled)) {
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
            Write-LogEntry -Message "ExchangeOnlineManagement-Modul wurde erfolgreich installiert" -Type Success
            return $true
        }
        else {
            Write-LogEntry -Message "ExchangeOnlineManagement-Modul ist bereits installiert" -Type Info
            return $true
        }
    }
    catch {
        Write-LogEntry -Message "Fehler bei der Installation des ExchangeOnlineManagement-Moduls: $($_.Exception.Message)" -Type Error
        throw $_.Exception
    }
}

# Verbindet mit Exchange Online
function Connect-ExchangeOnlineService {
    [CmdletBinding()]
    param()
    
    try {
        # Status aktualisieren und loggen
        Update-GuiText -TextElement $global:txtStatus -Message "Verbindung wird hergestellt..."
        Write-LogEntry -Message "Verbindungsversuch zu Exchange Online mit ModernAuth" -Type Info
        
        # Prüfen, ob Modul installiert ist
        if (-not (Test-ExchangeOnlineModuleInstalled)) {
            throw "ExchangeOnlineManagement Modul ist nicht installiert. Bitte installieren Sie das Modul über den 'Installiere Module' Button."
        }
        
        # Modul laden
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        
        # ModernAuth-Verbindung herstellen (nutzt automatisch die Standardbrowser-Authentifizierung)
        Connect-ExchangeOnline -ShowBanner:$false -ShowProgress $true -ErrorAction Stop
        
        # Bei erfolgreicher Verbindung
        Write-LogEntry -Message "Exchange Online Verbindung hergestellt mit ModernAuth (MFA)" -Type Success
        Update-GuiText -TextElement $global:txtStatus -Message "Mit Exchange verbunden (MFA)" -Color $global:connectedBrush
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Update-GuiText -TextElement $global:txtStatus -Message "Fehler beim Verbinden: $errorMsg"
        Write-LogEntry -Message "Fehler beim Verbinden: $errorMsg" -Type Error
        
        # Zeige Fehlermeldung an den Benutzer
        try {
            [System.Windows.MessageBox]::Show(
                "Fehler bei der Verbindung zu Exchange Online:`n$errorMsg", 
                "Verbindungsfehler", 
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
        catch {
            # Fallback, falls MessageBox fehlschlägt
            Write-Host "Fehler bei der Verbindung zu Exchange Online: $errorMsg" -ForegroundColor Red
        }
        
        return $false
    }
}

# Trennt die Verbindung zu Exchange Online
function Disconnect-ExchangeOnlineService {
    [CmdletBinding()]
    param()
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-LogEntry -Message "Exchange Online Verbindung getrennt" -Type Info
        Update-GuiText -TextElement $global:txtStatus -Message "Exchange Verbindung getrennt"
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Update-GuiText -TextElement $global:txtStatus -Message "Fehler beim Trennen der Verbindung: $errorMsg"
        Write-LogEntry -Message "Fehler beim Trennen der Verbindung: $errorMsg" -Type Error
        
        # Zeige Fehlermeldung an den Benutzer
        try {
            [System.Windows.MessageBox]::Show(
                "Fehler beim Trennen der Verbindung zu Exchange Online:`n$errorMsg", 
                "Verbindungsfehler", 
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
        catch {
            # Fallback, falls MessageBox fehlschlägt
            Write-Host "Fehler beim Trennen der Verbindung: $errorMsg" -ForegroundColor Red
        }
        
        return $false
    }
}

# Prüft, ob eine Exchange Online Verbindung besteht
function Test-EXOConnection {
    [CmdletBinding()]
    param()
    
    try {
        # Versuche einen einfachen Exchange-Befehl auszuführen
        Get-AcceptedDomain -ErrorAction Stop | Out-Null
        $script:isConnected = $true
        return $true
    }
    catch {
        $script:isConnected = $false
        return $false
    }
}

# Liefert den aktuellen Verbindungsstatus zurück
function Get-EXOConnectionState {
    [CmdletBinding()]
    param()
    
    return $script:isConnected
}

# -------------------------------------------------
# Abschnitt: Kalenderberechtigungen
# -------------------------------------------------
function Get-CalendarPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Test-EmailAddress -Email $MailboxUser)) {
            throw "Ungültige E-Mail-Adresse für Postfach."
        }
        
        Write-DebugMessage -Message "Rufe Kalenderberechtigungen ab für: $MailboxUser" -Type Info
        
        # Prüfe deutsche und englische Kalenderordner
        $permissions = $null
        try {
            # Versuche mit deutschem Pfad
            $identity = "${MailboxUser}:\Kalender"
            Write-DebugMessage -Message "Versuche deutschen Kalenderpfad: $identity" -Type Info
            $permissions = Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop
        } 
        catch {
            try {
                # Versuche mit englischem Pfad
                $identity = "${MailboxUser}:\Calendar"
                Write-DebugMessage -Message "Versuche englischen Kalenderpfad: $identity" -Type Info
                $permissions = Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop
            } 
            catch {
                # Wenn beide fehlschlagen, wirf den ursprünglichen Fehler
                throw $_.Exception
            }
        }
        
        Write-DebugMessage -Message "Kalenderberechtigungen abgerufen: $($permissions.Count) Einträge gefunden" -Type Success
        Write-LogEntry -Message "Kalenderberechtigungen für $MailboxUser erfolgreich abgerufen: $($permissions.Count) Einträge." -Type Info
        return $permissions
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage -Message "Fehler beim Abrufen der Kalenderberechtigungen: $errorMsg" -Type Error
        Write-LogEntry -Message "Fehler beim Abrufen der Kalenderberechtigungen: $errorMsg" -Type Error
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
        if (-not (Test-EmailAddress -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Test-EmailAddress -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage -Message "Füge Kalenderberechtigung hinzu/aktualisiere: $SourceUser -> $TargetUser ($Permission)" -Type Info
        
        # Prüfe ob Berechtigung bereits existiert
        $calendarExists = $false
        $identityDE = "${SourceUser}:\Kalender"
        $identityEN = "${SourceUser}:\Calendar"
        $identity = $null
        
        try {
            # Versuche vorhandene Berechtigung mit Get-MailboxFolderPermission zu prüfen
            Write-DebugMessage -Message "Prüfe bestehende Berechtigungen (DE): $identityDE" -Type Info
            $existingPerm = Get-MailboxFolderPermission -Identity $identityDE -User $TargetUser -ErrorAction SilentlyContinue
            
            if ($existingPerm) {
                $calendarExists = $true
                $identity = $identityDE
            } 
            else {
                # Versuche englischen Kalender
                Write-DebugMessage -Message "Prüfe bestehende Berechtigungen (EN): $identityEN" -Type Info
                $existingPerm = Get-MailboxFolderPermission -Identity $identityEN -User $TargetUser -ErrorAction SilentlyContinue
                
                if ($existingPerm) {
                    $calendarExists = $true
                    $identity = $identityEN
                }
            }
        } 
        catch {
            # Bei Fehler wird automatisch Add-MailboxFolderPermission versucht
            $errorMsg = $_.Exception.Message
            Write-DebugMessage -Message "Fehler bei der Prüfung bestehender Berechtigungen: $errorMsg" -Type Warning
            
            Write-DebugMessage -Message "Versuche Kalender-Existenz zu prüfen" -Type Info
            if (Get-MailboxFolderPermission -Identity $identityDE -ErrorAction SilentlyContinue) {
                $identity = $identityDE
            } 
            else if (Get-MailboxFolderPermission -Identity $identityEN -ErrorAction SilentlyContinue) {
                $identity = $identityEN
            }
            else {
                throw "Konnte keinen Kalenderordner für $SourceUser finden."
            }
        }
        
        # Je nachdem ob Berechtigung existiert, update oder add
        if ($calendarExists) {
            Write-DebugMessage -Message "Aktualisiere bestehende Berechtigung: $identity ($Permission)" -Type Info
            Set-MailboxFolderPermission -Identity $identity -User $TargetUser -AccessRights $Permission -ErrorAction Stop
            
            if ($null -ne $global:txtStatus) {
                Update-GuiText -TextElement $global:txtStatus -Message "Kalenderberechtigung aktualisiert." -Color $global:connectedBrush
            }
            
            Write-DebugMessage -Message "Kalenderberechtigung erfolgreich aktualisiert" -Type Success
            Write-LogEntry -Message "Kalenderberechtigung aktualisiert: $SourceUser -> $TargetUser mit $Permission" -Type Success
        } 
        else {
            Write-DebugMessage -Message "Füge neue Berechtigung hinzu: $identity ($Permission)" -Type Info
            Add-MailboxFolderPermission -Identity $identity -User $TargetUser -AccessRights $Permission -ErrorAction Stop
            
            if ($null -ne $global:txtStatus) {
                Update-GuiText -TextElement $global:txtStatus -Message "Kalenderberechtigung hinzugefügt." -Color $global:connectedBrush
            }
            
            Write-DebugMessage -Message "Kalenderberechtigung erfolgreich hinzugefügt" -Type Success
            Write-LogEntry -Message "Kalenderberechtigung hinzugefügt: $SourceUser -> $TargetUser mit $Permission" -Type Success
        }
        
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage -Message "Fehler beim Hinzufügen/Aktualisieren der Kalenderberechtigung: $errorMsg" -Type Error
        Write-LogEntry -Message "Fehler beim Hinzufügen/Aktualisieren der Kalenderberechtigung: $errorMsg" -Type Error
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
        if (-not (Test-EmailAddress -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Test-EmailAddress -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage -Message "Entferne Kalenderberechtigung: $SourceUser -> $TargetUser" -Type Info
        
        # Prüfe deutsche und englische Kalenderordner
        $removed = $false
        
        try {
            $identityDE = "${SourceUser}:\Kalender"
            Write-DebugMessage -Message "Prüfe deutsche Kalenderberechtigungen: $identityDE" -Type Info
            
            # Prüfe ob Berechtigung existiert
            $existingPerm = Get-MailboxFolderPermission -Identity $identityDE -User $TargetUser -ErrorAction SilentlyContinue
            
            if ($existingPerm) {
                Remove-MailboxFolderPermission -Identity $identityDE -User $TargetUser -Confirm:$false -ErrorAction Stop
                $removed = $true
                Write-DebugMessage -Message "Kalenderberechtigung aus deutschem Kalender entfernt" -Type Success
            }
            else {
                Write-DebugMessage -Message "Keine Berechtigung im deutschen Kalender gefunden" -Type Info
            }
        } 
        catch {
            $errorMsg = $_.Exception.Message
            Write-DebugMessage -Message "Fehler beim Entfernen der deutschen Kalenderberechtigungen: $errorMsg" -Type Warning
            # Bei Fehler einfach weitermachen und englischen Pfad versuchen
        }
        
        if (-not $removed) {
            try {
                $identityEN = "${SourceUser}:\Calendar"
                Write-DebugMessage -Message "Prüfe englische Kalenderberechtigungen: $identityEN" -Type Info
                
                # Prüfe ob Berechtigung existiert
                $existingPerm = Get-MailboxFolderPermission -Identity $identityEN -User $TargetUser -ErrorAction SilentlyContinue
                
                if ($existingPerm) {
                    Remove-MailboxFolderPermission -Identity $identityEN -User $TargetUser -Confirm:$false -ErrorAction Stop
                    $removed = $true
                    Write-DebugMessage -Message "Kalenderberechtigung aus englischem Kalender entfernt" -Type Success
                }
                else {
                    Write-DebugMessage -Message "Keine Berechtigung im englischen Kalender gefunden" -Type Info
                }
            } 
            catch {
                $errorMsg = $_.Exception.Message
                Write-DebugMessage -Message "Fehler beim Entfernen der englischen Kalenderberechtigungen: $errorMsg" -Type Warning
            }
        }
        
        if ($removed) {
            Write-LogEntry -Message "Kalenderberechtigung entfernt: $SourceUser -> $TargetUser" -Type Success
            return $true
        } 
        else {
            Write-DebugMessage -Message "Keine Kalenderberechtigung zum Entfernen gefunden" -Type Warning
            Write-LogEntry -Message "Keine Kalenderberechtigung gefunden zum Entfernen: $SourceUser -> $TargetUser" -Type Warning
            return $false
        }
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage -Message "Fehler beim Entfernen der Kalenderberechtigung: $errorMsg" -Type Error
        Write-LogEntry -Message "Fehler beim Entfernen der Kalenderberechtigung: $errorMsg" -Type Error
        return $false
    }
}

# -------------------------------------------------
# Abschnitt: Postfachberechtigungen
# -------------------------------------------------
function Add-MailboxFullAccessPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUser,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Test-EmailAddress -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Test-EmailAddress -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage -Message "Füge Postfachberechtigung hinzu: $SourceUser -> $TargetUser (FullAccess)" -Type Info
        
        # Prüfen, ob die Berechtigung bereits existiert
        $existingPermissions = Get-MailboxPermission -Identity $SourceUser -User $TargetUser -ErrorAction SilentlyContinue
        $fullAccessExists = $existingPermissions | Where-Object { $_.AccessRights -like "*FullAccess*" }
        
        if ($fullAccessExists) {
            Write-DebugMessage -Message "Berechtigung existiert bereits, keine Änderung notwendig" -Type Warning
            Write-LogEntry -Message "Postfachberechtigung bereits vorhanden: $SourceUser -> $TargetUser" -Type Warning
            return $true
        }
        
        # Berechtigung hinzufügen
        Add-MailboxPermission -Identity $SourceUser -User $TargetUser -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop
        
        Write-DebugMessage -Message "Postfachberechtigung erfolgreich hinzugefügt" -Type Success
        Write-LogEntry -Message "Postfachberechtigung hinzugefügt: $SourceUser -> $TargetUser (FullAccess)" -Type Success
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage -Message "Fehler beim Hinzufügen der Postfachberechtigung: $errorMsg" -Type Error
        Write-LogEntry -Message "Fehler beim Hinzufügen der Postfachberechtigung: $errorMsg" -Type Error
        return $false
    }
}

function Remove-MailboxFullAccessPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUser,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUser
    )
    
    try {
        # Eingabevalidierung
        if (-not (Test-EmailAddress -Email $SourceUser)) {
            throw "Ungültige E-Mail-Adresse für Quellpostfach."
        }
        if (-not (Test-EmailAddress -Email $TargetUser)) {
            throw "Ungültige E-Mail-Adresse für Zielbenutzer."
        }
        
        Write-DebugMessage -Message "Entferne Postfachberechtigung: $SourceUser -> $TargetUser" -Type Info
        
        # Prüfen, ob die Berechtigung existiert
        $existingPermissions = Get-MailboxPermission -Identity $SourceUser -User $TargetUser -ErrorAction SilentlyContinue
        if (-not $existingPermissions) {
            Write-DebugMessage -Message "Keine Berechtigung zum Entfernen gefunden" -Type Warning
            Write-LogEntry -Message "Keine Postfachberechtigung zum Entfernen gefunden: $SourceUser -> $TargetUser" -Type Warning
            return $false
        }
        
        # Berechtigung entfernen
        Remove-MailboxPermission -Identity $SourceUser -User $TargetUser -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
        
        Write-DebugMessage -Message "Postfachberechtigung erfolgreich entfernt" -Type Success
        Write-LogEntry -Message "Postfachberechtigung entfernt: $SourceUser -> $TargetUser" -Type Success
        return $true
    } 
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage -Message "Fehler beim Entfernen der Postfachberechtigung: $errorMsg" -Type Error
        Write-LogEntry -Message "Fehler beim Entfernen der Postfachberechtigung: $errorMsg" -Type Error
        return $false
    }
}

function Get-MailboxPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxUser
    )
    
    try {
        Write-DebugMessage -Message "Postfachberechtigungen abrufen: Validiere Benutzereingabe" -Type Info
        
        # E-Mail-Format überprüfen
        if (-not (Test-EmailAddress -Email $MailboxUser)) {
            if (-not ($MailboxUser -match "^[a-zA-Z0-9\s.-]+$")) {
                throw "Ungültiges Format für Postfach. Bitte geben Sie eine gültige E-Mail-Adresse oder einen Benutzernamen ein."
            }
        }
        
        Write-DebugMessage -Message "Postfachberechtigungen abrufen für: $MailboxUser" -Type Info
        
        # Postfachberechtigungen abrufen
        Write-DebugMessage -Message "Rufe Postfachberechtigungen ab für: $MailboxUser" -Type Info
        $permissions = Get-MailboxPermission -Identity $MailboxUser | Where-Object { 
            $_.User -notlike "NT AUTHORITY\SELF" -and 
            $_.User -notlike "S-1-5*" -and 
            $_.User -notlike "NT AUTHORITY\SYSTEM" -and
            $_.IsInherited -eq $false 
        }
        
        # SendAs-Berechtigungen abrufen
        $sendAsPermissions = Get-RecipientPermission -Identity $MailboxUser | Where-Object { 
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
                Identity = $perm.Identity
                User = $perm.User
                AccessRights = $perm.AccessRights -join ", "
                SendAs = $hasSendAs
                IsInherited = $perm.IsInherited
            }
            
            Write-DebugMessage -Message "Postfachberechtigung verarbeitet: $($perm.User) -> $($perm.AccessRights -join ', ')" -Type Info
            $resultCollection += $entry
        }
        
        # SendAs-Berechtigungen hinzufügen, die nicht bereits in Postfachberechtigungen enthalten sind
        foreach ($sendPerm in $sendAsPermissions) {
            $existingEntry = $resultCollection | Where-Object { $_.User -eq $sendPerm.Trustee }
            
            if ($null -eq $existingEntry) {
                $entry = [PSCustomObject]@{
                    Identity = $sendPerm.Identity
                    User = $sendPerm.Trustee
                    AccessRights = "SendAs"
                    SendAs = $true
                    IsInherited = $false
                }
                $resultCollection += $entry
            }
        }
        
        # Wenn keine Berechtigungen gefunden wurden
        if ($resultCollection.Count -eq 0) {
            # Füge "NT AUTHORITY\SELF" hinzu, der normalerweise vorhanden ist
            $selfPerm = Get-MailboxPermission -Identity $MailboxUser | Where-Object { $_.User -like "NT AUTHORITY\SELF" } | Select-Object -First 1
            
            if ($null -ne $selfPerm) {
                $entry = [PSCustomObject]@{
                    Identity = $selfPerm.Identity
                    User = "NT AUTHORITY\SELF (Standard)"
                    AccessRights = "FullAccess"
                    SendAs = $false
                    IsInherited = $selfPerm.IsInherited
                }
                $resultCollection += $entry
            }
            else {
                $entry = [PSCustomObject]@{
                    Identity = $MailboxUser
                    User = "Keine benutzerdefinierten Berechtigungen gefunden"
                    AccessRights = "Nur Standardberechtigungen"
                    SendAs = $false
                    IsInherited = $true
                }
                $resultCollection += $entry
            }
        }
        
        Write-DebugMessage -Message "Postfachberechtigungen abgerufen und verarbeitet: $($resultCollection.Count) Einträge gefunden" -Type Success
        
        # Wichtig: Rückgabe als Array für die GUI-Darstellung
        return ,$resultCollection
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage -Message "Fehler beim Abrufen der Postfachberechtigungen: $errorMsg" -Type Error
        Write-LogEntry -Message "Fehler beim Abrufen der Postfachberechtigungen für $MailboxUser`: $errorMsg" -Type Error
        return @()
    }
}

# -------------------------------------------------
# Abschnitt: E-Mail-Validierung
# -------------------------------------------------
function Test-EmailAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email
    )
    
    try {
        if ([string]::IsNullOrEmpty($Email)) {
            return $false
        }
        
        # Einfache Validierung für E-Mail-Format
        $regex = '^[\w\.\-]+@([\w\-]+\.)+[a-zA-Z]{2,}$'
        return $Email -match $regex
    }
    catch {
        Write-DebugMessage -Message "Fehler bei der E-Mail-Validierung: $($_.Exception.Message)" -Type Error
        return $false
    }
}

# Exportiere Modul-Funktionen
Export-ModuleMember -Function Test-ExchangeOnlineModuleInstalled, Install-ExchangeOnlineModule, Connect-ExchangeOnlineService, Disconnect-ExchangeOnlineService, 
                             Test-EXOConnection, Get-EXOConnectionState, Get-CalendarPermissions,
                             Add-CalendarPermission, Remove-CalendarPermission, Add-MailboxFullAccessPermission,
                             Remove-MailboxFullAccessPermission, Get-MailboxPermissions, Test-EmailAddress