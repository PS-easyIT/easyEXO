# UIManagement.psm1
# Modul zur Verwaltung der UI-Elemente und Event-Handler mit robuster Fehlerbehandlung

# Element-Cache für die Verwaltung von UI-Elementen
$script:UIElements = @{}

function Register-UIElement {
    <#
    .SYNOPSIS
        Registriert ein UI-Element im globalen Cache.
    .DESCRIPTION
        Fügt ein UI-Element zum Cache hinzu, sodass es später sicher abgerufen werden kann.
        Ermöglicht die Überprüfung auf null-Werte, um Fehler zu vermeiden.
    .PARAMETER Name
        Der Name des UI-Elements.
    .PARAMETER Element
        Das UI-Element selbst.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        
        [Parameter(Mandatory=$true)]
        [System.Object]$Element
    )
    
    try {
        # Element im Cache speichern
        $script:UIElements[$Name] = $Element
        Write-Debug "UI-Element '$Name' erfolgreich registriert."
    }
    catch {
        Write-Warning "Fehler beim Registrieren des UI-Elements '$Name': $_"
    }
}

function Get-UIElement {
    <#
    .SYNOPSIS
        Ruft ein UI-Element aus dem Cache ab.
    .DESCRIPTION
        Sucht ein UI-Element im Cache und gibt es zurück, wenn es existiert.
        Falls das Element nicht existiert, wird null zurückgegeben.
    .PARAMETER Name
        Der Name des UI-Elements.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name
    )
    
    try {
        # Prüfen, ob das Element im Cache existiert
        if ($script:UIElements.ContainsKey($Name)) {
            return $script:UIElements[$Name]
        }
        else {
            Write-Debug "UI-Element '$Name' nicht im Cache gefunden."
            return $null
        }
    }
    catch {
        Write-Warning "Fehler beim Abrufen des UI-Elements '$Name': $_"
        return $null
    }
}

function Register-UIElements {
    <#
    .SYNOPSIS
        Registriert alle UI-Elemente aus einem XAML-Window im Cache.
    .DESCRIPTION
        Durchsucht ein WPF-Window nach UI-Elementen mit Namen und registriert sie im Cache.
    .PARAMETER Window
        Das WPF-Window, das die UI-Elemente enthält.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Window]$Window
    )
    
    try {
        Write-Debug "Registriere alle benannten UI-Elemente aus dem Window..."
        
        # Alle Elemente mit Namen finden und registrieren
        $type = [System.Windows.FrameworkElement]
        $elementsWithName = Get-ChildItemsRecursive -Element $Window -Type $type | 
                           Where-Object { -not [string]::IsNullOrEmpty($_.Name) }
        
        foreach ($element in $elementsWithName) {
            Register-UIElement -Name $element.Name -Element $element
        }
        
        Write-Debug "UI-Elemente erfolgreich registriert. Anzahl: $($elementsWithName.Count)"
    }
    catch {
        Write-Warning "Fehler beim Registrieren der UI-Elemente: $_"
    }
}

function Get-ChildItemsRecursive {
    <#
    .SYNOPSIS
        Gibt alle Kind-Elemente eines WPF-Elements rekursiv zurück.
    .DESCRIPTION
        Durchsucht ein WPF-Element rekursiv nach allen Kind-Elementen eines bestimmten Typs.
    .PARAMETER Element
        Das WPF-Element, das durchsucht werden soll.
    .PARAMETER Type
        Der Typ der Elemente, die gefunden werden sollen.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Object]$Element,
        
        [Parameter(Mandatory=$true)]
        [type]$Type
    )
    
    try {
        # Liste für die gefundenen Elemente
        $items = New-Object System.Collections.ArrayList
        
        # Aktuelle Ebene prüfen
        if ($Element -is $Type) {
            [void]$items.Add($Element)
        }
        
        # ContentControl-Elemente prüfen
        if ($Element -is [System.Windows.Controls.ContentControl]) {
            $content = $Element.Content
            if ($content -is $Type) {
                [void]$items.Add($content)
            }
            
            # Rekursiv nach weiteren Elementen suchen
            if ($content -is [System.Windows.DependencyObject]) {
                $childItems = Get-ChildItemsRecursive -Element $content -Type $Type
                if ($childItems) {
                    [void]$items.AddRange($childItems)
                }
            }
        }
        # Panel-Elemente prüfen
        elseif ($Element -is [System.Windows.Controls.Panel]) {
            foreach ($child in $Element.Children) {
                if ($child -is $Type) {
                    [void]$items.Add($child)
                }
                
                # Rekursiv nach weiteren Elementen suchen
                if ($child -is [System.Windows.DependencyObject]) {
                    $childItems = Get-ChildItemsRecursive -Element $child -Type $Type
                    if ($childItems) {
                        [void]$items.AddRange($childItems)
                    }
                }
            }
        }
        # ItemsControl-Elemente prüfen
        elseif ($Element -is [System.Windows.Controls.ItemsControl]) {
            foreach ($item in $Element.Items) {
                if ($item -is $Type) {
                    [void]$items.Add($item)
                }
                
                # Rekursiv nach weiteren Elementen suchen
                if ($item -is [System.Windows.DependencyObject]) {
                    $childItems = Get-ChildItemsRecursive -Element $item -Type $Type
                    if ($childItems) {
                        [void]$items.AddRange($childItems)
                    }
                }
            }
        }
        
        return $items
    }
    catch {
        Write-Warning "Fehler bei der rekursiven Suche nach Kind-Elementen: $_"
        return $null
    }
}

function Register-SafeEventHandler {
    <#
    .SYNOPSIS
        Registriert einen Event-Handler für ein UI-Element mit Fehlerprüfung.
    .DESCRIPTION
        Registriert einen Event-Handler für ein UI-Element, falls das Element existiert.
        Verhindert Fehler bei nicht existierenden Elementen.
    .PARAMETER ElementName
        Der Name des UI-Elements.
    .PARAMETER Event
        Der Name des Events (z.B. "Click").
    .PARAMETER ScriptBlock
        Der ScriptBlock, der als Event-Handler ausgeführt werden soll.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ElementName,
        
        [Parameter(Mandatory=$true)]
        [string]$Event,
        
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock
    )
    
    try {
        # Element aus dem Cache abrufen
        $element = Get-UIElement -Name $ElementName
        
        # Prüfen, ob das Element existiert
        if ($null -eq $element) {
            Write-Warning "Event-Handler für '$ElementName.$Event' konnte nicht registriert werden. Element nicht gefunden."
            return
        }
        
        # Methode zum Hinzufügen des Event-Handlers ermitteln
        $addMethod = "Add_$Event"
        
        # Prüfen, ob die Methode existiert
        if (-not ($element | Get-Member -Name $addMethod -MemberType Method)) {
            Write-Warning "Element '$ElementName' unterstützt das Event '$Event' nicht."
            return
        }
        
        # Event-Handler mit try-catch Block umhüllen
        $safeScriptBlock = {
            param($sender, $eventArgs)
            
            try {
                # Original ScriptBlock ausführen
                & $ScriptBlock $sender $eventArgs
            }
            catch {
                # Fehler abfangen und protokollieren
                $errorMsg = "Fehler im Event-Handler für '$ElementName.$Event': $_"
                Write-Warning $errorMsg
                
                # Ggf. UI-Feedback für Benutzer
                try {
                    [System.Windows.MessageBox]::Show(
                        "Ein unerwarteter Fehler ist aufgetreten. Bitte wenden Sie sich an den Support.`n`nDetails: $errorMsg",
                        "Fehler",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Error
                    )
                }
                catch {
                    # Ignorieren, falls MessageBox-Fehler
                }
            }
        }
        
        # Event-Handler registrieren
        $element.$addMethod.Invoke($safeScriptBlock)
        Write-Debug "Event-Handler für '$ElementName.$Event' erfolgreich registriert."
    }
    catch {
        Write-Warning "Fehler beim Registrieren des Event-Handlers für '$ElementName.$Event': $_"
    }
}

function Initialize-DynamicElement {
    <#
    .SYNOPSIS
        Erstellt ein dynamisches UI-Element und registriert es im Cache.
    .DESCRIPTION
        Erstellt ein neues UI-Element zur Laufzeit und registriert es im Cache.
        Nützlich für dynamisch erzeugte UI-Elemente, die später über ihren Namen abgerufen werden sollen.
    .PARAMETER Name
        Der Name des zu erstellenden Elements.
    .PARAMETER Type
        Der Typ des Elements (z.B. "System.Windows.Controls.Button").
    .PARAMETER Properties
        Ein Hashtable mit Eigenschaften, die auf das Element angewendet werden sollen.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        
        [Parameter(Mandatory=$true)]
        [string]$Type,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$Properties = @{}
    )
    
    try {
        # Element-Typ laden
        $elementType = $Type -as [type]
        if ($null -eq $elementType) {
            Write-Warning "Unbekannter Element-Typ: $Type"
            return $null
        }
        
        # Neues Element erstellen
        $element = New-Object $elementType
        $element.Name = $Name
        
        # Eigenschaften anwenden
        foreach ($key in $Properties.Keys) {
            try {
                $element.$key = $Properties[$key]
            }
            catch {
                Write-Warning "Fehler beim Setzen der Eigenschaft '$key' für Element '$Name': $_"
            }
        }
        
        # Element im Cache registrieren
        Register-UIElement -Name $Name -Element $element
        
        return $element
    }
    catch {
        Write-Warning "Fehler beim Erstellen des Elements '$Name': $_"
        return $null
    }
}

# Exportieren der Funktionen
Export-ModuleMember -Function Register-UIElement, Get-UIElement, Register-UIElements, 
                             Register-SafeEventHandler, Initialize-DynamicElement