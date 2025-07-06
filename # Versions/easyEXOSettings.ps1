# easyEXOSettings.ps1
# Module for Exchange Online organization settings management
# Part of the easyEXO project

<#
.SYNOPSIS
    Module for Exchange Online organization settings management.
.DESCRIPTION
    Provides functions for managing Exchange Online organization settings.
    Implements a GUI tab in the easyEXO tool for managing Set-OrganizationConfig settings.
.NOTES
    Part of the easyEXO project
    Author: PhinIT
#>

# Import required modules and assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

#region Global Variables
# -----------------------------------------------
# Global Variables
# -----------------------------------------------
$script:tabEXOSettings = $null
$script:organizationConfigSettings = @{}
$script:currentOrganizationConfig = $null
$script:txtStatus = $null
$script:LoggingEnabled = $false
$script:LogFilePath = Join-Path -Path $PSScriptRoot -ChildPath "logs\easyEXOSettings.log"
#endregion Global Variables

#region Main Functions
# -----------------------------------------------
# Main Functions
# -----------------------------------------------

#region Tab Initialization
<#
.SYNOPSIS
    Initializes the EXO Settings tab.
.DESCRIPTION
    Sets up event handlers and initializes UI elements for the Exchange Online settings tab.
.PARAMETER TabItem
    The TabItem control that contains the EXO settings UI.
.RETURNS
    Boolean indicating success or failure of initialization.
#>
function Initialize-EXOSettingsTab {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.TabItem]$TabItem
    )
    
    try {
        Write-DebugMessage "Initializing EXO Settings tab" -Type "Info"
        $script:tabEXOSettings = $TabItem
        
        # Find status textfield
        $script:txtStatus = $TabItem.FindName("txtStatus")
        if ($null -eq $script:txtStatus) {
            # Try to find global status field if tab-specific one is not available
            $script:txtStatus = $TabItem.Parent.Parent.FindName("txtStatus")
        }
        
        # Event handler for help link
        if ($null -ne $TabItem.FindName("helpLinkEXOSettings")) {
            $helpLinkEXOSettings = $TabItem.FindName("helpLinkEXOSettings")
            $helpLinkEXOSettings.Add_MouseLeftButtonDown({
                Start-Process "https://learn.microsoft.com/de-de/powershell/module/exchange/set-organizationconfig?view=exchange-ps"
            })
        }
        
        # Event handler for "Load Current Settings" button
        if ($null -ne $TabItem.FindName("btnGetOrganizationConfig")) {
            $btnGetOrganizationConfig = $TabItem.FindName("btnGetOrganizationConfig")
            $btnGetOrganizationConfig.Add_Click({
                Get-CurrentOrganizationConfig
            })
        }
        
        # Event handler for "Save Settings" button
        if ($null -ne $TabItem.FindName("btnSetOrganizationConfig")) {
            $btnSetOrganizationConfig = $TabItem.FindName("btnSetOrganizationConfig")
            $btnSetOrganizationConfig.Add_Click({
                Set-CustomOrganizationConfig
            })
        }
        
        # Event handler for "Export Configuration" button
        if ($null -ne $TabItem.FindName("btnExportOrgConfig")) {
            $btnExportOrgConfig = $TabItem.FindName("btnExportOrgConfig")
            $btnExportOrgConfig.Add_Click({
                Export-OrganizationConfig
            })
        }
        
        # Initialize additional UI elements
        Initialize-OrganizationConfigCheckboxes
        
        # Create log directory if logging is enabled
        if ($script:LoggingEnabled) {
            $logDirectory = Split-Path -Path $script:LogFilePath -Parent
            if (-not (Test-Path -Path $logDirectory)) {
                New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
            }
        }
        
        Write-DebugMessage "EXO Settings tab was successfully initialized" -Type "Success"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Error initializing EXO Settings tab: $errorMsg" -Type "Error"
        return $false
    }
}
#endregion Tab Initialization

#region UI Elements Initialization
<#
.SYNOPSIS
    Initializes all checkboxes and controls in the OrganizationConfig tab.
.DESCRIPTION
    Sets up event handlers for all UI elements that are used to modify Exchange Online organization settings.
#>
function Initialize-OrganizationConfigCheckboxes {
    [CmdletBinding()]
    param()
    
    try {
        Write-DebugMessage "Initializing OrganizationConfig UI elements" -Type "Info"
        
        # Process all tab panels with checkbox objects
        $tabOrgSettings = $script:tabEXOSettings.FindName("tabOrgSettings")
        if ($null -eq $tabOrgSettings) {
            throw "TabControl 'tabOrgSettings' not found"
        }
        
        # Process all TabItems
        foreach ($tabItem in $tabOrgSettings.Items) {
            $panel = $tabItem.Content
            
            # Find all StackPanels in Grid
            $grid = $panel -as [System.Windows.Controls.Grid]
            if ($null -ne $grid) {
                foreach ($child in $grid.Children) {
                    if ($child -is [System.Windows.Controls.StackPanel]) {
                        # Find all checkboxes in StackPanel and assign event handlers
                        foreach ($control in $child.Children) {
                            if ($control -is [System.Windows.Controls.CheckBox]) {
                                $checkbox = $control
                                
                                # Event handler for checkbox changes
                                $checkbox.Add_Checked({
                                    $cb = $_.Source
                                    $paramName = $cb.Name.Replace("chk", "")
                                    $script:organizationConfigSettings[$paramName] = $true
                                    Write-DebugMessage "Checkbox $($cb.Name) activated: $paramName = $true" -Type "Info"
                                })
                                
                                $checkbox.Add_Unchecked({
                                    $cb = $_.Source
                                    $paramName = $cb.Name.Replace("chk", "")
                                    $script:organizationConfigSettings[$paramName] = $false
                                    Write-DebugMessage "Checkbox $($cb.Name) deactivated: $paramName = $false" -Type "Info"
                                })
                            }
                            
                            # Process ComboBoxes
                            if ($control -is [System.Windows.Controls.ComboBox]) {
                                $comboBox = $control
                                
                                # Event handler for ComboBox changes
                                $comboBox.Add_SelectionChanged({
                                    $cmb = $_.Source
                                    $paramName = $cmb.Name.Replace("cmb", "")
                                    $selectedItem = $cmb.SelectedItem
                                    if ($null -ne $selectedItem) {
                                        $value = $selectedItem.Content.ToString()
                                        # For timeout values, extract only the time value (e.g. "01:00:00 (1h)" -> "01:00:00")
                                        if ($value -match '(\d{2}:\d{2}:\d{2})') {
                                            $value = $matches[1]
                                        }
                                        $script:organizationConfigSettings[$paramName] = $value
                                        Write-DebugMessage "ComboBox $($cmb.Name) changed: $paramName = $value" -Type "Info"
                                    }
                                })
                            }
                            
                            # Process TextBoxes for input fields
                            if ($control -is [System.Windows.Controls.TextBox]) {
                                $textBox = $control
                                
                                # Event handler for TextBox changes
                                $textBox.Add_TextChanged({
                                    $txt = $_.Source
                                    $paramName = $txt.Name.Replace("txt", "")
                                    $script:organizationConfigSettings[$paramName] = $txt.Text
                                    Write-DebugMessage "TextBox $($txt.Name) changed: $paramName = $($txt.Text)" -Type "Info"
                                })
                            }
                        }
                    }
                }
            }
        }
        
        # Special handling for ActivityTimeout
        $cmbActivityTimeout = $script:tabEXOSettings.FindName("cmbActivityTimeout")
        if ($null -ne $cmbActivityTimeout) {
            $cmbActivityTimeout.Add_SelectionChanged({
                $cmb = $_.Source
                $selectedItem = $cmb.SelectedItem
                if ($null -ne $selectedItem) {
                    $timeoutValue = ($selectedItem.Content -split ' ')[0] # Extract "01:00:00" from "01:00:00 (1h)"
                    $script:organizationConfigSettings["ActivityBasedAuthenticationTimeout"] = $timeoutValue
                    Write-DebugMessage "Timeout interval changed: ActivityBasedAuthenticationTimeout = $timeoutValue" -Type "Info"
                }
            })
        }
        
        # Special handling for LargeAudienceThreshold
        $cmbLargeAudienceThreshold = $script:tabEXOSettings.FindName("cmbLargeAudienceThreshold")
        if ($null -ne $cmbLargeAudienceThreshold) {
            $cmbLargeAudienceThreshold.Add_SelectionChanged({
                $cmb = $_.Source
                $selectedItem = $cmb.SelectedItem
                if ($null -ne $selectedItem) {
                    $thresholdValue = [int]$selectedItem.Content
                    $script:organizationConfigSettings["MailTipsLargeAudienceThreshold"] = $thresholdValue
                    Write-DebugMessage "Large audience threshold changed: MailTipsLargeAudienceThreshold = $thresholdValue" -Type "Info"
                }
            })
        }
        
        # Special handling for Rate Limit settings
        $txtPowerShellMaxConcurrency = $script:tabEXOSettings.FindName("txtPowerShellMaxConcurrency")
        if ($null -ne $txtPowerShellMaxConcurrency) {
            $txtPowerShellMaxConcurrency.Add_TextChanged({
                $txt = $_.Source
                if ([int]::TryParse($txt.Text, [ref]$null)) {
                    $script:organizationConfigSettings["PowerShellMaxConcurrency"] = [int]$txt.Text
                    Write-DebugMessage "PowerShellMaxConcurrency changed: $($txt.Text)" -Type "Info"
                }
            })
        }
        
        $txtPowerShellMaxCmdletQueueDepth = $script:tabEXOSettings.FindName("txtPowerShellMaxCmdletQueueDepth")
        if ($null -ne $txtPowerShellMaxCmdletQueueDepth) {
            $txtPowerShellMaxCmdletQueueDepth.Add_TextChanged({
                $txt = $_.Source
                if ([int]::TryParse($txt.Text, [ref]$null)) {
                    $script:organizationConfigSettings["PowerShellMaxCmdletQueueDepth"] = [int]$txt.Text
                    Write-DebugMessage "PowerShellMaxCmdletQueueDepth changed: $($txt.Text)" -Type "Info"
                }
            })
        }
        
        $txtPowerShellMaxCmdletsExecutionDuration = $script:tabEXOSettings.FindName("txtPowerShellMaxCmdletsExecutionDuration")
        if ($null -ne $txtPowerShellMaxCmdletsExecutionDuration) {
            $txtPowerShellMaxCmdletsExecutionDuration.Add_TextChanged({
                $txt = $_.Source
                if ([int]::TryParse($txt.Text, [ref]$null)) {
                    $script:organizationConfigSettings["PowerShellMaxCmdletsExecutionDuration"] = [int]$txt.Text
                    Write-DebugMessage "PowerShellMaxCmdletsExecutionDuration changed: $($txt.Text)" -Type "Info"
                }
            })
        }
        
        # Special handling for InformationBarrierMode
        $cmbInformationBarrierMode = $script:tabEXOSettings.FindName("cmbInformationBarrierMode")
        if ($null -ne $cmbInformationBarrierMode) {
            $cmbInformationBarrierMode.Add_SelectionChanged({
                $cmb = $_.Source
                $selectedItem = $cmb.SelectedItem
                if ($null -ne $selectedItem) {
                    $value = $selectedItem.Content.ToString()
                    $script:organizationConfigSettings["InformationBarrierMode"] = $value
                    Write-DebugMessage "Information Barrier Mode changed: $value" -Type "Info"
                }
            })
        }
        
        # Special handling for EwsApplicationAccessPolicy
        $cmbEwsAppAccessPolicy = $script:tabEXOSettings.FindName("cmbEwsAppAccessPolicy")
        if ($null -ne $cmbEwsAppAccessPolicy) {
            $cmbEwsAppAccessPolicy.Add_SelectionChanged({
                $cmb = $_.Source
                $selectedItem = $cmb.SelectedItem
                if ($null -ne $selectedItem) {
                    $value = $selectedItem.Content.ToString()
                    $script:organizationConfigSettings["EwsApplicationAccessPolicy"] = $value
                    Write-DebugMessage "EWS Application Access Policy changed: $value" -Type "Info"
                }
            })
        }
        
        # Special handling for OfficeFeatures
        $cmbOfficeFeatures = $script:tabEXOSettings.FindName("cmbOfficeFeatures")
        if ($null -ne $cmbOfficeFeatures) {
            $cmbOfficeFeatures.Add_SelectionChanged({
                $cmb = $_.Source
                $selectedItem = $cmb.SelectedItem
                if ($null -ne $selectedItem) {
                    $value = $selectedItem.Content.ToString()
                    $script:organizationConfigSettings["OfficeFeatures"] = $value
                    Write-DebugMessage "Office Features changed: $value" -Type "Info"
                }
            })
        }
        
        # Special handling for SearchQueryLanguage
        $cmbSearchQueryLanguage = $script:tabEXOSettings.FindName("cmbSearchQueryLanguage")
        if ($null -ne $cmbSearchQueryLanguage) {
            $cmbSearchQueryLanguage.Add_SelectionChanged({
                $cmb = $_.Source
                $selectedItem = $cmb.SelectedItem
                if ($null -ne $selectedItem) {
                    $value = $selectedItem.Content.ToString()
                    $script:organizationConfigSettings["SearchQueryLanguage"] = $value
                    Write-DebugMessage "Search Query Language changed: $value" -Type "Info"
                }
            })
        }
        
        Write-DebugMessage "OrganizationConfig UI elements successfully initialized" -Type "Success"
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Error initializing OrganizationConfig UI elements: $errorMsg" -Type "Error"
    }
}
#endregion UI Elements Initialization

#region Organization Config Management
<#
.SYNOPSIS
    Retrieves and displays current organization settings.
.DESCRIPTION
    Retrieves the current Exchange Online organization configuration and updates
    the UI controls with the current values.
#>
function Get-CurrentOrganizationConfig {
    [CmdletBinding()]
    param()
    
    try {
        # Check if we're connected to Exchange
        if (-not (Confirm-ExchangeConnection)) {
            Show-MessageBox -Message "Please connect to Exchange Online first." -Title "Not Connected" -Type "Warning"
            return
        }
        
        Write-DebugMessage "Retrieving current organization settings" -Type "Info"
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Loading organization settings..."
        }
        
        # Retrieve organization settings
        $script:currentOrganizationConfig = Get-OrganizationConfig -ErrorAction Stop
        $configProperties = $script:currentOrganizationConfig | Get-Member -MemberType Properties | 
                            Where-Object { $_.Name -notlike "__*" } | 
                            Select-Object -ExpandProperty Name
        
        # Update UI controls with current values
        foreach ($property in $configProperties) {
            $value = $script:currentOrganizationConfig.$property
            
            # Find and update checkbox
            $checkBoxName = "chk$property"
            $checkBox = $script:tabEXOSettings.FindName($checkBoxName)
            if ($null -ne $checkBox -and $checkBox -is [System.Windows.Controls.CheckBox]) {
                # Set checkbox value without triggering events
                $checkBox.IsChecked = [bool]$value
                $script:organizationConfigSettings[$property] = [bool]$value
            }
            
            # Find and update ComboBox
            $comboBoxName = "cmb$property"
            $comboBox = $script:tabEXOSettings.FindName($comboBoxName)
            if ($null -ne $comboBox -and $comboBox -is [System.Windows.Controls.ComboBox]) {
                # Set ComboBox selection
                foreach ($item in $comboBox.Items) {
                    $content = $item.Content.ToString()
                    if ($content.StartsWith($value) -or $content -eq $value) {
                        $comboBox.SelectedItem = $item
                        $script:organizationConfigSettings[$property] = $value
                        break
                    }
                }
            }
            
            # Find and update TextBox
            $textBoxName = "txt$property"
            $textBox = $script:tabEXOSettings.FindName($textBoxName)
            if ($null -ne $textBox -and $textBox -is [System.Windows.Controls.TextBox]) {
                $textBox.Text = $value
                $script:organizationConfigSettings[$property] = $value
            }
        }
        
        # Set additional special ComboBox values
        
        # Set ActivityTimeout correctly
        $cmbActivityTimeout = $script:tabEXOSettings.FindName("cmbActivityTimeout")
        if ($null -ne $cmbActivityTimeout -and $null -ne $script:currentOrganizationConfig.ActivityBasedAuthenticationTimeout) {
            $timeoutValue = $script:currentOrganizationConfig.ActivityBasedAuthenticationTimeout.ToString()
            foreach ($item in $cmbActivityTimeout.Items) {
                if ($item.Content.StartsWith($timeoutValue)) {
                    $cmbActivityTimeout.SelectedItem = $item
                    $script:organizationConfigSettings["ActivityBasedAuthenticationTimeout"] = $timeoutValue
                    break
                }
            }
        }
        
        # Set LargeAudienceThreshold correctly
        $cmbLargeAudienceThreshold = $script:tabEXOSettings.FindName("cmbLargeAudienceThreshold")
        if ($null -ne $cmbLargeAudienceThreshold -and $null -ne $script:currentOrganizationConfig.MailTipsLargeAudienceThreshold) {
            $thresholdValue = $script:currentOrganizationConfig.MailTipsLargeAudienceThreshold.ToString()
            foreach ($item in $cmbLargeAudienceThreshold.Items) {
                if ($item.Content -eq $thresholdValue) {
                    $cmbLargeAudienceThreshold.SelectedItem = $item
                    $script:organizationConfigSettings["MailTipsLargeAudienceThreshold"] = [int]$thresholdValue
                    break
                }
            }
        }
        
        # Set Rate Limit settings
        $txtPowerShellMaxConcurrency = $script:tabEXOSettings.FindName("txtPowerShellMaxConcurrency")
        if ($null -ne $txtPowerShellMaxConcurrency) {
            $txtPowerShellMaxConcurrency.Text = $script:currentOrganizationConfig.PowerShellMaxConcurrency
        }
        
        $txtPowerShellMaxCmdletQueueDepth = $script:tabEXOSettings.FindName("txtPowerShellMaxCmdletQueueDepth")
        if ($null -ne $txtPowerShellMaxCmdletQueueDepth) {
            $txtPowerShellMaxCmdletQueueDepth.Text = $script:currentOrganizationConfig.PowerShellMaxCmdletQueueDepth
        }
        
        $txtPowerShellMaxCmdletsExecutionDuration = $script:tabEXOSettings.FindName("txtPowerShellMaxCmdletsExecutionDuration")
        if ($null -ne $txtPowerShellMaxCmdletsExecutionDuration) {
            $txtPowerShellMaxCmdletsExecutionDuration.Text = $script:currentOrganizationConfig.PowerShellMaxCmdletsExecutionDuration
        }
        
        # Set InformationBarrierMode
        $cmbInformationBarrierMode = $script:tabEXOSettings.FindName("cmbInformationBarrierMode")
        if ($null -ne $cmbInformationBarrierMode -and $null -ne $script:currentOrganizationConfig.InformationBarrierMode) {
            $barrierMode = $script:currentOrganizationConfig.InformationBarrierMode.ToString()
            foreach ($item in $cmbInformationBarrierMode.Items) {
                if ($item.Content.ToString() -eq $barrierMode) {
                    $cmbInformationBarrierMode.SelectedItem = $item
                    $script:organizationConfigSettings["InformationBarrierMode"] = $barrierMode
                    break
                }
            }
        }
        
        # Set EwsApplicationAccessPolicy
        $cmbEwsAppAccessPolicy = $script:tabEXOSettings.FindName("cmbEwsAppAccessPolicy")
        if ($null -ne $cmbEwsAppAccessPolicy -and $null -ne $script:currentOrganizationConfig.EwsApplicationAccessPolicy) {
            $policyValue = $script:currentOrganizationConfig.EwsApplicationAccessPolicy.ToString()
            foreach ($item in $cmbEwsAppAccessPolicy.Items) {
                if ($item.Content.ToString() -eq $policyValue) {
                    $cmbEwsAppAccessPolicy.SelectedItem = $item
                    $script:organizationConfigSettings["EwsApplicationAccessPolicy"] = $policyValue
                    break
                }
            }
        }
        
        # Set OfficeFeatures
        $cmbOfficeFeatures = $script:tabEXOSettings.FindName("cmbOfficeFeatures")
        if ($null -ne $cmbOfficeFeatures -and $null -ne $script:currentOrganizationConfig.OfficeFeatures) {
            $featuresValue = $script:currentOrganizationConfig.OfficeFeatures.ToString()
            foreach ($item in $cmbOfficeFeatures.Items) {
                if ($item.Content.ToString() -eq $featuresValue) {
                    $cmbOfficeFeatures.SelectedItem = $item
                    $script:organizationConfigSettings["OfficeFeatures"] = $featuresValue
                    break
                }
            }
        }
        
        # Set SearchQueryLanguage
        $cmbSearchQueryLanguage = $script:tabEXOSettings.FindName("cmbSearchQueryLanguage")
        if ($null -ne $cmbSearchQueryLanguage -and $null -ne $script:currentOrganizationConfig.SearchQueryLanguage) {
            $languageValue = $script:currentOrganizationConfig.SearchQueryLanguage.ToString()
            foreach ($item in $cmbSearchQueryLanguage.Items) {
                if ($item.Content.ToString() -eq $languageValue) {
                    $cmbSearchQueryLanguage.SelectedItem = $item
                    $script:organizationConfigSettings["SearchQueryLanguage"] = $languageValue
                    break
                }
            }
        }
        
        # Display configuration in TextBox
        $configText = $script:currentOrganizationConfig | Format-List | Out-String
        $txtOrganizationConfig = $script:tabEXOSettings.FindName("txtOrganizationConfig")
        if ($null -ne $txtOrganizationConfig) {
            $txtOrganizationConfig.Text = $configText
        }
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organization settings successfully loaded."
        }
        Write-DebugMessage "Organization settings successfully retrieved and displayed" -Type "Success"
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Error loading organization settings: $errorMsg"
        }
        Write-DebugMessage "Error retrieving organization settings: $errorMsg" -Type "Error"
        Show-MessageBox -Message "Error loading organization settings: $errorMsg" -Title "Error" -Type "Error"
    }
}

<#
.SYNOPSIS
    Saves organization settings based on UI controls.
.DESCRIPTION
    Applies the settings configured in the UI to the Exchange Online organization configuration.
#>
function Set-CustomOrganizationConfig {
    [CmdletBinding()]
    param()
    
    try {
        # Check if we're connected to Exchange
        if (-not (Confirm-ExchangeConnection)) {
            Show-MessageBox -Message "Please connect to Exchange Online first." -Title "Not Connected" -Type "Warning"
            return
        }
        
        # Show confirmation dialog
        $confirmResult = Show-MessageBox -Message "Do you really want to save the organization settings?" -Title "Save Settings" -Type "YesNo"
        if ($confirmResult -ne "Yes") {
            return
        }
        
        Write-DebugMessage "Saving organization settings" -Type "Info"
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Saving organization settings..."
        }
        
        # Prepare parameters for Set-OrganizationConfig
        $params = @{}
        foreach ($key in $script:organizationConfigSettings.Keys) {
            $params[$key] = $script:organizationConfigSettings[$key]
        }
        
        # Update organization settings
        Set-OrganizationConfig @params -ErrorAction Stop
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organization settings successfully saved."
        }
        Write-DebugMessage "Organization settings successfully saved" -Type "Success"
        Show-MessageBox -Message "The organization settings have been successfully saved." -Title "Success" -Type "Info"
        
        # Reload current configuration to see changes
        Get-CurrentOrganizationConfig
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Error saving organization settings: $errorMsg"
        }
        Write-DebugMessage "Error saving organization settings: $errorMsg" -Type "Error"
        Show-MessageBox -Message "Error saving organization settings: $errorMsg" -Title "Error" -Type "Error"
    }
}

<#
.SYNOPSIS
    Exports organization settings to a file.
.DESCRIPTION
    Exports the current Exchange Online organization configuration to either CSV or text file.
#>
function Export-OrganizationConfig {
    [CmdletBinding()]
    param()
    
    try {
        # Check if we're connected to Exchange
        if (-not (Confirm-ExchangeConnection)) {
            Show-MessageBox -Message "Please connect to Exchange Online first." -Title "Not Connected" -Type "Warning"
            return
        }
        
        # Check if current configuration is available
        if ($null -eq $script:currentOrganizationConfig) {
            Get-CurrentOrganizationConfig
        }
        
        # Show SaveFileDialog
        $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*"
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
            # Export as CSV
            $script:currentOrganizationConfig | 
                Select-Object * -ExcludeProperty RunspaceId, PSComputerName, PSShowComputerName, PSSourceJobInstanceId | 
                Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        }
        else {
            # Export as text
            $script:currentOrganizationConfig | Format-List | Out-File -FilePath $exportPath -Encoding utf8
        }
        
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Organization settings exported to $exportPath."
        }
        Write-DebugMessage "Organization settings successfully exported to $exportPath" -Type "Success"
        Show-MessageBox -Message "The organization settings have been successfully exported to '$exportPath'." -Title "Export Successful" -Type "Info"
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($null -ne $script:txtStatus) {
            $script:txtStatus.Text = "Error exporting organization settings: $errorMsg"
        }
        Write-DebugMessage "Error exporting organization settings: $errorMsg" -Type "Error"
        Show-MessageBox -Message "Error exporting organization settings: $errorMsg" -Title "Error" -Type "Error"
    }
}
#endregion Organization Config Management
#endregion Main Functions

#region Helper Functions
# -----------------------------------------------
# Helper Functions
# -----------------------------------------------

<#
.SYNOPSIS
    Verifies the Exchange Online connection.
.DESCRIPTION
    Checks if an active Exchange Online session exists and verifies 
    it by running a simple command.
.RETURNS
    Boolean indicating if a valid connection exists.
#>
function Confirm-ExchangeConnection {
    [CmdletBinding()]
    param()
    
    try {
        # Check if an active Exchange Online session exists
        $exoSession = Get-PSSession | Where-Object { 
            $_.ConfigurationName -eq "Microsoft.Exchange" -and 
            $_.State -eq "Opened" -and 
            $_.Availability -eq "Available" 
        }
        
        if ($null -eq $exoSession) {
            Write-DebugMessage "No active Exchange Online connection found" -Type "Warning"
            return $false
        }
        
        # Try to execute a simple command to verify the connection
        $null = Get-OrganizationConfig -ErrorAction Stop
        
        Write-DebugMessage "Exchange Online connection confirmed" -Type "Info"
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugMessage "Error with Exchange Online connection: $errorMsg" -Type "Error"
        return $false
    }
}

<#
.SYNOPSIS
    Displays message boxes to the user.
.DESCRIPTION
    Shows a Windows MessageBox with specified text, title, and type.
.PARAMETER Message
    The message to display in the MessageBox.
.PARAMETER Title
    The title of the MessageBox.
.PARAMETER Type
    The type of MessageBox to display: Info, Warning, Error, or YesNo.
.RETURNS
    The result of the MessageBox dialog.
#>
function Show-MessageBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Title = "Exchange Online Settings",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "YesNo")]
        [string]$Type = "Info"
    )
    
    $button = [System.Windows.MessageBoxButton]::OK
    $icon = [System.Windows.MessageBoxImage]::Information
    
    switch ($Type) {
        "Warning" { $icon = [System.Windows.MessageBoxImage]::Warning }
        "Error" { $icon = [System.Windows.MessageBoxImage]::Error }
        "YesNo" { 
            $button = [System.Windows.MessageBoxButton]::YesNo
            $icon = [System.Windows.MessageBoxImage]::Question
        }
    }
    
    $result = [System.Windows.MessageBox]::Show($Message, $Title, $button, $icon)
    return $result
}

<#
.SYNOPSIS
    Logs debug messages with timestamp and color coding.
.DESCRIPTION
    Outputs debug messages to the console and optionally logs them to a file.
.PARAMETER Message
    The message to log.
.PARAMETER Type
    The type of message: Info, Warning, Error, or Success.
#>
function Write-DebugMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Debug", "Info", "Warning", "Error", "Success")]
        [string]$Type = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    switch ($Type) {
        "Debug" { $color = "Gray"; $prefix = "DEBUG" }
        "Info" { $color = "Cyan"; $prefix = "INFO" }
        "Warning" { $color = "Yellow"; $prefix = "WARN" }
        "Error" { $color = "Red"; $prefix = "ERR!" }
        "Success" { $color = "Green"; $prefix = "OK!" }
    }
    
    # Output with color and timestamp
    Write-Host "[$timestamp] [$prefix] $Message" -ForegroundColor $color
    
    # Log to file if enabled
    if ($script:LoggingEnabled) {
        try {
            $logEntry = "[$timestamp] [$prefix] $Message"
            Add-Content -Path $script:LogFilePath -Value $logEntry -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            # If logging fails, output error to console but don't interrupt script
            Write-Host "[$timestamp] [ERR!] Failed to write to log file: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

<#
.SYNOPSIS
    Enables or disables debug logging to a file.
.DESCRIPTION
    Controls whether debug messages are logged to a file in addition to console output.
.PARAMETER Enable
    If true, enables logging to a file. If false, disables logging.
.PARAMETER LogFilePath
    Optional path to the log file. If not specified, uses the default path.
#>
function Set-DebugLogging {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Enable,
        
        [Parameter(Mandatory = $false)]
        [string]$LogFilePath
    )
    
    $script:LoggingEnabled = $Enable
    
    if ($PSBoundParameters.ContainsKey('LogFilePath')) {
        $script:LogFilePath = $LogFilePath
    }
    
    if ($Enable) {
        $logDirectory = Split-Path -Path $script:LogFilePath -Parent
        if (-not (Test-Path -Path $logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }
        
        Write-DebugMessage "Debug logging enabled. Log file: $script:LogFilePath" -Type "Info"
    }
    else {
        Write-DebugMessage "Debug logging disabled" -Type "Info"
    }
}
#endregion Helper Functions

# -----------------------------------------------
# Export required functions
# -----------------------------------------------
Export-ModuleMember -Function Initialize-EXOSettingsTab, Get-CurrentOrganizationConfig, Set-CustomOrganizationConfig, Export-OrganizationConfig, Confirm-ExchangeConnection, Set-DebugLogging

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA7SqhvZU+By7x2
# 67Yt7Ip5XMPXFXp7MOKJMu3j0gwqlqCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCCJ2NCG6EvvMfqVkC4PWrqdF5CbwFXADaNH7q5Za0XfnjANBgkqhkiG
# 9w0BAQEFAASCAQBp7EdnlKknzHZSPlRTWyBu2Vd6VTr1d9Hu+1Eh3YAAdVKBIj6G
# iwNKlwn0B5tstKsdwD0tCN4Lpym1aUhllXeVcd5k8/nxMePEhEoaAk3tXLX8++HE
# sSrsVB3Jz3mL4pgscp+9onGXjp9vAQEoKQ2d7qGvlDSnL9s2ujKcYkYHiRtiZw3i
# Zg6GSr8BrypBqRvRh3TwY9S3IF+bFDT9wugyPfKP/w4/sERMRrrIvnhYIxkoFuge
# Dex8yS4Ms0kXsPVSmTre4oGLlquEig59XaubtUk4PhY2pMt0iR8GzPviwlf74M8P
# CsJuikF33cU4kX1EtGz9cBX7G7rQkZENxLV1oYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE0NFowLwYJKoZIhvcNAQkEMSIEIKxf1SwVEMNFbpm84LGUh0GTgI+O
# cXrOkjDW9KekQQhjMA0GCSqGSIb3DQEBAQUABIICAHgL8BfyG8RZI/JK5rYpPo1M
# h29Rh3WlLS5IdXeRh2iB/ys8PAW4mIBkyCTUhWWB2Wr8N0hkCaFUkuveaoH/HIcy
# bVVmAIMbV9N1Nzi8UANz/QZnuW2CoJgzyM1OTJYPgSgRivU6Dk9grM6PIz5OtR6E
# PmdSziWT1gfu9bMe3W6qx9T4rKNkN8b2MuSat8umaGyIzP2CB7YJIz8Wp+uAmukU
# lliBymvtLF/cAdCGzOM5yENgJp+QXeoeNE45+DKCgIRg/L6GASfHmf//3Wogro4t
# BcNulgbT006xEeAeX9iRscZa/y0zQuXtfc69JwsVfaNWjPAMEPuYGacmDmFS+FN3
# xt1WTXtRSVqYH1v8wRDNA21WR3NCBhEyM/T2TFiq23l1uCYFwfiQcka4XFD7Oul0
# /pLJNrqakX2hUn5jmmd/wKGwDhoQa9n7IjgR/fDovzZCMeg3S5FWiNNDOBddrhG3
# tunU37T63h66ALqkPmwhbAMTJZjRM8cHzl/cpHuCYYD6R74z01LSfQpU+w575BU8
# Cd594yF7x6pk59CLp1nXijYNq9rGuC8CC+rUkhPBEMnyScbn/A9wuT2+N9XxLjZj
# TOQhnOVLVfAInd+5DBZJSXTjQMBQRw/QeMDoQs8h74bygdZV5H3rKxZzuJn1NgPH
# aEAK7FTyh1LNA48covQR
# SIG # End signature block
