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
