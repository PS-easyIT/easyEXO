@{
    RootModule = 'EXOConfig.psm1'
    ModuleVersion = '1.0.0'
    GUID = 'c7a35710-5e30-4a17-88b1-e521d6398bef'
    Author = 'Andreas Hepp'
    CompanyName = 'psscripts.de'
    Copyright = '(c) 2025 Andreas Hepp - Alle Rechte vorbehalten.'
    Description = 'Konfigurations-Modul für easyEXO - stellt Funktionen für Konfigurationsmanagement bereit'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Initialize-EXOConfig',
        'Get-ConfigValue',
        'Set-ConfigValue',
        'Get-CurrentConfig',
        'Refresh-ConfigFromFile',
        'Stop-ConfigWatcher',
        'Get-IniContent',
        'Set-IniContent'
    )
    CmdletsToExport = @()
    VariablesToExport = '*'
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('Configuration', 'INI', 'Settings')
            LicenseUri = ''
            ProjectUri = ''
            ReleaseNotes = 'Initiale Version des Konfigurations-Moduls'
        }
    }
}
