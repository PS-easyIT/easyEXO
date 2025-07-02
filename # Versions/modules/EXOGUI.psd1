@{
    RootModule = 'EXOGUI.psm1'
    ModuleVersion = '1.0.0'
    GUID = '5e6c7f85-42c6-4b53-a1c6-94d8a79c9cc1'
    Author = 'Andreas Hepp'
    CompanyName = 'psscripts.de'
    Copyright = '(c) 2025 Andreas Hepp - Alle Rechte vorbehalten.'
    Description = 'GUI-Modul für easyEXO - stellt Funktionen für die Benutzeroberfläche bereit'
    PowerShellVersion = '5.1'
    RequiredModules = @()
    FunctionsToExport = @(
        'Initialize-EXOGUI',
        'Load-XAML',
        'Initialize-ExchangeOnlineLimits',
        'Initialize-GroupsTab'
    )
    CmdletsToExport = @()
    VariablesToExport = '*'
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('GUI', 'WPF', 'Exchange')
            LicenseUri = ''
            ProjectUri = ''
            ReleaseNotes = 'Initiale Version des GUI-Moduls'
        }
    }
}
