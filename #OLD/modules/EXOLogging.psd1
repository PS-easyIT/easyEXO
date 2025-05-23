@{
    RootModule = 'EXOLogging.psm1'
    ModuleVersion = '1.0.0'
    GUID = '7db13dfb-c9af-4fe7-8f90-200e24a86b77'
    Author = 'Andreas Hepp'
    CompanyName = 'psscripts.de'
    Copyright = '(c) 2025 Andreas Hepp - Alle Rechte vorbehalten.'
    Description = 'Logging-Modul für easyEXO - stellt Funktionen für Debugging und Logging bereit'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Initialize-EXOLogging',
        'Write-DebugMessage',
        'Write-LogEntry',
        'Update-GuiText'
    )
    CmdletsToExport = @()
    VariablesToExport = '*'
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('Logging', 'Debug', 'GUI', 'Exchange')
            LicenseUri = ''
            ProjectUri = ''
            ReleaseNotes = 'Initiale Version des Logging-Moduls'
        }
    }
}
