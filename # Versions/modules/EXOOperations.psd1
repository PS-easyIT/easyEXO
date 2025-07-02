@{
    RootModule = 'EXOOperations.psm1'
    ModuleVersion = '1.0.0'
    GUID = 'b425c9b8-32e0-4b54-a7ea-d55362f89025'
    Author = 'Andreas Hepp'
    CompanyName = 'psscripts.de'
    Copyright = '(c) 2025 Andreas Hepp - Alle Rechte vorbehalten.'
    Description = 'Operations-Modul für easyEXO - stellt Funktionen für Exchange Online Operationen bereit'
    PowerShellVersion = '5.1'
    RequiredModules = @('ExchangeOnlineManagement')
    FunctionsToExport = @(
        'Connect-ExchangeOnlineService',
        'Disconnect-ExchangeOnlineService',
        'Test-ExchangeOnlineModuleInstalled',
        'Install-ExchangeOnlineModule',
        'Get-CalendarPermissions',
        'Add-CalendarPermission',
        'Remove-CalendarPermission',
        'Get-MailboxPermissions',
        'Add-MailboxPermission',
        'Remove-MailboxPermission',
        'Add-SendAsPermission',
        'Remove-SendAsPermission',
        'Add-SendOnBehalfPermission',
        'Remove-SendOnBehalfPermission',
        'Get-PublicFolderPermissions',
        'Add-PublicFolderPermission',
        'Remove-PublicFolderPermission',
        'Search-PublicFolders',
        'Get-ExchangeLimits',
        'Set-ExchangeLimits',
        'Set-DefaultCalendarPermission',
        'Set-AnonymousCalendarPermission',
        'Set-DefaultCalendarPermissionForAll',
        'Set-AnonymousCalendarPermissionForAll',
        'Get-GroupsAction',
        'New-GroupAction',
        'Remove-GroupAction',
        'Add-GroupMemberAction',
        'Remove-GroupMemberAction',
        'Get-GroupMembersAction',
        'Invoke-AuditCommand',
        'Export-AuditResults'
    )
    CmdletsToExport = @()
    VariablesToExport = '*'
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('Exchange', 'EXO', 'Operation', 'Permission')
            LicenseUri = ''
            ProjectUri = ''
            ReleaseNotes = 'Initiale Version des Operations-Moduls'
        }
    }
}
