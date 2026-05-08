@{
    RootModule        = 'MaintWindowMgrCommon.psm1'
    ModuleVersion     = '1.0.1'
    GUID              = 'a1d2e3f4-5678-9abc-def0-123456789abc'
    Author            = 'Jason Ulbright'
    Description       = 'Maintenance window management for MECM device collections.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging
        'Initialize-Logging'
        'Write-Log'

        # CM Connection
        'Connect-CMSite'
        'Disconnect-CMSite'
        'Test-CMConnection'

        # Data Retrieval
        'Get-AllMaintenanceWindows'
        'Get-CollectionMaintenanceWindows'
        'Get-DeviceCollectionSummary'
        'Get-CollectionsWithoutWindows'
        'ConvertTo-HumanSchedule'
        'Get-NextOccurrences'

        # Folder hierarchy (SMS_ObjectContainerNode)
        'Get-CMCollectionFolderTree'
        'Get-CMCollectionFolderMap'

        # CRUD
        'New-ManagedMaintenanceWindow'
        'Set-ManagedMaintenanceWindow'
        'Remove-ManagedMaintenanceWindow'

        # Schedule Helpers
        'New-WindowSchedule'
        'Get-PatchTuesday'
        'New-PatchTuesdaySchedule'

        # Templates
        'Get-WindowTemplates'
        'Save-WindowTemplate'
        'Remove-WindowTemplate'
        'Apply-WindowTemplate'

        # Bulk Operations
        'Import-MaintenanceWindowsCsv'
        'Copy-MaintenanceWindowToCollections'
        'Set-BulkWindowEnabled'

        # Export
        'Export-MaintenanceWindowsCsv'
        'Export-MaintenanceWindowsHtml'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
