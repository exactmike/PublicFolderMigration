Function Export-PFMUserPublicFolderTree
{
    [cmdletbinding()]
    param(
        [switch]$ExportPermissions
    )
    $allUserPublicFolders = @(Get-PFMUserPublicFolderTree -ExchangeOrganization $ExchangeOrganization)
    $ExportFile = Export-Data -DataToExport $allUserPublicFolders -DataToExportTitle UserPublicFolderTree -Depth 3 -DataType json -ReturnExportFilePath -ErrorAction Stop
    Write-Information -MessageData "Exported UserPublicFolderTree File: $ExportFile" -Tags Notification
    if ($ExportPermissions)
    {
        $GetPublicFolderClientPermissionParams = @{
            Identity    = ''
            ErrorAction = 'Stop'
        }
        $PublicFolderUserPermissions = @(
            $allUserPublicFolders | ForEach-Object
            {
                $GetPublicFolderClientPermissionParams.Identity = $_.EntryID
                Get-PublicFolderClientPermission @GetPublicFolderClientPermissionParams | Select-Object -Property Identity, User -ExpandProperty AccessRights
            }
        )
        $ExportFile = Export-Data -DataToExport $PublicFolderUserPermissions -DataToExportTitle UserPublicFolderPermissions -Depth 1 -DataType csv -ReturnExportFilePath
        Write-Information -MessageData "Exported UserPublicFolderPermissions File: $ExportFile" -Tags Notification
    }
}
