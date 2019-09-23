Function Export-PFMMailPublicFolder
{
    [cmdletbinding()]
    param(
    )
    $allMailPublicFolders = @(Get-PFMAllMailPublicFolder)
    $ExportFile = Export-Data -DataToExport $allMailPublicFolders -DataToExportTitle MailPublicFolders -Depth 3 -DataType json -ReturnExportFilePath -ErrorAction Stop
    Write-Information -MessageData "Exported MailPublicFolders File: $ExportFile" -Tags Notification
}