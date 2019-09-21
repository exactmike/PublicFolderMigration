Function ExportResumeID
{

    [CmdletBinding()]
    param
    (
        $ID
        ,
        $nextPermissionID
        ,
        $outputFolderPath
        ,
        $TimeStamp
        ,
        $ResumeIndex
    )
    $ExportFilePath = Join-Path -Path $outputFolderPath -ChildPath $($TimeStamp + "ExchangePermissionExportResumeID.xml")
    $Identities = @{
        NextPermissionIdentity = $nextPermissionID
        ResumeID               = $ID
        ResumeIndex            = $ResumeIndex
    }
    Export-Clixml -Depth 1 -Path $ExportFilePath -InputObject $Identities -Encoding UTF8
    $ExportFilePath

}
