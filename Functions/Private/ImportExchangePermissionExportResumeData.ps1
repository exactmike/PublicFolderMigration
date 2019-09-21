Function ImportExchangePermissionExportResumeData
{

    [CmdletBinding()]
    param
    (
        [parameter(Mandatory)]
        $path
    )
    $ImportedExchangePermissionsExportResumeData = Import-Clixml -Path $path -ErrorAction Stop
    $parentpath = Split-Path -Path $path -Parent
    $ResumeIDFilePath = Join-Path -path $parentpath -ChildPath $($ImportedExchangePermissionsExportResumeData.TimeStamp + 'ExchangePermissionExportResumeID.xml')
    $ResumeIDs = Import-Clixml -Path $ResumeIDFilePath -ErrorAction Stop
    $ImportedExchangePermissionsExportResumeData.ResumeID = $ResumeIDs.ResumeID
    $ImportedExchangePermissionsExportResumeData.NextPermissionIdentity = $ResumeIDs.NextPermissionIdentity
    $ImportedExchangePermissionsExportResumeData.ResumeIndex = $ResumeIDs.ResumeIndex
    $ImportedExchangePermissionsExportResumeData

}
