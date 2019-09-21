Function ExportExchangePermissionExportResumeData
{

    [CmdletBinding()]
    param
    (
        $ExchangePermissionsExportParameters
        ,
        $ExcludedPublicFoldersEntryIDHash
        ,
        $ExcludedTrusteeGuidHash
        ,
        $SIDHistoryRecipientHash
        ,
        $InScopeFolders
        ,
        $InScopeMailPublicFoldersHash
        ,
        $ObjectGUIDHash
        ,
        $outputFolderPath
        ,
        $ExportedExchangePublicFolderPermissionsFile
        ,
        $TimeStamp
    )
    GetCallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -Name VerbosePreference
    $ExchangePermissionExportResumeData = @{
        ExchangePermissionsExportParameters         = $ExchangePermissionsExportParameters
        ExcludedPublicFoldersEntryIDHash            = $ExcludedPublicFoldersEntryIDHash
        ExcludedTrusteeGuidHash                     = $ExcludedTrusteeGuidHash
        SIDHistoryRecipientHash                     = $SIDHistoryRecipientHash
        InScopeFolders                              = $InScopeFolders
        InScopeMailPublicFoldersHash                = $InScopeMailPublicFoldersHash
        ObjectGUIDHash                              = $ObjectGUIDHash
        ExportedExchangePublicFolderPermissionsFile = $ExportedExchangePublicFolderPermissionsFile
        TimeStamp                                   = $TimeStamp
    }
    $ExportFilePath = Join-Path -Path $outputFolderPath -ChildPath $($TimeStamp + "ExchangePermissionExportResumeData.xml")
    Export-Clixml -Depth 2 -Path $ExportFilePath -InputObject $ExchangePermissionExportResumeData -Encoding UTF8
    $ExportFilePath

}
