function Find-OrphanedMailEnabledPublicFolder
{
    [cmdletbinding()]
    Param
    (
        [parameter(Mandatory)]
        $ExchangeOrganization
    )
    #Try to get all Mail Enabled Public Folder Objects in the Organization
    try
    {
        $MailEnabledPublicFolders = @(Get-AllMailPublicFolder -ExchangeOrganization $ExchangeOrganization -ErrorAction Stop)
    }
    catch
    {
        $_
        Return
    }
    #Try to get all User Public Folders from the Organization public folder tree
    try
    {
        $ExchangePublicFolders = @(Get-UserPublicFolderTree -ExchangeOrganization $ExchangeOrganization -ErrorAction Stop)
    }
    catch
    {
        $_
        Return
    }
    #Try to get a Mail Enabled Public Folder for each Public Folder
    $splat = @{
        cmdlet               = 'Get-MailPublicFolder'
        ErrorAction          = 'Stop'
        Splat                = @{
            Identity      = ''
            ErrorAction   = 'SilentlyContinue'
            WarningAction = 'SilentlyContinue'
        }
        ExchangeOrganization = $ExchangeOrganization
    }
    $message = "Get-MailPublicFolder for each Public Folder"
    WriteLog -Message $message -EntryType Attempting -Verbose
    $ExchangePublicFoldersMailEnabled = @(
        $PFCount = 0
        foreach ($pf in $ExchangePublicFolders)
        {
            $PFCount++
            $splat.splat.Identity = $pf.ParentPath + '\' + $pf.Name
            Write-Progress -Activity $message -Status "Get-MailPublicFolder -Identity $($splat.Splat.Identity)" -CurrentOperation "$PFCount of $($ExchangePublicFolders.Count)" -PercentComplete $PFCount/$($ExchangePublicFolders.Count)*100
            Invoke-ExchangeCommand @splat
        }
    )
    Write-Progress -Activity $message -Status "Completed" -CurrentOperation "Completed" -PercentComplete 100 -Completed
    WriteLog -Message $message -EntryType Succeeded -Verbose

    $message = 'Build Hashtables to Compare Results of Get-MailPublicFolder with Per Public Folder Get-MailPublicFolder'
    WriteLog -Message $message -EntryType Attempting -Verbose
    $MEPFHashByDN = $MailEnabledPublicFolders | Group-Object -Property DistinguishedName -AsHashTable
    $EPFMEHashByDN = $ExchangePublicFoldersMailEnabled | Group-Object -Property DistinguishedName -AsHashTable
    WriteLog -Message $message -EntryType Succeeded -Verbose
    $message = 'Compare Results of Get-MailPublicFolder with Per Public Folder Get-MailPublicFolder'
    WriteLog -Message $message -EntryType Attempting -Verbose
    $MEPFWithoutEPFME = @(
        foreach ($MEPF in $MailEnabledPublicFolders)
        {
            if (-not $EPFMEHashByDN.ContainsKey($MEPF.DistinguishedName))
            {
                $MEPF
            }
        }
    )
    $EPFMEWithoutMEPF = @(
        foreach ($EPFME in $ExchangePublicFoldersMailEnabled)
        {
            if (-not $MEPFHashByDN.ContainsKey($EPFME.DistinguishedName))
            {
                $EPFME
            }
        }
    )
    if ($EPFMEWithoutMEPF.Count -ge 1)
    {
        $message = "Found Public Folders which are mail enabled but for which no mail enabled public folder object was found with get-mailpublicfolder.  Exporting Data."
        WriteLog -message $message -Verbose
        $file1 = Export-Data -DataToExport $EPFMEWithoutMEPF -DataToExportTitle 'PublicFoldersMissingMailEnabledObject' -Depth 3 -DataType json -ReturnExportFilePath
        $file2 = Export-Data -DataToExport $EPFMEWithoutMEPF -DataToExportTitle 'PublicFoldersMissingMailEnabledObject' -DataType csv -ReturnExportFilePath
        WriteLog -Message "Exported Files: $file1,$file2" -Verbose
    }
    if ($MEPFWithoutEPFME.Count -ge 1)
    {
        $message = "Found Mail Enabled Public Folders for which no public folder object was found.  Exporting Data."
        WriteLog -message $message -Verbose
        $file1 = Export-Data -DataToExport $MEPFWithoutEPFME -DataToExportTitle 'MailEnabledPublicFolderMissingPublicFolderObject' -Depth 3 -DataType json -ReturnExportFilePath
        $file2 = Export-Data -DataToExport $MEPFWithoutEPFME -DataToExportTitle 'MailEnabledPublicFolderMissingPublicFolderObject' -DataType csv -ReturnExportFilePath
        WriteLog -Message "Exported Files: $file1,$file2" -Verbose
    }
}
#end function Find-OrphanedMailEnabledPublicFolders
