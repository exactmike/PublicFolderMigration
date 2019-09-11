Function Find-OrphanedMailEnabledPublicFolders
{

    [cmdletbinding()]
    Param(
        [parameter(Mandatory)]
        $ExchangeOrganization
    )
    #End Params
    #Try to get all Mail Enabled Public Folder Objects in the Organization
    try
    {
        $MailEnabledPublicFolders = @(Get-AllMailPublicFolder -ErrorAction Stop)
    }
    catch
    {
        $_
        Return
    }
    #Try to get all User Public Folders from the Organization public folder tree
    try
    {
        $ExchangePublicFolders = @(Get-UserPublicFolderTree -ErrorAction Stop)
    }
    catch
    {
        $_
        Return
    }
    #Try to get a Mail Enabled Public Folder for each Public Folder
    $getMailPublicFolderParams = @{
            Identity      = ''
            ErrorAction   = 'SilentlyContinue'
            WarningAction = 'SilentlyContinue'
    }
    $message = "Get-MailPublicFolder for each Public Folder"
    Write-Information -Message $message -Tags Attempting
    $ExchangePublicFoldersMailEnabled = @(
        $PFCount = 0
        foreach ($pf in $ExchangePublicFolders)
        {
            $PFCount++
            $getMailPublicFolderParams.Identity = $pf.ParentPath + '\' + $pf.Name #have to use this format because Get-MailPublicFolder does not accept the unambiguous EntryID as an identity!
            Write-Progress -Activity $message -Status "Get-MailPublicFolder -Identity $($splat.Splat.Identity)" -CurrentOperation "$PFCount of $($ExchangePublicFolders.Count)" -PercentComplete $PFCount/$($ExchangePublicFolders.Count)*100
            Get-MailPublicFolder @splat
        }
    )
    Write-Progress -Activity $message -Status "Completed" -CurrentOperation "Completed" -PercentComplete 100 -Completed
    Write-Information -Message $message -Tags Succeeded

    $message = 'Build Hashtables to Compare Results of Get-MailPublicFolder with Per Public Folder Get-MailPublicFolder'
    Write-Information -Message $message -Tags Attempting
    $MEPFHashByDN = $MailEnabledPublicFolders | Group-Object -Property DistinguishedName -AsHashTable
    $EPFMEHashByDN = $ExchangePublicFoldersMailEnabled | Group-Object -Property DistinguishedName -AsHashTable
    Write-Information -Message $message -Tags Succeeded
    $message = 'Compare Results of Get-MailPublicFolder with Per Public Folder Get-MailPublicFolder'
    Write-Information -Message $message -Tags Attempting
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
        Write-Information -message $message -Tags Notification
        $file1 = Export-Data -DataToExport $EPFMEWithoutMEPF -DataToExportTitle 'PublicFoldersMissingMailEnabledObject' -Depth 3 -DataType json -ReturnExportFilePath
        $file2 = Export-Data -DataToExport $EPFMEWithoutMEPF -DataToExportTitle 'PublicFoldersMissingMailEnabledObject' -DataType csv -ReturnExportFilePath
        Write-Information -Message "Exported Files: $file1,$file2" -Tags Notification
    }
    if ($MEPFWithoutEPFME.Count -ge 1)
    {
        $message = "Found Mail Enabled Public Folders for which no public folder object was found.  Exporting Data."
        Write-Information -message $message -Tags Notification
        $file1 = Export-Data -DataToExport $MEPFWithoutEPFME -DataToExportTitle 'MailEnabledPublicFolderMissingPublicFolderObject' -Depth 3 -DataType json -ReturnExportFilePath
        $file2 = Export-Data -DataToExport $MEPFWithoutEPFME -DataToExportTitle 'MailEnabledPublicFolderMissingPublicFolderObject' -DataType csv -ReturnExportFilePath
        Write-Information -Message "Exported Files: $file1,$file2" -Tags Notification
    }

}
