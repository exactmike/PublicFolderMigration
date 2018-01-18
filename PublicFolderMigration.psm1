###############################################################################################
#Core Public Folder Migration Module Functions
###############################################################################################

function Find-OrphanedMailEnabledPublicFolders
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
            cmdlet = 'Get-MailPublicFolder'
            ErrorAction = 'Stop'
            Splat = @{
                Identity = ''
                ErrorAction = 'SilentlyContinue'
                WarningAction = 'SilentlyContinue'
            }
            ExchangeOrganization = $ExchangeOrganization
        }
        $message = "Get-MailPublicFolder for each Public Folder"
        WriteLog -Message $message -EntryType Attempting -Verbose
        $ExchangePublicFoldersMailEnabled = @(
            $PFCount = 0
            foreach ($pf in $ExchangePublicFolders) {
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
                    Write-Output -InputObject $MEPF
                }
            }
        )
        $EPFMEWithoutMEPF = @(
            foreach ($EPFME in $ExchangePublicFoldersMailEnabled)
            {
                if (-not $MEPFHashByDN.ContainsKey($EPFME.DistinguishedName))
                {
                    Write-Output -InputObject $EPFME
                }
            }
        )
        if ($EPFMEWithoutMEPF.Count -ge 1) {
            $message = "Found Public Folders which are mail enabled but for which no mail enabled public folder object was found with get-mailpublicfolder.  Exporting Data."
            WriteLog -message $message -Verbose
            $file1 = Export-Data -DataToExport $EPFMEWithoutMEPF -DataToExportTitle 'PublicFoldersMissingMailEnabledObject' -Depth 3 -DataType json -ReturnExportFilePath
            $file2 = Export-Data -DataToExport $EPFMEWithoutMEPF -DataToExportTitle 'PublicFoldersMissingMailEnabledObject' -DataType csv -ReturnExportFilePath
            WriteLog -Message "Exported Files: $file1,$file2" -Verbose
        }
        if ($MEPFWithoutEPFME.Count -ge 1) {
            $message = "Found Mail Enabled Public Folders for which no public folder object was found.  Exporting Data."
            WriteLog -message $message -Verbose
            $file1 = Export-Data -DataToExport $MEPFWithoutEPFME -DataToExportTitle 'MailEnabledPublicFolderMissingPublicFolderObject' -Depth 3 -DataType json -ReturnExportFilePath
            $file2 = Export-Data -DataToExport $MEPFWithoutEPFME -DataToExportTitle 'MailEnabledPublicFolderMissingPublicFolderObject' -DataType csv -ReturnExportFilePath
            WriteLog -Message "Exported Files: $file1,$file2" -Verbose
        }
    }
#end function Find-OrphanedMailEnabledPublicFolders
function Export-UserPublicFolderTree
    {
        [cmdletbinding()]
        param
        (
            $ExchangeOrganization
            ,
            [switch]$ExportPermissions
        )
        $allUserPublicFolders = @(Get-UserPublicFolderTree -ExchangeOrganization $ExchangeOrganization)
        $ExportFile = Export-Data -DataToExport $allUserPublicFolders -DataToExportTitle UserPublicFolderTree -Depth 3 -DataType json -ReturnExportFilePath -ErrorAction Stop
        WriteLog -Message "Exported UserPublicFolderTree File: $ExportFile" -Verbose
        if ($ExportPermissions)
        {
            $splat = @{
                ExchangeOrganization = $ExchangeOrganization
                Cmdlet = 'Get-PublicFolderClientPermission'
                ErrorAction = 'Stop'
                Splat = @{
                    Identity = ''
                    ErrorAction = 'Stop'
                }
            }
            $PublicFolderUserPermissions = @(
            $allUserPublicFolders | ForEach-Object
                {
                    $splat.splat.Identity = $_.EntryID
                    Invoke-ExchangeCommand @splat | Select-Object -Property Identity,User -ExpandProperty AccessRights
                }
            )
            $ExportFile = Export-Data -DataToExport $PublicFolderUserPermissions -DataToExportTitle UserPublicFolderPermissions -Depth 1 -DataType csv -ReturnExportFilePath
            WriteLog -Message "Exported UserPublicFolderPermissions File: $ExportFile" -Verbose
        }
    }
#end function Export-UserPublicFolderTree
function Export-MailPublicFolder
    {
        [cmdletbinding()]
        param
        (
            $ExchangeOrganization
        )
        $allMailPublicFolders = @(Get-AllMailPublicFolder -ExchangeOrganization $ExchangeOrganization)
        $ExportFile = Export-Data -DataToExport $allMailPublicFolders -DataToExportTitle MailPublicFolders -Depth 3 -DataType json -ReturnExportFilePath -ErrorAction Stop
        WriteLog -Message "Exported MailPublicFolders File: $ExportFile" -Verbose
    }
#end function Export-MailPublicFolder
function Get-UserPublicFolderTree
    {
        [cmdletbinding()]
        param
        (
            [parameter(Mandatory)]
            $ExchangeOrganization
        )
        #Get All Public Folders
        $splat = @{
            cmdlet = 'Get-PublicFolder'
            ErrorAction = 'Stop'
            Splat = @{
                Recurse = $true
                Identity = '\'
            }
            ExchangeOrganization = $ExchangeOrganization
        }
        try 
        {
            $message = "Get All Mail Enabled Public Folder Objects"
            WriteLog -Message $message -EntryType Attempting -Verbose
            $ExchangePublicFolders = @(Invoke-ExchangeCommand @splat)
            WriteLog -Message $message -EntryType Succeeded -Verbose
            Write-Output -InputObject $ExchangePublicFolders
        }
        catch
        {
            $myerror = $_
            WriteLog -Message $message -EntryType Failed -Verbose -ErrorLog
            WriteLog -Message $myerror.tostring() -ErrorLog
            $myerror
        }
    }
#end function Get-UserPublicFolderTree
function Get-AllMailPublicFolder
    {
        [cmdletbinding()]
        param
        (
            [parameter(Mandatory)]
            $ExchangeOrganization
        )
        #Get all mail enabled public folders
        $splat = @{
            cmdlet = 'Get-MailPublicFolder'
            ErrorAction = 'Stop'
            splat = @{
                ResultSize = 'Unlimited'
                ErrorAction = 'stop'
                WarningAction = 'SilentlyContinue'
            }
            ExchangeOrganization = $ExchangeOrganization
        }
        try 
        {
            $message = "Get All Mail Enabled Public Folder Objects"
            WriteLog -Message $message -EntryType Attempting -Verbose
            $MailEnabledPublicFolders = @(Invoke-ExchangeCommand @splat)
            WriteLog -Message $message -EntryType Succeeded -Verbose
            Write-Output -InputObject $MailEnabledPublicFolders
        }
        catch
        {
            $myerror = $_
            WriteLog -Message $message -EntryType Failed -Verbose -ErrorLog
            WriteLog -Message $myerror.tostring() -ErrorLog
            $myerror
        }
    }
#end function Get-AllMailPublicFolder
function Get-PFMMoveRequest
    {
        [cmdletbinding()]
        param(
            $ExchangeOrganization
            ,
            $BatchName
        )
        if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true)
        {
            WriteLog -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
            throw {"Connect to Exchange Organization $ExchangeOrganization Failed"}
        }#End If
        $Message = "Get all existing $BatchName move requests"
        WriteLog -message $Message -Verbose -EntryType Attempting
        $splat = @{
        cmdlet = 'Get-PublicFolderMailboxMigrationRequest'
        ExchangeOrganization = $ExchangeOrganization
        ErrorAction = 'Stop'
        splat = @{
            BatchName = 'MigrationService:' + $BatchName
            ResultSize = 'Unlimited'
            ErrorAction = 'Stop'
        }#innersplat
        }#outersplat
        $Script:mr = @(Invoke-ExchangeCommand @splat)
        $Script:fmr = @($mr | Where-Object -FilterScript {$_.status -eq 'Failed'})
        $Script:ipmr = @($mr | Where-Object {$_.status -eq 'InProgress'})
        $Script:smr = @($mr | Where-Object {$_.status -eq 'Suspended'})
        $Script:asmr = @($mr | Where-Object {$_.status -in ('AutoSuspended','Synced')})
        $Script:cmr = @($mr | Where-Object {$_.status -like 'Completed*'})
        $Script:qmr = @($mr | Where-Object {$_.status -eq 'Queued'})
        $Script:ncmr = @($mr | Where-Object {$_.status -notlike 'Completed*'})
        WriteLog -message $Message -Verbose -EntryType Succeeded
    }
#end function Get-PFMMoveRequest
function Get-PFMMoveRequestReport
    {
        [cmdletbinding()]
        param
        (
            [parameter(Mandatory)]
            [string]$BatchName
            ,
            [parameter(Mandatory)]
            [ValidateSet('Monitoring','FailureAnalysis','BatchCompletionMonitoring')]
            [string]$operation
            ,
            [datetime]$FailedSince
            ,
            [parameter()]
            [ValidateSet('All','Failed','InProgress','NotCompleted','LargeItemFailure','CommunicationFailure')]
            [string]$StatsOperation
            ,
            [switch]$passthru
            ,
            [Parameter(Mandatory)]
            [string]$ExchangeOrganization #convert to dynamic parameter later
        )
        Process
        {
            Get-PFMMoveRequest -ExchangeOrganization $ExchangeOrganization -BatchName $BatchName
            switch ($operation)
            {
                'FailureAnalysis'
                {
                    if ($passthru -and -not $PSBoundParameters.ContainsKey('StatsOperation'))
                    {$Script:fmr}
                }
                'Monitoring'
                {
                    if ($passthru -and -not $PSBoundParameters.ContainsKey('StatsOperation'))
                    {$Script:mr}
                }
                'BatchCompletionMonitoring'
                {
                    if ($passthru -and -not $PSBoundParameters.ContainsKey('StatsOperation'))
                    {$Script:cmr}
                }
            }
            switch ($statsoperation)
            {
                'All'
                {
                    $logstring = "Getting request statistics for all $BatchName move requests." 
                    WriteLog -Message $logstring -EntryType Attempting 
                    $RecordCount=$Script:mr.count
                    $b=0
                    $Script:mrs = @(
                        foreach ($request in $Script:mr)
                        {
                            $b++
                            Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b/$RecordCount*100)
                            if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true) {
                                WriteLog -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                                throw {"Connect to Exchange Organization $ExchangeOrganization Failed"}
                            }#End If
                            $splat = @{
                            Identity = $($request.requestguid)
                            }
                            Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization
                        }
                    )
                    $Script:ipmrs = @($Script:mrs | where-object {$psitem.status -like 'InProgress'})
                    $Script:fmrs = @($Script:mrs | where-object {$psitem.status -like 'Failed'})
                    $Script:asmrs = @($Script:mrs | where-object {$psitem.status -like 'Synced'})
                    $Script:cmrs = @($Script:mrs |  where-object {$psitem.status -like 'Completed*'})
                    $script:ncmrs = @($script:mrs | Where-Object {$psitem.status -notlike 'Completed*'})
                    if ($passthru)
                    {$Script:mrs}
                }
                'Failed'
                {
                    $logstring = "Getting Statistics for all failed $BatchName move requests."
                    WriteLog -Message $logstring -EntryType Attempting
                    $RecordCount=$Script:fmr.Count
                    $b=0
                    $Script:fmrs = @(
                        foreach ($request in $fmr)
                        {
                            $b++
                            Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b/$RecordCount*100)
                            if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true) {
                                WriteLog -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                                throw {"Connect to Exchange Organization $ExchangeOrganization Failed"}
                            }#End If
                            $splat = @{
                            Identity = $($request.requestguid)
                            }
                            Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization
                        }
                    )
                    if ($FailedSince)
                    {
                        $logstring =  "Filtering Statistics for $BatchName move requests failed since $FailedSince."
                        WriteLog -Message $logstring -EntryType Attempting -Verbose
                        $script:fsfmrs = @($Script:fmrs | Where-Object {$_.FailureTimeStamp -gt $FailedSince})
                        if ($passthru)
                        {$Script:fsfmrs}
                    }
                    else
                    {
                        if ($passthru)
                        {$Script:fmrs}
                    }
                }
                'InProgress'
                {
                    $logstring = "Getting Statistics for all in progress $BatchName move requests."
                    WriteLog -Message $logstring -EntryType Attempting
                    $RecordCount=$Script:ipmr.Count
                    $b=0
                    $Script:ipmrs = @(
                        foreach ($request in $ipmr)
                        {
                            $b++
                            Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b/$RecordCount*100)
                            if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true) {
                                WriteLog -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                                throw {"Connect to Exchange Organization $ExchangeOrganization Failed"}
                            }#End If
                            $splat = @{
                            Identity = $($request.requestguid)
                            }
                            Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization
                        }
                    )
                    if ($passthru)
                    {$Script:ipmrs}
                }
                'NotCompleted'
                {
                    $logstring = "Getting move request statistics for not completed $BatchName move requests." 
                    WriteLog -Message $logstring -EntryType Attempting
                    $RecordCount=$Script:ncmr.count
                    $b=0
                    $Script:ncmrs = @(
                        foreach ($request in $Script:ncmr )
                        {
                            $b++
                            Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b/$RecordCount*100)
                            if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true) {
                                WriteLog -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                                throw {"Connect to Exchange Organization $ExchangeOrganization Failed"}
                            }#End If
                            $splat = @{
                            Identity = $($request.requestguid)
                            }
                            Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization
                        }
                    )
                    if ($passthru)
                    {$Script:ncmrs}
                }
                'LargeItemFailure'
                {
                    $logstring = "Getting Statistics for all failed $BatchName move requests." 
                    WriteLog -Message $logstring -EntryType Attempting -Verbose
                    $RecordCount=$Script:fmr.count
                    $b=0
                    $Script:fmrs = @(
                    foreach ($request in $fmr)
                        {
                            $b++
                            Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b/$RecordCount*100)
                            if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true) {
                                WriteLog -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                                throw {"Connect to Exchange Organization $ExchangeOrganization Failed"}
                            }#End If
                            $splat = @{
                            Identity = $($request.requestguid)
                            }
                            Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization
                        }
                    )
                    if ($failedsince)
                    {
                        $logstring =  "Filtering Statistics for $BatchName move requests failed since $FailedSince."
                        WriteLog -Message $logstring -EntryType Attempting -Verbose
                        $script:prelifmrs = @($Script:fmrs | Where-Object {$_.FailureTimeStamp -gt $FailedSince -and $_.FailureType -eq 'TooManyLargeItemsPermanentException'})
                    }
                    else
                    {
                        $logstring =  "Getting Statistics for all large item failed $BatchName move requests."
                        WriteLog -Message $logstring -EntryType Attempting -Verbose
                        $prelifmrs = @($Script:fmrs | Where-Object {$_.FailureType -eq 'TooManyLargeItemsPermanentException'})
                    }
                    $RecordCount=$prelifmrs.count
                    $b=0
                    $Script:lifmrs = @(
                        foreach ($request in $prelifmrs)
                        {
                            $b++
                            Write-Progress -Activity "Getting move request statistics for all large item failed $BatchName move requests." -Status "Processing Record $b  of $RecordCount." -PercentComplete ($b/$RecordCount*100)
                            if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true) {
                                WriteLog -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                                throw {"Connect to Exchange Organization $ExchangeOrganization Failed"}
                            }#End If
                            $splat = @{
                            Identity = $($request.requestguid)
                            IncludeReport = $true
                            }
                            $request | ForEach-Object {Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization} | 
                            Select-Object -Property Alias,AllowLargeItems,ArchiveDomain,ArchiveGuid,BadItemLimit,BadItemsEncountered,BatchName,BytesTransferred,BytesTransferredPerMinute,CompleteAfter,CompletedRequestAgeLimit,CompletionTimestamp,DiagnosticInfo,Direction,DisplayName,DistinguishedName,DoNotPreserveMailboxSignature,ExchangeGuid,FailureCode,FailureSide,FailureTimestamp,FailureType,FinalSyncTimestamp,Flags,Identity,IgnoreRuleLimitErrors,InitialSeedingCompletedTimestamp,InternalFlags,IsOffline,IsValid,ItemsTransferred,LargeItemLimit,LargeItemsEncountered,LastUpdateTimestamp,MailboxIdentity,Message,MRSServerName,OverallDuration,PercentComplete,PositionInQueue,Priority,Protect,QueuedTimestamp,RecipientTypeDetails,RemoteArchiveDatabaseGuid,RemoteArchiveDatabaseName,RemoteCredentialUsername,RemoteDatabaseGuid,RemoteDatabaseName,RemoteGlobalCatalog,RemoteHostName,SourceArchiveDatabase,SourceArchiveServer,SourceArchiveVersion,SourceDatabase,SourceServer,SourceVersion,StartAfter,StartTimestamp,Status,StatusDetail,Suspend,SuspendedTimestamp,SuspendWhenReadyToComplete,SyncStage,TargetArchiveDatabase,TargetArchiveServer,TargetArchiveVersion,TargetDatabase,TargetDeliveryDomain,TargetServer,TargetVersion,TotalArchiveItemCount,TotalArchiveSize,TotalDataReplicationWaitDuration,TotalFailedDuration,TotalFinalizationDuration,TotalIdleDuration,TotalInProgressDuration,TotalMailboxItemCount,TotalMailboxSize,TotalProxyBackoffDuration,TotalQueuedDuration,TotalStalledDueToCIDuration,TotalStalledDueToHADuration,TotalStalledDueToMailboxLockedDuration,TotalStalledDueToReadCpu,TotalStalledDueToReadThrottle,TotalStalledDueToReadUnknown,TotalStalledDueToWriteCpu,TotalStalledDueToWriteThrottle,TotalStalledDueToWriteUnknown,TotalSuspendedDuration,TotalTransientFailureDuration,ValidationMessage,WorkloadType,
                            @{n="BadItemList";e={@($_.Report.BadItems)}},@{n="LargeItemList";e={@($_.Report.LargeItems)}}
                        }
                    )
                    if ($passthru)
                    {$Script:lifmrs}
                }
                'CommunicationFailure'
                {
                    $logstring = "Getting Statistics for all communication error failed $BatchName move requests."
                    WriteLog -Message $logstring -EntryType Attempting
                    if ($FailedSince)
                    {
                        $preCEfmrs = @($Script:fmrs | Where-Object {$_.FailureType -eq 'CommunicationErrorTransientException' -and $_.FailureTimeStamp -gt $FailedSince})
                    }
                    else
                    {
                        $preCEfmrs = @($Script:fmrs | Where-Object {$_.FailureType -eq 'CommunicationErrorTransientException'})
                    }
                    $RecordCount=$preCEfmrs.count
                    $b=0
                    $Script:cefmrs = @(
                        foreach ($request in $preCEfmrs)
                        {
                            $b++
                            Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b/$RecordCount*100)
                            Connect-Exchange -ExchangeOrganization $ExchangeOrganization > $null
                            $request | ForEach-Object {Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -string "-Identity $($_.alias) -IncludeReport" -ExchangeOrganization $ExchangeOrganization} | Select-Object -Property Alias,AllowLargeItems,ArchiveDomain,ArchiveGuid,BadItemLimit,BadItemsEncountered,BatchName,BytesTransferred,BytesTransferredPerMinute,CompleteAfter,CompletedRequestAgeLimit,CompletionTimestamp,DiagnosticInfo,Direction,DisplayName,DistinguishedName,DoNotPreserveMailboxSignature,ExchangeGuid,FailureCode,FailureSide,FailureTimestamp,FailureType,FinalSyncTimestamp,Flags,Identity,IgnoreRuleLimitErrors,InitialSeedingCompletedTimestamp,InternalFlags,IsOffline,IsValid,ItemsTransferred,LargeItemLimit,LargeItemsEncountered,LastUpdateTimestamp,MailboxIdentity,Message,MRSServerName,OverallDuration,PercentComplete,PositionInQueue,Priority,Protect,QueuedTimestamp,RecipientTypeDetails,RemoteArchiveDatabaseGuid,RemoteArchiveDatabaseName,RemoteCredentialUsername,RemoteDatabaseGuid,RemoteDatabaseName,RemoteGlobalCatalog,RemoteHostName,SourceArchiveDatabase,SourceArchiveServer,SourceArchiveVersion,SourceDatabase,SourceServer,SourceVersion,StartAfter,StartTimestamp,Status,StatusDetail,Suspend,SuspendedTimestamp,SuspendWhenReadyToComplete,SyncStage,TargetArchiveDatabase,TargetArchiveServer,TargetArchiveVersion,TargetDatabase,TargetDeliveryDomain,TargetServer,TargetVersion,TotalArchiveItemCount,TotalArchiveSize,TotalDataReplicationWaitDuration,TotalFailedDuration,TotalFinalizationDuration,TotalIdleDuration,TotalInProgressDuration,TotalMailboxItemCount,TotalMailboxSize,TotalProxyBackoffDuration,TotalQueuedDuration,TotalStalledDueToCIDuration,TotalStalledDueToHADuration,TotalStalledDueToMailboxLockedDuration,TotalStalledDueToReadCpu,TotalStalledDueToReadThrottle,TotalStalledDueToReadUnknown,TotalStalledDueToWriteCpu,TotalStalledDueToWriteThrottle,TotalStalledDueToWriteUnknown,TotalSuspendedDuration,TotalTransientFailureDuration,ValidationMessage,WorkloadType,@{n="TotalTransientFailureMinutes";e={@($_.TotalTransientFailureDuration.TotalMinutes)}},@{n="TotalStalledDueToMailboxLockedMinutes";e={@($_.TotalStalledDueToMailboxLockedDuration.TotalMinutes)}}
                    }
                )
                if ($passthru)
                {$Script:cefmrs}
                }
            }
        }#end Process
    }
#end function Get-PFMMoveRequestReport