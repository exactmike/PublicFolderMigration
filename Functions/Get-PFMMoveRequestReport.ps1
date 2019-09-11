    Function Get-PFMMoveRequestReport
    {
        
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory)]
        [string]$BatchName
        ,
        [parameter(Mandatory)]
        [ValidateSet('Monitoring', 'FailureAnalysis', 'BatchCompletionMonitoring')]
        [string]$operation
        ,
        [datetime]$FailedSince
        ,
        [parameter()]
        [ValidateSet('All', 'Failed', 'InProgress', 'NotCompleted', 'LargeItemFailure', 'CommunicationFailure')]
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
                { $Script:fmr }
            }
            'Monitoring'
            {
                if ($passthru -and -not $PSBoundParameters.ContainsKey('StatsOperation'))
                { $Script:mr }
            }
            'BatchCompletionMonitoring'
            {
                if ($passthru -and -not $PSBoundParameters.ContainsKey('StatsOperation'))
                { $Script:cmr }
            }
        }
        switch ($statsoperation)
        {
            'All'
            {
                $logstring = "Getting request statistics for all $BatchName move requests."
                Write-Log -Message $logstring -EntryType Attempting
                $RecordCount = $Script:mr.count
                $b = 0
                $Script:mrs = @(
                    foreach ($request in $Script:mr)
                    {
                        $b++
                        Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b / $RecordCount * 100)
                        if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true)
                        {
                            Write-Log -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                            throw { "Connect to Exchange Organization $ExchangeOrganization Failed" }
                        }#End If
                        $splat = @{
                            Identity = $($request.requestguid)
                        }
                        Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization
                    }
                )
                $Script:ipmrs = @($Script:mrs | where-object { $psitem.status -like 'InProgress' })
                $Script:fmrs = @($Script:mrs | where-object { $psitem.status -like 'Failed' })
                $Script:asmrs = @($Script:mrs | where-object { $psitem.status -like 'Synced' })
                $Script:cmrs = @($Script:mrs | where-object { $psitem.status -like 'Completed*' })
                $script:ncmrs = @($script:mrs | Where-Object { $psitem.status -notlike 'Completed*' })
                if ($passthru)
                { $Script:mrs }
            }
            'Failed'
            {
                $logstring = "Getting Statistics for all failed $BatchName move requests."
                Write-Log -Message $logstring -EntryType Attempting
                $RecordCount = $Script:fmr.Count
                $b = 0
                $Script:fmrs = @(
                    foreach ($request in $fmr)
                    {
                        $b++
                        Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b / $RecordCount * 100)
                        if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true)
                        {
                            Write-Log -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                            throw { "Connect to Exchange Organization $ExchangeOrganization Failed" }
                        }#End If
                        $splat = @{
                            Identity = $($request.requestguid)
                        }
                        Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization
                    }
                )
                if ($FailedSince)
                {
                    $logstring = "Filtering Statistics for $BatchName move requests failed since $FailedSince."
                    Write-Log -Message $logstring -EntryType Attempting -Verbose
                    $script:fsfmrs = @($Script:fmrs | Where-Object { $_.FailureTimeStamp -gt $FailedSince })
                    if ($passthru)
                    { $Script:fsfmrs }
                }
                else
                {
                    if ($passthru)
                    { $Script:fmrs }
                }
            }
            'InProgress'
            {
                $logstring = "Getting Statistics for all in progress $BatchName move requests."
                Write-Log -Message $logstring -EntryType Attempting
                $RecordCount = $Script:ipmr.Count
                $b = 0
                $Script:ipmrs = @(
                    foreach ($request in $ipmr)
                    {
                        $b++
                        Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b / $RecordCount * 100)
                        if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true)
                        {
                            Write-Log -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                            throw { "Connect to Exchange Organization $ExchangeOrganization Failed" }
                        }#End If
                        $splat = @{
                            Identity = $($request.requestguid)
                        }
                        Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization
                    }
                )
                if ($passthru)
                { $Script:ipmrs }
            }
            'NotCompleted'
            {
                $logstring = "Getting move request statistics for not completed $BatchName move requests."
                Write-Log -Message $logstring -EntryType Attempting
                $RecordCount = $Script:ncmr.count
                $b = 0
                $Script:ncmrs = @(
                    foreach ($request in $Script:ncmr )
                    {
                        $b++
                        Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b / $RecordCount * 100)
                        if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true)
                        {
                            Write-Log -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                            throw { "Connect to Exchange Organization $ExchangeOrganization Failed" }
                        }#End If
                        $splat = @{
                            Identity = $($request.requestguid)
                        }
                        Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization
                    }
                )
                if ($passthru)
                { $Script:ncmrs }
            }
            'LargeItemFailure'
            {
                $logstring = "Getting Statistics for all failed $BatchName move requests."
                Write-Log -Message $logstring -EntryType Attempting -Verbose
                $RecordCount = $Script:fmr.count
                $b = 0
                $Script:fmrs = @(
                    foreach ($request in $fmr)
                    {
                        $b++
                        Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b / $RecordCount * 100)
                        if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true)
                        {
                            Write-Log -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                            throw { "Connect to Exchange Organization $ExchangeOrganization Failed" }
                        }#End If
                        $splat = @{
                            Identity = $($request.requestguid)
                        }
                        Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization
                    }
                )
                if ($failedsince)
                {
                    $logstring = "Filtering Statistics for $BatchName move requests failed since $FailedSince."
                    Write-Log -Message $logstring -EntryType Attempting -Verbose
                    $script:prelifmrs = @($Script:fmrs | Where-Object { $_.FailureTimeStamp -gt $FailedSince -and $_.FailureType -eq 'TooManyLargeItemsPermanentException' })
                }
                else
                {
                    $logstring = "Getting Statistics for all large item failed $BatchName move requests."
                    Write-Log -Message $logstring -EntryType Attempting -Verbose
                    $prelifmrs = @($Script:fmrs | Where-Object { $_.FailureType -eq 'TooManyLargeItemsPermanentException' })
                }
                $RecordCount = $prelifmrs.count
                $b = 0
                $Script:lifmrs = @(
                    foreach ($request in $prelifmrs)
                    {
                        $b++
                        Write-Progress -Activity "Getting move request statistics for all large item failed $BatchName move requests." -Status "Processing Record $b  of $RecordCount." -PercentComplete ($b / $RecordCount * 100)
                        if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true)
                        {
                            Write-Log -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
                            throw { "Connect to Exchange Organization $ExchangeOrganization Failed" }
                        }#End If
                        $splat = @{
                            Identity      = $($request.requestguid)
                            IncludeReport = $true
                        }
                        $request | ForEach-Object { Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -splat $splat -ExchangeOrganization $ExchangeOrganization } |
                        Select-Object -Property Alias, AllowLargeItems, ArchiveDomain, ArchiveGuid, BadItemLimit, BadItemsEncountered, BatchName, BytesTransferred, BytesTransferredPerMinute, CompleteAfter, CompletedRequestAgeLimit, CompletionTimestamp, DiagnosticInfo, Direction, DisplayName, DistinguishedName, DoNotPreserveMailboxSignature, ExchangeGuid, FailureCode, FailureSide, FailureTimestamp, FailureType, FinalSyncTimestamp, Flags, Identity, IgnoreRuleLimitErrors, InitialSeedingCompletedTimestamp, InternalFlags, IsOffline, IsValid, ItemsTransferred, LargeItemLimit, LargeItemsEncountered, LastUpdateTimestamp, MailboxIdentity, Message, MRSServerName, OverallDuration, PercentComplete, PositionInQueue, Priority, Protect, QueuedTimestamp, RecipientTypeDetails, RemoteArchiveDatabaseGuid, RemoteArchiveDatabaseName, RemoteCredentialUsername, RemoteDatabaseGuid, RemoteDatabaseName, RemoteGlobalCatalog, RemoteHostName, SourceArchiveDatabase, SourceArchiveServer, SourceArchiveVersion, SourceDatabase, SourceServer, SourceVersion, StartAfter, StartTimestamp, Status, StatusDetail, Suspend, SuspendedTimestamp, SuspendWhenReadyToComplete, SyncStage, TargetArchiveDatabase, TargetArchiveServer, TargetArchiveVersion, TargetDatabase, TargetDeliveryDomain, TargetServer, TargetVersion, TotalArchiveItemCount, TotalArchiveSize, TotalDataReplicationWaitDuration, TotalFailedDuration, TotalFinalizationDuration, TotalIdleDuration, TotalInProgressDuration, TotalMailboxItemCount, TotalMailboxSize, TotalProxyBackoffDuration, TotalQueuedDuration, TotalStalledDueToCIDuration, TotalStalledDueToHADuration, TotalStalledDueToMailboxLockedDuration, TotalStalledDueToReadCpu, TotalStalledDueToReadThrottle, TotalStalledDueToReadUnknown, TotalStalledDueToWriteCpu, TotalStalledDueToWriteThrottle, TotalStalledDueToWriteUnknown, TotalSuspendedDuration, TotalTransientFailureDuration, ValidationMessage, WorkloadType,
                        @{n = "BadItemList"; e = { @($_.Report.BadItems) } }, @{n = "LargeItemList"; e = { @($_.Report.LargeItems) } }
                    }
                )
                if ($passthru)
                { $Script:lifmrs }
            }
            'CommunicationFailure'
            {
                $logstring = "Getting Statistics for all communication error failed $BatchName move requests."
                Write-Log -Message $logstring -EntryType Attempting
                if ($FailedSince)
                {
                    $preCEfmrs = @($Script:fmrs | Where-Object { $_.FailureType -eq 'CommunicationErrorTransientException' -and $_.FailureTimeStamp -gt $FailedSince })
                }
                else
                {
                    $preCEfmrs = @($Script:fmrs | Where-Object { $_.FailureType -eq 'CommunicationErrorTransientException' })
                }
                $RecordCount = $preCEfmrs.count
                $b = 0
                $Script:cefmrs = @(
                    foreach ($request in $preCEfmrs)
                    {
                        $b++
                        Write-Progress -Activity $logstring -Status "Processing Record $b of $RecordCount." -PercentComplete ($b / $RecordCount * 100)
                        Connect-Exchange -ExchangeOrganization $ExchangeOrganization > $null
                        $request | ForEach-Object { Invoke-ExchangeCommand -cmdlet Get-PublicFolderMailboxMigrationRequestStatistics -string "-Identity $($_.alias) -IncludeReport" -ExchangeOrganization $ExchangeOrganization } | Select-Object -Property Alias, AllowLargeItems, ArchiveDomain, ArchiveGuid, BadItemLimit, BadItemsEncountered, BatchName, BytesTransferred, BytesTransferredPerMinute, CompleteAfter, CompletedRequestAgeLimit, CompletionTimestamp, DiagnosticInfo, Direction, DisplayName, DistinguishedName, DoNotPreserveMailboxSignature, ExchangeGuid, FailureCode, FailureSide, FailureTimestamp, FailureType, FinalSyncTimestamp, Flags, Identity, IgnoreRuleLimitErrors, InitialSeedingCompletedTimestamp, InternalFlags, IsOffline, IsValid, ItemsTransferred, LargeItemLimit, LargeItemsEncountered, LastUpdateTimestamp, MailboxIdentity, Message, MRSServerName, OverallDuration, PercentComplete, PositionInQueue, Priority, Protect, QueuedTimestamp, RecipientTypeDetails, RemoteArchiveDatabaseGuid, RemoteArchiveDatabaseName, RemoteCredentialUsername, RemoteDatabaseGuid, RemoteDatabaseName, RemoteGlobalCatalog, RemoteHostName, SourceArchiveDatabase, SourceArchiveServer, SourceArchiveVersion, SourceDatabase, SourceServer, SourceVersion, StartAfter, StartTimestamp, Status, StatusDetail, Suspend, SuspendedTimestamp, SuspendWhenReadyToComplete, SyncStage, TargetArchiveDatabase, TargetArchiveServer, TargetArchiveVersion, TargetDatabase, TargetDeliveryDomain, TargetServer, TargetVersion, TotalArchiveItemCount, TotalArchiveSize, TotalDataReplicationWaitDuration, TotalFailedDuration, TotalFinalizationDuration, TotalIdleDuration, TotalInProgressDuration, TotalMailboxItemCount, TotalMailboxSize, TotalProxyBackoffDuration, TotalQueuedDuration, TotalStalledDueToCIDuration, TotalStalledDueToHADuration, TotalStalledDueToMailboxLockedDuration, TotalStalledDueToReadCpu, TotalStalledDueToReadThrottle, TotalStalledDueToReadUnknown, TotalStalledDueToWriteCpu, TotalStalledDueToWriteThrottle, TotalStalledDueToWriteUnknown, TotalSuspendedDuration, TotalTransientFailureDuration, ValidationMessage, WorkloadType, @{n = "TotalTransientFailureMinutes"; e = { @($_.TotalTransientFailureDuration.TotalMinutes) } }, @{n = "TotalStalledDueToMailboxLockedMinutes"; e = { @($_.TotalStalledDueToMailboxLockedDuration.TotalMinutes) } }
                    }
                )
                if ($passthru)
                { $Script:cefmrs }
            }
        }
    }

    }

