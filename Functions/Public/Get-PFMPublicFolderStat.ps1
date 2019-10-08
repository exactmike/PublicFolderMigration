#Add publicfolderobject with replicas attribute processing, public folder path processing, and 'all' processing.
function Get-PFMPublicFolderStat
{
    <#
    .SYNOPSIS
    Gets the statistics for one or more, or all, Exchange 2010 Public Folders.
    .DESCRIPTION
    Gets the statistics for one or more, or all, Exchange 2010 Public Folders. Default is for All, with the PublicFolderPath parameter you can specify multiple public folders, with the PublicFolderInfoObject parameter you can specify multiple public folders and their replicas.
    .PARAMETER PublicFolderMailboxServer
    This parameter specifies the Exchange 2010 server(s) from which to retrieve statistics. If this is omitted, all Exchange servers hosting a Public Folder Database are targeted. Alternatively, use the Public Folder Info Object parameter.
    .PARAMETER PublicFolderPath
    Accepts one or more public folder paths and retrieves the statistics for them from each public folder database server specified, or from all public folder database servers if none were specified.
    .PARAMETER PublicFolderInfoObject
    Accepts one or more of the output objects from Get-PFMPublicFolderTree. Results in statistics being returned for each replica of the public folder info object submitted, but not from public folder servers which do not have a replica.
    .PARAMETER Passthru
    Controls whether the public folder statistics objects are returned to the PowerShell pipeline for further processing.
    .PARAMETER OutputFolderPath
    Mandatory parameter for the already existing directory location where you want public folder stats output to be placed.  Operational log files will also go to this location.
    .PARAMETER OutputFormat
    Mandatory parameter used to specify whether you want csv, json, xml, clixml or any combination of these.
    .PARAMETER SendEmail
    This switch will set the script to send an email report.  To use this parameter you must have already used the Set-PFMEmailConfiguration cmdlet to configure your email settings.
    .EXAMPLE
    PS C:\> Connect-PFMExchange -ExchangeOnPremisesServer PublicFolderServer1.us.wa.contoso.com -credential $cred
    PS C:\> Get-PFMPublicFolderReplicationReport -OutputFolderPath c:\PFReports -OutputFormats csv,html
    Gets public folder tree data from PublicFolderServer1.us.wa.contoso.com and public folder stats data from all other public folder database servers in the Exchange Organization
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSAvoidInvokingEmptyMembers",
        "",
        Justification = "Intentionally Uses Non-Constant Members for Stats Processing from multiple servers"
    )]
    [CmdletBinding(ConfirmImpact = 'None', DefaultParameterSetName = 'All')]
    [OutputType([System.Object[]])]
    param
    (
        [parameter(ParameterSetName = 'Path')]
        [parameter(ParameterSetName = 'All')]
        [string[]]$PublicFolderMailboxServer = @()
        ,
        [parameter(ParameterSetName = 'Path', Mandatory)]
        [string[]]$PublicFolderPath = @()
        ,
        [parameter(ParameterSetName = 'InfoObject', Mandatory)]
        [psobject[]]$PublicFolderInfoObject
        ,
        [parameter()]
        [switch]$Passthru
        ,
        [parameter(Mandatory)]
        [ValidateScript( { TestIsWriteableDirectory -path $_ })]
        [string]$OutputFolderPath
        ,
        [parameter(Mandatory)]
        [ExportDataOutputFormat[]]$Outputformat
        ,
        [parameter()]
        #Add ValidateScript to verify Email Configuration is set
        [ValidateScript( {
                if ($null -eq $script:EmailConfiguration)
                {
                    Write-Warning -message 'You must run Set-PFMEmailConfiguration before use the SendEmail parameter'
                    $false
                }
                else
                {
                    $true
                } })]
        [switch]$SendEmail
        ,
        [parameter()]
        [ValidateSet('Unicode', 'BigEndianUnicode', 'Ascii', 'Default', 'UTF8', 'UTF8NOBOM', 'UTF7', 'UTF32')]
        [string]$Encoding = 'UTF8'
    )
    Confirm-PFMExchangeConnection -PSSession $Script:PSSession
    $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
    $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetPublicFolderStat.log')
    $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetPublicFolderStat-ERRORS.log')
    WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
    $ExchangeOrganization = Invoke-Command -Session $Script:PSSession -ScriptBlock { Get-OrganizationConfig | Select-Object -ExpandProperty Identity | Select-Object -ExpandProperty Name }
    WriteLog -Message "Exchange Session is Running in Exchange Organzation $ExchangeOrganization" -EntryType Notification
    #region ValidateParameters
    #if the user specified public folder mailbox servers, validate them:
    function Get-PublicFolderMailboxServerDatabase
    {
        [cmdletbinding()]
        param(
            $PublicFolderMailboxServer
        )
        switch ($PublicFolderMailboxServer.Count)
        {
            0
            {
                $ServerDatabases = @(
                    Invoke-Command -Session $script:PSSession -ScriptBlock {
                        Get-PublicFolderDatabase
                    } | Select-Object -Property @{n = 'DatabaseName'; e = { $_.Name } }, @{n = 'ServerName'; e = { $_.Server } }, @{n = 'ServerFQDN'; e = { $_.RpcClientAccessServer } }
                )

            }
            { $_ -gt 0 }
            {
                $ServerDatabases = @(
                    foreach ($Server in $PublicFolderMailboxServer)
                    {

                        Invoke-Command -Session $script:PSSession -scriptblock {
                            Get-PublicFolderDatabase -server $using:Server -ErrorAction SilentlyContinue
                        } | Select-Object -Property @{n = 'DatabaseName'; e = { $_.Name } }, @{n = 'ServerName'; e = { $_.Server } }, @{n = 'ServerFQDN'; e = { $_.RpcClientAccessServer } }
                    }
                )
                if ($ServerDatabases.Count -ne $PublicFolderMailboxServer.Count)
                {
                    Write-Error "One or more of the specified PublicFolderMailboxServers $($PublicFolderMailboxServer -join ', ') does not host a public folder database."
                    Return $null
                }
            }
        }
        $ServerDatabases
    }
    $ServerDatabase = @(Get-PublicFolderMailboxServerDatabase -PublicFolderMailboxServer $PublicFolderMailboxServer)
    $PublicFolderMailboxServerNames = $ServerDatabase.ServerName -join ', '
    WriteLog -Message "Public Folder Mailbox Servers Included: $PublicFolderMailboxServerNames" -EntryType Notification -Verbose
    #region GetPublicFolderStats
    #Make Server PSSessions
    $connectSessionFailure = [System.Collections.Generic.List[String]]::new()
    $connectSessionSuccess = [System.Collections.Generic.List[String]]::new()
    foreach ($s in $ServerDatabase)
    {
        $ConnectPFMExchangeParams = @{
            ExchangeOnPremisesServer = $s.ServerFQDN
            IsParallel               = $true
            ErrorAction              = 'Stop'
        }
        if ($null -ne $Script:PSSessionOption)
        {
            $ConnectPFMExchangeParams.PSSessionOption = $Script:PSSessionOption
        }
        try
        {
            Connect-PFMExchange @ConnectPFMExchangeParams
            writelog -message "Connected Parallel PSSession to $($s.ServerFQDN) for Stats operations" -entrytype Notification -verbose
            $connectSessionSuccess.Add($s.ServerFQDN)
        }
        catch
        {
            Writelog -message "Unable to connect a remote Exchange Powershell session to $($s.ServerFQDN)" -entryType Failed -Verbose
            $connectSessionFailure.Add($s.ServerFQDN)
        }
    }
    $ServerDatabaseToProcess, $ServerDatabaseRetry = $ServerDatabase.where( { $_.ServerFQDN -in $connectSessionSuccess }, 'Split')
    if ($connectSessionFailure.Count -ge 1)
    {
        writelog -message "Connect Session Failures: $($connectSessionFailure -join ',')" -entrytype Notification
        if ($PSCmdlet.ParameterSetName -in @('InfoObject', 'Path'))
        {
            throw('Not all required or specified public folder servers were connected to for stats operations. Quitting to avoid incomplete data return')
            Return $null
        }
    }
    if ($connectSessionSuccess.count -eq 0)
    {
        throw('None of the specified public folder servers were connected to for stats operations. Quitting to avoid incomplete data return')
        Return $null
    }
    #Get Stats from successful connections
    $publicFolderStats =
    @(
        # if the user specified public folder path then only retrieve stats for the specified folders.
        # This can be significantly faster than retrieving stats for all public folders
        switch ($PSCmdlet.ParameterSetName)
        {
            'All'
            {
                #Start the jobs
                $StatsJob = @(
                    foreach ($s in $ServerDatabaseToProcess)
                    {
                        $ServerSession = Get-PFMParallelPSSession -name $s.ServerFQDN
                        Confirm-PFMExchangeConnection -IsParallel -PSSession $ServerSession
                        $ServerSession = Get-PFMParallelPSSession -name $s.ServerFQDN
                        Write-Verbose "Starting Job to retrieve stats for all Public Folders from $($s.ServerFQDN)"
                        $ServerName = $s.ServerName
                        #avoid NULL output by testing for results while still suppressing errors with SilentlyContinue
                        Invoke-Command -Session $ServerSession -ScriptBlock {
                            Get-PublicFolderStatistics -Server $using:ServerName -ResultSize Unlimited -ErrorAction SilentlyContinue
                        } -AsJob -JobName $s.ServerFQDN
                    }
                )
                #Monitor the jobs
                $CompletedJobCount = 0
                $StatsJobCount = $StatsJob.Count
                $StatsJobStopWatch = [System.Diagnostics.Stopwatch]::new()
                $StatsJobStopWatch.Start()
                do
                {
                    $States = $StatsJob | Measure-Property -property State -ashashtable
                    $CompletedJobCount = $states.Completed
                    $ElapsedTimeString = "{0} Days, {1} Hours, {2} Minutes, {3} Seconds" -f $StatsJobStopWatch.Elapsed.Days, $StatsJobStopWatch.Elapsed.Hours, $StatsJobStopWatch.Elapsed.Minutes, $StatsJobStopWatch.Elapsed.Seconds
                    $WriteProgressParams = @{
                        Activity         = 'Monitoring Public Folder Statistics Retrieval Jobs'
                        CurrentOperation = "Monitoring $StatsJobCount Jobs"
                        PercentComplete  = $($CompletedJobCount / $StatsJobCount * 100)
                        Status           = "$CompletedJobCount of $StatsJobCount Completed. Elapsed time: $ElapsedTimeString"
                    }
                    Write-Progress @WriteProgressParams
                    Start-Sleep -Seconds 10
                }
                until
                (
                    $CompletedJobCount -eq $StatsJobsCount -or ($null -ne $States.Failed -and $states.Failed.Count -ge 1)
                )
                $StatsJobStopWatch.Stop()

                #Retrieve the Jobs
                $States = $StatsJob | Measure-Property -property State -ashashtable
                $CompletedJobCount = $states.Completed
                $ElapsedTimeString = "{0} Days, {1} Hours, {2} Minutes, {3} Seconds" -f $StatsJobStopWatch.Elapsed.Days, $StatsJobStopWatch.Elapsed.Hours, $StatsJobStopWatch.Elapsed.Minutes, $StatsJobStopWatch.Elapsed.Seconds
                $WriteProgressParams = @{
                    Activity         = 'Monitoring Public Folder Statistics Retrieval Jobs'
                    CurrentOperation = "Monitored $StatsJobCount Jobs"
                    PercentComplete  = 100
                    Status           = "$CompletedJobCount of $StatsJobCount Jobs Completed. Elapsed time: $ElapsedTimeString"
                    Completed        = $true
                }
                Write-Progress @WriteProgressParams
                switch ($CompletedJobCount -eq $StatsJobsCount)
                {
                    $true
                    {
                        $StatsJobStopWatch.Reset()
                        $StatsJobStopWatch.Start()
                        $ReceivedJobCount = 0

                        Foreach ($job in $StatsJob)
                        {
                            $ElapsedTimeString = "{0} Days, {1} Hours, {2} Minutes, {3} Seconds" -f $StatsJobStopWatch.Elapsed.Days, $StatsJobStopWatch.Elapsed.Hours, $StatsJobStopWatch.Elapsed.Minutes, $StatsJobStopWatch.Elapsed.Seconds
                            $WriteProgressParams = @{
                                Activity         = 'Receiving Public Folder Statistics From Jobs'
                                CurrentOperation = "Receiving from $($job.Name)."
                                PercentComplete  = $($ReceivedJobCount / $StatsJobCount * 100)
                                Status           = "$ReceivedJobCount of $StatsJobCount Jobs Received. Elapsed time: $ElapsedTimeString"
                            }
                            Write-Progress @WriteProgressParams
                            $customProperties = @(
                                '*'
                                @{n = 'ServerName'; e = { $job.Name } }
                                @{n = 'SizeInBytes'; e = { $_.TotalItemSize.ToString().split(('(', ')'))[1].replace(',', '').replace(' bytes', '') -as [long] } }
                            )
                            $theStats = Receive-Job -job $job -ErrorAction Stop
                            if ($null -ne $thestats)
                            {
                                $thestats | Select-Object -ExcludeProperty ServerName -Property $customProperties
                            }
                            Remove-Job -Job $job
                            $ReceivedJobCount++
                        }
                        $ElapsedTimeString = "{0} Days, {1} Hours, {2} Minutes, {3} Seconds" -f $StatsJobStopWatch.Elapsed.Days, $StatsJobStopWatch.Elapsed.Hours, $StatsJobStopWatch.Elapsed.Minutes, $StatsJobStopWatch.Elapsed.Seconds
                        $WriteProgressParams = @{
                            Activity         = 'Receiving Public Folder Statistics From Jobs'
                            CurrentOperation = "Completed"
                            PercentComplete  = 100
                            Status           = "$ReceivedJobCount of $StatsJobCount Jobs Received. Elapsed time: $ElapsedTimeString"
                            Completed        = $true
                        }
                        Write-Progress @WriteProgressParams
                    }
                    $false
                    {
                        throw("some jobs failed to complete")
                    }
                }
            }
            'InfoObject'
            {
                $count = 0
                $recordCount = 0
                ($PublicFolderInfoObject).foreach( { $recordCount += $_.ReplicaCount })
                foreach ($i in $PublicFolderInfoObject)
                {
                    foreach ($r in $i.Replicas)
                    {
                        $s = ($ServerDatabase).where( { $_.DatabaseName -eq $r })
                        $customProperties =
                        @(
                            '*'
                            @{n = 'ServerName'; e = { $s.ServerFQDN } }
                            #this is necessary b/c powershell remoting makes the attributes deserialized and the value in bytes is not available directly.  Code below should work in EMS locally and in remote powershell sessions
                            @{n = 'SizeInBytes'; e = { $_.TotalItemSize.ToString().split(('(', ')'))[1].replace(',', '').replace(' bytes', '') -as [long] } }
                        )
                        $NOstatsProperties =
                        @{
                            AdminDisplayName         = $null
                            AssociatedItemCount      = $null
                            ContactCount             = $null
                            CreationTime             = $null
                            DeletedItemCount         = 0
                            EntryId                  = $i.EntryID
                            ExpiryTime               = $null
                            FolderPath               = $null
                            Identity                 = $i.Identity
                            IsDeletePending          = $null
                            IsValid                  = $null
                            ItemCount                = 0
                            LastAccessTime           = $null
                            LastModificationTime     = $null
                            LastUserAccessTime       = $null
                            LastUserModificationTime = $null
                            MapiIdentity             = $i.MapiIdentity
                            Name                     = $null
                            OriginatingServer        = $null
                            OwnerCount               = $null
                            ServerName               = $s.ServerFQDN
                            StorageGroupName         = $null
                            TotalAssociatedItemSize  = $null
                            TotalDeletedItemSize     = $null
                            TotalItemSize            = $null
                            DatabaseName             = $s.DatabaseName
                            SizeInBytes              = $null
                        }
                        $count++
                        $currentOperationString = "Getting Stats for $($i.EntryID) from Server $($s.ServerFQDN)"
                        $WriteProgressParams = @{
                            Activity         = 'Retrieving Public Folder Stats for Selected Public Folders'
                            CurrentOperation = $currentOperationString
                            PercentComplete  = $($count / $RecordCount * 100)
                            Status           = "Retrieving Stats for folder replica instance $count of $RecordCount"
                        }
                        Write-Progress @WriteProgressParams
                        $ServerSession = Get-PFMParallelPSSession -name $s.ServerFQDN
                        #makes sure the session is working, if not updates it
                        Confirm-PFMExchangeConnection -IsParallel -PSSession $ServerSession
                        #gets the session again from the $script:ParallelPSsession arraylist
                        $ServerSession = Get-PFMParallelPSSession -name $s.ServerFQDN
                        $ServerName = $s.ServerName
                        $thestats = $(
                            Invoke-Command -Session $ServerSession -ScriptBlock {
                                Get-PublicFolderStatistics -Identity $($using:FolderID).EntryID -Server $using:ServerName -ErrorAction SilentlyContinue
                            }
                        )
                        if ($null -ne $thestats)
                        {
                            $thestats | Select-Object -ExcludeProperty ServerName -Property $customProperties
                        }
                        else
                        {
                            New-Object -TypeName psobject -Property $NOstatsProperties
                        }

                    }
                }
            }
            'Path'
            {
                $count = 0
                $recordCount = $PublicFolderPath.Count * $ServerDatabase.Count
                foreach ($p in $PublicFolderPath)
                {
                    foreach ($s in $ServerDatabase)
                    {
                        $customProperties =
                        @(
                            '*'
                            @{n = 'ServerName'; e = { $s.ServerFQDN } }
                            #this is necessary b/c powershell remoting makes the attributes deserialized and the value in bytes is not available directly.  Code below should work in EMS locally and in remote powershell sessions
                            @{n = 'SizeInBytes'; e = { $_.TotalItemSize.ToString().split(('(', ')'))[1].replace(',', '').replace(' bytes', '') -as [long] } }
                        )
                        $NOstatsProperties =
                        @{
                            AdminDisplayName         = $null
                            AssociatedItemCount      = $null
                            ContactCount             = $null
                            CreationTime             = $null
                            DeletedItemCount         = 0
                            EntryId                  = $null
                            ExpiryTime               = $null
                            FolderPath               = $null
                            Identity                 = $p
                            IsDeletePending          = $null
                            IsValid                  = $null
                            ItemCount                = 0
                            LastAccessTime           = $null
                            LastModificationTime     = $null
                            LastUserAccessTime       = $null
                            LastUserModificationTime = $null
                            MapiIdentity             = $p
                            Name                     = $null
                            OriginatingServer        = $null
                            OwnerCount               = $null
                            ServerName               = $s.ServerFQDN
                            StorageGroupName         = $null
                            TotalAssociatedItemSize  = $null
                            TotalDeletedItemSize     = $null
                            TotalItemSize            = $null
                            DatabaseName             = $s.DatabaseName
                            SizeInBytes              = $null
                        }
                        $count++
                        $currentOperationString = "Getting Stats for $p from Server $($s.ServerFQDN)"
                        $WriteProgressParams = @{
                            Activity         = 'Retrieving Public Folder Stats for Selected Public Folders'
                            CurrentOperation = $currentOperationString
                            PercentComplete  = $($count / $RecordCount * 100)
                            Status           = "Retrieving Stats for folder replica instance $count of $RecordCount"
                        }
                        Write-Progress @WriteProgressParams
                        $ServerSession = Get-PFMParallelPSSession -name $s.ServerFQDN
                        #makes sure the session is working, if not updates it
                        Confirm-PFMExchangeConnection -IsParallel -PSSession $ServerSession
                        #gets the session again from the $script:ParallelPSsession arraylist
                        $ServerSession = Get-PFMParallelPSSession -name $s.ServerFQDN
                        $ServerName = $s.ServerName
                        $thestats = $(
                            Invoke-Command -Session $ServerSession -ScriptBlock {
                                Get-PublicFolderStatistics -Identity $($using:FolderID).EntryID -Server $using:ServerName -ErrorAction SilentlyContinue
                            }
                        )
                        if ($null -ne $thestats)
                        {
                            $thestats | Select-Object -ExcludeProperty ServerName -Property $customProperties
                        }
                        else
                        {
                            New-Object -TypeName psobject -Property $NOstatsProperties
                        }
                    }
                }
            }
        }
    )
    #check for condition where there are no public folders and/or no public folder replicas on the specified servers
    if ($publicFolderStats.Count -eq 0)
    {
        $message = 'There are no public folder replicas hosted on the specified servers.'
        WriteLog -Message $message -EntryType Failed -Verbose -ErrorLog
        Write-Error $message
        return
    }
    else
    {
        WriteLog -Message "Count of Stats objects returned: $($publicFolderStats.count)" -EntryType Notification -Verbose
    }
    #endregion GetPublicFolderStats
    $CreatedFilePath = @(
        foreach ($of in $Outputformat)
        {
            writelog -message "Exporting statistics data to format $of" -entryType Notification -Verbose
            Export-Data -ExportFolderPath $OutputFolderPath -DataToExportTitle 'PublicFolderStats' -ReturnExportFilePath -Encoding $Encoding -DataFormat $of -DataToExport $publicFolderStats -verbose
        }
    )
    WriteLog -Message "Output files created: $($CreatedFilePath -join '; ')" -entryType Notification -verbose
    if ($true -eq $Passthru)
    {
        $publicFolderStats
    }
}