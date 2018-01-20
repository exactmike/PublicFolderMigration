Function Connect-ExchangeOrganization
    {
        [CmdletBinding(DefaultParameterSetName = 'ExchangeOnline')]
        param
        (
            [parameter(Mandatory,ParameterSetName = 'ExchangeOnline')]
            [switch]$ExchangeOnline
            ,
            [parameter(Mandatory,ParameterSetName = 'ExchangeOnPremises')]
            [string]$ExchangeOnPremisesServer
            ,
            [parameter(Mandatory)]
            [pscredential]$Credential
            ,
            [System.Management.Automation.Remoting.PSSessionOption]$PSSessionOption
        )
        $script:Credential = $Credential
        #since this is user facing we always assume that if called the existing session needs to be replaced
        if ($null -ne $script:PsSession -and $script:PsSession -is [System.Management.Automation.Runspaces.PSSession])
        {
            Remove-PSSession -Session $script:PsSession -ErrorAction SilentlyContinue
        }
        $GetExchangePSSessionParams = @{
            ErrorAction = 'Stop'
            Credential = $script:Credential
        }
        if ($null -ne $PSSessionOption)
        {
            $script:PSSessionOption = $PSSessionOption
            $GetExchangePSSessionParams.PSSessionOption = $script:PSSessionOption
        }
        switch ($PSCmdlet.ParameterSetName)
        {
            'ExchangeOnline'
            {
                $Script:OrganizationType = 'ExchangeOnline'
                $GetExchangePSSessionParams.ExchangeOnline = $true
            }
            'ExchangeOnPremises'
            {
                $Script:OrganizationType = 'ExchangeOnPremises'
                $Script:ExchangeOnPremisesServer = $ExchangeOnPremisesServer
                $GetExchangePSSessionParams.ExchangeServer = $script:ExchangeOnPremisesServer
            }
        }
        $script:PsSession = GetExchangePSSession @GetExchangePSSessionParams
        $script:ConnectExchangeOrganizationCompleted = $true
    }
#end function Connect-ExchangeOrganization
function Get-PublicFolderReplicationReport
    {
        <#
        .SYNOPSIS
        Generates a report for Exchange 2010 Public Folder Replication.
        .DESCRIPTION
        This script will generate a report for Exchange 2010 Public Folder Replication. It returns general information, such as total number of public folders, total items in all public folders, total size of all items, the top 10 largest folders, and more. Additionally, it lists each Public Folder and the replication status on each server. By default, this script will scan the entire Exchange environment in the current domain and all public folders. This can be limited by using the -PublicFolderMailboxServer and -PublicFolderPath parameters.
        .PARAMETER PublicFolderMailboxServer
        This parameter specifies the Exchange 2010 server(s) to scan. If this is omitted, all Exchange servers hosting a Public Folder Database are scanned.
        .PARAMETER PublicFolderPath
        This parameter specifies the Public Folder(s) to scan. If this is omitted, all public folders are scanned.
        .PARAMETER Recurse
        When used in conjunction with the FolderPath parameter, this will include all child Public Folders of the Folders listed in Folder Path.
        .PARAMETER AsHTML
        Specifying this switch will have this script output HTML, rather than the result objects. This is independent of the Filename or SendEmail parameters and only controls the console output of the script.
        .PARAMETER Passthrough
        Controls whether the ReportMatrix of Public Folder Stats is returned to the pipeline instead of just being consumed in output to email, file, or html. 
        .PARAMETER Filename
        Providing a Filename will save the HTML report to a file.
        .PARAMETER SendEmail
        This switch will set the script to send an HTML email report. If this switch is specified, then the To, From and SmtpServers are required.
        .PARAMETER To
        When SendEmail is used, this sets the recipients of the email report.
        .PARAMETER From
        When SendEmail is used, this sets the sender of the email report.
        .PARAMETER SmtpServer
        When SendEmail is used, this is the SMTP Server to send the report through.
        .PARAMETER Subject
        When SendEmail is used, this sets the subject of the email report.
        .PARAMETER IncludeSystemPublicFolders
        This parameter specifies to include System Public Folders when scanning all public folders. If this is omitted, System Public Folders are omitted.
        .PARAMETER LargestPublicFolderReportCount
        This parameter allows control of the count largest public folders data in the report object.
        #>
        [cmdletbinding()]
        param
        (
            [parameter()]
            [string[]]$PublicFolderMailboxServer = @()
            ,
            [parameter()]
            [string[]]$PublicFolderPath = @()
            ,
            [parameter()]
            [switch]$Recurse
            ,
            [parameter()]
            [switch]$IncludeSystemPublicFolders
            ,
            [parameter()]
            [validateset('RawReplicationData','ReportObject')]
            [string]$PipelineData
            ,
            [parameter(Mandatory)]
            [ValidateScript({TestIsWriteableDirectory -path $_})]
            [string]$outputFolderPath
            ,
            [parameter()]
            [ValidateScript({$_ | TestEmailAddress})]
            [string[]]$To
            ,
            [parameter()]
            [ValidateScript({TestEmailAddress -EmailAddress $_})]
            [string]$From
            ,
            [parameter()]
            [ValidateScript({TestTCPConnection -port 25 -ComputerName $_})]
            [string]$SmtpServer
            ,
            [parameter()]
            [string]$Subject
            ,
            [parameter()]
            [switch]$HTMLBody
            ,
            [parameter()]
            [validateset('html','csv')]
            [string[]]$outputformats
            ,
            [parameter()]
            [switch]$SendEmail
            ,
            [parameter()]
            [int]$LargestPublicFolderReportCount = 10
        )
        Begin
        {
            switch ($script:ConnectExchangeOrganizationCompleted)
            {
                $true
                {
                    switch (TestExchangePSSession -PSSession $script:PSSession)
                    {
                        $true
                        {
                            WriteLog -Message 'Using Existing PSSession' -EntryType Notification
                        }
                        $false
                        {
                            WriteLog -Message 'Removing Existing Failed PSSession' -EntryType Notification
                            Remove-PSSession -Session $script:PsSession -ErrorAction SilentlyContinue
                            WriteLog -Message 'Establishing New PSSession to Exchange Organization' -EntryType Notification
                            $GetExchangePSSessionParams = GetGetExchangePSSessionParams
                            $script:PsSession = GetExchangePSSession @GetExchangePSSessionParams
                        }
                    }
                }
                $false
                {
                    WriteUserInstructionError
                }
            }
            $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
            $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'PublicFolderReplicationAndStatisticsReport.log')
            $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'PublicFolderReplicationAndStatisticsReport-ERRORS.log')
            WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
            $ExchangeOrganization = Invoke-Command -Session $Script:PSSession -ScriptBlock {Get-OrganizationConfig | Select-Object -ExpandProperty Identity | Select-Object -ExpandProperty Name}
            WriteLog -Message "Exchange Session is Running in Exchange Organzation $ExchangeOrganization" -EntryType Notification
            #region ValidateParameters
                if ($true -eq $SendEmail)
                {
                       if (-not $To.Count -gt 0)
                    {
                        Write-Error 'The -To parameter is required when including email as an output method. If this parameter was used, verify that valid email addresses were specified.'
                        return
                    }
                    if ([string]::IsNullOrEmpty($From))
                    {
                        Write-Error 'The -From parameter is not valid. This parameter is required when including email as an output method.'
                        return
                    }
                    if ([string]::IsNullOrEmpty($SmtpServer))
                    {
                        Write-Error 'You must specify a SmtpServer. This parameter is required when including email as an output method.'
                        return
                    }
                }#end if $SendEmail
                #if the user specified public folder mailbox servers, validate them:
                if ($PublicFolderMailboxServer.Count -ge 1) 
                {
                    foreach ($Server in $PublicFolderMailboxServer) {
                        $VerifyPFDatabase = @(
                            Invoke-Command -Session $script:PSSession -scriptblock {
                                Get-PublicFolderDatabase -server $using:Server -IncludePreExchange2010 -ErrorAction SilentlyContinue
                            }
                        )
                        if ($VerifyPFDatabase.Count -ne 1) {
                            Write-Error "$server is either not a Mailbox server or does not host a public folder database."
                            Return
                        }
                    }
                }#if publicFolderMailboxServer count over 0
                #if the user did not specify the public folder mailbox servers to include, include all of them
                if ($PublicFolderMailboxServer.Count -lt 1)
                {
                    $PublicFolderMailboxServer = @(
                        Invoke-Command -Session $script:PSSession -ScriptBlock {
                            Get-PublicFolderDatabase -includePreExchange2010 | Select-Object -ExpandProperty ServerName
                        }
                    )
                }
            #endregion ValidateParameters
        }#Begin
        End
        {
            #region BuildServerAndDatabaseLists
            $PublicFolderMailboxServerNames = $PublicFolderMailboxServer -join ', '
            WriteLog -Message "Public Folder Mailbox Servers Included: $PublicFolderMailboxServerNames" -EntryType Notification -Verbose
            #Build Server/Database Hash Tables for later reporting activities
            $PublicFolderMailboxServerDatabases = @{}
            $PublicFolderDatabaseMailboxServers = @{}
            foreach ($server in $PublicFolderMailboxServer) 
            {
                $PublicFolderDatabase = $(
                    Invoke-Command -Session $script:PSSession -ScriptBlock {
                        Get-PublicFolderDatabase -Server $Using:Server -includePreExchange2010   
                    }
                )
                $PublicFolderMailboxServerDatabases.$Server = $PublicFolderDatabase.Name
                $PublicFolderDatabaseMailboxServers.$($PublicFolderDatabase.Name) = $Server
            }
            #endregion BuildServerAndDatabaseLists
            #region BuildPublicFolderList
            #Set up the parameters for Get-PublicFolder
            $GetPublicFolderParams = @{}
            if ($Recurse) {
                $GetPublicFolderParams.Recurse = $true
                $GetPublicFolderParams.ResultSize = 'Unlimited'
            }
            $FolderIDs = @(
                #if the user specified specific public folder paths, get those
                if ($PublicFolderPath.Count -ge 1) 
                {
                    $publicFolderPathString = $PublicFolderPath -join ', '
                    WriteLog -Message "Retrieving Public Folders in the following Path(s): $publicFolderPathString" -EntryType Notification
                    foreach($Path in $PublicFolderPath) {
                        Invoke-Command -Session $script:PSSession -ScriptBlock {
                            Get-PublicFolder $using:Path @using:GetPublicFolderParams
                        } | Select-Object -property @{n='EntryID';e={$_.EntryID.tostring()}},@{n='Identity';e={$_.Identity.tostring()}},Name,Replicas
                    }
                }
                #otherwise, get all default public folders
                else 
                {
                    WriteLog -message 'Retrieving All Default (Non-System) Public Folders from IPM_SUBTREE' -EntryType Notification
                    Invoke-Command -Session $script:PSSession -ScriptBlock {
                        Get-PublicFolder -Recurse -ResultSize Unlimited
                    } | Select-Object -property @{n='EntryID';e={$_.EntryID.tostring()}},@{n='Identity';e={$_.Identity.tostring()}},Name,Replicas
                    if ($IncludeSystemPublicFolders) {
                        WriteLog -Message 'Retrieving All System Public Folders from NON_IPM_SUBTREE' -EntryType Notification
                        Invoke-Command -Session $script:PSSession -ScriptBlock {
                            Get-PublicFolder \Non_IPM_SUBTREE -Recurse -ResultSize Unlimited
                        } | Select-Object -property @{n='EntryID';e={$_.EntryID.tostring()}},@{n='Identity';e={$_.Identity.tostring()}},Name,Replicas
                    }
                }
            )
            #filter any duplicates if the user specified public folder paths
            WriteLog -Message 'Sorting and De-duplicating retrieved Public Folders.' -EntryType Notification
            if ($PublicFolderPath.Count -ge 1) {$FolderIDs = @($FolderIDs | Select-Object -Unique -Property *)}
            #sort folders by path
            $FolderIDs = @($FolderIDs | Sort-Object Identity)
            $publicFoldersRetrievedCount = $FolderIDs.Count
            WriteLog -Message "Count of Public Folders Retrieved: $publicFoldersRetrievedCount" -EntryType Notification
            #endregion BuildPublicFolderList
            #region GetPublicFolderStats
            $publicFolderStatsFromSelectedServers = 
            @(
                # if the user specified public folder path then only retrieve stats for the specified folders.  
                # This can be significantly faster than retrieving stats for all public folders
                if ($PublicFolderPath.Count -ge 1) 
                {
                    $count = 0
                    $RecordCount = $FolderIDs.Count * $PublicFolderMailboxServer.Count
                    foreach ($FolderID in $FolderIDs)
                    { 
                        foreach ($Server in $PublicFolderMailboxServer)
                        {
                            $customProperties = 
                            @(
                                '*'
                                @{n='ServerName';e={$Server}}
                                #this is necessary b/c powershell remoting makes the attributes deserialized and the value in bytes is not available directly.  Code below should work in EMS locally and in remote powershell sessions
                                @{n='SizeInBytes';e={$_.TotalItemSize.ToString().split(('(',')'))[1].replace(',','').replace(' bytes','') -as [long]}}
                            )
                            $NOstatsProperties = 
                            @{
                                AdminDisplayName = $null
                                AssociatedItemCount = $null
                                ContactCount = $null
                                CreationTime = $null
                                DeletedItemCount = 0
                                EntryId = $FolderID.EntryID
                                ExpiryTime = $null
                                FolderPath = $null
                                Identity = $FolderID.EntryID
                                IsDeletePending = $null
                                IsValid = $null
                                ItemCount = 0
                                LastAccessTime = $null
                                LastModificationTime = $null
                                LastUserAccessTime = $null
                                LastUserModificationTime = $null
                                MapiIdentity =  $FolderID.Name
                                Name = $FolderID.Name
                                OriginatingServer = $null
                                OwnerCount = $null
                                ServerName = $Server
                                StorageGroupName = $null
                                TotalAssociatedItemSize = $null
                                TotalDeletedItemSize = $null
                                TotalItemSize = $null
                                DatabaseName = $($PublicFolderMailboxServerDatabases.$Server)
                                SizeInBytes = $null
                            }
                            if ($FolderID.Replicas -contains $PublicFolderMailboxServerDatabases.$Server) 
                            {
                                $count++
                                $currentOperationString = "Getting Stats for $($FolderID.Identity) from Server $Server."
                                Write-Progress -Activity 'Retrieving Public Folder Stats for Selected Public Folders' -CurrentOperation $currentOperationString -PercentComplete $($count/$RecordCount*100) -Status "Retrieving Stats for folder replica instance $count of $RecordCount"
                                WriteLog -Message $currentOperationString -EntryType Notification -Verbose
                                #Error Action Silently Continue because some servers may not have a replica and we don't care about that error in this context
                                $thestats = $(
                                    Invoke-Command -Session $script:PSSession -ScriptBlock {
                                        Get-PublicFolderStatistics -Identity $($using:FolderID).EntryID -Server $using:Server -ErrorAction SilentlyContinue 
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
                            else 
                            {
                                New-Object -TypeName psobject -Property $NOstatsProperties
                            }
                        }#foreach $Server
                    }#foreach $FolderID
                    Write-Progress -Activity 'Retrieving Public Folder Stats for Selected Public Folders' -CurrentOperation Completed -Completed -Status Completed
                }
                # Get statistics for all public folders on all selected servers
                # This is significantly faster than trying to get folders one by one by name
                else 
                {
                    $count = 0
                    $RecordCount = $PublicFolderMailboxServer.Count
                    foreach ($Server in $PublicFolderMailboxServer) {
                        $customProperties = @(
                            '*'
                            @{n='ServerName';e={$Server}}
                            @{n='SizeInBytes';e={$_.TotalItemSize.ToString().split(('(',')'))[1].replace(',','').replace(' bytes','') -as [long]}}
                        )
                        Write-Verbose "Retrieving Stats for all Public Folders from $Server"
                        Write-Progress -Activity 'Retrieving Public Folder Stats' -CurrentOperation $Server -PercentComplete $($count/$RecordCount*100) -Status "Retrieving Stats for Server $count of $RecordCount"
                        #avoid NULL output by testing for results while still suppressing errors with SilentlyContinue
                        $thestats = @(
                            Invoke-Command -Session $script:PSSession -ScriptBlock {
                                Get-PublicFolderStatistics -Server $using:Server -ResultSize Unlimited -ErrorAction SilentlyContinue 
                            }
                        )
                        if ($thestats -ne $null)
                        {
                            $thestats | Select-Object -ExcludeProperty ServerName -Property $customProperties
                        }
                    }
                } 
            )
            #check for condition where there are no public folders and/or no public folder replicas on the specified servers
            if ($publicFolderStatsFromSelectedServers.Count -eq 0)
            {
                $message = 'There are no public folder replicas hosted on the specified servers.'
                WriteLog -Message $message -EntryType Failed -Verbose -ErrorLog
                Write-Error $message
                return
            }
            else 
            {
                WriteLog -Message "Count of Stats objects returned: $($publicFolderStatsFromSelectedServers.count)" -EntryType Notification -Verbose
            }
            #endregion GetPublicFolderStats
            #region BuildStatsLookupHash
            #create the hash table
            $publicFolderStatsLookup = @{}
            #Populate the hashtable - one key/value pair per EntryID plus Server
            foreach ($Stats in ($publicFolderStatsFromSelectedServers | where-object -FilterScript {$_.EntryID -ne $NULL})) {
                $Key = $($Stats.EntryID.tostring() + '_' + $Stats.ServerName)
                $Value = $Stats;
                $PublicFolderStatsLookup.$Key = $Value
            }
            #endregion BuildStatsLookupHash
            #region BuildResultMatrix
            $ResultMatrix = 
            @(
                $count = 0
                $RecordCount = $FolderIDs.Count
                foreach($Folder in $FolderIDs)
                { 
                    $count++
                    $currentOperationString= "Processing Report for Folder $($folder.EntryID) with name $($Folder.Identity)"
                    Write-Progress -Activity 'Building Data Matrix of Public Folder Stats for output and reporting.' -Status 'Compiling Data' -CurrentOperation $currentOperationString -PercentComplete ($count/$RecordCount*100)
                    #WriteLog -Message $currentOperationString -EntryType Notification -Verbose
                    $resultItem = @{
                        EntryID = $Folder.EntryID
                        FolderPath = $Folder.Identity
                        Name = $folder.name
                        ConfiguredReplicas = $($folder.replicas -join ',')
                        Data = @(
                            #Get all the stats entries for this folder from each server using the EntryID + Server Key lookup
                            foreach ($Server in $PublicFolderMailboxServer) 
                            {
                                $publicFolderStatsLookup.$($Folder.EntryID + '_' + $Server) | Where-Object -FilterScript {$_} |
                                ForEach-Object {
                                    New-Object PSObject -Property @{
                                            AdminDisplayName = $_.AdminDisplayName
                                            AssociatedItemCount = $_.AssociatedItemCount
                                            ContactCount = $_.ContactCount
                                            CreationTime = $_.CreationTime
                                            DeletedItemCount = $_.DeletedItemCount
                                            EntryId = $_.EntryID.tostring()
                                            ExpiryTime = $_.ExpiryTime
                                            FolderPath = $_.FolderPath
                                            Identity = $_.Identity.tostring()
                                            IsDeletePending = $_.IsDeletePending
                                            IsValid = $_.IsValid
                                            ItemCount = $_.ItemCount
                                            LastAccessTime = $_.LastAccessTime
                                            LastModificationTime = $_.LastModificationTime
                                            LastUserAccessTime = $_.LastUserAccessTime
                                            LastUserModificationTime = $_.LastUserModificationTime
                                            MapiIdentity =  $_.MapiIdentity
                                            Name = $_.Name
                                            OwnerCount = $_.OwnerCount
                                            TotalAssociatedItemSize = $_.TotalAssociatedItemSize
                                            TotalDeletedItemSize = $_.TotalDeletedItemSize
                                            ServerName = $_.ServerName
                                            DatabaseName = $_.DatabaseName
                                            TotalItemSize = $_.TotalItemSize
                                            SizeInBytes = $_.SizeInBytes
                                    }
                                }
                            }
                        )
                    }
                    #Get Max Total Item Size in Bytes across Replicas 
                    $resultItem.TotalBytes = $resultItem.Data | Measure-Object -Property SizeInBytes -Maximum | Select-Object -ExpandProperty Maximum
                    #Get Max Total Item Size human friendly based on max Bytes
                    $resultItem.TotalItemSize = $resultItem.Data | Where-Object -FilterScript {$_.SizeInBytes -eq $resultItem.TotalBytes} | Select-Object -First 1 -ExpandProperty TotalItemSize
                    #Get Max Item Count
                    $resultItem.ItemCount = $resultItem.Data | Measure-Object -Property ItemCount -Maximum | Select-Object -ExpandProperty Maximum
                    $resultItem.LastAccessTime = $resultItem.Data | Measure-Object -Property LastAccessTime -Maximum | Select-Object -ExpandProperty Maximum
                    $resultItem.LastModificationTime = $resultItem.Data | Measure-Object -Property LastModificationTime -Maximum | Select-Object -ExpandProperty Maximum
                    $resultItem.LastUserAccessTime = $resultItem.Data | Measure-Object -Property LastUserAccessTime -Maximum | Select-Object -ExpandProperty Maximum
                    $resultItem.LastUserModificationTime = $resultItem.Data | Measure-Object -Property LastUserModificationTime -Maximum | Select-Object -ExpandProperty Maximum
                    $replCheck = $true
                    foreach($dataRecord in $resultItem.Data) {
                        if ($resultItem.ItemCount -eq 0 -or $resultItem.ItemCount -eq $null)
                        {
                            $progress = 100
                        }
                        else
                        {
                            try 
                            {
                                $ErrorActionPreference = 'Stop'
                                $progress = ([Math]::Round($dataRecord.ItemCount / ($resultItem.ItemCount)  * 100, 0))
                                $ErrorActionPreference = 'Continue'
                            }
                            catch
                            {
                                $progress = $null
                                WriteLog -Message "Server: $($dataRecord.Server), Database: $($dataRecord.Databasename), ItemCount: $($dataRecord.ItemCount), TotalItemCount: $($resultItem.ItemCount)" -EntryType Failed -ErrorLog
                                WriteLog -Message $_.tostring() -Verbose -ErrorLog
                                $ErrorActionPreference = 'Continue'
                            }
                        }
                        if ($progress -lt 100)
                        {
                            $replCheck = $false
                        }
                        $dataRecord | Add-Member -MemberType NoteProperty -Name 'Progress' -Value $progress
                    }
                    $resultItem.ReplicationCompleteOnIncludedServers=$replCheck
                    #output result object
                    New-Object PSObject -Property $resultItem
                }#Foreach
                Write-Progress -Activity 'Building Data Matrix of Public Folder Stats for output and reporting.' -Status 'Compiling Data' -CurrentOperation $currentOperationString -Completed
            )#$ResultMatrix
            #endregion BuildResultMatrix
            #Build the Report Object
            [pscustomobject]$ReportObject = @{
            #region BuildReportObject
                TimeStamp = Get-Date -Format yyyyMMdd-HHmm
                IncludedPublicFolderServersAndDatabases = $($(foreach ($server in $PublicFolderMailboxServer) {"$Server ($($PublicFolderMailboxServerDatabases.$server))"}) -join ',')
                IncludedPublicFoldersCount = $ResultMatrix.Count
                TotalSizeOfIncludedPublicFoldersInBytes = $ResultMatrix | Measure-Object -Property TotalBytes -Sum | Select-Object -ExpandProperty Sum
                TotalItemCountFromIncludedPublicFolders = $ResultMatrix | Measure-Object -Property ItemCount -Sum | Select-Object -ExpandProperty Sum
                IncludedContainerOrEmptyPublicFoldersCount = @($ResultMatrix | Where-Object -FilterScript {$_.ItemCount -eq 0}).Count
                IncludedReplicationIncompletePublicFolders = @($ResultMatrix | Where-Object -FilterScript {$_.ReplicationCompleteOnIncludedServers -eq $false}).Count
                LargestPublicFolders = @($ResultMatrix | Sort-Object TotalBytes -Descending | Select-Object -First $LargestPublicFolderReportCount)
                PublicFoldersWithIncompleteReplication = @(
                    Foreach ($result in ($ResultMatrix | Where-Object -FilterScript {$_.ReplicationCompleteOnIncludedServers -eq $false})) 
                    {
                        [pscustomobject]@{
                            EntryID = $result.EntryID
                            FolderPath = $Result.FolderPath
                            ItemCount = $Result.ItemCount
                            TotalItemSize = $Result.TotalItemSize
                            ConfiguredReplicaDatabases = $result.ConfiguredReplicas
                            ConfiguredReplicaServers = 
                            $(
                                $databases = $result.ConfiguredReplicas.split(',')
                                $servers = $databases | foreach {$PublicFolderDatabaseMailboxServers.$_}
                                $Servers -join ','
                            )
                            CompleteServers = 
                            $(
                                $CompleteServers = $result.Data | Where-Object {$_.Progress -eq 100} | Select-Object -ExpandProperty ServerName
                                $CompleteServers -join ','
                            )
                            CompleteDatabases = 
                            $(
                                $CompleteDatabases = $result.Data | Where-Object {$_.Progress -eq 100} | Select-Object -ExpandProperty ServerName
                                $CompleteDatabases -join ','
                            )
                            IncompleteServers = 
                            $(
                                $IncompleteServers = $result.Data | Where-Object {$_.Progress -lt 100} | Select-Object -ExpandProperty ServerName
                                $IncompleteServers -join ','
                            )
                            IncompleteDatabases = 
                            $(
                                $IncompleteDatabases = $result.Data | Where-Object {$_.Progress -lt 100} | Select-Object -ExpandProperty DatabaseName
                                $IncompleteDatabases -join ','
                            )
                        }#pscustomobject
                    }#Foreach
                )
                ReplicationReportByServerPercentage = @(
                    Foreach ($result in $ResultMatrix) 
                    {
                        $RRObject = [pscustomobject]@{
                            FolderPath = $result.FolderPath
                            EntryID = $result.EntryID
                            HighestItemCount = $result.ItemCount
                            HighestBytesCount = $result.totalBytes
                        }#pscustomobject
                        Foreach ($Server in $PublicFolderMailboxServer) 
                        {
                            $ResultItem = $result.Data | Where-Object -FilterScript {$_.ServerName -eq $Server}
                            $PropertyName1 = $Server + '-%'
                            $PropertyName2 = $Server + '-Count'
                            $PropertyName3 = $server + '-SizeInBytes'
                            if ($resultItem -eq $null) 
                            {
                                $RRObject | Add-Member -NotePropertyName $PropertyName1 -NotePropertyValue 'N/A'
                                $RRObject | Add-Member -NotePropertyName $PropertyName2 -NotePropertyValue 'N/A'
                                $RRObject | Add-Member -NotePropertyName $PropertyName3 -NotePropertyValue 'N/A'
                            }#if
                            else 
                            {
                                $RRObject | Add-Member -NotePropertyName $PropertyName1 -NotePropertyValue $resultItem.Progress
                                $RRObject | Add-Member -NotePropertyName $PropertyName2 -NotePropertyValue $resultItem.itemCount
                                $RRObject | Add-Member -NotePropertyName $PropertyName3 -NotePropertyValue $resultItem.SizeInBytes
                            }#else
                        }#Foreach
                    $RRObject
                    }#Foreach
                )
            }
            $ReportObject.NonContainerOrEmptyPublicFoldersCount = $ReportObject.IncludedPublicFoldersCount - $ReportObject.IncludedContainerOrEmptyPublicFoldersCount
            $ReportObject.AverageSizeOfIncludedPublicFolders = [Math]::Round($ReportObject.TotalSizeOfIncludedPublicFoldersInBytes/$ReportObject.NonContainerOrEmptyPublicFoldersCount, 0)
            $ReportObject.AverageItemCountFromIncludedPublicFolders = [Math]::Round($ReportObject.TotalItemCountFromIncludedPublicFolders / $ReportObject.NonContainerOrEmptyPublicFoldersCount, 0)
            #endregion BuildReportObject
            #region PipelineDataOutput
            if (-not [string]::IsNullOrWhiteSpace($PipelineData)) {
                switch ($PipelineData)
                {
                    'RawReplicationData' 
                    {$ResultMatrix}
                    'ReportObject'
                    {$ReportObject}
                }
                #$ReportObject
            }#if $passthrough - output the report data as objects
            #endregion PipelineDataOutput
            #region GenerateHTMLOutput
            if (('html' -in $outputformats) -or $HTMLBody)
            {
                $html = GetHTMLReport -ReportObject $ReportObject -ResultMatrix $ResultMatrix -PublicFolderMailboxServer $PublicFolderMailboxServer
            }#if to generate HTML output if required/requested
            #endregion GenerateHTMLOutput
            #region GenerateOutputFormats
            $outputfiles = @(
                if ('csv' -in $outputformats)
                {
                    $CSVOutputReports = @{
                        PubliFolderEnvironmentSummary = [pscustomobject]@{
                            ReportTimeStamp = $ReportObject.TimeStamp
                            IncludedPublicFolderServersAndDatabases = $ReportObject.IncludedPublicFolderServersAndDatabases
                            IncludedPublicFoldersCount = $ReportObject.IncludedPublicFoldersCount
                            TotalSizeOfIncludedPublicFoldersInBytes = $ReportObject.TotalSizeOfIncludedPublicFoldersInBytes
                            TotalItemCountFromIncludedPublicFolders = $ReportObject.TotalItemCountFromIncludedPublicFolders
                            IncludedContainerOrEmptyPublicFoldersCount = $ReportObject.IncludedContainerOrEmptyPublicFoldersCount
                            IncludedReplicationIncompletePublicFolders = $ReportObject.IncludedReplicationIncompletePublicFolders
                        }
                        LargestPublicFolders = $ReportObject.LargestPublicFolders | Select-Object FolderPath,TotalItemSize,ItemCount
                        PublicFoldersWithIncompleteReplication = $ReportObject.PublicFoldersWithIncompleteReplication
                        ReplicationReportPerReplicaDetails = $ReportObject.ReplicationReportByServerPercentage
                        PublicFolderStatisticsFromAllReplicas = $resultMatrix | foreach-object {
                            $parent = $_
                            $parent.data | foreach-object {
                                [pscustomobject]@{
                                    EntryID = $parent.EntryID
                                    Name = $parent.Name
                                    FolderPath = $parent.FolderPath
                                    ConfiguredReplicas = $parent.ConfiguredReplicas
                                    MaxTotalBytes = $Parent.TotalBytes
                                    MaxItemCount = $Parent.ItemCount
                                    MaxLastAccessTime = $Parent.LastAccessTime
                                    MaxLastModificationTime = $Parent.LastModificationTime
                                    MaxLastUserAccessTime = $Parent.LastUserAccessTime
                                    MaxLastUserModificationTime = $Parent.LastUserModificationTime
                                    AdminDisplayName = $_.AdminDisplayName
                                    AssociatedItemCount = $_.AssociatedItemCount
                                    ContactCount = $_.ContactCount
                                    CreationTime = $_.CreationTime
                                    DatabaseName = $_.DatabaseName
                                    DeletedItemCount = $_.DeletedItemCount
                                    ExpiryTime = $_.ExpiryTime
                                    Identity = $_.Identity
                                    IsDeletePending = $_.IsDeletePending
                                    IsValid = $_.IsValid
                                    ItemCount = $_.ItemCount
                                    LastAccessTime = $_.LastAccessTime
                                    LastModificationTime = $_.LastModificationTime
                                    LastUserAccessTime = $_.LastUserAccessTime
                                    LastUserModificationTime = $_.LastUserModificationTime
                                    MapiIdentity = $_.MapiIdentity
                                    OwnerCount = $_.OwnerCount
                                    Progress = $_.Progress
                                    ServerName = $_.ServerName
                                    SizeInBytes = $_.SizeInBytes
                                    TotalAssociatedItemSize = $_.TotalAssociatedItemSize
                                    TotalDeletedItemSize = $_.TotalDeletedItemSize
                                    TotalItemSize = $_.TotalItemSize
                                }
                            }
                        }
                    }#end CSVOutputReports
                    foreach ($key in $CSVOutputReports.keys)
                    {
                        $outputFileName = $BeginTimeStamp + $key + '.csv'
                        $outputFilePath = Join-Path -path $outputFolderPath -ChildPath $outputFileName 
                        $CSVOutputReports.$key | Export-CSV -path $outputFilePath -Encoding UTF8 -NoTypeInformation
                        $outputFilePath
                    }
                }
                if ('html' -in $outputformats)
                {
                    $HTMLFileName = $BeginTimeStamp + 'PublicFolderEnvironmentAndReplicationReport.html'
                    $HTMLFilePath = Join-Path -path $outputFolderPath -ChildPath $HTMLFileName
                    $html | Out-File -FilePath $HTMLFilePath -Encoding utf8
                    $HTMLFilePath
                }
            )
            #endregion GenerateOutputFormats
            #region SendMail
            if ($true -eq $SendEmail)
            {
                if ([string]::IsNullOrEmpty($Subject))
                {
                    $Subject = 'Public Folder Environment and Replication Status Report'
                }
                $SendMailMessageParams = @{
                    Subject = $Subject
                    Attachments = $outputfiles
                    To = $to
                    From = $from
                    SMTPServer = $SmtpServer
                }
                if ($HTMLBody)
                {
                    $SendMailMessageParams.BodyAsHTML
                    $SendMailMessageParams.Body = $html
                }
                else
                {
                    $SendMailMessageParams.Body = "Public Folder Environment and Replication Status Report Attached."
                }
                Send-MailMessage @SendMailMessageParams
            }#end if $SendMail
            #endregion SendMail
        }#end
    }
#end function Get-PublicFolderReplicationReport