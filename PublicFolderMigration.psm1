###############################################################################################
#Core Public Folder Migration Module Functions
###############################################################################################
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
        .PARAMETER NoAttachment
        When SendEmail is used, specifying this switch will set the email report to not include the HTML Report as an attachment. It will still be sent in the body of the email.
        .PARAMETER IncludeSystemPublicFolders
        This parameter specifies to include System Public Folders when scanning all public folders. If this is omitted, System Public Folders are omitted.
        .PARAMETER LargestPublicFolderReportCount
        This parameter allows control of the count largest public folders data in the report object.
        #>
        [cmdletbinding()]
        param(
            [string[]]$PublicFolderMailboxServer = @()
            ,
            [string[]]$PublicFolderPath = @()
            ,
            [switch]$Recurse
            ,
            [switch]$IncludeSystemPublicFolders
            ,
            [parameter()]
            [validateset('RawReplicationData','ReportObject')]
            [string]$PipelineData
            ,
            [string]$FileFolderPath
            ,
            [string[]]$To
            ,
            [string]$From
            ,
            [string]$SmtpServer
            ,
            [string]$Subject
            ,
            [switch]$HTMLBody
            ,
            [parameter()]
            [validateset('html','csv')]
            [string[]]$outputformats
            ,
            [parameter()]
            [validateset('email','files')]
            [string[]]$outputmethods
            ,
            [int]$LargestPublicFolderReportCount = 10
        )
        Begin
        {
            #region ValidateParameters
                if ('email' -in $outputmethods)
                {
                    [array]$newTo = @()
                    foreach($recipient in $To)
                    {
                        if ($recipient -imatch "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z0-9.-]+$")
                        {
                            $newTo += $recipient
                        }
                    }
                    $To = $newTo
                    if (-not $To.Count -gt 0)
                    {
                        Write-Error 'The -To parameter is required when using the -SendEmail switch. If this parameter was used, verify that valid email addresses were specified.'
                        return
                    }
                    if ($From -inotmatch "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z0-9.-]+$")
                    {
                        Write-Error 'The -From parameter is not valid. This parameter is required when using the -SendEmail switch.'
                        return
                    }
                    if ([string]::IsNullOrEmpty($SmtpServer))
                    {
                        Write-Error 'You must specify a SmtpServer. This parameter is required when using the -SendEmail switch.'
                        return
                    }
                    if ((Test-Connection $SmtpServer -Quiet -Count 2) -ne $true)
                    {
                        Write-Error "The SMTP server specified ($SmtpServer) could not be contacted."
                        return
                    }
                }#end if email in outputmethods
                if ('files' -in $outputmethods -or 'email' -in $outputmethods)
                {
                    if (-not (Test-Path -Path $FileFolderPath))
                    {
                        Write-Error "$FileFolderPath failed validation."
                        Return
                    }
                    if ($FileFolderPath -notlike '*\'){
                        $FileFolderPath = $FileFolderPath + '\'
                    }
                }
                #if the user specified public folder mailbox servers, validate them:
                if ($PublicFolderMailboxServer.Count -ge 1) 
                {
                    foreach ($Server in $PublicFolderMailboxServer) {
                        $VerifyPFDatabase = @(Get-PublicFolderDatabase -server $Server -IncludePreExchange2010)
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
                        Get-PublicFolderDatabase -includePreExchange2010 | Select-Object -ExpandProperty ServerName
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
                $PublicFolderDatabase = Get-PublicFolderDatabase -Server $Server -includePreExchange2010
                $PublicFolderMailboxServerDatabases.$Server = $PublicFolderDatabase | Select-Object -ExpandProperty Name
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
                    WriteLog -Message "Retrieving Public Folders in the following Path(s): $publicFolderPathString" -EntryType Notification -Verbose
                    foreach($Path in $PublicFolderPath) {
                        Get-PublicFolder $Path @GetPublicFolderParams | Select-Object -property @{n='EntryID';e={$_.EntryID.tostring()}},@{n='Identity';e={$_.Identity.tostring()}},Name,Replicas
                    }
                }
                #otherwise, get all default public folders
                else 
                {
                    WriteLog -message 'Retrieving All Default (Non-System) Public Folders from IPM_SUBTREE' -EntryType Notification -Verbose
                    Get-PublicFolder -Recurse -ResultSize Unlimited | Select-Object -property @{n='EntryID';e={$_.EntryID.tostring()}},@{n='Identity';e={$_.Identity.tostring()}},Name,Replicas
                    if ($IncludeSystemPublicFolders) {
                        WriteLog -Message 'Retrieving All System Public Folders from NON_IPM_SUBTREE' -EntryType Notification -Verbose
                        Get-PublicFolder \Non_IPM_SUBTREE -Recurse -ResultSize Unlimited | Select-Object -property @{n='EntryID';e={$_.EntryID.tostring()}},@{n='Identity';e={$_.Identity.tostring()}},Name,Replicas
                    }
                }
            )
            #filter any duplicates if the user specified public folder paths
            WriteLog -Message 'Sorting and De-duplicating retrieved Public Folders.' -EntryType Notification -Verbose
            if ($PublicFolderPath.Count -ge 1) {$FolderIDs = @($FolderIDs | Select-Object -Unique -Property *)}
            #sort folders by path
            $FolderIDs = @($FolderIDs | Sort-Object Identity)
            $publicFoldersRetrievedCount = $FolderIDs.Count
            WriteLog -Message "Count of Public Folders Retrieved: $publicFoldersRetrievedCount" -EntryType Notification -Verbose
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
                    foreach ($FolderID in $FolderIDs) { 
                        foreach ($Server in $PublicFolderMailboxServer){
                            $customProperties = 
                            @(
                                '*'
                                @{n='ServerName';e={$Server}}
                                #this is necessary b/c powershell remoting makes the attributes deserialized and the value in bytes is not available directly.  Code below should work in EMS locally and in remote powershell sessions
                                @{n='SizeInBytes';e={$_.TotalItemSize.ToString().split(('(',')'))[1].replace(',','').replace(' bytes','') -as [long]}}
                            )
                            $NOstatsProperties = 
                            @{
                                'ServerName'=$Server
                                'SizeInBytes'=$null
                                'Progress'=0
                                'ItemCount'=0
                                'TotalItemSize'=$null
                                'DatabaseName'=$PublicFolderMailboxServerDatabases.$Server
                                'LastModificationTime'=$null
                        }
                            if ($FolderID.Replicas -contains $PublicFolderMailboxServerDatabases.$Server) 
                            {
                                $count++
                                $currentOperationString = "Getting Stats for $($FolderID.Identity) from Server $Server."
                                Write-Progress -Activity 'Retrieving Public Folder Stats for Selected Public Folders' -CurrentOperation $currentOperationString -PercentComplete $($count/$RecordCount*100) -Status "Retrieving Stats for folder replica instance $count of $RecordCount"
                                WriteLog -Message $currentOperationString -EntryType Notification -Verbose
                                #Error Action Silently Continue because some servers may not have a replica and we don't care about that error in this context
                                $thestats = Get-PublicFolderStatistics -Identity $FolderID.EntryID -Server $Server -ErrorAction SilentlyContinue 
                                if ($thestats) {$thestats | Select-Object -ExcludeProperty ServerName -Property $customProperties}
                                else {
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
                        $thestats = Get-PublicFolderStatistics -Server $Server -ResultSize Unlimited -ErrorAction SilentlyContinue 
                        if ($thestats) {$thestats | Select-Object -ExcludeProperty ServerName -Property $customProperties}
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
                                            'ServerName' = $_.ServerName
                                            'DatabaseName' = $_.DatabaseName
                                            'TotalItemSize' = $_.TotalItemSize
                                            'ItemCount' = $_.ItemCount
                                            'SizeInBytes' = $_.SizeInBytes
                                            'LastModificationTime' = $_.LastModificationTime
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
                    $replCheck = $true
                    foreach($dataRecord in $resultItem.Data) {
                        if ($resultItem.ItemCount -eq 0 -or $resultItem.ItemCount -eq $null)
                        {
                            $progress = 100
                        } else {
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
                $html = Get-HTMLReport -ReportObject -ResultMatrix -PublicFolderMailboxServer
            }#if to generate HTML output if required/requested
            #endregion GenerateHTMLOutput
            #region GenerateOutputFormats
            if ('files' -in $outputmethods -or 'email' -in $outputmethods) #files output to FileFolderpath requested
            {
                $outputfiles = @(
                    if ('csv' -in $outputformats) {
                        #Create the additional summary output object(s) for CSV
                        $PubliFolderEnvironmentSummary = [pscustomobject]@{
                            ReportTimeStamp = $ReportObject.TimeStamp
                            IncludedPublicFolderServersAndDatabases = $ReportObject.IncludedPublicFolderServersAndDatabases
                            IncludedPublicFoldersCount = $ReportObject.IncludedPublicFoldersCount
                            TotalSizeOfIncludedPublicFoldersInBytes = $ReportObject.TotalSizeOfIncludedPublicFoldersInBytes
                            TotalItemCountFromIncludedPublicFolders = $ReportObject.TotalItemCountFromIncludedPublicFolders
                            IncludedContainerOrEmptyPublicFoldersCount = $ReportObject.IncludedContainerOrEmptyPublicFoldersCount
                            IncludedReplicationIncompletePublicFolders = $ReportObject.IncludedReplicationIncompletePublicFolders
                        }
                        $LargestPublicFolders = $ReportObject.LargestPublicFolders | Select-Object FolderPath,TotalItemSize,ItemCount
                        #create the csv files
                        Export-Data -ExportFolderPath $FileFolderPath -DataToExportTitle PublicFolderEnvironmentSummary -DataToExport $PubliFolderEnvironmentSummary -DataType csv -ReturnExportFilePath
                        Export-Data -ExportFolderPath $FileFolderPath -DataToExportTitle LargestPublicFolders -DataToExport $LargestPublicFolders -DataType csv -ReturnExportFilePath
                        Export-Data -ExportFolderPath $FileFolderPath -DataToExportTitle PublicFoldersWithIncompleteReplication -DataToExport $ReportObject.PublicFoldersWithIncompleteReplication -DataType csv -ReturnExportFilePath
                        Export-Data -ExportFolderPath $FileFolderPath -DataToExportTitle ReplicationReportByServerPercentage -DataToExport $ReportObject.ReplicationReportByServerPercentage -DataType csv -ReturnExportFilePath
                    }
                    if ('html' -in $outputformats)
                    {
                        $HTMLFilePath = $FileFolderPath + $(Get-TimeStamp) + 'PublicFolderEnvironmentAndReplicationReport.html'
                        $html | Out-File -FilePath $HTMLFilePath 
                        $HTMLFilePath
                    }
                )
            }#if files or email in outputmethods
            #endregion GenerateOutputFormats
            #region SendMail
            if ('email' -in $outputmethods)
            {
                if ([string]::IsNullOrEmpty($Subject)) {
                    $Subject = 'Public Folder Environment and Replication Status Report'
                }
                $SendMailMessageParams = @{
                    Subject = $Subject
                    Attachments = $outputfiles
                    To = $to
                    From = $from
                    Body = if ($HTMLBody) {$html} else {"Public Folder Environment and Replication Status Report Attached."}
                    SMTPServer = $SmtpServer
                }
                Send-MailMessage @SendMailMessageParams
            }#if email in outputmethods
            #endregion SendMail
        }#end
    }
#end function Get-PublicFolderReplicationReport
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