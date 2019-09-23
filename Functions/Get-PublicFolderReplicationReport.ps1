    Function Get-PublicFolderReplicationReport
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
        [string[]]$PublicFolderMailboxServer = @(),
        [string[]]$PublicFolderPath = @(),
        [switch]$Recurse,
        [switch]$IncludeSystemPublicFolders,
        [parameter()]
        [validateset('RawReplicationData', 'ReportObject')]
        [string]$PipelineData,
        [string]$FileFolderPath,
        [string[]]$To,
        [string]$From,
        [string]$SmtpServer,
        [string]$Subject,
        [switch]$HTMLBody,
        [parameter()]
        [validateset('html', 'csv')]
        [string[]]$outputformats,
        [parameter()]
        [validateset('email', 'files')]
        [string[]]$outputmethods,
        [int]$LargestPublicFolderReportCount = 10
    )
    Begin
    {
        Function GetTimeStamp
        {
            [string]$Stamp = Get-Date -Format yyyyMMdd-HHmmss
            #$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) #check if this is faster to use than Get-Date
            $Stamp
        }
        #region ValidateParameters
        if ('email' -in $outputmethods)
        {
            [array]$newTo = @()
            foreach ($recipient in $To)
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
        }#if email in outputmethods
        if ('files' -in $outputmethods -or 'email' -in $outputmethods)
        {
            if (-not (Test-Path -Path $FileFolderPath))
            {
                Write-Error "$FileFolderPath failed validation."
                Return
            }
            if ($FileFolderPath -notlike '*\')
            {
                $FileFolderPath = $FileFolderPath + '\'
            }
        }
        #if the user specified public folder mailbox servers, validate them:
        if ($PublicFolderMailboxServer.Count -ge 1)
        {
            foreach ($Server in $PublicFolderMailboxServer)
            {
                $VerifyPFDatabase = @(Get-PublicFolderDatabase -server $Server)
                if ($VerifyPFDatabase.Count -ne 1)
                {
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
        Write-Information -Message "Public Folder Mailbox Servers Included: $PublicFolderMailboxServerNames" -Tags Notification
        #Build Server/Database Hash Tables for later reporting activities
        $PublicFolderMailboxServerDatabases = @{ }
        $PublicFolderDatabaseMailboxServers = @{ }
        foreach ($server in $PublicFolderMailboxServer)
        {
            $PublicFolderDatabase = Get-PublicFolderDatabase -Server $Server -includePreExchange2010
            $PublicFolderMailboxServerDatabases.$Server = $PublicFolderDatabase | Select-Object -ExpandProperty Name
            $PublicFolderDatabaseMailboxServers.$($PublicFolderDatabase.Name) = $Server
        }
        #endregion BuildServerAndDatabaseLists
        #region BuildPublicFolderList
        #Set up the parameters for Get-PublicFolder
        $GetPublicFolderParams = @{ }
        if ($Recurse)
        {
            $GetPublicFolderParams.Recurse = $true
            $GetPublicFolderParams.ResultSize = 'Unlimited'
        }
        $FolderIDs = @(
            #if the user specified specific public folder paths, get those
            if ($PublicFolderPath.Count -ge 2)
            {
                $publicFolderPathString = $PublicFolderPath -join ', '
                Write-Information -Message "Retrieving Public Folders in the following Path(s): $publicFolderPathString" -Tags Notification -InformationAction Continue
                foreach ($Path in $PublicFolderPath)
                {
                    Get-PublicFolder $Path @GetPublicFolderParams | Select-Object -property @{n = 'EntryID'; e = { $_.EntryID.tostring() } }, @{n = 'Identity'; e = { $_.Identity.tostring() } }, Name, Replicas
                }
            }
            #otherwise, get all default public folders
            else
            {
                Write-Information -message 'Retrieving All Default (Non-System) Public Folders from IPM_SUBTREE' -Tags Notification -InformationAction Continue
                Get-PublicFolder -Recurse -ResultSize Unlimited | Select-Object -property @{n = 'EntryID'; e = { $_.EntryID.tostring() } }, @{n = 'Identity'; e = { $_.Identity.tostring() } }, Name, Replicas
                if ($IncludeSystemPublicFolders)
                {
                    Write-Information -Message 'Retrieving All System Public Folders from NON_IPM_SUBTREE' -Tags Notification -InformationAction Continue
                    Get-PublicFolder \Non_IPM_SUBTREE -Recurse -ResultSize Unlimited | Select-Object -property @{n = 'EntryID'; e = { $_.EntryID.tostring() } }, @{n = 'Identity'; e = { $_.Identity.tostring() } }, Name, Replicas
                }
            }
        )
        #filter any duplicates if the user specified public folder paths
        Write-Information -Message 'Sorting and De-duplicating retrieved Public Folders.' -Tags Notification -InformationAction Continue
        if ($PublicFolderPath.Count -ge 2) { $FolderIDs = @($FolderIDs | Select-Object -Unique -Property *) }
        #sort folders by path
        $FolderIDs = @($FolderIDs | Sort-Object Identity)
        $publicFoldersRetrievedCount = $FolderIDs.Count
        Write-Information -Message "Count of Public Folders Retrieved: $publicFoldersRetrievedCount" -Tags Notification -InformationAction Continue
        #endregion BuildPublicFolderList
        #region GetPublicFolderStats
        $publicFolderStatsFromSelectedServers =
        @(
            # if the user specified public folder path then only retrieve stats for the specified folders.
            # This can be significantly faster than retrieving stats for all public folders
            if ($PublicFolderPath.Count -ge 2)
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
                            @{n = 'ServerName'; e = { $Server } }
                            #this is necessary b/c powershell remoting makes the attributes deserialized and the value in bytes is not available directly.  Code below should work in EMS locally and in remote powershell sessions
                            @{n = 'SizeInBytes'; e = { $_.TotalItemSize.ToString().split(('(', ')'))[1].replace(',', '').replace(' bytes', '') -as [long] } }
                        )
                        $NOstatsProperties =
                        @{
                            'ServerName'           = $Server
                            'SizeInBytes'          = $null
                            'Progress'             = 0
                            'ItemCount'            = 0
                            'TotalItemSize'        = $null
                            'DatabaseName'         = $PublicFolderMailboxServerDatabases.$Server
                            'LastModificationTime' = $null
                        }
                        if ($FolderID.Replicas -contains $PublicFolderMailboxServerDatabases.$Server)
                        {
                            $count++
                            $currentOperationString = "Getting Stats for $($FolderID.Identity) from Server $Server."
                            Write-Progress -Activity 'Retrieving Public Folder Stats for Selected Public Folders' -CurrentOperation $currentOperationString -PercentComplete $($count / $RecordCount * 100) -Status "Retrieving Stats for folder replica instance $count of $RecordCount"
                            Write-Information -Message $currentOperationString -Tags Notification -InformationAction Continue
                            #Error Action Silently Continue because some servers may not have a replica and we don't care about that error in this context
                            $thestats = Get-PublicFolderStatistics -Identity $FolderID.EntryID -Server $Server -ErrorAction SilentlyContinue
                            if ($thestats) { $thestats | Select-Object -ExcludeProperty ServerName -Property $customProperties }
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
                foreach ($Server in $PublicFolderMailboxServer)
                {
                    $customProperties = @(
                        '*'
                        @{n = 'ServerName'; e = { $Server } }
                        @{n = 'SizeInBytes'; e = { $_.TotalItemSize.ToString().split(('(', ')'))[1].replace(',', '').replace(' bytes', '') -as [long] } }
                    )
                    Write-Verbose "Retrieving Stats for all Public Folders from $Server"
                    Write-Progress -Activity 'Retrieving Public Folder Stats' -CurrentOperation $Server -PercentComplete $($count / $RecordCount * 100) -Status "Retrieving Stats for Server $count of $RecordCount"
                    #avoid NULL output by testing for results while still suppressing errors with SilentlyContinue
                    $thestats = Get-PublicFolderStatistics -Server $Server -ResultSize Unlimited -ErrorAction SilentlyContinue
                    if ($thestats) { $thestats | Select-Object -ExcludeProperty ServerName -Property $customProperties }
                }
            }
        )
        #check for condition where there are no public folders and/or no public folder replicas on the specified servers
        if ($publicFolderStatsFromSelectedServers.Count -eq 0)
        {
            $message = 'There are no public folder replicas hosted on the specified servers.'
            Write-Information -Message $message -Tags Failure -InformationAction Continue
            Write-Error $message
            return
        }
        else
        {
            Write-Information -Message "Count of Stats objects returned: $($publicFolderStatsFromSelectedServers.count)" -Tags Notification -InformationAction Continue
        }
        #endregion GetPublicFolderStats
        #region BuildStatsLookupHash
        #create the hash table
        $publicFolderStatsLookup = @{ }
        #Populate the hashtable - one key/value pair per EntryID plus Server
        foreach ($Stats in ($publicFolderStatsFromSelectedServers | where-object -FilterScript { $_.EntryID -ne $NULL }))
        {
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
            foreach ($Folder in $FolderIDs)
            {
                $count++
                $currentOperationString = "Processing Report for Folder $($folder.EntryID) with name $($Folder.Identity)"
                Write-Progress -Activity 'Building Data Matrix of Public Folder Stats for output and reporting.' -Status 'Compiling Data' -CurrentOperation $currentOperationString -PercentComplete ($count / $RecordCount * 100)
                #Write-Information -Message $currentOperationString -Tags Notification -InformationAction Continue
                $resultItem = @{
                    EntryID            = $Folder.EntryID
                    FolderPath         = $Folder.Identity
                    Name               = $folder.name
                    ConfiguredReplicas = $($folder.replicas -join ',')
                    Data               = @(
                        #Get all the stats entries for this folder from each server using the EntryID + Server Key lookup
                        foreach ($Server in $PublicFolderMailboxServer)
                        {
                            $publicFolderStatsLookup.$($Folder.EntryID + '_' + $Server) | Where-Object -FilterScript { $_ } |
                            ForEach-Object {
                                New-Object PSObject -Property @{
                                    'ServerName'           = $_.ServerName
                                    'DatabaseName'         = $_.DatabaseName
                                    'TotalItemSize'        = $_.TotalItemSize
                                    'ItemCount'            = $_.ItemCount
                                    'SizeInBytes'          = $_.SizeInBytes
                                    'LastModificationTime' = $_.LastModificationTime
                                }
                            }
                        }
                    )
                }
                #Get Max Total Item Size in Bytes across Replicas
                $resultItem.TotalBytes = $resultItem.Data | Measure-Object -Property SizeInBytes -Maximum | Select-Object -ExpandProperty Maximum
                #Get Max Total Item Size human friendly based on max Bytes
                $resultItem.TotalItemSize = $resultItem.Data | Where-Object -FilterScript { $_.SizeInBytes -eq $resultItem.TotalBytes } | Select-Object -First 1 -ExpandProperty TotalItemSize
                #Get Max Item Count
                $resultItem.ItemCount = $resultItem.Data | Measure-Object -Property ItemCount -Maximum | Select-Object -ExpandProperty Maximum
                $replCheck = $true
                foreach ($dataRecord in $resultItem.Data)
                {
                    if ($resultItem.ItemCount -eq 0 -or $resultItem.ItemCount -eq $null)
                    {
                        $progress = 100
                    }
                    else
                    {
                        try
                        {
                            $ErrorActionPreference = 'Stop'
                            $progress = ([Math]::Round($dataRecord.ItemCount / ($resultItem.ItemCount) * 100, 0))
                            $ErrorActionPreference = 'Continue'
                        }
                        catch
                        {
                            $progress = $null
                            Write-Information -Message "Server: $($dataRecord.Server), Database: $($dataRecord.Databasename), ItemCount: $($dataRecord.ItemCount), TotalItemCount: $($resultItem.ItemCount)" -Tags Failure -InformationAction Continue
                            Write-Information -Message $_.tostring() -Tags Failure -InformationAction Continue
                            $ErrorActionPreference = 'Continue'
                        }
                    }
                    if ($progress -lt 100)
                    {
                        $replCheck = $false
                    }
                    $dataRecord | Add-Member -MemberType NoteProperty -Name 'Progress' -Value $progress
                }
                $resultItem.ReplicationCompleteOnIncludedServers = $replCheck
                #output result object
                New-Object PSObject -Property $resultItem
            }#Foreach
            Write-Progress -Activity 'Building Data Matrix of Public Folder Stats for output and reporting.' -Status 'Compiling Data' -CurrentOperation $currentOperationString -Completed
        )#$ResultMatrix
        #endregion BuildResultMatrix
        #Build the Report Object
        [pscustomobject]$ReportObject = @{
            #region BuildReportObject
            TimeStamp                                  = Get-Date -Format yyyyMMdd-HHmm
            IncludedPublicFolderServersAndDatabases    = $($(foreach ($server in $PublicFolderMailboxServer) { "$Server ($($PublicFolderMailboxServerDatabases.$server))" }) -join ',')
            IncludedPublicFoldersCount                 = $ResultMatrix.Count
            TotalSizeOfIncludedPublicFoldersInBytes    = $ResultMatrix | Measure-Object -Property TotalBytes -Sum | Select-Object -ExpandProperty Sum
            TotalItemCountFromIncludedPublicFolders    = $ResultMatrix | Measure-Object -Property ItemCount -Sum | Select-Object -ExpandProperty Sum
            IncludedContainerOrEmptyPublicFoldersCount = @($ResultMatrix | Where-Object -FilterScript { $_.ItemCount -eq 0 }).Count
            IncludedReplicationIncompletePublicFolders = @($ResultMatrix | Where-Object -FilterScript { $_.ReplicationCompleteOnIncludedServers -eq $false }).Count
            LargestPublicFolders                       = @($ResultMatrix | Sort-Object TotalBytes -Descending | Select-Object -First $LargestPublicFolderReportCount)
            PublicFoldersWithIncompleteReplication     = @(
                Foreach ($result in ($ResultMatrix | Where-Object -FilterScript { $_.ReplicationCompleteOnIncludedServers -eq $false }))
                {
                    [pscustomobject]@{
                        FolderPath                 = $Result.FolderPath
                        ItemCount                  = $Result.ItemCount
                        TotalItemSize              = $Result.TotalItemSize
                        ConfiguredReplicaDatabases = $result.ConfiguredReplicas
                        ConfiguredReplicaServers   =
                        $(
                            $databases = $result.ConfiguredReplicas.split(',')
                            $servers = $databases | foreach { $PublicFolderDatabaseMailboxServers.$_ }
                            $Servers -join ','
                        )
                        CompleteServers            =
                        $(
                            $CompleteServers = $result.Data | Where-Object { $_.Progress -eq 100 } | Select-Object -ExpandProperty ServerName
                            $CompleteServers -join ','
                        )
                        CompleteDatabases          =
                        $(
                            $CompleteDatabases = $result.Data | Where-Object { $_.Progress -eq 100 } | Select-Object -ExpandProperty ServerName
                            $CompleteDatabases -join ','
                        )
                        IncompleteServers          =
                        $(
                            $IncompleteServers = $result.Data | Where-Object { $_.Progress -lt 100 } | Select-Object -ExpandProperty ServerName
                            $IncompleteServers -join ','
                        )
                        IncompleteDatabases        =
                        $(
                            $IncompleteDatabases = $result.Data | Where-Object { $_.Progress -lt 100 } | Select-Object -ExpandProperty DatabaseName
                            $IncompleteDatabases -join ','
                        )
                    }#pscustomobject
                }#Foreach
            )
            ReplicationReportByServerPercentage        = @(
                Foreach ($result in $ResultMatrix)
                {
                    $RRObject = [pscustomobject]@{
                        FolderPath        = $result.FolderPath
                        HighestItemCount  = $result.ItemCount
                        HighestBytesCount = $result.totalBytes
                    }#pscustomobject
                    Foreach ($Server in $PublicFolderMailboxServer)
                    {
                        $ResultItem = $result.Data | Where-Object -FilterScript { $_.ServerName -eq $Server }
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
        $ReportObject.AverageSizeOfIncludedPublicFolders = [Math]::Round($ReportObject.TotalSizeOfIncludedPublicFoldersInBytes / $ReportObject.NonContainerOrEmptyPublicFoldersCount, 0)
        $ReportObject.AverageItemCountFromIncludedPublicFolders = [Math]::Round($ReportObject.TotalItemCountFromIncludedPublicFolders / $ReportObject.NonContainerOrEmptyPublicFoldersCount, 0)
        #endregion BuildReportObject
        #region PipelineDataOutput
        if (-not [string]::IsNullOrWhiteSpace($PipelineData))
        {
            switch ($PipelineData)
            {
                'RawReplicationData'
                { $ResultMatrix }
                'ReportObject'
                { $ReportObject }
            }
            #$ReportObject
        }#if $passthrough - output the report data as objects
        #endregion PipelineDataOutput
        #region GenerateHTMLOutput
        if (('html' -in $outputformats) -or $HTMLBody)
        {
            $html = @"
<html>
<style>
body
{
font-family:Arial,sans-serif;
font-size:8pt;
}
table
{
border-collapse:collapse;
font-size:8pt;
font-family:Arial,sans-serif;
border-collapse:collapse;
min-width:400px;
}
table,th, td
{
border: 1px solid black;
}
th
{
text-align:center;
font-size:18;
font-weight:bold;
}
</style>
<body>
<font size="1" face="Arial,sans-serif">
<h1 align="center">Exchange Public Folder Replication Report</h1>
<h4 align="center">Generated $([DateTime]::Now)</h3>

</font><h2>Overall Summary</h2>
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="2">Public Folder Environment Summary</th></tr>
<tr><td>Included Public Folder Servers</td><td>$($ReportObject.IncludedPublicFolderServersAndDatabases)</td></tr>
<tr><td>Count of Included Public Folders</td><td>$($ReportObject.IncludedPublicFoldersCount)</td></tr>
<tr><td>Count of Included Container or Empty Public Folders (0 Item Count)</td><td>$($ReportObject.IncludedContainerOrEmptyPublicFoldersCount)</td></tr>
<tr><td>Count of Included Public Folders with Incomplete Replication on Included Servers</td><td>$($ReportObject.IncludedReplicationIncompletePublicFolders)</td></tr>
<tr><td>Count of Total Items in Included Public Folders</td><td>$($ReportObject.TotalItemCountFromIncludedPublicFolders)</td></tr>
<tr><td>Total Size of Included Public Folder Items (Bytes)</td><td>$($ReportObject.TotalSizeOfIncludedPublicFoldersInBytes)</td></tr>
<tr><td>Average Size of Included Public Folders (Non Empty/Container)</td><td>$($ReportObject.AverageSizeOfIncludedPublicFolders)</td></tr>
<tr><td>Average Item Count of Included Public Folders (Non Empty/Container)</td><td>$($ReportObject.AverageItemCountFromIncludedPublicFolders)</td></tr>
</table>
<br />
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="3">Largest Public Folders by Size</th></tr>
<tr style="background-color:#E9E9E9;font-weight:bold"><td>Folder Path</td><td>Size</td><td>Item Count</td></tr>
$(
if (-not $ReportObject.LargestPublicFolders.Count -gt 0)
{
    "<tr><td colspan='3'>No Largest Public Folders Report is Included in this report.</td></tr>"
} else {
    foreach($sizeResult in $reportObject.LargestPublicFolders)
    {
        "<tr><td>$($sizeResult.FolderPath)</td><td>$($sizeResult.TotalItemSize)</td><td>$($sizeResult.ItemCount)</td></tr>`r`n"
    }
}
)
</table>
</font><h2>Public Folder Replication Results</h2>
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="5">Folders with Incomplete Replication on Included Servers</th></tr>
<tr style="background-color:#E9E9E9;font-weight:bold"><td>Folder Path</td><td>Item Count</td><td>Size</td><td>Servers with Replicas Configured</td><td>Servers with Replication Incomplete</td></tr>
$(
if (-not $ReportObject.PublicFoldersWithIncompleteReplication.Count -gt 0)
{
    "<tr><td colspan='4'>There are no public folders with incomplete replication.</td></tr>"
} else {
    foreach($IncompleteFolder in $ReportObject.PublicFoldersWithIncompleteReplication)
    {
        "<tr><td>$($IncompleteFolder.FolderPath)</td><td>$($IncompleteFolder.ItemCount)</td><td>$($IncompleteFolder.TotalItemSize)</td><td>$($IncompleteFolder.ConfiguredReplicaServers)</td><td>$($IncompleteFolder.IncompleteServers)</td></tr>`r`n"
    }
}
)
</table>
<br />
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="$($PublicFolderMailboxServer.Count + 1)">Public Folder Replication Information</th></tr>
<tr style="background-color:#E9E9E9;font-weight:bold"><td>Folder Path</td>
$(
foreach($rServer in $PublicFolderMailboxServer)
{
    "<td>$($rServer)</td>"
}
)
</tr>
$(
if (-not $ResultMatrix.Count -gt 0)
{
    "<tr><td colspan='$($PublicFolderMailboxServer.Count + 1)'>There are no public folders in this report.</td></tr>"
}
foreach($rItem in $ResultMatrix)
{
    "<tr><td>$($rItem.FolderPath)</td>"
    foreach($rServer in $PublicFolderMailboxServer)
    {
        $(
        $rDataItem = $rItem.Data | Where-Object { $_.ServerName -eq $rServer }
        if ($rDataItem -eq $null)
        {
            '<td>N/A</td>'
        } else {
            if ($rDataItem.Progress -ne 100)
            {
                $color = '#FC2222'
            } else {
                $color = '#A9FFB5'
            }
            "<td style='background-color:$($color)'><div title='$($rDataItem.TotalItemSize) of $($rItem.TotalItemSize) and $($rDataItem.ItemCount) of $($rItem.ItemCount) items.'>$($rDataItem.Progress)%</div></td>"
        }
        )
    }
    '</tr>'
}
)
</table>
</body>
</html>
"@
        }#if to generate HTML output if required/requested
        #endregion GenerateHTMLOutput
        #region GenerateOutputFormats
        if ('files' -in $outputmethods -or 'email' -in $outputmethods) #files output to FileFolderpath requested
        {
            $outputfiles = @(
                if ('csv' -in $outputformats)
                {
                    #Create the additional summary output object(s) for CSV
                    $PubliFolderEnvironmentSummary = [pscustomobject]@{
                        ReportTimeStamp                            = $ReportObject.TimeStamp
                        IncludedPublicFolderServersAndDatabases    = $ReportObject.IncludedPublicFolderServersAndDatabases
                        IncludedPublicFoldersCount                 = $ReportObject.IncludedPublicFoldersCount
                        TotalSizeOfIncludedPublicFoldersInBytes    = $ReportObject.TotalSizeOfIncludedPublicFoldersInBytes
                        TotalItemCountFromIncludedPublicFolders    = $ReportObject.TotalItemCountFromIncludedPublicFolders
                        IncludedContainerOrEmptyPublicFoldersCount = $ReportObject.IncludedContainerOrEmptyPublicFoldersCount
                        IncludedReplicationIncompletePublicFolders = $ReportObject.IncludedReplicationIncompletePublicFolders
                    }
                    $LargestPublicFolders = $ReportObject.LargestPublicFolders | Select-Object FolderPath, TotalItemSize, ItemCount
                    #create the csv files
                    Export-Data -ExportFolderPath $FileFolderPath -DataToExportTitle PublicFolderEnvironmentSummary -DataToExport $PubliFolderEnvironmentSummary -DataType csv -ReturnExportFilePath
                    Export-Data -ExportFolderPath $FileFolderPath -DataToExportTitle LargestPublicFolders -DataToExport $LargestPublicFolders -DataType csv -ReturnExportFilePath
                    Export-Data -ExportFolderPath $FileFolderPath -DataToExportTitle PublicFoldersWithIncompleteReplication -DataToExport $ReportObject.PublicFoldersWithIncompleteReplication -DataType csv -ReturnExportFilePath
                    Export-Data -ExportFolderPath $FileFolderPath -DataToExportTitle ReplicationReportByServerPercentage -DataToExport $ReportObject.ReplicationReportByServerPercentage -DataType csv -ReturnExportFilePath
                }
                if ('html' -in $outputformats)
                {
                    $HTMLFilePath = $FileFolderPath + $(GetTimeStamp) + 'PublicFolderEnvironmentAndReplicationReport.html'
                    $html | Out-File -FilePath $HTMLFilePath
                    $HTMLFilePath
                }
            )
        }#if files or email in outputmethods
        #endregion GenerateOutputFormats
        #region SendMail
        if ('email' -in $outputmethods)
        {
            if ([string]::IsNullOrEmpty($Subject))
            {
                $Subject = 'Public Folder Environment and Replication Status Report'
            }
            $SendMailMessageParams = @{
                Subject     = $Subject
                Attachments = $outputfiles
                To          = $to
                From        = $from
                Body        = if ($HTMLBody) { $html } else { "Public Folder Environment and Replication Status Report Attached." }
                SMTPServer  = $SmtpServer
            }
            Send-MailMessage @SendMailMessageParams
        }#if email in outputmethods
        #endregion SendMail
    }#end

    }
