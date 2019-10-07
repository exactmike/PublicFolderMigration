Function Get-PFMPublicFolderTree
{
    <#
    .SYNOPSIS
    Gets Public Folder Information Objects for all or specified public folders.
    .DESCRIPTION
    Gets Public Folder Information Objects for all or specified public folders.  Returns a custom public folder information object which includes the following attributes:
    EntryID (A Unique Identifier), Name, Identity, MapiIdentity, ParentPath, HasSubFolders, ReplicasString (delimited with ;), Replicas, ReplicaCount, UseDatabaseReplicationSchedule, ReplicationScheduleString, ReplicationSchedule, PerUserReadStateEnabled, FolderType, MailEnabled, HiddenFromAddressListsEnabled, MaxItemSize, UseDatabaseQuotaDefaults, IssueWarningQuota, ProhibitPostQuota, UseDatabaseRetentionDefaults, RetainDeletedItemsFor, UseDatabaseAgeDefaults, AgeLimit, HasRules, HasModerator, IsValid
    .PARAMETER PublicFolderMailboxServer
    This parameter specifies the Exchange 2010 server from which to retrieve folder information to generate the Public Folder Information Objects.
    .PARAMETER PublicFolderPath
    This parameter specifies the Public Folder(s) to scan. If this is omitted, all public folders are scanned (except System Public Folders - see the IncludeSystemPublicFolders parameter). Include the leading '\'.
    .PARAMETER Recurse
    When used in conjunction with the FolderPath parameter, this will include all child Public Folders of the Folders listed in Folder Path.
    .PARAMETER Passthru
    Controls whether the public folder information objects are returned to the PowerShell pipeline for further processing.
    .PARAMETER OutputFolderPath
    Mandatory parameter for the already existing directory location where you want public folder replication and stats reports to be placed.  Operational log files will also go to this location.
    .PARAMETER OutputFormats
    Mandatory parameter used to specify whether you want csv, json, xml, clixml or any combination of these.
    .PARAMETER IncludeSystemPublicFolders
    This parameter specifies to include System Public Folders when scanning all public folders. If this is omitted, System Public Folders are omitted.
    .EXAMPLE
    Connect-PFMExchange -ExchangeOnPremisesServer USCLTEX10PF01.us.clt.contoso.com -credential $cred
    Get-PFMPublicFolderTree -OutputFolderPath c:\PFReports -OutputFormats csv,json,xml -PublicFolderMailboxServer USCLTEX10PF01

    Gets public folder tree data from USCLTEX10PF01.us.clt.contoso.com and exports it to csv, json, and xml formats in c:\PFReports
    #>
    [CmdletBinding(ConfirmImpact = 'None')]
    [OutputType([System.Object[]])]
    param
    (
        [parameter(Mandatory)]
        [string]$PublicFolderMailboxServer
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
        [ValidateSet('Unicode', 'BigEndianUnicode', 'Ascii', 'Default', 'UTF8', 'UTF8NOBOM', 'UTF7', 'UTF32')]
        [string]$Encoding = 'UTF8'
    )

    Confirm-PFMExchangeConnection -PSSession $Script:PSSession
    $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
    $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetPublicFolderTree.log')
    $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetPublicFolderTree-ERRORS.log')
    WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
    $ExchangeOrganization = Invoke-Command -Session $Script:PSSession -ScriptBlock { Get-OrganizationConfig | Select-Object -ExpandProperty Identity | Select-Object -ExpandProperty Name }
    WriteLog -Message "Exchange Session is Running in Exchange Organzation $ExchangeOrganization" -EntryType Notification
    #region ValidateParameters
    $VerifyPFDatabase = @(
        Invoke-Command -Session $script:PSSession -scriptblock {
            Get-PublicFolderDatabase -server $using:PublicFolderMailboxServer -ErrorAction SilentlyContinue
        }
    )
    if ($VerifyPFDatabase.Count -ne 1)
    {
        Write-Error "$PublicFolderMailboxServer does not host a public folder database."
        Return
    }
    #Using/Abusing? switch here.  Switch wants to unroll the array so using scriptblock options
    $publicFolderPathType = switch ($null) #types are Root, SingleNonRoot, MultipleWithRoot, MultipleNonRoot
    {
        { $PublicFolderPath.Count -eq 0 }
        { 'Root' }
        { $PublicFolderPath.Count -eq 1 -and $PublicFolderPath[0] -eq '\' }
        { 'Root' }
        { $PublicFolderPath.Count -eq 1 -and $PublicFolderPath[0] -ne '\' }
        { 'SingleNonRoot' }
        { $PublicFolderPath.Count -ge 2 -and $PublicFolderPath -contains '\' }
        { 'MultipleWithRoot' }
        { $PublicFolderPath.Count -ge 2 -and $PublicFolderPath -notcontains '\' }
        { 'MultipleNonRoot' }
        { $null -eq $PublicFolderPath }
        { 'Root' }
        Default
        { 'Root' }
    }
    writelog -Message "PublicFolder Path Type specified by user parameters: $PublicFolderPathType"  -EntryType Notification -verbose
    #endregion ValidateParameters
    #region BuildServerAndDatabaseLists
    WriteLog -Message "Validated Public Folder Mailbox Server To Query: $PublicFolderMailboxServer" -EntryType Notification -Verbose

    #setup property list to retrieve
    $PropertyList = @(
        @{n = 'EntryID'; e = { $_.EntryID.tostring() } }
        'Name'
        @{n = 'Identity'; e = { $_.Identity.tostring() } }
        @{n = 'MapiIdentity'; e = { $_.MapiIdentity.tostring() } }
        'ParentPath'
        'HasSubFolders'
        @{n = 'ReplicasString'; e = { $_.Replicas -join ';' } }
        'Replicas'
        @{n = 'ReplicaCount'; e = { $_.Replicas.count } }
        'UseDatabaseReplicationSchedule'
        @{n = 'ReplicationScheduleString'; e = { $_.ReplicationSchedule -join ';' } }
        'ReplicationSchedule'
        'PerUserReadStateEnabled'
        'FolderType'
        'MailEnabled'
        'HiddenFromAddressListsEnabled'
        'MaxItemSize'
        'UseDatabaseQuotaDefaults'
        'IssueWarningQuota'
        'ProhibitPostQuota'
        'UseDatabaseRetentionDefaults'
        'RetainDeletedItemsFor'
        'UseDatabaseAgeDefaults'
        'AgeLimit'
        'HasRules'
        'HasModerator'
        'IsValid'
    )

    $GetPublicFolderParams = @{ }
    if ($Recurse)
    {
        $GetPublicFolderParams.Recurse = $true
        $GetPublicFolderParams.ResultSize = 'Unlimited'
    }
    $Folders = @(
        switch ($publicFolderPathType)
        {
            { $_ -in @('SingleNonRoot') } #if the user specified specific public folder paths, get those
            {
                $publicFolderPathString = $PublicFolderPath -join ', '
                $path = $PublicFolderPath[0]
                WriteLog -Message "Retrieving Public Folders in the following Path: $publicFolderPathString" -EntryType Notification
                Invoke-Command -Session $script:PSSession -ScriptBlock {
                    Get-PublicFolder -Identity $Using:path  @using:GetPublicFolderParams
                } | Select-Object -property $PropertyList
            }
            { $_ -in @('MultipleNonRoot') } #if the user specified specific public folder paths, get those
            {
                $publicFolderPathString = $PublicFolderPath -join ', '
                foreach ($path in $PublicFolderPath)
                {
                    WriteLog -Message "Retrieving Public Folders in the following Path(s): $publicFolderPathString" -EntryType Notification
                    Invoke-Command -Session $script:PSSession -ScriptBlock {
                        Get-PublicFolder -Identity $using:path @using:GetPublicFolderParams
                    } | Select-Object -property $PropertyList
                }
            }
            { $_ -in @('Root', 'MultipleWithRoot') } #otherwise, get all default public folders
            {
                WriteLog -message 'Retrieving All Default (Non-System) Public Folders from IPM_SUBTREE' -EntryType Notification
                Invoke-Command -Session $script:PSSession -ScriptBlock {
                    Get-PublicFolder -Recurse -ResultSize Unlimited
                } | Select-Object -property $PropertyList
                if ($IncludeSystemPublicFolders)
                {
                    WriteLog -Message 'Retrieving All System Public Folders from NON_IPM_SUBTREE' -EntryType Notification
                    Invoke-Command -Session $script:PSSession -ScriptBlock {
                        Get-PublicFolder \Non_IPM_SUBTREE -Recurse -ResultSize Unlimited
                    } | Select-Object -property $PropertyList
                }
            }
        }
    )
    #filter any duplicates if the user specified public folder paths
    if ($publicFolderPathType -in @('MultipleNonRoot'))
    {
        WriteLog -Message 'Sorting and De-duplicating retrieved Public Folders.' -EntryType Notification -verbose
        $Folders = @($Folders | Sort-Object -Unique -Property EntryID)
        $Folders = @($Folders | Sort-Object -Unique -Property Identity)
    }
    #sort folders by path
    $publicFoldersRetrievedCount = $Folders.Count
    WriteLog -Message "Count of Public Folders Retrieved: $publicFoldersRetrievedCount" -EntryType Notification -verbose
    #endregion BuildPublicFolderList
    $CreatedFilePath = @(
        foreach ($of in $Outputformats)
        {
            Export-Data -ExportFolderPath $OutputFolderPath -DataToExportTitle 'PublicFolderTree' -ReturnExportFilePath -Encoding $Encoding -DataFormat $of -DataToExport $Folders
        }
    )
    WriteLog -Message "Output files created: $($CreatedFilePath -join '; ')" -entryType Notification -verbose
    if ($true -eq $Passthru)
    {
        $Folders
    }
}
