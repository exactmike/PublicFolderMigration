<#
    .SYNOPSIS
    Processes Empty Public Folder Information Objects for possible removal
    .DESCRIPTION
    Accepts one or more EntryIDs or other Public Folder Unique Identifiers and processes them for removal if they meet the required validations.
    .PARAMETER PublicFolderMailboxServer
    Use to  specifies the Exchange server from which to retrieve folder information to generate the Public Folder Information Objects.
    .PARAMETER Identity
    Use to specify the identity(ies) of the public folder(s) to be validated for and processed for removal.
    .PARAMETER Validations
    Use to specify the validations to run before processing a public folder for removal. NoSubfolders,NotMailEnabled,NoItems
    .PARAMETER DateValidations
    Use to specify a set of time validations to run before processing a public folder for removal.  Create a set of time validations using New-PFMTimeValidationSet and New-PFMTimeValidation.  This parameter accepts the name of a previously created set.
    .PARAMETER Passthru
    Controls whether the public folder validation/removal result objects are returned to the PowerShell pipeline for further processing.
    .PARAMETER OutputFolderPath
    Mandatory parameter for the already existing directory location where you want public folder replication and stats reports to be placed.  Operational log files will also go to this location.
    .PARAMETER OutputFormat
    Mandatory parameter used to specify whether you want csv, json, xml, clixml or any combination of these.
    .EXAMPLE
    Connect-PFMExchange -ExchangeOnPremisesServer USCLTEX10PF01.us.clt.contoso.com -credential $cred
    Invoke-PFMEmptyPublicFolderRemoval

    If public folders are on Exchange 2010, the ExchangeOnPremisesServer must be an Exchange 2010 server.
    Gets public folder tree data from USCLTEX10PF01.us.clt.contoso.com and exports it to csv, json, and xml formats in c:\PFReports
#>
Function Invoke-PFMRemovePublicFolder
{
    [cmdletbinding()]
    [OutputType([psobject])]
    param(
        [parameter()]
        $PublicFolderMailboxServer
        ,
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]]$Identity
        ,
        [parameter()]
        [FolderValidation[]]$Validations
        ,
        [parameter()]
        #[FolderActivityTimeValidation[]]
        [psobject[]]$TimeValidations
        ,
        [parameter(Mandatory)]
        [ValidateScript( { TestIsWriteableDirectory -path $_ })]
        [string]$OutputFolderPath
        ,
        [parameter()]
        [ValidateSet('Unicode', 'BigEndianUnicode', 'Ascii', 'Default', 'UTF8', 'UTF8NOBOM', 'UTF7', 'UTF32')]
        [string]$Encoding = 'UTF8'
    )
    begin
    {
        Confirm-PFMExchangeConnection -PSSession $Script:PSSession
        $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
        $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'InvokeRemovePublicFolder.log')
        $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'InvokeRemovePublicFolder-ERRORS.log')
        WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
        $ExchangeOrganization = Invoke-Command -Session $Script:PSSession -ScriptBlock { Get-OrganizationConfig | Select-Object -ExpandProperty Identity | Select-Object -ExpandProperty Name }
        WriteLog -Message "Exchange Session is Running in Exchange Organzation $ExchangeOrganization" -EntryType Notification
        $GetPublicFolderParams = @{
            Recurse     = $False
            #ResultSize  = 'Unlimited'
            ErrorAction = 'Stop'
        }
        if ($Validations -contains 'NoItems' -or $TimeValidations.Count -ge 1 )
        {
            $ServerDatabase = @(
                Invoke-Command -Session $script:PSSession -ScriptBlock {
                    Get-PublicFolderDatabase
                } | Select-Object -Property @{n = 'DatabaseName'; e = { $_.Name } }, @{n = 'ServerName'; e = { $_.Server } }, @{n = 'ServerFQDN'; e = { $_.RpcClientAccessServer } }
            )
            $DatabaseServerLookup = @{ }
            $ServerDatabase.foreach( { $DatabaseServerLookup.$($_.DatabaseName) = $_ })
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
                    WriteLog -message "Connected Parallel PSSession to $($s.ServerFQDN) for Stats operations" -entrytype Notification -verbose
                    $connectSessionSuccess.Add($s.ServerFQDN)
                }
                catch
                {
                    WriteLog -message "Unable to connect a remote Exchange Powershell session to $($s.ServerFQDN)" -entryType Failed -Verbose
                    $connectSessionFailure.Add($s.ServerFQDN)
                }
            }
            $ServerDatabaseToProcess, $ServerDatabaseRetry = $ServerDatabase.where( { $_.ServerFQDN -in $connectSessionSuccess }, 'Split')
            if ($connectSessionFailure.Count -ge 1)
            {
                WriteLog -message "Connect Session Failures: $($connectSessionFailure -join ',')" -entrytype Notification
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
        }
        $ValidationRecords = [System.Collections.ArrayList]::new()
    }
    process
    {
        :nexti foreach ($i in $Identity)
        {
            #setup the result object
            $result = [pscustomobject]@{
                InputIdentity       = $i
                FoundEntryID        = ''
                FoundIdentity       = ''
                FoldersFound        = $null
                ValidationResults   = [System.Collections.ArrayList]::new()
                ValidatedForRemoval = $false
                RemovalResult       = $null
                RemovalError        = ''
            }
            #region getfolder
            $folder = @(
                try
                {
                    Invoke-Command -ErrorAction Stop -Session $script:PSSession -ScriptBlock {
                        Get-PublicFolder -Identity $Using:i  @using:GetPublicFolderParams
                    } | Select-Object -property $Script:PublicFolderPropertyList
                }
                catch { }
            )
            $result.FoldersFound = $folder.Count
            if ($folder.count -ne 1 -or $null -eq $folder[0])
            {
                $result
                continue :nexti
            }
            else
            {
                $foundfolder = $folder[0]
                $result.FoundEntryID = $foundfolder.entryID
                $result.FoundIdentity = $foundfolder.Identity
            }
            #endregion getfolder
            if ($Validations -contains 'NoItems' -or $TimeValidations.Count -ge 1 )
            {
                $folderstats = @(
                    foreach ($r in $foundfolder.replicas)
                    {
                        $ServerSession = Get-PFMParallelPSSession -name $DatabaseServerLookup.$r.ServerFQDN
                        Confirm-PFMExchangeConnection -IsParallel -PSSession $ServerSession
                        $ServerSession = Get-PFMParallelPSSession -name $DatabaseServerLookup.$r.ServerFQDN
                        Invoke-Command -Session $ServerSession -ScriptBlock {
                            Get-PublicFolderStatistics -Identity $using:EntryID -Server $using:ServerName -ErrorAction SilentlyContinue
                        }
                    }
                )
            }
            #region validate
            foreach ($v in $Validations)
            {
                $vResult = [pscustomobject]@{ Name = $v; Result = $null }
                switch ($v)
                {
                    'NoSubFolders'
                    {
                        if ($false -eq $foundfolder.hasSubfolders)
                        {
                            $vResult.Result = $true
                        }
                        else
                        {
                            $vResult.Result = $false
                        }
                    }
                    'NotMailEnabled'
                    {
                        if ($false -eq $foundfolder.mailenabled)
                        {
                            $vResult.Result = $true
                        }
                        else
                        {
                            $vResult.Result = $false
                        }
                    }
                    'NoItems'
                    {
                        $max = 0
                        $folderstats.foreach( { if ($max -lt $_.itemcount) { $max = $_.itemcount } })
                        if ($max -eq 0)
                        {
                            $vResult.Result = $true
                        }
                        else
                        {
                            $vResult.Result = $false
                        }
                    }
                }
                $null = $result.ValidationResults.add($vResult)
            }
            if ($result.ValidationResults.Result -notcontains $false)
            {
                $result.ValidatedForRemoval = $true
            }
            #endregion validate
            #output to pipeline
            $result
            #output to Validation Records
            $ValidationRecords.add($(ConvertTo-Json -InputObject $result -Compress))
        }
    }
    end
    {
        $ValidationCount = $ValidationRecords.count
        $RecordFilePath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'PFValidationForRemovalRecords.log')
        WriteLog -Message "$validationCount Public Folders processed for Removal Validation" -entryType Notification -verbose
        WriteLog -Message "Public Folder Validation for Removal Records being sent to file $RecordFilePath" -entryType Notification -verbose
        $ValidationRecords | Out-File -FilePath $RecordFilePath -Encoding $Encoding
    }
}
