<#
    .SYNOPSIS
    Processes Public Folders for the specified validations and emits an overall result object
    .DESCRIPTION
    Accepts one or more EntryIDs or other Public Folder Unique Identifiers and processes the specified validations.
    .PARAMETER PublicFolderMailboxServer
    Use to  specifies the Exchange server from which to retrieve folder information to generate the Public Folder Information Objects.
    .PARAMETER Identity
    Use to specify the identity(ies) of the public folder(s) to be validated for and processed for removal.
    .PARAMETER Validations
    Use to specify the validations to run before processing a public folder for removal. NoSubfolders,NotMailEnabled,NoItems
    .PARAMETER DateValidations
    Use to specify a set of time validations to run before processing a public folder for removal.  Create a set of time validations using New-PFMTimeValidationSet and New-PFMTimeValidation.  This parameter accepts the name of a previously created set.
    .PARAMETER OutputFolderPath
    Mandatory parameter for the already existing directory location where you want public folder replication and stats reports to be placed.  Operational log files will also go to this location.
    .EXAMPLE
    Connect-PFMExchange -ExchangeOnPremisesServer USCLTEX10PF01.us.clt.contoso.com -credential $cred
    Invoke-PFMEmptyPublicFolderRemoval

    If public folders are on Exchange 2010, the ExchangeOnPremisesServer must be an Exchange 2010 server.
    Gets public folder tree data from USCLTEX10PF01.us.clt.contoso.com and exports it to csv, json, and xml formats in c:\PFReports
#>
Function Invoke-PFMValidatePublicFolder
{
    [cmdletbinding()]
    [OutputType([psobject])]
    param(
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]]$Identity
        ,
        [parameter()]
        [FolderValidation[]]$Validations
        ,
        #[parameter()]
        #[FolderActivityTimeValidation[]]
        #[psobject[]]$TimeValidations
        #,
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
        $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'InvokeValidatePublicFolder.log')
        $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'InvokeValidatePublicFolder-ERRORS.log')
        WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
        WriteLog -Message "Exchange Session is Running in Exchange Organzation $script:ExchangeOrganization" -EntryType Notification
        $GetPublicFolderParams = @{
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
        }
        $ValidationRecords = [System.Collections.ArrayList]::new()
    }
    process
    {
        :nexti foreach ($i in $Identity)
        {
            #setup the result object
            $result = [pscustomobject]@{
                PSTypeName         = 'PublicFolderValidation'
                InputIdentity      = $i
                FoundEntryID       = ''
                FoundIdentity      = ''
                FoldersFound       = $null
                ValidationResults  = [System.Collections.ArrayList]::new()
                Validated          = $false
                ValidatedTimeStamp = $null
                ActionName         = ''
                ActionResult       = $null
                ActionError        = ''
                ActionTimeStamp    = $null
            }
            #region getfolder
            Confirm-PFMExchangeConnection -PSSession $Script:PSSession
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
                Confirm-PFMExchangeConnection -PSSession $Script:PSSession
                $EntryID = $foundfolder.EntryID
                $folderstats = @(
                    foreach ($r in $foundfolder.replicas)
                    {
                        $ServerFQDN = $DatabaseServerLookup.$r.ServerFQDN
                        Invoke-Command -Session $Script:PSSession -ScriptBlock {
                            Get-PublicFolderStatistics -Identity $using:EntryID -Server $using:ServerFQDN -ErrorAction SilentlyContinue
                        }
                    }
                )
            }
            #region validate
            foreach ($v in $Validations)
            {
                $vResult = [pscustomobject]@{ Name = $v; Result = $null ; EvaluatedAttribute = ''; EvaluatedValue = $null }
                switch ($v)
                {
                    'NoSubFolders'
                    {
                        $vResult.EvaluatedAttribute = 'hasSubfolders'
                        $vResult.EvaluatedValue = $foundfolder.hasSubfolders
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
                        $vResult.EvaluatedAttribute = 'MailEnabled'
                        $vResult.EvaluatedValue = $foundfolder.MailEnabled
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
                        $vResult.EvaluatedAttribute = 'ItemCount'
                        $max = 0
                        $folderstats.foreach( { if ($max -lt $_.itemcount) { $max = $_.itemcount } })
                        $vResult.EvaluatedValue = $max
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
                $result.Validated = $true
                $result.ValidatedTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
            }
            #endregion validate

            #output to pipeline
            $result
            #output to Validation Records
            $null = $ValidationRecords.add($(ConvertTo-Json -InputObject $result -Compress))
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
