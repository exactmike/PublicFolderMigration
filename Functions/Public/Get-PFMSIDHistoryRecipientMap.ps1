Function Get-PFMSIDHistoryRecipientMap
{

    [cmdletbinding(DefaultParameterSetName = 'UserInitiated')]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [parameter(Mandatory, ParameterSetName = 'ModuleInitiated')]
        [Alias('ExchangeSession')]
        [System.Management.Automation.Runspaces.PSSession]$ExchangePSSession
        ,
        [parameter(Mandatory, ParameterSetName = 'ModuleInitiated')]
        [System.Management.Automation.Runspaces.PSSession]$ADPSSession
        ,
        [Parameter(Mandatory, ParameterSetName = 'UserInitiated')]
        [ValidateScript( { TestIsWriteableDirectory -Path $_ })]
        $OutputFolderPath
        ,
        [parameter(ParameterSetName = 'UserInitiated')]
        [ValidateSet('Unicode', 'BigEndianUnicode', 'Ascii', 'Default', 'UTF8', 'UTF8NOBOM', 'UTF7', 'UTF32')]
        [string]$Encoding = 'UTF8'
        ,
        [parameter(ParameterSetName = 'UserInitiated')]
        [switch]$Passthru
    )

    switch ($PSCmdlet.ParameterSetName)
    {
        'UserInitiated'
        {
            $ExchangePSSession = $script:PSSession
            $ADPSSession = $script:ADPSSession
        }
        'ModuleInitiated'
        {

        }
    }
    Confirm-PFMExchangeConnection -PSSession $ExchangePSSession
    Confirm-PFMActiveDirectoryConnection -PSSession $ADPSSession
    $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
    $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetSIDHistoryRecipientMap.log')
    $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetSIDHistoryRecipientMap-ERRORS.log')
    #$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
    WriteLog -Message "Exchange Session is Running in Exchange Organzation $script:ExchangeOrganization" -EntryType Notification
    [ExportDataOutputFormat[]]$Outputformat = 'json', 'clixml'

    $ldapfilter = "(&(legacyExchangeDN=*)(sidhistory=*))"

    GetCallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -Name VerbosePreference

    #Region GetSIDHistoryUsers
    Try
    {
        $message = "Get AD Objects with Exchange Attributes and SIDHistory from AD Global Catalog"
        WriteLog -Message $message -EntryType Attempting
        Invoke-Command -Session $ADPSSession -ScriptBlock { Set-Location -Path 'GC:\' -ErrorAction Stop } -ErrorAction Stop
        $sidHistoryUsers = @(
            Invoke-Command -Session $ADPSSession -ScriptBlock {
                Get-ADObject -ldapfilter $using:ldapfilter -Properties sidhistory, legacyExchangeDN -ErrorAction Stop
            }
        )
        WriteLog -Message $message -EntryType Succeeded
    }
    Catch
    {
        $myError = $_
        WriteLog -Message $message -EntryType Failed -ErrorLog
        WriteLog -Message $myError.tostring() -ErrorLog
        throw("Failed: $Message")
    }
    WriteLog -Message "Got $($sidHistoryUsers.count) AD Objects with Exchange Attributes and SIDHistory from AD Global Catalog" -EntryType Notification
    #EndRegion GetSIDHistoryObjects

    $sidhistoryusercounter = 0
    $SIDHistoryRecipientHash = @{ }
    Foreach ($shu in $sidhistoryusers)
    {
        $sidhistoryusercounter++
        $message = 'Generating hash of SIDHistory SIDs and Recipient objects...'
        $ProgressInterval = [int]($($sidhistoryusers.Count) * .01)
        if ($($sidhistoryusercounter) % $ProgressInterval -eq 0)
        {
            Write-Progress -Activity $message -status "Items processed: $($sidhistoryusercounter) of $($sidhistoryusers.Count)" -percentComplete (($sidhistoryusercounter / $($sidhistoryusers.Count)) * 100)
        }
        $splat = @{Identity = $shu.ObjectGuid.guid; ErrorAction = 'SilentlyContinue' } #is this a good assumption?
        $sidhistoryuserrecipient = $Null
        $sidhistoryuserrecipient = Invoke-Command -Session $ExchangePSSession -ScriptBlock { Get-Recipient @using:splat } -ErrorAction SilentlyContinue
        If ($null -ne $sidhistoryuserrecipient)
        {
            Foreach ($sidhistorysid in $shu.sidhistory)
            {
                $SIDHistoryRecipientHash.$($sidhistorysid) = $sidhistoryuserrecipient
            }#End Foreach
        }#end If
    }#End Foreach

    $ResultCount = $SIDHistoryRecipientHash.count
    WriteLog -Message "Count of SIDHistory Recipients Retrieved: $ResultCount" -EntryType Notification -verbose
    switch ($PSCmdlet.ParameterSetName)
    {
        'UserInitiated'
        {
            $CreatedFilePath = @(
                foreach ($of in $Outputformat)
                {
                    Export-PFMData -ExportFolderPath $OutputFolderPath -DataToExportTitle 'SidHistoryRecipientHash' -ReturnExportFilePath -Encoding $Encoding -DataFormat $of -DataToExport $SIDHistoryRecipientHash
                }
            )
            WriteLog -Message "Output files created: $($CreatedFilePath -join '; ')" -entryType Notification -verbose
        }
        'ModuleInitiated'
        {
            $Passthru = $true
        }
    }

    if ($true -eq $Passthru)
    {
        $SIDHistoryRecipientHash
    }
}
