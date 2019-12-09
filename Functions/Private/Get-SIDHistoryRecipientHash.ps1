Function Get-SIDHistoryRecipientHash
{

    [cmdletbinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [parameter(Mandatory)]
        [Alias('ExchangeSession')]
        [System.Management.Automation.Runspaces.PSSession]$ExchangePSSession
        ,
        [parameter(Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]$ADPSSession
    )

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
                $SIDHistoryRecipientHash.$($sidhistorysid.value) = $sidhistoryuserrecipient
            }#End Foreach
        }#end If
    }#End Foreach
    $SIDHistoryRecipientHash
}
