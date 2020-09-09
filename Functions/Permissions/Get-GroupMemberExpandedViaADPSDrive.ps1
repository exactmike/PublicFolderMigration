Function Get-GroupMemberExpandedViaADPS
{

    [CmdletBinding()]
    param
    (
        [string]$Identity #distinguishedName
        ,
        [parameter(Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]$ExchangeSession
        ,
        [parameter(Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]$ADPSSession
        ,
        $hrPropertySet
        ,
        $ObjectGUIDHash
        ,
        $DomainPrincipalHash
        ,
        $SIDHistoryRecipientHash
        ,
        [hashtable]$UnfoundIdentitiesHash
    )
    #enumerate groups: http://stackoverflow.com/questions/8055338/listing-users-in-ad-group-recursively-with-powershell-script-without-cmdlets/8055996#8055996
    $LDAPFilter = "(&(memberof:1.2.840.113556.1.4.1941:=$($Identity))(objectCategory=user))"

    GetCallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -Name VerbosePreference
    Invoke-Command -Session $ADPSSession -ScriptBlock { Set-Location -path 'GC:\' -ErrorAction Stop } -ErrorAction Stop

    $TrusteeObjects = @(
        Try
        {
            Invoke-Command -Session $ADPSSession -ScriptBlock {
                Get-ADObject -ldapfilter $Using:LDAPFilter -ErrorAction Stop
            } -ErrorAction Stop
        }
        Catch
        {
            $myError = $_
            WriteLog -Message $myError.tostring() -ErrorLog -EntryType Failed -Verbose
        }
    )

    foreach ($to in $TrusteeObjects)
    {
        $TrusteeIdentity = $to.objectguid.guid
        $trusteeRecipient = GetTrusteeObject -TrusteeIdentity $TrusteeIdentity -HRPropertySet $HRPropertySet -ObjectGUIDHash $ObjectGUIDHash -DomainPrincipalHash $DomainPrincipalHash -SIDHistoryHash $SIDHistoryRecipientHash -ExchangeSession $ExchangeSession -UnfoundIdentitiesHash $UnFoundIdentitiesHash
        if ($null -ne $trusteeRecipient) { $trusteeRecipient }
    }

}
