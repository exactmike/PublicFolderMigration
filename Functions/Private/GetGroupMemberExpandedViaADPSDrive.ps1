Function GetGroupMemberExpandedViaADPSDrive
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
        [string]$ADPSDriveName
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
    Push-Location
    $ADPSDrivePath = $ADPSDriveName + ':\'
    Set-Location -Path $ADPSDrivePath -ErrorAction Stop

    $TrusteeObjects = @(
        Try
        {
            Get-ADObject -ldapfilter $LDAPFilter -ErrorAction Stop
        }
        Catch
        {
            $myError = $_
            WriteLog -Message $myError.tostring() -ErrorLog -EntryType Failed -Verbose
        }
    )

    Pop-Location

    foreach ($to in $TrusteeObjects)
    {
        $TrusteeIdentity = $to.objectguid.guid
        $trusteeRecipient = GetTrusteeObject -TrusteeIdentity $TrusteeIdentity -HRPropertySet $HRPropertySet -ObjectGUIDHash $ObjectGUIDHash -DomainPrincipalHash $DomainPrincipalHash -SIDHistoryHash $SIDHistoryRecipientHash -ExchangeSession $ExchangeSession -UnfoundIdentitiesHash $UnFoundIdentitiesHash
        if ($null -ne $trusteeRecipient) { $trusteeRecipient }
    }

}
