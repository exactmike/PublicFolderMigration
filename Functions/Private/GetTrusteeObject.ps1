Function GetTrusteeObject
{

    [CmdletBinding()]
    param
    (
        [parameter(Mandatory)]
        [AllowNull()]
        [string]$TrusteeIdentity
        ,
        [string[]]$HRPropertySet
        ,
        [hashtable]$ObjectGUIDHash
        ,
        [hashtable]$DomainPrincipalHash
        ,
        [hashtable]$SIDHistoryHash
        ,
        [hashtable]$UnfoundIdentitiesHash
        ,
        [System.Management.Automation.Runspaces.PSSession]$ExchangeSession
    )
    $trusteeObject = $(
        $AddToLookup = $null
        Write-Verbose -Message "Getting Object for TrusteeIdentity $TrusteeIdentity"
        switch ($TrusteeIdentity)
        {
            { $null -eq $TrusteeIdentity }
            {
                $null
                Write-Verbose -Message 'Trustee Identity is NULL'
                break
            }
            { $UnfoundIdentitiesHash.ContainsKey($_) }
            {
                $null
                Write-Verbose -Message 'Found Trustee in UnfoundIdentitiesHash'
                break
            }
            { $ObjectGUIDHash.ContainsKey($_) }
            {
                $ObjectGUIDHash.$($_)
                Write-Verbose -Message 'Found Trustee in ObjectGUIDHash'
                break
            }
            { $DomainPrincipalHash.ContainsKey($_) }
            {
                $DomainPrincipalHash.$($_)
                Write-Verbose -Message 'Found Trustee in DomainPrincipalHash'
                break
            }
            { $SIDHistoryHash.ContainsKey($_) }
            {
                $SIDHistoryHash.$($_)
                Write-Verbose -Message 'Found Trustee in SIDHistoryHash'
                break
            }
            Default
            {
                if ($Script:ExchangeOrganizationType -eq 'ExchangeOnline' -and $TrusteeIdentity -like '*\*')
                {
                    Write-Verbose -Message 'In Exchange Online and Trustee Identity is Domain Principal Format'
                    $null
                }
                else
                {
                    Write-Verbose -Message 'Performing new Trustee Recipient Lookup Attempts'
                    $splat = @{
                        Identity    = $TrusteeIdentity
                        ErrorAction = 'SilentlyContinue'
                    }
                    Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-Recipient @using:splat } -ErrorAction SilentlyContinue -OutVariable AddToLookup
                    if ($null -eq $AddToLookup -or $AddToLookup.count -eq 0)
                    {
                        Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-User @using:splat } -ErrorAction SilentlyContinue -OutVariable AddToLookup
                    }
                    if ($null -eq $AddToLookup -or $AddToLookup.count -eq 0)
                    {
                        Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-Group @using:splat } -ErrorAction SilentlyContinue -OutVariable AddToLookup
                    }
                }
            }
        }
    )
    #if we found a 'new' object add it to the lookup hashtables
    if ($null -ne $AddToLookup -and $AddToLookup.count -eq 1)
    {
        Write-Verbose -Message "Found Trustee $TrusteeIdentity via new lookup"
        $AddToLookup | Select-Object -Property $HRPropertySet | ForEach-Object -Process { $ObjectGUIDHash.$($_.ExchangeGuid.Guid) = $_ } -ErrorAction SilentlyContinue
        Write-Verbose -Message "ObjectGUIDHash Count is $($ObjectGUIDHash.count)"
        $AddToLookup | Select-Object -Property $HRPropertySet | ForEach-Object -Process { $ObjectGUIDHash.$($_.Guid.Guid) = $_ } -ErrorAction SilentlyContinue
        if ($TrusteeIdentity -like '*\*' -or $TrusteeIdentity -like '*@*')
        {
            $AddToLookup | Select-Object -Property $HRPropertySet | ForEach-Object -Process { $DomainPrincipalHash.$($TrusteeIdentity) = $_ } -ErrorAction SilentlyContinue
            Write-Verbose -Message "DomainPrincipalHash Count is $($DomainPrincipalHash.count)"
        }
    }
    #if we found nothing, add the Identity to the UnfoundIdentitiesHash
    if ($null -eq $trusteeObject -and $null -ne $TrusteeIdentity)
    {
        if (-not [string]::IsNullOrEmpty($TrusteeIdentity))
        {
            if (-not $UnfoundIdentitiesHash.ContainsKey($TrusteeIdentity))
            {
                $UnfoundIdentitiesHash.$TrusteeIdentity = $null
            }
        }
    }
    if ($null -ne $trusteeObject -and $trusteeObject.Count -ge 2)
    {
        Write-Verbose -Message "Trustee Identity $TrusteeIdentity is Ambiguous"
        #TrusteeIdentity is ambiguous.  Need to implement and AmbiguousIdentitiesHash for testing/reporting
        $trusteeObject = $null
    }
    $trusteeObject

}
