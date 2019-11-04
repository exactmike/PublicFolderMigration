Function GetSendASPermissionsViaExchange
{

    [cmdletbinding()]
    param
    (
        $TargetPublicFolder
        ,
        [parameter(Mandatory)]
        $TargetMailPublicFolder
        ,
        [System.Management.Automation.Runspaces.PSSession]$ExchangeSession
        ,
        [hashtable]$ObjectGUIDHash
        ,
        [hashtable]$excludedTrusteeGUIDHash
        ,
        [bool]$dropInheritedPermissions
        ,
        [hashtable]$DomainPrincipalHash
        ,
        [hashtable]$UnfoundIdentitiesHash
        ,
        $ExchangeOrganization
        ,
        $HRPropertySet #Property set for recipient object inclusion in object lookup hashtables
    )
    GetCallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -Name VerbosePreference
    switch ($Script:ExchangeOrganizationType)
    {
        'ExchangeOnline'
        {
            $command = 'Get-RecipientPermission'
            $splat = @{
                ErrorAction  = 'Stop'
                ResultSize   = 'Unlimited'
                Identity     = $TargetMailPublicFolder.guid.guid
                AccessRights = 'SendAs'
            }
            try
            {
                $saRawPermissions = @(Invoke-Command -Session $ExchangeSession -ScriptBlock { &($using:command) @using:splat } -ErrorAction Stop)
            }
            catch
            {
                $saRawPermissions = @()
                $myerror = $_
                WriteLog -Message $myerror.tostring() -ErrorLog -Verbose -EntryType Failed
            }
        }
        'ExchangeOnPremises'
        {
            $command = 'Get-ADPermission'
            $splat = @{
                ErrorAction = 'Stop'
                Identity    = $TargetMailPublicFolder.distinguishedname
            }
            #Get All AD Permissions
            try
            {
                $saRawPermissions = @(Invoke-Command -Session $ExchangeSession -ScriptBlock { &($using:command) @using:splat } -ErrorAction Stop)
            }
            catch
            {
                $saRawPermissions = @()
                $myerror = $_
                WriteLog -Message $myerror.tostring() -ErrorLog -Verbose -EntryType Failed
            }
            #Filter out just the Send-AS Permissions
            $saRawPermissions = @($saRawPermissions | Where-Object -FilterScript { $_.ExtendedRights -contains 'Send-As' })
        }
    }
    #Drop Inherited Permissions if Requested
    if ($dropInheritedPermissions)
    {
        $saRawPermissions = @($saRawPermissions | Where-Object -FilterScript { $_.IsInherited -eq $false })
    }
    $IdentityProperty = switch ($Script:ExchangeOrganizationType) { 'ExchangeOnline' { 'Trustee' } 'ExchangeOnPremises' { 'User' } }
    #Drop Self Permissions
    $saRawPermissions = @($saRawPermissions | Where-Object -FilterScript { $_.$IdentityProperty -ne 'NT AUTHORITY\SELF' })
    #Lookup Trustee Recipients and export permission if found
    foreach ($sa in $saRawPermissions)
    {
        $trusteeRecipient = GetTrusteeObject -TrusteeIdentity $sa.$IdentityProperty -HRPropertySet $HRPropertySet -ObjectGUIDHash $ObjectGUIDHash -DomainPrincipalHash $DomainPrincipalHash -SIDHistoryHash $SIDHistoryRecipientHash -ExchangeSession $ExchangeSession -UnfoundIdentitiesHash $UnFoundIdentitiesHash
        switch ($null -eq $trusteeRecipient)
        {
            $true
            {
                $npeoParams = @{
                    TargetPublicFolder         = $TargetPublicFolder
                    TargetMailPublicFolder     = $TargetMailPublicFolder
                    TrusteeIdentity            = $sa.$IdentityProperty
                    TrusteeRecipientObject     = $null
                    PermissionType             = 'SendAs'
                    AssignmentType             = 'Undetermined'
                    IsInherited                = $sa.IsInherited
                    SourceExchangeOrganization = $ExchangeOrganization
                }
                NewPermissionExportObject @npeoParams
            }#end $true
            $false
            {
                if (-not $excludedTrusteeGUIDHash.ContainsKey($trusteeRecipient.guid.guid))
                {
                    $npeoParams = @{
                        TargetPublicFolder         = $TargetPublicFolder
                        TargetMailPublicFolder     = $TargetMailPublicFolder
                        TrusteeIdentity            = $sa.$IdentityProperty
                        TrusteeRecipientObject     = $trusteeRecipient
                        PermissionType             = 'SendAs'
                        AssignmentType             = switch -Wildcard ($trusteeRecipient.RecipientTypeDetails) { $null { 'Undetermined' } '*group*' { 'GroupMembership' } Default { 'Direct' } }
                        IsInherited                = $sa.IsInherited
                        SourceExchangeOrganization = $ExchangeOrganization
                    }
                    NewPermissionExportObject @npeoParams
                }
            }#end $false
        }#end switch
    }#end foreach

}
