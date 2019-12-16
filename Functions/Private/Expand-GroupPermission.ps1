Function Expand-GroupPermission
{

    [CmdletBinding()]
    [OutputType([PSObject[]], [System.Array])]
    param
    (
        [psobject[]]$Permission
        ,
        $TargetPublicFolder
        ,
        [parameter()]
        [AllowNull()]
        $TargetMailPublicFolder
        ,
        [hashtable]$ObjectGUIDHash
        ,
        [hashtable]$SIDHistoryHash
        ,
        $excludedTrusteeGUIDHash
        ,
        [hashtable]$UnfoundIdentitiesHash
        ,
        $HRPropertySet
        ,
        [System.Management.Automation.Runspaces.PSSession]$ExchangeSession
        ,
        [System.Management.Automation.Runspaces.PSSession]$ADPSSession
        ,
        $dropExpandedParentGroupPermissions
        ,
        [switch]$UseExchangeCommandsInsteadOfADOrLDAP
    )
    GetCallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -Name VerbosePreference
    if ($PSBoundParameters.ContainsKey('ADPSSession'))
    {
        Invoke-Command -Session $ADPSSession -ScriptBlock { Set-Location -path 'GC:\' -ErrorAction Stop } -ErrorAction Stop
    }
    $gPermissions = @($Permission | Where-Object -FilterScript { $_.TrusteeRecipientTypeDetails -like '*Group*' })
    $ngPermissions = @($Permission | Where-Object -FilterScript { $_.TrusteeRecipientTypeDetails -notlike '*Group*' -or $null -eq $_.TrusteeRecipientTypeDetails })
    if ($gPermissions.Count -ge 1)
    {
        $expandedPermissions = @(
            foreach ($gp in $gPermissions)
            {
                Write-Verbose -Message "Expanding Group $($gp.TrusteeObjectGUID)"
                #check if we have already expanded this group . . .
                switch ($script:ExpandedGroupsNonGroupMembershipHash.ContainsKey($gp.TrusteeObjectGUID))
                {
                    $true
                    {
                        #if so, get the terminal trustee objects from the expansion hashtable
                        $UserTrustees = $script:ExpandedGroupsNonGroupMembershipHash.$($gp.TrusteeObjectGUID)
                        Write-Verbose -Message "Previously Expanded Group $($gp.TrusteeObjectGUID) Members Count: $($userTrustees.count)"
                    }
                    $false
                    {
                        #if not, get the terminal trustee objects now
                        if ($UseExchangeCommandsInsteadOfADOrLDAP -eq $true)
                        {
                            $UserTrustees = @(GetGroupMemberExpandedViaExchange -Identity $gp.TrusteeObjectGUID -ExchangeSession $exchangeSession -hrPropertySet $HRPropertySet -ObjectGUIDHash $ObjectGUIDHash -DomainPrincipalHash $DomainPrincipalHash -SIDHistoryRecipientHash $SIDHistoryRecipientHash -UnFoundIdentitiesHash $UnfoundIdentitiesHash )
                        }
                        else
                        {
                            $UserTrustees = @(Get-GroupMemberExpandedViaADPSDrive -Identity $gp.TrusteeDistinguishedName -ExchangeSession $exchangeSession -hrPropertySet $HRPropertySet -ObjectGUIDHash $ObjectGUIDHash -DomainPrincipalHash $DomainPrincipalHash -SIDHistoryRecipientHash $SIDHistoryRecipientHash -UnfoundIdentitiesHash $UnfoundIdentitiesHash -ADPSSession $ADPSSession)
                        }
                        #and add them to the expansion hashtable
                        $script:ExpandedGroupsNonGroupMembershipHash.$($gp.TrusteeObjectGUID) = $UserTrustees
                        Write-Verbose -Message "Newly Expanded Group $($gp.TrusteeObjectGUID) Members Count: $($userTrustees.count)"
                    }
                }
                foreach ($u in $UserTrustees)
                {
                    $trusteeRecipient = $u
                    switch ($null -eq $trusteeRecipient)
                    {
                        $true
                        {
                            #no point in doing anything here
                        }#end $true
                        $false
                        {
                            if (-not $excludedTrusteeGUIDHash.ContainsKey($trusteeRecipient.guid.guid))
                            {
                                $npeoParams = @{
                                    TargetPublicFolder         = $TargetPublicFolder
                                    TargetMailPublicFolder     = $TargetMailPublicFolder
                                    TrusteeIdentity            = $trusteeRecipient.guid.guid
                                    TrusteeRecipientObject     = $trusteeRecipient
                                    TrusteeGroupObjectGUID     = $gp.TrusteeObjectGUID
                                    PermissionType             = $gp.PermissionType
                                    AccessRights               = $gp.AccessRights
                                    AssignmentType             = 'GroupMembership'
                                    SourceExchangeOrganization = $Script:ExchangeOrganization
                                    IsInherited                = $gp.IsInherited
                                    ParentPermissionIdentity   = $gp.PermissionIdentity
                                }
                                NewPermissionExportObject @npeoParams
                            }
                        }#end $false
                    }#end switch
                }#end foreach (user)
            }#end foreach (permission)
        )#expandedPermissions
        if ($expandedPermissions.Count -ge 1)
        {
            #remove any self permissions that came in through expansion
            $expandedPermissions = @($expandedPermissions | Where-Object -FilterScript { $_.TargetObjectGUID -ne $_.TrusteeObjectGUID })
        }
        if ($dropExpandedParentGroupPermissions)
        {
            @($ngPermissions; $expandedPermissions)
        }
        else
        {
            @($ngPermissions; $gPermissions; $expandedPermissions)
        }
    }
    else
    {
        $Permission
    }

}
