Function GetClientPermission
{

    [cmdletbinding()]
    param
    (
        $TargetPublicFolder
        ,
        [parameter()]
        [AllowNull()]
        $TargetMailPublicFolder
        ,
        [System.Management.Automation.Runspaces.PSSession]$ExchangeSession
        ,
        [hashtable]$ObjectGUIDHash
        ,
        [hashtable]$excludedTrusteeGUIDHash
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
    $splat = @{Identity = $TargetPublicFolder.EntryID.tostring(); ErrorAction = 'Stop' }
    try
    {
        $RawClientPermissions = @(Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-PublicFolderClientPermission @using:splat } -ErrorAction Stop)
    }
    catch
    {
        $myerror = $_
        WriteLog -Message $myerror.tostring() -ErrorLog -Verbose -EntryType Failed
        $RawClientPermissions = @()
    }
    foreach ($rcp in $RawClientPermissions)
    {
        switch -Wildcard ($rcp.User)
        {
            'Default'
            { $trusteeRecipient = $null }
            'Anonymous'
            { $trusteeRecipient = $null }
            'NT User:*'
            {
                $TrusteeIdentity = $rcp.User.split(':')[1]
                $trusteeRecipient = GetTrusteeObject -TrusteeIdentity $TrusteeIdentity -HRPropertySet $HRPropertySet -ObjectGUIDHash $ObjectGUIDHash -DomainPrincipalHash $DomainPrincipalHash -SIDHistoryHash $SIDHistoryRecipientHash -ExchangeSession $ExchangeSession -UnfoundIdentitiesHash $UnFoundIdentitiesHash
            }
            Default
            { $trusteeRecipient = GetTrusteeObject -TrusteeIdentity $rcp.user -HRPropertySet $HRPropertySet -ObjectGUIDHash $ObjectGUIDHash -DomainPrincipalHash $DomainPrincipalHash -SIDHistoryHash $SIDHistoryRecipientHash -ExchangeSession $ExchangeSession -UnfoundIdentitiesHash $UnFoundIdentitiesHash }
        }
        switch ($null -eq $trusteeRecipient)
        {
            $true
            {
                $npeoParams = @{
                    TargetPublicFolder         = $TargetPublicFolder
                    TargetMailPublicFolder     = $TargetMailPublicFolder
                    TrusteeIdentity            = $rcp.User
                    TrusteeRecipientObject     = $null
                    PermissionType             = 'ClientPermission'
                    AccessRights               = $rcp.AccessRights -join '|'
                    AssignmentType             = 'Undetermined'
                    IsInherited                = $false
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
                        TrusteeIdentity            = $rcp.User
                        TrusteeRecipientObject     = $trusteeRecipient
                        PermissionType             = 'ClientPermission'
                        AccessRights               = $rcp.AccessRights -join '|'
                        AssignmentType             = switch -Wildcard ($trusteeRecipient.RecipientTypeDetails) { '*group*' { 'GroupMembership' } $null { 'Undetermined' } Default { 'Direct' } }
                        IsInherited                = $false
                        SourceExchangeOrganization = $ExchangeOrganization
                    }
                    NewPermissionExportObject @npeoParams
                }
            }#end $false
        }#end switch
    }#end foreach fa

}
