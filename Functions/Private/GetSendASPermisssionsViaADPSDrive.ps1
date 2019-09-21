Function GetSendASPermisssionsViaADPSDrive
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
        [parameter(Mandatory)]
        [string]$ADPSDriveName
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
        [bool]$ExchangeOrganizationIsInExchangeOnline = $false
        ,
        $HRPropertySet #Property set for recipient object inclusion in object lookup hashtables
    )
    GetCallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -Name VerbosePreference

    #use the AD Drive Provided
    Push-Location
    $ADPSDrivePath = $ADPSDriveName + ':\'
    Set-Location -Path $ADPSDrivePath -ErrorAction Stop

    #Well-known GUID for Send As Permissions, see function Get-SendASRightGUID
    $SendASRight = [GUID]'ab721a54-1e2f-11d0-9819-00aa0040529b'

    $saRawPermissions = @(
        Try
        {
            $RawACEs = @((Get-ACL -Path $TargetMailPublicFolder.DistinguishedName -ErrorAction Stop).Access)
            $SendASACEs = $RawACEs | Where-Object -FilterScript { (($_.ObjectType -eq $SendASRight) -or ($_.ActiveDirectoryRights -eq 'GenericAll')) -and ($_.AccessControlType -eq 'Allow') }
            $SendASNotSelf = $SendASACEs | Where-Object -FilterScript { $_.IdentityReference.tostring() -ne "NT AUTHORITY\SELF" }
            $SendAsNotSelf | Select-Object -Property identityreference, IsInherited
            # Where-Object -FilterScript {($_.identityreference.ToString().split('\')[0]) -notin $ExcludedTrusteeDomains} #not doing this part yet
            # Where-Object -FilterScript {$_.identityreference.tostring() -notin $ExcludedTrustees} #we do this below now
        }
        Catch
        {
            $myerror = $_
            WriteLog -Message $myerror.tostring() -ErrorLog -Verbose -EntryType Failed
        }
    )
    #WriteLog -message "Found $($saRawPermissions.Count) SendAS Permisisons"
    Pop-Location

    if ($dropInheritedPermissions -eq $true)
    {
        $saRawPermissions = @($saRawPermissions | Where-Object -FilterScript { $_.IsInherited -eq $false })
        #WriteLog -message "Found $($saRawPermissions.count) non-inherited SendAS Permissions"
    }

    #Lookup Trustee Recipients and export permission if found
    foreach ($sa in $saRawPermissions)
    {
        $trusteeRecipient = GetTrusteeObject -TrusteeIdentity $sa.IdentityReference.tostring() -HRPropertySet $HRPropertySet -ObjectGUIDHash $ObjectGUIDHash -DomainPrincipalHash $DomainPrincipalHash -SIDHistoryHash $SIDHistoryRecipientHash -ExchangeSession $ExchangeSession -ExchangeOrganizationIsInExchangeOnline $ExchangeOrganizationIsInExchangeOnline -UnfoundIdentitiesHash $UnFoundIdentitiesHash
        switch ($null -eq $trusteeRecipient)
        {
            $true
            {
                $npeoParams = @{
                    TargetPublicFolder         = $TargetPublicFolder
                    TargetMailPublicFolder     = $TargetMailPublicFolder
                    TrusteeIdentity            = $sa.IdentityReference.tostring()
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
                        TrusteeIdentity            = $sa.IdentityReference.tostring()
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
