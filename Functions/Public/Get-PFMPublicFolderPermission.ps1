Function Get-PFMPublicFolderPermission
{
    <#
    .SYNOPSIS
        Gets public folder permission objects for client, sendas, and sendonbehalf types.
    .DESCRIPTION
        Gets public folder permission objects for client, sendas, and sendonbehalf types for all public folders, or for a selected subset by tree or entryid.
        Can optionally exclude public folders, permission holders (trustees), expand group permissions, and drop inherited permissions.

    .EXAMPLE
        Get-PFMPublicFolderPermission
        Gets permissions for all public folders. uses the default settings for inclusion of various permisison types
    .INPUTS
        Inputs (if any)
    .OUTPUTS
        Output (if any)
    .NOTES
        General notes
    #>
    [cmdletbinding(DefaultParameterSetName = 'AllPublicFolders', ConfirmImpact = 'none')]
    [OutputType([System.Object[]])]
    param
    (
        [parameter(ParameterSetName = 'Scoped', Mandatory)]
        [string[]]$PublicFolderPath
        ,
        [parameter(ParameterSetName = 'Scoped')]
        [switch]$Recurse
        ,
        [parameter(ParameterSetName = 'EntryID', Mandatory)]
        [string[]]$PublicFolderEntryID
        ,
        [parameter(ParameterSetName = 'InfoObject', Mandatory)]
        [psobject[]]$PublicFolderInfoObject
        ,
        [Parameter(Mandatory)]
        [ValidateScript( { TestIsWriteableDirectory -Path $_ })]
        $OutputFolderPath
        ,
        #Public Folder identities to exclude from permissions gathering (use folder name, full path, or EntryID).  EntryID is preferred as it is guaranteed to be unique.
        [parameter()]
        [string[]]$ExcludedIdentities
        ,
        [parameter()]#These will be resolved to trustee objects and permisisons with these trustees will be omitted from output
        [string[]]$ExcludedTrusteeIdentities
        ,
        [Parameter()]#include public folder client permissions
        [bool]$IncludeClientPermission = $true
        ,
        [Parameter()]#include sendas permissions
        [bool]$IncludeSendAs = $true
        ,
        [Parameter()]#include sendonbehalf permissions
        [bool]$IncludeSendOnBehalf = $true
        ,
        #Expand group permissions to individual trustees if possible
        [bool]$ExpandGroups = $true
        ,
        #Drop the original group permission if ExpandGroups is True
        [bool]$DropExpandedParentGroupPermissions = $false
        ,
        #Drop inherited permissions
        [bool]$DropInheritedPermissions = $false
        ,
        #lookup SIDHistory for matching SIDs in permissions to an actual trustee
        [switch]$IncludeSIDHistory
        ,
        #exclude output where the resulting permission is 'none'
        [switch]$ExcludeNonePermissionOutput
        ,
        [switch]$ExcludedIdentitiesAreEntryID
    )#End Param
    Begin
    {
        Confirm-PFMExchangeConnection -PSSession $Script:PSSession
        If ($script:ExchangeOrganizationType -eq 'ExchangeOnPremises')
        {
            Confirm-PFMActiveDirectoryConnection -PSSession $script:ADPSSession
        }
        $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
        $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetPublicFolderPermission.log')
        $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetPublicFolderPermission-ERRORS.log')
        #$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
        WriteLog -Message "Exchange Session is Running in Exchange Organzation $script:ExchangeOrganization" -EntryType Notification
        switch ($script:ExchangeOrganizationType)
        {
            'ExchangeOnline'
            {
                if ($True -eq $IncludeSIDHistory)
                {
                    throw ('You cannot include SidHistory when your Exchange Organization is in Exchange Online.')
                }
            }
            'ExchangeOnPremises'
            {
                If ($true -eq $IncludeSidHistory -or $true -eq $IncludeSendAs -or $true -eq $ExpandGroups)
                {
                    Confirm-PFMActiveDirectoryConnection -pssession $script:ADPSSession
                }
            }
        }
        #Configure properties to retain in memory / hashtables for retrieved public folders and Recipients
        $PFPropertySet = @('EntryID', 'Identity', 'Name', 'ParentPath', 'FolderType', 'Has*', 'HiddenFromAddressListsEnabled', '*Quota', 'MailEnabled', 'Replicas', 'ReplicationSchedule', 'RetainDeletedItemsFor', 'Use*')
        $HRPropertySet = @('*name*', '*addr*', 'RecipientType*', '*Id', 'Identity', 'GrantSendOnBehalfTo')
        $ExportedExchangePublicFolderPermissionsFile = Join-Path -Path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'ExchangePublicFolderPermissions.csv')
        $ResumeIndex = 0
        [uint32]$Script:PermissionIdentity = 0
        #create a property set for storing of recipient data during processing.  We don't need all attributes in memory/storage.
        #Region GetExcludedRecipients

        if ($PSBoundParameters.ContainsKey('ExcludedIdentities'))
        {
            Confirm-PFMExchangeConnection -PSSession $script:PSSession
            switch ($true -eq $ExcludedIdentitiesAreEntryID)
            {
                $false
                {
                    try
                    {
                        $message = "Get public folder object(s) from Exchange Organization $ExchangeOrganization for the $($ExcludedIdentities.Count) ExcludedIdentities provided."
                        WriteLog -Message $message -EntryType Attempting -verbose
                        $excludedPublicFolders = @(
                            $ExcludedIdentities | ForEach-Object {
                                $splat = @{
                                    Identity    = $_
                                    ErrorAction = 'Stop'
                                }
                                Invoke-Command -Session $Script:PSSession -ScriptBlock { Get-PublicFolder @Using:splat | Select-Object -Property $using:PFPropertySet } -ErrorAction 'Stop'
                            }
                        )
                        WriteLog -Message $message -EntryType Succeeded -verbose
                    }
                    Catch
                    {
                        $myError = $_
                        WriteLog -Message $message -EntryType Failed -ErrorLog
                        WriteLog -Message $myError.tostring() -ErrorLog
                        throw("Failed: $Message")
                    }
                    WriteLog -Message "Found $($excludedPublicFolders.count) Public Folders for Exclusion from processing" -EntryType Notification
                    $excludedPublicFoldersEntryIDHash = @{ }
                    $excludedPublicFolders.foreach( { $excludedPublicFoldersEntryIDHash.$($_.EntryID.tostring()) = $_ })
                }
                $true
                {
                    WriteLog -Message "Processing $($ExcludedIdentities.count) EntryIDs for Exclusion from processing" -EntryType Notification
                    $excludedPublicFoldersEntryIDHash = @{ }
                    $ExcludedIdentities.foreach( { $excludedPublicFoldersEntryIDHash.$($_) = $_ } )
                }
            }
        }
        else
        {
            $excludedPublicFoldersEntryIDHash = @{ }
        }
        #EndRegion GetExcludedRecipients

        #Region GetExcludedTrustees
        if ($PSBoundParameters.ContainsKey('ExcludedTrusteeIdentities'))
        {
            Confirm-PFMExchangeConnection -PSSession $script:PSSession
            try
            {
                $message = "Get recipent object(s) from Exchange Organization $ExchangeOrganization for the $($ExcludedTrusteeIdentities.Count) ExcludedTrusteeIdentities provided."
                WriteLog -Message $message -EntryType Attempting -verbose
                $excludedTrusteeRecipients = @(
                    $ExcludedTrusteeIdentities | ForEach-Object {
                        $splat = @{
                            Identity    = $_
                            ErrorAction = 'Stop'
                        }
                        Invoke-Command -Session $Script:PSSession -ScriptBlock { Get-Recipient @Using:splat | Select-Object -Property $using:HRPropertySet } -ErrorAction 'Stop'
                    }
                )
                WriteLog -Message $message -EntryType Succeeded -verbose
            }
            Catch
            {
                $myError = $_
                WriteLog -Message $message -EntryType Failed -ErrorLog
                WriteLog -Message $myError.tostring() -ErrorLog
                throw("Failed: $Message")
            }
            WriteLog -Message "Got $($excludedTrusteeRecipients.count) Excluded Trustee Objects" -EntryType Notification -verbose
            $excludedTrusteeGUIDHash = @{ }
            $excludedTrusteeRecipients.foreach( { $excludedTrusteeGUIDHash.$($_.GUID.ToString()) = $_ })
        }
        else
        {
            $excludedTrusteeGUIDHash = @{ }
        }
        #EndRegion GetExcludedTrustees

        #Region GetInScopePublicFolders
        Try
        {
            switch ($PSCmdlet.ParameterSetName)
            {
                'Scoped'
                {
                    Write-Information -MessageData "Not Implemented, rewriting to use Get-PFMPublicFolderTree" -InformationAction Continue
                    Write-Warning -message "Not Implemented, rewriting to use Get-PFMPublicFolderTree" -WarningAction Stop
                }#end Scoped
                'AllPublicFolders'
                {
                    Write-Information -MessageData "Not Implemented, rewriting to use Get-PFMPublicFolderTree" -InformationAction Continue
                    Write-Warning -message "Not Implemented, rewriting to use Get-PFMPublicFolderTree" -WarningAction Stop
                }#end AllMailboxes
                'EntryID'
                {
                    Write-Information -MessageData "Not Implemented, rewriting to use Get-PFMPublicFolderTree" -InformationAction Continue
                    Write-Warning -message "Not Implemented, rewriting to use Get-PFMPublicFolderTree" -WarningAction Stop
                }
                'InfoObject'
                {
                    $InScopeFolders = $PublicFolderInfoObject
                }
            }#end Switch
        }#end try
        Catch
        {
            $myError = $_
            WriteLog -Message $message -EntryType Failed -ErrorLog
            WriteLog -Message $myError.tostring() -ErrorLog
            throw("Failed: $Message")
        }
        $InScopeFolderCount = $InScopeFolders.count
        WriteLog -Message "Got $InScopeFolderCount In Scope Folder Objects" -EntryType Notification
        #EndRegion GetInScopePublicFolders

        #Region GetInScopeMailPublicFolders
        If ($true -eq $IncludeSendAs -or $true -eq $IncludeSendOnBehalf)
        {
            Confirm-PFMExchangeConnection -PSSession $script:PSSessio
            $message = 'Get Mail Enabled Public Folders To support retrieval of SendAS and/or SendOnBehalf Permissions and for additional output information for ClientPermissions.'
            WriteLog -message $message -entryType Attempting -verbose
            $PossibleMailEnabledPF = $InScopeFolders.where( { ($_.MailEnabled -is [bool] -and $_.MailEnabled -eq $true) -or $_.MailEnabled -eq 'TRUE' })
            $InScopeMailPublicFolders = @(GetMailPublicFolderPerUserPublicFolder -ExchangeSession $script:PSSession -PublicFolder $PossibleMailEnabledPF -ErrorAction Stop)
            WriteLog -message $message -entryType Succeeded -verbose
            WriteLog -Message "Got $($InScopeMailPublicFolders.count) In Scope Mail Public Folder Objects" -EntryType Notification -verbose
            $InScopeMailPublicFoldersHash = @{ }
            $InScopeMailPublicFolders.foreach( { $InScopeMailPublicFoldersHash.$($_.EntryID.ToString()) = $_ })
        }
        else
        {
            $InScopeMailPublicFoldersHash = @{ }
        }
        #EndRegion GetInScopeMailPublicFolders

        #Region GetSIDHistoryData
        if ($IncludeSIDHistory -eq $true)
        {
            Confirm-PFMActiveDirectoryConnection -PSSession $script:ADPSSession
            Confirm-PFMExchangeConnection -PSSession $script:PSSession
            $SIDHistoryRecipientHash = Get-SIDHistoryRecipientHash -ExchangePSSession $Script:PSSession -ADPSSession $Script:ADPSSession -ErrorAction Stop
        }
        else
        {
            $SIDHistoryRecipientHash = @{ }
        }
        #EndRegion GetSIDHistoryData

        #Region BuildLookupHashTables
        #these have to be populated as we go
        WriteLog -Message "Building Recipient Lookup HashTables" -EntryType Notification
        $DomainPrincipalHash = @{ }
        $UnfoundIdentitiesHash = @{ }
        $ObjectGUIDHash = @{ }
        if ($expandGroups -eq $true)
        {
            $script:ExpandedGroupsNonGroupMembershipHash = @{ }
        }
        #EndRegion BuildLookupHashtables
    }
    End
    {
        #Set Up to Loop through Public Folders
        $message = "First Permission Identity will be $($Script:PermissionIdentity)"
        WriteLog -message $message -EntryType Notification
        $ISRCounter = $ResumeIndex
        $ExportedPermissions = @(
            :nextISR for
            (
                $i = $ResumeIndex
                $i -le $InScopeFolderCount - 1
                $(if ($Recovering) { $i = $ResumeIndex } else { $i++ })
                #$ISR in $InScopeFolders[$ResumeIndex..$()]
            )
            {
                Confirm-PFMExchangeConnection -PSSession $script:PSSession
                if ($true -eq $IncludeSendAs -or $true -eq $IncludeSendOnBehalf)
                {
                    Confirm-PFMActiveDirectoryConnection -PSSession $script:ADPSSession
                }
                $Recovering = $false
                $ISRCounter++
                $ISR = $InScopeFolders[$i]
                $ID = $ISR.EntryID.tostring()
                if ($excludedPublicFoldersEntryIDHash.ContainsKey($ID))
                {
                    #WriteLog -Message "Excluding Excluded Folder with EntryID $ID"
                    continue nextISR
                }
                if ($InScopeMailPublicFoldersHash.ContainsKey($ID))
                {
                    $ISRR = $InScopeMailPublicFoldersHash.$ID
                }
                else
                {
                    $ISRR = $null
                }
                $message = "Collect permissions for $($ID)"
                Write-Progress -Activity $message -status "Items processed: $($ISRCounter) of $($InScopeFolderCount)" -percentComplete (($ISRCounter / $InScopeFolderCount) * 100)
                Try
                {
                    Confirm-PFMExchangeConnection -PSSession $Script:PSSession
                    WriteLog -Message $message -EntryType Attempting
                    $PermissionExportObjects = @(
                        If ($IncludeSendOnBehalf -and $InScopeMailPublicFoldersHash.ContainsKey($ID))
                        {
                            #WriteLog -Message "Getting SendOnBehalf Permissions for Target $ID" -entryType Notification
                            GetSendOnBehalfPermission -TargetPublicFolder $ISR -TargetMailPublicFolder $ISRR -ObjectGUIDHash $ObjectGUIDHash -ExchangeSession $Script:PSSession -ExcludedTrusteeGUIDHash $excludedTrusteeGUIDHash -ExchangeOrganization $ExchangeOrganization -HRPropertySet $HRPropertySet -DomainPrincipalHash $DomainPrincipalHash -UnfoundIdentitiesHash $UnfoundIdentitiesHash
                        }
                        If ($IncludeClientPermission)
                        {
                            #WriteLog -Message "Getting Client Permissions for Target $ID" -entryType Notification
                            GetClientPermission -TargetPublicFolder $ISR -TargetMailPublicFolder $ISRR -ObjectGUIDHash $ObjectGUIDHash -ExchangeSession $Script:PSSession -excludedTrusteeGUIDHash $excludedTrusteeGUIDHash -ExchangeOrganization $ExchangeOrganization -DomainPrincipalHash $DomainPrincipalHash -HRPropertySet $HRPropertySet -UnfoundIdentitiesHash $UnfoundIdentitiesHash
                        }
                        If ($IncludeSendAs -and $InScopeMailPublicFoldersHash.ContainsKey($ID))
                        {
                            #WriteLog -Message "Getting SendAS Permissions for Target $ID" -entryType Notification
                            switch ($script:ExchangeOrganizationType)
                            {
                                'ExchangeOnline'
                                {
                                    #WriteLog -Message "Getting SendAS Permissions for Target $ID Via Exchange Commands" -entryType Notification
                                    GetSendASPermissionsViaExchange -TargetPublicFolder $ISRR -TargetMailPublicFolder $ISRR -ExchangeSession $Script:PSSession -ObjectGUIDHash $ObjectGUIDHash -excludedTrusteeGUIDHash $ -dropInheritedPermissions $dropInheritedPermissions -DomainPrincipalHash $DomainPrincipalHash -ExchangeOrganization $ExchangeOrganization -HRPropertySet $HRPropertySet -UnfoundIdentitiesHash $UnfoundIdentitiesHash
                                }
                                'ExchangeOnPremises'
                                {
                                    #WriteLog -Message "Getting SendAS Permissions for Target $ID Via AD Commands" -entryType Notification
                                    Get-SendASPermisssionsViaADPS -TargetPublicFolder $ISR -TargetMailPublicFolder $ISRR -ExchangeSession $Script:PSSession -ADPSSession $script:ADPSSession -ObjectGUIDHash $ObjectGUIDHash -excludedTrusteeGUIDHash $excludedTrusteeGUIDHash -dropInheritedPermissions $dropInheritedPermissions -DomainPrincipalHash $DomainPrincipalHash -ExchangeOrganization $ExchangeOrganization -HRPropertySet $HRPropertySet -UnfoundIdentitiesHash $UnfoundIdentitiesHash -ADPSDriveName $ADPSDriveName
                                }
                            }
                        }
                    )
                    if ($expandGroups -eq $true)
                    {
                        #WriteLog -Message "Expanding Group Based Permissions for Target $ID" -entryType Notification
                        $splat = @{
                            Permission              = $PermissionExportObjects
                            ObjectGUIDHash          = $ObjectGUIDHash
                            SIDHistoryHash          = $SIDHistoryRecipientHash
                            excludedTrusteeGUIDHash = $excludedTrusteeGUIDHash
                            UnfoundIdentitiesHash   = $UnfoundIdentitiesHash
                            HRPropertySet           = $HRPropertySet
                            exchangeSession         = $Script:PSSession
                            TargetPublicFolder      = $ISR
                            TargetMailPublicFolder  = $ISRR
                        }
                        if ($dropExpandedParentGroupPermissions -eq $true)
                        { $splat.dropExpandedParentGroupPermissions = $true }
                        switch ($Script:ExchangeOrganizationType)
                        {
                            'ExchangeOnline'
                            {
                                $splat.UseExchangeCommandsInsteadOfADOrLDAP = $true
                            }
                            'ExchangeOnPremises'
                            {
                                $splat.ADPSSession = $Script:ADPSSession
                            }
                        }
                        $PermissionExportObjects = @(Expand-GroupPermission @splat)
                    }

                    if ($PermissionExportObjects.Count -eq 0 -and -not $ExcludeNonePermissionOutput -eq $true)
                    {
                        $GPEOParams = @{
                            TargetPublicFolder         = $ISR
                            TargetMailPublicFolder     = $ISRR
                            TrusteeIdentity            = 'Not Applicable'
                            TrusteeRecipientObject     = $null
                            PermissionType             = 'None'
                            AssignmentType             = 'None'
                            SourceExchangeOrganization = $ExchangeOrganization
                            None                       = $true
                        }
                        $NonPerm = NewPermissionExportObject @GPEOParams
                        $NonPerm
                    }
                    elseif ($PermissionExportObjects.Count -gt 0)
                    {
                        $PermissionExportObjects
                    }
                    WriteLog -Message $message -EntryType Succeeded
                }
                Catch
                {
                    WriteLog -Message $message -EntryType Failed
                }
            }#Foreach recipient in set
        )# end ExportedPermissions
        if ($ExportedPermissions.Count -ge 1)
        {
            Try
            {
                $message = "Export $($ExportedPermissions.Count) Exported Permissions to File $ExportedExchangePublicFolderPermissionsFile."
                WriteLog -Message $message -EntryType Attempting -verbose
                $ExportedPermissions | Export-Csv -Path $ExportedExchangePublicFolderPermissionsFile -NoClobber -Encoding UTF8 -ErrorAction Stop -NoTypeInformation
                WriteLog -Message $message -EntryType Succeeded -verbose
            }
            Catch
            {
                $myerror = $_
                WriteLog -Message $message -EntryType Failed -ErrorLog -Verbose
                WriteLog -Message $myError.tostring() -ErrorLog
                WriteLog -Message "Saving Exported Permissions to Global Variable $($BeginTimeStamp + "ExportedExchangePermissions") for recovery/manual export if desired/required." -verbose
                Set-Variable -Name $($BeginTimeStamp + "ExportedExchangePermissions") -Value $ExportedPermissions -Scope Global
            }
        }
        else
        {
            WriteLog -Message "No Permissions were generated for export by this operation.  Check the logs for errors if this is unexpected." -EntryType Notification -Verbose
        }
    }#end End

}
