    Function Export-PublicFolderPermission
    {

        [cmdletbinding(DefaultParameterSetName = 'AllPublicFolders')]
        param
        (
            [Parameter(ParameterSetName = 'AllPublicFolders',Mandatory)]
            [parameter(ParameterSetName = 'Scoped',Mandatory)]
            [ValidateScript({TestIsWriteableDirectory -Path $_})]
            $OutputFolderPath
            ,
            [parameter()]
            [ValidateScript({TestADPSDrive -name $_ -IsRootofDirectory})]
            $ADPSDriveName
            ,
            [parameter(ParameterSetName = 'Scoped')]
            [switch]$Recurse
            ,
            [parameter(ParameterSetName = 'Scoped')]
            [string[]]$PublicFolderPath = @()
            ,
            #Public Folder identities to exclude from permissions gathering (use folder name, full path, or EntryID).  EntryID is preferred as it is guaranteed to be unique.
            [parameter()]
            [string[]]$ExcludedIdentities
            ,
            [parameter()]#These will be resolved to trustee objects
            [string[]]$ExcludedTrusteeIdentities
            ,
            [parameter(ParameterSetName = 'Scoped')]
            [Parameter(ParameterSetName = 'AllPublicFolders')]
            [bool]$IncludeClientPermission = $true
            ,
            [parameter(ParameterSetName = 'Scoped')]
            [Parameter(ParameterSetName = 'AllPublicFolders')]
            [bool]$IncludeSendAs = $true
            ,
            [parameter(ParameterSetName = 'Scoped')]
            [Parameter(ParameterSetName = 'AllPublicFolders')]
            [bool]$IncludeSendOnBehalf = $true
            ,
            [bool]$ExpandGroups = $true
            ,
            [bool]$DropExpandedParentGroupPermissions = $false
            ,
            [bool]$DropInheritedPermissions = $false
            ,
            [switch]$IncludeSIDHistory
            ,
            [switch]$ExcludeNonePermissionOutput
            ,
            [switch]$EnableResume
            ,
            [switch]$KeepExportedPermissionsInGlobalVariable
            ,
            [Parameter(ParameterSetName = 'Resume',Mandatory)]
            [ValidateScript({Test-Path -Path $_})]
            [string]$ResumeFile
        )#End Param
        Begin
        {
            $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
            $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'ExchangePublicFolderPermissionsExportOperations.log')
            $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'ExchangePublicFolderPermissionsExportOperations-ERRORS.log')
            #$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            switch ($script:ConnectExchangeOrganizationCompleted)
            {
                $true
                {
                    switch (TestExchangePSSession -PSSession $script:PSSession)
                    {
                        $true
                        {
                            WriteLog -Message 'Using Existing PSSession' -EntryType Notification
                        }
                        $false
                        {
                            WriteLog -Message 'Removing Existing Failed PSSession' -EntryType Notification
                            RemoveExchangePSSession -Session $script:PsSession
                            WriteLog -Message 'Establishing New PSSession to Exchange Organization' -EntryType Notification
                            $GetExchangePSSessionParams = GetGetExchangePSSessionParams
                            $script:PsSession = GetExchangePSSession @GetExchangePSSessionParams
                        }
                    }
                }
                $false
                {
                    WriteUserInstructionError
                }
            }
            $ExchangeOrganization = Invoke-Command -Session $Script:PSSession -ScriptBlock {Get-OrganizationConfig | Select-Object -ExpandProperty Identity | Select-Object -ExpandProperty Name}
            $ExchangeOrganizationIsInExchangeOnline = $ExchangeOrganization -like '*.onmicrosoft.com'

            if ($ExchangeOrganizationIsInExchangeOnline -eq $false -and $null -eq $ADPSDriveName -and ($includeSidHistory -or $IncludeSendAs -or $ExpandGroups))
            {
                throw ('You need to use the ADPSDrive name parameter to provide an existing PowerShell Active Directory PSdrive connection to the AD forest where Exchange is installed')
            }
            if ($ExchangeOrganizationIsInExchangeOnline -eq $true -and $IncludeSIDHistory)
            {
                throw ('You cannot include SidHistory when your Exchange Organization is in Exchange Online.')
            }

            #Configure properties to retain in memory / hashtables for retrieved public folders and Recipients
            $PFPropertySet = @('EntryID','Identity','Name','ParentPath','FolderType','Has*','HiddenFromAddressListsEnabled','*Quota','MailEnabled','Replicas','ReplicationSchedule','RetainDeletedItemsFor','Use*')
            $HRPropertySet = @('*name*','*addr*','RecipientType*','*Id','Identity','GrantSendOnBehalfTo')
            switch ($PSCmdlet.ParameterSetName -eq 'Resume')
            {
                $true
                {
                    $ImportedExchangePermissionsExportResumeData = ImportExchangePermissionExportResumeData -Path $ResumeFile
                    $ExcludedPublicFoldersEntryIDHash = $ImportedExchangePermissionsExportResumeData.ExcludedPublicFoldersEntryIDHash
                    $ExcludedTrusteeGuidHash = $ImportedExchangePermissionsExportResumeData.ExcludedTrusteeGuidHash
                    $InScopeFolders = $ImportedExchangePermissionsExportResumeData.InScopeFolders
                    $InScopeMailPublicFoldersHash = $ImportedExchangePermissionsExportResumeData.InScopeMailPublicFoldersHash
                    $InScopeFolderCount = $InScopeFolders.count
                    $ResumeIdentity = $ImportedExchangePermissionsExportResumeData.ResumeID
                    [uint32]$Script:PermissionIdentity = $ImportedExchangePermissionsExportResumeData.NextPermissionIdentity
                    $ExportedExchangePublicFolderPermissionsFile = $ImportedExchangePermissionsExportResumeData.ExportedExchangePublicFolderPermissionsFile
                    $ResumeIndex = $ImportedExchangePermissionsExportResumeData.ResumeIndex
                    foreach ($v in $ImportedExchangePermissionsExportResumeData.ExchangePermissionsExportParameters)
                    {
                        if ($v.name -ne 'ExchangeSession') #why are we doing this?
                        {
                            Set-Variable -Name $v.name -Value $v.value -Force
                        }
                    }
                    WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
                    WriteLog -Message "Exchange Session is Running in Exchange Organzation $ExchangeOrganization" -EntryType Notification
                    if ($null -eq $ResumeIndex -or $ResumeIndex.gettype().name -notlike '*int*')
                    {
                        $message = "ResumeIndex is invalid.  Check/Edit the *ResumeID.xml file for a valid ResumeIdentity GUID."
                        WriteLog -Message $message -ErrorLog -EntryType Failed
                        Throw($message)
                    }
                    WriteLog -Message "Resume index set to $ResumeIndex based on ResumeIdentity $resumeIdentity" -EntryType Notification
                }
                $false
                {
                    WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
                    WriteLog -Message "Exchange Session is Running in Exchange Organzation $ExchangeOrganization" -EntryType Notification
                    $ExportedExchangePublicFolderPermissionsFile = Join-Path -Path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'ExportedExchangePublicFolderPermissions.csv')
                    $ResumeIndex = 0
                    [uint32]$Script:PermissionIdentity = 0
                    #create a property set for storing of recipient data during processing.  We don't need all attributes in memory/storage.
                    #Region GetExcludedRecipients
                    if ($PSBoundParameters.ContainsKey('ExcludedIdentities'))
                    {
                        try
                        {
                            $message = "Get public folder object(s) from Exchange Organization $ExchangeOrganization for the $($ExcludedIdentities.Count) ExcludedIdentities provided."
                            WriteLog -Message $message -EntryType Attempting -verbose
                            $excludedPublicFolders = @(
                                $ExcludedIdentities | ForEach-Object {
                                    $splat = @{
                                        Identity = $_
                                        ErrorAction = 'Stop'
                                    }
                                    Invoke-Command -Session $Script:PSSession -ScriptBlock {Get-PublicFolder @Using:splat | Select-Object -Property $using:PFPropertySet} -ErrorAction 'Stop'
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
                        WriteLog -Message "Got $($excludedPublicFolders.count) Excluded Objects" -EntryType Notification
                        $excludedPublicFoldersEntryIDHash = $excludedPublicFolders | Group-Object -Property EntryID -AsString -AsHashTable -ErrorAction Stop
                    }
                    else
                    {
                        $excludedPublicFoldersEntryIDHash = @{}
                    }
                    #EndRegion GetExcludedRecipients

                    #Region GetExcludedTrustees
                    if ($PSBoundParameters.ContainsKey('ExcludedTrusteeIdentities'))
                    {
                        try
                        {
                            $message = "Get recipent object(s) from Exchange Organization $ExchangeOrganization for the $($ExcludedTrusteeIdentities.Count) ExcludedTrusteeIdentities provided."
                            WriteLog -Message $message -EntryType Attempting -verbose
                            $excludedTrusteeRecipients = @(
                                $ExcludedTrusteeIdentities | ForEach-Object {
                                    $splat = @{
                                        Identity = $_
                                        ErrorAction = 'Stop'
                                    }
                                    Invoke-Command -Session $Script:PSSession -ScriptBlock {Get-Recipient @Using:splat | Select-Object -Property $using:HRPropertySet} -ErrorAction 'Stop'
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
                        $excludedTrusteeGUIDHash = $excludedTrusteeRecipients | Group-Object -Property GUID -AsString -AsHashTable -ErrorAction Stop
                    }
                    else
                    {
                        $excludedTrusteeGUIDHash = @{}
                    }
                    #EndRegion GetExcludedTrustees

                    #Region GetInScopePublicFolders
                    Try
                    {
                        switch ($PSCmdlet.ParameterSetName)
                        {
                            'Scoped'
                            {
                                WriteLog -Message "Operation: Scoped Permission retrieval for Public Folders with $($PublicFolderPath.Count) Public Folder Path(s) provided."
                                $message = "Get Public Folder object(s) for each provided Identity in Exchange Organization $ExchangeOrganization."
                                WriteLog -Message $message -EntryType Attempting -verbose
                                $InScopeFolders = @(
                                    $PublicFolderPath | ForEach-Object {
                                        $Splat = @{
                                            Identity = $_
                                            ErrorAction = 'Stop'
                                        }
                                        if ($Recurse -eq $true) {$Splat.Recurse = $true}
                                        Invoke-Command -Session $Script:PSSession -ScriptBlock {Get-PublicFolder @Using:splat | Select-Object -Property $Using:PFPropertySet} -ErrorAction Stop
                                    }
                                )
                                WriteLog -Message $message -EntryType Succeeded -verbose
                            }#end Scoped
                            'AllPublicFolders'
                            {
                                WriteLog -Message "Operation: Permission retrieval for all Public Folders."
                                $message = "Get all available Public Folder objects (from the non-system subtree) in Exchange Organization $ExchangeOrganization."
                                WriteLog -Message $message -EntryType Attempting -verbose
                                $splat = @{
                                    ResultSize = 'Unlimited'
                                    ErrorAction = 'Stop'
                                    Recurse = $true
                                }
                                $InScopeFolders = @(Invoke-Command -Session $Script:PSSession -ScriptBlock {Get-PublicFolder @Using:splat | Select-Object -Property $Using:PFPropertySet} -ErrorAction Stop)
                                WriteLog -Message $message -EntryType Succeeded -verbose
                            }#end AllMailboxes
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
                    #EndRegion GetInScopeFolders

                    #Region GetInScopeMailPublicFolders
                    $message = 'Get Mail Enabled Public Folders To support retrieval of SendAS and/or SendOnBehalf Permissions and for additional output information for ClientPermissions.'
                    WriteLog -message $message -entryType Attempting -verbose
                    $InScopeMailPublicFolders = @(GetMailPublicFolderPerUserPublicFolder -ExchangeSession $script:PSSession -PublicFolder $InScopeFolders -ErrorAction Stop)
                    WriteLog -message $message -entryType Succeeded -verbose
                    WriteLog -Message "Got $($InScopeMailPublicFolders.count) In Scope Mail Public Folder Objects" -EntryType Notification -verbose
                    $InScopeMailPublicFoldersHash = $InScopeMailPublicFolders | Group-Object -AsHashTable -Property EntryID -AsString
                    if ($null -eq $InScopeMailPublicFoldersHash) {$InScopeMailPublicFoldersHash = @{}}
                    #Region GetInScopeMailPublicFolders

                    #Region GetSIDHistoryData
                    if ($IncludeSIDHistory -eq $true)
                    {
                        $SIDHistoryRecipientHash = GetSIDHistoryRecipientHash -ADPSDriveName $ADPSDriveName -ExchangeSession $Script:PSSession -ErrorAction Stop
                    }
                    else
                    {
                        $SIDHistoryRecipientHash = @{}
                    }
                    #EndRegion GetSIDHistoryData
                }
            }
            # Setup for Possible Resume if requested by the user
            if ($EnableResume -eq $true)
            {
                $ExportExchangePermissionsExportResumeData = @{
                    excludedPublicFoldersEntryIDHash = $excludedPublicFoldersEntryIDHash
                    ExcludedTrusteeGuidHash = $ExcludedTrusteeGuidHash
                    SIDHistoryRecipientHash = $SIDHistoryRecipientHash
                    InScopeFolders = $InScopeFolders
                    InScopeMailPublicFoldersHash = $InScopeMailPublicFoldersHash
                    outputFolderPath = $outputFolderPath
                    ExportedExchangePublicFolderPermissionsFile = $ExportedExchangePublicFolderPermissionsFile
                    TimeStamp = $BeginTimeStamp
                    ErrorAction = 'Stop'
                }
                switch ($PSCmdlet.ParameterSetName -eq 'Resume')
                {
                    $true
                    {
                        $ExportExchangePermissionsExportResumeData.ExchangePermissionsExportParameters = $ImportedExchangePermissionsExportResumeData.ExchangePermissionsExportParameters
                    }
                    $false
                    {
                        $ExportExchangePermissionsExportResumeData.ExchangePermissionsExportParameters = @(GetAllParametersWithAValue -boundparameters $PSBoundParameters -allparameters $MyInvocation.MyCommand.Parameters)
                    }
                }
                $message = "Enable Resume and Export Resume Data"
                WriteLog -Message $message -EntryType Attempting
                $ResumeFile = ExportExchangePermissionExportResumeData @ExportExchangePermissionsExportResumeData
                $message = $message + " to file $ResumeFile"
                WriteLog -Message $message -EntryType Succeeded
            }
            #Region BuildLookupHashTables
            #these have to be populated as we go
            WriteLog -Message "Building Recipient Lookup HashTables" -EntryType Notification
            $DomainPrincipalHash = @{}
            $UnfoundIdentitiesHash = @{}
            $ObjectGUIDHash = @{}
            if ($expandGroups -eq $true)
            {
                $script:ExpandedGroupsNonGroupMembershipHash = @{}
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
                    $(if ($Recovering) {$i = $ResumeIndex} else {$i++})
                    #$ISR in $InScopeFolders[$ResumeIndex..$()]
                )
                {
                    $Recovering = $false
                    $ISRCounter++
                    $ISR = $InScopeFolders[$i]
                    $ID = $ISR.EntryID.tostring()
                    if ($excludedPublicFoldersEntryIDHash.ContainsKey($ID))
                    {
                        WriteLog -Message "Excluding Excluded Folder with EntryID $ID"
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
                    Write-Progress -Activity $message -status "Items processed: $($ISRCounter) of $($InScopeFolderCount)" -percentComplete (($ISRCounter / $InScopeFolderCount)*100)
                    Try
                    {
                        WriteLog -Message $message -EntryType Attempting
                        $PermissionExportObjects = @(
                            If ($IncludeSendOnBehalf -and $InScopeMailPublicFoldersHash.ContainsKey($ID))
                            {
                                WriteLog -Message "Getting SendOnBehalf Permissions for Target $ID" -entryType Notification
                                GetSendOnBehalfPermission -TargetPublicFolder $ISR -TargetMailPublicFolder $ISRR -ObjectGUIDHash $ObjectGUIDHash -ExchangeSession $Script:PSSession -ExcludedTrusteeGUIDHash $excludedTrusteeGUIDHash -ExchangeOrganization $ExchangeOrganization -HRPropertySet $HRPropertySet -DomainPrincipalHash $DomainPrincipalHash -UnfoundIdentitiesHash $UnfoundIdentitiesHash
                            }
                            If ($IncludeClientPermission)
                            {
                                WriteLog -Message "Getting Client Permissions for Target $ID" -entryType Notification
                                GetClientPermission -TargetPublicFolder $ISR -TargetMailPublicFolder $ISRR -ObjectGUIDHash $ObjectGUIDHash -ExchangeSession $Script:PSSession -excludedTrusteeGUIDHash $excludedTrusteeGUIDHash -ExchangeOrganization $ExchangeOrganization -DomainPrincipalHash $DomainPrincipalHash -HRPropertySet $HRPropertySet -UnfoundIdentitiesHash $UnfoundIdentitiesHash
                            }
                            If ($IncludeSendAs -and $InScopeMailPublicFoldersHash.ContainsKey($ID))
                            {
                                WriteLog -Message "Getting SendAS Permissions for Target $ID" -entryType Notification
                                if ($ExchangeOrganizationIsInExchangeOnline)
                                {
                                    #WriteLog -Message "Getting SendAS Permissions for Target $ID Via Exchange Commands" -entryType Notification
                                    GetSendASPermissionsViaExchange -TargetPublicFolder $ISRR -TargetMailPublicFolder $ISRR -ExchangeSession $Script:PSSession -ObjectGUIDHash $ObjectGUIDHash -excludedTrusteeGUIDHash $ -dropInheritedPermissions $dropInheritedPermissions -DomainPrincipalHash $DomainPrincipalHash -ExchangeOrganization $ExchangeOrganization -ExchangeOrganizationIsInExchangeOnline $ExchangeOrganizationIsInExchangeOnline -HRPropertySet $HRPropertySet -UnfoundIdentitiesHash $UnfoundIdentitiesHash
                                }
                                else
                                {
                                    #WriteLog -Message "Getting SendAS Permissions for Target $ID Via AD Commands" -entryType Notification
                                    GetSendASPermisssionsViaADPSDrive -TargetPublicFolder $ISR -TargetMailPublicFolder $ISRR -ExchangeSession $Script:PSSession -ObjectGUIDHash $ObjectGUIDHash -excludedTrusteeGUIDHash $excludedTrusteeGUIDHash -dropInheritedPermissions $dropInheritedPermissions -DomainPrincipalHash $DomainPrincipalHash -ExchangeOrganization $ExchangeOrganization -ExchangeOrganizationIsInExchangeOnlin $ExchangeOrganizationIsInExchangeOnline -HRPropertySet $HRPropertySet -UnfoundIdentitiesHash $UnfoundIdentitiesHash -ADPSDriveName $ADPSDriveName
                                }
                            }
                        )
                        if ($expandGroups -eq $true)
                        {
                            WriteLog -Message "Expanding Group Based Permissions for Target $ID" -entryType Notification
                            $splat = @{
                                Permission = $PermissionExportObjects
                                ObjectGUIDHash = $ObjectGUIDHash
                                SIDHistoryHash = $SIDHistoryRecipientHash
                                excludedTrusteeGUIDHash = $excludedTrusteeGUIDHash
                                UnfoundIdentitiesHash = $UnfoundIdentitiesHash
                                HRPropertySet = $HRPropertySet
                                exchangeSession = $Script:PSSession
                                TargetPublicFolder = $ISR
                                TargetMailPublicFolder = $ISRR
                            }
                            if ($dropExpandedParentGroupPermissions -eq $true)
                            {$splat.dropExpandedParentGroupPermissions = $true}
                            if ($ExchangeOrganizationIsInExchangeOnline)
                            {
                                $splat.UseExchangeCommandsInsteadOfADOrLDAP = $true
                            }
                            else
                            {
                                $splat.ADPSDriveName = $ADPSDriveName
                            }
                            $PermissionExportObjects = @(ExpandGroupPermission @splat)
                        }
                        if (TestExchangePSSession -PSSession $Script:PSSession)
                        {
                            if ($PermissionExportObjects.Count -eq 0 -and -not $ExcludeNonePermissionOutput -eq $true)
                            {
                                $GPEOParams = @{
                                    TargetPublicFolder = $ISR
                                    TargetMailPublicFolder = $ISRR
                                    TrusteeIdentity = 'Not Applicable'
                                    TrusteeRecipientObject = $null
                                    PermissionType = 'None'
                                    AssignmentType = 'None'
                                    SourceExchangeOrganization = $ExchangeOrganization
                                    None = $true
                                }
                                $NonPerm = NewPermissionExportObject @GPEOParams
                                Write-Output $NonPerm
                            }
                            elseif ($PermissionExportObjects.Count -gt 0)
                            {
                                Write-Output $PermissionExportObjects
                            }
                            WriteLog -Message $message -EntryType Succeeded
                        }
                        else
                        {
                            WriteLog -Message 'Removing Existing Failed PSSession' -EntryType Notification -verbose
                            Remove-PSSession -Session $script:PsSession -ErrorAction SilentlyContinue
                            WriteLog -Message 'Establish New PSSession to Exchange Organization' -EntryType Attempting -verbose
                            $GetExchangePSSessionParams = GetGetExchangePSSessionParams
                            try
                            {
                                Start-Sleep -Seconds 10
                                $script:PsSession = GetExchangePSSession @GetExchangePSSessionParams
                                WriteLog -Message 'Establish New PSSession to Exchange Organization' -EntryType Succeeded -verbose
                                $ResumeIndex = $i
                                $ISRCounter--
                                $Recovering = $true
                                continue nextISR
                            }
                            catch
                            {
                                $myerror = $_
                                WriteLog -Message 'Establish New PSSession to Exchange Organization' -EntryType Failed
                                WriteLog -Message $myerror.tostring() -ErrorLog -Verbose
                                WriteLog -Message $message -EntryType Failed -ErrorLog -Verbose
                                $exitmessage = "Testing Showed that Exchange Session Failed/Disconnected during permission processing for ID $ID."
                                WriteLog -Message $exitmessage -EntryType Notification -ErrorLog -Verbose
                                if ($EnableResume -eq $true)
                                {
                                    WriteLog -Message "Resume File $ResumeFile is available to resume this operation after you have re-connected the Exchange Session" -Verbose
                                    WriteLog -Message "Resume Recipient ID is $ID" -Verbose
                                    $ResumeIDFile = ExportResumeID -ID $ID -outputFolderPath $OutputFolderPath -TimeStamp $BeginTimeStamp -NextPermissionID $Script:PermissionIdentity -ResumeIndex $i
                                    WriteLog -Message "Resume ID $ID exported to file $resumeIDFile" -Verbose
                                    WriteLog -Message "Next Permission Identity $($Script:PermissionIdentity) exported to file $resumeIDFile" -Verbose
                                    $message = "Run `'Get-ExchangePermission -ResumeFile $ResumeFile`' and also specify any common parameters desired (such as -verbose) since common parameters are not included in the Resume Data File."
                                    WriteLog -Message $message -EntryType Notification -verbose
                                }
                                Break nextISR
                            }
                        }
                    }
                    Catch
                    {
                        $myerror = $_
                        WriteLog -Message $message -EntryType Failed -ErrorLog -Verbose
                        $exitmessage = "Exchange Session Failed/Disconnected during permission processing for ID $ID. The next Log entry is the error from the Exchange Session."
                        WriteLog -Message $exitmessage -EntryType Notification -ErrorLog -Verbose
                        WriteLog -Message $myError.tostring() -ErrorLog -Verbose
                        WriteLog -Message 'Removing Existing Failed PSSession' -EntryType Notification
                        Remove-PSSession -Session $script:PsSession -ErrorAction SilentlyContinue
                        WriteLog -Message 'Establish New PSSession to Exchange Organization' -EntryType Attempting
                        $GetExchangePSSessionParams = GetGetExchangePSSessionParams
                        try
                        {
                            Start-Sleep -Seconds 10
                            $script:PsSession = GetExchangePSSession @GetExchangePSSessionParams
                            WriteLog -Message 'Establish New PSSession to Exchange Organization' -EntryType Succeeded
                            $ResumeIndex = $i
                            $ISRCounter--
                            $Recovering = $true
                            continue nextISR
                        }
                        catch
                        {
                            $myerror = $_
                            WriteLog -Message 'Establish New PSSession to Exchange Organization' -EntryType Failed
                            WriteLog -Message $myerror.tostring() -ErrorLog -Verbose
                            WriteLog -Message $message -EntryType Failed -ErrorLog -Verbose
                            $exitmessage = "Testing Showed that Exchange Session Failed/Disconnected during permission processing for ID $ID."
                            WriteLog -Message $exitmessage -EntryType Notification -ErrorLog -Verbose
                            if ($EnableResume -eq $true)
                            {
                                WriteLog -Message "Resume File $ResumeFile is available to resume this operation after you have re-connected the Exchange Session" -Verbose
                                WriteLog -Message "Resume Recipient ID is $ID" -Verbose
                                $ResumeIDFile = ExportResumeID -ID $ID -outputFolderPath $OutputFolderPath -TimeStamp $BeginTimeStamp -NextPermissionID $Script:PermissionIdentity -ResumeIndex $i
                                WriteLog -Message "Resume ID $ID exported to file $resumeIDFile" -Verbose
                                WriteLog -Message "Next Permission Identity $($Script:PermissionIdentity) exported to file $resumeIDFile" -Verbose
                                $message = "Run `'Get-ExchangePermission -ResumeFile $ResumeFile`' and also specify any common parameters desired (such as -verbose) since common parameters are not included in the Resume Data File."
                                WriteLog -Message $message -EntryType Notification -verbose
                            }
                            Break nextISR
                        }
                    }
                }#Foreach recipient in set
            )# end ExportedPermissions
            if ($ExportedPermissions.Count -ge 1)
            {
                Try
                {
                    $message = "Export $($ExportedPermissions.Count) Exported Permissions to File $ExportedExchangePublicFolderPermissionsFile."
                    WriteLog -Message $message -EntryType Attempting -verbose
                    switch ($PSCmdlet.ParameterSetName -eq 'Resume')
                    {
                        $true
                        {
                            $ExportedPermissions | Export-Csv -Path $ExportedExchangePublicFolderPermissionsFile -Append -Encoding UTF8 -ErrorAction Stop -NoTypeInformation #-Force
                        }
                        $false
                        {
                            $ExportedPermissions | Export-Csv -Path $ExportedExchangePublicFolderPermissionsFile -NoClobber -Encoding UTF8 -ErrorAction Stop -NoTypeInformation
                        }
                    }
                    WriteLog -Message $message -EntryType Succeeded -verbose
                    if ($KeepExportedPermissionsInGlobalVariable -eq $true)
                    {
                        WriteLog -Message "Saving Exported Permissions to Global Variable $($BeginTimeStamp + "ExportedExchangePermissions") for recovery/manual export." -Verbose
                        Set-Variable -Name $($BeginTimeStamp + "ExportedExchangePermissions") -Value $ExportedPermissions -Scope Global
                    }
                }
                Catch
                {
                    $myerror = $_
                    WriteLog -Message $message -EntryType Failed -ErrorLog -Verbose
                    WriteLog -Message $myError.tostring() -ErrorLog
                    WriteLog -Message "Saving Exported Permissions to Global Variable $($BeginTimeStamp + "ExportedExchangePermissions") for recovery/manual export if desired/required.  This is separate from performing a Resume with a Resume file." -verbose
                    Set-Variable -Name $($BeginTimeStamp + "ExportedExchangePermissions") -Value $ExportedPermissions -Scope Global
                }
            }
            else
            {
                WriteLog -Message "No Permissions were generated for export by this operation.  Check the logs for errors if this is unexpected." -EntryType Notification -Verbose
            }
        }#end End

    }

