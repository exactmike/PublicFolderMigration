cmdletbinding()]
param(
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential] $Credential
)

function WriteErrorSummary
{
    param ($folder, $operation, $errorMessage, $commandtext)

    WriteOperationSummary $folder.Guid $operation $errorMessage $commandtext
    $script:errorsEncountered++
}

# Writes the current progress
function WriteProgress
{
    param($statusFormat, $statusProcessed, $statusTotal)
    $WPParams = @{
        Activity        = $LocalizedStrings.ProgressBarActivity
        Status          = ($statusFormat -f $statusProcessed, $statusTotal)
        PercentComplete = (100 * ($script:itemsProcessed + $statusProcessed) / $script:totalItems)
    }
    Write-Progress @WPParams
}

# Invokes New-SyncMailPublicFolder to create a new MEPF object on AD
function NewMailEnabledPublicFolder
{
    param ($localFolder)

    if ($localFolder.PrimarySmtpAddress.ToString() -eq "")
    {
        $errorMsg = ($LocalizedStrings.FailedToCreateMailPublicFolderEmptyPrimarySmtpAddress -f $localFolder.Guid)
        Write-Error $errorMsg
        WriteErrorSummary $localFolder $LocalizedStrings.CreateOperationName $errorMsg ""
        return
    }

    # preserve the ability to reply via Outlook's nickname cache post-migration
    $emailAddressesArray = $localFolder.EmailAddresses.ToStringArray() + ("x500:" + $localFolder.LegacyExchangeDN)

    $newParams = @{ }
    AddNewOrSetCommonParameters $localFolder $emailAddressesArray $newParams

    [string]$commandText = (FormatCommand $script:NewSyncMailPublicFolderCommand $newParams)

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText
    }

    try
    {
        $result = &$script:NewSyncMailPublicFolderCommand @newParams
        WriteOperationSummary $localFolder $LocalizedStrings.CreateOperationName $LocalizedStrings.CsvSuccessResult $commandText

        if (-not $WhatIf)
        {
            $script:ObjectsCreated++
        }
    }
    catch
    {
        WriteErrorSummary $localFolder $LocalizedStrings.CreateOperationName $error[0].Exception.Message $commandText
        Write-Error $_
    }
}

# Invokes Remove-SyncMailPublicFolder to remove a MEPF from AD
function RemoveMailEnabledPublicFolder
{
    param ($remoteFolder)

    $removeParams = @{ }
    $removeParams.Add("Identity", $remoteFolder.DistinguishedName)
    $removeParams.Add("Confirm", $false)
    $removeParams.Add("WarningAction", [System.Management.Automation.ActionPreference]::SilentlyContinue)
    $removeParams.Add("ErrorAction", [System.Management.Automation.ActionPreference]::Stop)

    if ($WhatIf)
    {
        $removeParams.Add("WhatIf", $true)
    }

    [string]$commandText = (FormatCommand $script:RemoveSyncMailPublicFolderCommand $removeParams)

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText
    }

    try
    {
        &$script:RemoveSyncMailPublicFolderCommand @removeParams
        WriteOperationSummary $remoteFolder $LocalizedStrings.RemoveOperationName $LocalizedStrings.CsvSuccessResult $commandText

        if (-not $WhatIf)
        {
            $script:ObjectsDeleted++
        }
    }
    catch
    {
        WriteErrorSummary $remoteFolder $LocalizedStrings.RemoveOperationName $_.Exception.Message $commandText
        Write-Error $_
    }
}

# Invokes Set-MailPublicFolder to update the properties of an existing MEPF
function UpdateMailEnabledPublicFolder
{
    param ($localFolder, $remoteFolder)

    $localEmailAddresses = $localFolder.EmailAddresses.ToStringArray()
    $localEmailAddresses += ("x500:" + $localFolder.LegacyExchangeDN) # preserve the ability to reply via Outlook's nickname cache post-migration
    $emailAddresses = ConsolidateEmailAddresses $localEmailAddresses $remoteFolder.EmailAddresses $remoteFolder.LegacyExchangeDN

    $setParams = @{ }
    $setParams.Add("Identity", $remoteFolder.DistinguishedName)

    if ($script:mailEnabledSystemFolders.Contains($localFolder.Guid))
    {
        $setParams.Add("IgnoreMissingFolderLink", $true)
    }

    AddNewOrSetCommonParameters $localFolder $emailAddresses $setParams

    [string]$commandText = (FormatCommand $script:SetMailPublicFolderCommand $setParams)

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText
    }

    try
    {
        &$script:SetMailPublicFolderCommand @setParams
        WriteOperationSummary $remoteFolder $LocalizedStrings.UpdateOperationName $LocalizedStrings.CsvSuccessResult $commandText

        if (-not $WhatIf)
        {
            $script:ObjectsUpdated++
        }
    }
    catch
    {
        WriteErrorSummary $remoteFolder $LocalizedStrings.UpdateOperationName $_.Exception.Message $commandText
        Write-Error $_
    }
}

# Adds the common set of parameters between New and Set cmdlets to the given dictionary
function AddNewOrSetCommonParameters
{
    param ($localFolder, $emailAddresses, [System.Collections.IDictionary]$parameters)

    $windowsEmailAddress = $localFolder.WindowsEmailAddress.ToString()
    if ($windowsEmailAddress -eq "")
    {
        $windowsEmailAddress = $localFolder.PrimarySmtpAddress.ToString()
    }

    $parameters.Add("Alias", $localFolder.Alias.Trim())
    $parameters.Add("DisplayName", $localFolder.DisplayName.Trim())
    $parameters.Add("EmailAddresses", $emailAddresses)
    $parameters.Add("ExternalEmailAddress", $localFolder.PrimarySmtpAddress.ToString())
    $parameters.Add("HiddenFromAddressListsEnabled", $localFolder.HiddenFromAddressListsEnabled)
    $parameters.Add("Name", $localFolder.Name.Trim())
    $parameters.Add("OnPremisesObjectId", $localFolder.Guid)
    $parameters.Add("WindowsEmailAddress", $windowsEmailAddress)
    $parameters.Add("ErrorAction", [System.Management.Automation.ActionPreference]::Stop)

    if ($WhatIf)
    {
        $parameters.Add("WhatIf", $true)
    }
}

# Finds out the cloud-only email addresses and merges those with the values current persisted in the on-premises object
function ConsolidateEmailAddresses
{
    param($localEmailAddresses, $remoteEmailAddresses, $remoteLegDN)

    # Check if the email address in the existing cloud object is present on-premises; if it is not, then the address was either:
    # 1. Deleted on-premises and must be removed from cloud
    # 2. or it is a cloud-authoritative address and should be kept
    $remoteAuthoritative = @()
    foreach ($remoteAddress in $remoteEmailAddresses)
    {
        if ($remoteAddress.StartsWith("SMTP:", [StringComparison]::InvariantCultureIgnoreCase))
        {
            $found = $false
            $remoteAddressParts = $remoteAddress.Split($script:proxyAddressSeparators) # e.g. SMTP:alias@domain
            if ($remoteAddressParts.Length -ne 3)
            {
                continue # Invalid SMTP proxy address (it will be removed)
            }

            foreach ($localAddress in $localEmailAddresses)
            {
                # note that the domain part of email addresses is case insensitive while the alias part is case sensitive
                $localAddressParts = $localAddress.Split($script:proxyAddressSeparators)
                if ($localAddressParts.Length -eq 3 -and
                    $remoteAddressParts[0].Equals($localAddressParts[0], [StringComparison]::InvariantCultureIgnoreCase) -and
                    $remoteAddressParts[1].Equals($localAddressParts[1], [StringComparison]::InvariantCulture) -and
                    $remoteAddressParts[2].Equals($localAddressParts[2], [StringComparison]::InvariantCultureIgnoreCase))
                {
                    $found = $true
                    break
                }
            }

            if (-not $found)
            {
                foreach ($domain in $script:authoritativeDomains)
                {
                    if ($remoteAddressParts[2] -eq $domain)
                    {
                        $found = $true
                        break
                    }
                }

                if (-not $found)
                {
                    # the address on the remote object is from a cloud authoritative domain and should not be removed
                    $remoteAuthoritative += $remoteAddress
                }
            }
        }
        elseif ($remoteAddress.StartsWith("X500:", [StringComparison]::InvariantCultureIgnoreCase) -and
            $remoteAddress.Substring(5) -eq $remoteLegDN)
        {
            $remoteAuthoritative += $remoteAddress
        }
    }

    $localEmailAddresses + $remoteAuthoritative
}

# Formats the command and its parameters to be printed on console or to file
function FormatCommand
{
    param ([string]$command, [System.Collections.IDictionary]$parameters)

    $commandText = New-Object System.Text.StringBuilder
    [void]$commandText.Append($command)
    foreach ($name in $parameters.Keys)
    {
        [void]$commandText.AppendFormat(" -{0}:", $name)

        $value = $parameters[$name]
        if ($value -isnot [Array])
        {
            [void]$commandText.AppendFormat("`"{0}`"", $value)
        }
        elseif ($value.Length -eq 0)
        {
            [void]$commandText.Append("@()")
        }
        else
        {
            [void]$commandText.Append("@(")
            foreach ($subValue in $value)
            {
                [void]$commandText.AppendFormat("`"{0}`",", $subValue)
            }

            [void]$commandText.Remove($commandText.Length - 1, 1)
            [void]$commandText.Append(")")
        }
    }

    return $commandText.ToString()
}

$script:ObjectsCreated, $script:ObjectsUpdated, $script:ObjectsDeleted = 0, 0, 0
$script:NewSyncMailPublicFolderCommand = "New-EXOSyncMailPublicFolder"
$script:SetMailPublicFolderCommand = "Set-EXOMailPublicFolder"
$script:RemoveSyncMailPublicFolderCommand = "Remove-EXOSyncMailPublicFolder"
[char[]]$script:proxyAddressSeparators = ':', '@'
$script:errorsEncountered = 0
$script:authoritativeDomains = $null
$script:mailEnabledSystemFolders = New-Object 'System.Collections.Generic.HashSet[Guid]'
$script:WellKnownSystemFolders = @(
    "\NON_IPM_SUBTREE\EFORMS REGISTRY",
    "\NON_IPM_SUBTREE\OFFLINE ADDRESS BOOK",
    "\NON_IPM_SUBTREE\SCHEDULE+ FREE BUSY",
    "\NON_IPM_SUBTREE\schema-root",
    "\NON_IPM_SUBTREE\Events Root")

#load hashtable of localized string
Import-LocalizedData -BindingVariable LocalizedStrings -FileName SyncMailPublicFolders.strings.psd1



try
{
    #InitializeExchangeOnlineRemoteSession

    WriteInfoMessage $LocalizedStrings.LocalMailPublicFolderEnumerationStart

    # During finalization, Public Folders deployment is locked for migration, which means the script cannot invoke
    # Get-PublicFolder as that operation would fail. In that case, the script cannot determine which mail public folder
    # objects are linked to system folders under the NON_IPM_SUBTREE.
    $allSystemFoldersInAD = @()




    $localFolders = @(Get-MailPublicFolder -ResultSize:Unlimited -IgnoreDefaultScope | Sort-Object Guid)
    $remoteFolders = @(Get-EXOMailPublicFolder -ResultSize:Unlimited | Sort-Object OnPremisesObjectId)

    $missingOnPremisesGuid = @()
    $pendingRemoves = @()
    $pendingUpdates = @{ }
    $pendingAdds = @{ }

    $localIndex = 0
    $remoteIndex = 0
    while ($localIndex -lt $localFolders.Length -and $remoteIndex -lt $remoteFolders.Length)
    {
        $local = $localFolders[$localIndex]
        $remote = $remoteFolders[$remoteIndex]

        if ($remote.OnPremisesObjectId -eq "")
        {
            # This folder must be processed based on PrimarySmtpAddress
            $missingOnPremisesGuid += $remote
            $remoteIndex++
        }
        elseif ($local.Guid.ToString() -eq $remote.OnPremisesObjectId)
        {
            $pendingUpdates.Add($local.Guid, (New-Object PSObject -Property @{ Local = $local; Remote = $remote }))
            $localIndex++
            $remoteIndex++
        }
        elseif ($local.Guid.ToString() -lt $remote.OnPremisesObjectId)
        {
            if (-not $script:mailEnabledSystemFolders.Contains($local.Guid))
            {
                $pendingAdds.Add($local.Guid, $local)
            }

            $localIndex++
        }
        else
        {
            $pendingRemoves += $remote
            $remoteIndex++
        }
    }

    # Remaining folders on $localFolders collection must be added to Exchange Online
    while ($localIndex -lt $localFolders.Length)
    {
        $local = $localFolders[$localIndex]

        if (-not $script:mailEnabledSystemFolders.Contains($local.Guid))
        {
            $pendingAdds.Add($local.Guid, $local)
        }

        $localIndex++
    }

    # Remaining folders on $remoteFolders collection must be removed from Exchange Online
    while ($remoteIndex -lt $remoteFolders.Length)
    {
        $remote = $remoteFolders[$remoteIndex]
        if ($remote.OnPremisesObjectId -eq "")
        {
            # This folder must be processed based on PrimarySmtpAddress
            $missingOnPremisesGuid += $remote
        }
        else
        {
            $pendingRemoves += $remote
        }

        $remoteIndex++
    }

    if ($missingOnPremisesGuid.Length -gt 0)
    {
        # Process remote objects missing the OnPremisesObjectId using the PrimarySmtpAddress as a key instead.
        $missingOnPremisesGuid = @($missingOnPremisesGuid | Sort-Object PrimarySmtpAddress)
        $localFolders = @($localFolders | Sort-Object PrimarySmtpAddress)

        $localIndex = 0
        $remoteIndex = 0
        while ($localIndex -lt $localFolders.Length -and $remoteIndex -lt $missingOnPremisesGuid.Length)
        {
            $local = $localFolders[$localIndex]
            $remote = $missingOnPremisesGuid[$remoteIndex]

            if ($local.PrimarySmtpAddress.ToString() -eq $remote.PrimarySmtpAddress.ToString())
            {
                # Make sure the PrimarySmtpAddress has no duplicate on-premises; otherwise, skip updating all objects with duplicate address
                $j = $localIndex + 1
                while ($j -lt $localFolders.Length)
                {
                    $next = $localFolders[$j]
                    if ($local.PrimarySmtpAddress.ToString() -ne $next.PrimarySmtpAddress.ToString())
                    {
                        break
                    }

                    WriteErrorSummary $next $LocalizedStrings.UpdateOperationName ($LocalizedStrings.PrimarySmtpAddressUsedByAnotherFolder -f $local.PrimarySmtpAddress, $local.Guid) ""

                    # If there were a previous match based on OnPremisesObjectId, remove the folder operation from add and update collections
                    $pendingAdds.Remove($next.Guid)
                    $pendingUpdates.Remove($next.Guid)
                    $j++
                }

                $duplicatesFound = $j - $localIndex - 1
                if ($duplicatesFound -gt 0)
                {
                    # If there were a previous match based on OnPremisesObjectId, remove the folder operation from add and update collections
                    $pendingAdds.Remove($local.Guid)
                    $pendingUpdates.Remove($local.Guid)
                    $localIndex += $duplicatesFound + 1

                    WriteErrorSummary $local $LocalizedStrings.UpdateOperationName ($LocalizedStrings.PrimarySmtpAddressUsedByOtherFolders -f $local.PrimarySmtpAddress, $duplicatesFound) ""
                    WriteWarningMessage ($LocalizedStrings.SkippingFoldersWithDuplicateAddress -f ($duplicatesFound + 1), $local.PrimarySmtpAddress)
                }
                elseif ($pendingUpdates.Contains($local.Guid))
                {
                    # If we get here, it means two different remote objects match the same local object (one by OnPremisesObjectId and another by PrimarySmtpAddress).
                    # Since that is an ambiguous resolution, let's skip updating the remote objects.
                    $ambiguousRemoteObj = $pendingUpdates[$local.Guid].Remote
                    $pendingUpdates.Remove($local.Guid)

                    $errorMessage = ($LocalizedStrings.AmbiguousLocalMailPublicFolderResolution -f $local.Guid, $ambiguousRemoteObj.Guid, $remote.Guid)
                    WriteErrorSummary $local $LocalizedStrings.UpdateOperationName $errorMessage ""
                    WriteWarningMessage $errorMessage
                }
                else
                {
                    # Since there was no match originally using OnPremisesObjectId, the local object was treated as an add to Exchange Online.
                    # In this way, since we now found a remote object (by PrimarySmtpAddress) to update, we must first remove the local object from the add list.
                    $pendingAdds.Remove($local.Guid)
                    $pendingUpdates.Add($local.Guid, (New-Object PSObject -Property @{ Local = $local; Remote = $remote }))
                }

                $localIndex++
                $remoteIndex++
            }
            elseif ($local.PrimarySmtpAddress.ToString() -gt $remote.PrimarySmtpAddress.ToString())
            {
                # There are no local objects using the remote object's PrimarySmtpAddress
                $pendingRemoves += $remote
                $remoteIndex++
            }
            else
            {
                $localIndex++
            }
        }

        # All objects remaining on the $missingOnPremisesGuid list no longer exist on-premises
        while ($remoteIndex -lt $missingOnPremisesGuid.Length)
        {
            $pendingRemoves += $missingOnPremisesGuid[$remoteIndex]
            $remoteIndex++
        }
    }

    $script:totalItems = $pendingRemoves.Length + $pendingUpdates.Count + $pendingAdds.Count

    # At this point, we know all changes that need to be synced to Exchange Online.

    # Find out the authoritative AcceptedDomains on-premises so that we don't accidently remove cloud-only email addresses during updates
    $script:authoritativeDomains = @(Get-AcceptedDomain | Where-Object { $_.DomainType -eq "Authoritative" } | ForEach-Object { $_.DomainName.ToString() })

    # Finally, let's perfom the actual operations against Exchange Online
    $script:itemsProcessed = 0
    for ($i = 0; $i -lt $pendingRemoves.Length; $i++)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusRemoving $i $pendingRemoves.Length
        RemoveMailEnabledPublicFolder $pendingRemoves[$i]
    }

    $script:itemsProcessed += $pendingRemoves.Length
    $updatesProcessed = 0
    foreach ($folderPair in $pendingUpdates.Values)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusUpdating $updatesProcessed $pendingUpdates.Count
        UpdateMailEnabledPublicFolder $folderPair.Local $folderPair.Remote
        $updatesProcessed++
    }

    $script:itemsProcessed += $pendingUpdates.Count
    $addsProcessed = 0
    foreach ($localFolder in $pendingAdds.Values)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusCreating $addsProcessed $pendingAdds.Count
        NewMailEnabledPublicFolder $localFolder
        $addsProcessed++
    }

    Write-Progress -Activity $LocalizedStrings.ProgressBarActivity -Status ($LocalizedStrings.ProgressBarStatusCreating -f $pendingAdds.Count, $pendingAdds.Count) -Completed
    WriteInfoMessage ($LocalizedStrings.SyncMailPublicFolderObjectsComplete -f $script:ObjectsCreated, $script:ObjectsUpdated, $script:ObjectsDeleted)

    if ($script:errorsEncountered -gt 0)
    {
        WriteWarningMessage ($LocalizedStrings.ErrorsFoundDuringImport -f $script:errorsEncountered, (Get-Item $CsvSummaryFile).FullName)
    }
