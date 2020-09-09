# .SYNOPSIS
#    Syncs mail-enabled public folder objects from the local Exchange deployment into O365. It uses the local Exchange deployment
#    as master to determine what changes need to be applied to O365. The script will create, update or delete mail-enabled public
#    folder objects on O365 Active Directory when appropriate.
#
# .DESCRIPTION
#    The script must be executed from an Exchange 2007 or 2010 Management Shell window providing access to mail public folders in
#    the local Exchange deployment. Then, using the credentials provided, the script will create a session against Exchange Online,
#    which will be used to manipulate O365 Active Directory objects remotely.
#
#    Copyright (c) 2014 Microsoft Corporation. All rights reserved.
#
#    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# .PARAMETER Credential
#    Exchange Online user name and password.
#
# .PARAMETER CsvSummaryFile
#    The file path where sync operations and errors will be logged in a CSV format.
#
# .PARAMETER ConnectionUri
#    The Exchange Online remote PowerShell connection uri. If you are an Office 365 operated by 21Vianet customer in China, use "https://partner.outlook.cn/PowerShell".
#
# .PARAMETER Confirm
#    The Confirm switch causes the script to pause processing and requires you to acknowledge what the script will do before processing continues. You don't have to specify
#    a value with the Confirm switch.
#
# .PARAMETER Force
#    Force the script execution and bypass validation warnings.
#
# .PARAMETER WhatIf
#    The WhatIf switch instructs the script to simulate the actions that it would take on the object. By using the WhatIf switch, you can view what changes would occur
#    without having to apply any of those changes. You don't have to specify a value with the WhatIf switch.
#
# .EXAMPLE
#    .\Sync-MailPublicFolders.ps1 -Credential (Get-Credential) -CsvSummaryFile:sync_summary.csv
#
#    This example shows how to sync mail-public folders from your local deployment to Exchange Online. Note that the script outputs a CSV file listing all operations executed, and possibly errors encountered, during sync.
#
# .EXAMPLE
#    .\Sync-MailPublicFolders.ps1 -Credential (Get-Credential) -CsvSummaryFile:sync_summary.csv -ConnectionUri:"https://partner.outlook.cn/PowerShell"
#
#    This example shows how to use a different URI to connect to Exchange Online and sync mail-public folders from your local deployment.
#
param(
    [Parameter()]
    [System.Management.Automation.PSCredential] $Credential,

    <#     [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string] $CsvSummaryFile, #>

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string] $OutputFolderPath,

    [Parameter()]
    [bool] $Confirm = $true,

    [Parameter()]
    [switch] $Force,

    [Parameter()]
    [switch] $WhatIf
)

# Writes a dated information message to console
function WriteInfoMessage()
{
    param ($message)
    Write-Host "[$($(Get-Date).ToString())]" $message;
}

# Writes a dated warning message to console
function WriteWarningMessage()
{
    param ($message)
    Write-Warning ("[{0}] {1}" -f (Get-Date), $message);
}

# Writes a verbose message to console
function WriteVerboseMessage()
{
    param ($message)
    Write-Host "[VERBOSE] $message" -ForegroundColor Green -BackgroundColor Black;
}

# Writes an error importing a mail public folder to the CSV summary
function WriteErrorSummary()
{
    param ($folder, $operation, $errorMessage, $commandtext)

    WriteOperationSummary $folder.Guid $operation $errorMessage $commandtext;
    $script:errorsEncountered++;
}

# Writes the operation executed and its result to the output CSV
function WriteOperationSummary()
{
    param ($folder, $operation, $result, $commandtext)

    $columns = @(
        (Get-Date).ToString(),
        $folder.Guid,
        $operation,
        (EscapeCsvColumn $result),
        (EscapeCsvColumn $commandtext)
    );

    Add-Content $CsvSummaryFile -Value ("{0},{1},{2},{3},{4}" -f $columns);
}

#Escapes a column value based on RFC 4180 (http://tools.ietf.org/html/rfc4180)
function EscapeCsvColumn()
{
    param ([string]$text)

    if ($text -eq $null)
    {
        return $text;
    }

    $hasSpecial = $false;
    for ($i = 0; $i -lt $text.Length; $i++)
    {
        $c = $text[$i];
        if ($c -eq $script:csvEscapeChar -or
            $c -eq $script:csvFieldDelimiter -or
            $script:csvSpecialChars -contains $c)
        {
            $hasSpecial = $true;
            break;
        }
    }

    if (-not $hasSpecial)
    {
        return $text;
    }

    $ch = $script:csvEscapeChar.ToString([System.Globalization.CultureInfo]::InvariantCulture);
    return $ch + $text.Replace($ch, $ch + $ch) + $ch;
}

# Writes the current progress
function WriteProgress()
{
    param($statusFormat, $statusProcessed, $statusTotal)
    Write-Progress -Activity $LocalizedStrings.ProgressBarActivity `
        -Status ($statusFormat -f $statusProcessed, $statusTotal) `
        -PercentComplete (100 * ($script:itemsProcessed + $statusProcessed) / $script:totalItems);
}

# Create a tenant PSSession against Exchange Online.
function InitializeExchangeOnlineRemoteSession()
{
    WriteInfoMessage $LocalizedStrings.CreatingRemoteSession;

    $oldWarningPreference = $WarningPreference;
    $oldVerbosePreference = $VerbosePreference;

    try
    {
        $VerbosePreference = $WarningPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue;
        <#         $sessionOption = (New-PSSessionOption -SkipCACheck);
        $script:session = New-PSSession -ConnectionURI:$ConnectionUri `
            -ConfigurationName:Microsoft.Exchange `
            -AllowRedirection `
            -Authentication:"Basic" `
            -SessionOption:$sessionOption `
            -Credential:$Credential `
            -ErrorAction:SilentlyContinue;

        if ($script:session -eq $null)
        {
            Write-Error ($LocalizedStrings.FailedToCreateRemoteSession -f $error[0].Exception.Message);
            Exit;
        }
        else
        {
            $result = Import-PSSession -Session $script:session `
                -Prefix "EXO" `
                -AllowClobber;

            if (-not $?)
            {
                Write-Error ($LocalizedStrings.FailedToImportRemoteSession -f $error[0].Exception.Message);
                Remove-PSSession $script:session;
                Exit;
            }
        }#>
        Connect-ExchangeOnline -Prefix 'PFS' -ErrorAction Stop
    }
    finally
    {
        $WarningPreference = $oldWarningPreference;
        $VerbosePreference = $oldVerbosePreference;
    }

    WriteInfoMessage $LocalizedStrings.RemoteSessionCreatedSuccessfully;
}

# Invokes New-SyncMailPublicFolder to create a new MEPF object on AD
function NewMailEnabledPublicFolder()
{
    param ($localFolder)

    if ($localFolder.PrimarySmtpAddress.ToString() -eq "")
    {
        $errorMsg = ($LocalizedStrings.FailedToCreateMailPublicFolderEmptyPrimarySmtpAddress -f $localFolder.Guid);
        Write-Error $errorMsg;
        WriteErrorSummary $localFolder $LocalizedStrings.CreateOperationName $errorMsg "";
        return;
    }

    # preserve the ability to reply via Outlook's nickname cache post-migration
    $emailAddressesArray = $localFolder.EmailAddresses.ToStringArray() + ("x500:" + $localFolder.LegacyExchangeDN);

    $newParams = @{};
    AddNewOrSetCommonParameters $localFolder $emailAddressesArray $newParams;

    [string]$commandText = (FormatCommand $script:NewSyncMailPublicFolderCommand $newParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }

    try
    {
        $result = &$script:NewSyncMailPublicFolderCommand @newParams;
        WriteOperationSummary $localFolder $LocalizedStrings.CreateOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsCreated++;
        }
    }
    catch
    {
        WriteErrorSummary $localFolder $LocalizedStrings.CreateOperationName $error[0].Exception.Message $commandText;
        Write-Error $_;
    }
}

# Invokes Remove-SyncMailPublicFolder to remove a MEPF from AD
function RemoveMailEnabledPublicFolder()
{
    param ($remoteFolder)

    $removeParams = @{};
    $removeParams.Add("Identity", $remoteFolder.DistinguishedName);
    $removeParams.Add("Confirm", $false);
    $removeParams.Add("WarningAction", [System.Management.Automation.ActionPreference]::SilentlyContinue);
    $removeParams.Add("ErrorAction", [System.Management.Automation.ActionPreference]::Stop);

    if ($WhatIf)
    {
        $removeParams.Add("WhatIf", $true);
    }

    [string]$commandText = (FormatCommand $script:RemoveSyncMailPublicFolderCommand $removeParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }

    try
    {
        &$script:RemoveSyncMailPublicFolderCommand @removeParams;
        WriteOperationSummary $remoteFolder $LocalizedStrings.RemoveOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsDeleted++;
        }
    }
    catch
    {
        WriteErrorSummary $remoteFolder $LocalizedStrings.RemoveOperationName $_.Exception.Message $commandText;
        Write-Error $_;
    }
}

# Invokes Set-MailPublicFolder to update the properties of an existing MEPF
function UpdateMailEnabledPublicFolder()
{
    param ($localFolder, $remoteFolder)

    $localEmailAddresses = $localFolder.EmailAddresses.ToStringArray();
    $localEmailAddresses += ("x500:" + $localFolder.LegacyExchangeDN); # preserve the ability to reply via Outlook's nickname cache post-migration
    $emailAddresses = ConsolidateEmailAddresses $localEmailAddresses $remoteFolder.EmailAddresses $remoteFolder.LegacyExchangeDN;

    $setParams = @{};
    $setParams.Add("Identity", $remoteFolder.DistinguishedName);

    if ($script:mailEnabledSystemFolders.Contains($localFolder.Guid))
    {
        $setParams.Add("IgnoreMissingFolderLink", $true);
    }

    AddNewOrSetCommonParameters $localFolder $emailAddresses $setParams;

    [string]$commandText = (FormatCommand $script:SetMailPublicFolderCommand $setParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }

    try
    {
        &$script:SetMailPublicFolderCommand @setParams;
        WriteOperationSummary $remoteFolder $LocalizedStrings.UpdateOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsUpdated++;
        }
    }
    catch
    {
        WriteErrorSummary $remoteFolder $LocalizedStrings.UpdateOperationName $_.Exception.Message $commandText;
        Write-Error $_;
    }
}

# Adds the common set of parameters between New and Set cmdlets to the given dictionary
function AddNewOrSetCommonParameters()
{
    param ($localFolder, $emailAddresses, [System.Collections.IDictionary]$parameters)

    $windowsEmailAddress = $localFolder.WindowsEmailAddress.ToString();
    if ($windowsEmailAddress -eq "")
    {
        $windowsEmailAddress = $localFolder.PrimarySmtpAddress.ToString();
    }

    $parameters.Add("Alias", $localFolder.Alias.Trim());
    $parameters.Add("DisplayName", $localFolder.DisplayName.Trim());
    $parameters.Add("EmailAddresses", $emailAddresses);
    $parameters.Add("ExternalEmailAddress", $localFolder.PrimarySmtpAddress.ToString());
    $parameters.Add("HiddenFromAddressListsEnabled", $localFolder.HiddenFromAddressListsEnabled);
    $parameters.Add("Name", $localFolder.Name.Trim());
    $parameters.Add("OnPremisesObjectId", $localFolder.Guid);
    $parameters.Add("WindowsEmailAddress", $windowsEmailAddress);
    $parameters.Add("ErrorAction", [System.Management.Automation.ActionPreference]::Stop);

    if ($WhatIf)
    {
        $parameters.Add("WhatIf", $true);
    }
}

# Finds out the cloud-only email addresses and merges those with the values current persisted in the on-premises object
function ConsolidateEmailAddresses()
{
    param($localEmailAddresses, $remoteEmailAddresses, $remoteLegDN)

    # Check if the email address in the existing cloud object is present on-premises; if it is not, then the address was either:
    # 1. Deleted on-premises and must be removed from cloud
    # 2. or it is a cloud-authoritative address and should be kept
    $remoteAuthoritative = @();
    foreach ($remoteAddress in $remoteEmailAddresses)
    {
        if ($remoteAddress.StartsWith("SMTP:", [StringComparison]::InvariantCultureIgnoreCase))
        {
            $found = $false;
            $remoteAddressParts = $remoteAddress.Split($script:proxyAddressSeparators); # e.g. SMTP:alias@domain
            if ($remoteAddressParts.Length -ne 3)
            {
                continue; # Invalid SMTP proxy address (it will be removed)
            }

            foreach ($localAddress in $localEmailAddresses)
            {
                # note that the domain part of email addresses is case insensitive while the alias part is case sensitive
                $localAddressParts = $localAddress.Split($script:proxyAddressSeparators);
                if ($localAddressParts.Length -eq 3 -and
                    $remoteAddressParts[0].Equals($localAddressParts[0], [StringComparison]::InvariantCultureIgnoreCase) -and
                    $remoteAddressParts[1].Equals($localAddressParts[1], [StringComparison]::InvariantCulture) -and
                    $remoteAddressParts[2].Equals($localAddressParts[2], [StringComparison]::InvariantCultureIgnoreCase))
                {
                    $found = $true;
                    break;
                }
            }

            if (-not $found)
            {
                foreach ($domain in $script:authoritativeDomains)
                {
                    if ($remoteAddressParts[2] -eq $domain)
                    {
                        $found = $true;
                        break;
                    }
                }

                if (-not $found)
                {
                    # the address on the remote object is from a cloud authoritative domain and should not be removed
                    $remoteAuthoritative += $remoteAddress;
                }
            }
        }
        elseif ($remoteAddress.StartsWith("X500:", [StringComparison]::InvariantCultureIgnoreCase) -and
            $remoteAddress.Substring(5) -eq $remoteLegDN)
        {
            $remoteAuthoritative += $remoteAddress;
        }
    }

    return $localEmailAddresses + $remoteAuthoritative;
}

# Formats the command and its parameters to be printed on console or to file
function FormatCommand()
{
    param ([string]$command, [System.Collections.IDictionary]$parameters)

    $commandText = New-Object System.Text.StringBuilder;
    [void]$commandText.Append($command);
    foreach ($name in $parameters.Keys)
    {
        [void]$commandText.AppendFormat(" -{0}:", $name);

        $value = $parameters[$name];
        if ($value -isnot [Array])
        {
            [void]$commandText.AppendFormat("`"{0}`"", $value);
        }
        elseif ($value.Length -eq 0)
        {
            [void]$commandText.Append("@()");
        }
        else
        {
            [void]$commandText.Append("@(");
            foreach ($subValue in $value)
            {
                [void]$commandText.AppendFormat("`"{0}`",", $subValue);
            }

            [void]$commandText.Remove($commandText.Length - 1, 1);
            [void]$commandText.Append(")");
        }
    }

    return $commandText.ToString();
}

################ DECLARING GLOBAL VARIABLES ################
$script:session = $null;
$script:verbose = $VerbosePreference -eq [System.Management.Automation.ActionPreference]::Continue;

$script:csvSpecialChars = @("`r", "`n");
$script:csvEscapeChar = '"';
$script:csvFieldDelimiter = ',';

$script:ObjectsCreated = $script:ObjectsUpdated = $script:ObjectsDeleted = 0;
$script:NewSyncMailPublicFolderCommand = "New-PFSSyncMailPublicFolder";
$script:SetMailPublicFolderCommand = "Set-PFSMailPublicFolder";
$script:RemoveSyncMailPublicFolderCommand = "Remove-PFSSyncMailPublicFolder";
[char[]]$script:proxyAddressSeparators = ':', '@';
$script:errorsEncountered = 0;
$script:authoritativeDomains = $null;
$script:mailEnabledSystemFolders = New-Object 'System.Collections.Generic.HashSet[Guid]';
$script:WellKnownSystemFolders = @(
    "\NON_IPM_SUBTREE\EFORMS REGISTRY",
    "\NON_IPM_SUBTREE\OFFLINE ADDRESS BOOK",
    "\NON_IPM_SUBTREE\SCHEDULE+ FREE BUSY",
    "\NON_IPM_SUBTREE\schema-root",
    "\NON_IPM_SUBTREE\Events Root");

#load hashtable of localized string
Import-LocalizedData -BindingVariable LocalizedStrings -FileName SyncMailPublicFolders.strings.psd1

#minimum supported exchange version to run this script
$minSupportedVersion = 8
################ END OF DECLARATION #################
$CsvSummaryFile = Join-Path -Path $OutputFolderPath -ChildPath 'SyncMailPublicFoldersSummary.csv'

if (Test-Path $CsvSummaryFile)
{
    Remove-Item $CsvSummaryFile -Confirm:$Confirm -Force;
}

# Write the output CSV headers
$csvFile = New-Item -Path $CsvSummaryFile -ItemType File -Force -ErrorAction:Stop -Value ("#{0},{1},{2},{3},{4}`r`n" -f $LocalizedStrings.TimestampCsvHeader,
    $LocalizedStrings.IdentityCsvHeader,
    $LocalizedStrings.OperationCsvHeader,
    $LocalizedStrings.ResultCsvHeader,
    $LocalizedStrings.CommandCsvHeader);

<# $localServerVersion = (Get-ExchangeServer $env:COMPUTERNAME -ErrorAction:Stop).AdminDisplayVersion;
# This script can run from Exchange 2007 Management shell and above
if ($localServerVersion.Major -lt $minSupportedVersion)
{
    Write-Error ($LocalizedStrings.LocalServerVersionNotSupported -f $localServerVersion) -ErrorAction:Continue;
    Exit;
} #>

try
{
    InitializeExchangeOnlineRemoteSession;

    WriteInfoMessage $LocalizedStrings.LocalMailPublicFolderEnumerationStart;

    # During finalization, Public Folders deployment is locked for migration, which means the script cannot invoke
    # Get-PublicFolder as that operation would fail. In that case, the script cannot determine which mail public folder
    # objects are linked to system folders under the NON_IPM_SUBTREE.
    $lockedForMigration = (Get-OrganizationConfig).PublicFoldersLockedForMigration;
    $allSystemFoldersInAD = @();
    if (-not $lockedForMigration)
    {
        # See https://technet.microsoft.com/en-us/library/bb397221(v=exchg.141).aspx#Trees
        # Certain WellKnownFolders in pre-E15 are created with prefix such as OWAScratchPad, StoreEvents.
        # For instance, StoreEvents folders have the following pattern: "\NON_IPM_SUBTREE\StoreEvents{46F83CF7-2A81-42AC-A0C6-68C7AA49FF18}\internal1"
        $storeEventAndOwaScratchPadFolders = @(Get-PublicFolder \NON_IPM_SUBTREE -GetChildren -ResultSize:Unlimited | Where-Object { $_.Name -like "StoreEvents*" -or $_.Name -like "OWAScratchPad*" });
        $allSystemFolderParents = $storeEventAndOwaScratchPadFolders + @($script:WellKnownSystemFolders | Get-PublicFolder -ErrorAction:SilentlyContinue);
        $allSystemFoldersInAD = @($allSystemFolderParents | Get-PublicFolder -Recurse -ResultSize:Unlimited | Get-MailPublicFolder -ErrorAction:SilentlyContinue);

        foreach ($systemFolder in $allSystemFoldersInAD)
        {
            [void]$script:mailEnabledSystemFolders.Add($systemFolder.Guid);
        }
    }
    else
    {
        WriteWarningMessage $LocalizedStrings.UnableToDetectSystemMailPublicFolders;
    }

    if ($script:verbose)
    {
        WriteVerboseMessage ($LocalizedStrings.SystemFoldersSkipped -f $script:mailEnabledSystemFolders.Count);
        $allSystemFoldersInAD | Sort-Object Alias | Format-Table -a | Out-String | Write-Host -ForegroundColor Green -BackgroundColor Black;
    }

    $localFolders = @(Get-MailPublicFolder -ResultSize:Unlimited -IgnoreDefaultScope | Sort-Object Guid);
    WriteInfoMessage ($LocalizedStrings.LocalMailPublicFolderEnumerationCompleted -f $localFolders.Length);

    if ($localFolders.Length -eq 0 -and $Force -eq $false)
    {
        WriteWarningMessage $LocalizedStrings.ForceParameterRequired;
        Exit;
    }

    WriteInfoMessage $LocalizedStrings.RemoteMailPublicFolderEnumerationStart;
    $remoteFolders = @(Get-PFSMailPublicFolder -ResultSize:Unlimited | Sort-Object OnPremisesObjectId);
    WriteInfoMessage ($LocalizedStrings.RemoteMailPublicFolderEnumerationCompleted -f $remoteFolders.Length);

    $missingOnPremisesGuid = @();
    $pendingRemoves = @();
    $pendingUpdates = @{};
    $pendingAdds = @{};

    $localIndex = 0;
    $remoteIndex = 0;
    while ($localIndex -lt $localFolders.Length -and $remoteIndex -lt $remoteFolders.Length)
    {
        $local = $localFolders[$localIndex];
        $remote = $remoteFolders[$remoteIndex];

        if ($remote.OnPremisesObjectId -eq "")
        {
            # This folder must be processed based on PrimarySmtpAddress
            $missingOnPremisesGuid += $remote;
            $remoteIndex++;
        }
        elseif ($local.Guid.ToString() -eq $remote.OnPremisesObjectId)
        {
            $pendingUpdates.Add($local.Guid, (New-Object PSObject -Property @{ Local = $local; Remote = $remote }));
            $localIndex++;
            $remoteIndex++;
        }
        elseif ($local.Guid.ToString() -lt $remote.OnPremisesObjectId)
        {
            if (-not $script:mailEnabledSystemFolders.Contains($local.Guid))
            {
                $pendingAdds.Add($local.Guid, $local);
            }

            $localIndex++;
        }
        else
        {
            $pendingRemoves += $remote;
            $remoteIndex++;
        }
    }

    # Remaining folders on $localFolders collection must be added to Exchange Online
    while ($localIndex -lt $localFolders.Length)
    {
        $local = $localFolders[$localIndex];

        if (-not $script:mailEnabledSystemFolders.Contains($local.Guid))
        {
            $pendingAdds.Add($local.Guid, $local);
        }

        $localIndex++;
    }

    # Remaining folders on $remoteFolders collection must be removed from Exchange Online
    while ($remoteIndex -lt $remoteFolders.Length)
    {
        $remote = $remoteFolders[$remoteIndex];
        if ($remote.OnPremisesObjectId -eq "")
        {
            # This folder must be processed based on PrimarySmtpAddress
            $missingOnPremisesGuid += $remote;
        }
        else
        {
            $pendingRemoves += $remote;
        }

        $remoteIndex++;
    }

    if ($missingOnPremisesGuid.Length -gt 0)
    {
        # Process remote objects missing the OnPremisesObjectId using the PrimarySmtpAddress as a key instead.
        $missingOnPremisesGuid = @($missingOnPremisesGuid | Sort-Object PrimarySmtpAddress);
        $localFolders = @($localFolders | Sort-Object PrimarySmtpAddress);

        $localIndex = 0;
        $remoteIndex = 0;
        while ($localIndex -lt $localFolders.Length -and $remoteIndex -lt $missingOnPremisesGuid.Length)
        {
            $local = $localFolders[$localIndex];
            $remote = $missingOnPremisesGuid[$remoteIndex];

            if ($local.PrimarySmtpAddress.ToString() -eq $remote.PrimarySmtpAddress.ToString())
            {
                # Make sure the PrimarySmtpAddress has no duplicate on-premises; otherwise, skip updating all objects with duplicate address
                $j = $localIndex + 1;
                while ($j -lt $localFolders.Length)
                {
                    $next = $localFolders[$j];
                    if ($local.PrimarySmtpAddress.ToString() -ne $next.PrimarySmtpAddress.ToString())
                    {
                        break;
                    }

                    WriteErrorSummary $next $LocalizedStrings.UpdateOperationName ($LocalizedStrings.PrimarySmtpAddressUsedByAnotherFolder -f $local.PrimarySmtpAddress, $local.Guid) "";

                    # If there were a previous match based on OnPremisesObjectId, remove the folder operation from add and update collections
                    $pendingAdds.Remove($next.Guid);
                    $pendingUpdates.Remove($next.Guid);
                    $j++;
                }

                $duplicatesFound = $j - $localIndex - 1;
                if ($duplicatesFound -gt 0)
                {
                    # If there were a previous match based on OnPremisesObjectId, remove the folder operation from add and update collections
                    $pendingAdds.Remove($local.Guid);
                    $pendingUpdates.Remove($local.Guid);
                    $localIndex += $duplicatesFound + 1;

                    WriteErrorSummary $local $LocalizedStrings.UpdateOperationName ($LocalizedStrings.PrimarySmtpAddressUsedByOtherFolders -f $local.PrimarySmtpAddress, $duplicatesFound) "";
                    WriteWarningMessage ($LocalizedStrings.SkippingFoldersWithDuplicateAddress -f ($duplicatesFound + 1), $local.PrimarySmtpAddress);
                }
                elseif ($pendingUpdates.Contains($local.Guid))
                {
                    # If we get here, it means two different remote objects match the same local object (one by OnPremisesObjectId and another by PrimarySmtpAddress).
                    # Since that is an ambiguous resolution, let's skip updating the remote objects.
                    $ambiguousRemoteObj = $pendingUpdates[$local.Guid].Remote;
                    $pendingUpdates.Remove($local.Guid);

                    $errorMessage = ($LocalizedStrings.AmbiguousLocalMailPublicFolderResolution -f $local.Guid, $ambiguousRemoteObj.Guid, $remote.Guid);
                    WriteErrorSummary $local $LocalizedStrings.UpdateOperationName $errorMessage "";
                    WriteWarningMessage $errorMessage;
                }
                else
                {
                    # Since there was no match originally using OnPremisesObjectId, the local object was treated as an add to Exchange Online.
                    # In this way, since we now found a remote object (by PrimarySmtpAddress) to update, we must first remove the local object from the add list.
                    $pendingAdds.Remove($local.Guid);
                    $pendingUpdates.Add($local.Guid, (New-Object PSObject -Property @{ Local = $local; Remote = $remote }));
                }

                $localIndex++;
                $remoteIndex++;
            }
            elseif ($local.PrimarySmtpAddress.ToString() -gt $remote.PrimarySmtpAddress.ToString())
            {
                # There are no local objects using the remote object's PrimarySmtpAddress
                $pendingRemoves += $remote;
                $remoteIndex++;
            }
            else
            {
                $localIndex++;
            }
        }

        # All objects remaining on the $missingOnPremisesGuid list no longer exist on-premises
        while ($remoteIndex -lt $missingOnPremisesGuid.Length)
        {
            $pendingRemoves += $missingOnPremisesGuid[$remoteIndex];
            $remoteIndex++;
        }
    }

    $script:totalItems = $pendingRemoves.Length + $pendingUpdates.Count + $pendingAdds.Count;

    $PendingAddsFilePath = Join-Path -Path $OutputFolderPath -ChildPath 'SyncMailPublicFoldersPendingAdds.xml'
    $pendingAdds | Export-Clixml -Path $PendingAddsFilePath -Force

    $PendingRemovesFilePath = Join-Path -Path $OutputFolderPath -ChildPath 'SyncMailPublicFoldersPendingRemoves.xml'
    $pendingRemoves | Export-Clixml -Path $PendingRemovesFilePath -Force

    $PendingUpdatesFilePath = Join-Path -Path $OutputFolderPath -ChildPath 'SyncMailPublicFoldersPendingUpdates.xml'
    $pendingUpdates | Export-Clixml -Path $PendingUpdatesFilePath -Force

    # At this point, we know all changes that need to be synced to Exchange Online. Let's prompt the admin for confirmation before proceeding.
    if ($Confirm -eq $true -and $script:totalItems -gt 0)
    {
        $title = $LocalizedStrings.ConfirmationTitle;
        $message = ($LocalizedStrings.ConfirmationQuestion -f $pendingAdds.Count, $pendingUpdates.Count, $pendingRemoves.Length);
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription $LocalizedStrings.ConfirmationYesOption, `
            $LocalizedStrings.ConfirmationYesOptionHelp;

        $no = New-Object System.Management.Automation.Host.ChoiceDescription $LocalizedStrings.ConfirmationNoOption, `
            $LocalizedStrings.ConfirmationNoOptionHelp;

        [System.Management.Automation.Host.ChoiceDescription[]]$options = $no, $yes;
        $confirmation = $host.ui.PromptForChoice($title, $message, $options, 0);
        if ($confirmation -eq 0)
        {
            Exit;
        }
    }

    # Find out the authoritative AcceptedDomains on-premises so that we don't accidently remove cloud-only email addresses during updates
    $script:authoritativeDomains = @(Get-AcceptedDomain | Where-Object { $_.DomainType -eq "Authoritative" } | ForEach-Object { $_.DomainName.ToString() });

    # Finally, let's perfom the actual operations against Exchange Online
    $script:itemsProcessed = 0;
    for ($i = 0; $i -lt $pendingRemoves.Length; $i++)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusRemoving $i $pendingRemoves.Length;
        RemoveMailEnabledPublicFolder $pendingRemoves[$i];
    }

    $script:itemsProcessed += $pendingRemoves.Length;
    $updatesProcessed = 0;
    foreach ($folderPair in $pendingUpdates.Values)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusUpdating $updatesProcessed $pendingUpdates.Count;
        UpdateMailEnabledPublicFolder $folderPair.Local $folderPair.Remote;
        $updatesProcessed++;
    }

    $script:itemsProcessed += $pendingUpdates.Count;
    $addsProcessed = 0;
    foreach ($localFolder in $pendingAdds.Values)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusCreating $addsProcessed $pendingAdds.Count;
        NewMailEnabledPublicFolder $localFolder;
        $addsProcessed++;
    }

    Write-Progress -Activity $LocalizedStrings.ProgressBarActivity -Status ($LocalizedStrings.ProgressBarStatusCreating -f $pendingAdds.Count, $pendingAdds.Count) -Completed;
    WriteInfoMessage ($LocalizedStrings.SyncMailPublicFolderObjectsComplete -f $script:ObjectsCreated, $script:ObjectsUpdated, $script:ObjectsDeleted);

    if ($script:errorsEncountered -gt 0)
    {
        WriteWarningMessage ($LocalizedStrings.ErrorsFoundDuringImport -f $script:errorsEncountered, (Get-Item $CsvSummaryFile).FullName);
    }
}
finally
{
    if ($script:session -ne $null)
    {
        Remove-PSSession $script:session;
    }
}