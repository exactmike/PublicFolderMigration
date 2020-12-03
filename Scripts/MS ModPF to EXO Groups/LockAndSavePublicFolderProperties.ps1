# .SYNOPSIS
# LockAndSavePublicFolderProperties.ps1
#    It locks down the public folders which are being migrated to Groups, and back up
#    the properties in case the user want to fallback to public folders.
#
# .DESCRIPTION
#    Script performs the following actions
#      1. It saves PublicFolderClientPermissions of migrating public folders.
#      2. It reads permission of each user and assign back the permissions except
#         create, edit or delete permissions.
#      3. Mail-disable and save mail properties of any mail enabled public folder which gets migrated.
#      4. Add smtp addresses of mail public folders to the Proxy Address list of Groups to 
#         which each public folder gets migrated.
#
#    After the execution of the script
#      1. All migrating public folders will be read-only to those users who had access to public folder content.
#      2. Any user who didn't had read permission will not gain read permission by lockdown.
#      3. Any mails sent to mail enabled public folder will be routed to target group.
#
# .PARAMETER MappingCsv
#    The public folder to group mapping csv file which was provided for the migration batch.
#
# .PARAMETER BackupDir
#    The directory to which user want to save the permissions and other properties as backup files.
#
# .PARAMETER ArePublicFoldersOnPremises
#    Tells if public folders are on-premises. Set to '$true' if public folders are remote, else set to '$false'.
#
# .PARAMETER Credential
#    Exchange Online user name and password.
#
# .PARAMETER ConnectionUri
#    The Exchange Online remote PowerShell connection uri. If you are an Office 365 operated by 21Vianet customer in China, use "https://partner.outlook.cn/PowerShell".
#
# .PARAMETER WhatIf
#    The WhatIf switch instructs the script to simulate the actions that it would take on the object. By using the WhatIf switch, you can view what changes would occur
#    without having to apply any of those changes. You don't have to specify a value with the WhatIf switch.
#
# .EXAMPLE
#    .\LockAndSavePublicFolderProperties.ps1 -MappingCsv .\map.csv -BackupDir C:\PFToGroupMigration\ -WhatIf
#    .\LockAndSavePublicFolderProperties.ps1 -MappingCsv .\map.csv -BackupDir C:\PFToGroupMigration\ -ArePublicFoldersOnPremises $true -Credential (Get-Credential)
#
# Copyright (c) 2017 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.


Param(

    [Parameter(Mandatory=$true, HelpMessage="The input csv used to create migration batch")]
    [ValidateNotNullOrEmpty()]
    [string] $MappingCsv,

    [Parameter(Mandatory=$true, HelpMessage="Choose directory to backup current public folder permissions")]
    [ValidateNotNullOrEmpty()]
    [string] $BackupDir,

    [Parameter(Mandatory = $false, HelpMessage = "Enter '`$true' if public folders are on-premises)")]
    [ValidateNotNullOrEmpty()]
    [bool] $ArePublicFoldersOnPremises = $false,

    [Parameter(Mandatory=$false, HelpMessage = "Enter the Exchange Online admin credential")]
    [System.Management.Automation.PSCredential] $Credential,

    [Parameter(Mandatory=$false, HelpMessage = "Enter the Exchange Online remote PowerShell connection uri")]
    [ValidateNotNullOrEmpty()]
    [string] $ConnectionUri = "https://outlook.office365.com/powerShell-liveID",

    [Parameter(Mandatory=$false)]
    [switch] $WhatIf = $false
)

###################### START OF DEFAULTS ######################

if ($WhatIf)
{
   $prefix = "test_"; 
}
else
{
   $prefix = $null; 
}

$permissionListFile = Join-Path $BackupDir ($prefix + "PfPermissions.csv");
$pfToGrpMappingCsv  = Join-Path $BackupDir ($prefix + "PfToGroupMapping.csv");
$pfMailPropCsv      = Join-Path $BackupDir ($prefix + "PfMailProperties.csv");
$logPath            = Join-Path $BackupDir ($prefix + "PfLockdown_summary.log");

$updateRolesLookupForLockdown = @{`
                                     "None"              = "None";`
                                     "AvailabilityOnly"  = "AvailabilityOnly";`
                                     "LimitedDetails"    = "LimitedDetails";`
                                     "Contributor"       = "FolderVisible";`
                                     "Reviewer"          = "ReadItems", "FolderVisible";`
                                     "NonEditingAuthor"  = "ReadItems", "FolderVisible";`
                                     "Author"            = "ReadItems", "FolderVisible";`
                                     "Editor"            = "ReadItems", "FolderVisible";`
                                     "PublishingAuthor"  = "ReadItems", "CreateSubfolders", "FolderVisible";`
                                     "PublishingEditor"  = "ReadItems", "CreateSubfolders", "FolderVisible";`
                                     "Owner"             = "ReadItems", "CreateSubfolders", "FolderContact", "FolderVisible";`
                                 }

$allowedListOfPermissionsForCustomRole = "ReadItems", "CreateSubfolders", "FolderContact", "FolderVisible";

###################### END OF DEFAULTS ######################

# Load necessary helper functions
. ".\WriteLog.ps1"
. ".\RetryScriptBlock.ps1"

# Load localized strings
Import-LocalizedData -BindingVariable LocalizedStrings -FileName LockAndSavePublicFolderProperties.strings.psd1

# Function to copy email addresses to group
function SetEmailIdsToGroup
{
    Param ($targetGroupMailbox, $pfEmailIds, $sendOnBehalfTo);

    $Error[0] = $null;
    if ($ArePublicFoldersOnPremises)
    {
        if ($sendOnBehalfTo)
        {
            Set-EXOUnifiedGroup $targetGroupMailbox -EmailAddresses @{Add=$pfEmailIds} -GrantSendOnBehalfTo @{Add=$sendOnBehalfTo} 2> $null;
        }
        else
        {
            Set-EXOUnifiedGroup $targetGroupMailbox -EmailAddresses @{Add=$pfEmailIds} 2> $null;
        }
    }
    else
    {
        if ($sendOnBehalfTo)
        {
            Set-UnifiedGroup $targetGroupMailbox -EmailAddresses @{Add=$pfEmailIds} -GrantSendOnBehalfTo @{Add=$sendOnBehalfTo} 2> $null;
        }
        else
        {
            Set-UnifiedGroup $targetGroupMailbox -EmailAddresses @{Add=$pfEmailIds} 2> $null;
        }
    }

    if ($Error[0])
    {
        return $false;
    }
    else
    {
        return $true;
    }
}

# Function to add SendAs permissions of public folder to group
function AddSendAsPermissionToGroup
{
    Param ($groupId, $trustee)

    $Error[0] = $null;
    if ($ArePublicFoldersOnPremises)
    {
        Add-EXORecipientPermission -Identity $groupId -Trustee $trustee -AccessRights SendAs -ErrorAction SilentlyContinue -confirm:$false;
    }
    else
    {
        Add-RecipientPermission -Identity $groupId -Trustee $trustee -AccessRights SendAs -ErrorAction SilentlyContinue -confirm:$false;
    }

    if ($Error[0])
    {
        return $false;
    }
    else
    {
        return $true;
    }
}

# Create a tenant PSSession against Exchange Online.
function InitializeExchangeOnlineRemoteSession()
{
    $sessionOption = (New-PSSessionOption -SkipCACheck);
    $script:session = New-PSSession -ConnectionURI:$ConnectionUri `
            -ConfigurationName:Microsoft.Exchange `
            -AllowRedirection `
            -Authentication:"Basic" `
            -SessionOption:$sessionOption `
            -Credential:$Credential `
            -ErrorAction:SilentlyContinue;
        
    if ($script:session -eq $null)
    {
        WriteLog -Path $logPath -Level Error -Message ($LocalizedStrings.FailedToCreateRemoteSession -f $error[0].Exception.Message);
        Exit;
    }
    else
    {
        $result = Import-PSSession -Session $script:session `
                -Prefix "EXO" `
                -AllowClobber;

        if (!$?)
        {
            WriteLog -Path $logPath -Level Error -LogOnly -Message ($LocalizedStrings.FailedToImportRemoteSession -f $error[0].Exception.Message);
            Remove-PSSession $script:session;
            Exit;
        }
    }

    WriteLog -Path $logPath -Message $LocalizedStrings.RemoteSessionCreatedSuccessfully;
}

####################################################################################################
# Script starts here
####################################################################################################

if ($WhatIf)
{
    WriteLog -Path $logPath -Message $LocalizedStrings.WhatIfEnabled;
}

if ($ArePublicFoldersOnPremises)
{
    # E2010 Snap-in is added for Get-RecipientPermission cmdlet.
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;

    Get-RecipientPermission > $null;
    if (!$?)
    {
        # User may not have enough permissions to run Get-RecipientPermission cmdlet.
        WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.UnsuccessfulGetRecipientPermissionCmdlet);
        Remove-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;
        return;
    }
}

if ($ArePublicFoldersOnPremises -and ((Get-ExchangeServer $env:COMPUTERNAME -ErrorAction:Stop).AdminDisplayVersion.Major -eq 14))
{
    $isE14OnPrem = $true;
}
else
{
    $isE14OnPrem = $false;
}

# Checking the existence of csv file.
if (!(Test-Path $MappingCsv))
{
    WriteLog -Path $logPath -Level Error -Message ($LocalizedStrings.MappingCsvNotFound -f $MappingCsv);
    return;
}

$pfToGrpMapping = Import-Csv $MappingCsv;

# Checking expected columns in the csv provided.
$invalidRows = $pfToGrpMapping | ?{$_.FolderPath -eq $null -or $_.TargetGroupMailbox -eq $null};

if ($invalidRows)
{
    WriteLog -Path $logPath -Level Error -Message $LocalizedStrings.IncorrectCsvFormat;
    return;
}

# Script will exit if the backup files found.
# Preventing careless re-run of the script, as it will lead to loss of actual Permissions backup
# and creating a new backup file of lockdown permissions.
if ((Test-Path $permissionListFile) -or (Test-Path $pfToGrpMappingCsv) -or (Test-Path $pfMailPropCsv))
{
    WriteLog -Path $logPath -Level Error -Message $LocalizedStrings.BackupCsvAlreadyExist;
    return;
}

WriteLog -Path $logPath -Message $LocalizedStrings.ExportingPFPermissions;

# Getting public folder permissions
$pfsBeingMigrated = $pfToGrpMapping | %{$_.FolderPath};
if ($isE14OnPrem)
{
    # ADRecipient object is not available in 2010. Hence get the user name from ActiveDirectoryIdentity in the user object. 
    $accessRights = $pfsBeingMigrated  | Get-PublicFolderClientPermission | Select Identity,
                                                                                     AccessRights,
                                                                                     User,
                                                                                     @{Name="Name"; Expression={$_.User.ActiveDirectoryIdentity.Name}};
}
else
{
    $accessRights = $pfsBeingMigrated  | Get-PublicFolderClientPermission | Select Identity,
                                                                                     AccessRights,
                                                                                     User,
                                                                                     @{Name="Name"; Expression={$_.User.ADRecipient.Name}},
                                                                                     @{Name="PrimarySmtpAddress"; Expression={$_.User.ADRecipient.PrimarySmtpAddress}};
}

# Checking if the public folders are already in locked down state.
# If there's a permission for any public folder with Create/Update/Delete permission,
# we consider that the public folders are not in lockdown state.
$alreadyLockedDown = $true;

foreach ($accessRightItem in $accessRights)
{
    if ($updateRolesLookupForLockdown.ContainsKey([string]$accessRightItem.AccessRights))
    {
        if (!("None","Reviewer", "AvailabilityOnly","LimitedDetails" -contains [string]$accessRightItem.AccessRights))
        {
            $alreadyLockedDown = $false;
            break;
        }
    }
    else
    {
        # Finding if there is any create/update/delete permission
        $updateAccessRights = $accessRightItem.AccessRights | ?{$allowedListOfPermissionsForCustomRole -notcontains $_};

        if ($updateAccessRights)
        {
            $alreadyLockedDown = $false;
            break;
        }
    }
}

# If the public folders are already locked down, warn the user that the backup file is not found.
if ($alreadyLockedDown)
{
    WriteLog -Path $logPath -Message $LocalizedStrings.PfsAlreadyInLockedState;
    WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.WarnBackupFilesNotFound -f $permissionListFile); 
    return;
}

try
{ 
    # if public folders are on-premises, create an EXO remote session
    if ($ArePublicFoldersOnPremises)
    {
        # Check if EXO credentials are provide.
        if (!$Credential)
        {
            WriteLog -Path $logPath -Level Warn -Message $LocalizedStrings.CredentialNotFound;
            $Credential = Get-Credential;
        }

        WriteLog -Path $logPath -Message $LocalizedStrings.CreatingRemoteSession;
        InitializeExchangeOnlineRemoteSession;
    }

    # Exporting public folder Permissions to default backup location
    $accessRightsToExport = @();
    foreach ($accessRightItem in $accessRights)
    {
        $row = New-Object psobject -Property @{
                       Identity           = [string] $accessRightItem.Identity;
                       User               = [string] $accessRightItem.User;
                       AccessRights       = [string] $accessRightItem.AccessRights;
                       PrimarySmtpAddress = [string] $accessRightItem.PrimarySmtpAddress;
                       Name               = [string] $accessRightItem.Name; 
               }

        $accessRightsToExport += $row;
    }

    $accessRightsToExport | Export-Csv $permissionListFile -Encoding UTF8;
    WriteLog -Path $logPath -Message ($LocalizedStrings.ExportPermissionsSuccessful -f $PermissionListFile);

    # Exporting mail public folder Properties to default backup location
    $pfToGrpMapping.FolderPath | Get-MailPublicFolder -ErrorAction SilentlyContinue | Export-Csv $pfMailPropCsv -Encoding UTF8;
    WriteLog -Path $logPath -Message ($LocalizedStrings.ExportMailPropertiesSuccessful -f $pfMailPropCsv);

    $rows = @();
    foreach ($pfToGrpMappingItem in $pfToGrpMapping)
    {
        # Obtaining mail properties.
        $mailEnabledPf = (Get-MailPublicFolder $pfToGrpMappingItem.FolderPath -ErrorAction SilentlyContinue) | Select EmailAddresses,
                                                                                                                      ExternalEmailAddress,
                                                                                                                      EmailAddressPolicyEnabled,
                                                                                                                      GrantSendOnBehalfTo,
                                                                                                                      PrimarySmtpAddress;
        # If public folder is mail enabled
        if ($mailEnabledPf)
        {
            $pfEmailIds     = $mailEnabledPf.EmailAddresses;
            $extEmailAddr   = $mailEnabledPf.ExternalEmailAddress;
            $primarySmtpAddr= [string]$mailEnabledPf.PrimarySmtpAddress;
            $sendOnBehalfTo = $mailEnabledPf.GrantSendOnBehalfTo | %{[string](Get-Recipient $_).PrimarySmtpAddress};
            $sendAsList     = Get-RecipientPermission $primarySmtpAddr -ErrorAction SilentlyContinue | %{$_.Trustee} | Get-Recipient | %{[string]$_.PrimarySmtpAddress};
            $folderPath     = $pfToGrpMappingItem.FolderPath;
            $groupId        = $pfToGrpMappingItem.TargetGroupMailbox;

            # Converting type of primary smtp address from "SMTP" to "smtp" to avoid replacing 
            # group's primary smtp address.
            $pfEmailIds = $pfEmailIds | %{"smtp:" + ([string]$_).Split(":")[1]};
            if ($isE14OnPrem)
            {
                # We can't set the externalEmailAddress for mail-enabled publicfolder in E14
                $extEmailAddr = $null;
            }

            # SendAs and SendOnBehalfTo permissions that are not already present in group.
            if ($ArePublicFoldersOnPremises)
            {
                $sendAsOfGroup = Get-EXORecipientPermission $groupId | %{$_.Trustee} | Get-EXORecipient | %{$_.PrimarySmtpAddress};
                $sendAsAddedByScript = $sendAsList | ?{$sendAsOfGroup -notcontains $_};

                $sendOnBehalfToOfGroup = (Get-EXOUnifiedGroup $groupId).GrantSendOnBehalfTo | %{ Get-EXORecipient $_ -ErrorAction SilentlyContinue } | %{[string] $_.PrimarySmtpAddress};
                $sendOnBehalfToAddedByScript = $sendOnBehalfTo | ?{$sendOnBehalfToOfGroup -notcontains $_};

                $migratedUserListOfSendOnBehalfTo = @();
                foreach ($user in $sendOnBehalfToAddedByScript)
                {
                    $migratedUser = Get-EXORecipient $user -ErrorAction SilentlyContinue;
                    if (!$migratedUser)
                    {
                        WriteLog -Path $logPath -Level Warn ($LocalizedStrings.SkippingNotMigratedUser -f $user);
                    }
                    else
                    {
                        $migratedUserListOfSendOnBehalfTo += $user;
                    }
                }

                $sendOnBehalfToAddedByScript = $migratedUserListOfSendOnBehalfTo;

                $migratedUserListOfSendAs = @();
                foreach ($user in $sendAsAddedByScript)
                {
                    $migratedUser = Get-EXORecipient $user -ErrorAction SilentlyContinue;
                    if (!$migratedUser)
                    {
                        WriteLog -Path $logPath -Level Warn ($LocalizedStrings.SkippingNotMigratedUser -f $user);
                    }
                    else
                    {
                        $migratedUserListOfSendAs += $user;
                    }
                }

                $sendAsAddedByScript = $migratedUserListOfSendAs;
             }
            else
            {
                $sendAsOfGroup = ([string] (Get-RecipientPermission $groupId).Trustee).Split();
                $sendAsAddedByScript = $sendAsList | ?{$sendAsOfGroup -notcontains $_};

                $sendOnBehalfToOfGroup = ([string] (Get-UnifiedGroup $groupId).GrantSendOnBehalfTo).Split();
                $sendOnBehalfToAddedByScript = $sendOnBehalfTo | ?{$sendOnBehalfToOfGroup -notcontains $_};
            }

            # Mail-Disabling public folder
            if (!$WhatIf)
            {
                Disable-MailPublicFolder $folderPath -Confirm:$false;
            }

            WriteLog -Path $logPath -Message ($LocalizedStrings.PfMailDisabled -f $folderPath);

            # Retry loop for assigning email ids group
            if ($WhatIf)
            {
                WriteLog -Path $logPath -Message ($LocalizedStrings.PfPropertiesCopiedToGroup -f $folderPath, $groupId, [string]$pfEmailIds, [string]$sendOnBehalfToAddedByScript);
            }
            else
            {
                ExecuteWithRetries -ScriptToRetry:${function:SetEmailIdsToGroup} `
                                   -ArgumentList:@($groupId, $pfEmailIds, $sendOnBehalfToAddedByScript) `
                                   -ErrorMessageOnFailures:$LocalizedStrings.SettingSMTPToGroupFailed `
                                   -MessageIfSucceeded:($LocalizedStrings.SMTPAddressesCopiedFromMailPfToGroup -f $folderPath, $groupId) `
                                   -LogPath:$logPath;
            }

            # Retry loop for giving SendAs permission of mail-enabled public folder to corresponding group
            if ($sendAsAddedByScript)
            {
                if ($WhatIf)
                {
                    WriteLog -Path $logPath -Message ($LocalizedStrings.SendAsPermsCopiedToGroup -f $folderPath, $groupId, [string]$sendAsAddedByScript);
                }
                else
                {
                    $sendAsAddedByScript | %{`
                        ExecuteWithRetries -ScriptToRetry ${function:AddSendAsPermissionToGroup} `
                                           -ArgumentList @($groupId, $_) `
                                           -ErrorMessageOnFailures $LocalizedStrings.AddingSendAsToGroupFailed `
                                           -NumberOfRetries 3 `
                                           -LogPath $logPath; `
                    };
                }
            }
        }
        else
        {
            # Removing the value of any mail property variable, in case of a non-mail public folder
            $pfEmailIds                  = $null;
            $extEmailAddr                = $null;
            $sendOnBehalfTo              = $null;
            $pfEmailIdsStr               = $null;
            $sendAsList                  = $null;
            $sendOnBehalfToAddedByScript = $null;
            $sendAsAddedByScript         = $null;
        }

        $row = New-Object psobject -Property @{
                        Identity = $pfToGrpMappingItem.FolderPath;
                        EmailAddresses = $pfEmailIds -join " ";
                        UnifiedGroup = $pfToGrpMappingItem.TargetGroupMailbox;
                        ExternalEmailAddress = $extEmailAddr;
                        EmailAddressPolicyEnabled = $mailEnabledPf.EmailAddressPolicyEnabled;
                        GrantSendOnBehalfTo = $sendOnBehalfTo -join " ";
                        SendAsList = $sendAsList -join " ";
                        SendOnBehalfToAddedByScript = $sendOnBehalfToAddedByScript -join " ";
                        SendAsAddedByScript = $sendAsAddedByScript -join " ";
                }

        $rows += $row;
    }

    # Exporting PF Identity, Group Identity, Mail properties of PF to default location.
    $rows | Export-Csv $pfToGrpMappingCsv -Encoding UTF8;

    WriteLog -Path $logPath -Message ($LocalizedStrings.ExportMailPfPropertiesAndGroupSuccessful -f $pfToGrpMappingCsv);

    WriteLog -Path $logPath -Message $LocalizedStrings.LockingPfsByRemovingPerms;

    # We update the permissions of every custom role or predefined role in a way that
    # 1. No user will have create/update/delete permission on contents
    # 2. If the user didn't had access to read the contents, it stays same even after lockdown.
    # 3. Permissions assigned will be a subset of {"ReadItems", "CreateSubfolders", "FolderContact", "FolderVisible"}
    foreach ($accessRightItem in $accessRights)
    {
        if ($updateRolesLookupForLockdown.ContainsKey([string]$accessRightItem.AccessRights))
        {
            $newAccessRights = $updateRolesLookupForLockdown.Get_Item([string]$accessRightItem.AccessRights);
        }
        else
        {
            $newAccessRights = $accessRightItem.AccessRights | ?{ $allowedListOfPermissionsForCustomRole -contains $_};
            if (!($newAccessRights))
            {
                $newAccessRights = "None";
            }
        }

        $user = [string] $accessRightItem.User;
        $identity = $accessRightItem.Identity;

        # Checking if the user exists. 
        if ($user -ne "default" -and $user -ne "anonymous")
        {   
            $uniqueUser = $accessRightItem.Name;

            if (!$uniqueUser)
            {
                WriteLog -Path $logPath -Level Warn ($LocalizedStrings.SkippingUser -f $user);
                continue;
            }
        }
        else
        {
            $uniqueUser = $user;
        }

        if ($WhatIf)
        {
            WriteLog -Path $logPath -Message ($LocalizedStrings.RemovingPfPerm -f $user, $identity);
            WriteLog -Path $logPath -Message ($LocalizedStrings.AddingPfPerm -f $user, $identity, [string]$newAccessRights);
        }
        else
        {
            WriteLog -Path $logPath -LogOnly -Message ($LocalizedStrings.RemovingPfPerm -f $user, $identity);

            if ($isE14OnPrem)
            {
                # E14 demands the parameter AccessRights in Remove-PublicFolderClientPermission cmdlet
                Remove-PublicFolderClientPermission -Confirm:$false -Identity $identity -User $uniqueUser -AccessRights $accessRightItem.AccessRights;
            }
            else
            {
                # Later versions of Exchange don't have the parameter AccessRight in Remove-PublicFolderClientPermission cmdlet
                Remove-PublicFolderClientPermission -Confirm:$false -Identity $identity -User $uniqueUser;
            }

            WriteLog -Path $logPath -LogOnly -Message ($LocalizedStrings.AddingPfPerm -f $user, $identity, [string]$newAccessRights);
            Add-PublicFolderClientPermission -Identity $identity -User $uniqueUser -AccessRights $newAccessRights;
        }
    }
}
finally
{
    if ($script:session -ne $null)
    {
        Remove-PSSession $script:session;
    }

    if ($ArePublicFoldersOnPremises)
    {
        Remove-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;
    }
}

if (!$WhatIf)
{
    WriteLog -Path $logPath -Message $LocalizedStrings.LockdownWithReadOnlyPermsSuccessful;
    WriteLog -Path $logPath -Message $LocalizedStrings.PfLockdownComplete;
}
# SIG # Begin signature block
# MIIdugYJKoZIhvcNAQcCoIIdqzCCHacCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUgvpaFDaooavnpYH5pE4fSAul
# LXygghhTMIIEwjCCA6qgAwIBAgITMwAAAMM7uBDWq3WchAAAAAAAwzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwOTA3MTc1ODUx
# WhcNMTgwOTA3MTc1ODUxWjCBsjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046
# RDIzNi0zN0RBLTk3NjExJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
# cnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCiOG2wDGVZj5ZH
# gCl0ZaExy6HZQZ9T2uupPuxqtiWqXH2oIj762GqMc1JPYHkpEo5alygWdvB3D6FS
# qpA8al+mGJTMktlA+ydstLPRr6CBoEF+hm6RBzwVlsN9z6BVppwIZWt2lEVG6r1Y
# W1y1rb0d4FsA8qwRSI0sB8sAw9IHXi/J4Jd6klQvw2m6oLXl9C73/1DldPPZYGOT
# DQ98RxIaYewvksnNqblmvFpOx8Kuedkxl4jtAKl0F/2+QqRfU32OAiCiYFgZIgOP
# B4A8UbHmLIyn7pNqtom4NqMiZz9G4Bm5bwILhElYcZPMq/P1Hr38/WoAD99WAm3W
# FpXSFZejAgMBAAGjggEJMIIBBTAdBgNVHQ4EFgQUc3cXeGMQ8QV4IbaO4PEw84WH
# F6gwHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJ
# oEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMv
# TWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYB
# BQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
# 9w0BAQUFAAOCAQEASOPK1ntqWwaIWnNINY+LlmHQ4Q88h6TON0aE+6cZ2RrjBUU4
# 9STkyQ2lgvKpmIkQYWJbuNRh65IJ1HInwhD8XWd0f0JXAIrzTlL0zw3SdbrtyZ9s
# P4NxqyjQ23xBiI/d13CrtfTAVlGYIY1Ahl80+0KGyuUzJLTi9350/gHaI0Jz3irw
# rJ+htxF1UW/NT0AYJyRYe2el9JhgeudeKOKav3fQBlzALQmk4Ekoyq3muJHGoqfe
# No4zsP/M+WQ6oBMlUq8/49sg/ryuP0EeVtNiePuxPmX5i6Knzpd3rPgKPS+9Tq1d
# KLts1K4rjpASoKSs8Ubv3rwQSw0O/zTd1bc8EjCCBgAwggPooAMCAQICEzMAAADD
# Dpun2LLc9ywAAAAAAMMwDQYJKoZIhvcNAQELBQAwfjELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2ln
# bmluZyBQQ0EgMjAxMTAeFw0xNzA4MTEyMDIwMjRaFw0xODA4MTEyMDIwMjRaMHQx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xHjAcBgNVBAMTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBALtX1zjRsQZ/SS2pbbNjn3q6tjohW7SYro3UpIGgxXXFLO+CQCq3gVN382MB
# CrzON4QDQENXgkvO7R+2/YBtycKRXQXH3FZZAOEM61fe/fG4kCe/dUr8dbJyWLbF
# SJszYgXRlZSlvzkirY0STUZi2jIZzqoiXFZIsW9FyWd2Yl0wiKMvKMUfUCrZhtsa
# ESWBwvT1Zy7neR314hx19E7Mx/znvwuARyn/z81psQwLYOtn5oQbm039bUc6x9nB
# YWHylRKhDQeuYyHY9Jkc/3hVge6leegggl8K2rVTGVQBVw2HkY3CfPFUhoDhYtuC
# cz4mXvBAEtI51SYDDYWIMV8KC4sCAwEAAaOCAX8wggF7MB8GA1UdJQQYMBYGCisG
# AQQBgjdMCAEGCCsGAQUFBwMDMB0GA1UdDgQWBBSnE10fIYlV6APunhc26vJUiDUZ
# rzBRBgNVHREESjBIpEYwRDEMMAoGA1UECxMDQU9DMTQwMgYDVQQFEysyMzAwMTIr
# YzgwNGI1ZWEtNDliNC00MjM4LTgzNjItZDg1MWZhMjI1NGZjMB8GA1UdIwQYMBaA
# FEhuZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFf
# MjAxMS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEA
# TZdPNH7xcJOc49UaS5wRfmsmxKUk9N9E1CS6s2oIiZmayzHncJv/FB2wBzl/5DA7
# EyLeDsiVZ7tufvh8laSQgjeTpoPTSQLBrK1Z75G3p2YADqJMJdTc510HAsooNGU7
# OYOtlSqOyqDoCDoc/j57QEmUTY5UJQrlsccK7nE3xpteNvWnQkT7vIewDcA12SaH
# X/9n7yh094owBBGKZ8xLNWBqIefDjQeDXpurnXEfKSYJEdT1gtPSNgcpruiSbZB/
# AMmoW+7QBGX7oQ5XU8zymInznxWTyAbEY1JhAk9XSBz1+3USyrX59MJpX7uhnQ1p
# gyfrgz4dazHD7g7xxIRDh+4xnAYAMny3IIq5CCPqVrAY1LK9Few37WTTaxUCI8aK
# M4c60Zu2wJZZLKABU4QBX/J7wXqw7NTYUvZfdYFEWRY4J1O7UPNecd/311HcMdUa
# YzUql36fZjdfz1Uz77LKvCwjqkQe7vtnSLToQsMPilFYokYCYSZaGb9clOmoQHDn
# WzBMfIDUUGeipe4O6z218eV5HuH1WBlvu4lteOIgWCX/5Eiz5q/xskAEF0ZQ1Axs
# kRR97sri9ibeGzsEZ1EuD6QX90L/P5GJMfinvLPlOlLcKjN/SmSRZdhlEbbbare0
# bFL8v4txFsQsznOaoOldCMFFRaUphuwBMW1edMZWMQswggYHMIID76ADAgECAgph
# Fmg0AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20x
# GTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0
# MDMxMzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# ITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP
# 7tGn0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySH
# nfL0Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUo
# Ri4nrIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABK
# R2YRJylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSf
# rx54QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGn
# MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMP
# MAsGA1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQO
# rIJgQFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZ
# MBcGCgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJv
# b3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1Ud
# HwRJMEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3By
# b2R1Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYI
# KwYBBQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWlj
# cm9zb2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3
# DQEBBQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKi
# jG1iuFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV
# 3U+rkuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5
# nGctxVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tO
# i3/FNSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbM
# UVbonXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXj
# pKh0NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh
# 0EPpK+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLax
# aj2JoXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWw
# ymO0eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma
# 7kng9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCB3ow
# ggVioAMCAQICCmEOkNIAAAAAAAMwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDExMB4XDTExMDcwODIwNTkwOVoX
# DTI2MDcwODIxMDkwOVowfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKvw+nIQHC6t2G6qghBNNLry
# tlghn0IbKmvpWlCquAY4GgRJun/DDB7dN2vGEtgL8DjCmQawyDnVARQxQtOJDXlk
# h36UYCRsr55JnOloXtLfm1OyCizDr9mpK656Ca/XllnKYBoF6WZ26DJSJhIv56sI
# UM+zRLdd2MQuA3WraPPLbfM6XKEW9Ea64DhkrG5kNXimoGMPLdNAk/jj3gcN1Vx5
# pUkp5w2+oBN3vpQ97/vjK1oQH01WKKJ6cuASOrdJXtjt7UORg9l7snuGG9k+sYxd
# 6IlPhBryoS9Z5JA7La4zWMW3Pv4y07MDPbGyr5I4ftKdgCz1TlaRITUlwzluZH9T
# upwPrRkjhMv0ugOGjfdf8NBSv4yUh7zAIXQlXxgotswnKDglmDlKNs98sZKuHCOn
# qWbsYR9q4ShJnV+I4iVd0yFLPlLEtVc/JAPw0XpbL9Uj43BdD1FGd7P4AOG8rAKC
# X9vAFbO9G9RVS+c5oQ/pI0m8GLhEfEXkwcNyeuBy5yTfv0aZxe/CHFfbg43sTUkw
# p6uO3+xbn6/83bBm4sGXgXvt1u1L50kppxMopqd9Z4DmimJ4X7IvhNdXnFy/dygo
# 8e1twyiPLI9AN0/B4YVEicQJTMXUpUMvdJX3bvh4IFgsE11glZo+TzOE2rCIF96e
# TvSWsLxGoGyY0uDWiIwLAgMBAAGjggHtMIIB6TAQBgkrBgEEAYI3FQEEAwIBADAd
# BgNVHQ4EFgQUSG5k5VAF04KqFzc3IrVtqMp1ApUwGQYJKwYBBAGCNxQCBAweCgBT
# AHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgw
# FoAUci06AjGQQ7kUBU7h6qfHMdEjiTQwWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDov
# L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0
# MjAxMV8yMDExXzAzXzIyLmNybDBeBggrBgEFBQcBAQRSMFAwTgYIKwYBBQUHMAKG
# Qmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0
# MjAxMV8yMDExXzAzXzIyLmNydDCBnwYDVR0gBIGXMIGUMIGRBgkrBgEEAYI3LgMw
# gYMwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# ZG9jcy9wcmltYXJ5Y3BzLmh0bTBABggrBgEFBQcCAjA0HjIgHQBMAGUAZwBhAGwA
# XwBwAG8AbABpAGMAeQBfAHMAdABhAHQAZQBtAGUAbgB0AC4gHTANBgkqhkiG9w0B
# AQsFAAOCAgEAZ/KGpZjgVHkaLtPYdGcimwuWEeFjkplCln3SeQyQwWVfLiw++MNy
# 0W2D/r4/6ArKO79HqaPzadtjvyI1pZddZYSQfYtGUFXYDJJ80hpLHPM8QotS0LD9
# a+M+By4pm+Y9G6XUtR13lDni6WTJRD14eiPzE32mkHSDjfTLJgJGKsKKELukqQUM
# m+1o+mgulaAqPyprWEljHwlpblqYluSD9MCP80Yr3vw70L01724lruWvJ+3Q3fMO
# r5kol5hNDj0L8giJ1h/DMhji8MUtzluetEk5CsYKwsatruWy2dsViFFFWDgycSca
# f7H0J/jeLDogaZiyWYlobm+nt3TDQAUGpgEqKD6CPxNNZgvAs0314Y9/HG8VfUWn
# duVAKmWjw11SYobDHWM2l4bf2vP48hahmifhzaWX0O5dY0HjWwechz4GdwbRBrF1
# HxS+YWG18NzGGwS+30HHDiju3mUv7Jf2oVyW2ADWoUa9WfOXpQlLSBCZgB/QACnF
# sZulP0V3HjXG0qKin3p6IvpIlR+r+0cjgPWe+L9rt0uX4ut1eBrs6jeZeRhL/9az
# I2h15q/6/IvrC4DqaTuv/DDtBEyO3991bWORPdGdVk5Pv4BXIqF4ETIheu9BCrE/
# +6jMpF3BoYibV3FWTkhFwELJm3ZbCoBIa/15n8G9bW1qyVJzEw16UM0xggTRMIIE
# zQIBATCBlTB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgw
# JgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExAhMzAAAAww6b
# p9iy3PcsAAAAAADDMAkGBSsOAwIaBQCggeUwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkE
# MRYEFHJGFE36PMZ5eo3pS7/4m7XTlcsOMIGEBgorBgEEAYI3AgEMMXYwdKBMgEoA
# TABvAGMAawBBAG4AZABTAGEAdgBlAFAAdQBiAGwAaQBjAEYAbwBsAGQAZQByAFAA
# cgBvAHAAZQByAHQAaQBlAHMALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBAFnoSISR9V/svL7Yeqm/
# z5cOMqyb3r/3kt8hfMBXPswV3Ki7/TW2ssnBmwfLrtPDimLLolOawQBwkuoaOOQX
# NQ2p09CkrJSS3oagtBH3GTfWhXG4teILqp+9LGKuByz09EfO+goYBSTpq44bn/Nf
# xivtVpZ3uPWOQYdkF5GmXtTPoT0OFf2oJ+/xpEy2OnSQbstfWcSx9rZunkh4nLsr
# 5xM/ylpg8B1+4UOCYVzWIr1xhUZnPjzEUeho3XKOHJwofkA8gptFKaQ+TzCtBOIb
# x0bs9SX1E91XUpmiWjXDALY3PL8MN7tdLa6QLgHpp5oZysqGraf5aRS1cP4nxUap
# fMihggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBAhMzAAAAwzu4ENardZyEAAAAAADDMAkGBSsOAwIaBQCgXTAY
# BgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNzEwMDUw
# NjAxNTlaMCMGCSqGSIb3DQEJBDEWBBRdO3U4TlS7vKhoHkeb/NQ0MNRD1jANBgkq
# hkiG9w0BAQUFAASCAQAT/ufsfjw+PFoMCIAJbEEuHxKXr+otEVF2naygdoiDgxmN
# Qw68tEEaH1CU3mtQRvLixgoBlSRvPs9g5/tpeWo+EMrVXLw0IdKmF5eoLnFOSLCw
# CytXrtCTyZwHPfffXr4xA7whpxhZKy833pLqkt+Pp2LPUfMaShZzulOxWdaTx4i6
# ENVST7A9xQyKgZmRBzO5M7cjSv/9af//ceYe27zqsmq/PpOUtoMMm4+0YUq/THZN
# 1CgGcVBaJi8tLuFhAx3Azg6J4MgLKSBNxEcrC1MC1rZ7IR82AsPe2YEbmo2ehwSy
# Xyscdrg5sZ8mj5D+9o8JPABmkUMaMLSnQZXpmlvD
# SIG # End signature block
