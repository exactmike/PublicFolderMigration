# .SYNOPSIS
# UnlockAndRestorePublicFolderProperties.ps1
#    It recovers the public folders with the Permissions it had before the lockdown process,
#    if the user want to fallback to public folders. However, any new content posted to group
#    will not be available in public folder.
#
# .DESCRIPTION
#    Script performs the following actions
#      1. It resumes PublicFolderClientPermissions of migrated public folders from backup file.
#      2. Mail-Enable and resume mail properties of any mail public folder which had migrated.
#      3. Resume the smtp addresses of mail public folders from the Proxy Address list of Groups.
#
#   After the execution of the script
#      1. All public folders which got migrated will resume it's permissions it had, at the 
#         time of locking public folder
#      2. All mails sent to smtp addresses of mail public folder had, will be routed to the 
#         mail public folders itself (No longer be redirected to group).
#      3. Any new item(s) posted to the group can only be accessed only from the group; after 
#         restore process, public folder will not contain any new data that was posted in group.
#
# .PARAMETER BackupDir
#    The directory that user had saved the permissions and other properties at the 
#    time of locking public folders.
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
#    .\UnlockAndRestorePublicFolderProperties.ps1 -BackupDir C:\PFToGroupMigration\ -WhatIf
#    .\UnlockAndRestorePublicFolderProperties.ps1 -BackupDir C:\PFToGroupMigration -ArePublicFoldersOnPremises $true -Credential (Get-Credential)
#
# Copyright (c) 2017 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Param(

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string] $BackupDir,

    [Parameter(Mandatory = $false, HelpMessage = "Enter '`$true' if public folders are on-premises)")]
    [ValidateNotNullOrEmpty()]
    [bool] $ArePublicFoldersOnPremises = $false,

    [Parameter(Mandatory=$false, HelpMessage = "Enter the Exchange Online admin credential(userName and password")]
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

$permissionListCsv  = Join-Path $BackupDir ($prefix + "PfPermissions.csv");
$pfToGrpMappingCsv  = Join-Path $BackupDir ($prefix + "PfToGroupMapping.csv");
$pfMailPropCsv      = Join-Path $BackupDir ($prefix + "PfMailProperties.csv");
$logPath            = Join-Path $BackupDir ($prefix + "PfLockdown_summary.log");

###################### END OF DEFAULTS ######################

# Load necessary helper functions
. ".\WriteLog.ps1"
. ".\RetryScriptBlock.ps1"

# Load localized strings
Import-LocalizedData -BindingVariable LocalizedStrings -FileName UnlockAndRestorePublicFolderProperties.strings.psd1

# Function to remove SendAs permission from group
function RemoveSendAsPermissionFromGroup
{
    Param ($groupId, $trustee)

    if ($ArePublicFoldersOnPremises)
    {
        Remove-EXORecipientPermission -Identity $groupId -Trustee $trustee -AccessRights SendAs -confirm:$false;
    }
    else
    {
        Remove-RecipientPermission -Identity $groupId -Trustee $trustee -AccessRights SendAs -confirm:$false;
    }

    return $?;
}

# Function to add SendAs permission to public folder
function AddSendAsPermissionToPf
{
    Param ($primarySmtpOfPf, $trustee)

    Add-RecipientPermission -Identity $primarySmtpOfPf -Trustee $trustee -AccessRights SendAs -confirm:$false;
    return $?;
}

# Function to add SendOnBehalfTo permission to public folder
function AddSendOnBehalfToPermissionToPf
{
    Param ($primarySmtpOfPf, $user)

    Set-MailPublicFolder $primarySmtpOfPf -GrantSendOnBehalfTo @{Add=$user};
    return $?;
}

# function to enable mail public folder with the properties
function EnableMailPfWithProperties
{
    Param ($identity, $pfEmailIds, $extAddr)

    $Error[0] = $null;
    if ($extAddr)
    {
        Set-MailPublicFolder $identity -EmailAddresses $pfEmailIds -ExternalEmailAddress $extAddr -EmailAddressPolicyEnabled $false 2> $null;
    }
    else
    {
        Set-MailPublicFolder $identity -EmailAddresses $pfEmailIds -EmailAddressPolicyEnabled $false 2> $null;
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

# If the backup files are not found, script will terminate.
$MissingFiles = @();

if (!(Test-Path $permissionListCsv))
{
    $MissingFiles += $permissionListCsv;
}

if (!(Test-Path $pfToGrpMappingCsv))
{
    $MissingFiles += $pfToGrpMappingCsv;
}

if (!(Test-Path $pfMailPropCsv))
{
    $MissingFiles += $pfMailPropCsv;
}

if ($MissingFiles)
{
    $MissingFiles = $MissingFiles -Join ", ";
    WriteLog -Path $logPath -Level Error -Message ($LocalizedStrings.BackupNotFound -f $MissingFiles);
    return;
}

#Importing PermissionList and MailProperties of public folders.
WriteLog -Path $logPath -Message $LocalizedStrings.ReadingPfPerms;
$accessRights = Import-Csv $permissionListCsv;
$pfEmailIdsAndGroup = Import-Csv $pfToGrpMappingCsv;

if ($accessRights.Length -eq 0)
{
    WriteLog -Path $logPath -Level Error -Message ($LocalizedStrings.IncorrectCsv -f $permissionListCsv);
    return;
}

if ($pfEmailIdsAndGroup.Length -eq 0)
{
    WriteLog -Path $logPath -Level Error -Message ($LocalizedStrings.IncorrectCsv -f $pfToGrpMappingCsv);
    return;
}

WriteLog -Path $logPath -Message ($LocalizedStrings.ImportingBackupFilesSuccessful -f $permissionListCsv, $pfToGrpMappingCsv);

WriteLog -Path $logPath -Message ($LocalizedStrings.RestoringPfPerms -f $permissionListCsv);

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

    foreach ($accessRight in $accessRights)
    {
        $user = ([string]$accessRight.User).ToLower();

        # Checking if the user exists.
        if ($user -ne "default" -and $user -ne "anonymous")
        {
            $userSmtpAddress = $accessRight.PrimarySmtpAddress;

            if ($isE14OnPrem)
            {
                $uniqueUser = [string] (Get-Recipient $user -ErrorAction SilentlyContinue).Name;
            }
            elseif ($userSmtpAddress)
            {
                # User existed at the time of lockdown. Checking the user still exists.
                $uniqueUser = [string] (Get-Recipient $userSmtpAddress -ErrorAction SilentlyContinue).Name;
            }
            else
            {
                # Invalid user from the time of lockdown.
                $uniqueUser = $null;
            }

            if (!$uniqueUser)
            {
                WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.SkippingUser -f $user);
                continue;
            }
        }
        else
        {
            $uniqueUser = $user;
        }

        if ($WhatIf)
        {
            WriteLog -Path $logPath -Message ($LocalizedStrings.RemovingPfPermission -f $uniqueUser, [string]$accessRight.Identity);
            WriteLog -Path $logPath -Message ($LocalizedStrings.AddPfPermission -f $uniqueUser, [string]$accessRight.AccessRights, [string]$accessRight.Identity);
            continue;
        }

        # Removing any current access right for that particular user on that particular public folder, before we add an access right from backup entry.

        if ($isE14OnPrem)
        {
            # In E14 we need to specify the AccessRight parameter in Remove-PublicFolderClientPermission, whereas in the
            # latest versions the parameter AccessRights is not present.
            $perm = Get-PublicFolderClientPermission -Identity $accessRight.Identity -User $uniqueUser;
            if ($perm)
            {
                Remove-PublicFolderClientPermission -Identity $accessRight.Identity -User $uniqueUser -AccessRights $perm.AccessRights -ErrorAction SilentlyContinue -Confirm:$false;
            }
        }
        else
        {
            Remove-PublicFolderClientPermission -Identity $accessRight.Identity -User $uniqueUser -ErrorAction SilentlyContinue -Confirm:$false;
        }

        # Add the access right from the backup file entry.
        Add-PublicFolderClientPermission -Identity $accessRight.Identity -User $uniqueUser -AccessRights $accessRight.AccessRights.Split();
        if (!$?)
        {
            WriteLog -Path $logPath -Level Error -LogOnly -Message $Error[0].Exception; 
        }
    }

    WriteLog -Path $logPath -Message $LocalizedStrings.RestoredPfPerms;

    WriteLog -Path $logPath -Message $LocalizedStrings.MailEnablingPfs;

    foreach ($pfEmailIdsAndGroupItem in $pfEmailIdsAndGroup)
    {
        if (!($pfEmailIdsAndGroupItem.EmailAddresses))
        {
            # This public folder was not mail enabled at the time of lock down.
            continue;
        }

        $pfEmailIds = $pfEmailIdsAndGroupItem.EmailAddresses.Split();
        $identity = $pfEmailIdsAndGroupItem.Identity;
        $sendOnBehalfTo = $pfEmailIdsAndGroupItem.GrantSendOnBehalfTo;
        $sendOnBehalfToList = $sendOnBehalfTo.Split();

        $sendOnBehalfToAddedByScript = ($pfEmailIdsAndGroupItem.SendOnBehalfToAddedByScript).Split();
        $sendAsAddedByScript = ($pfEmailIdsAndGroupItem.SendAsAddedByScript).Split();

        $extAddr = $pfEmailIdsAndGroupItem.ExternalEmailAddress;
        $emailAddressPolicyEnabled = [System.Convert]::ToBoolean($pfEmailIdsAndGroupItem.EmailAddressPolicyEnabled);
        $sendAsList = ($pfEmailIdsAndGroupItem.SendAsList).Split();
        $groupId = $pfEmailIdsAndGroupItem.UnifiedGroup;

        if ($WhatIf)
        {
            WriteLog -Path $logPath -Message ($LocalizedStrings.RemovedPropertiesFromGroup -f $groupId, [string]$SendAsAddedByScript, [string]$sendOnBehalfToAddedByScript, [string]$pfEmailIds);
            WriteLog -Path $logPath -Message ($LocalizedStrings.MailEnabledPf -f $identity);
            WriteLog -Path $logPath -Message ($LocalizedStrings.AddedPropertiesBackToPf -f $identity, [string]$SendAsList, [string]$sendOnBehalfToList, $emailAddressPolicyEnabled);
            continue;
        }

        # Removing the public folder smtp addresses from Group's proxy address list.
        if ($ArePublicFoldersOnPremises)
        {
            Set-EXOUnifiedGroup $groupId -EmailAddresses @{Remove=$pfEmailIds};
        }
        else
        {
            Set-UnifiedGroup $groupId -EmailAddresses @{Remove=$pfEmailIds};
        }
        # Removing SendOnBehalfTo from group, that had assigned at the time of lockdown
        if ($sendOnBehalfToAddedByScript)
        {
            # Trying whole list (silently) first.
            $Error[0] = $null;

            if ($ArePublicFoldersOnPremises)
            {
                Set-EXOUnifiedGroup $groupId -GrantSendOnBehalfTo @{Remove=$sendOnBehalfToAddedByScript} 2> $null;
            }
            else
            {
                Set-UnifiedGroup $groupId -GrantSendOnBehalfTo @{Remove=$sendOnBehalfToAddedByScript} 2> $null;
            }

            if ($Error[0])
            {
                # There could be one or more invalid users.
                # Trying the list items one by one.
                foreach ($sendOnBehalfToItem in $sendOnBehalfToAddedByScript)
                {
                    if ($ArePublicFoldersOnPremises)
                    {
                        Set-EXOUnifiedGroup $groupId -GrantSendOnBehalfTo @{Remove=$sendOnBehalfToItem};
                    }
                    else
                    {
                        Set-UnifiedGroup $groupId -GrantSendOnBehalfTo @{Remove=$sendOnBehalfToItem};
                    }

                    if (!$?)
                    {
                        WriteLog -Path $logPath -Level Error -LogOnly -Message $Error[0].Exception;
                    }
                }
            }
        }

        # Removing SendAs permission from the group, that had assigned at the time of lockdown.
        if ($sendAsAddedByScript)
        {
            $sendAsAddedByScript | %{`
                ExecuteWithRetries -ScriptToRetry ${function:RemoveSendAsPermissionFromGroup} `
                                   -ArgumentList @($groupId, $_) `
                                   -ErrorMessageOnFailures ($LocalizedStrings.RemovingSendAsFromGroupFailed -f $groupId) `
                                   -NumberOfRetries 3 `
                                   -LogPath $logPath; `
            };
        }

        # Mail Enabling public folder.
        Enable-MailPublicFolder -Identity $identity;

        ExecuteWithRetries -ScriptToRetry ${function:EnableMailPfWithProperties} `
                           -ArgumentList @($identity, $pfEmailIds, $extAddr) `
                           -ErrorMessageOnFailures ($LocalizedStrings.RestoringSMTPFailed -f $identity) `
                           -MessageIfSucceeded ($LocalizedStrings.RestoringSMTPSucceeded -f $identity) `
                           -LogPath $logPath;

        # Setting EmailAddressPolicyEnabled to original value.
        Set-MailPublicFolder -Identity $identity -EmailAddressPolicyEnabled $emailAddressPolicyEnabled;

        $primarySmtpOfPf = [string](Get-MailPublicFolder $identity).PrimarySmtpAddress;

        # Adding SendOnBehalfTo to public folder, that existed at the time of lockdown
        if ($sendOnBehalfToList)
        {
            # Trying whole list (silently) first
            $Error[0] = $null;
            Set-MailPublicFolder $primarySmtpOfPf -GrantSendOnBehalfTo @{Add=$sendOnBehalfToList} 2> $null;
            if ($Error[0])
            {
                # There could be one or more invalid users.
                # Trying the list items one by one. 
                $sendOnBehalfToList | %{`
                    ExecuteWithRetries -ScriptToRetry ${function:AddSendOnBehalfToPermissionToPf} `
                                       -ArgumentList @($primarySmtpOfPf, $_) `
                                       -ErrorMessageOnFailures ($LocalizedStrings.AddingSendOnBehalfToPermissionFailed -f $primarySmtpOfPf, $_) `
                                       -NumberOfRetries 3 `
                                       -LogPath $logPath; `
                };
            }
        }

        # Adding SendAs permission to the public folder, that was present at the time of lockdown.
        if ($sendAsList)
        {
            $sendAsList | %{`
                ExecuteWithRetries -ScriptToRetry ${function:AddSendAsPermissionToPf} `
                                   -ArgumentList @($primarySmtpOfPf, $_) `
                                   -ErrorMessageOnFailures ($LocalizedStrings.AddingSendAsToPfFailed -f $identity) `
                                   -NumberOfRetries 3 `
                                   -LogPath $logPath; `
            };
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
    WriteLog -Path $logPath -Message ($LocalizedStrings.MailEnabledAndRestoredProperties -f $identity);
    WriteLog -Path $logPath -Message $LocalizedStrings.PfRecoveryComplete;
}
# SIG # Begin signature block
# MIIdxAYJKoZIhvcNAQcCoIIdtTCCHbECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqq9O3+cZxb35cUZKclSEBCN/
# QL6gghhSMIIEwTCCA6mgAwIBAgITMwAAAMKgCcU3dun2zQAAAAAAwjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwOTA3MTc1ODUx
# WhcNMTgwOTA3MTc1ODUxWjCBsTELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEMMAoGA1UECxMDQU9DMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpD
# M0IwLTBGNkEtNDExMTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vy
# dmljZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJzfPT5gT5YLgF72
# 8Ipv/kMSm0FRtZmMMXMdDBrWM+LOObrNAITBA0w185w4qccTOzXIgsFlOyvvyGfI
# jH+4zLekfpL8U7DuccyDVdS3Lg70hYBCEJll0SwAhfpHR1D4NQaeIRnhnlRuSUwy
# 7LqOxCE6If90dH0+OaVlxiKHw7R5RgeO50m15BHI+6v9US70IZ8JFqRkfLpk52bh
# LNfnossW+CHvAFPVQ0uThMOaoESnJsmban0QaExZvftxreTrz2QQcVw74Y29CYbZ
# RUTIy4zIpuM/i5oBLj9mwf9CogC0rQibwWfEvPyiFuOZ/ncDX5I8KVHa4Y1LoFQq
# YWk/EEkCAwEAAaOCAQkwggEFMB0GA1UdDgQWBBTjHnnY/MhgLBEZmBJtobBujc6d
# rDAfBgNVHSMEGDAWgBQjNPjZUkZwCu1A+3b7syuwwzWzDzBUBgNVHR8ETTBLMEmg
# R6BFhkNodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9N
# aWNyb3NvZnRUaW1lU3RhbXBQQ0EuY3JsMFgGCCsGAQUFBwEBBEwwSjBIBggrBgEF
# BQcwAoY8aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3Nv
# ZnRUaW1lU3RhbXBQQ0EuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3
# DQEBBQUAA4IBAQAoNFRrsA/+bdu8IJvKoxcry0vIPw0qzrUya7ud9MrJ/pp9EO01
# OFrXqbFfuPW0niqZt7hYrs7bzwSlmbBItCkImv0GCLS/3cf0Vl/c0NxUpn8TUjoo
# +qwnPF3qRGUzcwrI/3Xl9EfoDlc8jWd2f5FqrjeQdmkdOUmtxSnVt1kbW+Fnjlyl
# 1q8aWpkXXgNrBD29iXQV7BklsvtzSVLB32UTZqADm/yzqPC+osWN2eHED2nag1w0
# 51bq++5Pc2mA/UbJeqv+J9VhQwyTGoFdCjE9ygfd7aASPsxiAsRBsNRlylFMjePA
# nFZyI0P0rM+CW09Q641SEKIKbT6T1ww+8ByJMIIGADCCA+igAwIBAgITMwAAAMMO
# m6fYstz3LAAAAAAAwzANBgkqhkiG9w0BAQsFADB+MQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBTaWdu
# aW5nIFBDQSAyMDExMB4XDTE3MDgxMTIwMjAyNFoXDTE4MDgxMTIwMjAyNFowdDEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEeMBwGA1UEAxMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEAu1fXONGxBn9JLalts2Oferq2OiFbtJiujdSkgaDFdcUs74JAKreBU3fzYwEK
# vM43hANAQ1eCS87tH7b9gG3JwpFdBcfcVlkA4QzrV9798biQJ791Svx1snJYtsVI
# mzNiBdGVlKW/OSKtjRJNRmLaMhnOqiJcVkixb0XJZ3ZiXTCIoy8oxR9QKtmG2xoR
# JYHC9PVnLud5HfXiHHX0TszH/Oe/C4BHKf/PzWmxDAtg62fmhBubTf1tRzrH2cFh
# YfKVEqENB65jIdj0mRz/eFWB7qV56CCCXwratVMZVAFXDYeRjcJ88VSGgOFi24Jz
# PiZe8EAS0jnVJgMNhYgxXwoLiwIDAQABo4IBfzCCAXswHwYDVR0lBBgwFgYKKwYB
# BAGCN0wIAQYIKwYBBQUHAwMwHQYDVR0OBBYEFKcTXR8hiVXoA+6eFzbq8lSINRmv
# MFEGA1UdEQRKMEikRjBEMQwwCgYDVQQLEwNBT0MxNDAyBgNVBAUTKzIzMDAxMitj
# ODA0YjVlYS00OWI0LTQyMzgtODM2Mi1kODUxZmEyMjU0ZmMwHwYDVR0jBBgwFoAU
# SG5k5VAF04KqFzc3IrVtqMp1ApUwVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljQ29kU2lnUENBMjAxMV8yMDEx
# LTA3LTA4LmNybDBhBggrBgEFBQcBAQRVMFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljQ29kU2lnUENBMjAxMV8y
# MDExLTA3LTA4LmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBN
# l080fvFwk5zj1RpLnBF+aybEpST030TUJLqzagiJmZrLMedwm/8UHbAHOX/kMDsT
# It4OyJVnu25++HyVpJCCN5Omg9NJAsGsrVnvkbenZgAOokwl1NznXQcCyig0ZTs5
# g62VKo7KoOgIOhz+PntASZRNjlQlCuWxxwrucTfGm1429adCRPu8h7ANwDXZJodf
# /2fvKHT3ijAEEYpnzEs1YGoh58ONB4Nem6udcR8pJgkR1PWC09I2Bymu6JJtkH8A
# yahb7tAEZfuhDldTzPKYifOfFZPIBsRjUmECT1dIHPX7dRLKtfn0wmlfu6GdDWmD
# J+uDPh1rMcPuDvHEhEOH7jGcBgAyfLcgirkII+pWsBjUsr0V7DftZNNrFQIjxooz
# hzrRm7bAllksoAFThAFf8nvBerDs1NhS9l91gURZFjgnU7tQ815x3/fXUdwx1Rpj
# NSqXfp9mN1/PVTPvssq8LCOqRB7u+2dItOhCww+KUViiRgJhJloZv1yU6ahAcOdb
# MEx8gNRQZ6Kl7g7rPbXx5Xke4fVYGW+7iW144iBYJf/kSLPmr/GyQAQXRlDUDGyR
# FH3uyuL2Jt4bOwRnUS4PpBf3Qv8/kYkx+Ke8s+U6UtwqM39KZJFl2GURtttqt7Rs
# Uvy/i3EWxCzOc5qg6V0IwUVFpSmG7AExbV50xlYxCzCCBgcwggPvoAMCAQICCmEW
# aDQAAAAAABwwDQYJKoZIhvcNAQEFBQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZ
# MBcGCgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJv
# b3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5MB4XDTA3MDQwMzEyNTMwOVoXDTIxMDQw
# MzEzMDMwOVowdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEh
# MB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEAn6Fssd/bSJIqfGsuGeG94uPFmVEjUK3O3RhOJA/u
# 0afRTK10MCAR6wfVVJUVSZQbQpKumFwwJtoAa+h7veyJBw/3DgSY8InMH8szJIed
# 8vRnHCz8e+eIHernTqOhwSNTyo36Rc8J0F6v0LBCBKL5pmyTZ9co3EZTsIbQ5ShG
# Lieshk9VUgzkAyz7apCQMG6H81kwnfp+1pez6CGXfvjSE/MIt1NtUrRFkJ9IAEpH
# ZhEnKWaol+TTBoFKovmEpxFHFAmCn4TtVXj+AZodUAiFABAwRu233iNGu8QtVJ+v
# HnhBMXfMm987g5OhYQK1HQ2x/PebsgHOIktU//kFw8IgCwIDAQABo4IBqzCCAacw
# DwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQUIzT42VJGcArtQPt2+7MrsMM1sw8w
# CwYDVR0PBAQDAgGGMBAGCSsGAQQBgjcVAQQDAgEAMIGYBgNVHSMEgZAwgY2AFA6s
# gmBAVieX5SUT/CrhClOVWeSkoWOkYTBfMRMwEQYKCZImiZPyLGQBGRYDY29tMRkw
# FwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9v
# dCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHmCEHmtFqFKoKWtTHNY9AcTLmUwUAYDVR0f
# BEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJv
# ZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3JsMFQGCCsGAQUFBwEBBEgwRjBEBggr
# BgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNy
# b3NvZnRSb290Q2VydC5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcN
# AQEFBQADggIBABCXisNcA0Q23em0rXfbznlRTQGxLnRxW20ME6vOvnuPuC7UEqKM
# bWK4VwLLTiATUJndekDiV7uvWJoc4R0Bhqy7ePKL0Ow7Ae7ivo8KBciNSOLwUxXd
# T6uS5OeNatWAweaU8gYvhQPpkSokInD79vzkeJkuDfcH4nC8GE6djmsKcpW4oTmc
# Zy3FUQ7qYlw/FpiLID/iBxoy+cwxSnYxPStyC8jqcD3/hQoT38IKYY7w17gX606L
# f8U1K16jv+u8fQtCe9RTciHuMMq7eGVcWwEXChQO0toUmPU8uWZYsy0v5/mFhsxR
# VuidcJRsrDlM1PZ5v6oYemIp76KbKTQGdxpiyT0ebR+C8AvHLLvPQ7Pl+ex9teOk
# qHQ1uE7FcSMSJnYLPFKMcVpGQxS8s7OwTWfIn0L/gHkhgJ4VMGboQhJeGsieIiHQ
# Q+kr6bv0SMws1NgygEwmKkgkX1rqVu+m3pmdyjpvvYEndAYR7nYhv5uCwSdUtrFq
# PYmhdmG0bqETpr+qR/ASb/2KMmyy/t9RyIwjyWa9nR2HEmQCPS2vWY+45CHltbDK
# Y7R4VAXUQS5QrJSwpXirs6CWdRrZkocTdSIvMqgIbqBbjCW/oO+EyiHW6x5PyZru
# SeD3AWVviQt9yGnI5m7qp5fOMSn/DsVbXNhNG6HY+i+ePy5VFmvJE6P9MIIHejCC
# BWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJv
# b3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcN
# MjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2
# WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSH
# fpRgJGyvnkmc6Whe0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQ
# z7NEt13YxC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHml
# SSnnDb6gE3e+lD3v++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3o
# iU+EGvKhL1nkkDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6
# nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6ep
# ZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf
# 28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCn
# q47f7Fufr/zdsGbiwZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx
# 7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O
# 9JawvEagbJjS4NaIjAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0G
# A1UdDgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMA
# dQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAW
# gBRyLToCMZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8v
# Y3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQy
# MDExXzIwMTFfMDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZC
# aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQy
# MDExXzIwMTFfMDNfMjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCB
# gzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9k
# b2NzL3ByaW1hcnljcHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABf
# AHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEB
# CwUAA4ICAQBn8oalmOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LR
# bYP+vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r
# 4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb
# 7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6v
# mSiXmE0OPQvyCInWH8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/
# sfQn+N4sOiBpmLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad2
# 5UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUf
# FL5hYbXw3MYbBL7fQccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWx
# m6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMj
# aHXmr/r8i+sLgOppO6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7
# qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCBNwwggTY
# AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAm
# BgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAADDDpun
# 2LLc9ywAAAAAAMMwCQYFKw4DAhoFAKCB8DAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQx
# FgQUZDWDiJqo0IyYRkcwkzUvf4h/LMEwgY8GCisGAQQBgjcCAQwxgYAwfqBWgFQA
# VQBuAGwAbwBjAGsAQQBuAGQAUgBlAHMAdABvAHIAZQBQAHUAYgBsAGkAYwBGAG8A
# bABkAGUAcgBQAHIAbwBwAGUAcgB0AGkAZQBzAC4AcABzADGhJIAiaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQBXFI2S
# c9N+hrPXli8xD6R1sEGNB/ATkHRtwcuYPneDm7gdJG0jpIqS1SO8eoNZ13zWRa5w
# DTGQgYDNwiRRcSdWQyvvNTaZ7ZDa+ngiQTETVIBGGo8zK+5qqzTm2P0ITmBUmAy9
# S3ypHxgcxxc4hzICLQifDrR5590oyAm5GkxEcE3MuD8QBXvk+iZpqOZJoHWlQ/dK
# KBAjQjhde2aFLbtdpoSx0vKmEA8EyK0Z/7LLUh023KCkgElRtPlq1aA+jvOfyBQQ
# 0FAfi2F8AX6lm1W+0djLrmeDko+38UOQcG3eauutdajbe2v4SDOU32Nn3tvEjfTZ
# 8ydymOMAa0aIHiXroYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAMKgCcU3dun2zQAAAAAAwjAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMTcxMDA1MDYwMjE5WjAjBgkqhkiG9w0BCQQxFgQUMajBaF1oi999VHOmIqv9
# uwIQSLIwDQYJKoZIhvcNAQEFBQAEggEAlXrJkgW0Wcl45gHankR4mXt93cHNl+xz
# WlLRa4PWYsEvdFLOhHjhJQnvGtcPgZDeAFSf920SUbpylmUxaEkNDpFYuOuZYxfR
# RM5NUXp2WmQu0EMRgTPy06pphuQgLK5HQcdcQdSY58I+vh9RX2EusxIz8tBtInpD
# PrvEJiuUwAcl3JasczMFtJJ4TBqbR1+KUpU1lDuPWLVWMLGBI//NcLuqVZKg9YdD
# ct92TIDmBV4dBnqOdGNdTr9164P4wJ2rmfjWykjbhXKZHmVHv6O8VchVXgS8ac8S
# c8zapf9cRksfyq8uO7mYCpMofPZxn0tekyNKIutwvTtK5sxxXurAag==
# SIG # End signature block
