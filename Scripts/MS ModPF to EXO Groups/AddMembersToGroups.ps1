# .SYNOPSIS
# AddMembersToGroups.ps1
#    This script reads the permission entries for each public folder (from the backup file if the public folders are locked) and
#    adds the users with specific permission entries as either owner/member to the respective group based on their access rights.
#
# .DESCRIPTION
#    1. It reads the permission entries from the backup file stored during lock down if public folders are locked, else uses 'Get-PublicFolderClientPermission' to get the permissions.
#    2. It adds the users with permission roles, "Owner, PublishingEditor, Editor, PublishingAuthor, Author" as members to the group
#    3. It also adds users having atleast "ReadItems, CreateItems, FolderVisible, EditOwnedItems, DeleteOwnedItems" access rights as members to the corresponding group.
#    4. It adds the users with permission role, "Owner" as owners to the group.
#    5. It throws a warning when the default permission is Author and above, suggesting the user to make the group 'public'.
#
# .PARAMETER MappingCsv
#    The public folder to group mapping csv file which was used to create the migration batch.
#
# .PARAMETER ArePublicFoldersLocked
#    Tells if public folders are locked. Set to '$true' if public folders are locked, else set to '$false'.
#
# .PARAMETER BackupDir
#    The directory to which user want to save the logs and read permissions from the back up file, if public folders are locked.
#
# .PARAMETER ArePublicFoldersOnPremises
#    Tells if public folders are on-premises. Set to '$true' if public folders are on premises, else set to '$false'.
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
#    .\AddMembersToGroups.ps1 -MappingCsv PFToGroupMap.csv -ArePublicFoldersLocked $true -BackupDir "C:\PFToGroupMigration\"
#
#    This example shows how to invoke the script when public folders are in exchange online
#
# .EXAMPLE
#    .\AddMembersToGroups.ps1 -MappingCsv PFToGroupMap.csv -ArePublicFoldersLocked $false -BackupDir "C:\PFToGroupMigration\" -ArePublicFoldersOnPremises $true -Credential (Get-Credential) -ConnectionUri "https://partner.outlook.cn/PowerShell"
#
#    This example shows how to invoke the script when public folders are on-premises
#
# .EXAMPLE
#    .\AddMembersToGroups.ps1 -MappingCsv PFToGroupMap.csv -ArePublicFoldersLocked $false -BackupDir "C:\PFToGroupMigration\" -ArePublicFoldersOnPremises $true -Credential (Get-Credential) -ConnectionUri "https://partner.outlook.cn/PowerShell" -WhatIf
#
#    This example shows how to use the 'WhatIf' parameter.
#
# Copyright (c) 2017 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Param(

    [Parameter(Mandatory = $true, HelpMessage = "Mapping csv used to create the migration batch")]
    [ValidateNotNullOrEmpty()]
    [string] $MappingCsv,

    [Parameter(Mandatory = $false, HelpMessage = "Enter '`$true' if public folders are locked")]
    [bool] $ArePublicFoldersLocked = $false,

    [Parameter(Mandatory = $true, HelpMessage = "Directory to write log and read permissions of locked public folders from back up file")]
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
                -AllowClobber `
                -ErrorAction:SilentlyContinue;

        if (!$? -or ($result -eq $null))
        {
            WriteLog -Path $logPath -Level Error -Message ($LocalizedStrings.FailedToImportRemoteSession -f $error[0].Exception.Message);
            Remove-PSSession $script:session;
            Exit;
        }
    }

    WriteLog -Path $logPath -Message $LocalizedStrings.RemoteSessionCreatedSuccessfully;
}

# Returns the members of the security group
function GetMembersOfSecurityGroup()
{
    param ($securityGroup);

    $securityGroup = Get-DistributionGroup $securityGroup | ?{$_.Name -Like $securityGroup};
    $securityGroupMembers = Get-DistributionGroupMember $securityGroup.PrimarySmtpAddress.ToString() | select Name, PrimarySmtpAddress, RecipientType;

    if (!$?)
    {
        WriteLog -Path $logPath -Level Error -LogOnly -Message $error[0];
    }

    return $securityGroupMembers;
}

# Adds the user to the dictionary, UniqueUserList along with the details, if the user is a SecurityGroup/ValidUser/InvalidUser,
# Returns $true when the user is of invalid recipient type and does not add such users to UniqueUserList
function AddUserToList()
{
    param ([string]$user, [string]$smtpAddress, [string]$userType, [string]$accessRight);
    
    if ($smtpAddress)
    {
        #Check if it is a security group
        if ($userType -eq [RecipientType]::MailUniversalSecurityGroup)
        {
            $UniqueUserList.Add($user, [UserDetails]::SecurityGroup);
        }
        elseif (($userType -eq [RecipientType]::UserMailbox) -or ($userType -eq [RecipientType]::MailUser))
        {
            $UniqueUserList.Add($user, [UserDetails]::ValidUser);
            $SmtpAddressOfUsers.Add($user, $smtpAddress);
        }
        else
        {
            WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.InvalidRecipientType -f $user, $userType, $accessRight);
            return $true;
        }
    }
    else
    {
        # When smtp address is null, it implies that the user does not exist
        $UniqueUserList.Add($user, [UserDetails]::InvalidUser);
    }

    return $false;
}

# Checks if the user is in exchange online and adds the user to list
function ValidateAndAddEXOUserToList()
{
    param ([string]$user, [string]$smtpAddress, [string]$userType, [string]$accessRight);

    # Using smtp address for uniqueness as "Name" will not be unique accross on-premises and exchange online
    $recipient = Get-EXORecipient $smtpAddress -ErrorAction SilentlyContinue | ?{$_.PrimarySmtpAddress -like $smtpAddress} | select RecipientType;

    if ($recipient)
    {
        $userType = $recipient.RecipientType;
    }
    else
    {
        $smtpAddress = $null;
    }

    return AddUserToList $user $smtpAddress $userType $accessRight;
}

# Validates the user and adds the user to list
function ValidateAndAddUserToList()
{
    param ([string]$user, [string]$smtpAddress, [string]$userType, [string]$accessRight);

    if ($smtpAddress)
    {
        $recipient = Get-Recipient $smtpAddress -ErrorAction SilentlyContinue | ?{$_.PrimarySmtpAddress -like $smtpAddress} | select PrimarySmtpAddress,RecipientType;
    }
    else
    {
        $recipient = Get-Recipient $user -ErrorAction SilentlyContinue | ?{$_.Name -like $user} | select PrimarySmtpAddress,RecipientType;
    }

    if ($recipient)
    {
        if (!$smtpAddress)
        {
            $smtpAddress = $recipient.PrimarySmtpAddress.ToString();
        }

        # User type will be null when public folders are locked and permissions are read from the back up file
        # (as user/recipient type is not stored during lockdown).
        if (!$userType)
        {
            $userType = $recipient.RecipientType;
        }

        if ($ArePublicFoldersOnPremises)
        {
            if ($userType -ne [RecipientType]::MailUniversalSecurityGroup)
            {
                return ValidateAndAddEXOUserToList $user $smtpAddress $userType $accessRights;
            }
        }
    }
    else
    {
        $smtpAddress = $null;
    }

    return AddUserToList $user $smtpAddress $userType $accessRight;
}

# Returns true if the user has enough access right to be added as member of the group, else returns false
function IsAccessRightSufficient()
{
    param ($accessRights);

    if ($accessRights.Count -eq 1)
    {
        if($EligibleRoles -Contains $accessRights)
        {
            return $true;
        }

        return $false;
    }

    $missingAccessRights = $NecessaryAccessRightsForMember | ?{$accessRights -notcontains $_};
    if ($missingAccessRights)
    {
        # One or more access rights required for the user to be added as a member is missing. Hence, return false
        return $false;
    }

    return $true;
}

# Processes each permission entry and decides if the user needs to be added as a member or owner to the group
function ProcessPermissionEntry()
{
    param ([string]$user, [string]$validUser, $accessRights, $owners, $members);

    if ($validUser -eq [UserDetails]::ValidUser)
    {
        if (IsAccessRightSufficient $accessRights)
        {
            # Users having 'owner' permission on the pf will be added as Owners to the group.
            # Users having other access rights will be added as members to the group.
            $smtpAddress = $SmtpAddressOfUsers[$user];
            $members.Add($smtpAddress) > $null;
            if($accessRights.Contains("Owner"))
            {
                $owners.Add($smtpAddress) > $null;
            }
        } 
        else
        {
            # Users with access right None/FolderVisible/CreateSubfolders should be skipped.
            WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.UserSkipped -f $user, [string] $accessRights);
        }
    }
    else
    {
        WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.InvalidUser -f $user, [string] $accessRights);
    }
}

################ DECLARING GLOBAL VARIABLES ################

$permissionListCsvPath = Join-Path $BackupDir "PfPermissions.csv";
$logPath = Join-Path $BackupDir "AddMembersToGroups-log.log";

# Creating a dictionary to store the users and their validity
$UniqueUserList = @{};

# Creating a dictionary to store the users and their validity
$SmtpAddressOfUsers = @{};

# List of roles that make the user eligible to be added as a member of the group
$EligibleRoles = @("Owner","PublishingEditor","Editor","PublishingAuthor","Author");

# List of access rights that are necessary for a user to be added as a member of the group
$NecessaryAccessRightsForMember = @("ReadItems","CreateItems","FolderVisible","EditOwnedItems","DeleteOwnedItems");


Add-Type -TypeDefinition @"
   public enum UserDetails
   {
       SecurityGroup,
       ValidUser,
       InvalidUser
   }
"@

Add-Type -TypeDefinition @"
   public enum RecipientType
   {
       MailUniversalSecurityGroup,
       UserMailbox,
       MailUser
   }
"@

################ END OF DECLARATION #################

# Load function to write logs
. ".\WriteLog.ps1"

# Load localized strings
Import-LocalizedData -BindingVariable LocalizedStrings -FileName AddMembersToGroups.strings.psd1

if (!(Test-Path $MappingCsv))
{ 
    WriteLog -Path $logPath -Level Error -Message $LocalizedStrings.MappingCsvNotFound;
    return;
}

# Load and validate the mapping csv
$pfToGrpMapping = Import-Csv $MappingCsv;

$invalidRows = $pfToGrpMapping | ?{$_.FolderPath -eq $null -or $_.TargetGroupMailbox -eq $null}
if ($invalidRows)
{
    WriteLog -Path $logPath -Level Error -Message $LocalizedStrings.IncorrectCsv;
    return;
}

try
{
    # If the public folders are on-premises, create an exchange online remote session
    if ($ArePublicFoldersOnPremises)
    {
        # Check if exchange online credential is provided.
        if (!$Credential)
        {
            WriteLog -Path $logPath -Level Warn -Message $LocalizedStrings.CredentialNotFound;
            $Credential = Get-Credential;
        }

        WriteLog -Path $logPath -Message $LocalizedStrings.CreatingRemoteSession;
        InitializeExchangeOnlineRemoteSession;
    }

    # If the public folders are locked, check for the back up file with permission list and read the permissions from the file,
    # else get the permissions using, "GetPublicFolderClientPermission"
    if ($ArePublicFoldersLocked)
    {
        if (!(Test-Path $permissionListCsvPath))
        {
            WriteLog -Path $logPath -Level Error -Message ($LocalizedStrings.PermissionFileMissing -f $BackupDir);
            return;
        }
    
        WriteLog -Path $logPath -Message ($LocalizedStrings.ReadingPermissionsFromFile -f $permissionListCsvPath);
        $PermissionList = Import-Csv $permissionListCsvPath;
    }
    else
    {
        WriteLog -Path $logPath -Message $LocalizedStrings.ReadingPermissions;
        $pfsBeingMigrated = $pfToGrpMapping | %{$_.FolderPath};


        if ($ArePublicFoldersOnPremises -and ((Get-ExchangeServer $env:COMPUTERNAME -ErrorAction:Stop).AdminDisplayVersion.Major -eq 14))
        {
            # ADRecipient object is not available in 2010. Hence get the user name from ActiveDirectoryIdentity in the user object. 
            $PermissionList = $pfsBeingMigrated | Get-PublicFolderClientPermission | Select Identity,
                                                                                             AccessRights,
                                                                                             User,
                                                                                             @{Name="Name";Expression={$_.User.ActiveDirectoryIdentity.Name}};
        }
        else
        {
            $PermissionList = $pfsBeingMigrated | Get-PublicFolderClientPermission | Select Identity,
                                                                                             AccessRights,
                                                                                             User,
                                                                                             @{Name="Name";Expression={$_.User.ADRecipient.Name}},
                                                                                             @{Name="PrimarySmtpAddress";Expression={$_.User.ADRecipient.PrimarySmtpAddress}},
                                                                                             @{Name="RecipientType";Expression={$_.User.ADRecipient.RecipientType}};
        }
        if (!$?)
        {
            WriteLog -Path $logPath -Level Error -LogOnly -Message $error[0];
        }
    }

    # Process the permission entries of each public folder in the mapping csv and add members to the respective group
    foreach ($pfEmailIdsAndGroupItem in $pfToGrpMapping)
    {
        $pfIdentity = $pfEmailIdsAndGroupItem.FolderPath;
        $group = $pfEmailIdsAndGroupItem.TargetGroupMailbox;
    
        WriteLog -Path $logPath -Message ($LocalizedStrings.AddingMembersToGroup -f $group, $pfIdentity);

        # Get permission entries for the public folder being processed
        $permissionEntries = $PermissionList | ?{[string]$_.Identity -eq $pfIdentity};
        if (!$permissionEntries)
        {
            if ($ArePublicFoldersLocked)
            {
                WriteLog -Path $logPath -Level Error -Message ($LocalizedStrings.PermissionEntriesMissingInFile -f $permissionListCsvPath, $pfidentity);
            }
            else
            {
                WriteLog -Path $logPath -Level Error -Message ($LocalizedStrings.PermissionEntriesMissing -f $pfidentity);
            }
        
            continue;
        }

        $accessRightsOfSpecificUsers = $permissionEntries | ?{!([string]$_.User -eq "Default" -or [string]$_.User -eq "Anonymous")};
        if (!$accessRightsOfSpecificUsers)
        {
            WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.FolderHasOnlyDefaultPermissions -f $pfidentity, $group);
        }

        # List of users to be added as owners of the group
        $owners = New-Object System.Collections.Generic.HashSet[string];

        # List of users to be added as members of the group
        $members = New-Object System.Collections.Generic.HashSet[string];

        # Dictionary of security groups and their access rights to be processed after processing the explicit permissions
        $securityGroups = @{};

        # List of users having explicit permissions;
        $usersWithExplicitPermission = New-Object System.Collections.Generic.HashSet[string];

        foreach ($permission in $accessRightsOfSpecificUsers)
        {
            $user = [string] $permission.User;
            $accessRights = $permission.AccessRights;
            $userName = $permission.Name;
            $smtpAddress = $permission.PrimarySmtpAddress;
            $userType = [string] $permission.RecipientType;

            # When the userName for a user is null, it implies that ADRecipient or ActiveDirectoryIdentity for the user is not available and the user is invalid
            if (!$userName)
            {
                WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.InvalidUser -f $user, [string] $accessRights);
                WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.UserNameIsNull -f $user);
                continue;
            }

            if (!$UniqueUserList.ContainsKey($userName))
            {
                if ($ArePublicFoldersLocked)
                {
                    $isUserTypeInvalid = ValidateAndAddUserToList $userName $smtpAddress $userType $accessRights;
                }
                else
                {
                    if ($ArePublicFoldersOnPremises -and ($userType -ne [RecipientType]::MailUniversalSecurityGroup))
                    {
                        if ($smtpAddress)
                        {
                            # Validate the user in exchange online and add the user to list
                            $isUserTypeInvalid = ValidateAndAddEXOUserToList $userName $smtpAddress $userType $accessRights;
                        }
                        else
                        {
                            # Run Get-Recipient first to get the smtpAdress
                            # smtpAddress will be null in 2010 as ADRecipient object is not available.
                            $isUserTypeInvalid = ValidateAndAddUserToList $userName $smtpAddress $userType $accessRights;
                        }
                    }
                    else
                    {
                        $isUserTypeInvalid = AddUserToList $userName $smtpAddress $userType $accessRights;
                    }
                }
            }

            # Skip users with invalid recipient type
            if ($isUserTypeInvalid)
            {
                continue;
            }

            $userType = $UniqueUserList[$userName];
            if ($userType -eq [UserDetails]::SecurityGroup)
            {
                $securityGroups.Add($userName, $accessRights);
            }
            else
            {
                $usersWithExplicitPermission.Add($userName) > $null;
                ProcessPermissionEntry $userName $userType $accessRights $owners $members;
            }
        }

        foreach ($key in $securityGroups.Keys)
        {
            $securityGroupMembers = GetMembersOfSecurityGroup($key);
            if(!$securityGroupMembers)
            {
                WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.SecurityGroupHasNoMembers -f $key);
            }

            for ($i = 0; $i -lt $securityGroupMembers.count; $i++)
            {
                $member = $securityGroupMembers[$i];
                $user = [string] $member.name;

                # Do not process those users in security groups, who have explicit permissions.
                if ($usersWithExplicitPermission.Contains($user))
                {
                    continue;
                }

                if ([string] $member.RecipientType -eq [RecipientType]::MailUniversalSecurityGroup)
                {
                    # Add members of the nested security group to list
                    $securityGroupMembers += GetMembersOfSecurityGroup($user);
                    continue;
                }

                if (!$UniqueUserList.ContainsKey($user))
                {
                    if ($ArePublicFoldersOnPremises)
                    {
                        # Validate if the user exists in EXO before adding it to the list
                        $isUserTypeInvalid = ValidateAndAddEXOUserToList $user $member.PrimarySmtpAddress $member.RecipientType $securityGroups[$key];
                    }
                    else
                    {
                        # "Get-DistributionGroupMembers" returns only those members who exist
                        # Hence validation is not required and user can be added to the list directly
                        $isUserTypeInvalid = AddUserToList $user $member.PrimarySmtpAddress $member.RecipientType $securityGroups[$key];
                    }
                }

                # Skip users with invalid recipient type
                if ($isUserTypeInvalid)
                {
                    continue;
                }

                ProcessPermissionEntry $user $UniqueUserList[$user] $securityGroups[$key] $owners $members;
            }
        }

        if ($WhatIf)
        {
            WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.AddingMembersAndOwners -f $group, $members, $owners);
            continue;
        }

        # Add members and owners to the group
        WriteLog -Path $logPath -LogOnly -Message ($LocalizedStrings.AddingMembersAndOwners -f $group, $members, $owners);
        if ($members -ne $null)
        {
            if ($ArePublicFoldersOnPremises)
            {
                Add-EXOUnifiedGroupLinks -Identity $group -LinkType Members -Links $members;
            }
            else
            {
                Add-UnifiedGroupLinks -Identity $group -LinkType Members -Links $members;
            }

            if (!$?)
            {
                WriteLog -Path $logPath -Level Error -LogOnly -Message $error[0];
            }
        }

        if ($owners -ne $null)
        {
            if ($ArePublicFoldersOnPremises)
            {
                Add-EXOUnifiedGroupLinks -Identity $group -LinkType Owners -Links $owners;
            }
            else
            {
                Add-UnifiedGroupLinks -Identity $group -LinkType Owners -Links $owners;
            }

            if (!$?)
            {
                WriteLog -Path $logPath -Level Error -LogOnly -Message $error[0];
            }
        }

        $default = $permissionEntries | ?{[string]$_.User -eq "Default"};
        if (IsAccessRightSufficient $default.AccessRights)
        {
            WriteLog -Path $logPath -Level Warn -Message ($LocalizedStrings.DefaultPermissionNotNone -f $pfIdentity, [string] $default.AccessRights, $group);
        }
    }
}
finally
{
    if ($script:session -ne $null)
    {
        Remove-PSSession $script:session;
    }
}

WriteLog -Path $logPath -Message $LocalizedStrings.AddingMembersSuccessful;
WriteLog -Path $logPath -Message $LocalizedStrings.CommandToAddMembers;
# SIG # Begin signature block
# MIIdmwYJKoZIhvcNAQcCoIIdjDCCHYgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/ljvaM+NPQjR6ynKqksWVUU/
# zWugghhTMIIEwjCCA6qgAwIBAgITMwAAAL+RbPt8GiTgIgAAAAAAvzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwOTA3MTc1ODQ5
# WhcNMTgwOTA3MTc1ODQ5WjCBsjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046
# NTdDOC0yRDE1LTFDOEIxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
# cnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCt7X+GwPaidVcV
# TRT2yohV/L1dpTMCvf4DHlCY0GUmhEzD4Yn22q/qnqZTHDd8IlI/OHvKhWC9ksKE
# F+BgBHtUQPSg7s6+ZXy69qX64r6m7X/NYizeK31DsScLsDHnqsbnwJaNZ2C2u5hh
# cKsHvc8BaSsv/nKlr6+eg2iX2y9ai1uB1ySNeunEtdfchAr1U6Qb7AJHrXMTdKl8
# ptLov67aFU0rRRMwQJOWHR+o/gQa9v4z/f43RY2PnMRoF7Dztn6ditoQ9CgTiMdS
# MtsqFWMAQNMt5bZ8oY1hmgkSDN6FwTjVyUEE6t3KJtgX2hMHjOVqtHXQlud0GR3Z
# LtAOMbS7AgMBAAGjggEJMIIBBTAdBgNVHQ4EFgQU5GwaORrHk1i0RjZlB8QAt3kX
# nBEwHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJ
# oEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMv
# TWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYB
# BQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
# 9w0BAQUFAAOCAQEAjt62jcZ+2YBqm7RKit827DRU9OKioi6HEERT0X0bL+JjUTu3
# 7k4piPcK3J/0cfktWuPjrYSuySa/NbkmlvAhQV4VpoWxipx3cZplF9HK9IH4t8AD
# YDxUI5u1xb2r24aExGIzWY+1uH92bzTKbAjuwNzTMQ1z10Kca4XXPI4HFZalXxgL
# fbjCkV3IKNspU1TILV0Dzk0tdKAwx/MoeZN1HFcB9WjzbpFnCVH+Oy/NyeJOyiNE
# 4uT/6iyHz1+XCqf2nIrV/DXXsJYKwifVlOvSJ4ZrV40MYucq3lWQuKERfXivLFXl
# dKyXQrS4eeToRPSevRisc0GBYuZczpkdeN5faDCCBgAwggPooAMCAQICEzMAAADD
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
# +6jMpF3BoYibV3FWTkhFwELJm3ZbCoBIa/15n8G9bW1qyVJzEw16UM0xggSyMIIE
# rgIBATCBlTB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgw
# JgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExAhMzAAAAww6b
# p9iy3PcsAAAAAADDMAkGBSsOAwIaBQCggcYwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkE
# MRYEFBG1cv1msWgjS77thK1TRzi61v8GMGYGCisGAQQBgjcCAQwxWDBWoC6ALABB
# AGQAZABNAGUAbQBiAGUAcgBzAFQAbwBHAHIAbwB1AHAAcwAuAHAAcwAxoSSAImh0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAE
# ggEAom9HnpVFrtmfQNusZBy3RrpOg/82NpWtY4xm6xuHL0RGoKnQLQPfTH0MFgpG
# suLEspWDw4YK7VkCW7TYTb27tIfgayG7m0l57t0IldYeJHCyX4DFFBfpF5CPPzfm
# ln38Q6/xQRjQU0EEh3iH1N9g7+LjDdXeOZdzwBFtOXSOsJgdwzsr7qXIbWgJMNAm
# K4VIPpEO9YN1/EfrcPh9S3twBRN8L29szctheyUxdvwP6NvRZSJT5rbG+T5ugDLV
# SBzYTj+43tQliYbSSRQGctDo3i+rNUkH4FcUw0SD336xggyfp20APZwxe9hrzN/C
# d2ufwTQ58U0SuhiDwMxbl7f6/aGCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIB
# ATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYD
# VQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAAC/kWz7fBok4CIAAAAA
# AL8wCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZI
# hvcNAQkFMQ8XDTE3MTAwNTA2MDAzMlowIwYJKoZIhvcNAQkEMRYEFIpCLssO421N
# +xmNTK1o1zhKw8ccMA0GCSqGSIb3DQEBBQUABIIBAHi2H6RTBTAZvjHWZO2UCkNA
# 6N5LiMpHItLJ9LH8tREf5GSroJ5hncX27+NYo2BeEzs58JCTvOdV9ZslZhDdnb6U
# fi6LDtpjyJRihibwsx3h5HkYEVJRFVLUSkc1fnc//OFAPogJSYkCkI1yH5NlU4YC
# xzsTL0VkDYp7LhnYWVRcme2cOEHasSyhzVxaO8T3GN0adlyTZS4yy840StqPcayl
# p9AKIh8r5giaoH6BhpeuBRLNaeQ7drvfYPfu8Fja/d7KQMDaUi9BRmnC1V14z+16
# 8MXvLVMC8ESk/f2PW2Eg2gpU+DTbrjPV7XrckxNwgrTOvk5rfdrKG8ntFPmCSTs=
# SIG # End signature block
