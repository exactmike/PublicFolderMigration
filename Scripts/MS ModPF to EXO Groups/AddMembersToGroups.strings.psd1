ConvertFrom-StringData @'
###PSLOC
MappingCsvNotFound = Public folder to Groups Mapping csv is not found. Please verify the provided path.
IncorrectCsv = The mapping csv is either empty or does not have the expected columns. Please ensure that it contains 'FolderPath' and 'TargetGroupMailbox' columns with appropriate values.
CredentialNotFound = Exchange Online credential not found. Please provide Exchange Online admin credential for the remote PowerShell login.
CreatingRemoteSession = Creating an Exchange Online remote PowerShell session...
FailedToCreateRemoteSession = Unable to create a remote PowerShell session to Exchange Online. The error is as follows: "{0}".
FailedToImportRemoteSession = Exchange Online remote PowerShell session could not be imported. The error is as follows: "{0}".
RemoteSessionCreatedSuccessfully = Exchange Online remote PowerShell session created successfully.
PermissionFileMissing = Back up permission file "PfPermissions.csv" missing in directory {0}! Please provide the correct path and try again. If public folders are not locked, please set 'ArePublicFoldersLocked' to '$false'.
ReadingPermissionsFromFile = Reading permissions from file, {0}.
ReadingPermissions = Getting permissions of public folders.
PermissionEntriesMissingInFile = No permission entries are found in file '{0}' for public folder '{1}'! Please check if the correct file exists in the path provided and try again. If public folders are not locked, please set 'ArePublicFoldersLocked' to '$false'.
PermissionEntriesMissing = No permission entries are found for the public folder '{0}'.
FolderHasOnlyDefaultPermissions = The public folder '{0}' has no permission entries for users, other than 'Default' and 'Anonymous'. Hence no members are added to the group '{1}'.
AddingMembersToGroup = Adding members and owners to the group, '{0}' based on the permission entries of public folder, {1}.
InvalidRecipientType = Skipping user with invalid recipient type! User-{0}; RecipientType-{1}; AccessRight-{2}.
UserSkipped = The user '{0}' has access rights '{1}' which isn't sufficient to be added as a member to the group, hence skipping the user.
InvalidUser = Skipping the user '{0}' with access right '{1}', as the user does not exist!
UserNameIsNull = User with DisplayName '{0}' not found (user name is null).
SecurityGroupHasNoMembers = The security group '{0}' has no members.
AddingMembersAndOwners =  Updating links of the group '{0}' by adding the following list of users as members and owners of the group respectively. Members - '{1}'; Owners - '{2}'.
DefaultPermissionNotNone = Default permission for the public folder '{0}' is '{1}', but only users with explicit permission entries were added as members to the group '{2}'. Please change the privacy setting of the group to 'Public' if you need it to be accessible to everyone.
AddingMembersSuccessful = All the users with explicit permissions (except None, FolderVisible, and CreateSubFolders) to access input public folders have been added as Owners/Members to the respective group successfully!
CommandToAddMembers = Note: Please use the following cmdlet to add new members to the group if required. 'Add-UnifiedGroupLinks -Identity <Group> -LinkType [Owners | Members] -Links <list of users>'.
###PSLOC
'@