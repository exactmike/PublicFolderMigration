ConvertFrom-StringData @'
###PSLOC
WhatIfEnabled = IMPORTANT!!! WhatIf parameter is set, therefore no changes will be made. The messages shown are just a preview of the actual execution. All backup files are created with a 'test_' prefix.
UnsuccessfulGetRecipientPermissionCmdlet = Get-RecipientPermission cmdlet could not be executed. Please make sure the user is a member of role groups 'Organization Management' and 'Recipient Management' and try again.
MappingCsvNotFound = Public folder to group mapping csv is not found. Please verify the provided path {0}.
IncorrectCsvFormat = The mapping csv is either empty or does not have the expected columns. Please ensure that it contains 'FolderPath' and 'TargetGroupMailbox' columns with appropriate values.
BackupCsvAlreadyExist = Backup files already exist. Preventing further execution, as the original permissions of the public folders can be permanently lost. Please provide a different backup location and try again.
ExportingPFPermissions = Exporting public folder permissions..
PfsAlreadyInLockedState = Public folders being migrated are already in locked down state. No action is performed on public folders.
WarnBackupFilesNotFound = Public folder permissions backup file is not found ({0}). Restoring permissions is not possible.
CredentialNotFound = Exchange Online credential not found. Please provide Exchange Online admin credential for the remote PowerShell login.
CreatingRemoteSession = Creating an Exchange Online remote Powershell session...
FailedToCreateRemoteSession = Unable to create a remote PowerShell session to Exchange Online. The error is as follows: '{0}'.
FailedToImportRemoteSession = Exchange Online remote Powershell session could not be imported. The error is as follows: '{0}'.
RemoteSessionCreatedSuccessfully = Exchange Online remote Powershell session created successfully.
ExportPermissionsSuccessful = Successfully saved public folder permissions to {0}.
ExportMailPropertiesSuccessful = Mail properties of mail enabled public folders are successfully saved to {0}.
SkippingUser = Skipping permissions of user {0} as user is not found. The user value in public folder client permission entry could not be resolved.
SkippingNotMigratedUser = Skipping permissions of user {0} as user is not found in Exchange Online.
PfMailDisabled = Public folder {0} is mail disabled.
PfPropertiesCopiedToGroup = The following properties are copied from public folder {0} to Group {1}: SMTP addresses {2}, send on behalf to permission to users {3}.
SettingSMTPToGroupFailed = Setting SMTP address to group failed..
SMTPAddressesCopiedFromMailPfToGroup = SMTP addresses of {0} have been added as proxy addresses to group {1}.
SendAsPermsCopiedToGroup = The SendAs permissions of following users are copied from public folder {0} to group {1}: {2}.
AddingSendAsToGroupFailed = Adding SendAs permission to group failed..
ExportMailPfPropertiesAndGroupSuccessful = Mail properties of mail enabled public folders, along with the groups to which those properties are exported, are successfully saved to {0}.
LockingPfsByRemovingPerms = Locking down migrating public folders by removing permissions..
RemovingPfPerm = Removing permissions of user {0} on public folder {1}.
AddingPfPerm = Adding permissions ({2}) for user {0} on {1}.
LockdownWithReadOnlyPermsSuccessful = Public folders being migrated are successfully locked down with ReadOnly permission.
PfLockdownComplete = Public folder lockdown is complete.
###PSLOC
'@