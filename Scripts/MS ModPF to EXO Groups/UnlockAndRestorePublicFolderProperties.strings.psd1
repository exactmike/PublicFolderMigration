ConvertFrom-StringData @'
###PSLOC
WhatIfEnabled = IMPORTANT!!! WhatIf parameter is set, therefore no changes will be made. The messages shown are just a preview of the actual execution. All backup files are expected with a 'test_' prefix. If actual lockdown is performed, please create a copy of the backup files with prefix 'text_' to see the preview of this script.
UnsuccessfulGetRecipientPermissionCmdlet = Get-RecipientPermission cmdlet could not be executed. Please make sure the user is a member of role groups 'Organization Management' and 'Recipient Management' and try again.
BackupNotFound = Following backup files for recovery does not exists: {0}. Please move the files to backup directory and re-run the script.
ReadingPfPerms = Reading public folder permissions..
IncorrectCsv = Incorrect csv {0} provided.
ImportingBackupFilesSuccessful = Successfully imported {0} and {1}.
RestoringPfPerms = Restoring public folder permissions from backup file ({0}) ..
CredentialNotFound = Exchange Online credential not found. Please provide Exchange Online admin credential for the remote PowerShell login.
CreatingRemoteSession = Creating an Exchange Online remote PowerShell session...
FailedToCreateRemoteSession = Unable to create a remote PowerShell session to Exchange Online. The error is as follows: '{0}'.
FailedToImportRemoteSession = Exchange Online remote PowerShell session could not be imported. The error is as follows: '{0}'.
RemoteSessionCreatedSuccessfully = Exchange Online remote PowerShell session created successfully.
SkippingUser = Skipping permissions of user {0} as user is not found. Get-Recipient failed with this user value obtained from backup file.
RemovingPfPermission = Removing permission of user {0} on public folder {1}.
AddPfPermission = Adding permissions '{1}' to user {0} on public folder {2}.
RestoredPfPerms = Successfully restored all public folder permissions from backup file.
MailEnablingPfs = Mail enabling all migrating public folders..
MailEnabledAndRestoredProperties = Mail enabled {0} and restored mail properties.
RemovedPropertiesFromGroup = The following properties are removed from group '{0}': SendAs permissions of users {1}, SendOnBehalfTo permissions of users {2}, SMTP addresses {3}.
MailEnabledPf = Public folder {0} is mail enabled.
RemovingSendAsFromGroupFailed = Removing SendAs permission from group {0} failed..
SendAsPermissionRemovedFromGroup = SendAs permission of {0} have been removed from group {1}.
RestoringSMTPFailed = Restoring SMTP address of public folder {0} failed..
RestoringSMTPSucceeded = Successfully Restored SMTP address of public folder {0}.
AddingSendOnBehalfToPermissionFailed = Adding SendOnBehalfTo permission of user '{1}' to public folder '{0}' is failed. 
AddingSendAsToPfFailed = Adding SendAs permission to public folder {0} failed..
AddedPropertiesBackToPf = The following properties are added back to public folder '{0}': SendAs permissions of users {1}, SendOnBehalfTo permissions of users {2}, emailAddressPolicyEnabled {3}.
PfRecoveryComplete = Public folders successfully restored.
###PSLOC
'@