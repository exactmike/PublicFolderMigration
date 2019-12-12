# PublicFolderManagement

PublicFolderManagement is intended to be a collection of functions to support the management and/or migration of Exchange Public Folders. It presents the following functions:

As a historical note, PublicFolderManagement started out as a re-write of Get-PublicFolderReplicationReport.ps1 from the Technet Sript Gallery. https://gallery.technet.microsoft.com/office/Exchange-2010-Public-944df6ee.

## Improvements that inspired this module

Some of the reasons for a re-write of Get-PublicFolderReplicationReport.ps1 that inspired this module in the first place were:

- Improve performance in large public folder tree environments (several of the items below accomplish this)
- Automatically target only mailbox servers with a public folder database if values for the PublicFolderMailboxServer are not provided.
- Rename parameters such as ComputerName and FolderPath because of potential ambiguity of the purpose of the parameter.
- Provide capability to include system public folders when processing the entire public folder tree
- Eliminate redundant retrievals of the Public Folder Tree and make the build of the Public Folder Tree array more efficient
- Allow for public folder statistics retrieval for a subset of public folders but still retrieve all statistics for faster performance in case the user has requested the complete tree.
- improve interactive usage by reporting progress via logging to screen/log file and/or write-progress
- resolve issues with statistics calculation and reporting due to serialization of data in powershell remoting situations
- provide additional data points per public folder and in aggregate reporting, such as Last Modification Time, DatabaseName and ServerName; included servers, databases, folders; identification of empty/container public folders, etc.
- separate data produced from the presentation of the data in html

## Development Plans

1. Add functions that improve on the Microsoft provided functions for synchronization and migration of mail public folders to modern public folders in Exchange On Premises or Exchange Online since those scripts have a number of unresolved problems.
2. Add additional functions for automation of various aspects of public folder migration to modern public folders.

## Usage Example

Connect-PFMExchange -Credential $admincred -ExchangeOnPremisesServer 'pfdatabaseserver.au.contoso.com'
$pftree = Get-PFMPublicFolderTree -OutputFolderPath d:\Reports -OutputFormat csv,json -encoding utf8 -passthru
Get-PFMPublicFolderDatabase -outputfolderpath d:\Reports -OutputFormat csv,xml
Get-PFMPublicFolderStat -PublicFolderInfoObject $pftree -outputFolderPath d:\reports -outputformat json,csv -encoding utf8
Connect-PFMActiveDirectory -Credential $admincred -DomainController 'dc22.au.contoso.com'
Get-PFMPublicFolderPermission -OutputFolderPath d:\Reports -includeSendAs $true -includeSendOnBehalf $true -DropInheritedPermissions $true -IncludeSIDHistory -ExpandGroups $false

## License

PublicFolderMigration is released under the MIT License
