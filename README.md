# PublicFolderMigration

PublicFolderMigration is intended to be a collection of functions to support the migration of Exchange Public Folders. At present it contains the following functions:

Get-PublicFolderReplicationReport

    Based on Get-PublicFolderReplicationReport.ps1 from the Technet Sript Gallery.  Get-PublicFolderReplicationReport retains HTML formatting, parameters (with some exceptions), and some logical flow/progression with [Get-PublicFolderReplicationReport.ps1](https://gallery.technet.microsoft.com/office/Exchange-2010-Public-944df6ee) but is otherwise extensively revised.

Find-OrphanedMailEnabledPublicFolders

    Compares Recipient Objects from Get-MailPublicFolder with the Public Folder Tree to identify Mail Enabled Public Folder Recipient Objects that are no longer associated with a Public Folder.  These orphaned objects need to be fixed or removed before a Public Folder Migration can be completed successfully.

Get-AllMailPublicFolder

Get-UserPublicFolderTree

Export-MailbPublicFolder

Export-UserPublicFolderTree

## Improvements and How it Works

Some of the reasons for a re-write of Get-PublicFolderReplicationReport.ps1 to the replacement function Get-PublicFolderReplicationReport

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

See the inline help for how to as well as [the original documentation](https://gallery.technet.microsoft.com/office/Exchange-2010-Public-944df6ee).

## Development Plans

1. Move parameter validation into the parameters where practical/possible
2. Continue to improve separation of the data generated from the presentation format(s) provided.
3. Add public folder synchronization functions that improve on the Microsoft provided functions for synchronization of mail public folders to Exchange Online since those scripts have a number of unresolved problems.
4. Add additional functions for automation of various aspects of public folder migration to Exchange 2013 or later or Exchage Online.
5. Add options to Get-PublicFolderReplicationReport for output of csv, json, or xml data for further analysis

## Example

TBA

## License

PublicFolderMigration is released under the MIT License
