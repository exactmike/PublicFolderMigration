# PublicFolderMigration

PublicFolderMigration is intended to be a collection of functions to support the migration of Exchange Public Folders. At present it contains one function - Get-PublicFolderReplicationReport which is an almost complete re-write of the script hosted here: https://gallery.technet.microsoft.com/office/Exchange-2010-Public-944df6ee

## Improvements and How it Works
To be added: additional documentation for the reasons for a re-write of Get-PublicFolderReplicationReport.ps1.  
See the inline help for how to as well as the original documentation at https://gallery.technet.microsoft.com/office/Exchange-2010-Public-944df6ee
Replace Write-Log with your own logging function.  Write-Log requirement will be addressed in a future version and/or the dependency removed in a future release or branch.

## Development Plans

1. Move parameter validation into the parameters where practical/possible
2. Continue to improve separation of the data generated from the presentation format(s) provided.
3. Add public folder synchronization functions that improve on the Microsoft provided functions for synchronization of mail public folders to Exchange Online since those scripts have a number of unresolved problems.
4. Add additional functions for automation of various aspects of public folder migration to Exchange 2013 or later or Exchange Online. 

## Example
TBA

## License

PublicFolderMigration is released under the MIT License

