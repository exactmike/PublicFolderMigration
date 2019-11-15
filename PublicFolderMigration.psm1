#Requires -Version 5.1
###############################################################################################
#Public Folder Migration Module Variables
###############################################################################################
$ModuleVariableNames = (
    'ConnectExchangeOrganizationCompleted',
    'ExchangeCredential',
    'EmailConfiguration',
    'ExchangeOrganizationType',
    'ExchangeOrganization',
    'ExchangeOnPremisesServer',
    'ParallelPSSession',
    'PSSession',
    'PSSessionOption',
    'UseAlternateParallelism',
    'PublicFolderPropertyList'
)
$ModuleVariableNames.ForEach( { Set-Variable -Scope Script -Name $_ -Value $null })

enum ExportDataOutputFormat { csv; json; xml; clixml }
enum FolderValidation { NoSubFolders; NotMailEnabled; NoItems }
enum FolderActivityTime { CreationTime; LastAccessTime; LastModificationTime; LastUserAccessTime; LastUserModificationTime }
enum ItemActivityTime { LastUserModificationTime; LastUserAccessTime; CreationTime }

###############################################################################################
#Public Folder Migration Module Functions
###############################################################################################
$AllFunctionFiles = Get-ChildItem -Recurse -File -Path $(Join-Path -Path $PSScriptRoot -ChildPath 'Functions')
#$PublicFunctionFiles = $AllFunctionFiles.where( { $_.PSParentPath -like '*\Public' })
#$PrivateFunctionFiles = $AllFunctionFiles.where( { $_.PSParentPath -like '*\Private' })
foreach ($ff in $AllFunctionFiles) { if ($ff.fullname -like '*.ps1') { . $ff.fullname } }

###############################################################################################
#Public Folder Migration Module Removal Routines
###############################################################################################
#Clean up objects that will exist in the Global Scope due to no fault of our own . . .
$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove =
{
    if ($null -ne $Script:PSSession) { Remove-PSSession -Session $script:Pssession }
    if ($null -ne $script:ParallelPSSession)
    {
        foreach ($session in $script:ParallelPSSession)
        {
            Remove-PSSession -Session $session
        }
    }
}