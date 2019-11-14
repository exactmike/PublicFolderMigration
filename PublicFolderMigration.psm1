###############################################################################################
#Public Folder Migration Module Variables
###############################################################################################
$ModuleVariableNames = (
    'ConnectExchangeOrganizationCompleted',
    'ExchangeCredential',
    'EmailConfiguration',
    'ExchangeOrganizationType',
    'ExchangeOnPremisesServer',
    'ParallelPSSession',
    'PSSession',
    'PSSessionOption',
    'UseAlternateParallelism'
)
$ModuleVariableNames.ForEach( { Set-Variable -Scope Script -Name $_ -Value $null })

enum ExportDataOutputFormat { csv; json; xml; clixml }
enum EmptyFolderValidation { NoSubFolders; NotMailEnabled; NoItems }
enum FolderActivityTime { CreationTime; LastAccessTime; LastModificationTime; LastUserAccessTime; LastUserModificationTime }

###############################################################################################
#Public Folder Migration Module Functions
###############################################################################################
$AllFunctionFiles = Get-ChildItem -Recurse -File -Path $(Join-Path -Path $PSScriptRoot -ChildPath 'Functions')
#$PublicFunctionFiles = $AllFunctionFiles.where( { $_.PSParentPath -like '*\Public' })
#$PrivateFunctionFiles = $AllFunctionFiles.where( { $_.PSParentPath -like '*\Private' })
foreach ($ff in $AllFunctionFiles) { . $ff.fullname }

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