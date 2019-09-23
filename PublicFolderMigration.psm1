###############################################################################################
#Public Folder Migration Module Variables
###############################################################################################
$ModuleVariableNames = (
    'ConnectExchangeOrganizationCompleted',
    'Credential',
    'EmailConfiguration',
    'ExchangeOrganizationType',
    'ExchangeOnPremisesServer',
    'ParallelPSSession',
    'PSSession',
    'PSSessionOption',
    'PublicFolderMailboxServerSessions'
)
$ModuleVariableNames.ForEach( { Set-Variable -Scope Script -Name $_ -Value $null })

###############################################################################################
#Public Folder Migration Module Functions
###############################################################################################
$AllFunctionFiles = Get-ChildItem -Recurse -File -Path $(Join-Path -Path $PSScriptRoot -ChildPath 'Functions')
#$PublicFunctionFiles = $AllFunctionFiles.where( { $_.PSParentPath -like '*\Public' })
#$PrivateFunctionFiles = $AllFunctionFiles.where( { $_.PSParentPath -like '*\Private' })
foreach ($ff in $AllFunctionFiles) { . $ff.fullname }
#Export-ModuleMember -Function $PublicFunctionFiles.foreach( { $_.BaseName })
