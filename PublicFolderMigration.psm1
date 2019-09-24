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
    'PSSessionOption'
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


###############################################################################################
#Public Folder Migration Module Removal Routines
###############################################################################################
#Clean up objects that will exist in the Global Scope due to no fault of our own . . .
$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove =
{
    if ($null -ne $Script:PSSession) {Remove-PSSession -Session $script:Pssession}
    if ($null -ne $script:ParallelPSSession)
    {
        foreach ($session in $script:ParallelPSSession)
        {
            Remove-PSSession -Session $session
        }
    }
}