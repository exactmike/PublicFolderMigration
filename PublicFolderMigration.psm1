###############################################################################################
#Core Public Folder Migration Module Functions
###############################################################################################
$script:ConnectExchangeOrganizationCompleted = $false
. $(Join-Path $PSScriptRoot 'PublicFunctions.ps1')
. $(Join-Path $PSScriptRoot 'PrivateFunctions.ps1')
. $(Join-Path $PSScriptRoot 'SupportingFunctions.ps1')
. $(Join-Path $PSScriptRoot 'ThirdPartyFunctions.ps1')
. $(Join-Path $PSScriptRoot 'ModuleVariableFunctions.ps1')
. $(Join-Path $PSScriptRoot '\Functions\Private\GetHTMLReportFunction.ps1')
