###############################################################################################
#Core Public Folder Migration Module Functions
###############################################################################################
$script:ConnectExchangeOrganizationCompleted = $false
$script:EmailConfiguration = $null
#Imports all the function files, public and private
$AllFunctionFiles = Get-ChildItem -Recurse -File -Path $(Join-Path -Path $PSScriptRoot -ChildPath 'Functions')
$PublicFunctionFiles = $AllFunctionFiles.where({$_.PSParentPath -like '*\Public'})
$PrivateFunctionFiles = $AllFunctionFiles.where({$_.PSParentPath -like '*\Private'})
foreach ($ff in $AllFunctionFiles) {. $ff.fullname}
Export-ModuleMember -Function $PublicFunctionFiles.foreach({$_.BaseName})
