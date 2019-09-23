$CommandName = $MyInvocation.MyCommand.Name.Replace(".Tests.ps1", "")
$Script:ModuleRoot = (Split-Path -Path $PSScriptRoot -Parent)
Write-Information -MessageData "Module Root is $script:ModuleRoot" -InformationAction Continue
$Script:ModuleFile = $Script:ModuleFile = Get-Item $ModuleRoot\*.psm1
Write-Information -MessageData "Module File is $($script:ModuleFile.FullName)" -InformationAction Continue
$Script:ModuleName = $Script:ModuleFile.BaseName
Write-Information -MessageData "Module Name is $Script:ModuleName" -InformationAction Continue
$script:ModuleFullPath = $Script:ModuleFile.FullName
Write-Information -MessageData "Removing Module $Script:ModuleName" -InformationAction Continue
Remove-Module -Name $Script:ModuleName -Force -ErrorAction SilentlyContinue
Write-Information -MessageData "Import Module $Script:ModuleName" -InformationAction Continue
Import-Module -Name $Script:ModuleFullPath -Force

Describe "$CommandName Unit Tests" -Tag 'UnitTests' {
    Context "Validate parameters" {
        $defaultParamCount = 11
        [object[]]$params = (Get-ChildItem "function:\$CommandName").Parameters.Keys
        $knownParameters = 'ExchangeOnline','ExchangeOnPremisesServer','Credential','PSSessionOption','IsParallel'
        $paramCount = $knownParameters.Count
        It "Should contain specific parameters" {
            ( (Compare-Object -ReferenceObject $knownParameters -DifferenceObject $params -IncludeEqual | Where-Object SideIndicator -eq "==").Count ) | Should Be $paramCount
        }
        It "Should only contain $paramCount parameters" {
            $params.Count - $defaultParamCount | Should Be $paramCount
        }
    }
}

Write-Information -MessageData "Removing Module $Script:ModuleName" -InformationAction Continue
Remove-Module -Name $Script:ModuleName -Force -ErrorAction SilentlyContinue