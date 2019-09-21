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

Describe "Public commands have Pester tests" -Tag 'Build' {
    $commands = Get-Command -Module $Script:ModuleName

    foreach ($command in $commands.Name)
    {
        $file = Get-ChildItem -Path "$Script:ModuleRoot\Tests" -Include "$command.Tests.ps1" -Recurse
        It "Should have a Pester test for [$command]" {
            $file.FullName | Should Not BeNullOrEmpty
        }
    }
}

Write-Information -MessageData "Removing Module $Script:ModuleName" -InformationAction Continue
Remove-Module -Name $Script:ModuleName -Force -ErrorAction SilentlyContinue