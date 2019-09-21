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
Describe "Public commands in $script:ModuleName have comment-based or external help" -Tags 'Build' {
    $functions = Get-Command -Module $Script:ModuleName
    $help = foreach ($function in $functions) {
        Get-Help -Name $function.Name
    }

    foreach ($node in $help)
    {
        Context $node.Name {
            It "Should have a Description or Synopsis" {
                ($node.Description + $node.Synopsis) | Should Not BeNullOrEmpty
            }

            It "Should have an Example"  {
                $node.Examples | Should Not BeNullOrEmpty
                $node.Examples | Out-String | Should -Match ($node.Name)
            }

            foreach ($parameter in $node.Parameters.Parameter)
            {
                if ($parameter -notmatch 'WhatIf|Confirm')
                {
                    It "Should have a Description for Parameter [$($parameter.Name)]" {
                        $parameter.Description.Text | Should Not BeNullOrEmpty
                    }
                }
            }
        }
    }
}

Write-Information -MessageData "Removing Module $Script:ModuleName" -InformationAction Continue
Remove-Module -Name $Script:ModuleName -Force -ErrorAction SilentlyContinue