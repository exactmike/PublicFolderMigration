$Script:ModuleRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$Script:ModuleFile = $Script:ModuleFile = Get-ChildItem $ModuleRoot\*\*.psm1
$Script:ModuleName = $Script:ModuleFile.BaseName
$script:ModuleFullPath = $Script:ModuleFile.FullName
$Script:ModuleName
$script:ModuleFullPath
Remove-Module -Name $Script:ModuleName -Force -ErrorAction SilentlyContinue
Import-Module -Name $Script:ModuleFullPath -Force
Describe "Public commands have comment-based or external help" -Tags 'Build' {
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