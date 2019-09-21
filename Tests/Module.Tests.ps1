$Script:ModuleRoot = (Split-Path -Path $PSScriptRoot -Parent)
Write-Information -MessageData "Module Root is $script:ModuleRoot" -InformationAction Continue
$Script:ModuleFile = $Script:ModuleFile = Get-Item $Script:ModuleRoot\*.psm1
Write-Information -MessageData "Module File is $($script:ModuleFile.FullName)" -InformationAction Continue
$Script:ModuleName = $Script:ModuleFile.BaseName
Write-Information -MessageData "Module Name is $Script:ModuleName" -InformationAction Continue
$script:ModuleFullPath = $Script:ModuleFile.FullName
Write-Information -MessageData "Removing Module $Script:ModuleName" -InformationAction Continue
Remove-Module -Name $Script:ModuleName -Force -ErrorAction SilentlyContinue
Write-Information -MessageData "Import Module $Script:ModuleName" -InformationAction Continue
Import-Module -Name $Script:ModuleFullPath -Force

Describe "All commands pass PSScriptAnalyzer rules" -Tag 'Build' {
    $rules = "$Script:ModuleRoot\PSScriptAnalyzerSettings.psd1"
    $scripts = Get-ChildItem -Path $ModuleRoot -Include '*.ps1', '*.psm1', '*.psd1' -Recurse |
        Where-Object -filterscript {$_.FullName -notmatch 'Classes' -and $_.FullName -notmatch 'Tests'}

    foreach ($script in $scripts)
    {
        Context $script.FullName {
            $results = Invoke-ScriptAnalyzer -Path $script.FullName -Settings $rules
            if ($results)
            {
                foreach ($rule in $results)
                {
                    It $("Should {0} Severity:{1} Line {2}: {3}" -f $rule.RuleName, $rule.Severity, $rule.Line, $rule.Message) {
                        $message = "violated"
                        $message | Should Be ""
                    }
                }
            }
            else
            {
                It "Should not fail any rules" {
                    $results | Should BeNullOrEmpty
                }
            }
        }
    }
}

Describe "Public commands have Pester tests" -Tag 'Build' {
    $commands = Get-Command -Module $ModuleName

    foreach ($command in $commands.Name)
    {
        $file = Get-ChildItem -Path "$ModuleRoot\Tests" -Include "$command.Tests.ps1" -Recurse
        It "Should have a Pester test for [$command]" {
            $file.FullName | Should Not BeNullOrEmpty
        }
    }
}

Write-Information -MessageData "Removing Module $Script:ModuleName" -InformationAction Continue
Remove-Module -Name $Script:ModuleName -Force -ErrorAction SilentlyContinue