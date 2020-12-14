@{
    Rules = @{
        PSUseCompatibleSyntax = @{
            # This turns the rule on (setting it to false will turn it off)
            Enable         = $true

            # List the targeted versions of PowerShell here
            TargetVersions = @(
                '4.0',
                '5.1'
            )
        }
    }
}