    Function Get-PFMVariable
    {
        
        param
        (
        [string]$Name
        )
            Get-Variable -Scope Script -Name $name
    
    }

