    Function Remove-PFMVariable
    {
        
        param
        (
            [string]$Name
        )
        Remove-Variable -Scope Script -Name $name
    
    }

