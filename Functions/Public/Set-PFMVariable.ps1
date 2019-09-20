    Function Set-PFMVariable
    {
        
        param
        (
            [string]$Name
            ,
            $Value
        )
        Set-Variable -Scope Script -Name $Name -Value $value
    
    }

