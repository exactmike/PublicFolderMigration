Function Get-PFMVariableValue
{

    param
    (
        [string]$Name
    )
    Get-Variable -Scope Script -Name $name -ValueOnly

}
