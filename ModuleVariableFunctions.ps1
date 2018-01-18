###############################################################################################
#Module Variable Functions
###############################################################################################
function Get-PFMVariable
    {
        param
        (
        [string]$Name
        )
            Get-Variable -Scope Script -Name $name
    }
#end function Get-PFMVariable
function Get-PFMVariableValue
    {
        param
        (
            [string]$Name
        )
            Get-Variable -Scope Script -Name $name -ValueOnly
    }
#end function Get-PFMVariableValue
function Set-PFMVariable
    {
        param
        (
            [string]$Name
            ,
            $Value
        )
        Set-Variable -Scope Script -Name $Name -Value $value
    }
#end function Set-PFMVariable    
function New-PFMVariable
    {
        param
        (
            [string]$Name
            ,
            $Value
        )
        New-Variable -Scope Script -Name $name -Value $Value
    }
#end function New-PFMVariable
function Remove-PFMVariable
    {
        param
        (
            [string]$Name
        )
        Remove-Variable -Scope Script -Name $name
    }
#end function Remove-PFMVariable