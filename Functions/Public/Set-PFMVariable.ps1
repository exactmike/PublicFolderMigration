Function Set-PFMVariable
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification = "Sets or Creates in-memory object only."
    )]
    [CmdletBinding(
        ConfirmImpact = 'None'
    )]
    param
    (
        [string]$Name
        ,
        $Value
    )
    Set-Variable -Scope Script -Name $Name -Value $value

}
