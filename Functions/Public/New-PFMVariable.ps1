Function New-PFMVariable
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification = "Creates in-memory object only."
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
    New-Variable -Scope Script -Name $name -Value $Value

}
