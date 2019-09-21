Function Remove-PFMVariable
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification = "Removes in-memory object only."
    )]
    [CmdletBinding(
        ConfirmImpact = 'None'
    )]
    param
    (
        [string]$Name
    )
    Remove-Variable -Scope Script -Name $name

}
