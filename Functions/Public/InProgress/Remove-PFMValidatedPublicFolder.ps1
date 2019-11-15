Function Remove-PFMValidatedPublicFolder
{
    [cmdletbinding()]
    [OutputType([pscustomobject])]
    param(
        [PublicFolderValidation]$PublicFolderValidation
    )
    $PublicFolderValidation
}