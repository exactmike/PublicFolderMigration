Function Remove-PFMValidatedPublicFolder
{
    [cmdletbinding()]
    [OutputType([pscustomobject])]
    param(
        [PSTypeName("PublicFolderValidation")]$PublicFolderValidation
    )
    $PublicFolderValidation
}