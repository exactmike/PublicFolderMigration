Function Remove-PFMValidatedPublicFolder
{
    [cmdletbinding()]
    [OutputType([pscustomobject])]
    param(
        [parameter(Mandatory, ValueFromPipeline)]
        [PSTypeName("PublicFolderValidation")]$PublicFolderValidation
    )
    $PublicFolderValidation
}