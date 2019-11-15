Function Remove-PFMValidatedPublicFolder
{
    [cmdletbinding()]
    [OutputType([pscustomobject])]
    param(
        [parameter(Mandatory, ValueFromPipeline)]
        [PSTypeName("PublicFolderValidation")]$PublicFolderValidation
    )
    process
    {
        foreach ($pfv in $PublicFolderValidation)
        {
            $pfv
        }
    }
}