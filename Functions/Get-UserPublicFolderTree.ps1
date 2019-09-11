Function Get-UserPublicFolderTree
{

    [cmdletbinding()]
    param(
    )
    #Get All Public Folders
    $getPublicFolderParams = @{
            Recurse  = $true
            Identity = '\'
    }
    try
    {
        $message = "Get All Mail Enabled Public Folder Objects"
        Write-Information -Message $message -Tags Attempting
        Get-PublicFolder @getPublicFolderParams
        Write-Information -Message $message -Tags Succeeded
    }
    catch
    {
        $myerror = $_
        Write-Information -Message $message -Tags Errors
        Write-Information -Message $myerror.tostring() -Tags Errors
        $myerror
    }
}
