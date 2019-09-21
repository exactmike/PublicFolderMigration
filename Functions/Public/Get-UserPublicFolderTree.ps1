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
        Write-Information -MessageData $message -Tags Attempts
        Get-PublicFolder @getPublicFolderParams
        Write-Information -MessageData $message -Tags Successes
    }
    catch
    {
        $myerror = $_
        Write-Information -MessageData $message -Tags Errors
        Write-Information -MessageData $myerror.tostring() -Tags Errors
        $myerror
    }
}
