Function Get-AllMailPublicFolder
{
    [cmdletbinding()]
    param(
    )
    #Get all mail enabled public folders
    $getMailPublicFolderParams = @{
        ResultSize    = 'Unlimited'
        ErrorAction   = 'stop'
        WarningAction = 'SilentlyContinue'
    }
    try
    {
        $message = "Get All Mail Enabled Public Folder Objects"
        Write-Information -Message $message -Tags Attempting
        Get-MailPublicFolder @getMailPublicFolderParams
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
