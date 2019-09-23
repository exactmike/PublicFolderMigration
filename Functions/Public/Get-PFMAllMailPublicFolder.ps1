Function Get-PFMAllMailPublicFolder
{
    [cmdletbinding()]
    [OutputType([System.Object[]])]
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
        Write-Information -MessageData $message -Tags Attempts
        Get-MailPublicFolder @getMailPublicFolderParams
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
