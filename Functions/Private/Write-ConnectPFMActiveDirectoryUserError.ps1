Function Write-ConnectPFMActiveDirectoryUserError
{

    $message = "You must call the Connect-PFMActiveDirectory function  before calling any other cmdlets which require an active Active Directory connection."
    throw($message)

}
