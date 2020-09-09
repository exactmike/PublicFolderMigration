Function Write-ConnectPFMExchangeUserError
{

    $message = "You must call the Connect-PFMExchange function (without the IsParallel parameter) before calling any other cmdlets which require an active Exchange Organization connection."
    throw($message)

}
