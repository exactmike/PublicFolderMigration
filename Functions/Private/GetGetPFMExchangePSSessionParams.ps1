Function GetGetPFMExchangePSSessionParams
{

    $GetPFMExchangePSSessionParams = @{
        ErrorAction = 'Stop'
        Credential  = $script:Credential
    }
    if ($null -ne $script:PSSessionOption -and $script:PSSessionOption -is [System.Management.Automation.Remoting.PSSessionOption])
    {
        $GetPFMExchangePSSessionParams.PSSessionOption = $script:PSSessionOption
    }
    switch ($Script:ExchangeOrganizationType)
    {
        'ExchangeOnline'
        {
            $GetPFMExchangePSSessionParams.ExchangeOnline = $true
        }
        'ExchangeOnPremises'
        {
            $GetPFMExchangePSSessionParams.ExchangeServer = $script:ExchangeOnPremisesServer
        }
    }
    $GetPFMExchangePSSessionParams

}
