Function Connect-ExchangeOrganization
{

    [CmdletBinding(DefaultParameterSetName = 'ExchangeOnline')]
    param
    (
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnline')]
        [switch]$ExchangeOnline
        ,
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnPremises')]
        [string]$ExchangeOnPremisesServer
        ,
        [parameter(Mandatory)]
        [pscredential]$Credential
        ,
        [System.Management.Automation.Remoting.PSSessionOption]$PSSessionOption
    )
    $script:Credential = $Credential
    #since this is user facing we always assume that if called the existing session needs to be replaced
    if ($null -ne $script:PsSession -and $script:PsSession -is [System.Management.Automation.Runspaces.PSSession])
    {
        Remove-PSSession -Session $script:PsSession -ErrorAction SilentlyContinue
    }
    $GetExchangePSSessionParams = @{
        ErrorAction = 'Stop'
        Credential  = $script:Credential
    }
    if ($null -ne $PSSessionOption)
    {
        $script:PSSessionOption = $PSSessionOption
        $GetExchangePSSessionParams.PSSessionOption = $script:PSSessionOption
    }
    switch ($PSCmdlet.ParameterSetName)
    {
        'ExchangeOnline'
        {
            $Script:OrganizationType = 'ExchangeOnline'
            $GetExchangePSSessionParams.ExchangeOnline = $true
        }
        'ExchangeOnPremises'
        {
            $Script:OrganizationType = 'ExchangeOnPremises'
            $Script:ExchangeOnPremisesServer = $ExchangeOnPremisesServer
            $GetExchangePSSessionParams.ExchangeServer = $script:ExchangeOnPremisesServer
        }
    }
    $script:PsSession = GetExchangePSSession @GetExchangePSSessionParams
    $script:ConnectExchangeOrganizationCompleted = $true

}
