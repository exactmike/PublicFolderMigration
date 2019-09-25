Function Connect-PFMExchange
{

    [CmdletBinding(DefaultParameterSetName = 'ExchangeOnPremises')]
    param
    (
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnline')]
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnlineParallel')]
        [switch]$ExchangeOnline
        ,
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnPremises')]
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnPremisesParallel')]
        [ValidatePattern('(?=^.{1,254}$)(^(?:(?!\d+\.|-)[a-zA-Z0-9_\-]{1,63}(?<!-)\.?)+(?:[a-zA-Z]{2,})$)')]
        [string]$ExchangeOnPremisesServer
        ,
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnPremises')]
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnline')]
        [pscredential]$Credential
        ,
        [System.Management.Automation.Remoting.PSSessionOption]$PSSessionOption
        ,
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnPremisesParallel')]
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnlineParallel')]
        [switch]$IsParallel
    )

    #Force user to run Connect-PFMExchange for organization before IsParallel
    if (
            ($null -eq $ConnectExchangeOrganizationCompleted -or $false -eq $ConnectExchangeOrganizationCompleted) -and
            $true -eq $IsParallel
        )
    {
        Write-ConnectPFMExchangeUserError
    }

    #set module variables for credential and exchange organization type if not IsParallel
    if ($true -ne $IsParallel)
    {
        $script:ExchangeCredential = $Credential
        $script:ExchangeOrganizationType = $script:ExchangeOrganizationType
    }

    #since this is user facing we always assume that if called the existing session needs to be replaced
    if ($false -eq $IsParallel -and $null -ne $script:PsSession)
    {
        Remove-PSSession -Session $script:PsSession -ErrorAction SilentlyContinue
        $script:PSSession = $null
    }

    #BuildParamsToGetTheRequiredSession
    $GetPFMExchangePSSessionParams = @{
        ErrorAction = 'Stop'
        Credential  = $script:ExchangeCredential
    }
    if ($null -ne $PSSessionOption)
    {
        $script:PSSessionOption = $PSSessionOption
        $GetPFMExchangePSSessionParams.PSSessionOption = $PSSessionOption
    }
    switch -Wildcard ($PSCmdlet.ParameterSetName)
    {
        'ExchangeOnline*'
        {
            $GetPFMExchangePSSessionParams.ExchangeOnline = $true
        }
        'ExchangeOnPremises*'
        {
            $GetPFMExchangePSSessionParams.ExchangeServer = $ExchangeOnPremisesServer
        }
    }

    #Get the Required Exchange Session
    $ExchangeSession = Get-PFMExchangePSSession @GetPFMExchangePSSessionParams

    Switch ($IsParallel)
    {
        $false
        {
            $script:PsSession = $ExchangeSession
            $script:ConnectExchangeOrganizationCompleted = $true
        }
        $true
        {
            Add-PFMParallelPSSession -Session $ExchangeSession
        }
    }
}
