Function Connect-PFMExchange
{

    [CmdletBinding(DefaultParameterSetName = 'ExchangeOnPremises')]
    param
    (
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnline')]
        [switch]$ExchangeOnline
        ,
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnPremises')]
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnPremisesParallel')]
        [ValidatePattern('(?=^.{1,254}$)(^(?:(?!\d+\.|-)[a-zA-Z0-9_\-]{1,63}(?<!-)\.?)+(?:[a-zA-Z]{2,})$)')]
        [string]$ExchangeOnPremisesServer
        ,
        [parameter(Mandatory)]
        [pscredential]$Credential
        ,
        [System.Management.Automation.Remoting.PSSessionOption]$PSSessionOption
        ,
        [parameter(ParameterSetName = 'ExchangeOnPremisesParallel')]
        [switch]$IsParallel
    )
    $script:Credential = $Credential
    #since this is user facing we always assume that if called the existing session needs to be replaced
    if ($false -eq $IsParallel -and $null -ne $script:PsSession -and $script:PsSession -is [System.Management.Automation.Runspaces.PSSession])
    {
        Remove-PSSession -Session $script:PsSession -ErrorAction SilentlyContinue
    }
    $GetPFMExchangePSSessionParams = @{
        ErrorAction = 'Stop'
        Credential  = $script:Credential
    }
    if ($null -ne $PSSessionOption)
    {
        $script:PSSessionOption = $PSSessionOption
        $GetPFMExchangePSSessionParams.PSSessionOption = $script:PSSessionOption
    }
    switch ($PSCmdlet.ParameterSetName)
    {
        'ExchangeOnline'
        {
            $Script:ExchangeOrganizationType = 'ExchangeOnline'
            $GetPFMExchangePSSessionParams.ExchangeOnline = $true
        }
        'ExchangeOnPremises'
        {
            $Script:ExchangeOrganizationType = 'ExchangeOnPremises'
            $Script:ExchangeOnPremisesServer = $ExchangeOnPremisesServer
            $GetPFMExchangePSSessionParams.ExchangeServer = $script:ExchangeOnPremisesServer
        }
    }
    Switch ($IsParallel)
    {
        $false
        {
            $script:PsSession = Get-PFMExchangePSSession @GetPFMExchangePSSessionParams
            $script:ConnectExchangeOrganizationCompleted = $true
        }
        $true
        {
            if ($null -eq $script:ParallelPSSession)
            {$script:ParallelPSSession = [System.Collections.ArrayList]::new()}
            [void]$script:ParallelPSSession.Add($(Get-PFMExchangePSSession @Get-PFMExchangePSSessionParams))
        }
    }

}
