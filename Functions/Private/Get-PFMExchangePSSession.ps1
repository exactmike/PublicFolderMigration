Function Get-PFMExchangePSSession
{
    [CmdletBinding(DefaultParameterSetName = 'ExchangeOnline')]
    [OutputType([System.Management.Automation.Runspaces.PSSession])]
    param
    (
        [parameter(Mandatory)]
        [pscredential]$Credential
        ,
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnline')]
        [switch]$ExchangeOnline
        ,
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnPremises')]
        [string]$ExchangeServer
        ,
        [System.Management.Automation.Remoting.PSSessionOption]$PSSessionOption
    )
    $NewPsSessionParams = @{
        ErrorAction       = 'Stop'
        ConfigurationName = 'Microsoft.Exchange'
        Credential        = $Credential
    }
    if ($null -ne $PSSessionOption)
    {
        $NewPsSessionParams.PSSessionOption = $PSSessionOption
    }
    switch ($PSCmdlet.ParameterSetName)
    {
        'ExchangeOnline'
        {
            $NewPsSessionParams.ConnectionURI = 'https://outlook.office365.com/powershell-liveid/'
            $NewPsSessionParams.Authentication = 'Basic'
            $NewPsSessionParams.AllowRedirection = $True
            $NewPsSessionParams.Name = 'ExchangeOnline'
        }
        'ExchangeOnPremises'
        {
            $NewPsSessionParams.ConnectionURI = 'http://' + $ExchangeServer + '/PowerShell/'
            $NewPsSessionParams.Authentication = 'Kerberos'
            $NewPsSessionParams.Name = $ExchangeServer
        }
    }
    $ExchangeSession = New-PSSession @NewPsSessionParams
    if ($PSCmdlet.ParameterSetName -eq 'ExchangeOnPremises')
    {
        Invoke-Command -Session $ExchangeSession -ScriptBlock { Set-ADServerSettings -ViewEntireForest $true -ErrorAction 'Stop' } -ErrorAction Stop
    }
    $ExchangeSession

}
