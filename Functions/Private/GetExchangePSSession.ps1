    Function GetExchangePSSession
    {
        
        [CmdletBinding(DefaultParameterSetName = 'ExchangeOnline')]
        param
        (
            [parameter(Mandatory)]
            [pscredential]$Credential = $script:Credential
            ,
            [parameter(Mandatory,ParameterSetName = 'ExchangeOnline')]
            [switch]$ExchangeOnline
            ,
            [parameter(Mandatory,ParameterSetName = 'ExchangeOnPremises')]
            [string]$ExchangeServer
            ,
            [System.Management.Automation.Remoting.PSSessionOption]$PSSessionOption
        )
        $NewPsSessionParams = @{
            ErrorAction = 'Stop'
            ConfigurationName = 'Microsoft.Exchange'
            Credential = $Credential
        }
        switch ($PSCmdlet.ParameterSetName)
        {
            'ExchangeOnline'
            {
                $NewPsSessionParams.ConnectionURI = 'https://outlook.office365.com/powershell-liveid/'
                $NewPsSessionParams.Authentication = 'Basic'
            }
            'ExchangeOnPremises'
            {
                $NewPsSessionParams.ConnectionURI = 'http://' + $ExchangeServer + '/PowerShell/'
                $NewPsSessionParams.Authentication = 'Kerberos'
            }
        }
        $ExchangeSession = New-PSSession @NewPsSessionParams
        if ($PSCmdlet.ParameterSetName -eq 'ExchangeOnPremises')
        {
            Invoke-Command -Session $ExchangeSession -ScriptBlock {Set-ADServerSettings -ViewEntireForest $true -ErrorAction 'Stop'} -ErrorAction Stop
        }
        $ExchangeSession
    
    }

