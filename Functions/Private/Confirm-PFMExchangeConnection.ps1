function Confirm-PFMExchangeConnection
{
    switch ($script:ConnectExchangeOrganizationCompleted)
    {
        $true
        {
            switch (Test-PFMExchangePSSession -PSSession $script:PSSession)
            {
                $true
                {
                    WriteLog -Message 'Using Existing PSSession' -EntryType Notification
                }
                $false
                {
                    WriteLog -Message 'Removing Existing Failed PSSession' -EntryType Notification
                    Remove-PSSession -Session $script:PsSession -ErrorAction SilentlyContinue
                    $script:PSSession = $null
                    WriteLog -Message 'Establishing New PSSession to Exchange Organization' -EntryType Notification
                    $GetPFMExchangePSSessionParams = @{
                        ErrorAction = 'Stop'
                        Credential  = $Script:ExchangeCredential
                    }
                    if ($null -ne $Script:PSSessionOption)
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
                            $GetPFMExchangePSSessionParams.$ExchangeServer = $ExchangeOnPremisesServer
                        }
                    }
                    $script:PsSession = Get-PFMExchangePSSession @GetPFMExchangePSSessionParams
                }
            }
        }
        $false
        {
            Write-ConnectPFMExchangeUserError
        }
        $null
        {
            Write-ConnectPFMExchangeUserError
        }
    }
}