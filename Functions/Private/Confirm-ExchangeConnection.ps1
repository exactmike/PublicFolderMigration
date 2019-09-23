function Confirm-ExchangeConnection
{
    switch ($script:ConnectExchangeOrganizationCompleted)
    {
        $true
        {
            switch (TestExchangePSSession -PSSession $script:PSSession)
            {
                $true
                {
                    WriteLog -Message 'Using Existing PSSession' -EntryType Notification
                }
                $false
                {
                    WriteLog -Message 'Removing Existing Failed PSSession' -EntryType Notification
                    Remove-PSSession -Session $script:PsSession -ErrorAction SilentlyContinue
                    WriteLog -Message 'Establishing New PSSession to Exchange Organization' -EntryType Notification
                    $GetExchangePSSessionParams = GetGetExchangePSSessionParams
                    $script:PsSession = GetExchangePSSession @GetExchangePSSessionParams
                }
            }
        }
        $false
        {
            WriteUserInstructionError
        }
        $null
        {
            WriteUserInstructionError
        }
    }
}