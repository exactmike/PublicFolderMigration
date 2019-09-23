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
                    WriteLog -Message 'Establishing New PSSession to Exchange Organization' -EntryType Notification
                    $GetPFMExchangePSSessionParams = GetGetPFMExchangePSSessionParams
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