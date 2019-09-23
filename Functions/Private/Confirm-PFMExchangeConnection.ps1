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
                    $Get-PFMExchangePSSessionParams = GetGet-PFMExchangePSSessionParams
                    $script:PsSession = Get-PFMExchangePSSession @Get-PFMExchangePSSessionParams
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