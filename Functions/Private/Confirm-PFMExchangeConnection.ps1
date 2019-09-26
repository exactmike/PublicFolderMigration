function Confirm-PFMExchangeConnection
{
    [cmdletbinding(DefaultParameterSetName = 'OrganizationConnection')]
    Param(
        [parameter(Mandatory, ParameterSetName = 'ParallelConnection')]
        [switch]$IsParallel
        ,
        [parameter(Mandatory, ParameterSetName = 'ParallelConnection')]
        [parameter(Mandatory, ParameterSetName = 'OrganizationConnection')]
        [AllowNull()]
        [System.Management.Automation.Runspaces.PSSession]$PSSession
    )
    switch ($script:ConnectExchangeOrganizationCompleted)
    {
        $true
        {
            switch (Test-PFMExchangePSSession -PSSession $PSSession)
            {
                $true
                {
                    WriteLog -Message 'Using Existing PSSession' -EntryType Notification
                }
                $false
                {
                    #Remove storage of the existing session
                    WriteLog -Message 'Removing Existing Failed PSSession' -EntryType Notification
                    switch ($PSCmdlet.ParameterSetName)
                    {
                        'OrganizationConnection'
                        {
                            Remove-PSSession -Session $PsSession -ErrorAction SilentlyContinue
                            $script:PSSession = $null
                        }
                        'ParallelConnection'
                        {
                            #nothing at this stage Add-PFMParallelPSSession does the work to update $script:ParalellPSSession
                        }
                    }
                    #establish a new session
                    WriteLog -Message "Establishing New PSSession to $($PSSession.Name)" -EntryType Notification
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
                            $GetPFMExchangePSSessionParams.ExchangeServer = $PSSession.Name
                        }
                    }
                    $NewPSSession = Get-PFMExchangePSSession @GetPFMExchangePSSessionParams
                    #Update storage of the updated session
                    switch ($PSCmdlet.ParameterSetName)
                    {
                        'OrganizationConnection'
                        {
                            $Script:PSSession = $NewPSSession
                        }
                        'ParallelConnection'
                        {
                            Add-PFMParallelPSSession -PSSession $NewPSSession
                        }
                    }
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