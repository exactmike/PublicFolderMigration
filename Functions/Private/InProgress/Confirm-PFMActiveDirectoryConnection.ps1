function Confirm-PFMActiveDirectoryConnection
{
    [cmdletbinding()]
    Param(
        [parameter(Mandatory)]
        [AllowNull()]
        [System.Management.Automation.Runspaces.PSSession]$PSSession
    )
    switch ($script:ConnectActiveDirectoryCompleted)
    {
        $true
        {
            switch (Test-PFMActiveDirectoryPSSession -PSSession $PSSession)
            {
                $false
                {
                    #Remove storage of the existing session
                    WriteLog -Message "Removing Existing Failed PSSession: $($PSSession.Name)" -EntryType Notification
                    Remove-PSSession -Session $PsSession -ErrorAction SilentlyContinue
                    $script:PSSession = $null
                    #establish a new session
                    WriteLog -Message "Establishing New PSSession to $($PSSession.Name)" -EntryType Notification
                    $GetPFMActiveDirectoryPSSessionParams = @{
                        ErrorAction = 'Stop'
                        Credential  = $Script:ADCredential
                    }
                    if ($null -ne $Script:PSSessionOption)
                    {
                        $GetPFMActiveDirectoryPSSessionParams.PSSessionOption = $script:PSSessionOption
                    }
                    $GetPFMActiveDirectoryPSSessionParams.DomainController = $PSSession.Name
                    $NewPSSession = Get-PFMActiveDirectoryPSSession @GetPFMActiveDirectoryPSSessionParams
                    #Update storage of the updated session
                    $Script:ADPSSession = $NewPSSession
                }
            }
        }
        $false
        {
            Write-ConnectPFMActiveDirectoryUserError
        }
        $null
        {
            Write-ConnectPFMActiveDirectoryUserError
        }
    }
}