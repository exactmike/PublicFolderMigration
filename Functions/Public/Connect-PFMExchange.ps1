Function Connect-PFMExchange
{
    <#
.SYNOPSIS
    Establishes a connection to the target Exchange Organization
.DESCRIPTION
    Establishes a connection to the target Exchange Organization and stores the PSSessionOption and Credential for re-connection or parallel connectdions (for parallel/long-running reporting operations where re-connection or additional connections to public folder servers will be necessary)
.PARAMETER ExchangeOnline
Switch parameter used to indicate your exchange organization is in Exchange Online
.PARAMETER ExchangeOnPremisesServer
String parameter which requires the full FQDN of the on premises Exchange server you want to use for PFM operations.
.PARAMETER Credential
Credential parameter requires a PSCredential object which will be used to connect to your Exchange organization and run get-* commands against public folders and related objects. This credential will be stored in a module scope variable to support re-connection or parallel connections to public folder servers during long running operations.  The credential is removed with Remove-Module.
.PARAMETER PSSessionOption
PSSessionOption parameter accepts a PSSessionOption object to configure PSSessionOptions for environments where proxy options or other PSSessionOptions are required for successful Exchange Powershell connections.
.PARAMETER IsParallel
Intended for internal module use only, this parameter is used when creating one or more secondare Exchange PSSessions for public folder statistics operations performed in parallel by Get-PublicFolderReplicationReport.
.PARAMETER UseAlternateParallelism
Intended for internal module use only, this parameter is used when creating one or more secondare Exchange PSSessions for public folder statistics operations performed in parallel by Get-PublicFolderReplicationReport.

.EXAMPLE
    PS C:\> $cred = get-credential
    PS C:\> Connect-PFMExchange -ExchangeOnPremisesServer PublicFolderServer1.us.wa.contoso.com -credential $cred
    Connects to the Exchange On Premises server via Exchange Powershell and stores the PSSession for subsequent use by other PFM commands.  Stores the credential for reconnect or parallel connect scenarios.
.INPUTS
    Inputs
        [string] ExchangeOnPremisesServer
        [pscredential] Exchange Public Folder Administrator Credential
.OUTPUTS
    Output
        No direct output.
#>
    [CmdletBinding(DefaultParameterSetName = 'ExchangeOnPremises')]
    #[OutputType([System.Management.Automation.Runspaces.PSSession])]
    param
    (
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnline')]
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnlineParallel')]
        [switch]$ExchangeOnline
        ,
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnPremises')]
        [parameter(Mandatory, ParameterSetName = 'ExchangeOnPremisesParallel')]
        [ValidateScript( {
                if ($_ -match '(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+[a-zA-Z]{2,63}$)')
                { $true }
                else { Write-Warning -message "Parameter ExchangeOnPremisesServer requires an FQDN"; $false }
            })]
        [string]$ExchangeOnPremisesServer
        ,
        [parameter(ParameterSetName = 'ExchangeOnPremises')]
        [switch]$UseAlternateParallelism
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
        ,
        [parameter(ParameterSetName = 'ExchangeOnPremises')]
        [parameter(ParameterSetName = 'ExchangeOnPremisesParallel')]
        [switch]$UseBasicAuth
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
        $script:ExchangeOnPremisesServer = $ExchangeOnPremisesServer
        if ($true -eq $UseAlternateParallelism) { $script:UseAlternateParallelism = $true } else { $script:UseAlternateParallelism = $false }
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
    if ($true -eq $UseBasicAuth)
    {
        $GetPFMExchangePSSessionParams.UseBasicAuth = $true
    }

    #Get the Required Exchange Session
    Switch ($IsParallel)
    {
        $false
        {
            $ExchangeSession = Get-PFMExchangePSSession @GetPFMExchangePSSessionParams
            $script:PsSession = $ExchangeSession
            $script:ConnectExchangeOrganizationCompleted = $true
            switch -wildcard ($PSCmdlet.ParameterSetName)
            { 'ExchangeOnPremises*' { $Script:ExchangeOrganizationType = 'ExchangeOnPremises' } 'ExchangeOnline*' { $Script:ExchangeOrganizationType = 'ExchangeOnline' } }
            $script:ExchangeOrganization = Invoke-Command -Session $Script:PSSession -ScriptBlock { Get-OrganizationConfig | Select-Object -ExpandProperty Identity | Select-Object -ExpandProperty Name }
        }
        $true
        {
            $GetPFMExchangePSSessionParams.IsParallel = $true
            switch ($null -eq $script:ParallelPSSession)
            {
                $true
                {
                    $ExchangeSession = Get-PFMExchangePSSession @GetPFMExchangePSSessionParams
                    Add-PFMParallelPSSession -PSSession $ExchangeSession
                }
                $false
                {
                    $existingSessionIndex = (GetArrayIndexForProperty -array $script:ParallelPSSession -property Name -Value $ExchangeOnPremisesServer)
                    if ($null -ne $existingSessionIndex -and $existingSessionIndex -ne -1)
                    {
                        #There's an existing session
                        switch (Test-PFMExchangePSSession -PSSession $script:ParallelPSSession[$existingSessionIndex])
                        {
                            $true
                            { }#which is working
                            $false
                            {
                                #which is not working
                                $ExchangeSession = Get-PFMExchangePSSession @GetPFMExchangePSSessionParams
                                Add-PFMParallelPSSession -PSSession $ExchangeSession
                            }
                        }
                    }
                    else
                    {
                        #there's no existing session
                        $ExchangeSession = Get-PFMExchangePSSession @GetPFMExchangePSSessionParams
                        Add-PFMParallelPSSession -PSSession $ExchangeSession
                    }
                }
            }
        }
    }
}
