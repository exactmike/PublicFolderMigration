Function Connect-PFMActiveDirectory
{
    <#
.SYNOPSIS
    Establishes a connection to the target Active Directory Forest
.DESCRIPTION
    Establishes a connection to the target Active Directory Forest and stores the PSSessionOption and Credential for re-connection or parallel connectdions (for parallel/long-running reporting operations where re-connection or additional connections to public folder servers will be necessary)
.PARAMETER DomainController
String parameter which requires the full FQDN of the on premises Active Directory Domain Controller server you want to use for PFM operations.
.PARAMETER Credential
Credential parameter requires a PSCredential object which will be used to connect to your AD organization and run get-* commands against public folders and related objects.  Best results if this account is a domain admin with enterprise admin membership. This credential will be stored in a module scope variable to support re-connection or parallel connections to public folder servers during long running operations.  The credential is removed with Remove-Module.
.PARAMETER PSSessionOption
PSSessionOption parameter accepts a PSSessionOption object to configure PSSessionOptions for environments where proxy options or other PSSessionOptions are required for successful Exchange Powershell connections.

.EXAMPLE
    PS C:\> $cred = get-credential
    PS C:\> Connect-PFMActiveDirectory -DomainController DomainController1.us.wa.contoso.com -credential $cred
    Connects to the Exchange On Premises server via Exchange Powershell and stores the PSSession for subsequent use by other PFM commands.  Stores the credential for reconnect or parallel connect scenarios.
.INPUTS
    Inputs
        [string] ExchangeOnPremisesServer
        [pscredential] AD Forest Enterprise ADmin
.OUTPUTS
    Output
        No direct output.
#>
    [CmdletBinding()]
    #[OutputType([System.Management.Automation.Runspaces.PSSession])]
    param
    (
        [parameter(Mandatory)]
        [ValidateScript( {
                if ($_ -match '(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+[a-zA-Z]{2,63}$)')
                { $true }
                else { Write-Warning -message "Parameter DomainController requires an FQDN"; $false }
            })]
        [string]$DomainController
        ,
        [parameter(Mandatory)]
        [pscredential]$Credential
        ,
        [System.Management.Automation.Remoting.PSSessionOption]$PSSessionOption
        ,
        [parameter()]
        [switch]$UseBasicAuth
    )

    #set module variables for credential and Active Directory
    $script:ADCredential = $Credential
    $script:DomainController = $DomainController

    #since this is user facing we always assume that if called the existing session needs to be replaced
    if ($null -ne $script:ADPSSession)
    {
        Remove-PSSession -Session $script:ADPSSession -ErrorAction SilentlyContinue
        $script:ADPSSession = $null
    }

    #BuildParamsToGetTheRequiredSession
    $GetPFMActiveDirectoryPSSessionParams = @{
        ErrorAction = 'Stop'
        Credential  = $script:ADCredential
    }
    if ($null -ne $PSSessionOption)
    {
        $script:PSSessionOption = $PSSessionOption
        $GetPFMActiveDirectoryPSSessionParams.PSSessionOption = $PSSessionOption
    }

    $GetPFMActiveDirectoryPSSessionParams.DomainController = $DomainController

    if ($true -eq $UseBasicAuth)
    {
        $GetPFMActiveDirectoryPSSessionParams.UseBasicAuth = $true
    }

    #Get the Required Domain Controller Session

    $ADPSSession = Get-PFMActiveDirectoryPSSession @GetPFMActiveDirectoryPSSessionParams
    $script:ADPSSession = $ADPSSession
    $script:ConnectActiveDirectoryCompleted = $true
    $script:ADForest = Invoke-Command -Session $Script:ADPSSession -ScriptBlock { Set-Location -Path 'AD:\'; Get-ADForest | Select-Object -ExpandProperty Name }
    Write-Information -MessageData "Connected to AD Forest $script:ADForest"
}
