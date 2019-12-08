Function Get-PFMExchangePSSession
{
    [CmdletBinding()]
    [OutputType([System.Management.Automation.Runspaces.PSSession])]
    param
    (
        [parameter(Mandatory)]
        [pscredential]$Credential
        ,
        [parameter(Mandatory)]
        [string]$DomainController
        ,
        [System.Management.Automation.Remoting.PSSessionOption]$PSSessionOption
        ,
        [switch]$IsParallel
        ,
        [switch]$UseBasicAuth
    )
    $NewPsSessionParams = @{
        ErrorAction = 'Stop'
        Credential  = $Credential
    }
    if ($null -ne $PSSessionOption)
    {
        $NewPsSessionParams.PSSessionOption = $PSSessionOption
    }
    switch ($true -eq $UseBasicAuth)
    {
        $true
        { $NewPsSessionParams.Authentication = 'Basic' }
        $false
        { $NewPsSessionParams.Authentication = 'Kerberos' }
    }
    $NewPsSessionParams.Name = $DomainController
    $ActiveDirectorySession = New-PSSession @NewPsSessionParams
    $ADModuleImported = $false
    try
    {
        Invoke-Command -Session $ActiveDirectorySession -ScriptBlock { Import-Module ActiveDirectory -ErrorAction Stop } -ErrorAction Stop
        $ADModuleImported = $true
    }
    catch
    {
        throw ($_)
    }
    if ($true -eq $ADModuleImported)
    {
        $ActiveDirectorySession
    }
}
