Function Get-PFMActiveDirectoryPSSession
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
        [switch]$UseBasicAuth
    )
    $NewPsSessionParams = @{
        ErrorAction = 'Stop'
        Credential  = $Credential
        EnableNetworkAccess = $true
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
    $NewPSSessionParams.Computer = $DomainController
    $ActiveDirectorySession = New-PSSession @NewPsSessionParams
    $ADModuleImported = $false
    $GCDriveCreated = $false
    try
    {
        Invoke-Command -Session $ActiveDirectorySession -ScriptBlock { Import-Module ActiveDirectory -ErrorAction Stop } -ErrorAction Stop
        $ADModuleImported = $true
    }
    catch
    {
        throw ($_)
    }
    try
    {
        $null = Invoke-Command -Session $ActiveDirectorySession -ScriptBlock { New-PSDrive -PSProvider ActiveDirectory -GlobalCatalog -Root '' -Name 'GC' -ErrorAction Stop } -ErrorAction Stop
        $GCDriveCreated = $true
    }
    catch
    {
        throw ($_)
    }
    if ($true -eq $ADModuleImported -and $true -eq $GCDriveCreated)
    {
        $ActiveDirectorySession
    }
}
