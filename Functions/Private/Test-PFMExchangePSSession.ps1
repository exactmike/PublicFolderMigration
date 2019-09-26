Function Test-PFMExchangePSSession
{

    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]$PSSession
    )
    switch ($PSSession.State -eq 'Opened')
    {
        $true
        {
            Try
            {
                $TestCommandResult = invoke-command -Session $PSSession -ScriptBlock { Get-OrganizationConfig -ErrorAction Stop | Select-Object -ExpandProperty Identity | Select-Object -ExpandProperty Name } -ErrorAction Stop
                $(-not [string]::IsNullOrEmpty($TestCommandResult))
            }
            Catch
            {
                $false
            }
        }
        $false
        {
            $false
        }
    }

}
