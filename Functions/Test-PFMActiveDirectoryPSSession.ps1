Function Test-PFMActiveDirectoryPSSession
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
                $TestCommandResult = Invoke-Command -Session $PSSession -ScriptBlock { Get-PSDrive -PSProvider ActiveDirectory -ErrorAction Stop | Select-Object -ExpandProperty Name } -ErrorAction Stop
                'AD' -in $TestCommandResult -and 'GC' -in $TestCommandResult
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
