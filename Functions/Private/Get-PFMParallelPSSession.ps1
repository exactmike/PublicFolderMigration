Function Get-PFMParallelPSsession
{
    [cmdletbinding()]
    [outputtype([System.Management.Automation.Runspaces.PSSession])]
    Param(
        [parameter(Mandatory)]
        [string]$Name
    )
    $SessionIndex = GetArrayIndexForProperty -array $Script:ParallelPSSession -Property 'Name' -value $Name
    $PSSession = $Script:ParallelPSSession[$SessionIndex]
    if ($null -ne $PSSession)
    {
        $PSSession
    }
}
