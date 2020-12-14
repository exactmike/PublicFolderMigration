function Add-PFMParallelPSSession
{
    param (
        [System.Management.Automation.Runspaces.PSSession]$PSSession
    )

    if ($null -eq $script:ParallelPSSession)
    {
        $script:ParallelPSSession = New-Object -TypeName System.Collections.ArrayList
    }

    $existingSessionIndex = (GetArrayIndexForProperty -array $script:ParallelPSSession -property Name -Value $PSSession.Name)
    if ($null -ne $existingSessionIndex -and $existingSessionIndex -ne -1)
    {
        Remove-PSSession -Session $script:ParallelPSSession[$existingSessionIndex]
        $script:ParallelPSSession.Remove($script:ParallelPSSession[$existingSessionIndex])
    }

    [void]$script:ParallelPSSession.Add($PSSession)
}
