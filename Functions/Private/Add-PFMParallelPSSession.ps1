function Add-PFMParallelPSSession
{
    param (
        $Session
    )

    if ($null -eq $script:ParallelPSSession)
    {
        $script:ParallelPSSession = [System.Collections.ArrayList]::new()
    }

    $existingSessionIndex = (GetArrayIndexForProperty -array $script:ParallelPSSession -property Name -Value $Session.Name)
    if ($null -ne $existingSessionIndex -and $existingSessionIndex -ne -1)
    {
        Remove-PSSession -Session $script:ParallelPSSession[$existingSessionIndex]
        $script:ParallelPSSession.Remove($script:ParallelPSSession[$existingSessionIndex])
    }

    [void]$script:ParallelPSSession.Add($($Session))
}
