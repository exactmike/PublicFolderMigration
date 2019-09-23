function Add-PFMParallelPSSession
{
    param (
        $Session
    )

}
if ($null -eq $script:ParallelPSSession)
{$script:ParallelPSSession = [System.Collections.ArrayList]::new()}
[void]$script:ParallelPSSession.Add($())