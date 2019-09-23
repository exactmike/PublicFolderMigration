if (Test-PFMExchangePSSession -PSSession $Script:PSSession)
{
    #here is what you do if everythign is OK
}
else
{
    WriteLog -Message 'Removing Existing Failed PSSession' -EntryType Notification
    Remove-PSSession -Session $script:PsSession -ErrorAction SilentlyContinue
    WriteLog -Message 'Establish New PSSession to Exchange Organization' -EntryType Attempting
    $GetPFMExchangePSSessionParams = GetGetPFMExchangePSSessionParams
    try
    {
        Start-Sleep -Seconds 10
        $script:PsSession = Get-PFMExchangePSSession @GetPFMExchangePSSessionParams
        WriteLog -Message 'Establish New PSSession to Exchange Organization' -EntryType Succeeded
    }
    catch
    {
        $myerror = $_
        WriteLog -Message 'Establish New PSSession to Exchange Organization' -EntryType Failed
        WriteLog -Message $myerror.tostring() -ErrorLog -Verbose
        WriteLog -Message $message -EntryType Failed -ErrorLog -Verbose
        $exitmessage = "Testing Showed that Exchange Session Failed/Disconnected during permission processing for ID $ID."
        WriteLog -Message $exitmessage -EntryType Notification -ErrorLog -Verbose
        #Here is where you'd put code to enable a Resume operation
        #Break
    }
}