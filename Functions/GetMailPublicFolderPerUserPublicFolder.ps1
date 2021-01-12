Function GetMailPublicFolderPerUserPublicFolder
{

    [CmdletBinding()]
    param
    (
        [System.Management.Automation.Runspaces.PSSession]$ExchangeSession
        ,
        [psobject[]]$PublicFolder
    )
    begin
    {
        $message = "Get-MailPublicFolder for each Public Folder"
        WriteLog -Message $message -EntryType Attempting
        $PublicFolderCount = $PublicFolder.Count
    } # end begin
    process
    {
        foreach ($pf in $PublicFolder)
        {
            $CurrentPF++
            $GetMailPublicFolderParams = @{
                Identity      = $pf.Identity.tostring()
                ErrorAction   = 'SilentlyContinue'
                WarningAction = 'SilentlyContinue'
            }
            $InnerMessage = "Get-MailPublicFolder -Identity $($pf.Identity.tostring())"
            Write-Progress -Activity $message -Status $InnerMessage -CurrentOperation "$CurrentPF of $PublicFolderCount" -PercentComplete $($CurrentPF / $PublicFolderCount * 100)
            try
            {
                #output Selected object with additional properties from the Pf object
                $MEPF = Invoke-Command -Session $script:PSSession -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -ScriptBlock {
                    Get-MailPublicFolder @using:GetMailPublicFolderParams
                }
                if ($null -ne $MEPF)
                {
                    $CustomProperties = @(
                        '*'
                        @{n = 'PFIdentity'; e = { $pf.Identity.tostring() } }
                    )
                    if ($null -eq $MEPF.EntryID) #Exchange 2013 and later include the EntryID natively so check for that
                    {
                        $CustomProperties += @{n = 'EntryID'; e = { $pf.EntryID.tostring() } }
                    }
                    $MEPF | Select-Object -Property $CustomProperties
                }
            }
            catch
            {
                $myerror = $_
                WriteLog -message $InnerMessage -EntryType Failed
                WriteLog -message $myerror.tostring() -ErrorLog
            }
        }
    } # end Process
    end
    {
        WriteLog -Message $message -EntryType Succeeded
    }

}
