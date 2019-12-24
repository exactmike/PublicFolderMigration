Function Get-PFMMailPublicFolder
{
    [cmdletbinding(DefaultParameterSetName = 'RecipientOnly')]
    [OutputType([System.Object[]])]
    param(
        [parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'TreeMappedRecipient')]
        [psobject[]]$PublicFolderInfoObject
    )
    begin
    {
        Confirm-PFMExchangeConnection -PSSession $Script:PSSession
        $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
        $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetMailPublicFolder.log')
        $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetMailPublicFolder-ERRORS.log')
        #$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
        WriteLog -Message "Exchange Session is Running in Exchange Organzation $script:ExchangeOrganization" -EntryType Notification

    }
    process
    {
        Switch ($PSCmdlet.ParameterSetName)
        {
            'RecipientOnly'
            {
                Invoke-Command -ErrorAction Stop -WarningAction SilentlyContinue -Session $script:PSSession -ScriptBlock {
                    Get-MailPublicFolder -ResultSize Unlimited -ErrorAction Stop
                }
            }
            'TreeMappedRecipient'
            {
                foreach ($pf in $PublicFolderInfoObject)
                {
                    $CurrentPF++
                    $GetMailPublicFolderParams = @{
                        Identity      = $pf.Identity.tostring()
                        ErrorAction   = 'SilentlyContinue'
                        WarningAction = 'SilentlyContinue'
                    }
                    $Status = "Get-MailPublicFolder -Identity $($pf.Identity.tostring())"
                    Write-Progress -Activity $message -Status $Status -CurrentOperation "$CurrentPF of $($PublicFolderInfoObject.Count)" -PercentComplete $($CurrentPF / $PublicFolderInfoObject.Count * 100)
                    try
                    {
                        #output Selected object with additional properties from the Pf object
                        $MEPF = Invoke-Command -Session $script:PSSession -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -ScriptBlock {
                            Get-MailPublicFolder @using:GetMailPublicFolderParams
                        }
                        if ($null -ne $MEPF)
                        {
                            $MEPF | Select-Object -Property *, @{n = 'EntryID'; e = { $pf.EntryID.tostring() } }, @{n = 'PFIdentity'; e = { $pf.Identity.tostring() } }
                        }
                    }
                    catch
                    {
                        $myerror = $_
                        WriteLog -message $Status -EntryType Failed
                        WriteLog -message $myerror.tostring() -ErrorLog
                    }
                }
            }
        }
    }
}
