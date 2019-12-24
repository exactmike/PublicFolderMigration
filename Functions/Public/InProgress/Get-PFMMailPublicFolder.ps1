Function Get-PFMMailPublicFolder
{
    [cmdletbinding(DefaultParameterSetName = 'RecipientOnly')]
    [OutputType([System.Object[]])]
    param(
        [parameter(Mandatory, ParameterSetName = 'TreeMappedRecipient')]
        #PublicFolderInfoObjects to scope the retrieval of Mail Public Folders to the submitted objects.  Also includes EntryID and Identity in the resulting output objects.
        [psobject[]]$PublicFolderInfoObject
        ,
        [Parameter(Mandatory)]
        [ValidateScript( { TestIsWriteableDirectory -Path $_ })]
        $OutputFolderPath
        ,
        [parameter(Mandatory)]
        [ExportDataOutputFormat[]]$Outputformat
        ,
        [parameter()]
        [ValidateSet('Unicode', 'BigEndianUnicode', 'Ascii', 'Default', 'UTF8', 'UTF8NOBOM', 'UTF7', 'UTF32')]
        [string]$Encoding = 'UTF8'
        ,
        [parameter()]
        [switch]$Passthru
    )
    Confirm-PFMExchangeConnection -PSSession $Script:PSSession
    $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
    $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetMailPublicFolder.log')
    $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetMailPublicFolder-ERRORS.log')
    #$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
    WriteLog -Message "Exchange Session is Running in Exchange Organzation $script:ExchangeOrganization" -EntryType Notification
    $Results = @(
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
                    Write-Progress -Activity 'Get Mail Public Folder For Each Public Folder' -Status $Status -CurrentOperation "$CurrentPF of $($PublicFolderInfoObject.Count)" -PercentComplete $($CurrentPF / $PublicFolderInfoObject.Count * 100)
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
    )
    $ResultCount = $Results.Count
    WriteLog -Message "Count of Mail Enabled PublicFolders Retrieved: $ResultCount" -EntryType Notification -verbose
    $CreatedFilePath = @(
        foreach ($of in $Outputformat)
        {
            Export-PFMData -ExportFolderPath $OutputFolderPath -DataToExportTitle 'MailEnabledPublicFolders' -ReturnExportFilePath -Encoding $Encoding -DataFormat $of -DataToExport $Results
        }
    )
    WriteLog -Message "Output files created: $($CreatedFilePath -join '; ')" -entryType Notification -verbose
    if ($true -eq $Passthru)
    {
        $Results
    }
}
