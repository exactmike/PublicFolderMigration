Function Remove-PFMValidatedPublicFolder
{
    [cmdletbinding()]
    [OutputType([pscustomobject])]
    param(
        [parameter(Mandatory, ValueFromPipeline)]
        [PSTypeName("PublicFolderValidation")]$PublicFolderValidation
        ,
        [parameter(Mandatory)]
        [ValidateScript( { TestIsWriteableDirectory -path $_ })]
        [string]$OutputFolderPath
        ,
        [parameter()]
        [ValidateSet('Unicode', 'BigEndianUnicode', 'Ascii', 'Default', 'UTF8', 'UTF8NOBOM', 'UTF7', 'UTF32')]
        [string]$Encoding = 'UTF8'
        ,
        [parameter()]
        [switch]$Passthru
    )
    begin
    {
        Connect-PFMExchange -IsParallel -ExchangeOnPremisesServer $script:ExchangeOnPremisesServer
        $ParallelSession = Get-PFMParallelPSSession -name $Script:ExchangeOnPremisesServer
        $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
        $ResultPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'RemoveValidatedPublicFolder.log')
        $ErrorResultPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'RemoveValidatedPublicFolder-ERRORS.log')
    }
    process
    {
        foreach ($pfv in $PublicFolderValidation)
        {
            $haderror = $false
            if ($true -eq $pfv.Validated)
            {
                $pfv.ActionName = 'Remove'
                $RemovePFParams = @{
                    ErrorAction = 'Stop'
                    Confirm     = $false
                    Identity    = $pfv.FoundEntryID
                }
                Confirm-PFMExchangeConnection -IsParallel -PSSession $ParallelSession
                $ParallelSession = Get-PFMParallelPSSession -name $script:ExchangeOnPremisesServer
                try
                {
                    Invoke-Command -Session $ParallelSession -ScriptBlock {
                        Remove-PublicFolder @using:RemovePFParams
                    } -ErrorAction 'Stop'
                    $pfv.ActionResult = $true
                }
                catch
                {
                    $myerrorstring = $_.tostring()
                    $haderror = $true
                    $pfv.ActionResult = $false
                    $pfv.ActionError = $myerrorstring
                }
                $pfv.ActionTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
            }
            else
            {
                $pfv.ActionName = 'None'
                $pfv.ActionTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
            }
            $pfvjson = ConvertTo-Json -InputObject $pfv -Compress
            Out-File -InputObject $pfvjson -Append -FilePath $ResultPath -Encoding $Encoding
            if ($true -eq $haderror)
            {
                Out-File -InputObject $pfvjson -Append -FilePath $ErrorResultPath -Encoding $Encoding
            }
            if ($true -eq $Passthru)
            {
                $pfv
            }
        }
    }
    end
    {

    }
}