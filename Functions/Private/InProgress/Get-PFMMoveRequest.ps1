Function Get-PFMMoveRequest
{

    [cmdletbinding()]
    param(
        $ExchangeOrganization
        ,
        $BatchName
    )
    if ((Connect-Exchange -ExchangeOrganization $ExchangeOrganization) -ne $true)
    {
        Write-Log -Message "Connect to Exchange Organization $ExchangeOrganization" -ErrorLog -EntryType Failed
        throw { "Connect to Exchange Organization $ExchangeOrganization Failed" }
    }#End If
    $Message = "Get all existing $BatchName move requests"
    Write-Log -message $Message -Verbose -EntryType Attempting
    $splat = @{
        cmdlet               = 'Get-PublicFolderMailboxMigrationRequest'
        ExchangeOrganization = $ExchangeOrganization
        ErrorAction          = 'Stop'
        splat                = @{
            BatchName   = 'MigrationService:' + $BatchName
            ResultSize  = 'Unlimited'
            ErrorAction = 'Stop'
        }#innersplat
    }#outersplat
    $Script:mr = @(Invoke-ExchangeCommand @splat)
    $Script:fmr = @($mr | Where-Object -FilterScript { $_.status -eq 'Failed' })
    $Script:ipmr = @($mr | Where-Object { $_.status -eq 'InProgress' })
    $Script:smr = @($mr | Where-Object { $_.status -eq 'Suspended' })
    $Script:asmr = @($mr | Where-Object { $_.status -in ('AutoSuspended', 'Synced') })
    $Script:cmr = @($mr | Where-Object { $_.status -like 'Completed*' })
    $Script:qmr = @($mr | Where-Object { $_.status -eq 'Queued' })
    $Script:ncmr = @($mr | Where-Object { $_.status -notlike 'Completed*' })
    Write-Log -message $Message -Verbose -EntryType Succeeded

}
