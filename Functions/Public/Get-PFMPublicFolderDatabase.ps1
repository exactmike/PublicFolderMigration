function Get-PFMPublicFolderDatabase
{
    <#
    .SYNOPSIS
    Gets Public Folder Database Information Objects for all or specified public folder databases and servers
    .DESCRIPTION
    Gets Public Folder Database Information Objects for all or specified public folder databases and servers
    .PARAMETER Identity
    Specify the public folder database identity for the public folder database for which to get a Public Folder Database Information Object. If Identity and Server are not used the default is to return an Information Object for each Public FolderDatabase.
    .PARAMETER Server
    Specify the public folder server identity for the public folder server for which to get a Public Folder Database Information Object. If Identity and Server are not used the default is to return an Information Object for each Public FolderDatabase.
    .PARAMETER Passthru
    Controls whether the Public Folder Database Information Object(s) is/are returned to the PowerShell pipeline for further processing.
    .PARAMETER OutputFolderPath
    Mandatory parameter for the already existing directory location where you want public folder replication and stats reports to be placed.  Operational log files will also go to this location.
    .PARAMETER OutputFormat
    Mandatory parameter used to specify whether you want csv, json, xml, clixml or any combination of these.
    .EXAMPLE
    Connect-PFMExchange -ExchangeOnPremisesServer USCLTEX10PF01.us.clt.contoso.com -credential $cred
    Get-PFMPublicFolderDatabase -OutputFolderPath c:\PFReports -OutputFormats csv,json,xml -Server USCLTEX10PF01

    Gets a public folder database information object from USCLTEX10PF01.us.clt.contoso.com and exports it to csv, json, and xml formats in c:\PFReports
    #>
    [CmdletBinding(ConfirmImpact = 'None', DefaultParameterSetName = 'All')]
    [OutputType([System.Object[]])]
    param (
        [parameter(Mandatory, ParameterSetName = 'Identity')]
        [string[]]$Identity
        ,
        [parameter(Mandatory, ParameterSetName = 'Server')]
        [string[]]$Server
        ,
        [parameter()]
        [switch]$Passthru
        ,
        [parameter(Mandatory)]
        [ValidateScript( { TestIsWriteableDirectory -path $_ })]
        [string]$OutputFolderPath
        ,
        [parameter(Mandatory)]
        [ExportDataOutputFormat[]]$Outputformat
        ,
        [parameter()]
        [ValidateSet('Unicode', 'BigEndianUnicode', 'Ascii', 'Default', 'UTF8', 'UTF8NOBOM', 'UTF7', 'UTF32')]
        [string]$Encoding = 'UTF8'
    )

    begin
    {
        Confirm-PFMExchangeConnection -PSSession $Script:PSSession
        $BeginTimeStamp = Get-Date -Format yyyyMMdd-HHmmss
        $script:LogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetPublicFolderDatabase.log')
        $script:ErrorLogPath = Join-Path -path $OutputFolderPath -ChildPath $($BeginTimeStamp + 'GetPublicFolderDatabase-ERRORS.log')
        WriteLog -Message "Calling Invocation = $($MyInvocation.Line)" -EntryType Notification
        $ExchangeOrganization = Invoke-Command -Session $Script:PSSession -ScriptBlock { Get-OrganizationConfig | Select-Object -ExpandProperty Identity | Select-Object -ExpandProperty Name }
        WriteLog -Message "Exchange Session is Running in Exchange Organzation $ExchangeOrganization" -EntryType Notification
    }

    Process
    {
        $pfDatabaseInfoObjects = @(
            switch ($PSCmdlet.ParameterSetName)
            {
                'All'
                {
                    Invoke-Command -Session $script:PSSession -scriptblock { Get-PublicFolderDatabase -status }
                }
                'Identity'
                {
                    foreach ($i in $Identity)
                    {
                        Invoke-Command -Session $script:PSSession -scriptblock { Get-PublicFolderDatabase -status -Identity $using:i }
                    }
                }
                'Server'
                {
                    foreach ($s in $Server)
                    {
                        Invoke-Command -Session $script:PSSession -scriptblock { Get-PublicFolderDatabase -status -Server $using:s }
                    }
                }
            }
        )
    }

    end
    {
        $pfDatabaseInfoObjectsCount = $pfDatabaseInfoObjects.Count
        WriteLog -Message "Count of Public Folder Databases Retrieved: $pfDatabaseInfoObjectsCount " -EntryType Notification -verbose
        $CreatedFilePath = @(
            foreach ($of in $Outputformat)
            {
                Export-Data -ExportFolderPath $OutputFolderPath -DataToExportTitle 'PublicFolderDatabases' -ReturnExportFilePath -Encoding $Encoding -DataFormat $of -DataToExport $pfDatabaseInfoObjects
            }
        )
        WriteLog -Message "Output files created: $($CreatedFilePath -join '; ')" -entryType Notification -verbose
        if ($true -eq $Passthru)
        {
            $pfDatabaseInfoObjects
        }
    }
}