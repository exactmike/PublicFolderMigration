function Get-PublicFolderDatabaseReport
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Path to a file location for the report")]
        [Alias("PSPath")]
        [ValidateNotNullOrEmpty()]
        [string]
        $FilePath
    )

    begin
    {
    }

    process
    {
        Get-PublicFolderDatabase -status | ConvertTo-Json | Out-File -FilePath $FilePath
    }

    end
    {
    }
}