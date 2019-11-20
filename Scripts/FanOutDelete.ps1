Function Start-ComplexJob
{
    <#
        .SYNOPSIS
        Helps Start Complex Background Jobs with many arguments and functions using Start-Job.
        .DESCRIPTION
        Helps Start Complex Background Jobs with many arguments and functions using Start-Job.
        The primary utility is to bring custom functions from the current session into the background job.
        A secondary utility is to formalize the input for creation complex background jobs by using a hashtable template and splatting.
        .PARAMETER  Name
        The name of the background job which will be created.  A string.
        .PARAMETER  JobFunctions
        The name[s] of any local functions which you wish to export to the background job for use in the background job script.
        The definition of any function listed here is exported as part of the script block to the background job.
        .EXAMPLE
        $StartComplexJobParams = @{
        jobfunctions = @(
                'Connect-WAAD'
            ,'Get-TimeStamp'
            ,'Write-Log'
            ,'Write-EndFunctionStatus'
            ,'Write-StartFunctionStatus'
            ,'Export-Data'
            ,'Get-MatchingAzureADUsersAndExport'
        )
        name = "MatchingAzureADUsersAndExport"
        arguments = @($SourceData,$SourceDataFolder,$LogPath,$ErrorLogPath,$OnlineCred)
        script = [scriptblock]{
            $PSModuleAutoloadingPreference = "None"
            $sourcedata = $args[0]
            $sourcedatafolder = $args[1]
            $logpath = $args[2]
            $errorlogpath = $args[3]
            $credential = $args[4]
            Connect-WAAD -MSOnlineCred $credential
            Get-MatchingAzureADUsersAndExport
        }
        }
        Start-ComplexJob @StartComplexJobParams
    #>
    [cmdletbinding()]
    param
    (
        [string]$Name
        ,
        [string[]]$JobFunctions
        ,
        [psobject[]]$Arguments
        ,
        [string]$Script
    )
    #build functions to initialize in job
    $JobFunctionsText = ''
    foreach ($Function in $JobFunctions)
    {
        $FunctionText = 'function ' + (Get-Command -Name $Function).Name + "{`r`n" + (Get-Command -Name $Function).Definition + "`r`n}`r`n"
        $JobFunctionsText = $JobFunctionsText + $FunctionText
    }
    $ExecutionScript = $JobFunctionsText + $Script
    #$initializationscript = [scriptblock]::Create($script)
    $ScriptBlock = [scriptblock]::Create($ExecutionScript)
    $StartJobParams = @{
        Name         = $Name
        ArgumentList = $Arguments
        ScriptBlock  = $ScriptBlock
    }
    #$startjobparams.initializationscript = $initializationscript
    Start-Job @StartJobParams

}
Function New-SplitArrayRange
{

    <#
        .SYNOPSIS
        Provides Start and End Ranges to Split an array into a specified number of parts (new arrays) or parts (new arrays) with a specified number (size) of elements
        .PARAMETER inArray
        A one dimensional array you want to split
        .EXAMPLE
        Split-array -inArray @(1,2,3,4,5,6,7,8,9,10) -parts 3
        .EXAMPLE
        Split-array -inArray @(1,2,3,4,5,6,7,8,9,10) -size 3
        .NOTE
        Derived from https://gallery.technet.microsoft.com/scriptcenter/Split-an-array-into-parts-4357dcc1#content
        #>
    [cmdletbinding()]
    param(
        [parameter(Mandatory)]
        [array]$inputArray
        ,
        [parameter(Mandatory, ParameterSetName = 'Parts')]
        [int]$parts
        ,
        [parameter(Mandatory, ParameterSetName = 'Size')]
        [int]$size
    )
    switch ($PSCmdlet.ParameterSetName)
    {
        'Parts'
        {
            $PartSize = [Math]::Ceiling($inputArray.count / $parts)
        }#Parts
        'Size'
        {
            $PartSize = $size
            $parts = [Math]::Ceiling($inputArray.count / $size)
        }#Size
    }#switch
    for ($i = 1; $i -le $parts; $i++)
    {
        $start = (($i - 1) * $PartSize)
        $end = (($i) * $PartSize) - 1
        if ($end -ge $inputArray.count) { $end = $inputArray.count }
        $SplitArrayRange = [pscustomobject]@{
            Part  = $i
            Start = $start
            End   = $end
        }
        $SplitArrayRange
    }#for
}
Function Read-OpenFileDialog
{

    [cmdletbinding()]
    param(
        [string]$WindowTitle
        ,
        [string]$InitialDirectory
        ,
        [string]$Filter = 'All files (*.*)|*.*'
        ,
        [switch]$AllowMultiSelect
    )
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $WindowTitle
    if ($PSBoundParameters.ContainsKey('InitialDirectory')) { $openFileDialog.InitialDirectory = $InitialDirectory }
    $openFileDialog.Filter = $Filter
    if ($AllowMultiSelect) { $openFileDialog.MultiSelect = $true }
    $openFileDialog.ShowHelp = $true
    # Without this line the ShowDialog() function may hang depending on system configuration and running from console vs. ISE.
    $result = $openFileDialog.ShowDialog()
    switch ($Result)
    {
        'OK'
        {
            if ($AllowMultiSelect)
            {
                $openFileDialog.Filenames
            }
            else
            {
                $openFileDialog.Filename
            }
        }
        'Cancel'
        {
        }
    }
    $openFileDialog.Dispose()
    Remove-Variable -Name openFileDialog

}


$credential = Get-Credential
$ListDataFile = Read-OpenFileDialog -WindowTitle 'Select Data File to Process'
$ServersFile = Read-OpenFileDialog -WindowTitle 'Select Servers File to Process'
$ListData = Import-Csv $ListDataFile
$Servers = (Import-Csv $ServersFile).rpclientaccessserver

$SplitArrayRange = New-SplitArrayRange -inputArray $ListData -parts $Servers.count
$part = 0
foreach ($s in $Servers)
{
    $part++
    $thisSplitArrayRange = $SplitArrayRange.where( { $_.Part -eq $part })
    $SplitData = $ListData[$thisSplitArrayRange.Start..$thisSplitArrayRange.end]
    $StartComplexJobParams = @{
        name      = $s
        arguments = @($SplitData, $s, $credential)
        script    = [scriptblock] {
            Import-Module PublicFolderMigration
            $List = $args[0]
            $Server = $args[1]
            $credential = $args[2]
            Connect-PFMExchange -ExchangeOnPremisesServer $Server -Credential $credential
            $List.EntryID |
            Invoke-PFMValidatePublicFolder -Validations NoSubFolders, NotMailEnabled, NoItems -OutputFolderPath D:\PerficientReports\PFDeletions\ |
            Remove-PFMValidatedPublicFolder -OutputFolderPath D:\PerficientReports\PFDeletions\
        }
    }
    Start-ComplexJob @StartComplexJobParams
}
