param(
    $credential
    ,
    $ListDataFile
    ,
    $ServersFile
    ,
    $OutputFolderPath
)
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

$ListData = Import-Csv $ListDataFile
$Servers = Import-Csv $ServersFile

#$SplitArrayRange = New-SplitArrayRange -inputArray $ListData -parts $Servers.count
$part = 0
foreach ($s in $Servers)
{
    $part++
    #$thisSplitArrayRange = $SplitArrayRange.where( { $_.Part -eq $part })
    $SplitData = $ListData.where($_.rpcClientAccessServer -eq $s)
    $StartJobParams = @{
        name         = $s
        argumentlist = @($SplitData, $s, $credential, $part)
        scriptblock  = [scriptblock] {
            Import-Module PublicFolderMigration
            $List = $args[0]
            $Server = $args[1]
            $credential = $args[2]
            $part = $args[3]
            $RecordCount = $List.count
            $Record = 0
            $Interval = 10
            Connect-PFMExchange -ExchangeOnPremisesServer $Server -Credential $credential
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $List.EntryID |
            Invoke-PFMValidatePublicFolder -Validations NoSubFolders, NotMailEnabled, NoItems -OutputFolderPath $using:outputfolderpath |
            Remove-PFMValidatedPublicFolder -OutputFolderPath $using:outputfolderpath -passthru |
            & { process
                {
                    $Record++
                    if ($Record % $Interval -eq 0)
                    {
                        $SecondsPerRecord = [math]::Round($($(Stopwatch.Elapsed.TotalSeconds) / $Record), 1)
                        $RemainingRecords = $RecordCount - $Record
                        $SecondsRemaining = $SecondsPerRecord * $RemainingRecords
                        $ProgressParams = @{
                            Activity         = 'PF Removals'
                            Status           = "Minutes Elapsed:$($Stopwatch.Elapsed.TotalMinutes)"
                            PercentComplete  = $($Record / $RecordCount * 100)
                            SecondsRemaining = $SecondsRemaining
                            ParentID         = 0
                            ID               = $part
                        }
                        Write-Progress @ProgressParams
                    }
                } }
        }
    }
    Start-Job @StartJobParams
}
