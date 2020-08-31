param(
$PublicFolderTreeFile
,
$Servers 
,
$Credential
, 
$User
)

Function New-SplitArrayRange
{

    <#
        .SYNOPSIS
        Provides Start and End Ranges to Split an array into a specified number of new arrays or new arrays with a specified number (size) of elements
        .PARAMETER inputArray
        A one dimensional array you want to split
        .EXAMPLE
        Split-array -inputArray @(1,2,3,4,5,6,7,8,9,10) -parts 3
        .EXAMPLE
        Split-array -inputArray @(1,2,3,4,5,6,7,8,9,10) -size 3
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


$publicFolderTree = Import-CSV -Path $PublicFolderTreeFile -Encoding UTF8 | Select-Object EntryID -Skip 110
$ranges = New-SplitArrayRange -inputArray $PublicFolderTree -parts $servers.count 
#<#
foreach ($r in $ranges)
{
    $serverIndex = ($r.part - 1)
    $server = $servers[$serverIndex]
    $treepart = $PublicFolderTree[$r.start..$r.end]
    
    $startJobParams = @{
        name      = $server
        scriptblock    = [scriptblock] {
            Write-Verbose -Message "I'm running on server $using:server processing part $($($using:r).part)" -Verbose
            Import-Module PublicFolderMigration
            $MyTreepart = $using:treepart
            Write-Verbose -Message "imported Public Folder Information objects" -Verbose

            Connect-PFMExchange -ExchangeOnPremisesServer $using:server -Credential $using:Credential -verbose
            $sess = Get-PSSession   
            Import-PSSession -Session $sess
            $i = 0
            foreach ($e in $MyTreepart)
            {
                $i++
                if (($i%10) -eq 0)
                {
                    Write-Verbose -Message "Processing Permission $i of $($MyTreepart.count)" -Verbose
                }
                $null = Add-PublicFolderClientPermission -AccessRights 'Owner' -User $using:user -confirm:$false -Identity $e.entryID
            }
        }
    }
    #$StartJobParams
    Start-Job @StartJobParams 
}
#>