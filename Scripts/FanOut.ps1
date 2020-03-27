param(
    $PublicFolderTreeFile =
    #,
    $ServersFile
    ,
    $DomainControllersFile
    ,
    $Credential
    ,
    $SidHistoryMapFile
    ,
    $MailEnabledFoldersFile
    ,
    $outputfolderpath
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


$servers = Get-Content -Path $ServersFile
$domainControllers = Get-Content -Path $DomainControllersFile
$publicFolderTree = Import-CSV -Path $PublicFolderTreeFile -Encoding UTF8
$sidHistoryMap = Get-Content -Raw -Path $SidHistoryMapFile -Encoding UTF8 | ConvertFrom-Json -AsHashtable
$mailPublicFolders = Import-Csv -Path $MailEnabledFoldersFile -Encoding UTF8

$ranges = New-SplitArrayRange -inputArray $PublicFolderTree -parts $servers.count
$s = 0
#<#
foreach ($r in $ranges)
{
    $server = $servers[$s]
    $dc = $domainControllers[$s]
    $startJobParams = @{
        name        = $server
        scriptblock = [scriptblock] {
            $file = Join-Path -path $using:outputfolderpath -ChildPath $("" + $($using:r).part + ".csv")
            Write-Verbose -Message "I'm running on server $using:server and dc $using:dc processing file $File" -Verbose
            Import-Module PublicFolderMigration
            #$publicFolderTree = Import-CSV -Path $using:PublicFolderTreeFile -Encoding UTF8 -Verbose
            $sidHistoryMap = Get-Content -Raw -Path $using:SidHistoryMapFile -Encoding UTF8 | ConvertFrom-Json -AsHashtable -Verbose
            Write-Verbose -Message "imported SidHistory" -Verbose
            $mailPublicFolders = Import-Csv -Path $using:MailEnabledFoldersFile -Encoding UTF8 -Verbose
            Write-Verbose -message "imported Mail Folders" -Verbose
            $Treepart = Import-Csv -path $file -encoding utf8 -verbose
            Write-Verbose -Message "imported Public Folder Information objects" -Verbose

            Connect-PFMExchange -ExchangeOnPremisesServer $using:server -Credential $using:Credential -verbose
            Connect-PFMActiveDirectory -DomainController $using:dc -Credential $using:Credential -verbose
            $GPGPParams = @{
                PublicFolderInfoObject  = $Treepart
                OutputFolderPath        = $using:outputfolderpath
                IncludeClientPermission = $true
                IncludeSIDHistory       = $true
                IncludeSendAs           = $false
                IncludeSendOnBehalf     = $false
                ExpandGroups            = $false
                SidHistoryRecipientMap  = $SidHistoryMap
                MailPublicFolder        = $MailPublicFolders
                #Verbose = $true
                InformationAction       = 'Continue'
            }
            Get-PFMPublicFolderPermission @GPGPParams
        }
    }
    #$StartJobParams
    Start-Job @StartJobParams
    $s++
}
#>