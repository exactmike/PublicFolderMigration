function Get-HTMLReport
    {
        [cmdletbinding()]
        param
        (
            $ReportObject
            ,
            $ResultMatrix
            ,
            $PublicFolderMailboxServer
        )
        $html =
@"
<html>
<style>
    body
    {
        font-family:Arial,sans-serif;
        font-size:8pt;
    }
    table
    {
        border-collapse:collapse;
        font-size:8pt;
        font-family:Arial,sans-serif;
        border-collapse:collapse;
        min-width:400px;
    }
    table,th, td
    {
        border: 1px solid black;
    }
    th
    {
        text-align:center;
        font-size:18;
        font-weight:bold;
    }
</style>
<body>
<font size="1" face="Arial,sans-serif">
<h1 align="center">Exchange Public Folder Replication Report</h1>
<h4 align="center">Generated $([DateTime]::Now)</h3>

</font><h2>Overall Summary</h2>
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="2">Public Folder Environment Summary</th></tr>
<tr><td>Included Public Folder Servers</td><td>$($ReportObject.IncludedPublicFolderServersAndDatabases)</td></tr>
<tr><td>Count of Included Public Folders</td><td>$($ReportObject.IncludedPublicFoldersCount)</td></tr>
<tr><td>Count of Included Container or Empty Public Folders (0 Item Count)</td><td>$($ReportObject.IncludedContainerOrEmptyPublicFoldersCount)</td></tr>
<tr><td>Count of Included Public Folders with Incomplete Replication on Included Servers</td><td>$($ReportObject.IncludedReplicationIncompletePublicFolders)</td></tr>
<tr><td>Count of Total Items in Included Public Folders</td><td>$($ReportObject.TotalItemCountFromIncludedPublicFolders)</td></tr>
<tr><td>Total Size of Included Public Folder Items (Bytes)</td><td>$($ReportObject.TotalSizeOfIncludedPublicFoldersInBytes)</td></tr>
<tr><td>Average Size of Included Public Folders (Non Empty/Container)</td><td>$($ReportObject.AverageSizeOfIncludedPublicFolders)</td></tr>
<tr><td>Average Item Count of Included Public Folders (Non Empty/Container)</td><td>$($ReportObject.AverageItemCountFromIncludedPublicFolders)</td></tr>
</table>
<br />
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="3">Largest Public Folders by Size</th></tr>
<tr style="background-color:#E9E9E9;font-weight:bold"><td>Folder Path</td><td>Size</td><td>Item Count</td></tr>
$(
    if (-not $ReportObject.LargestPublicFolders.Count -gt 0)
    {
        "<tr><td colspan='3'>No Largest Public Folders Report is Included in this report.</td></tr>"
    }
    else
    {
        foreach($sizeResult in $reportObject.LargestPublicFolders)
        {
            "<tr><td>$($sizeResult.FolderPath)</td><td>$($sizeResult.TotalItemSize)</td><td>$($sizeResult.ItemCount)</td></tr>`r`n"
        }
    }
)
</table>
</font><h2>Public Folder Replication Results</h2>
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="5">Folders with Incomplete Replication on Included Servers</th></tr>
<tr style="background-color:#E9E9E9;font-weight:bold"><td>Folder Path</td><td>Item Count</td><td>Size</td><td>Servers with Replicas Configured</td><td>Servers with Replication Incomplete</td></tr>
$(
    if (-not $ReportObject.PublicFoldersWithIncompleteReplication.Count -gt 0)
    {
        "<tr><td colspan='4'>There are no public folders with incomplete replication.</td></tr>"
    }
    else
    {
        foreach($IncompleteFolder in $ReportObject.PublicFoldersWithIncompleteReplication)
        {
            "<tr><td>$($IncompleteFolder.FolderPath)</td><td>$($IncompleteFolder.ItemCount)</td><td>$($IncompleteFolder.TotalItemSize)</td><td>$($IncompleteFolder.ConfiguredReplicaServers)</td><td>$($IncompleteFolder.IncompleteServers)</td></tr>`r`n"
        }
    }
)
</table>
<br />
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="$($PublicFolderMailboxServer.Count + 1)">Public Folder Replication Information</th></tr>
<tr style="background-color:#E9E9E9;font-weight:bold"><td>Folder Path</td>
$(
    foreach($rServer in $PublicFolderMailboxServer)
    {
        "<td>$($rServer)</td>"
    }
)
</tr>
$(
    if (-not $ResultMatrix.Count -gt 0)
    {
        "<tr><td colspan='$($PublicFolderMailboxServer.Count + 1)'>There are no public folders in this report.</td></tr>"
    }
    foreach($rItem in $ResultMatrix)
    {
        "<tr><td>$($rItem.FolderPath)</td>"
        foreach($rServer in $PublicFolderMailboxServer)
        {
            $(
                $rDataItem = $rItem.Data | Where-Object { $_.ServerName -eq $rServer }
                if ($rDataItem -eq $null)
                {
                    '<td>N/A</td>'
                }
                else
                {
                    if ($rDataItem.Progress -ne 100)
                    {
                        $color = '#FC2222'
                    }
                    else
                    {
                        $color = '#A9FFB5'
                    }
                    "<td style='background-color:$($color)'><div title='$($rDataItem.TotalItemSize) of $($rItem.TotalItemSize) and $($rDataItem.ItemCount) of $($rItem.ItemCount) items.'>$($rDataItem.Progress)%</div></td>"
                }
            )
        }
        '</tr>'
    }
)
</table>
</body>
</html>
"@
        Write-Output -InputObject $html
    }
#end function Get-HTMLReport