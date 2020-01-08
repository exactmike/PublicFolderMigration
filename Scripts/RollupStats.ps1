param($SummarizedStatsFile)

$SummarizedStats = import-json -path $SummarizedStatsFile -encoding utf8
$SortedSummarizedStats = $SummarizedStats | Sort-Object -Property FolderPath

$i = 1;
$ic = 1;
foreach ($f in $SortedSummarizedStats)
{
    Write-Progress -Activity 'Rollup PF Stats Data' -Status "Processing $ic of $($SortedSummarizedStats.count)" -CurrentOperation "Processing $($f.name)" -PercentComplete $($ic / $SortedSummarizedStats.count * 100)
    $creationTime = [System.Collections.ArrayList]::new()
    $deletedItemCount = [System.Collections.ArrayList]::new()
    $itemCount = [System.Collections.ArrayList]::new()
    $LastAccessTime = [System.Collections.ArrayList]::new()
    $LastModificationTime = [System.Collections.ArrayList]::new()
    $LastUserModifiicationTime = [System.Collections.ArrayList]::new()
    $LastUserAccesstime = [System.Collections.ArrayList]::new()
    $TotalAssociatedItemSize = [System.Collections.ArrayList]::new()
    $TotalDeletedItemSize = [System.Collections.ArrayList]::new()
    $TotalItemSize = [System.Collections.ArrayList]::new()
    do
    {
        if ($SortedSummarizedStats[$i].FolderPath.replace('[', '`[').replace(']', '`]') + '\' -like $f.FolderPath.replace('[', '`[').replace(']', '`]') + '\*')
        {
            [void]$creationTime.Add($SortedSummarizedStats[$i].CreationTime)
            [void]$deletedItemCount.Add($SortedSummarizedStats[$i].DeletedItemCount)
            [void]$itemCount.Add([int64]$SortedSummarizedStats[$i].ItemCount)
            [void]$LastAccessTime.Add($SortedSummarizedStats[$i].LastAccessTime)
            [void]$LastModificationTime.Add($SortedSummarizedStats[$i].LastModificationTime)
            [void]$LastUserModifiicationTime.Add($SortedSummarizedStats[$i].LastUserModificationTime)
            [void]$LastUserAccesstime.Add($SortedSummarizedStats[$i].LastUserAccessTime)
            [void]$TotalAssociatedItemSize.Add([Int64]$SortedSummarizedStats[$i].TotalAssociatedItemSize)
            [void]$TotalDeletedItemSize.Add([Int64]$SortedSummarizedStats[$i].TotalDeletedItemSize)
            [void]$TotalItemSize.Add([Int64]$SortedSummarizedStats[$i].TotalItemSize)
        }
        $i++;
    }
    until ($SortedSummarizedStats[$i].FolderPath.replace('[', '`[').replace(']', '`]') -notlike $f.FolderPath.replace('[', '`[').replace(']', '`]') + '\*');

    [PSCustomObject]@{
        EntryID                         = $f.EntryID
        Name                            = $f.Name
        FolderPath                      = $f.FolderPath
        TreeMaxCreationTime             = $($max = $null; foreach ($e in $CreationTime) { if ($max -lt $e) { $max = $e } }; $max)
        TreeDeletedItemCount            = $($sum = 0; foreach ($e in $deletedItemCount) { $sum += $e } ; $sum)
        TreeItemCount                   = $($sum = 0; foreach ($e in $itemCount) { $sum += $e } ; $sum)
        TreeMaxLastAccessTime           = $($max = $null; foreach ($e in $LastAccessTime) { if ($max -lt $e) { $max = $e } }; $max)
        TreeMaxLastModificationTime     = $($max = $null; foreach ($e in $LastModificationTime) { if ($max -lt $e) { $max = $e } }; $max)
        TreeMaxLastUserAccessTime       = $($max = $null; foreach ($e in $LastUserAccesstime) { if ($max -lt $e) { $max = $e } }; $max)
        TreeMaxLastUserModificationTime = $($max = $null; foreach ($e in $LastUserModifiicationTime) { if ($max -lt $e) { $max = $e } }; $max)
        TreeTotalAssociatedItemSize     = $($sum = 0; foreach ($e in $TotalAssociatedItemSize) { $sum += $e }; $sum)
        TreeTotalDeletedItemSize        = $($sum = 0; foreach ($e in $TotalDeletedItemSize) { $sum += $e }; $sum)
        TreeTotalItemSize               = $($sum = 0; foreach ($e in $TotalItemSize) { $sum += $e }; $sum)
    }
    $ic++;
    $i = $ic;

}