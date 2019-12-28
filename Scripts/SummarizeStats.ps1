param($StatsFile)
$Stats = $(Import-JSON -Path $StatsFile -Encoding utf8)
$SortedStats = $Stats | Sort-Object -Property EntryID

$i = 0
$ic = 0
do
{
    $s = $SortedStats[$ic]
    Write-Progress -Activity 'Summarize Stats Data' -Status "Processing $ic of $($SortedStats.count)" -CurrentOperation "Processing $($s.name)" -PercentComplete $($ic / $SortedStats.count * 100)
    $AssociatedItemCount = [System.Collections.ArrayList]::new()
    $ContactCount = [System.Collections.ArrayList]::new()
    $creationTime = [System.Collections.ArrayList]::new()
    $deletedItemCount = [System.Collections.ArrayList]::new()
    $itemCount = [System.Collections.ArrayList]::new()
    $LastAccessTime = [System.Collections.ArrayList]::new()
    $LastModificationTime = [System.Collections.ArrayList]::new()
    $LastUserModifiicationTime = [System.Collections.ArrayList]::new()
    $LastUserAccesstime = [System.Collections.ArrayList]::new()
    $OwnerCount = [System.Collections.ArrayList]::new()
    $TotalAssociatedItemSize = [System.Collections.ArrayList]::new()
    $TotalDeletedItemSize = [System.Collections.ArrayList]::new()
    $TotalItemSize = [System.Collections.ArrayList]::new()
    do
    {
        if ($sortedStats[$i].EntryID -eq $s.EntryID)
        {
            [void]$AssociatedItemCount.Add($SortedStats[$i].AssociatedItemCount)
            [void]$ContactCount.Add($SortedStats[$i].ContactCount)
            [void]$creationTime.Add($SortedStats[$i].CreationTime)
            [void]$deletedItemCount.Add($SortedStats[$i].DeletedItemCount)
            [void]$itemCount.Add($SortedStats[$i].ItemCount)
            [void]$LastAccessTime.Add($SortedStats[$i].LastAccessTime)
            [void]$LastModificationTime.Add($SortedStats[$i].LastModificationTime)
            [void]$LastUserModifiicationTime.Add($SortedStats[$i].LastUserModificationTime)
            [void]$LastUserAccesstime.Add($SortedStats[$i].LastUserAccessTime)
            [void]$OwnerCount.Add($SortedStats[$i].OwnerCount)

            if ($null -ne $SortedStats[$i].TotalAssociatedItemSize)
            {
                [void]$TotalAssociatedItemSize.Add([Int64]$SortedStats[$i].TotalAssociatedItemSize.split('(')[1].split(' ')[0].replace(',', ''))
            }
            else { [void]$TotalAssociatedItemSize.Add($null) }

            if ($null -ne $SortedStats[$i].TotalDeletedItemSize)
            {
                [void]$TotalDeletedItemSize.Add([Int64]$SortedStats[$i].TotalDeletedItemSize.split('(')[1].split(' ')[0].replace(',', ''))
            }
            else { [void]$TotalDeletedItemSize.Add($null) }

            if ($null -ne $SortedStats[$i].TotalItemSize)
            {
                [void]$TotalItemSize.Add([Int64]$SortedStats[$i].TotalItemSize.split('(')[1].split(' ')[0].replace(',', ''))
            }
            else { [void]$TotalDeletedItemSize.Add($null) }
        }
        $i++
    }
    until ($sortedStats[$i].EntryID -ne $s.EntryID)

    [PSCustomObject]@{
        EntryID                  = $s.EntryID
        Name                     = $s.Name
        FolderPath               = '\' + $s.FolderPath
        StatCount                = $creationTime.Count
        AssociatedItemCount      = $($max = $null; foreach ($e in $AssociatedItemCount) { if ($max -lt $e) { $max = $e } }; $max)
        ContactCount             = $($max = $null; foreach ($e in $ContactCount) { if ($max -lt $e) { $max = $e } }; $max)
        CreationTime             = $($max = $null; foreach ($e in $CreationTime) { if ($max -lt $e) { $max = $e } }; $max)
        DeletedItemCount         = $($max = $null; foreach ($e in $deletedItemCount) { if ($max -lt $e) { $max = $e } }; $max)
        ItemCount                = $($max = $null; foreach ($e in $itemCount) { if ($max -lt $e) { $max = $e } }; $max)
        LastAccessTime           = $($max = $null; foreach ($e in $LastAccessTime) { if ($max -lt $e) { $max = $e } }; $max)
        LastModificationTime     = $($max = $null; foreach ($e in $LastModificationTime) { if ($max -lt $e) { $max = $e } }; $max)
        LastUserAccessTime       = $($max = $null; foreach ($e in $LastUserAccesstime) { if ($max -lt $e) { $max = $e } }; $max)
        LastUserModificationTime = $($max = $null; foreach ($e in $LastUserModifiicationTime) { if ($max -lt $e) { $max = $e } }; $max)
        OwnerCount               = $($max = $null; foreach ($e in $OwnerCount) { if ($max -lt $e) { $max = $e } }; $max)
        TotalAssociatedItemSize  = $($max = $null; foreach ($e in $TotalAssociatedItemSize) { if ($max -lt $e) { $max = $e } }; $max)
        TotalDeletedItemSize     = $($max = $null; foreach ($e in $TotalDeletedItemSize) { if ($max -lt $e) { $max = $e } }; $max)
        TotalItemSize            = $($max = $null; foreach ($e in $TotalItemSize) { if ($max -lt $e) { $max = $e } }; $max)
    }
    $ic = $i;
}
until ($ic -ge $SortedStats.Count)