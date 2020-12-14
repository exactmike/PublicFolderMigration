param($TreeFile)

$tree = import-json $TreeFile
$SortedTree = $Tree | Sort-Object -Property Identity
#$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$c = 1
foreach ($f in $sortedTree)
{
    Write-Progress -Activity 'Add Members' -Status "Processing $c of $($SortedTree.count)" -CurrentOperation "Processing $($f.name)" -PercentComplete $($c / $SortedTree.count * 100)
    Add-Member -InputObject $f -MemberType NoteProperty -Name FolderCount -Value 1 -Force;
    Add-Member -InputObject $f -MemberType NoteProperty -Name SubFolderEntryIDs -Value @() -Force;
    Add-Member -InputObject $f -MemberType NoteProperty -Name MailEnabledFolderCount -Value 0 -Force;
    Add-Member -InputObject $f -MemberType NoteProperty -Name ReplicaCount -Value $f.Replicas.count -Force;
    $c++
}
$i = 1;
$ic = 1;
foreach ($f in $SortedTree[1..$sortedTree.count])
{
    Write-Progress -Activity 'Summarize PF Data' -Status "Processing $ic of $($SortedTree.count)" -CurrentOperation "Processing $($f.name)" -PercentComplete $($ic / $SortedTree.count * 100)
    $fc = 0;
    $mefc = 0;
    if ($true -eq $f.mailenabled)
    { $mefc++ }
    $sfeids = New-Object -TypeName System.Collections.ArrayList
    do
    {
        if ($sortedTree[$i].identity.replace('[', '`[').replace(']', '`]') + '\' -like $f.identity.replace('[', '`[').replace(']', '`]') + '\*')
        {
            [void]$sfeids.add($SortedTree[$i].entryID)
            if ($true -eq $SortedTree[$i].mailenabled)
            {
                $mefc++
            }
            $fc++;
            $i++;
        }
    }
    until ($sortedTree[$i].Identity.replace('[', '`[').replace(']', '`]') -notlike $f.Identity.replace('[', '`[').replace(']', '`]') + '\*');
    #Write-Verbose -Message "Folder Count for $($f.identity) is $fc" -Verbose;
    #Write-Verbose -Message "SubFolder EntryIDs are: $($sfeids -join ';')" -Verbose;
    #Write-Verbose -Message "MailEnabledFolder Count is $mefc" -Verbose;
    $f.SubFolderEntryIDs = $sfeids
    $f.FolderCount = $fc;
    $f.MailEnabledFolderCount = $mefc;
    $ic++;
    $i = $ic;
    #Write-Verbose -Message "Index Counter is $ic" -Verbose;
}
