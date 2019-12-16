Function Export-PublicFolderToPST
{
    Param(
        [Parameter(Mandatory)]
        [string]$File,
        [Parameter()]
        [string]$PstPath,
        [Parameter()]
        [string]$LogPath
    )

    Function Get-TimeStamp
    {

        "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)

    }


    Function Write-Log
    {
        param(
            $Log
            ,
            $LogPath
        )

        Out-File -FilePath $LogPath -Append -InputObject $Log
    }

    #Import Data file to process against
    $Data = Import-Csv $File

    #Initialize Variables
    If (!($PstPath))
    {
        $PstPath = "C:\temp"
    }
    If (!($LogPath))
    {
        $LogPath = "C:\temp\LogFiles\"
    }

    $LogFile = $LogPath + (hostname) + "-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".txt"
    $WarningLog = $LogPath + (hostname) + "-Warning-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".txt"
    $i = 1

    #Transcript
    Start-Transcript -path $LogFile

    #Open up outlook
    $null = [Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook")
    $Outlook = New-Object -comobject Outlook.Application
    $Namespace = $Outlook.GetNameSpace("MAPI")
    $StoreID = ($Namespace.Stores | Where-Object { $_.DisplayName -like "Public Folders*" }).Storeid

    Foreach ($Obj in $Data)
    {

        #Define Location Reference
        $EntryID = $Obj.EntryID
        $Name = $Obj.Name
        $Guid = $Obj.Guid
        $PstFQDN = $PstPath + "\" + $Guid + ".pst"

        #Begin Processing
        Write-Progress -Activity "Exporting Public Folders" -Status "Working on folder: $($Name)" -PercentComplete (($i / ($Data.count)) * 100)
        Write-Information -MessageData "$(Get-TimeStamp) Processing Folder $($Name) - Folder Number $($i) of $($Data.count)" -InformationAction Continue
        Write-Information -MessageData "$(Get-TimeStamp) -FolderPath $($Obj.Identity)" -InformationAction Continue
        Write-Information -MessageData "$(Get-TimeStamp) -EntryID $($Obj.EntryID)" -InformationAction Continue
        Write-Information -MessageData "$(Get-TimeStamp) -PST File $($PstFQDN)" -InformationAction Continue
        Write-Information -MessageData "$(Get-TimeStamp) -Folder contains $($Obj.Itemcount) items for $(Invoke-SumSize $Obj.Totalitemsize)" -InformationAction Continue

        #Add PST to OLK Profile
        Write-Information -MessageData "$(Get-TimeStamp) -Added PST to Outlook" -InformationAction Continue

        $Namespace.AddStore($PstFQDN)
        $PstFolder = $Namespace.Session.Folders | Where-Object { $_.name -eq "Outlook Data File" }

        #Export to PST
        $Folders = $Namespace.GetFolderFromID($EntryID, $StoreID)

        Write-Information -MessageData "$(Get-TimeStamp) -Copying data to PST" -InformationAction Continue

        $null = $Folders.CopyTo($PstFolder)

        Write-Information -MessageData "$(Get-TimeStamp) -Finished Copying data to PST" -InformationAction Continue

        #Validate item counts in PST
        $Pst = $Namespace.Stores | Where-Object { $_.FilePath -eq $PstFQDN }
        $PstRoot = $Pst.GetRootFolder()
        $PstTopLevelFolders = $PstRoot.Folders
        $PstItemCounter = 0

        Write-Information -MessageData "$(Get-TimeStamp) -Validating PST Item Counts" -InformationAction Continue

        Foreach ($PstTopLevelFolder in $PstTopLevelFolders)
        {
            $PstCount = $PstTopLevelFolder.items.count
            $PstItemCounter += $PstCount
        }

        If ($PstItemCounter -eq $Obj.Itemcount)
        {
            Write-Information -MessageData "$(Get-TimeStamp) -PST contents match values in data file" -InformationAction Continue
        }

        Elseif ($PstItemCounter -lt $Obj.Itemcount)
        {
            Write-Warning -Message "$(Get-TimeStamp) PST contains few items than data file. Check warning log."
            Write-Log ("====== $(Get-TimeStamp) ======")
            Write-Log ("File: " + $PstFQDN)
            Write-Log ("-Item Count in Data file: " + ($Obj.Itemcount) + ", PSTfile:" + $PstItemCounter)
            Write-Log ("-EntryID: " + $Obj.EntryID)
            Write-Log ("-Folderpath: " + $Obj.Identity)
            # Add an alternate log stream with JSON output for later conversion back to objects/structured data for reporting
            # another random comment
        }

        #Rename and Remove PST
        $PstFolder.Name = $Name
        $PstStore = $Namespace.Session.Folders.GetLast()

        Write-Information -MessageData "$(Get-TimeStamp) -Removing PST from Outlook" -InformationAction Continue
        Write-Information -MessageData "" -InformationAction Continue
        Write-Information -MessageData "$(Get-TimeStamp) Processing Data File Percent Complete: $(($i/($Data.count)).tostring("P0"))" -InformationAction Continue

        $Namespace.RemoveStore($PstStore)

        Write-Information -MessageData "" -InformationAction Continue
        Write-Information -MessageData "--------------------------------------------------------------------------------------------" -InformationAction Continue
        Write-Information -MessageData "" -InformationAction Continue

        $i++
        #All done

    }
}
