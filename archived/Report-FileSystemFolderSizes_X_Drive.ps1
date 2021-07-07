$folderSizes = C:\Scripts\Get-DirStats.ps1 -Path D:\ -every

$folderHash = @{"_FolderTreeSize" = 0}

#$jaggedHashTableOfManagers["$manager"] += @($($user.AccountName).ToLower().Replace("sustainltd\",""))}
foreach ($folder in $folderSizes){
    $tier1 = "$($folder.Path.Split("\")[1])"
    $tier2 = "$($folder.Path.Split("\")[2])"
    $tier3 = "$($folder.Path.Split("\")[3])"

    if ($tier1 -ne $null -and $tier1 -ne "" ){ #Check that $tier1 isn't empty
        if(-not $folderHash.ContainsKey($tier1)){ #If it doesn't already exist, set it up
            $folderHash.Add($tier1, @{})
            $folderHash[$tier1].Add("_FolderTreeSize", 0)
            }
        $folderHash[$tier1].Set_Item("_FolderTreeSize", $($folderHash[$tier1]["_FolderTreeSize"] + $folder.Size)) #Then add the folder size to the _FolderTreeSize value
        }


    if ($tier2 -ne $null -and $tier2 -ne "" ){
        if (-not $folderHash[$tier1].ContainsKey($tier2)){
            $folderHash[$tier1].Add($tier2, @{})
            $folderHash[$tier1][$tier2].Add("_FolderTreeSize", 0)
            }
        $folderHash[$tier1][$tier2].Set_Item("_FolderTreeSize", $($folderHash[$tier1][$tier2]["_FolderTreeSize"] + $folder.Size))
        }


    if ($tier3 -ne $null -and $tier3 -ne "" ){
        if (-not $folderHash[$tier1][$tier2].ContainsKey($tier3) -and $tier3 -ne $null){
            $folderHash[$tier1][$tier2].Add($tier3, @{})
            $folderHash[$tier1][$tier2][$tier3].Add("_FolderTreeSize", 0)
            }
        $folderHash[$tier1][$tier2][$tier3].Set_Item("_FolderTreeSize", $($folderHash[$tier1][$tier2][$tier3]["_FolderTreeSize"] + $folder.Size))
        }

    #Finally, add the Size to thr root _FolderTreeSize
    $folderHash.Set_Item("_FolderTreeSize", $($folderHash["_FolderTreeSize"] + $folder.Size))
    }

$outHash = @{}

$outHash.Add("\",$folderHash["_FolderTreeSize"])
foreach($key in $folderHash.Keys){
    if ($key -ne "_FolderTreeSize"){$outHash.Add("\$Key",$folderHash[$Key]["_FolderTreeSize"])}
    foreach($key2 in $folderHash[$key].Keys){
        if ($key2 -ne "_FolderTreeSize"){$outHash.Add("\$Key\$key2",$folderHash[$key][$key2]["_FolderTreeSize"])}
        foreach($key3 in $folderHash[$key][$key2].Keys){
            if ($key3 -ne "_FolderTreeSize"){$outHash.Add("\$Key\$key2\$key3",$folderHash[$key][$key2][$key3]["_FolderTreeSize"])}
            }
        }
    }

$outHash.getEnumerator() | select name, value | Export-Csv '\\JimboJones\c$\Users\kevinm\Desktop\FileServerBreakdown.csv' -Encoding ASCII -NoTypeInformation

$outBigFolders= @()
foreach ($entry in $outHash.GetEnumerator()){
    if ($entry.Value -gt (2 * [Math]::Pow(1024,3))){
        $bigFolder = New-Object PSObject
        $bigFolder | Add-Member NoteProperty "FullPath" $entry.Name
        $bigFolder | Add-Member NoteProperty "Tier1" $entry.Name.Split("\")[1]
        $bigFolder | Add-Member NoteProperty "Tier2" $entry.Name.Split("\")[2]
        $bigFolder | Add-Member NoteProperty "Tier3" $entry.Name.Split("\")[3]
        $bigFolder | Add-Member NoteProperty "Bytes" $entry.Value
        $bigFolder | Add-Member NoteProperty "GB_$(Get-Date -Format "yyMMdd")" $($entry.Value / [Math]::Pow(1024,3))
        $outBigFolders += $bigFolder
        }
    }
$outBigFolders | Export-Csv "D:\Internal\ICT\Secure\DataStorage\X\FileServerBreakdown_BigFolders_$(Get-Date -Format "yyMMdd").csv" -Encoding ASCII -NoTypeInformation
