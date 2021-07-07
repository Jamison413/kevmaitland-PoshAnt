param([string]$filePathRoot)
#$filePathRoot = "D:\Clients\A2 Dominion\100530-Feasibility - A2 - Arthur Sanct House\"
#$filePathRoot = "D:\Clients"
$outDirList = @{}

foreach($directory in Get-ChildItem -Path $filePathRoot -Recurse -Directory){
    $mostRecentAccessedDate = $null
    $mostRecentModifiedDate = $null
    foreach($file in Get-ChildItem -Path $directory.FullName -File){
        if($file.LastAccessTime -gt $mostRecentAccessedDate){$mostRecentAccessedDate = $file.LastAccessTime}
        if($file.LastWriteTime -gt $mostRecentModifiedDate){$mostRecentModifiedDate = $file.LastWriteTime}
        }
    $dirObj = New-Object psobject
    $dirObj | Add-Member NoteProperty "FullPath" $directory.FullName
    $dirObj | Add-Member NoteProperty "NumTiers" $directory.FullName.Split("\").Count
    $dirObj | Add-Member NoteProperty "LastAccessed" $mostRecentAccessedDate
    $dirObj | Add-Member NoteProperty "LastModified" $mostRecentModifiedDate
    $outDirList.Add($directory.FullName+"\", $dirObj)
    }

foreach($dir in $outDirList.Keys){
    $i = $filePathRoot.Split("\").Count -1
    $dirParent = $filePathRoot
    do{
        $dirParent += $dir.Split("\")[$i]+"\"
        #Write-Host -ForegroundColor Cyan "`$dir = $dir"
        #Write-Host -ForegroundColor DarkYellow "`$dirParent = $dirParent"
        #Write-Host -ForegroundColor DarkYellow "`$i = $i"
        #Write-Host -ForegroundColor DarkYellow "`$outDirList[`$dirParent].LastAccessed = $($outDirList[$dirParent].LastAccessed)"
        if ($outDirList[$dirParent].LastAccessed -lt $outDirList[$dir].LastAccessed){$outDirList[$dirParent].LastAccessed = $outDirList[$dir].LastAccessed}
        if ($outDirList[$dirParent].LastModified -lt $outDirList[$dir].LastModified){$outDirList[$dirParent].LastModified = $outDirList[$dir].LastModified}
        $i ++
        }
    while($i -lt $dir.Split("\").Count -1)
    }

$outDirArray = @()
foreach($key in $outDirList.Keys){$outDirArray += $outDirList[$key]}
$outDirArray | Export-Csv "D:\Internal\ICT\Secure\DataStorage\X\FolderAgeBasedOnContent_$($filePathRoot.Replace(":",'').Replace("\",'_'))_$(Get-Date -Format "yyMMdd").csv" -Encoding ASCII -NoTypeInformation
