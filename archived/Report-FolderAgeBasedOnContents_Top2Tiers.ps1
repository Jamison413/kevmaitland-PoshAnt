param([string]$filePathRoot)
#$filePathRoot = "D:\Clients\A2 Dominion\100530-Feasibility - A2 - Arthur Sanct House\"
#$filePathRoot = "D:\Clients"
$outDirList = @{}

foreach($clientFolder in Get-ChildItem -Path $filePathRoot -Directory){
    $mostRecentAccessedDateClient = (Get-Date).AddYears(-555).ToShortDateString()
    $mostRecentModifiedDateClient = (Get-Date).AddYears(-555).ToShortDateString()
    foreach($projectFolder in Get-ChildItem -Path $clientFolder.FullName -Directory){
        $mostRecentAccessedDateProject = (Get-Date).AddYears(-555).ToShortDateString()
        $mostRecentModifiedDateProject = (Get-Date).AddYears(-555).ToShortDateString()
        foreach($file in Get-ChildItem -Path $projectFolder.FullName -File){
            if($file.LastAccessTime -gt (Get-Date $mostRecentAccessedDateProject)){$mostRecentAccessedDateProject = $file.LastAccessTime}
            if($file.LastWriteTime -gt (Get-Date $mostRecentModifiedDateProject)){$mostRecentModifiedDateProject = $file.LastWriteTime}
            if($file.LastAccessTime -gt (Get-Date $mostRecentAccessedDateClient)){$mostRecentAccessedDateClient = $file.LastAccessTime}
            if($file.LastWriteTime -gt (Get-Date $mostRecentModifiedDateClient)){$mostRecentModifiedDateClient = $file.LastWriteTime}
            }
        foreach($subfolder in Get-ChildItem -Path $projectFolder.FullName -Directory -Recurse){
            foreach($file in Get-ChildItem -Path $subFolder.FullName -File){
                if($file.LastAccessTime -gt (Get-Date $mostRecentAccessedDateProject)){$mostRecentAccessedDateProject = $file.LastAccessTime}
                if($file.LastWriteTime -gt (Get-Date $mostRecentModifiedDateProject)){$mostRecentModifiedDateProject = $file.LastWriteTime}
                if($file.LastAccessTime -gt (Get-Date $mostRecentAccessedDateClient)){$mostRecentAccessedDateClient = $file.LastAccessTime}
                if($file.LastWriteTime -gt (Get-Date $mostRecentModifiedDateClient)){$mostRecentModifiedDateClient = $file.LastWriteTime}
                }
            }
        $dirObj = New-Object psobject
        $dirObj | Add-Member NoteProperty "FullPath" $projectFolder.FullName
        $dirObj | Add-Member NoteProperty "NumTiers" $projectFolder.FullName.Split("\").Count
        $dirObj | Add-Member NoteProperty "LastAccessed" $mostRecentAccessedDateProject
        $dirObj | Add-Member NoteProperty "LastModified" $mostRecentModifiedDateProject
        $outDirList.Add($projectFolder.FullName+"\", $dirObj)
        }
    $dirObj = New-Object psobject
    $dirObj | Add-Member NoteProperty "FullPath" $clientFolder.FullName
    $dirObj | Add-Member NoteProperty "NumTiers" $clientFolder.FullName.Split("\").Count
    $dirObj | Add-Member NoteProperty "LastAccessed" $mostRecentAccessedDateClient
    $dirObj | Add-Member NoteProperty "LastModified" $mostRecentModifiedDateClient
    $outDirList.Add($clientFolder.FullName+"\", $dirObj)
    }



$outDirArray = @()
foreach($key in $outDirList.Keys){$outDirArray += $outDirList[$key]}
$outDirArray | Export-Csv "D:\Internal\ICT\Secure\DataStorage\X\FolderAgeBasedOnContent_$($filePathRoot.Replace(":",'').Replace("\",'_'))_$(Get-Date -Format "yyMMdd").csv" -Encoding ASCII -NoTypeInformation
