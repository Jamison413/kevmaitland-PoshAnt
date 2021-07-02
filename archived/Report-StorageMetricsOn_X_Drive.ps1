$numberOfTiersToUsePowershellFor = 3 #How many tiers beyond $commonPathBeginsAt

$archiveStub = "\\Sustainltd.local\archiveddata\Archives\X"
#$sourceStub = "\\sustainltd.local\Data\Clients\53 Shepherds Hill (Residents Association" #This is also the format the foldersToArchive list will be in
$sourceStub = "\\Sustainltd.local\Data"
$commonPathBeginsAt = $sourceStub.Split("\").Count
#$foldersToArchive = Get-Content "\\sustainltd.local\data\Internal\ICT\Secure\DataStorage\x\160219_FoldersToArchive.txt"
#$dirOutObj = @{}
$dirOutArray = @()

Function getSubfolderData([string]$shortFolderPath){
    $params = New-Object System.Collections.Arraylist
    $params.AddRange(@("/L","/S","/XJ","/XJD","/NJH","/NJS","/BYTES","/FP","/NC","/NDL","/TS","/R:0","/W:0"))
    $countPattern = "^\s{3}Files\s:\s+(?<Count>\d+).*"
    $sizePattern = "^\s{3}Bytes\s:\s+(?<Size>\d+(?:\.?\d+)\s[a-z]?).*"
    ((robocopy $shortFolderPath NULL $params)) | ForEach {
        If ($_ -match "(?<Size>\d+)\s(?<Date>\S+\s\S+)\s+(?<FullName>.*)") {
            New-Object PSObject -Property @{
                FullName = $matches.FullName
                Size = $matches.Size
                Date = [datetime]$matches.Date
                }
            } 
            Else {Write-Verbose ("{0}" -f $_)}
        }
    }
Function makeDirObject([string]$pFullPath,[string]$pNumTiers,[string]$pSize,[string]$pLastMod){
    $dirObj = New-Object psobject
    $dirObj | Add-Member NoteProperty "FullPath" $pFullPath
    $dirObj | Add-Member NoteProperty "NumTiers" $pNumTiers
    $dirObj | Add-Member NoteProperty "Size" $pSize
    $dirObj | Add-Member NoteProperty "LastModified" $pLastMod
    $dirObj
    }

Write-Host -ForegroundColor Yellow "Generating folder list..."
$i=0
$foldersToProcess = @()
$foldersToProcess += , (Get-Item $sourceStub)
do{
    Write-Host -ForegroundColor DarkYellow "`tGenerating Tier $i"
    $stopCounter = $foldersToProcess.Count
    $nextTierOfFolders = @()
    for ($j=0;$j -lt $stopCounter; $j++){
        #$nextTierOfFolders += Get-ChildItem $foldersToProcess[$j].FullName -Directory
        $foldersToProcess += Get-ChildItem $foldersToProcess[$j].FullName -Directory
        }
    #$foldersToProcess = $nextTierOfFolders
    $i++
    }
until ($i -ge $numberOfTiersToUsePowershellFor)

Write-Host -ForegroundColor Yellow "Processing folder list with RoboCopy..."
$i=0
foreach($folder in $foldersToProcess){
    if($folder.FullName.Split("\")[$commonPathBeginsAt] -ne $lastTier1){Write-Host -ForegroundColor DarkYellow "`tProcessing $($folder.FullName.Split("\")[$commonPathBeginsAt])...";$i=0}
    $folderData = getSubfolderData -shortFolderPath $folder.FullName #= get-item "\\Sustainltd.local\Data\Internal\ICT"
    $size = $($folderData | Measure-Object -Property Size -Sum).Sum
    $lastMod = $($folderData | Measure-Object -Property Date -Maximum).Maximum
    $dirOutArray += makeDirObject -pFullPath  $folder.FullName -pNumTiers $folder.FullName.Split("\").Count -pSize $size -pLastMod $lastMod
    $lastTier1 = $folder.FullName.Split("\")[$commonPathBeginsAt]
    $i++
    if($i%10 -eq 0){write-host -foregroundcolor DarkYellow "`t`t$i processed..."}
    }

$dirOutArray | Export-Csv "\\SustainLtd.local\Data\Internal\ICT\Secure\DataStorage\X\StorageMetrics_$(Get-Date -Format "yyMMdd").csv" -Encoding ASCII -NoTypeInformation




#    for($i=$commonPathBeginsAt;$i -lt $folder.FullName.Split("\").Count;$i++){
#        switch($i){
#            $commonPathBeginsAt+0 {if($dirOutObj.Keys -notcontains $folder.FullName.Split("\")[$i]){
#                $dirOutObj += @{$folder.FullName.Split("\")[$commonPathBeginsAt] = makeDirObject -pFullPath  $folder.FullName -pNumTiers $folder.FullName.Split("\").Count -pSize $size -pLastMod $lastMod}}
#                }
#            $commonPathBeginsAt+1 {if($dirOutObj[$folder.FullName.Split("\")[$commonPathBeginsAt]].Keys -notcontains $folder.FullName.Split("\")[$i]){$dirOutObj += makeDirObject -pFullPath  $folder.FullName -pNumTiers $folder.FullName.Split("\").Count -pSize $size -pLastMod $lastMod}}
#            $commonPathBeginsAt+2 {}
#            $commonPathBeginsAt+3 {}#
#
#            }
#        
#        }#
##
#
#    $dirObj = New-Object psobject
#    $dirObj | Add-Member NoteProperty "FullPath" $folder.FullName
#    $dirObj | Add-Member NoteProperty "NumTiers" $folder.FullName.Split("\").Count
#    $dirObj | Add-Member NoteProperty "Size" $size
#    $dirObj | Add-Member NoteProperty "LastModified" $lastMod
#
#    if($folder.FullName.Split("\")[$commonPathBeginsAt]){}
