function enumerate-fsDirStats(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [psobject]$fsDirStatBlob
        ,[parameter(Mandatory = $false)]
        [psobject]$prependPath = ""
        )
    $outBlob = New-Object psobject -Property @{
        Path = $prependPath+"\"+$fsDirStatBlob.Name
        Size = $fsDirStatBlob.Size
        LastModified = $fsDirStatBlob.LastModified
        Tier = $prependPath.Split("\").Count
        }
    $outBlob
    $fsDirStatBlob.Subfolders.Keys | % {
        $test = $_
        enumerate-fsDirStats -fsDirStatBlob $fsDirStatBlob.SubFolders[$test] -prependPath $($prependPath+"\"+$_)
        }

    }
function new-dirStatObject(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $false)]
        [datetime]$duffDate = [datetime]$(get-date).AddYears(-100)
        )
    New-Object psobject -Property @{
        Size = [long]0
        LastModified = $duffDate
        #Subfolders = [ordered]@{}
        Subfolders = @{}
        }
    }
Function scrape-subitems(){
    [CmdletBinding()]
    param([string]$rootPath)
    $params = New-Object System.Collections.Arraylist
    $params.AddRange(@("/L","/S","/XJ","/XJD","/NJH","/NJS","/BYTES","/FP","/NC","/NDL","/TS","/R:0","/W:0"))
    #$countPattern = "^\s{3}Files\s:\s+(?<Count>\d+).*"
    #$sizePattern = "^\s{3}Bytes\s:\s+(?<Size>\d+(?:\.?\d+)\s[a-z]?).*"
    $rootPath = $rootPath.TrimEnd("\")
    write-host $rootPath
    (robocopy $rootPath NULL $params) | ForEach {
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
function update-dirStats(){
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [psobject]$roboCopyLine

        ,[parameter(Mandatory = $true)]
        [psobject]$dirStatsObject
        )
    

    }


$startPoint = "G:\Shares and data\Documents & Data\"
$fsBlob = new-dirStatObject

scrape-subitems $startPoint | % {
    $temp = $_
    $remainingPath = ((Split-Path $temp.FullName)+"\").Replace($startPoint,"").Replace("\\","\")
    switch($remainingPath.Split("\").Count){
        {$_ -eq 1} {
            $fsBlob.Size += $temp.Size
            if($fsBlob.LastModified -lt $temp.Date){$fsBlob.LastModified = $temp.Date}
            }
        {$_ -eq 2} {
            if($fsBlob.Subfolders.Keys -notcontains $remainingPath.Split("\")[0]){
                $fsBlob.Subfolders.Add($remainingPath.Split("\")[0],$(new-dirStatObject))
                }
            $fsBlob.Size += $temp.Size
            $fsBlob.Subfolders[$remainingPath.Split("\")[0]].Size  += $temp.Size
            if($fsBlob.LastModified -lt $temp.Date){$fsBlob.LastModified = $temp.Date}
            if($fsBlob.Subfolders[$remainingPath.Split("\")[0]].LastModified -lt $temp.Date){$fsBlob.Subfolders[$remainingPath.Split("\")[0]].LastModified = $temp.Date}
            }
        {$_ -eq 3} {
            if($fsBlob.Subfolders.Keys -notcontains $remainingPath.Split("\")[0]){
                $fsBlob.Subfolders.Add($remainingPath.Split("\")[0],$(new-dirStatObject))
                }
            if($fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders.Keys -notcontains $remainingPath.Split("\")[1]){
                $fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders.Add($remainingPath.Split("\")[1],$(new-dirStatObject))
                }
            $fsBlob.Size += $temp.Size
            $fsBlob.Subfolders[$remainingPath.Split("\")[0]].Size  += $temp.Size
            $fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].Size  += $temp.Size
            if($fsBlob.LastModified -lt $temp.Date){$fsBlob.LastModified = $temp.Date}
            if($fsBlob.Subfolders[$remainingPath.Split("\")[0]].LastModified -lt $temp.Date){$fsBlob.Subfolders[$remainingPath.Split("\")[0]].LastModified = $temp.Date}
            if($fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].LastModified -lt $temp.Date){$fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].LastModified = $temp.Date}
            }        
        {$_ -gt 3} {
            if($fsBlob.Subfolders.Keys -notcontains $remainingPath.Split("\")[0]){
                $fsBlob.Subfolders.Add($remainingPath.Split("\")[0],$(new-dirStatObject))
                }
            if($fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders.Keys -notcontains $remainingPath.Split("\")[1]){
                $fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders.Add($remainingPath.Split("\")[1],$(new-dirStatObject))
                }
            if($fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].Subfolders.Keys -notcontains $remainingPath.Split("\")[2]){
                $fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].Subfolders.Add($remainingPath.Split("\")[2],$(new-dirStatObject))
                }
            $fsBlob.Size += $temp.Size
            $fsBlob.Subfolders[$remainingPath.Split("\")[0]].Size  += $temp.Size
            $fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].Size  += $temp.Size
            $fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].Subfolders[$remainingPath.Split("\")[2]].Size  += $temp.Size
            if($fsBlob.LastModified -lt $temp.Date){$fsBlob.LastModified = $temp.Date}
            if($fsBlob.Subfolders[$remainingPath.Split("\")[0]].LastModified -lt $temp.Date){$fsBlob.Subfolders[$remainingPath.Split("\")[0]].LastModified = $temp.Date}
            if($fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].LastModified -lt $temp.Date){$fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].LastModified = $temp.Date}
            if($fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].Subfolders[$remainingPath.Split("\")[2]].LastModified -lt $temp.Date){$fsBlob.Subfolders[$remainingPath.Split("\")[0]].Subfolders[$remainingPath.Split("\")[1]].Subfolders[$remainingPath.Split("\")[2]].LastModified = $temp.Date}
            }
default {write-host "Uh-oh"}
        }
    }

     
$outblobs = @()
$outblobs += enumerate-fsDirStats -fsDirStatBlob $fsBlob
$outblobs | Export-Csv -Path C:\Users\kev.maitland\folderdata3.csv -NoTypeInformation 
