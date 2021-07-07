#Script to allow non-administrators to connect in and see who has a specific file open

param([string]$pathToFile)
#$pathToFile = "D:\Internal\Operations\Resources Tracker v1.1.xlsm"
$localPath = $pathToFile.Replace("X:\","D:\").Replace("\\sustainltd.local\data\","D:\")

$result = Get-SmbOpenFile | ? {$_.Path -eq $localPath}
if ($result -ne $null){
    if ($result.ClientComputerName -ne ""){
        $name = $(nslookup $result.ClientComputerName)[3].replace("Name:    ","")
        }
    "$($result.ClientUserName.Replace("SUSTAINLTD\",'').replace("."," ")) on $name"
    }
    else{"No-one has $pathToFile open!"}