if($PSCommandPath){
    $InformationPreference = 2
    $VerbosePreference = 0
    $logFileLocation = "C:\ScriptLogs\"
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))$suffix`_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    Start-Transcript $transcriptLogName -Append
    }
else{
    $logFileLocation = "C:\ScriptLogs\"
    $transcriptLogName = "$($logFileLocation)delete-previousVersions_Transcript_$(Get-Date -Format "yyyy-MM-dd").log"
    }

$tokenSharePoint = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName SharePointBot) -grant_type client_credentials

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\KimbleBot.txt)
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/clients" -Credentials $adminCreds


$clientSiteId = "anthesisllc.sharepoint.com,68fbfc7c-e744-47bb-9e0b-9b9ee057e9b5,faed84bc-70be-4e35-bfbf-cdab31aeeb99"
$allClientDrives = get-graphDrives -tokenResponse $tokenSharePoint -siteGraphId $clientSiteId

$allClientDrives | % {
    $thisDrive = $_
    $thisList = get-graphList -tokenResponse $tokenSharePoint -graphDriveId $thisDrive.id
    Write-Output "Setting Versions=50 on [$($thisList.displayName)]"
    Set-PnPList -Identity $thisList.id -MajorVersions 50
    }

<# Test with a big folder (BEIS)
$beis = $($allClientDrives | ? {$_.webUrl -match "https://anthesisllc.sharepoint.com/clients/BEIS"})[0]
$fileToTest = get-graphDriveItems -tokenResponse $tokenSharePoint -driveGraphId $beis.id -folderPathRelativeToRoot "/P-1005764%20202111%20BEIS%20Heat%20Network%20Zoning%20Pilot/Data%20%26%20refs/GIS%20layers/GIS%20layers/OS%20data/Zoomstack%20cropped/OS_Open_Zoomstack%20local_buildings%20(3).gpkg" -returnWhat Item
$fileToTest2 = get-graphDriveItems -tokenResponse $tokenSharePoint -driveGraphId $beis.id -folderPathRelativeToRoot "/University of Lancaster DHN (E006118)/Data & refs/Site Visit Photos 12112019/20191112_135641.jpg" -returnWhat Item
$prevVersions = invoke-graphGet -tokenResponse $tokenSharePoint -graphQuery "/drives/$($beis.id)/items/$($fileToTest.id)/versions"
$prevVersions[166]
$testDelete = invoke-graphDelete -tokenResponse $tokenSharePoint -graphQuery "/drives/$($beis.id)/items/$($fileToTest.id)/versions/$($prevVersions[166].id)"
for($i=50; $i -lt $prevVersions.Count; $i++){
    $result = invoke-graphDelete -tokenResponse $tokenSharePoint -graphQuery "/drives/$($beis.id)/items/$($fileToTest.id)/versions/$($prevVersions[$i].id)"
    }

$beisList = get-graphList -tokenResponse $tokenSharePoint -graphDriveId $beis.id
$beisPnpList = Get-PnPList -Identity $beisList.id
$rootDriveItems = get-graphDriveItems -tokenResponse $tokenSharePoint -driveGraphId $beis.id -returnWhat Children

$pnpTest = Measure-Command { $rootDriveItems = Get-PnPListItem -List $beisList.id -PageSize 5000}
$graphTest =  Measure-Command { $graphItems = get-graphDriveItems -tokenResponse $tokenSharePoint -driveGraphId $beis.id -returnWhat Children -recursive} #There is no way to dump all DriveItems, so recursively querying Folders for children is the only way to get everything. And it's really slow.
$graphListTest =  Measure-Command { $graphListItems = invoke-graphGet -tokenResponse $tokenSharePoint -graphQuery "/drives/$($beis.id)/list/items"} #The /list endpoint *does* dump all ListItems, but ListItems are less helpful than DriveItems.
#>
#PNP is fastest (8 seconds), GraphList is 2nd (50 seconds), GraphDriveItems is slowest (500 seconds)
#Get-PnpListItem CAML is shite and can't combine -Query with -PageSize, so it fails on any List with >5000 items (https://github.com/pnp/PnP-PowerShell/issues/879), so we have to dump everythign and filter client-side. Still the fastest solution.

$totalStorageRecovered = 0
for($i=0; $i -lt $allClientDrives.Count; $i++){ #Iterate through all DocLibs in /clients
    $tokenSharePoint = test-graphBearerAccessTokenStillValid -tokenResponse $tokenSharePoint -renewTokenExpiringInSeconds 60 -aadAppCreds $(get-graphAppClientCredentials -appName SharePointBot)
    $thisDrive = $allClientDrives[$i]
    $storageSpaceRecoveredFromDocLib = 0 #Reset reclaimed storage count for this DocLib
    Write-Progress -activity "Processing DocLibs in Clients" -Status "[$i/$($allClientDrives.count)]" -PercentComplete ($i/ $allClientDrives.count *100)
    Write-Output "Processing [$($thisDrive.name)]"
    $thisList = get-graphList -tokenResponse $tokenSharePoint -graphDriveId $thisDrive.id
    $theseItems =  Get-PnPListItem -List $thisList.id -PageSize 5000
    $theseItemsToPrune = $theseItems | ? {$_.FieldValues.owshiddenversion -ge 50} #Get Files with >50 versions (not all versions will still exist, but this is a quick way of ruling out any files that *cannot* have >50 versions)
    Write-Output "`tProcessing [$($theseItemsToPrune.Count)] items"
    for($j=0; $j -lt $theseItemsToPrune.Count; $j++){
        Write-Progress -activity "Processing [$($theseItemsToPrune.Count)] files in [$($thisDrive.name)]" -Status "[$j/$($theseItemsToPrune.count)]" -PercentComplete ($j/ $theseItemsToPrune.count *100)
        $thisItemToPruneVersions = Get-PnPFileVersion -Url $theseItemsToPrune[$j].FieldValues.FileRef #Get the PreviousVersions
        $thisItemToPruneVersions = $thisItemToPruneVersions | Sort-Object Id -Descending #Explicitly sort the PreviousVersions, making the most recent top of the array
        Write-Output "`t`t[$($thisItemToPruneVersions.Count)] Previous Versions found for [$($theseItemsToPrune[$j].FieldValues.FileRef)]"
        $storageSpaceRecoveredFromFile = 0
        for ($k=50; $k -lt $thisItemToPruneVersions.Count; $k++){ #Skip the most recent 50 versions
            $storageSpaceRecoveredFromFile += $thisItemToPruneVersions[$k].Size
            try{
                #Remove-PnPFileVersion -Url $theseItemsToPrune[$j].FieldValues.FileRef -Identity $thisItemToPruneVersions[$k].ID -Recycle -Force #Remove any Previous Versions >50
                }
            catch{
                [array]$errorLog += New-Object -TypeName psobject @{DocLibId=$thisDrive.id;DocLibName=$thisDrive.name;FileId=$theseItemsToPrune[$j].id;FileName=$theseItemsToPrune[$j].FieldValues.FileLeafRef;webUrl=$theseItemsToPrune[$j].FieldValues.FileRef;PreviousVersionId=$thisItemToPruneVersions[$k].id;ErrorString=$_}
                }

            }
        Write-Output "`t`t`t[$([Math]::Round($storageSpaceRecoveredFromFile/1MB,2))]MB recovered from [$([Math]::Max(0,$thisItemToPruneVersions.Count-50))] deleted Previous Versions [$($thisItemToPruneVersions.Count)] Previous Versions were available"
        $storageSpaceRecoveredFromDocLib += $storageSpaceRecoveredFromFile
        }
    Write-Output "`t[$([Math]::Round($storageSpaceRecoveredFromDocLib/1GB,2))]GB recovered from [$([Math]::Max(0,$theseItemsToPrune.Count))] files in [$($thisDrive.webUrl)]"
    $totalStorageRecovered += $storageSpaceRecoveredFromDocLib
    }
Write-Output "`t[$([Math]::Round($totalStorageRecovered/1GB,2))]GB recovered from [$([Math]::Max(0,$allClientDrives.Count))] DocLibs in Clients Site"
Write-Output ""

<#CAML is shite
$query = "<View Scope=`"Recursive`">
 <Query>
<Where>
<Geq>
    <FieldRef Name='owshiddenversion' />
    <Value Type='Integer'>51</Value>
  </Geq>
</Where></Query></View>"


$query = New-Object -TypeName Microsoft.SharePoint.Client.CamlQuery
$query.ViewXml="<View Scope=`"Recursive`">
<Query>
<Where>
<Geq>
<FieldRef Name = 'owshiddenversion'/>
<Value Type = 'Integer'>51</Value>
</Geq>
</Where>
</Query><RowLimit Paged='TRUE'>1000</RowLimit>
</View>"
$query.AllowIncrementalResults = $true
$query.ListItemCollectionPosition = "PreviousListItemCollection?.ListItemCollectionPosition"

$rootDriveItemsToPrune = Get-PnPListItem -List $beisList.id -PageSize 5000 -Query $query

Get-PnPListItem -List $beisList.id -Id $fileToTest.id


$query = "<View>
 <Query>
<Where>
<Or>
<Contains>
    <FieldRef Name='eTag' />
    <Value Type='Text'>,51</Value>
  </Contains>
<Or>
<Contains>
    <FieldRef Name='eTag' />
    <Value Type='Text'>,52</Value>
  </Contains>
<Or>
  <Contains>
    <FieldRef Name='eTag' />
    <Value Type='Text'>,53</Value>
  </Contains>
  <Contains>
    <FieldRef Name='eTag' />
    <Value Type='Text'>,54</Value>
  </Contains>
</Or>
</Or>
</Or>
</Where>"




$query = "<View>
 <Query>
<Where>"
$queryEnd = ""
for ($i=51;$i -lt 52;$i++){
    $query += "<Or>
<Contains>
    <FieldRef Name='eTag' />
    <Value Type='Text'>,$i</Value>
  </Contains>"
    $queryEnd += "</Or>
    "
    }

$fullQuery = $query+$queryEnd+"</Where></Query></View>"
#>

<# Checking Recycle Bin to see if reducing Max Versions automatiaclly purges Previous Versions down the the limit (of course it fucking doesn't).
$rbin = Get-PnPRecycleBinItem
$rbin1 = Get-PnPRecycleBinItem -FirstStage
$rbin2 = Get-PnPRecycleBinItem -SecondStage
$fullBin = $rbin | Measure-Object -Property Size -Sum

([Math]::Round($($rbin | Measure-Object -Property Size -Sum).Sum/1GB,2))
([Math]::Round($($rbin1 | Measure-Object -Property Size -Sum).Sum/1GB,2))
([Math]::Round($($rbin2 | Measure-Object -Property Size -Sum).Sum/1GB,2))

$rbin1[2000] | select *

#deleted when?
$rbin1 | Group-Object -Property {$(Get-Date -f d $_.DeletedDate)} | sort-object Count

"https://anthesisllc.sharepoint.com/:u:/r/clients/BEIS/P-1005764%20202111%20BEIS%20Heat%20Network%20Zoning%20Pilot/Data%20%26%20refs/GIS%20layers/GIS%20layers/OS%20data/Zoomstack%20cropped/OS_Open_Zoomstack%20local_buildings%20(3).gpkg"
#>