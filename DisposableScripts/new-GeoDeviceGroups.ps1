$userBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponseTeamsBots = get-graphTokenResponse -aadAppCreds $userBotDetails

$geoUGs = get-graphGroups -tokenResponse $tokenResponseTeamsBots -filterGroupType Unified -filterDisplayNameStartsWith "All (" -selectAllProperties
$geoUGs.displayName 


$geoUGs[0]

$groupNames = $("Devices - All","Devices - iOS","Devices - Android","Devices - Win10","Devices - MacOS","Devices - VMs")
$regions = convertTo-arrayOfStrings $($geoUGs | % {get-3lettersInBrackets $_.displayName})
$regions[0] = "North America" #Relies on "All" being 0th
$regions | % {
    $thisRegion = $_
    $groupNames | % {
        write-host  $($_.Replace("-","- $thisRegion -"))
        $newGroup = new-graphGroup -tokenResponse $tokenResponseTeamsBots -groupDisplayName $($_.Replace("-","- $thisRegion -")) -groupType Security -membershipType Assigned -groupDescription $($_.Replace("-","- $thisRegion -")) #-groupOwners @("1c43a72f-1221-4b1b-889d-4d045895ed63") #-groupMembers @("groupbot@anthesisgroup.com") 
        if($_ -match "All"){
            $thisAll = $newGroup
            $thisGeoUG = $geoUGs | ? {$_.displayname -match "($thisRegion)"}
            if(![string]::IsNullOrWhiteSpace($thisGeoUG)){
                set-graphGroup -tokenResponse $tokenResponseTeamsBots -groupId $thisGeoUG.id -groupUGSyncInfoExtensionHash @{deviceGroupId=$thisAll.id}
                }
            }
        else{add-graphUsersToGroup -tokenResponse $tokenResponseTeamsBots -graphGroupId $thisAll.id -memberType members -graphUserIds $newGroup.id}
        #continue
        }
    continue
    }


get-graphGroups -tokenResponse $tokenResponseTeamsBots -filterDisplayNameStartsWith "IT Ad"

$allDeviceGroups = get-graphGroups -tokenResponse $tokenResponseTeamsBots -filterDisplayNameStartsWith "Devices -"
$allRegionalDeviceGroups = $allDeviceGroups | ? {$_.displayName -notmatch "All"}

$groupNames | % {
    $thisDeviceGroup = $_
    #$newGroup = new-graphGroup -tokenResponse $tokenResponseTeamsBots -groupDisplayName $($($_.Replace("-","- All -"))) -groupType Security -membershipType Assigned -groupDescription $($($_.Replace("-","- All -"))) #-groupOwners @("1c43a72f-1221-4b1b-889d-4d045895ed63") #-groupMembers @("groupbot@anthesisgroup.com") 
    $newGroup = get-graphGroups -tokenResponse $tokenResponseTeamsBots -filterDisplayName $($($thisDeviceGroup.Replace("-","- All -"))) 
    if($thisDeviceGroup -match "All"){$relevantDevicegroups = $allDeviceGroups | ? {$_.displayName -match "Devices - All - " -and $_.displayName -notmatch "All - All"}}
    else{$relevantDevicegroups = $allRegionalDeviceGroups | ? {$_.displayName -match $($thisDeviceGroup.Split(" ")[2])}}
    if($relevantDevicegroups -ne $null){add-graphUsersToGroup -tokenResponse $tokenResponseTeamsBots -graphGroupId $newGroup.id -memberType members -graphUserIds $relevantDevicegroups.id}
    }