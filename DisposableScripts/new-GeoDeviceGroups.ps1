$userBotDetails = get-graphAppClientCredentials -appName TeamsBot
$userBotTokenResponse = get-graphTokenResponse -aadAppCreds $userBotDetails

$geoUGs = get-graphGroups -tokenResponse $userBotTokenResponse -filterGroupType Unified -filterDisplayNameStartsWith "All (" -selectAllProperties
$geoUGs.displayName 


$geoUGs[0]

$groupNames = $("Devices - All","Devices - iOS","Devices - Android","Devices - Win10","Devices - MacOS","Devices - VMs")
$regions = convertTo-arrayOfStrings $($geoUGs | % {get-3lettersInBrackets $_.displayName})
$regions[0] = "North America" #Relies on "All" being 0th
$regions | % {
    $thisRegion = $_
    $groupNames | % {
        write-host  $($_.Replace("-","- $thisRegion -"))
        $newGroup = new-graphGroup -tokenResponse $userBotTokenResponse -groupDisplayName $($_.Replace("-","- $thisRegion -")) -groupType Security -membershipType Assigned -groupDescription $($_.Replace("-","- $thisRegion -")) #-groupOwners @("1c43a72f-1221-4b1b-889d-4d045895ed63") #-groupMembers @("groupbot@anthesisgroup.com") 
        if($_ -match "All"){
            $thisAll = $newGroup
            $thisGeoUG = $geoUGs | ? {$_.displayname -match "($thisRegion)"}
            if(![string]::IsNullOrWhiteSpace($thisGeoUG)){
                set-graphGroup -tokenResponse $userBotTokenResponse -groupId $thisGeoUG.id -groupUGSyncInfoExtensionHash @{deviceGroupId=$thisAll.id}
                }
            }
        else{add-graphUsersToGroup -tokenResponse $userBotTokenResponse -graphGroupId $thisAll.id -memberType members -graphUserIds $newGroup.id}
        #continue
        }
    continue
    }


get-graphGroups -tokenResponse $userBotTokenResponse -filterDisplayNameStartsWith "IT Ad"
