start-transcriptLog -thisScriptName "update-geoDeviceGroups"

$userBotDetails = get-graphAppClientCredentials -appName TeamsBot
$userBotTokenResponse = get-graphTokenResponse -aadAppCreds $userBotDetails

Write-host "Getting Geographic Unified Groups"
$geoUGs = get-graphGroups -tokenResponse $userBotTokenResponse -filterGroupType Unified -filterDisplayNameStartsWith "All (" -selectAllProperties
$geoUGs = $geoUGs.Where({($_.displayName -ne "All (All)") -and ($_.displayName -ne "All (EMELA)") })


Write-host "`tRetrieved [$($geoUGs.Count)] Geographic Unified Groups"


Write-host "Getting all AAD devices"
$allAadDevices = get-graphDevices -tokenResponse $userBotTokenResponse -includeOwners
$allAadDevices | ForEach-Object {Add-Member -InputObject $_ -MemberType NoteProperty -Name OwnerId -Value $_.registeredOwners[0].id}
$allAadDevices = $allAadDevices.Where({($_.Id -ne "58c0ce09-5ca8-4f33-a369-fefeb42a6fd3")}) #exclude Netmon

Write-host "`tRetrieved [$($allAadDevices.Count)] AD devices"


If(($geoUGs) -and ($allAadDevices)){

$duffDevices = @()
$geoUGs | ForEach-Object {
    $thisGeoUG = $_
    write-host "Processing [$($thisGeoUG.displayName)]"
    #continue
    if([string]::IsNullOrWhiteSpace($thisGeoUG.anthesisgroup_UGSync.deviceGroupId)){
        $allGroup = get-graphGroups -tokenResponse $userBotTokenResponse -filterDisplayName "Devices - $(get-3lettersInBrackets $thisGeoUG.displayName) - All"
        if([string]::IsNullOrWhiteSpace($allGroup)){
            $allGroup = new-graphGroup -tokenResponse $userBotTokenResponse -groupDisplayName "Devices - $(get-3lettersInBrackets $thisGeoUG.displayName) - All" -groupType Security -membershipType Assigned -groupDescription "Device Group for $(get-3lettersInBrackets $thisGeoUG.displayName) - All"
        }
        set-graphGroup -tokenResponse $userBotTokenResponse -groupId $thisGeoUG.id -groupUGSyncInfoExtensionHash @{deviceGroupId=$allGroup.id}
        $allAllGroup = get-graphGroups -tokenResponse $userBotTokenResponse -filterDisplayName "Devices - All - All" 
        add-graphUsersToGroup -tokenResponse $userBotTokenResponse -graphGroupId $allAllGroup.id -memberType members -graphUserIds $allGroup.id
        $subgroups = @("Win10","Win11","Win_Other","iOS","VMs","Android","MacOS")
        foreach ($group in $subgroups) {
            $subGroup = get-graphGroups -tokenResponse $userBotTokenResponse -filterDisplayName "Devices - $(get-3lettersInBrackets $thisGeoUG.displayName) - $group"
            if([string]::IsNullOrWhiteSpace($subGroup)){
                $subGroup = new-graphGroup -tokenResponse $userBotTokenResponse -groupDisplayName "Devices - $(get-3lettersInBrackets $thisGeoUG.displayName) - $group" -groupType Security -membershipType Assigned -groupDescription "Device Group for $(get-3lettersInBrackets $thisGeoUG.displayName) - All"
                add-graphUsersToGroup -tokenResponse $userBotTokenResponse -graphGroupId $allGroup.id -memberType members -graphUserIds $subgroup.id
            }
            $allSubGroup = get-graphGroups -tokenResponse $userBotTokenResponse -filterDisplayName "Devices - All - $group" 
            add-graphUsersToGroup -tokenResponse $userBotTokenResponse -graphGroupId $allSubGroup.id -memberType members -graphUserIds $subGroup.id
    
        }
    }
    $thisGeoUsers = get-graphUsersFromGroup -tokenResponse $userBotTokenResponse -groupId $thisGeoUG.id -memberType TransitiveMembers -returnOnlyLicensedUsers #-selectAllProperties
    $thisGeoUsers | ForEach-Object {Add-Member -InputObject $_ -MemberType NoteProperty -Name OwnerId -Value $_.id}
    #Get the devices owned by people in this Geographic Group
    #$thisGeoDevices = Compare-Object -ReferenceObject @($allAadDevices | Select-Object) -DifferenceObject @($thisGeoUsers | Select-Object) -Property OwnerId -IncludeEqual -ExcludeDifferent -PassThru
    $thisGeoDevices = @()
    ForEach($thisGeoUser in $thisGeoUsers){
        $thisUserGeoDevices = $allAadDevices.Where({$_.OwnerId -eq $thisGeoUser.OwnerId})
        $thisGeoDevices += $thisUserGeoDevices
    }
    #Get the old list of devices owned by people in this Geographic Group
    $thisGeoDevicesCurrent = get-graphUsersFromGroup -tokenResponse $userBotTokenResponse -groupId $thisGeoUG.anthesisgroup_UGSync.deviceGroupId -memberType TransitiveMembers
    $thisGeoDevicesCurrentDevices = $thisGeoDevicesCurrent | Where-Object {$_.'@odata.type' -ne "#microsoft.graph.group" } | ForEach-Object {get-graphDevices -tokenResponse $userBotTokenResponse -filterCustomEq @{"Id" = $_.Id} -includeOwners}
    $thisGeoDevicesCurrentGroups = $thisGeoDevicesCurrent | Where-Object {$_.'@odata.type' -eq "#microsoft.graph.group" }
    #Compare the old list with the new one
    $thisGeoDevicesDelta = Compare-Object -ReferenceObject @($thisGeoDevices | Select-Object) -DifferenceObject @($thisGeoDevicesCurrentDevices | Select-Object)  -Property id -PassThru

    $thisGeoDevicesDelta | ForEach-Object {
        #Find the appropriate subgroup
        if($_.model -eq "Virtual Machine"){
            ForEach($physicalId in $_.physicalIds){
                If(($physicalId.split(":")[0] -eq "[AzureResourceId]") -and ($physicalId -match "AVD-Personal")){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {($_.displayName -match "VMs") -and ($_.displayName -match "Personal")}} #AVD personal VMs
            }
        }
        elseif($_.operatingSystem -eq "Windows" -and ($_.operatingSystemVersion -match "^10.0.22" -or $_.operatingSystemVersion -eq "Windows 11")){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "Win11"}}
        elseif($_.operatingSystem -eq "Windows" -and ($_.operatingSystemVersion -match "^10" -or $_.operatingSystemVersion -eq "Windows 10")){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "Win10"}}
        elseif($_.operatingSystem -eq "Windows"){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "Win_Other"}}
        elseif($_.operatingSystem -match "Mac"){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "MacOS"}}
        elseif($_.operatingSystem -match "Android"){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "Android"}}
        elseif($_.operatingSystem -eq "iPad" -or $_.operatingSystem -eq "iPhone" -or $_.operatingSystem -eq "iOS"){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "iOS"}}
        else{
            Write-Warning "Device [$($_.id)] owned by [$($_.registeredOwners.userPrincipalName -join "; ")] cannot be categorised, so cannot be included in any GeoDevice Groups for [$($thisGeoUG.displayName)]"
            $duffDevices += $_
            $relevantGroup = $null
            }


        #Add/Remove the delta device as appropriate
        if(($_.SideIndicator -eq "<=") -and ($relevantGroup -ne $null)){
            try{
                Write-Host "`tAdding device [$($_.displayName)][$($_.id)] owned by [$($_.registeredOwners.userPrincipalName -join "; ")] to [$($relevantGroup.displayName)]"
                add-graphUsersToGroup -tokenResponse $userBotTokenResponse -graphGroupId $relevantGroup.id -memberType members -graphUserIds $_.id
                }
            catch{if($_.Exception -notmatch "(400)"){get-errorSummary $_}}
            }
        elseif(($_.SideIndicator -eq "=>") -and ($relevantGroup -ne $null)){
            try{
                Write-Host "`tRemoving device [$($_.displayName)][$($_.id)] owned by [$($_.registeredOwners.userPrincipalName -join "; ")] from [$($relevantGroup.displayName)]"
                remove-graphUsersFromGroup -tokenResponse $userBotTokenResponse -graphGroupId $relevantGroup.id -memberType members -graphUserIds $_.id
                }
            catch{if($_.Exception -notmatch "(404)"){get-errorSummary $_}}
            }
        else{Write-Warning "SideIndicator = [$($_.SideIndicator)] for Device [$($_.id)] owned by [$($_.registeredOwners.userPrincipalName -join "; ")]. It cannot be added/removed from [$($relevantGroup.displayName)]"}
        }
    }

}
Else{
write-host "Error: GeoUGs or AAD Devices not retrieved"
}
Stop-Transcript



