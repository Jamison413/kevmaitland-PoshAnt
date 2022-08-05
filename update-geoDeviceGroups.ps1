start-transcriptLog -thisScriptName "update-geoDeviceGroups"

$userBotDetails = get-graphAppClientCredentials -appName TeamsBot
$userBotTokenResponse = get-graphTokenResponse -aadAppCreds $userBotDetails

Write-host "Getting Geographic Unified Groups"
$geoUGs = get-graphGroups -tokenResponse $userBotTokenResponse -filterGroupType Unified -filterDisplayNameStartsWith "All (" -selectAllProperties
$geoUGs  = $geoUGs | Where-Object {$_.displayName -ne "All (All)"}
Write-host "`tRetrieved [$($geoUGs.Count)] Geographic Unified Groups"


Write-host "Getting all AAD devices"
$allAadDevices = get-graphDevices -tokenResponse $userBotTokenResponse -includeOwners
$allAadDevices | ForEach-Object {Add-Member -InputObject $_ -MemberType NoteProperty -Name OwnerId -Value $_.registeredOwners[0].id}
Write-host "`tRetrieved [$($allAadDevices.Count)] AD devices"

#Remove baseline testing devices (Nov 2021 - testing Jan 2022)
$allAadDevices = $allAadDevices.Where({($_.deviceId -ne "e0212c77-12be-4d86-8acb-e8efcfbf1ee8") -and ($_.deviceId -ne "b8584707-0c3f-4a1a-9418-187937c955c3") -and ($_.deviceId -ne "378f8497-1433-417e-a524-94fee14e8002") -and ($_.deviceId -ne "1065721f-4f3a-4ae8-8541-ff87e0b9343a")})



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
        $subgroups = @("Win10","iOS","VMs","Android","MacOS")
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
    $thisGeoDevices = Compare-Object -ReferenceObject @($allAadDevices | Select-Object) -DifferenceObject @($thisGeoUsers | Select-Object) -Property OwnerId -IncludeEqual -ExcludeDifferent -PassThru
    #Get the old list of devices owned by people in this Geographic Group
    $thisGeoDevicesCurrent = get-graphUsersFromGroup -tokenResponse $userBotTokenResponse -groupId $thisGeoUG.anthesisgroup_UGSync.deviceGroupId -memberType Members 
    $thisGeoDevicesCurrentDevices = $thisGeoDevicesCurrent | Where-Object {$_.'@odata.type' -ne "#microsoft.graph.group" }
    $thisGeoDevicesCurrentGroups = $thisGeoDevicesCurrent | Where-Object {$_.'@odata.type' -eq "#microsoft.graph.group" }
    #Compare the old list with the new one
    $thisGeoDevicesDelta = Compare-Object -ReferenceObject @($thisGeoDevices | Select-Object) -DifferenceObject @($thisGeoDevicesCurrentDevices | Select-Object)  -Property id -PassThru

    $thisGeoDevicesDelta | ForEach-Object {
        #Find the appropriate subgroup
        if($_.model -eq "Virtual Machine"){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "VMs"}}
        elseif($_.operatingSystem -eq "Windows" -and ($_.operatingSystemVersion -match "^10" -or $_.operatingSystemVersion -eq "Windows 10")){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "Win10"}}
        elseif($_.operatingSystem -match "Mac"){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "MacOS"}}
        elseif($_.operatingSystem -match "Android"){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "Android"}}
        elseif($_.operatingSystem -eq "iPad" -or $_.operatingSystem -eq "iPhone" -or $_.operatingSystem -eq "iOS"){$relevantGroup = $thisGeoDevicesCurrentGroups | Where-Object {$_.displayName -match "iOS"}}
        else{
            Write-Warning "Device [$($_.id)] owned by [$($_.registeredOwners.userPrincipalName -join "; ")] cannot be categorised, so cannot be included in any GeoDevice Groups for [$($thisGeoUG.displayName)]"
            $duffDevices += $_
            $relevantGroup = $null
            }
        
        #Add/Remove the delta device as appropriate
        if($_.SideIndicator -eq "<="){
            try{
                Write-Host "`tAdding device [$($_.displayName)][$($_.id)] owned by [$($_.registeredOwners.userPrincipalName -join "; ")] to [$($relevantGroup.displayName)]"
                add-graphUsersToGroup -tokenResponse $userBotTokenResponse -graphGroupId $relevantGroup.id -memberType members -graphUserIds $_.id
                }
            catch{if($_.Exception -notmatch "(400)"){get-errorSummary $_}}
            }
        elseif($_.SideIndicator -eq "=>"){
            try{
                Write-Host "`tRemoving device [$($_.displayName)][$($_.id)] owned by [$($_.registeredOwners.userPrincipalName -join "; ")] from [$($relevantGroup.displayName)]"
                remove-graphUsersFromGroup -tokenResponse $userBotTokenResponse -graphGroupId $relevantGroup.id -memberType members -graphUserIds $_.id
                }
            catch{get-errorSummary $_}
            }
        else{Write-Warning "SideIndicator = [$($_.SideIndicator)] for Device [$($_.id)] owned by [$($_.registeredOwners.userPrincipalName -join "; ")]. It cannot be added/removed from [$($relevantGroup.displayName)]"}
        }
    }

Stop-Transcript