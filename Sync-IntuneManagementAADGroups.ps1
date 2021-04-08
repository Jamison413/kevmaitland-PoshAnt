######################################################################################



                        #Windows Update Rings AAD Group#



######################################################################################




#Connect with IntuneBot initially to get device data from Intune
$IntuneBotDetails = get-graphAppClientCredentials -appName IntuneBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $IntuneBotDetails
If(!($tokenResponse.access_token)){
$TeamsReport = @{"Devices not Synced between Intune and AAD Group - IntuneBotConnection" = "Problem getting Graph Credentials"}
Exit
}
#Get all managed devices
$allDevices = invoke-graphGet -tokenResponse $tokenResponse -graphQuery "/deviceManagement/managedDevices" -useBetaEndPoint

#Re-connect with TeamsBot to get the AAD group 
$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails
If(!($tokenResponse.access_token)){
$TeamsReport = @{"Devices not Synced between Intune and AAD Group - TeamsBotConnection" = "Problem getting Graph Credentials"}
Exit
}
#Get all AAD devices (we can't get the AAD directory object id via Intune managed device object
$allAADDevices = get-graphDevices -tokenResponse $tokenResponse -filterOperatingSystem Windows 


################################
#                              #
#                              #
#           Filters            #
#                              #
#                              #
################################ 

#Win10 devices only
$IntuneDevices = $allDevices | ? {$_.operatingSystem -eq "Windows"}
$IntuneDevices = $IntuneDevices | ? {(($_.osVersion).split(".")[0]) -eq "10"}
#AzureAD joined devices only:  UserEnrollment = jointype of "azureADRegistered"  /  azureADJoined = jointype of "azureADJoined"
$IntuneDevices = $IntuneDevices | ? {$_.deviceEnrollmentType -eq "windowsAzureADJoin"}
#Ownership type of Company devices only
$IntuneDevices = $IntuneDevices | ? {$_.ownerType -eq "company"}
#Omit Virtual Machines
$IntuneDevices = $IntuneDevices | ? {$_.model -ne "Virtual Machine"}

#Iterate through each managed device and find the associated AAD object ID (we'll need this to add it to the group, we can't pull it back from the Managed Device endpoint in Graph)
$targetDevices = @()
ForEach($IntuneDevice in $IntuneDevices){
$thisIntunedevice = $IntuneDevice

$thisAADdevice = ""
$thisAADdevice = $allAADDevices | ? {$_.deviceId -eq $thisIntunedevice.azureADDeviceId}
    If(($thisAADdevice | Measure-Object).Count -eq 1){
    $thisIntunedevice | Add-Member -MemberType NoteProperty -Name FoundAADDirectoryObjectID -Value "$($thisAADdevice.id)" -Force
    $targetDevices += $thisIntunedevice
    }
    Else{
    Write-Host "Couldn't pin down an AAD object for this device: $($thisIntunedevice.deviceName)" -ForegroundColor Red
    }

}
Write-Host "We found $($targetDevices.Count) target devices" -ForegroundColor Yellow


################################
#                              #
#                              #
#         Group Sync           #
#                              #
#                              #
################################
 
Write-Host "Syncing devices to AAD group the Windows Update and Feature Rings are assigned with" -ForegroundColor Cyan

#Get all devices from the MDM - Windows Update Ring - All Intune Devices (Production) AAD group - which is assigned to the update an feature rings
$targetAADGroupId = "2dff2b3b-d432-4de2-8657-fa6ca58a9702"
$updateRingGroup = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $targetAADGroupId -memberType Members
Write-Host "Current count of devices in AAD group: $(($updateRingGroup | Measure-Object).count)" -ForegroundColor Yellow
#Compare by the AAD object ID we found earlier, if there is a mismatch, sync the device
$unSyncedDevices = Compare-Object -ReferenceObject $updateRingGroup.id -DifferenceObject $targetDevices.FoundAADDirectoryObjectID
$unSyncedDevices = $unSyncedDevices | ? {$_.SideIndicator -eq "=>"}
If($unSyncedDevices){
Write-Host "[We found $($unSynced.Count) devices not in our AAD group...syncing back up]" -ForegroundColor Cyan
    ForEach($unSyncedDevice in $unSyncedDevices){
    $thisUnsyncedDevice = $targetDevices | ? {$_.FoundAADDirectoryObjectID -eq $unSyncedDevice.InputObject}
    Write-Host "Adding device to ADD group: $($thisUnsyncedDevice.deviceName)..." -ForegroundColor Cyan
    invoke-graphPost -tokenResponse $tokenResponse -graphQuery "/groups/$($targetAADGroupId)/members/`$ref" -graphBodyHashtable @{"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($thisUnsyncedDevice.FoundAADDirectoryObjectID)"} -Verbose
}
}


################################
#                              #
#                              #
#         Reporting            #
#                              #
#                              #
################################ 

#Pull back the devices again to check the sync and report any mismatches to Teams
$updateRingGroup = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $targetAADGroupId -memberType Members
$unSyncedDevices = Compare-Object -ReferenceObject $updateRingGroup.id -DifferenceObject $targetDevices.FoundAADDirectoryObjectID
If($unSyncedDevices){$TeamsReport = @{"Devices not Synced between Intune and AAD Group for Windows Update Rings" = "$($unSyncedDevices.Count)"}}
If($TeamsReport){
    $report = @()
    $report += "***************Errors found in Sync-IntuneManagementAADGroups***************" + "<br>"
        ForEach($t in $TeamsReport){
        $report += "$($t.Keys)" + " - " + "$($t.Values)" + "<br><br>"
}
$report = $report | out-string
Send-MailMessage -To "8ed81bd4.anthesisgroup.com@amer.teams.ms" -From "groupbot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "test" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds
}



######################################################################################



                        #Intune Security Baseline AAD Group#



######################################################################################


#Clear session history as we want to stick to one automated script
Remove-Variable * -ErrorAction SilentlyContinue


#Kimblebot for Teams email-to-channel reporting
$Admin = "kimblebot@anthesisgroup.com"
$AdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass
$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass
connect-ToExo -credential $exoCreds

$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails
If(!($tokenResponse.access_token)){
$TeamsReport = @{"Users not Synced between Intune and AAD Group - TeamsBotConnection" = "Problem getting Graph Credentials"}
Exit
}


#using all gbr - we don't know if usage location is correct due to e1 licensed subcontractors
$allgbrgroup = "549dd0d0-251f-4c23-893e-9d0c31c2dc13"
$targetgroup = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $allgbrgroup -memberType Members -returnOnlyLicensedUsers

#Get membership from current Intune baseline AAD group
$IntuneBaselineAADgroup = "5e9fee91-b777-4b95-ac2e-b167676a48cd"
$IntuneBaselineAADgroupMembers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $IntuneBaselineAADgroup -memberType Members

#Remove anyone who isn't GB by usage location
$targetusers = @()

ForEach($person in $targetgroup){

$usagelocation = ""
$usagelocation = get-graphUsers -tokenResponse $tokenResponse -filterUpns $person.userprincipalname -selectAllProperties

$targetusers += $usagelocation
}
$GBRusers = $targetusers | ? {$_.usageLocation -eq "GB"}


$unSyncedUsers = Compare-Object -ReferenceObject $IntuneBaselineAADgroupMembers.userPrincipalName -DifferenceObject $GBRusers.userPrincipalName
If($unSyncedUsers){
    ForEach($unSyncedUser in $unSyncedUsers){

    If($unSyncedUser.SideIndicator -eq "<="){
    #remove
    remove-graphUsersFromGroup -tokenResponse $tokenResponse -graphGroupId $IntuneBaselineAADgroup -memberType Members -graphUserUpns $unSyncedUser.InputObject -Verbose
    }
    If($unSyncedUser.SideIndicator -eq "=>"){
    #add
    add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $IntuneBaselineAADgroup -memberType Members -graphUserIds $unSyncedUser.InputObject -Verbose
    }
    }
}


#Pull back the users again to check the sync and report any mismatches to Teams
$IntuneBaselineAADgroupMembers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $IntuneBaselineAADgroup -memberType Members
$unSyncedUsers = Compare-Object -ReferenceObject $IntuneBaselineAADgroupMembers.userPrincipalName -DifferenceObject $GBRusers.userPrincipalName
$unSyncedUsers = $unSyncedUsers | ? {($_.InputObject -ne "t0-kevin.maitland@anthesisgroup.com") -and ($_.InputObject -ne "t1-andrew.ost@anthesisgroup.com") -and ($_.InputObject -ne "t1-Emily.Pressey@anthesisgroup.com")}
If($unSyncedUsers){$TeamsReport = @{"Users not Synced between Intune and AAD Group for security baselines" = "$($unSyncedUsers.Count)"}}
If($TeamsReport){
    $report = @()
    $report += "***************Errors found in Sync-IntuneManagementAADGroups***************" + "<br>"
        ForEach($t in $TeamsReport){
        $report += "$($t.Keys)" + " - " + "$($t.Values)" + "<br><br>"
}
}
$report = $report | out-string
Send-MailMessage -To "8ed81bd4.anthesisgroup.com@amer.teams.ms" -From "groupbot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "test" -BodyAsHtml $report -Encoding UTF8 -Credential $exocreds

