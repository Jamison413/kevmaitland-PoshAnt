$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$teamBotTokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

connect-ToExo

#Set what we want to change
$displayNameString = "Diversity, Equity & Inclusion (GBR)‎"
$newdisplayNameString = "Diversity, Equity & Inclusion Team (GBR)"


#//GROUPS//

#Find all the stuff - use the unfied group as the central point of information, and get some more information from the old-style entrypoint for the groups (where sharepointurl is an available property without too much faff)
$graphgroup = get-graphGroupWithUGSyncExtensions -tokenResponse $teamBotTokenResponse -filterDisplayName $displayNameString
$unifiedgroup = Get-UnifiedGroup -Identity $graphgroup.id -IncludeAllProperties -verbose

$currentcombinedgroup = get-graphGroups -tokenResponse $teamBotTokenResponse -filterId $graphgroup.anthesisgroup_UGSync.combinedGroupId
$currentDataManagersgroup = get-graphGroups -tokenResponse $teamBotTokenResponse -filterId $graphgroup.anthesisgroup_UGSync.dataManagerGroupId
$currentMembersgroup = get-graphGroups -tokenResponse $teamBotTokenResponse -filterId $graphgroup.anthesisgroup_UGSync.memberGroupId
If($graphgroup.anthesisgroup_UGSync.sharedMailboxId){$currentSharedmailbox = Get-Mailbox -Identity $graphgroup.anthesisgroup_UGSync.sharedMailboxId} 

#Sanitise displayname for emails
$displayNameEmailFormat = $newdisplayNameString.Replace("North America","NA")
$displayNameEmailFormat = $displayNameEmailFormat -replace "\(*\)"
$displayNameEmailFormat = $displayNameEmailFormat -replace "\)*\("
$displayNameEmailFormat = $displayNameEmailFormat -replace "& "
$displayNameEmailFormat = $displayNameEmailFormat -replace ","
$displayNameEmailFormat = $displayNameEmailFormat.Replace(" ","_")


$365email = $displayNameEmailFormat + "_365" + "@anthesisgroup.com"
$combinedemail = $displayNameEmailFormat + "@anthesisgroup.com"
$datamanagersgroupemail = $displayNameEmailFormat + "_-_Data_Managers_Subgroup" + "@anthesisgroup.com"
$membersgroupemail = $displayNameEmailFormat + "_-_Members_Subgroup" + "@anthesisgroup.com"
$sharedmailboxemail = "Shared_Mailbox_-_" + "$displayNameEmailFormat" + "@anthesisgroup.com"

#Suggest changes and ask for sign off

Write-Host "---------------------------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "|                                           Table of Changes - Hit Y to Apply                                                   |" -ForegroundColor Yellow
write-host "---------------------------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "Sharepoint Site Name      :   $($graphgroup.displayName)  ->  $newdisplayNameString" -ForegroundColor Yellow
Write-Host "365 group name            :   $($graphgroup.displayName)  ->  $newdisplayNameString" -ForegroundColor Yellow
Write-Host "365 group email           :   $($graphgroup.mail)  ->  $365email" -ForegroundColor DarkGray
Write-Host "combined group name       :   $($currentcombinedgroup.displayName)  ->  $newdisplayNameString" -ForegroundColor Yellow
Write-Host "combined group email      :   $($currentcombinedgroup.mail)  ->  $combinedemail" -ForegroundColor DarkGray
Write-Host "Managers group name       :   $($currentDataManagersgroup.displayName)  ->  $($newdisplayNameString + " - Data Managers Subgroup")" -ForegroundColor Yellow
Write-Host "Managers group email      :   $($currentDataManagersgroup.mail)  ->  $datamanagersgroupemail" -ForegroundColor DarkGray
Write-Host "Members group name        :   $($currentMembersgroup.displayName)  ->  $($newdisplayNameString + " - Members Subgroup")" -ForegroundColor Yellow
Write-Host "Members group email       :   $($currentMembersgroup.mail)  ->  $membersgroupemail" -ForegroundColor DarkGray
If($currentSharedmailbox){
Write-Host "Shared Mailbox name       :   $($currentSharedmailbox.DisplayName)  ->  $("Shared Mailbox - " + $newdisplayNameString)" -ForegroundColor Yellow
Write-Host "Shared Mailbox email      :   $($currentSharedmailbox.PrimarySmtpAddress)  ->  $sharedmailboxemail" -ForegroundColor DarkGray
}
Else{Write-Host "No Shared Mailbox ID found on the Unified Group Extension Data" -ForegroundColor Cyan}
Write-Host "---------------------------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "|                                                                                                                               |" -ForegroundColor Yellow
write-host "---------------------------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow


$ans = Read-Host -Prompt "(y/n?)"
If($ans -eq "y"){

#Change all the stuff

#Change all display names (this will affect Teams also for the 365 rename)

[hashtable]$365group = @{"displayName" = $newdisplayNameString}
invoke-graphPatch -tokenResponse $teamBotTokenResponse -graphQuery "/groups/$($graphgroup.id)" -graphBodyHashtable $365group -Verbose
Set-DistributionGroup -Identity $graphgroup.anthesisgroup_UGSync.combinedGroupId -DisplayName $newdisplayNameString
Set-DistributionGroup -Identity $graphgroup.anthesisgroup_UGSync.dataManagerGroupId -DisplayName $($newdisplayNameString + " - Data Managers Subgroup")
Set-DistributionGroup -Identity $graphgroup.anthesisgroup_UGSync.memberGroupId -DisplayName $($newdisplayNameString + " - Members Subgroup")
If($currentSharedmailbox){Set-Mailbox -Identity $graphgroup.anthesisgroup_UGSync.sharedMailboxId -DisplayName $("Shared Mailbox - " + $newdisplayNameString)}


#Change all email addresses

Set-Group -Identity $graphgroup.id -WindowsEmailAddress $365email
Set-DistributionGroup -Identity $graphgroup.anthesisgroup_UGSync.combinedGroupId -PrimarySmtpAddress $combinedemail
Set-DistributionGroup -Identity $graphgroup.anthesisgroup_UGSync.dataManagerGroupId -PrimarySmtpAddress $datamanagersgroupemail
Set-DistributionGroup -Identity $graphgroup.anthesisgroup_UGSync.memberGroupId -PrimarySmtpAddress $membersgroupemail
If($currentSharedmailbox){Set-Mailbox -Identity $graphgroup.anthesisgroup_UGSync.sharedMailboxId -EmailAddresses "SMTP:$($sharedmailboxemail)"}

#Set group descriptions
Set-Group -Identity $graphgroup.id -Notes "Unified 365 Group for $($newdisplayNameString)"
Set-Group -Identity $graphgroup.anthesisgroup_UGSync.combinedGroupId -Notes "Mail-enabled Security Group for $($newdisplayNameString)"
Set-Group -Identity $graphgroup.anthesisgroup_UGSync.dataManagerGroupId -Notes "Mail-enabled Security Group for $($newdisplayNameString) Data Managers"
Set-Group -Identity $graphgroup.anthesisgroup_UGSync.memberGroupId -Notes "Mail-enabled Security Group for mirroring membership of $($newdisplayNameString) Unified Group"
}
Else{
Write-Host "Okay we won't change anything, feel free to manually amend the variables to change to what it needed" -ForegroundColor Red
}


#//SITES//


#We can't amend Sharepoint site names easily - you'll have to do it through the GUI


#Resource Site - again manual amend
Connect-PnPOnline -Url "https://anthesisllc-admin.sharepoint.com/" -UseWebLogin
Get-PnPTenantSite -WebTemplate SITEPAGEPUBLISHING#0

#You'll also have to change it manually on the Team Hub page - to save you a few clicks: https://anthesisllc.sharepoint.com/sites/TeamHub