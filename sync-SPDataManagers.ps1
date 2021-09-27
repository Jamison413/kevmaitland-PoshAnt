$Logname = "C:\ScriptLogs" + "\sync-SPDataManagers $(Get-Date -Format "yyMMdd").log"
Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)



#For pnp (Graph can't manage Sharepoint groups currently)
$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Downloads\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass


$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

#Check connection
If(!($tokenResponse.access_token)){
write-host "Error getting TeamsBot Credentials, exiting." -ForegroundColor Red
Exit
}


#Get members of 'Data Managers - Authorised (All) from 365' and sp groups from the team hub and client hub
$datamanagers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId "daf56fbd-ebce-457e-a10a-4fce50a2f99c" -memberType "Members"

#Get members of clients (unrestricted) - Modify
$clientsModifyMembers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId "5084cf5a-c05e-4f10-b2af-a74417beceee" -memberType "Members"

#Do any cleaning as a quick bodge until NS automation
$leavers = $clientsModifyMembers | Where-Object -Property "displayName" -Match "Ω_"
If($leavers){
remove-graphUsersFromGroup -tokenResponse $tokenResponse -graphGroupId "5084cf5a-c05e-4f10-b2af-a74417beceee" -memberType Members -graphUserUpns $leavers.userPrincipalName -Verbose
$clientsModifyMembers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId "5084cf5a-c05e-4f10-b2af-a74417beceee" -memberType "Members"
}
Else{
Write-Host "Clients (unrestircted) - Modify is already clean ('ish)" -ForegroundColor Yellow
}




#############################
#                           #
#          Team Hub         #
#                           #
#############################

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/sites/TeamHub/" -Credentials $adminCreds

$pnpconnection = Get-PnPConnection

#Check connection
If(!($pnpConnection)){
write-host "Error getting pnp connection, exiting." -ForegroundColor Red
Exit
}

$DataManagerSPOGroupName = "Internal - SPO Authorised Data Managers"
$MembersSPOGroupName = "Internal - SPO Authorised TeamHub Members"

$internalcurrentspdatamanagers = Get-PnPGroupMembers -Identity $DataManagerSPOGroupName
$internalcurrentspmembers = Get-PnPGroupMembers -Identity $MembersSPOGroupName

#We just compare and add/remove any new members

#Add managers
$newdatamanagers = Compare-Object -ReferenceObject $datamanagers.mail -DifferenceObject $internalcurrentspdatamanagers.Email | where-object -Property "SideIndicator" -EQ "<="
ForEach($newdatamanager in $newdatamanagers){
Write-Host "New Data Manager: Adding $($newdatamanager.InputObject) to $($DataManagerSPOGroupName)" -ForegroundColor Yellow
$spdatamanagers = Add-PnPUserToGroup -LoginName $($newdatamanager.InputObject) -Identity $DataManagerSPOGroupName
}

#Remove managers
$removeddatamanagers = Compare-Object -ReferenceObject $datamanagers.mail -DifferenceObject $internalcurrentspdatamanagers.Email | where-object -Property "SideIndicator" -EQ "=>"
$removeddatamanagers = $removeddatamanagers | Where-Object -property "inputObject" -ne "T1-Emily.Pressey@anthesisgroup.com" #this account is the Group Owner (can't add domain groups)
ForEach($removeddatamanager in $removeddatamanagers){
Write-Host "Removed Data Manager: Removing $($removeddatamanager.InputObject) from $($DataManagerSPOGroupName)" -ForegroundColor Yellow
$spdatamanagers = Remove-PnPUserFromGroup -LoginName $($removeddatamanager.InputObject) -Identity $DataManagerSPOGroupName
}

#Add members
$newmembers = Compare-Object -ReferenceObject $clientsModifyMembers.mail -DifferenceObject $internalcurrentspmembers.Email | where-object -Property "SideIndicator" -EQ "<="
ForEach($newmember in $newmembers){
Write-Host "New Member: Adding $($newmember.InputObject) to $($MembersSPOGroupName)" -ForegroundColor Yellow
$spmembers = Add-PnPUserToGroup -LoginName $($newmember.InputObject) -Identity $MembersSPOGroupName
}

#Remove members
$removedmembers = Compare-Object -ReferenceObject $clientsModifyMembers.mail -DifferenceObject $internalcurrentspmembers.Email | where-object -Property "SideIndicator" -EQ "=>"
$removedmembers = $removedmembers | Where-Object -property "inputObject" -ne "T1-Emily.Pressey@anthesisgroup.com" #this account is the Group Owner (can't add domain groups)
ForEach($removedmember in $removedmembers){
Write-Host "Removed Member: Removing $($removedmember.InputObject) from $($MembersSPOGroupName)" -ForegroundColor Yellow
$spmembers = Remove-PnPUserFromGroup  -LoginName $($removedmember.InputObject) -Identity $MembersSPOGroupName
}






#############################
#                           #
#        Client Hub         #
#                           #
#############################

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/clients/" -Credentials $adminCreds

$pnpconnection = Get-PnPConnection

#Check connection
If(!($pnpConnection)){
write-host "Error getting pnp connection, exiting." -ForegroundColor Red
Exit
}


$DataManagerSPOGroupName = "External - SPO Authorised Data Managers"
$MembersSPOGroupName = "External - Authorised Client Members"

$externalcurrentspdatamanagers = Get-PnPGroupMembers -Identity $DataManagerSPOGroupName
$externalcurrentspmembers = Get-PnPGroupMembers -Identity $MembersSPOGroupName

#We just compare and add/remove any new members to the two Sharepoint groups - external

#Add Managers
$newdatamanagers = Compare-Object -ReferenceObject $datamanagers.mail -DifferenceObject $externalcurrentspdatamanagers.Email | where-object -Property "SideIndicator" -EQ "<="
ForEach($newdatamanager in $newdatamanagers){
Write-Host "New Data Manager: Adding $($newdatamanager.InputObject) to $($DataManagerSPOGroupName)" -ForegroundColor Yellow
$spdatamanagers = Add-PnPUserToGroup -LoginName $($newdatamanager.InputObject) -Identity $DataManagerSPOGroupName
}

#Remove Managers
$removeddatamanagers = Compare-Object -ReferenceObject $datamanagers.mail -DifferenceObject $externalcurrentspdatamanagers.Email | where-object -Property "SideIndicator" -EQ "=>"
$removeddatamanagers = $removeddatamanagers | Where-Object -property "inputObject" -ne "T1-Emily.Pressey@anthesisgroup.com" #this account is the Group Owner (can't add domain groups)
ForEach($removeddatamanager in $removeddatamanagers){
Write-Host "Removed Data Manager: Removing $($removeddatamanager.InputObject) from SPDataManagers" -ForegroundColor Yellow
$spdatamanagers = Remove-PnPUserFromGroup -LoginName $($removeddatamanager.InputObject) -Identity $DataManagerSPOGroupName
}


#Add members
$newmembers = Compare-Object -ReferenceObject $clientsModifyMembers.mail -DifferenceObject $externalcurrentspmembers.Email | where-object -Property "SideIndicator" -EQ "<="
ForEach($newmember in $newmembers){
Write-Host "New Member: Adding $($newmember.InputObject) to $($MembersSPOGroupName)" -ForegroundColor Yellow
$spmembers = Add-PnPUserToGroup -LoginName $($newmember.InputObject) -Identity $MembersSPOGroupName
}

#Remove members
$removedmembers = Compare-Object -ReferenceObject $clientsModifyMembers.mail -DifferenceObject $externalcurrentspmembers.Email | where-object -Property "SideIndicator" -EQ "=>"
$removedmembers = $removedmembers | Where-Object -property "inputObject" -ne "T1-Emily.Pressey@anthesisgroup.com" #this account is the Group Owner (can't add domain groups)
ForEach($removedmember in $removedmembers){
Write-Host "Removed Member: Removing $($removedmember.InputObject) from $($MembersSPOGroupName)" -ForegroundColor Yellow
$spmembers = Remove-PnPUserFromGroup -LoginName $($removedmember.InputObject) -Identity $MembersSPOGroupName
}

#Add members
$newmembers = Compare-Object -ReferenceObject $clientsModifyMembers.mail -DifferenceObject $externalcurrentspmembers.Email | where-object -Property "SideIndicator" -EQ "<="
ForEach($newmember in $newmembers){
Write-Host "New Member: Adding $($newmember.InputObject) to $($MembersSPOGroupName)" -ForegroundColor Yellow
$spmembers = Add-PnPUserToGroup -LoginName $($newmember.InputObject) -Identity $MembersSPOGroupName
}

#Reports
If(!($error)){
$status = "Ok"
}
Else{
$status = "Error"
}
$syncSPDataManagersHash = @{
"reportType" = "sync-SPDataManagers";
"Status" = "$($status)"
"Notes" = "$($error)"
"LastRun" = "$(get-date)"
}
Update-graphListItem -tokenResponse $tokenResponse -serverRelativeSiteUrl "https://anthesisllc.sharepoint.com/teams/IT_Team_All_365/" -listName "IT reporting" -listitemId 14  -fieldHash $syncSPDataManagersHash 
#Remove members
$removedmembers = Compare-Object -ReferenceObject $clientsModifyMembers.mail -DifferenceObject $externalcurrentspmembers.Email | where-object -Property "SideIndicator" -EQ "=>"
$removedmembers = $removedmembers | Where-Object -property "inputObject" -ne "T1-Emily.Pressey@anthesisgroup.com" #this account is the Group Owner (can't add domain groups)
ForEach($removedmember in $removedmembers){
Write-Host "Removed Member: Removing $($removedmember.InputObject) from $($MembersSPOGroupName)" -ForegroundColor Yellow
$spmembers = Remove-PnPUserFromGroup -LoginName $($removedmember.InputObject) -Identity $MembersSPOGroupName
}			   

Stop-Transcript