$Logname = "C:\Scripts" + "\Logs" + "\sync-SPDataManagers $(Get-Date -Format "yyMMdd").log" #Check this location before live
Start-Transcript -Path $Logname -Append
Write-Host "Script started:" (Get-date)

Import-Module _PNP_Library_SPO
Import-Module _PS_Library_Graph.psm1

#For pnp (Graph can't manage Sharepoint groups currently)
$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/sites/TeamHub/" -Credentials $adminCreds

#For Graph
$teamBotDetails = Import-Csv "$env:USERPROFILE\Desktop\teambotdetails.txt"
$resource = "https://graph.microsoft.com"
$tenantId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.TenantId)
$clientId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.ClientID)
$redirect = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Redirect)
$secret   = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Secret)

$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $secret
    } 
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

#Get members of 'Data Managers - Authorised (All) from 365' and sp group 'SPDataManagers'
$datamanagers = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId "daf56fbd-ebce-457e-a10a-4fce50a2f99c" -memberType "Members"
$currentspdatamanagers = Get-PnPGroupMembers -Identity "SPDataManagers"

#We just compare and add/remove any new members

#Add members
$newdatamanagers = Compare-Object -ReferenceObject $datamanagers.mail -DifferenceObject $currentspdatamanagers.Email | where-object -Property "SideIndicator" -EQ "<="
ForEach($newdatamanager in $newdatamanagers){
Write-Host "New Data Manager: Adding $($newdatamanager.InputObject) to SPDataManagers" -ForegroundColor Yellow
$spdatamanagers = Add-PnPUserToGroup -LoginName $($newdatamanager.InputObject) -Identity "SPDataManagers"
}

#Remove members
$removeddatamanagers = Compare-Object -ReferenceObject $datamanagers.mail -DifferenceObject $currentspdatamanagers.Email | where-object -Property "SideIndicator" -EQ "=>"
ForEach($removeddatamanager in $removeddatamanagers){
Write-Host "Removed Data Manager: Removing $($removeddatamanager.InputObject) from SPDataManagers" -ForegroundColor Yellow
$spdatamanagers = Remove-PnPUserFromGroup -LoginName $($removeddatamanager.InputObject) -Identity "SPDataManagers"
}



<#Add each member to the sp group SPDataManagers (saved for rebuilding the list if there is an issue, might save a bit of effort)
ForEach($datamanager in $datamanagers){
Write-Host "Adding $($datamanager.mail) to SPDataManagers" -ForegroundColor Yellow
$spdatamanagers = Add-PnPUserToGroup -LoginName "$($datamanager.mail)" -Identity "SPDataManagers"
}
#>

Stop-Transcript