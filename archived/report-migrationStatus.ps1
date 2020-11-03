$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\Desktop\teambotdetails.txt"
$tokenResponseTeamBot = get-graphTokenResponse -aadAppCreds $teamBotDetails
$intuneBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\Desktop\intunebotdetails.txt"
$tokenResponseIntuneBot = get-graphTokenResponse -aadAppCreds $intuneBotDetails


#$allGBR = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterDisplayName "All (GBR)"
#$allGBR.id
$allUKUsers = get-graphUsersFromGroup -tokenResponse $tokenResponseTeamBot -groupId 549dd0d0-251f-4c23-893e-9d0c31c2dc13 -memberType TransitiveMembers -returnOnlyLicensedUsers

$intuneDevices = invoke-graphGet -tokenResponse $tokenResponseIntuneBot -graphQuery "/deviceManagement/managedDevices" -Verbose
$windowsDevices = $intuneDevices | ? {$_.operatingsystem -eq "Windows"}

$compare = Compare-Object -ReferenceObject $allUKUsers -DifferenceObject $windowsDevices -Property userPrincipalName -IncludeEqual -PassThru
$completed = $compare | ? {$_.SideIndicator -eq "=="}
$toMigrate = $compare | ? {$_.SideIndicator -eq "<="}
$weirdos = $compare | ? {$_.SideIndicator -eq "=>"}

Write-Host -ForegroundColor Yellow "Migrated"
Write-Host -ForegroundColor Yellow "--------"
$completed.userprincipalname | sort
Write-Host -ForegroundColor Yellow "ToDo"
Write-Host -ForegroundColor Yellow "----"
$toMigrate.userprincipalname | sort

Write-Host -ForegroundColor Yellow "% Complete"
Write-Host -ForegroundColor Yellow "----------"
"$([System.Math]::Floor(($completed.Count / ($completed.count + $toMigrate.Count) *100)))%"