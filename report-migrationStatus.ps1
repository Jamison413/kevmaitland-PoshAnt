$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$tokenResponseTeamBot = get-graphTokenResponse -aadAppCreds $teamBotDetails
$intuneBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\intunebot.txt"
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

$completed.userprincipalname | sort
$toMigrate.userprincipalname | sort

$completed.Count / ($completed.count + $toMigrate.Count)