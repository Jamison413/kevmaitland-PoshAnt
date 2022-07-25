#Creates Intune configuration profile for URI and a security group to assign user and device to


$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$teamBotTokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$IntuneBottokenResponse = get-graphTokenResponse -aadAppCreds $(get-graphAppClientCredentials -appName IntuneBot) -grant_type client_credentials 

#add user upn here to add back as admin:
$userUPN = ""

#add device name here to add user as admin to:
$targetDevice = ""

new-mdmLocalAdminPolicy -tokenResponseTeams $teamBotTokenResponse -tokenResponseIntune $IntuneBottokenResponse -userUPN $userUPN -deviceName $targetDevice -Verbose -overrideOtherPolicies -removeOtherMembers
