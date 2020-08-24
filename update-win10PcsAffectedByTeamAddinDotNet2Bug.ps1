#Find members of [Temp - Users affected by Team Addin .NET 2.0 bug] and move their devices to the remidiation group [Temp - Win10 PCs affected by Team Addin .NET 2.0 bug]
$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$teamBotTokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$affectedUsers = "0d437b9e-64e0-415d-9445-940cba619e31"
$affectedDevices = "0c8d3b2e-7ff3-4c05-84cd-8a38332f3a98"

update-graphGroupOfDevicesBasedOnOwners -tokenResponse $teamBotTokenResponse -userGroupId $affectedUsers -devicesGroupId $affectedDevices -deviceType Windows -Verbose