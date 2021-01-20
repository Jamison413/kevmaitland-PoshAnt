$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$teamBotTokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$tescoUsersGroupId = "3c48c759-8ce8-4eab-9dc3-984305593446"
$tescoDevicesGroupId = "35847d50-69c2-4b76-8737-5942954754c4"

$tescoUsers = get-graphUsersFromGroup -tokenResponse $teamBotTokenResponse -groupId $tescoUsersGroupId -memberType TransitiveMembers -returnOnlyLicensedUsers
$tescoWindowsDevices = get-graphDevices -tokenResponse $teamBotTokenResponse -filterOwnerIds $tescoUsers.id -filterOperatingSystem Windows
$currentTescoDevices = get-graphUsersFromGroup -tokenResponse $teamBotTokenResponse -groupId $tescoDevicesGroupId -memberType TransitiveMembers
if([string]::IsNullOrWhiteSpace($tescoWindowsDevices.Id)){$tescoWindowsDevices = @()}
if([string]::IsNullOrWhiteSpace($currentTescoDevices.Id)){$currentTescoDevices = @()}

$delta = Compare-Object -ReferenceObject $currentTescoDevices -DifferenceObject $tescoWindowsDevices -Property Id -PassThru

#Remove personal devices with workplace trust replationship
$delta = $delta.Where({$_.trustType -ne "Workplace"})

$toAdd = $delta | ? {$_.SideIndicator -eq "=>"}
If($toAdd){add-graphUsersToGroup -tokenResponse $teamBotTokenResponse -graphGroupId $tescoDevicesGroupId -memberType Members -graphUserIds $toAdd.id -Verbose}

$toRemove = $delta | ? {$_.SideIndicator -eq "<="}
If($toRemove){remove-graphUsersFromGroup -tokenResponse $teamBotTokenResponse -graphGroupId $tescoDevicesGroupId -memberType Members -graphUserIds $toRemove.id -Verbose}
