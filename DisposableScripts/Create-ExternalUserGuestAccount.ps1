
#add list of upns
$listofupns = ""

#connect
$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$formattedUpnsToInvite = convertTo-arrayOfEmailAddresses -blockOfText $listofupns

#process
ForEach($upn in $formattedUpnsToInvite){

#check if email exists
$emailFound = get-graphUsers -tokenResponse $tokenResponse -filterCustomEq @{"mail" = $upn} -Verbose}

#create guest account
$newGuestAccount = new-GraphGuestInvitation -tokenResponse $tokenResponse -invitedUserEmailAddress $upn -inviteRedirectUrl "https://myapps.microsoft.com/anthesisgroup.com" -sendInvitationMessage $false -Verbose
