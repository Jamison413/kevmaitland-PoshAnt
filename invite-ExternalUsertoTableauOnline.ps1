
#add list of upns
$listofupns = "
"

#connect
$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

$formattedUpnsToInvite = convertTo-arrayOfEmailAddresses -blockOfText $listofupns
$tableauOnlineViewersGroup = "5ab7fba6-01c4-4f02-9b5c-123f7c98e752" 

#process

ForEach($upn in $formattedUpnsToInvite){

#check if email exists
$emailFound = get-graphUsers -tokenResponse $tokenResponse -filterCustomEq @{"mail" = $upn} -Verbose

If($emailFound){
Write-Host "External user email already in AAD as Guest, adding to Tableau Viewers security group and emailing them SSO sign in url..." -ForegroundColor Green
#email exists in AAD as Guest already, add to Tableau Online Viewers group
            Try{
                add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $tableauOnlineViewersGroup -memberType members -graphUserIds $emailFound.id -Verbose
            }
            Catch{
                $error[0]
            }

}
Else{
Write-Host "Inviting external user $($upn) to AAD as Guest to access Tableau Online..." -ForegroundColor Green
Try{
    #email does not exist in AAD as Guest, invite and add to Tableau Online Viewers group
    $newGuestAccount = new-GraphGuestInvitation -tokenResponse $tokenResponse -invitedUserEmailAddress $upn -inviteRedirectUrl "https://myapps.microsoft.com/anthesisgroup.com" -sendInvitationMessage $false -Verbose
    $retrievedNewGuestAccount = $null
                Write-Host "Retrieving newly invited guest account..."
                    while(($retrievedNewGuestAccount | Measure-Object).Count -eq 0){
                    $retrievedNewGuestAccount = get-graphUsers -tokenResponse $tokenResponse -filterCustomEq @{"mail" = $newGuestAccount.invitedUserEmailAddress}
                    }
                Write-Host "Found guest account $($retrievedNewGuestAccount.mail), adding to Tableau Online Viewers group..." -ForegroundColor Green
                add-graphUsersToGroup -tokenResponse $tokenResponse -graphGroupId $tableauOnlineViewersGroup -memberType members -graphUserIds $retrievedNewGuestAccount.id -Verbose


}
catch{
$error[0]
}

$checkIfInGroup = get-graphUsersFromGroup -tokenResponse $tokenResponse -groupId $tableauOnlineViewersGroup -memberType Members -Verbose
If($checkIfInGroup.mail -contains $upn){
#if successful, send email to upn
#send-graphMailMessage -tokenResponse $tokenResponse -fromUpn "" -toAddresses "" -subject "Anthesis: You Have Now Been Added to Tableau Online" -bodyHtml $body -saveToSentItems $true -Verbose
}
Else{
Throw "Warning: Guest user not found in Tableau Online Viewers group"
}

}
}

