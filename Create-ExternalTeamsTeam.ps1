$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails


#Get the upn of the team site and copy here
$365GroupUPN = ""
    
$graphGroupExtended = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -filterUpn $365GroupUPN
    
If(($graphGroupExtended | Measure-Object) -ne 1){        
    Write-Verbose "Provisioning new MS Team (retroactively)"
    $graphTeam = new-graphTeam -tokenResponse $tokenResponse -groupId $graphGroupExtended.id -allowMemberCreateUpdateChannels $true -allowMemberDeleteChannels $false -Verbose:$VerbosePreference -ErrorAction Continue #Create the Team
    if(!$graphTeam){write-warning "Failed to provision Team [$($graphGroupExtended.DisplayName)] via Graph after 3 attempts. Try again later."}
}
Else{
Write-Host "More than one team found" -ForegroundColor Red
}


<# this part isn't finished - can't corroberate with the client name as the client list can be edited after external site creation
#Add a Website tab in the General channel linking back to the Client Site 
if($graphTeam){
Write-Host -f DarkYellow "`tCreating Website Tab in General channel to  Clients/Suppliers Site "
add-graphWebsiteTabToChannel -tokenResponse $tokenResponse -teamId $new365Group.id -channelName "General" -tabName "$($fullRequest.FieldValues.ClientName.Label) Client Data" -tabDestinationUrl $clientOrSupplierSiteDocLib.webUrl -Verbose
}
#>