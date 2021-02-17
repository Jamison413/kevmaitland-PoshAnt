Import-Module _PS_Library_Graph.psm1
Import-Module _PNP_Library_SPO.psm1
Import-Module _CSOM_Library-SPO.psm1
Import-Module MicrosoftTeams
Import-Module _PS_Library_UserManagement.psm1


#Conn - CSOM for SharepointUserID
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\kimblebot.txt) 
$spoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$peopleservicessite = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365/"
$conn = Connect-PnPOnline -Url $peopleservicessite -Credentials $spoCreds
$ctx = Get-PnPContext


#Conn - Graph for overall user profile
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


Write-Host "Emptying the staff list..." -ForegroundColor Yellow
#Empty the staff list first
$graphSiteId = "anthesisllc.sharepoint.com,cd82f435-8404-4c16-9ef5-c1e357ac5b96,2373d950-6dea-4ed5-9224-dea4c41c7da3"
$directoryListId = "009bb573-f305-402d-9b21-e6f597473256"
$allanthesians = get-graphListItems -tokenResponse $tokenResponse -graphSiteId $graphSiteId -listId $directoryListId -expandAllFields
$c = 0
ForEach($item in $allanthesians){
    $c++
    $graphItemId = $item.id
    Write-Host "Deleting item $($c)/$($allanthesians.count)" -ForegroundColor White
    delete-graphListItem -tokenResponse $tokenResponse -graphSiteId $graphSiteId -graphListId $directoryListId -graphItemId $graphItemId

}
Write-Host "Getting all graph users..." -ForegroundColor Yellow
#Get all graph users
$usersarray = get-graphUsers -tokenResponse $tokenResponse -filterLicensedUsers:$true -selectAllProperties:$true -Verbose
$allgraphusers = remove-mailboxesandbots -usersarray $usersarray



#Iterate through each graph and Sharepoint profile to get key details, we will iterate through the Exchange profile at the end for efficiency (for timezone early)
Write-Host "Getting key user details..." -ForegroundColor Yellow
#Set counter
$i = 1
$fullUsers = @()
ForEach($graphuser in $allgraphusers){


#Reset connection on 25th run
$i++
If($i -eq 25){
Write-Host "Resetting the connection!" -ForegroundColor Green

#Conn - CSOM for SharepointUserID
$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\kimblebot.txt) 
$spoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$peopleservicessite = "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365/"
$conn = Connect-PnPOnline -Url $peopleservicessite -Credentials $spoCreds
$ctx = Get-PnPContext


#Conn - Graph for overall user profile
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

$i = 1
}

Write-Host "Processing: $($graphuser.userPrincipalName)" -ForegroundColor White

#SPO - for Sharepoint lookup ID to input data into people picker columns
$spomanager = $null
$context = $ctx
#Fetch the users in Site Collection
$sharepointUsers = $context.Web.SiteUsers
$context.Load($sharepointUsers)
$context.ExecuteQuery()
$spoUser = $context.web.EnsureUser("i:0#.f|membership|$($graphuser.userPrincipalName)")
$context.Load($spoUser)
$context.ExecuteQuery()

#Graph and SPO manager
$graphQuery = "/users/$($graphuser.id)/manager"
$graphmanager = invoke-graphGet -tokenResponse $tokenResponse -graphQuery $graphQuery -ErrorAction SilentlyContinue
    If($graphmanager){
        $context = $ctx
        $sharepointUsers = $context.Web.SiteUsers
        $context.Load($sharepointUsers)
        $context.ExecuteQuery()
        $spomanager = $context.web.EnsureUser("i:0#.f|membership|$($graphmanager.userPrincipalName)")
        $context.Load($spomanager)
        $context.ExecuteQuery()
        }
    Else{
    Write-Host "WARNING: No line manager for $($graphuser.userPrincipalName)" -ForegroundColor White
    }

#EXO mailbox timezone
$exoTimezone = get-graphMailboxSettings -tokenResponse $tokenResponse -identity "$($graphuser.userPrincipalName)" -Verbose
#Philippine's uses several timezone names
If(($exoTimezone.timeZone -eq "Singapore Standard Time") -or ($exoTimezone.timeZone -eq "Taipei Standard Time") -or ($exoTimezone.timeZone -eq "China Standard Time")){
$exoTimezone = New-Object -TypeName psobject @{
"DisplayName" = "(UTC+08:00) Philippine Standard Time"
}
}
Else{
$exoTimezone = Get-TimeZone $exoTimezone.timeZone
}

#sharepoint timezone
$spoTimezone =  Get-PnPUserProfileProperty -Account $($graphuser.userPrincipalName)


$antUser = New-Object psobject -Property @{
    graphuser = $graphuser
    spouser = $spoUser
    exotimezone = $exoTimezone
    linemanager = $spomanager
    spotimezone = $spoTimezone
    teamslink = "https://teams.microsoft.com/l/chat/0/0?users=" + "$($graphuser.userPrincipalName)"
}
Write-Host "Adding: $($graphuser.displayName) to the running list" -ForegroundColor Yellow
$fullUsers += $antUser

}

Write-Host "Adding each user to the Directory..." -ForegroundColor Yellow
#Add each user to the Directory
$i = 1
ForEach($user in $fullUsers){

#Reset connection on 85th run
$i++
If($i -eq 85){
Write-Host "Resetting the connection!" -ForegroundColor Green

#Conn - Graph for overall user profile
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

$i = 1
}

#Conn - Graph for overall user profile
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

Write-Host "Adding $($user.graphuser.userPrincipalName) to the Directory" -ForegroundColor Yellow
If($user.linemanager.Id){
#Line manager
$body = "{
  `"fields`": {
    `"AnthesianLookupId`": `"$($user.spouser.Id)`",
    `"JobTitle`": `"$($user.graphuser.jobTitle)`",
    `"Community`": `"$($user.graphuser.department)`",
    `"Office_x0020_Phone`": `"$($user.graphuser.businessPhones)`",
    `"CellPhone`": `"$($user.graphuser.mobilePhone)`",
    `"City`": `"$($user.graphuser.city)`",
    `"Office`": `"$($user.graphuser.officeLocation)`",
    `"Country`": `"$($user.graphuser.country)`",
    `"Timezone`": `"$($user.exotimezone.DisplayName)`",
    `"ManagerLookupId`": `"$($user.linemanager.Id)`",
    `"ManagerEmail`": `"$($user.linemanager.Email)`",
    `"BusinessUnit`": `"$($user.graphuser.anthesisgroup_employeeInfo.businessUnit)`",
    `"Email`": `"$($user.graphuser.userPrincipalName)`",
    `"TeamsLink`": `"$($user.teamslink)`",
    `"UserGUID`": `"$($user.graphuser.id)`",
    `"plaintextname`": `"$($user.graphuser.displayName)`",

  }
}"
}
Else{
#No line manager
$body = "{
  `"fields`": {
    `"AnthesianLookupId`": `"$($user.spouser.Id)`",
    `"JobTitle`": `"$($user.graphuser.jobTitle)`",
    `"Community`": `"$($user.graphuser.department)`",
    `"Office_x0020_Phone`": `"$($user.graphuser.businessPhones)`",
    `"CellPhone`": `"$($user.graphuser.mobilePhone)`",
    `"City`": `"$($user.graphuser.city)`",
    `"Office`": `"$($user.graphuser.officeLocation)`",
    `"Country`": `"$($user.graphuser.country)`",
    `"Timezone`": `"$($user.exotimezone.DisplayName)`",
    `"BusinessUnit`": `"$($user.graphuser.anthesisgroup_employeeInfo.businessUnit)`",
    `"Email`": `"$($user.graphuser.userPrincipalName)`",
    `"TeamsLink`": `"$($user.teamslink)`",
    `"UserGUID`": `"$($user.graphuser.id)`",
    `"plaintextname`": `"$($user.graphuser.displayName)`",

  }
}"
}
$graphQuery = "https://graph.microsoft.com/v1.0/sites/anthesisllc.sharepoint.com,cd82f435-8404-4c16-9ef5-c1e357ac5b96,2373d950-6dea-4ed5-9224-dea4c41c7da3/lists/009bb573-f305-402d-9b21-e6f597473256/items"
$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$response = Invoke-RestMethod -Uri "$graphQuery" -Body $body -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Method Post -verbose
}





