$365creds = set-MsolCredentials
$restCredentials = new-spoCred -username $365creds.UserName -securePassword $365creds.Password
connect-ToMsol -credential $365creds
connect-ToExo -credential $365creds
connect-toAAD -credential $365creds
connect-toTeams -credential $365creds

$teamBotDetails = Import-Csv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
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

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/clients" -Credentials $365creds
$requests = Get-PnPListItem -List "External Client Site Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='String'>Awaiting creation</Value></Eq></Where></Query></View>"
if($requests){$selectedRequests = $requests | select {$_.FieldValues.Title},{$_.FieldValues.ClientName.Label},{$_.FieldValues.Site_x0020_Admin.LookupValue},{$_.FieldValues.Site_x0020_Owners.LookupValue -join ", "},{$_.FieldValues.Site_x0020_Members.LookupValue -join ", "},{$_.FieldValues.GUID.Guid} | Out-GridView -PassThru -Title "Highlight any requests to process and click OK"}
foreach ($currentRequest in $selectedRequests){
    $fullRequest = $requests | ? {$_.FieldValues.GUID.Guid -eq $currentRequest.'$_.FieldValues.GUID.Guid'}
    $managers = convertTo-arrayOfEmailAddresses ($fullRequest.FieldValues.Site_x0020_Owners.Email +","+ $fullRequest.FieldValues.Site_x0020_Admin.Email+","+ $((Get-PnPConnection).PSCredential.UserName)) | sort | select -Unique
    $members = convertTo-arrayOfEmailAddresses ($managers + $fullRequest.FieldValues.Site_x0020_Members.Email) | sort | select -Unique
    new-externalGroup -displayName $fullRequest.FieldValues.Title -managerUpns $managers -teamMemberUpns $members -membershipManagedBy 365 -tokenResponse $tokenResponse -alsoCreateTeam $false -pnpCreds $365creds -Verbose
    }

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/subs" -Credentials $365creds
$requests += Get-PnPListItem -List "External Subcontractor Site Requests" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='String'>Awaiting creation</Value></Eq></Where></Query></View>"

