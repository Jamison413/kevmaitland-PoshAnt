$teamBotDetails = Import-Csv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"

$tenantId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.TenantId)
$clientId = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.ClientID)
$secret   = decrypt-SecureString (ConvertTo-SecureString $teamBotDetails.Secret)


$reqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $secret
    } 
 
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody




$apiUrl = 'https://graph.microsoft.com/v1.0/Groups/'
$data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $apiUrl -Method Get
$Groups = ($Data | select-object Value).Value
$data.'@odata.context'

