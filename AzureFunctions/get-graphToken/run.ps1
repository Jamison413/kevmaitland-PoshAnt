using namespace System.Net
# Input bindings are passed in via param block
param($Request, $TriggerMetadata)
# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."
# Interact with query parameters or the body of the request
.$Scope = $Request.Query.Scope
if (-not $Scope) {
    $Scope = $Request.Body.Scope
}
#If parameter "Scope" has not been provided, we assume that graph.microsoft.com is the target resource
If (!$Scope) {
    $Scope = "https://graph.microsoft.com/"
}
$tokenAuthUri = $env:IDENTITY_ENDPOINT + "?resource=$Scope&api-version=2019-08-01"
$response = Invoke-RestMethod -Method Get -Headers @{"X-IDENTITY-HEADER"="$env:IDENTITY_HEADER"} -Uri $tokenAuthUri -UseBasicParsing
$accessToken = $response.access_token
#Invoke REST call to Graph API
$uri = 'https://graph.microsoft.com/v1.0/groups'
$authHeader = @{    
'Content-Type'='application/json'
'Authorization'='Bearer ' +  $accessToken
}
$result = (Invoke-RestMethod -Uri $uri -Headers $authHeader -Method Get -ResponseHeadersVariable RES).value
If ($result) {
    $body = $resul  t
    $StatusCode = [HttpStatusCode]::OK
}
Else {
    $body = $RES
    $StatusCode = [HttpStatusCode]::BadRequest}
# Associate values to output bindings by calling 'Push-OutputBinding'
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = $StatusCode
    Body = $body
})