$365creds = set-MsolCredentials
connect-to365 -credential $365creds

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


$displayName = "Diversity & Inclusivity (Global Steering)"
$managers = convertTo-arrayOfEmailAddresses ("fiona.place@anthesisgroup.com, verity.worthington@anthesisgroup.com") | sort | select -Unique
$members = @()
$members += convertTo-arrayOfEmailAddresses ("Paul.Crewe@anthesisgroup.com,Emma.Armstrong@anthesisgroup.com,Saga.Ekelin@anthesisgroup.com,Candan.Ergeneman@anthesisgroup.com,Merce.Autonell@anthesisgroup.com,Fiona.Place@anthesisgroup.com,Verity.Worthington@anthesisgroup.com") | sort | select -Unique
$members | % {
    $thisEmail = $_
    try{
        $dg = Get-DistributionGroup -Identity $thisEmail -ErrorAction Stop
        if($dg){
            $members += $(enumerate-nestedDistributionGroups -distributionGroupObject $dg -Verbose).WindowsLiveId
            $members = $members | ? {$_ -ne $thisEmail}
            }
        }
    catch{<# Anything that isn't an e-mail address for a Distribution Group will cause errors here, and we don't really care about them #>}
    }
$members = $members | Sort-Object | select -Unique
$managedBy = "365"

if($managedBy -eq "AAD"){$managers = "groupbot@anthesisgroup.com"}
new-teamGroup -displayName $displayName -managerUpns $managers -teamMemberUpns $members -membershipManagedBy $managedBy -tokenResponse $tokenResponse -pnpCreds $365creds -alsoCreateTeam $false -Verbose
