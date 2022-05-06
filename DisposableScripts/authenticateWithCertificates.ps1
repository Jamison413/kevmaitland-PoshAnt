$certName = $(whoami /upn)
$cert = New-SelfSignedCertificate -Subject "CN=$certName" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256

Export-Certificate -Cert $cert -FilePath "$env:USERPROFILE\Downloads\$certName.cer"   ## Specify your preferred location and replace {certificateName}

$userBotDetails = get-graphAppClientCredentials -appName UserBot
$userBotDetails.ClientID = "8acd9949-38f0-4688-961f-f162a7958bed"

Connect-MgGraph -ClientID 8acd9949-38f0-4688-961f-f162a7958bed -TenantId $userBotDetails.TenantId -CertificateName "CN=$env:USERNAME" ## Or -CertificateThumbprint instead of -CertificateName

$users = Get-MgUser -All

$ms = Measure-Command {$users = Get-MgUser -All}
$km = Measure-Command {$kUsers = get-graphUsers -tokenResponse $tokenTeams}

$ms.totalseconds
$km.totalseconds

$tokenTeams = get-graphTokenResponse -grant_type client_credentials -aadAppCreds $(get-graphAppClientCredentials -appName TeamsBot)

get-available365licensecount -licensetype All