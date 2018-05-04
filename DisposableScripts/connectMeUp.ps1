Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality
Import-Module _REST_Library-SPO.psm1
Import-Module _CSOM_Library-SPO
Import-Module *PNP*

$smtpServer = "anthesisgroup-com.mail.protection.outlook.com"
#$adCredentials = Get-Credential -Message "Enter local AD Administrator credentials to create a new user in AD" -UserName "$env:USERDOMAIN\username"
$msolCredentials = set-MsolCredentials #Set these once as a PSCredential object and use that to build the CSOM SharePointOnlineCredentials object and set the creds for REST
$restCredentials = new-spoCred -username $msolCredentials.UserName -securePassword $msolCredentials.Password
$csomCredentials = new-csomCredentials -username $msolCredentials.UserName -password $msolCredentials.Password
connect-ToMsol -credential $msolCredentials
connect-ToExo -credential $msolCredentials
connect-toAAD -credential $msolCredentials
connect-ToSpo -credential $msolCredentials
