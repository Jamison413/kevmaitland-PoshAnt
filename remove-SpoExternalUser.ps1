Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$sharePointCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

connect-ToSpo -credential $sharePointAdmin

$userToRemove = "Lasse.Kirkelykke@convatec.com"

$extUser = Get-SPOExternalUser -filter $userToRemove
if($extUser){
    Remove-SPOExternalUser $extUser
    }