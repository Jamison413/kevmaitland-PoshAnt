Import-Module _PS_Library_MSOL.psm1
Import-Module _PS_Library_GeneralFunctionality

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$sharePointCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

connect-ToSpo -credential $sharePointAdmin

$userToRemove = "emilypressey@hotmail.co.uk"
$siteToRemoveFrom = "https://anthesisllc.sharepoint.com/sites/external/UKResourcesCouncilWorkingGroupsWasteSectorPlan"

$extUser = Get-SPOExternalUser -filter $userToRemove
$extUser2 = Get-SPOExternalUser -filter $userToRemove -SiteUrl $siteToRemoveFrom
if($extUser){
    Remove-SPOExternalUser -UniqueIDs $extUser.UniqueId
    }
if($extUser){
    Remove-SPOExternalUser -UniqueIDs $extUser2.UniqueId
    }