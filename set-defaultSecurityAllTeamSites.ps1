#set-defaultSecurityAllTeamSites

Import-Module SharePointPnPPowerShellOnline
Import-Module _PNP_Library_SPO

$groupAdmin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString ""
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$sharePointCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
$sharePointCreds = set-MsolCredentials

connect-ToExo -credential $exoCreds
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $sharePointCreds

$teamSites = Get-PnPTenantSite -Template GROUP#0
$teamSites | ? {$_.Url -cmatch "Team" -and $_.Url -notmatch "Confidential"} | % {
    $thisTeamSite = $_
    set-standardTeamSitePermissions -teamSiteAbsoluteUrl $thisTeamSite.Url -adminCredentials $sharePointCreds -verboseLogging $verboseLogging
    }

