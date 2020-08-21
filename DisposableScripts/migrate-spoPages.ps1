$sharePointAdmin = "kimblebot@anthesisgroup.com"
#convertTo-localisedSecureString "KimbleBotPasswordHere"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content "$env:USERPROFILE\Desktop\KimbleBot.txt") 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/sites/Resources-Finance/" -Credentials $adminCreds
$sitePages = Get-PnPList "SitePages"
$pages = Get-PnPListItem -List $sitePages

$pages[0].FieldValues
Get-PnPClientSidePage 
for ($i=2; $i -lt $pages.Count; $i++ ) {
    $thisPage = $pages[$i]
    #continue
    Write-Host -f DarkYellow "`tCopying $($thisPage.FieldValues.FileLeafRef)"
    copy-spoPage -sourceUrl "https://anthesisllc.sharepoint.com$($thisPage.FieldValues.FileRef)" -destinationSite "https://anthesisllc.sharepoint.com/teams/External_-_NetSuite_Training_Materials_365" -pnpCreds $365creds -overwriteDestinationFile $true  | Out-Null
    copy-spoPage -sourceUrl "https://anthesisllc.sharepoint.com$($thisPage.FieldValues.FileRef)" -destinationSite "https://anthesisllc.sharepoint.com/teams/External_-_NetSuite_Training_Materials_365" -pnpCreds $adminCreds -overwriteDestinationFile $true  | Out-Null
    }