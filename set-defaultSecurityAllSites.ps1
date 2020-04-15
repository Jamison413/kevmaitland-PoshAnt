#set-defaultSecurityAllTeamSites

$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

#$groupAdmin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString ""
#$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
#$exoCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass

$sharePointAdmin = "kimblebot@anthesisgroup.com"
$sharePointAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\KimbleBot.txt) 
$sharePointCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sharePointAdmin, $sharePointAdminPass
#$sharePointCreds = set-MsolCredentials

#connect-ToExo -credential $exoCreds
#Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -Credentials $sharePointCreds

$groupAdmins = get-groupAdminRoleEmailAddresses -tokenResponse $tokenResponse 

$allUnifiedGroups = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -selectAllProperties
$excludeThese = @("teamstestingteam@anthesisgroup.com","apparel@anthesisgroup.com","AccountsPayable@anthesisgroup.com")
$groupsToProcess = $allUnifiedGroups | ? {$excludeThese -notcontains $_.mail -and $_.Displayname -notmatch "Confidential"}
$groupsToProcess | % {
    $tokenResponse = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponse -renewTokenExpiringInSeconds 300 -aadAppCreds $teamBotDetails
    $thisUnifiedGroup = $_
    Write-Host -f Yellow "[$($thisUnifiedGroup.displayName)][$($thisUnifiedGroup.id)][$($thisUnifiedGroup.mail)]"
    Try{
        set-standardSitePermissions -tokenResponse $tokenResponse -graphGroupExtended $thisUnifiedGroup -pnpCreds $sharePointCreds #-suppressEmailNotifications -Verbose:$VerbosePreference
        }
    Catch{
        #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Team [$($combinedMesg.displayName)] settings rolled back" -BodyAsHtml $body -To $($owners.mail) -Cc $itAdminEmailAddresses
        $body = "$($Error[0])<BR><BR>`r`n`r`n$($Error[1])<BR><BR>`r`n`r`n$($Error[2])<BR><BR>`r`n`r`n$($Error[3])"
        Send-MailMessage -From securitybot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Error: set-defaultSecurityAllTeamSites [$($thisUnifiedGroup.displayName)]" -BodyAsHtml $body -To kevin.maitland@anthesisgroup.com -Encoding UTF8
        }

    }

