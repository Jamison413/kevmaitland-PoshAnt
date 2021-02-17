#set-defaultSecurityAllTeamSites
$logFileLocation = "C:\ScriptLogs\"
$logFileName = "set-defaultSecurityAllTeamSites"
if($PSCommandPath){
    $transcriptLogName = "$logFileLocation$logFileName`_Transcript_$(Get-Date -Format "yyMMdd-hhmm").log"
    Start-Transcript $transcriptLogName
    }


$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
#$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\desktop\teambotdetails.txt"
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
Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com" -ClientId $teamBotDetails.ClientID -ClientSecret $teamBotDetails.Secret

$groupAdmins = get-groupAdminRoleEmailAddresses -tokenResponse $tokenResponse 

$allUnifiedGroups = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse -selectAllProperties #-filterDisplayName "Climate & Decarbonisation Community (GBR)"
$excludeThese = @("teamstestingteam@anthesisgroup.com","apparel@anthesisgroup.com","AccountsPayable@anthesisgroup.com")
$groupsToProcess = $allUnifiedGroups | ? {$excludeThese -notcontains $_.mail -and $_.Displayname -notmatch "Confidential"}
for($j=0; $j -lt $groupsToProcess.Count; $j++){
    Write-Progress -Activity "Process security on 365 Groups" -Status "[$j]/[$($groupsToProcess.Count)]"
    $tokenResponse = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponse -renewTokenExpiringInSeconds 300 -aadAppCreds $teamBotDetails
    $thisUnifiedGroup = $groupsToProcess[$j]
    Write-Host -f Yellow "[$j]/[$($groupsToProcess.Count)]: [$($thisUnifiedGroup.displayName)][$($thisUnifiedGroup.id)][$($thisUnifiedGroup.mail)]"
    Try{
        $error.Clear()
        set-standardSitePermissions -tokenResponse $tokenResponse -pnpAppCreds $teamBotDetails -graphGroupExtended $thisUnifiedGroup -pnpCreds $sharePointCreds -ErrorVariable Whoops -Verbose #-suppressEmailNotifications -Verbose:$VerbosePreference
        }
    Catch{
        Write-Host -f Red $(get-errorSummary $_)
        Write-Warning "Something went wrong processing [$($thisUnifiedGroup.displayName)][$($thisUnifiedGroup.id)][$($thisUnifiedGroup.mail)]"
        [string]$body ="<UL>"
        $thisUnifiedGroup.PSObject.Properties | ? {$_.Name -ne "anthesisgroup_UGSync"} | % {
            $body += "`t<LI><B>$($_.Name)</B>`r`n<BR>"
            $body += "`t$($_.Value)</LI>`r`n"
            }
        $body += "<LI><B>anthesisgroup_UGSync</B></LI><UL>"
        $thisUnifiedGroup.anthesisgroup_UGSync.PSObject.Properties | %{
            $body += "`t<LI><B>$($_.Name)</B>`r`n<BR>"
            $body += "`t$($_.Value)</LI>`r`n"
            }
        $body += "</UL>"

        for($i=0;$i -lt $error.Count; $i++){
            $body += "<B>Error [$($i+1)/$($error.Count)] *********************************************************</B><BR><UL>"
            $Error[$i].PSObject.Properties | % {
                $body += "`t<LI><B>$($_.Name)</B>`r`n<BR>"
                $body += "`t$($_.Value)</LI>`r`n"
                }
            if($error[$i].Exception.InnerException){
                $body += "<UL><B>Error [$($i+1).Exception.InnerException)</B>"
                $Error[$i].Exception.InnerException.PSObject.Properties | % {
                    $body += "`t<LI><B>$($_.Name)</B>`r`n<BR>"
                    $body += "`t$($_.Value)</LI>`r`n"
                    }
                if($error[$i].Exception.InnerException.InnerException){
                    $body += "<UL><B>Error [$($i+1).Exception.InnerException.InnerException)</B>"
                    $Error[$i].Exception.InnerException.InnerException.PSObject.Properties | % {
                        $body += "`t<LI><B>$($_.Name)</B>`r`n<BR>"
                        $body += "`t$($_.Value)</LI>`r`n"
                        }
                    $body += "</UL>`r`n`r`n<BR><BR>"
                    }

                $body += "</UL>`r`n`r`n<BR><BR>"
                }

            $body += "</UL>`r`n`r`n<BR><BR>"
            }
        Send-MailMessage -From securitybot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Error: set-defaultSecurityAllTeamSites [$($thisUnifiedGroup.displayName)]" -BodyAsHtml $body -To kevin.maitland@anthesisgroup.com -Encoding UTF8
        #Send-MailMessage -From groupbot@anthesisgroup.com -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject "Team [$($combinedMesg.displayName)] settings rolled back" -BodyAsHtml $body -To $($owners.mail) -Cc $itAdminEmailAddresses
        }

    }
Stop-Transcript