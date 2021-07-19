$logFileLocation = "C:\ScriptLogs\"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"sync-UnifiedGroupMembershipChanges_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"sync-UnifiedGroupMembershipChanges_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
if($PSCommandPath){
    $transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
    Start-Transcript $transcriptLogName -Append
    }


$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponseTeams = get-graphTokenResponse -aadAppCreds $teamBotDetails
$smtpBotDetails = get-graphAppClientCredentials -appName SmtpBot
$tokenResponseSmtp = get-graphTokenResponse -aadAppCreds $smtpBotDetails

#Unbelievbly, you still can't manage MESGs via Graph.
$groupAdmin = "groupbot@anthesisgroup.com"
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass
connect-ToExo -credential $adminCreds


$all365Groups = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponseTeams
$toExclude = @("Sym - Supply Chain","Apparel Team (All)","Teams Testing Team","All Homeworkers (All)","Archived Finance Team (North America)")
$365GroupsToProcess = $all365Groups | ? {$toExclude -notcontains $($_.DisplayName) -and $_.DisplayName -notmatch "Confidential"}# -and $_.DisplayName -notmatch "All "}

$adminEmailAddresses = get-groupAdminRoleEmailAddresses -tokenResponse $tokenResponseTeams

#$365GroupsToProcess | % {
$timeForFullCycle = Measure-Command {

    for($i=0;$i -lt $365GroupsToProcess.Count;$i++){
        $tokenResponseTeams = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseTeams -renewTokenExpiringInSeconds 300 -aadAppCreds $teamBotDetails  #Uncomment this when the whole sync processs takes >1h
        $tokenResponseSmtp = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponseSmtp -renewTokenExpiringInSeconds 300 -aadAppCreds $smtpBotDetails  #Uncomment this when the whole sync processs takes >1h
        $365Group = $365GroupsToProcess[$i]
        Write-Host "[$($i)]/[$($365GroupsToProcess.Count)]: [$($365Group.displayName)]"
        Write-Progress -Activity "Synchronising Group Memberships" -Status "[$($i)]/[$($365GroupsToProcess.Count)]: [$($365Group.displayName)]"
        try{
            #sync-groupMemberships_deprecated -UnifiedGroup $365Group -syncWhat Members -sourceGroup $365Group.CustomAttribute6 -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true}# -Verbose 
            sync-groupMemberships -tokenResponse $tokenResponseTeams -tokenResponseSmtp $tokenResponseSmtp -graphExtendedUG $365Group -syncWhat Members -sourceGroup $365Group.anthesisgroup_UGSync.masterMembershipList -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true
            }
        catch{
            Write-Host -ForegroundColor Red $(get-errorSummary $_)
            if(![string]::IsNullOrWhiteSpace($365Group.anthesisgroup_UGSync.classification) -and ![string]::IsNullOrWhiteSpace($365Group.anthesisgroup_UGSync.masterMembershipList)){
                try{ #If we've got enough data to automtacilly repair the broken group, try repairing and reporocessing the group
                    repair-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponseTeams -graphGroup $365Group -groupClassifcation $365Group.anthesisgroup_UGSync.classification -masterMembership $365Group.anthesisgroup_UGSync.masterMembershipList -createGroupsIfMissing -Verbose:$VerbosePreference
                    $365Group = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponseTeams -filterId $365Group.id
                    sync-groupMemberships -tokenResponse $tokenResponseTeams -tokenResponseSmtp $tokenResponseSmtp -graphExtendedUG $365Group -syncWhat Members -sourceGroup $365Group.anthesisgroup_UGSync.masterMembershipList -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true
                    }
                catch{
                    send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn $groupAdmin -toAddresses $adminEmailAddresses -subject "FAILED: sync-UnifiedGroupMembership [$($365Group.DisplayName)]" -bodyText "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Members`r`n`r`n$(get-errorSummary $_)" -priority high
                    #Send-MailMessage -To $adminEmailAddresses  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnfiedGroupMembership [$($365Group.DisplayName)]" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Members" -From "$env:COMPUTERNAME@anthesisgroup.com" 
                    continue
                    }
                }
            else{
                send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn $groupAdmin -toAddresses $adminEmailAddresses -subject "FAILED: sync-UnifiedGroupMembership [$($365Group.DisplayName)]" -bodyText "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Members`r`n`r`n$(get-errorSummary $_)" -priority high
                #Send-MailMessage -To $adminEmailAddresses  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnifiedGroupMembership [$($365Group.DisplayName)]" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Members" -From "$env:COMPUTERNAME@anthesisgroup.com"
                }
            continue
            }
        try{
            #sync-groupMemberships -UnifiedGroup $365Group -syncWhat Owners -sourceGroup AAD -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true} # -Verbose
            sync-groupMemberships -tokenResponse $tokenResponseTeams -tokenResponseSmtp $tokenResponseSmtp -graphExtendedUG $365Group -syncWhat Owners -sourceGroup AAD -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true
            }
        catch{        
            Write-Host -ForegroundColor Red $(get-errorSummary $_)
            if(![string]::IsNullOrWhiteSpace($365Group.anthesisgroup_UGSync.classification) -and ![string]::IsNullOrWhiteSpace($365Group.anthesisgroup_UGSync.masterMembershipList)){
                try{ #If we've got enough data to automtacilly repair the broken group, try repairing and reporocessing the group
                    repair-graphGroupUGSyncSchemaExtensions -tokenResponse $tokenResponseTeams -graphGroup $365Group -groupClassifcation $365Group.anthesisgroup_UGSync.classification -masterMembership $365Group.anthesisgroup_UGSync.masterMembershipList -createGroupsIfMissing -Verbose:$VerbosePreference
                    $365Group = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponseTeams -filterId $365Group.id
                    sync-groupMemberships -tokenResponse $tokenResponseTeams -tokenResponseSmtp $tokenResponseSmtp -graphExtendedUG $365Group -syncWhat Owners -sourceGroup AAD -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true
                    }
                catch{
                    send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn $groupAdmin -toAddresses $adminEmailAddresses -subject "FAILED: sync-UnifiedGroupMembership [$($365Group.DisplayName)]" -bodyText "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Owners`r`n`r`n$(get-errorSummary $_)" -priority high
                    #Send-MailMessage -To $adminEmailAddresses  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnfiedGroupMembership [$($365Group.DisplayName)]" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Owners" -From "$env:COMPUTERNAME@anthesisgroup.com"
                    continue
                    }
                }
            else{
                send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn $groupAdmin -toAddresses $adminEmailAddresses -subject "FAILED: sync-UnifiedGroupMembership [$($365Group.DisplayName)]" -bodyText "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Owners`r`n`r`n$(get-errorSummary $_)" -priority high
                #Send-MailMessage -To $adminEmailAddresses  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnfiedGroupMembership [$($365Group.DisplayName)]" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Owners" -From "$env:COMPUTERNAME@anthesisgroup.com"
                }
            continue
            }
        #$365GroupsToProcess = $365GroupsToProcess | ? {$_.Id -ne $365Group.Id}

        }

#if($365GroupsToProcess.Count -gt 0){
#    Send-MailMessage -To kevin.maitland@anthesisgroup.com  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnfiedGroupMembership [$($365GroupsToProcess.Count)] 365Groups remain unprocessed" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nGroups are: `r`n`t$($365GroupsToProcess.DisplayName -join "`r`n`t")" -From "$env:COMPUTERNAME@anthesisgroup.com"
#    }
    }
Write-Host "Processing complete at [$(get-date -Format s)] in [$($timeForFullCycle.TotalMinutes)] minutes ([$($timeForFullCycle.TotalSeconds)] seconds)"

Stop-Transcript
