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


$teamBotDetails = import-encryptedCsv -pathToEncryptedCsv "$env:USERPROFILE\OneDrive - Anthesis LLC\Desktop\teambotdetails.txt"
$tokenResponse = get-graphTokenResponse -aadAppCreds $teamBotDetails

#Unbelievbly, you still can't manage MESGs via Graph.
$groupAdmin = "groupbot@anthesisgroup.com"
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass
connect-ToExo -credential $adminCreds


$all365Groups = get-graphGroupWithUGSyncExtensions -tokenResponse $tokenResponse
$toExclude = @("Sym - Supply Chain","Apparel Team (All)","Teams Testing Team","All Homeworkers (All)")
$365GroupsToProcess = $all365Groups | ? {$toExclude -notcontains $($_.DisplayName) -and $_.DisplayName -notmatch "Confidential" -and $_.DisplayName -notmatch "All "}

$adminEmailAddresses = get-groupAdminRoleEmailAddresses -tokenResponse $tokenResponse

$365GroupsToProcess | % {
    #$tokenResponse = test-graphBearerAccessTokenStillValid -tokenResponse $tokenResponse -renewTokenExpiringInSeconds 30 -aadAppCreds $teamBotDetails  #Uncomment this when the whole sync processs takes >1h
    $365Group = $_
    try{
        #sync-groupMemberships_deprecated -UnifiedGroup $365Group -syncWhat Members -sourceGroup $365Group.CustomAttribute6 -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true}# -Verbose 
        sync-groupMemberships -tokenResponse $tokenResponse -graphExtendedUG $365Group -syncWhat Members -sourceGroup $365Group.anthesisgroup_UGSync.masterMembershipList -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true
        }
    catch{
        $_
        Send-MailMessage -To $adminEmailAddresses  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnfiedGroupMembership [$($365Group.DisplayName)]" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Members" -From "$env:COMPUTERNAME@anthesisgroup.com"
        continue
        }
    try{
        #sync-groupMemberships -UnifiedGroup $365Group -syncWhat Owners -sourceGroup AAD -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true} # -Verbose
        sync-groupMemberships -tokenResponse $tokenResponse -graphExtendedUG $365Group -syncWhat Owners -sourceGroup AAD -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true
        }
    catch{        
        $_
        Send-MailMessage -To $adminEmailAddresses  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnfiedGroupMembership [$($365Group.DisplayName)]" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Owners" -From "$env:COMPUTERNAME@anthesisgroup.com"
        continue
        }
    $365GroupsToProcess = $365GroupsToProcess | ? {$_.Id -ne $365Group.Id}
    }

if($365GroupsToProcess.Count -gt 0){
    Send-MailMessage -To kevin.maitland@anthesisgroup.com  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnfiedGroupMembership [$($365GroupsToProcess.Count)] 365Groups remain unprocessed" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nGroups are: `r`n`t$($365GroupsToProcess.DisplayName -join "`r`n`t")" -From "$env:COMPUTERNAME@anthesisgroup.com"
    }

Stop-Transcript
