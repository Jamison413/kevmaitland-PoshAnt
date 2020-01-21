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

Import-Module _PS_Library_MSOL
Import-Module _PS_Library_Groups
Import-Module _PS_Library_GeneralFunctionality

$groupAdmin = "groupbot@anthesisgroup.com"
#convertTo-localisedSecureString ""
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass

connect-ToMsol -credential $adminCreds
connect-toAAD -credential $adminCreds
connect-ToExo -credential $adminCreds

$all365Groups = Get-UnifiedGroup
$toExclude = @("Sym - Supply Chain","Apparel Team (All)","All North America","Business Development Team (GBR)","Pre Sales Team (All)","Teams Testing Team", "Finance Team (North America)","Finance Team (North America)")
$365GroupsToProcess = $all365Groups | ? {$toExclude -notcontains $($_.DisplayName) -and $_.DisplayName -notmatch "Confidential"}

$adminEmailAddresses = get-groupAdminRoleEmailAddresses

$365GroupsToProcess | % {
    $365Group = $_
    try{sync-groupMemberships -UnifiedGroup $365Group -syncWhat Members -sourceGroup $365Group.CustomAttribute6 -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true } #-Verbose }
    catch{
        $_
        Send-MailMessage -To $adminEmailAddresses  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnfiedGroupMembership [$($365Group.DisplayName)]" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Members" -From "$env:COMPUTERNAME@anthesisgroup.com"
        continue
        }
    try{sync-groupMemberships -UnifiedGroup $365Group -syncWhat Owners -sourceGroup AAD -adminEmailAddresses $adminEmailAddresses -enumerateSubgroups $true} # -Verbose}
    catch{        
        $_
        Send-MailMessage -To $adminEmailAddresses  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnfiedGroupMembership [$($365Group.DisplayName)]" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nError occurred synchronising Owners" -From "$env:COMPUTERNAME@anthesisgroup.com"
        continue
        }
    $365GroupsToProcess = $365GroupsToProcess | ? {$_.ExternalDirectoryObjectId -ne $365Group.ExternalDirectoryObjectId}
    }

if($365GroupsToProcess.Count -gt 0){
    Send-MailMessage -To kevin.maitland@anthesisgroup.com  -SmtpServer anthesisgroup-com.mail.protection.outlook.com -Subject "FAILED: sync-UnfiedGroupMembership [$($365GroupsToProcess.Count)] 365Groups remain unprocessed" -Priority High -Body "$_`r`n`r`nError recorded in [$transcriptLogName] on [$env:COMPUTERNAME]`r`n`r`nGroups are: `r`n`t$($365GroupsToProcess.DisplayName -join "`r`n`t")" -From "$env:COMPUTERNAME@anthesisgroup.com"
    }

Stop-Transcript
