$logFileLocation = "C:\ScriptLogs\"
$transcriptLogName = "$($logFileLocation+$(split-path $PSCommandPath -Leaf))_Transcript_$(Get-Date -Format "yyMMdd").log"
if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)){
    $fullLogPathAndName = $logFileLocation+"report-MFA_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = $logFileLocation+"report-MFA_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
else{
    $fullLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_FullLog_$(Get-Date -Format "yyMMdd").log"
    $errorLogPathAndName = "$($logFileLocation+$MyInvocation.MyCommand)_ErrorLog_$(Get-Date -Format "yyMMdd").log"
    }
Start-Transcript $transcriptLogName -Append

Import-Module _PS_Library_GeneralFunctionality
Import-Module _PS_Library_MSOL

$Admin = "kevin.maitland@anthesisgroup.com"
$AdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Desktop\Kev.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin, $AdminPass

connect-ToMsol -credential $adminCreds


#For Licensed User Accounts

$enabledUsers = Get-MsolUser -EnabledFilter EnabledOnly -All | ? {$_.UserType -eq "Member"} | ? {$_.IsLicensed -eq "True"}  
Write-Information "$($enabledUsers.Count) Enabled User accounts found"
$enabledUsersWithoutMFA = $enabledUsers | ? {[string]::IsNullOrWhiteSpace($_.StrongAuthenticationRequirements)} | Sort-Object UsageLocation, DisplayName 
Write-Information "$($enabledUsersWithoutMFA.Count) Enabled User accounts without MFA enabled found"
$enabledUsersWithMFA = $enabledUsers | ? {![string]::IsNullOrWhiteSpace($_.StrongAuthenticationRequirements)}
$suboptimalEnabledUsersWithMFA = $enabledUsersWithMFA | ? {"PhoneAppNotification" -ne $($_.StrongAuthenticationMethods | ? {$_.IsDefault -eq $true}).MethodType} | Sort-Object UsageLocation, DisplayName
Write-Information "$($suboptimalEnabledUsersWithMFA.Count) Suboptimally configured User accounts with MFA enabled found"
$optimalEnabledUsersWithMFA = $enabledUsersWithMFA | ? {"PhoneAppNotification" -eq $($_.StrongAuthenticationMethods | ? {$_.IsDefault -eq $true}).MethodType} | Sort-Object UsageLocation, DisplayName
Write-Information "$($OptimalEnabledUsersWithMFA.Count) Optimally configured User accounts with MFA enabled found"

#User MFA Stats

[INT]$TotalMFACount = $enabledUsers.Count
[INT]$TotalEnabledCount = $enabledUsersWithMFA.Count
#Extra deets
    [INT]$Totalsuboptimal = $suboptimalEnabledUsersWithMFA.Count
    [INT]$Totaloptimal = $optimalEnabledUsersWithMFA.Count
[INT]$TotalWithoutMFA = $enabledUsersWithoutMFA.Count
#Find the total percentage
[INT]$TotalMFAScore = $TotalEnabledCount/$TotalMFACount*100
[INT]$TotalOptimalMFAScore = $Totaloptimal/$TotalEnabledCount*100

Write-host "$TotalMFAScore%"
Write-host "$TotalOptimalMFAScore%"





$subject = "MFA status report"
$body = "<HTML><FONT FACE=`"Calibri`">Hello IT Team,`r`n`r`n<BR><BR>"
#$body += $ownerReport.To+"`r`n`r`n<BR><BR>"

$body += "The following users do not have MFA enabled on their accounts:"
$body += "      `r`n`t<BR><PRE>&#9;"
$enabledUsersWithoutMFA | % {$body += "$("$($_.UsageLocation)`t$($_.UserPrincipalName) `r`n`t")"}
$body += "</PRE>`r`n`r`n<BR>"

$body += "The following users do have MFA configured without App Notifications as their default:"
$body += "      `r`n`t<BR><PRE>&#9;"
$suboptimalEnabledUsersWithMFA | % {$body += $("$($_.UsageLocation)`t$($_.UserPrincipalName) `r`n`t")}
$body += "</PRE>`r`n`r`n<BR>"

$body += "The following users do have MFA configured with App Notifications as their default:"
$body += "      `r`n`t<BR><PRE>&#9;"
$OptimalEnabledUsersWithMFA | % {$body += $("$($_.UsageLocation)`t$($_.UserPrincipalName) `r`n`t")}
$body += "</PRE>`r`n`r`n<BR>"

$body += "Current Statistics for Anthesis:"
$body += "      `r`n`t<BR><PRE>&#9;"
$body += "Total MFA Score: $TotalMFAScore%"
$body += "      `r`n`t<BR><PRE>&#9;"
$body += "Total MFA Optimal Score: $TotalOptimalMFAScore%"


$body += "Love,`r`n`r`n<BR><BR>The Helpful MFA Robot</FONT></HTML>"
#Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulmfarobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
Write-Information $body
Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulmfarobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Credential $adminCreds
Write-Information "Message Sent (maybe)"
#$body
Stop-Transcript