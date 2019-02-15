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

$enabledUsers = Get-MsolUser -EnabledFilter EnabledOnly -All | ? {$_.UserType -eq "Member"}
Write-Information "$($enabledUsers.Count) Enabled User accounts found"
$enabledUsersWithoutMFA = $enabledUsers | ? {[string]::IsNullOrWhiteSpace($_.StrongAuthenticationRequirements)} | Sort-Object UsageLocation, DisplayName
Write-Information "$($enabledUsersWithoutMFA.Count) Enabled User accounts without MFA enabled found"
$enabledUsersWithMFA = $enabledUsers | ? {![string]::IsNullOrWhiteSpace($_.StrongAuthenticationRequirements)}
$suboptimalEnabledUsersWithMFA = $enabledUsersWithMFA | ? {"PhoneAppNotification" -ne $($_.StrongAuthenticationMethods | ? {$_.IsDefault -eq $true}).MethodType} | Sort-Object UsageLocation, DisplayName
Write-Information "$($suboptimalEnabledUsersWithMFA.Count) Suboptimally configured User accounts with MFA enabled found"



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
$body += "Love,`r`n`r`n<BR><BR>The Helpful MFA Robot</FONT></HTML>"
#Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulmfarobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
Write-Information $body
Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulmfarobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Credential $adminCreds
Write-Information "Message Sent (maybe)"
#$body
Stop-Transcript