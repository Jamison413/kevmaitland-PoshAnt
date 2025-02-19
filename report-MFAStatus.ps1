﻿$logFileLocation = "C:\ScriptLogs\"
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
Import-Module _PS_Library_Graph
Import-Module _PS_Library_MFA

$groupAdmin = "groupbot@anthesisgroup.com"
$groupAdminPass = ConvertTo-SecureString (Get-Content $env:USERPROFILE\Downloads\GroupBot.txt) 
$adminCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $groupAdmin, $groupAdminPass
connect-ToExo -credential $adminCreds


connect-ToMsol -credential $adminCreds

$teamBotDetails = get-graphAppClientCredentials -appName TeamsBot
$tokenResponseTeams = get-graphTokenResponse -aadAppCreds $teamBotDetails
$smtpBotDetails = get-graphAppClientCredentials -appName SmtpBot
$tokenResponseSmtp = get-graphTokenResponse -aadAppCreds $smtpBotDetails


#For Licensed User Accounts

$enabledUsers = Get-MsolUser -EnabledFilter EnabledOnly -All | ? {$_.UserType -eq "Member"} | ? {$_.IsLicensed -eq "True"}  
Write-Information "$($enabledUsers.Count) Enabled User accounts found"

#AVD MFA amends
$AVDGBR = get-graphUsersFromGroup -tokenResponse $tokenResponseTeams -groupId "1e15404b-f737-45f2-bc04-c3d0c3173ab5" -memberType Members -returnOnlyUsers
$AVDNA = get-graphUsersFromGroup -tokenResponse $tokenResponseTeams -groupId "c9e78db8-9ad1-43ed-a414-f9968e013ad5" -memberType Members -returnOnlyUsers

$enabledUsers = Compare-Object -ReferenceObject $enabledUsers -DifferenceObject $AVDGBR -Property "userPrincipalName" -IncludeEqual -PassThru #remove AVD GBR users
$enabledUsers = $enabledUsers.Where({$_.SideIndicator -ne "=="})
$enabledUsers = Compare-Object -ReferenceObject $enabledUsers -DifferenceObject $AVDNA -Property "userPrincipalName" -IncludeEqual -PassThru #remove AVD NA users
$enabledUsers = $enabledUsers.Where({$_.SideIndicator -ne "=="})

#Process remaining targetted users
$enabledUsersWithoutMFA = $enabledUsers | ? {[string]::IsNullOrWhiteSpace($_.StrongAuthenticationRequirements)} | Sort-Object UsageLocation, DisplayName 
Write-Information "$($enabledUsersWithoutMFA.Count) Enabled User accounts without MFA enabled found"
$enabledUsersWithMFA = $enabledUsers | ? {![string]::IsNullOrWhiteSpace($_.StrongAuthenticationRequirements)}
$suboptimalEnabledUsersWithMFA = $enabledUsersWithMFA | ? {"PhoneAppNotification" -ne $($_.StrongAuthenticationMethods | ? {$_.IsDefault -eq $true}).MethodType} | Sort-Object UsageLocation, DisplayName
Write-Information "$($suboptimalEnabledUsersWithMFA.Count) Suboptimally configured User accounts with MFA enabled found"
$optimalEnabledUsersWithMFA = $enabledUsersWithMFA | ? {"PhoneAppNotification" -eq $($_.StrongAuthenticationMethods | ? {$_.IsDefault -eq $true}).MethodType} | Sort-Object UsageLocation, DisplayName
Write-Information "$($OptimalEnabledUsersWithMFA.Count) Optimally configured User accounts with MFA enabled found"


#Microsoft now just prompts people to set up MFA as default so they might have MFA enabled already, find the difference and activate it for those who have already set up a method (it means no change):
#"If the user hasn't yet registered MFA authentication methods, they receive a prompt to register the next time they sign in using modern authentication (such as via a web browser)." https://docs.microsoft.com/en-us/azure/active-directory/authentication/howto-mfa-userstates
$userswithnomfa = @()
$usersnotactivatedwithMFA = @()
ForEach($user in $enabledUsersWithoutMFA){

$e = ""
$e = Get-MsolUser -UserPrincipalName "$($user.UserPrincipalName)"

If($e.StrongAuthenticationMethods.IsDefault -match "True"){
$usersnotactivatedwithMFA += $user
}
Else{
$userswithnomfa += $user
}
}

#Remove boxes and bots, and admin accounts we can't manage
$userswithnomfa = $userswithnomfa | where-object -Property "UserPrincipalName" -NE "acsmailboxaccess@anthesisgroup.com"
$userswithnomfa = $userswithnomfa | where-object -Property "UserPrincipalName" -NE "ACSSupport@anthesisgroup.com"
$userswithnomfa = $userswithnomfa | where-object -Property "UserPrincipalName" -NE "NewGroupBot@anthesisgroup.com"
$usersnotactivatedwithMFA = $usersnotactivatedwithMFA | where-object -Property "UserPrincipalName" -NE "acsmailboxaccess@anthesisgroup.com"
$usersnotactivatedwithMFA = $usersnotactivatedwithMFA | where-object -Property "UserPrincipalName" -NE "ACSSupport@anthesisgroup.com"
$usersnotactivatedwithMFA = $usersnotactivatedwithMFA | where-object -Property "UserPrincipalName" -NE "NewGroupBot@anthesisgroup.com"
$usersnotactivatedwithMFA = $usersnotactivatedwithMFA | where-object -Property "UserPrincipalName" -NE "T1-Emily.Pressey@anthesisgroup.com"
$usersnotactivatedwithMFA = $usersnotactivatedwithMFA | where-object -Property "UserPrincipalName" -NE "t1-andrew.ost@anthesisgroup.com"
$usersnotactivatedwithMFA = $usersnotactivatedwithMFA | where-object -Property "UserPrincipalName" -NE "tempga@sustain.co.uk"



#Enable it for anyone in Spain, remove them from the list
ForEach($user in $userswithnomfa){
    If(($user.UsageLocation -eq "ES") -or ($user.UsageLocation -eq "CO")){
    Write-Host "$($user.UserPrincipalName) is in ESP or CO, enabling MFA" -ForegroundColor Yellow
    
    
    #Create an empty StrongAuthenticationRequirement object
    $emptyAuthObject = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $emptyAuthObject.RelyingParty = "*"
    $emptyAuthObject.State = "Enabled"
    $emptyAuthObject.RememberDevicesNotIssuedBefore = (Get-Date)

    #Get the GUID for the SSPR Group
    #$ssprGroup = Get-MsolGroup -SearchString "SSPR Testers"
    [guid]$ssprGroupObjectId = "fee80bd5-6e2f-4888-a51c-9581cf64eb18" #This is the GUID for the SSPR Testers Group

    #Figure out who to run this for
    $ESPToEnable = convertTo-arrayOfEmailAddresses $user.UserPrincipalName


$ESPToEnable | % {
    $thisUser = Get-MsolUser -UserPrincipalName $_
    Write-Verbose "MFA is currently set to [$($thisUser.StrongAuthenticationRequirements.State)] for $_"
    if([string]::IsNullOrWhiteSpace($thisUser.StrongAuthenticationRequirements)){
        Write-Verbose "Enabling MFA for $_"
        Set-MsolUser -UserPrincipalName $thisUser.UserPrincipalName -StrongAuthenticationRequirements $emptyAuthObject
        }
    else{Write-Verbose "MFA already [$($thisUser.StrongAuthenticationRequirements.State)] for $_"}
    Add-MsolGroupMember -GroupObjectId $ssprGroupObjectId -GroupMemberType User -GroupMemberObjectId $thisUser.ObjectId
    }
#Remove them from the list
$userswithnomfa = $userswithnomfa | where-object -Property "userPrincipalName" -NE $thisUser.UserPrincipalName       
 }
}


#Enable MFA for anyone who already has it set up
ForEach($MFAregistereduser in $usersnotactivatedwithMFA){

    write-host "Enabling MFA for $($MFAregistereduser.UserPrincipalName)" -ForegroundColor Yellow

    #Create an empty StrongAuthenticationRequirement object
    $emptyAuthObject = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $emptyAuthObject.RelyingParty = "*"
    $emptyAuthObject.State = "Enabled"
    $emptyAuthObject.RememberDevicesNotIssuedBefore = (Get-Date)

    #Get the GUID for the SSPR Group
    #$ssprGroup = Get-MsolGroup -SearchString "SSPR Testers"
    [guid]$ssprGroupObjectId = "fee80bd5-6e2f-4888-a51c-9581cf64eb18" #This is the GUID for the SSPR Testers Group


    #Figure out who to run this for
    $upnsToEnable = convertTo-arrayOfEmailAddresses $MFAregistereduser.UserPrincipalName


$upnsToEnable | % {
    $thisUser = Get-MsolUser -UserPrincipalName $_
    Write-Verbose "MFA is currently set to [$($thisUser.StrongAuthenticationRequirements.State)] for $_"
    if([string]::IsNullOrWhiteSpace($thisUser.StrongAuthenticationRequirements)){
        Write-Verbose "Enabling MFA for $_"
        Set-MsolUser -UserPrincipalName $thisUser.UserPrincipalName -StrongAuthenticationRequirements $emptyAuthObject
        }
    else{Write-Verbose "MFA already [$($thisUser.StrongAuthenticationRequirements.State)] for $_"}
    Add-MsolGroupMember -GroupObjectId $ssprGroupObjectId -GroupMemberType User -GroupMemberObjectId $thisUser.ObjectId
    }
}



#Send a message to Teams to get Emily to chase anyone who's bypassed the prompts
$allnewuserrequestsoldlist = get-graphListItems -tokenResponse $tokenResponseTeams -serverRelativeSiteUrl "https://anthesisllc.sharepoint.com/teams/hr" -listName "New User Requests" -expandAllFields
$allnewuserrequestsnewlist = get-graphListItems -tokenResponse $tokenResponseTeams -serverRelativeSiteUrl "https://anthesisllc.sharepoint.com/teams/People_Services_Team_All_365" -listName "New Starter Details" -expandAllFields


$startstochase = @()
ForEach($user in $userswithnomfa){

$thisuserold = ""
$thisusernew = ""
$thisuserold = $allnewuserrequestsoldlist | Where-Object -Property "Fields" -Match "$($user.DisplayName)"
$thisusernew = $allnewuserrequestsnewlist | Where-Object -Property "Fields" -Match "$($user.DisplayName)"

If($thisusernew){
#from new list
$expandstartdate = $thisusernew | select -ExpandProperty "Fields" | select -Property "StartDate"
$expandstartdate = $expandstartdate.StartDate.Split("T")[0]
Write-Host "$($user.UserPrincipalName) start date: $($expandstartdate)" -ForegroundColor Cyan
}
Else{
#From old list
$expandstartdate = $thisuserold | select -ExpandProperty "Fields" | select -Property "Start_x0020_Date"
$expandstartdate = $expandstartdate.Start_x0020_Date.Split("T")[0]
Write-Host "$($user.UserPrincipalName) start date: $($expandstartdate)" -ForegroundColor Cyan
}
If($expandstartdate){

$startdate = $expandstartdate | get-date
If($startdate -lt (get-date)){
$userobj = @{
"User" = $user.UserPrincipalName;
"Start Date" = $startdate
} 
$startstochase += "$($userobj.user) started $($userobj.'Start Date')"
}
}
}

$subject = "MFA status report"
$body = "<HTML><FONT FACE=`"Calibri`">Hello IT Team, these guys need chasing for MFA if they have started`r`n`r`n<BR><BR>"
ForEach($persontochase in $startstochase){$body += "$($persontochase) `r`n<BR>"}
#Send-MailMessage -To "cb1d8222.anthesisgroup.com@amer.teams.ms" -From "thehelpfulmfarobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Credential $adminCreds
send-graphMailMessage -tokenResponse $tokenResponseSmtp -fromUpn $groupAdmin -toAddresses "cb1d8222.anthesisgroup.com@amer.teams.ms" -Subject $subject -bodyHtml $body 


<#

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


$body += "<BR><BR>Love,`r`n`r`n<BR>The Helpful MFA Robot</FONT></HTML>"
#Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulmfarobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8
Write-Information $body
Send-MailMessage -To "emily.pressey@anthesisgroup.com" -From "thehelpfulmfarobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Credential $adminCreds
Send-MailMessage -To "kevin.maitland@anthesisgroup.com" -From "thehelpfulmfarobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Credential $adminCreds
Send-MailMessage -To "andrew.ost@anthesisgroup.com" -From "thehelpfulmfarobot@anthesisgroup.com" -SmtpServer "anthesisgroup-com.mail.protection.outlook.com" -Subject $subject -BodyAsHtml $body -Encoding UTF8 -Credential $adminCreds
Write-Information "Message Sent (maybe)"
#$body

#>

Disconnect-ExchangeOnline -Confirm:$False
Stop-Transcript